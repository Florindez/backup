VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmAsientoContable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comprobantes Contables"
   ClientHeight    =   9300
   ClientLeft      =   1140
   ClientTop       =   960
   ClientWidth     =   14415
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
   Icon            =   "frmAsientoContable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9300
   ScaleWidth      =   14415
   ShowInTaskbar   =   0   'False
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   12240
      TabIndex        =   20
      Top             =   8280
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
      Left            =   600
      TabIndex        =   21
      Top             =   8280
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   1296
      Buttons         =   4
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "Consul&tar"
      Tag1            =   "3"
      ToolTipText1    =   "Consultar"
      Caption2        =   "&Anular"
      Tag2            =   "4"
      ToolTipText2    =   "Anular"
      Caption3        =   "&Buscar"
      Tag3            =   "5"
      ToolTipText3    =   "Buscar"
      UserControlWidth=   5700
   End
   Begin TabDlg.SSTab tabAsiento 
      Height          =   8985
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   14250
      _ExtentX        =   25135
      _ExtentY        =   15849
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Asientos"
      TabPicture(0)   =   "frmAsientoContable.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chkSimulacion"
      Tab(0).Control(1)=   "cboFondo"
      Tab(0).Control(2)=   "dtpFechaDesde"
      Tab(0).Control(3)=   "dtpFechaHasta"
      Tab(0).Control(4)=   "tdgConsulta"
      Tab(0).Control(5)=   "lblDescrip(2)"
      Tab(0).Control(6)=   "lblDescrip(1)"
      Tab(0).Control(7)=   "lblDescrip(0)"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmAsientoContable.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraResumen"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAccion"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   10680
         TabIndex        =   17
         Top             =   7320
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
      Begin VB.CheckBox chkSimulacion 
         Caption         =   "Simulación"
         Height          =   255
         Left            =   -74580
         TabIndex        =   23
         ToolTipText     =   "Marcar para ver los movimientos de la simulación"
         Top             =   1230
         Width           =   1455
      End
      Begin VB.Frame fraResumen 
         ForeColor       =   &H00800000&
         Height          =   7665
         Left            =   390
         TabIndex        =   60
         Top             =   480
         Width           =   13485
         Begin VB.CommandButton cmdContracuenta 
            Caption         =   "..."
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   12780
            TabIndex        =   48
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   3480
            Width           =   375
         End
         Begin VB.CheckBox chkContracuenta 
            Caption         =   "Contracuenta"
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   360
            TabIndex        =   47
            Top             =   3480
            Width           =   1545
         End
         Begin VB.ComboBox cboTipoDocumentoDet 
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
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   3060
            Width           =   4605
         End
         Begin VB.TextBox txtNumDocumentoDet 
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
            Left            =   8370
            MaxLength       =   20
            TabIndex        =   45
            Top             =   3060
            Width           =   4395
         End
         Begin VB.ComboBox cboTipoPersonaContraparte 
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
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   2700
            Width           =   3465
         End
         Begin VB.TextBox txtPersonaContraparte 
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
            Left            =   5460
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   2700
            Width           =   7305
         End
         Begin VB.CommandButton cmdBusqueda 
            Caption         =   "..."
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   12780
            TabIndex        =   42
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   2700
            Width           =   375
         End
         Begin VB.ComboBox cboMonedaContable 
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
            Left            =   8370
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   1440
            Width           =   4425
         End
         Begin VB.CheckBox chkMovContable 
            Caption         =   "Es Sólo Movimiento Contable"
            Height          =   255
            Left            =   9360
            TabIndex        =   39
            Top             =   4350
            Width           =   3345
         End
         Begin VB.CommandButton cmdActualizar 
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   37
            Top             =   5100
            Width           =   375
         End
         Begin TAMControls.TAMTextBox txtMontoMovimiento 
            Height          =   315
            Left            =   1950
            TabIndex        =   35
            Top             =   4320
            Width           =   2800
            _ExtentX        =   4948
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
            Container       =   "frmAsientoContable.frx":0044
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
         Begin VB.TextBox txtNumDocumento 
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
            Left            =   8370
            MaxLength       =   20
            TabIndex        =   34
            Top             =   1050
            Width           =   4395
         End
         Begin VB.TextBox txtDescripFileAnalitica 
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
            Left            =   4230
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   8000
            Width           =   8535
         End
         Begin VB.ComboBox cboTipoDocumento 
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
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1050
            Width           =   4605
         End
         Begin VB.ComboBox cboTipoAuxiliar 
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
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   8130
            Width           =   1005
         End
         Begin VB.ComboBox cboTipoFile 
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
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   8130
            Width           =   1065
         End
         Begin VB.TextBox txtDescripCuenta 
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
            Left            =   1950
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   29
            Top             =   1980
            Width           =   10815
         End
         Begin VB.CommandButton cmdBusqueda 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   12780
            TabIndex        =   28
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   2340
            Width           =   375
         End
         Begin VB.CommandButton cmdBusqueda 
            Caption         =   "..."
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   12780
            TabIndex        =   27
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   8220
            Width           =   375
         End
         Begin VB.TextBox txtDescripAuxiliar 
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
            Left            =   5460
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   2340
            Width           =   7305
         End
         Begin VB.CheckBox chkAjuste 
            Caption         =   "Ajuste por Tipo de Cambio"
            Height          =   255
            Left            =   12810
            TabIndex        =   18
            Top             =   1020
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.ComboBox cboVerifica 
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
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   8400
            Width           =   1635
         End
         Begin VB.ComboBox cboDigita 
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
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   8400
            Width           =   1635
         End
         Begin VB.TextBox txtDescripMovimiento 
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
            Left            =   1950
            TabIndex        =   12
            Top             =   4680
            Width           =   10935
         End
         Begin VB.ComboBox cboMonedaMovimiento 
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
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   3960
            Width           =   2800
         End
         Begin VB.ComboBox cboNaturaleza 
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
            Left            =   6240
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   3960
            Width           =   2800
         End
         Begin VB.TextBox txtCodAnalitica 
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
            Left            =   2520
            MaxLength       =   8
            TabIndex        =   11
            Top             =   2340
            Width           =   1605
         End
         Begin VB.TextBox txtCodFile 
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
            Left            =   1950
            MaxLength       =   3
            TabIndex        =   10
            Top             =   2340
            Width           =   555
         End
         Begin VB.CommandButton cmdBusqueda 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   12780
            TabIndex        =   22
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   1980
            Width           =   375
         End
         Begin VB.TextBox txtCodCuenta 
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
            Left            =   1950
            MaxLength       =   10
            TabIndex        =   9
            Top             =   1980
            Width           =   1665
         End
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Quitar detalle"
            Top             =   5940
            Width           =   375
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Agregar detalle"
            Top             =   5520
            Width           =   375
         End
         Begin VB.TextBox txtMontoAsiento 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1980
            TabIndex        =   7
            Top             =   8130
            Width           =   1155
         End
         Begin VB.ComboBox cboMoneda 
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
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1440
            Width           =   1395
         End
         Begin VB.TextBox txtDescripAsiento 
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
            MaxLength       =   100
            TabIndex        =   3
            Top             =   680
            Width           =   10815
         End
         Begin VB.ComboBox cboModulo 
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
            Left            =   4530
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   300
            Width           =   3135
         End
         Begin VB.TextBox txtHoraAsiento 
            Alignment       =   2  'Center
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
            Left            =   11910
            TabIndex        =   6
            Text            =   "00:00"
            Top             =   300
            Width           =   855
         End
         Begin MSComCtl2.DTPicker dtpFechaAsiento 
            Height          =   315
            Left            =   9150
            TabIndex        =   5
            Top             =   300
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
            Format          =   202244097
            CurrentDate     =   38068
         End
         Begin TAMControls.TAMTextBox txtTipoCambioMovimiento 
            Height          =   315
            Left            =   10560
            TabIndex        =   40
            Top             =   3960
            Width           =   1035
            _ExtentX        =   1826
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
            Container       =   "frmAsientoContable.frx":0060
            Text            =   "0.00000000"
            Decimales       =   8
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   8
            MaximoValor     =   999999999
         End
         Begin TrueOleDBGrid60.TDBGrid tdgMovimiento 
            Bindings        =   "frmAsientoContable.frx":007C
            Height          =   1365
            Left            =   930
            OleObjectBlob   =   "frmAsientoContable.frx":0098
            TabIndex        =   24
            Top             =   5070
            Width           =   11985
         End
         Begin TAMControls.TAMTextBox txtMontoContable 
            Height          =   315
            Left            =   6240
            TabIndex        =   49
            Top             =   4320
            Width           =   2805
            _ExtentX        =   4948
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
            Container       =   "frmAsientoContable.frx":A7D6
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
         Begin TAMControls.TAMTextBox txtTipoCambio 
            Height          =   315
            Left            =   5520
            TabIndex        =   50
            Top             =   1440
            Width           =   1035
            _ExtentX        =   1826
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
            Container       =   "frmAsientoContable.frx":A7F2
            Text            =   "0.00000000"
            Decimales       =   8
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   8
            MaximoValor     =   999999999
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   37
            Left            =   360
            TabIndex        =   97
            Top             =   1500
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Cambio"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   36
            Left            =   4320
            TabIndex        =   96
            Top             =   1500
            Width           =   1095
         End
         Begin VB.Label lblContracuenta 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1950
            TabIndex        =   95
            Top             =   3450
            Width           =   10815
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Doc."
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   35
            Left            =   360
            TabIndex        =   94
            Top             =   3120
            Width           =   1125
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Nro. Doc."
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   34
            Left            =   7140
            TabIndex        =   93
            Top             =   3090
            Width           =   1125
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Persona"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   33
            Left            =   360
            TabIndex        =   92
            Top             =   2730
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Mon. Cont."
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   32
            Left            =   7140
            TabIndex        =   91
            Top             =   1500
            Width           =   1155
         End
         Begin VB.Label lblDescripTC 
            AutoSize        =   -1  'True
            Caption         =   "(USD/SOL)"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   11760
            TabIndex        =   90
            Top             =   3990
            Width           =   1080
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Analitica"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   31
            Left            =   11970
            TabIndex        =   89
            Top             =   7350
            Width           =   495
         End
         Begin VB.Label lblDescrip 
            Caption         =   "File"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   30
            Left            =   7080
            TabIndex        =   88
            Top             =   8040
            Width           =   345
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Nro. Doc."
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   29
            Left            =   7140
            TabIndex        =   87
            Top             =   1080
            Width           =   1125
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Doc."
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   28
            Left            =   360
            TabIndex        =   86
            Top             =   1065
            Width           =   1125
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Auxiliar"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   27
            Left            =   270
            TabIndex        =   85
            Top             =   8010
            Width           =   765
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo File"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   26
            Left            =   330
            TabIndex        =   84
            Top             =   8280
            Width           =   705
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Auxiliar"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   25
            Left            =   4560
            TabIndex        =   83
            Top             =   2400
            Width           =   975
         End
         Begin VB.Line Line2 
            BorderColor     =   &H8000000C&
            BorderStyle     =   6  'Inside Solid
            X1              =   360
            X2              =   12660
            Y1              =   3870
            Y2              =   3870
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Moneda"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   24
            Left            =   360
            TabIndex        =   82
            Top             =   3990
            Width           =   1125
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Naturaleza"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   23
            Left            =   5040
            TabIndex        =   81
            Top             =   3990
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Descripción Mov."
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   21
            Left            =   360
            TabIndex        =   80
            Top             =   4695
            Width           =   1575
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Monto Cont."
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   20
            Left            =   5040
            TabIndex        =   79
            Top             =   4350
            Width           =   1125
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Monto"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   19
            Left            =   360
            TabIndex        =   78
            Top             =   4350
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Analítica"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   18
            Left            =   360
            TabIndex        =   77
            Top             =   2370
            Width           =   885
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Cuenta"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   17
            Left            =   360
            TabIndex        =   76
            Top             =   2025
            Width           =   975
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000C&
            BorderStyle     =   6  'Inside Solid
            X1              =   360
            X2              =   12990
            Y1              =   1840
            Y2              =   1840
         End
         Begin VB.Label lblTotalHaberMN 
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
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   10830
            TabIndex        =   55
            Top             =   6450
            Width           =   2085
         End
         Begin VB.Label lblTotalDebeMN 
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
            Left            =   7800
            TabIndex        =   54
            Top             =   6450
            Width           =   2055
         End
         Begin VB.Label lblTotalHaberME 
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
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   4920
            TabIndex        =   53
            Top             =   6450
            Width           =   1995
         End
         Begin VB.Label lblTotalDebeME 
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
            Left            =   1890
            TabIndex        =   52
            Top             =   6450
            Width           =   2055
         End
         Begin VB.Label lblDescrip 
            BackStyle       =   0  'Transparent
            Caption         =   "Debe ME"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   930
            TabIndex        =   75
            Top             =   6510
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            BackStyle       =   0  'Transparent
            Caption         =   "Haber ME"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   11
            Left            =   3990
            TabIndex        =   74
            Top             =   6480
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            BackStyle       =   0  'Transparent
            Caption         =   "Debe MN"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   12
            Left            =   6960
            TabIndex        =   73
            Top             =   6480
            Width           =   915
         End
         Begin VB.Label lblDescrip 
            BackStyle       =   0  'Transparent
            Caption         =   "Haber MN"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   13
            Left            =   9915
            TabIndex        =   72
            Top             =   6480
            Width           =   930
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Moneda"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   22
            Left            =   2130
            TabIndex        =   71
            Top             =   8100
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha Contabilización"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   16
            Left            =   7920
            TabIndex        =   70
            Top             =   8460
            Width           =   1935
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Revisión"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   15
            Left            =   4770
            TabIndex        =   69
            Top             =   8460
            Width           =   735
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Digitador"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   14
            Left            =   960
            TabIndex        =   68
            Top             =   8460
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Hora Creac."
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   9
            Left            =   10830
            TabIndex        =   67
            Top             =   330
            Width           =   1065
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Cambio"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   8
            Left            =   9360
            TabIndex        =   66
            Top             =   3990
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha Creac."
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   7
            Left            =   7890
            TabIndex        =   65
            Top             =   345
            Width           =   1275
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Monto"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   6
            Left            =   3210
            TabIndex        =   64
            Top             =   8130
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Módulo"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   5
            Left            =   3750
            TabIndex        =   63
            Top             =   345
            Width           =   735
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Descripción"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   4
            Left            =   360
            TabIndex        =   62
            Top             =   705
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Número"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   3
            Left            =   360
            TabIndex        =   61
            Top             =   330
            Width           =   735
         End
         Begin VB.Label lblFechaContable 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "99/99/9999"
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
            Left            =   10185
            TabIndex        =   56
            Top             =   8415
            Width           =   1155
         End
         Begin VB.Label lblNumAsiento 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000000000"
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
            Left            =   1950
            TabIndex        =   51
            Top             =   315
            Width           =   1515
         End
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
         Left            =   -73650
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   660
         Width           =   9315
      End
      Begin MSComCtl2.DTPicker dtpFechaDesde 
         Height          =   315
         Left            =   -62640
         TabIndex        =   1
         Top             =   660
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   202244097
         CurrentDate     =   38068
      End
      Begin MSComCtl2.DTPicker dtpFechaHasta 
         Height          =   315
         Left            =   -62640
         TabIndex        =   2
         Top             =   1125
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   202244097
         CurrentDate     =   38068
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmAsientoContable.frx":A80E
         Height          =   6285
         Left            =   -74580
         OleObjectBlob   =   "frmAsientoContable.frx":A828
         TabIndex        =   25
         Top             =   1650
         Width           =   13335
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Periodo Final"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   -64080
         TabIndex        =   59
         Top             =   1140
         Width           =   1335
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Periodo Inicial"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   -64080
         TabIndex        =   58
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fondo"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   -74580
         TabIndex        =   57
         Top             =   690
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmAsientoContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()                      As String, arrMoneda()              As String
Dim arrMonedaMovimiento()           As String, arrNaturaleza()          As String
Dim arrModulo()                     As String, arrTipoFile()            As String
Dim arrTipoAuxiliar()               As String, arrTipoDocumento()       As String
Dim arrTipoPersonaContraparte()     As String, arrTipoDocumentoDet()    As String
Dim arrMonedaContable()             As String

Dim strCodFondo                 As String, strCodMoneda                         As String
Dim strCodMonedaMovimiento      As String, strCodNaturaleza                     As String
Dim strCodModulo                As String, strEstado                            As String

Dim strCodMonedaContable            As String

Dim strCodPersonaContraparte        As String, strTipoPersonaContraparte        As String
Dim strDescripPersonaContraparte    As String

Dim strCodContracuenta              As String, strCodFileContracuenta                   As String
Dim strCodAnaliticaContracuenta     As String, strDescripFileAnaliticaContracuenta      As String
Dim strDescripContracuenta          As String, strTipoFileContracuenta                  As String


Dim strIndUltimoMoviiento       As String, strIndSoloMovimientoContable         As String
Dim strTipoDocumentoDet         As String


Dim strIndAuxiliar              As String, strTipoAuxiliar          As String
Dim strDescripCuenta            As String, strCodCuenta             As String
Dim strCodAuxiliar              As String, strDescripAuxiliar       As String
Dim strCodFile                  As String, strCodAnalitica          As String
Dim strTipoFile                 As String, strDescripFileAnalitica  As String
Dim intRegistro                 As Integer, strTipoDocumento        As String
Dim strNumDocumento             As String, adoRegistroAux           As ADODB.Recordset
Dim numSecMovimiento            As Long, strTipoProceso             As String
Dim adoConsulta                 As ADODB.Recordset
Dim strCodMonedaParEvaluacion   As String
Dim strCodMonedaParPorDefecto   As String
Dim adoRegistroAuxTC            As ADODB.Recordset
Dim strTipoCambioReemplazoXML   As String
Dim objTipoCambioReemplazoXML   As DOMDocument60
    

Public Sub Buscar()

    Dim strSQL          As String
    Dim strFechaDesde   As String, strFechaHasta As String
    Dim datFecha        As Date
    
    Set adoConsulta = New ADODB.Recordset
    
    strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
    datFecha = DateAdd("d", 1, dtpFechaHasta.Value)
    strFechaHasta = Convertyyyymmdd(datFecha)
    
    strSQL = "SELECT NumAsiento,DescripAsiento,FechaAsiento,ISNULL(AP.CodSigno,'ALL') AS DescripMoneda," & _
            "AC.CodMoneda,ValorTipoCambio,MontoAsiento,Convert(char(5),FechaAsiento,108) as Hora, TipoDocumento, NumDocumento " & _
            "FROM AsientoContable AC LEFT JOIN Moneda AP ON(AP.CodMoneda=AC.CodMoneda) " & _
            "WHERE (FechaAsiento >= '" & strFechaDesde & "' AND FechaAsiento < '" & strFechaHasta & "') AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND EstadoAsiento = '01' " & _
            "ORDER BY NumAsiento"
                        
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
    
End Sub


Private Sub TotalizarMovimientos()
    
    Dim strFechaDesde   As String, strFechaHasta As String
    Dim datFecha        As Date
    Dim strSQL          As String
    Dim intRegistro     As Integer
        
        
    Dim dblMontoDebe        As Double, dblMontoHaber        As Double
    Dim dblAcumuladoDebeMN  As Double, dblAcumuladoDebeME   As Double
    Dim dblAcumuladoHaberMN As Double, dblAcumuladoHaberME  As Double
    Dim intContador         As Integer

    'intContador = adoRegistroAux.RecordCount - 1
                            
    lblTotalDebeME.Caption = "0"
    lblTotalHaberME.Caption = "0"
    lblTotalDebeMN.Caption = "0"
    lblTotalHaberMN.Caption = "0"
    
    dblAcumuladoDebeMN = 0
    dblAcumuladoHaberMN = 0
    dblAcumuladoDebeME = 0
    dblAcumuladoHaberME = 0
        
    If Not adoRegistroAux.EOF And Not adoRegistroAux.BOF Then
        adoRegistroAux.MoveFirst
    End If
    
    While Not adoRegistroAux.EOF
        dblMontoDebe = CDbl(adoRegistroAux.Fields("MontoDebe"))
        dblMontoHaber = CDbl(adoRegistroAux.Fields("MontoHaber"))
    
        If adoRegistroAux.Fields("CodMonedaMovimiento") = Codigo_Moneda_Local Then
            dblAcumuladoDebeMN = dblAcumuladoDebeMN + dblMontoDebe
            dblAcumuladoHaberMN = dblAcumuladoHaberMN + dblMontoHaber
        Else
            dblAcumuladoDebeME = dblAcumuladoDebeME + dblMontoDebe
            dblAcumuladoHaberME = dblAcumuladoHaberME + dblMontoHaber
        End If
    
        adoRegistroAux.MoveNext
    Wend
    
    lblTotalDebeME.Caption = CStr(dblAcumuladoDebeME)
    lblTotalHaberME.Caption = CStr(dblAcumuladoHaberME)
    lblTotalDebeMN.Caption = CStr(dblAcumuladoDebeMN)
    lblTotalHaberMN.Caption = CStr(dblAcumuladoHaberMN)
    
            
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabAsiento
        .TabEnabled(0) = True
        .Tab = 0
    End With
    Call Buscar
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Diario General Analítico"
    
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Mayor General Analítico"
    
End Sub
Private Sub Deshabilita()

End Sub



Private Sub Habilita()

End Sub


Public Sub Imprimir()

    Call SubImprimir(1)
    
End Sub



'Private Sub LimpiarDatos()
'
'    txtCodCuenta.Text = Valor_Caracter
'    txtDescripCuenta.Text = Valor_Caracter
'    txtDescripFileAnalitica.Text = Valor_Caracter
'    txtDescripAuxiliar.Text = Valor_Caracter
'
'    strTipoAuxiliar = Valor_Caracter
'    strCodAuxiliar = Valor_Caracter
'    strCodCuenta = Valor_Caracter
'
'    txtCodFile.Text = Valor_Caracter
'    txtCodAnalitica.Text = Valor_Caracter
'    txtMontoMovimiento.Text = "0"
'    txtDescripMovimiento.Text = Valor_Caracter
'
'End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub SubImprimir(Index As Integer)

    Dim adoRegistro             As ADODB.Recordset
    Dim strSeleccionRegistro    As String
    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()

    Select Case Index
            
                
         Case 1
            

            gstrNumAsiento = Valor_Caracter
            
            If tabAsiento.Tab = 1 And (strEstado = Reg_Edicion Or strEstado = Reg_Consulta) Then
                '*** Comprobante seleccionado ***
'                strSeleccionRegistro = "{AsientoContable.NumAsiento} = '" & Trim(lblNumAsiento.Caption) & "'"
'                strSeleccionRegistro = strSeleccionRegistro & " AND {AsientoContableDetalle.NumAsiento} = '" & Trim(lblNumAsiento.Caption) & "'"
'                strSeleccionRegistro = strSeleccionRegistro & " AND {AsientoContableDetalle.FechaMovimiento} = '" & Convertyyyymmdd(dtpFechaAsiento.Value) & "'"
'                gstrSelFrml = strSeleccionRegistro
                gstrNumAsiento = CStr(CLng(Trim(lblNumAsiento.Caption)))
                gstrCodMonedaReporte = Codigo_Moneda_Local
                gstrSelFrml = "1"
            Else
                '*** Lista de comprobantes por rango de fecha ***
                strSeleccionRegistro = "{AsientoContable.FechaAsiento} IN 'Fch1' TO 'Fch2'"
                gstrSelFrml = strSeleccionRegistro
                'frmRangoFecha.Show vbModal
                frmFiltroReporte.strCodFondo = strCodFondo
                frmFiltroReporte.strCodAdministradora = gstrCodAdministradora
                
                frmFiltroReporte.chkOpcionFiltro(1).Enabled = False
                frmFiltroReporte.chkOpcionFiltro(1).Value = 0
                frmFiltroReporte.txtCodCuenta.Enabled = False
                frmFiltroReporte.cmdBusquedaCuenta.Enabled = False
                
                frmFiltroReporte.chkOpcionFiltro(2).Enabled = True
                frmFiltroReporte.chkOpcionFiltro(2).Value = 0
                frmFiltroReporte.txtNumAsiento.Enabled = False
                
                frmFiltroReporte.dtpFechaInicial = dtpFechaDesde
                frmFiltroReporte.dtpFechaFinal = dtpFechaHasta
                
                frmFiltroReporte.Show vbModal
                
            End If

            If gstrSelFrml <> "0" Then
                Set frmReporte = New frmVisorReporte
'INICIO: REVISAR NUEVA VERSION SPECTRUM 1_5
'                ReDim aReportParamS(9)
'FIN: REVISAR NUEVA VERSION SPECTRUM 1_5
                ReDim aReportParamS(8)
                ReDim aReportParamFn(5)
                ReDim aReportParamF(5)
                
                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "FechaDesde"
                aReportParamFn(2) = "FechaHasta"
                aReportParamFn(3) = "Hora"
                aReportParamFn(4) = "Fondo"
                aReportParamFn(5) = "NombreEmpresa"

                aReportParamF(0) = gstrLogin

                If tabAsiento.Tab = 1 And (strEstado = Reg_Edicion Or strEstado = Reg_Consulta) Then
                    aReportParamF(1) = dtpFechaAsiento
                    aReportParamF(2) = dtpFechaAsiento
                Else
                    aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
                    aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                End If
                
                aReportParamF(3) = Format(Time(), "hh:mm:ss")
                aReportParamF(4) = Trim(cboFondo.Text)
                aReportParamF(5) = gstrNombreEmpresa & Space(1)

                'SP
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora

                If tabAsiento.Tab = 1 And (strEstado = Reg_Edicion Or strEstado = Reg_Consulta) Then
                    aReportParamS(2) = Convertyyyymmdd(dtpFechaAsiento.Value)
                    aReportParamS(3) = Convertyyyymmdd(dtpFechaAsiento.Value)
                Else
                    aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                    aReportParamS(3) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))
                End If

                aReportParamS(4) = gstrCodMonedaReporte 'strCodMoneda
                aReportParamS(5) = gstrCodClaseTipoCambioFondo 'Codigo_Listar_Todos
                aReportParamS(6) = gstrValorTipoCambioCierre   '"0000000000"
                'aReportParamS(7) = Codigo_Listar_Todos

                If Trim(gstrNumAsiento) <> Valor_Caracter Then
                    aReportParamS(7) = Codigo_Listar_Individual '"I"
                    aReportParamS(8) = gstrNumAsiento
                Else
                    aReportParamS(7) = Codigo_Listar_Todos '"T"
                    aReportParamS(8) = "%"
                End If

'INICIO: REVISAR NUEVA VERSION SPECTRUM 1_5
'                If chkSimulacion.Value Then
'                    aReportParamS(9) = "1"
'                Else
'                    aReportParamS(9) = "0"
'                End If
'
'                gstrNameRepo = "LibroDiarioAnalitico"
'FIN: REVISAR NUEVA VERSION SPECTRUM 1_5

                If chkSimulacion.Value Then
                    gstrNameRepo = "SLibroDiario"
                Else
                    gstrNameRepo = "LibroDiarioMM"
                End If

                
            End If
                
                
        Case 2

            '*** Lista de comprobantes por rango de fecha ***
            strSeleccionRegistro = "{AsientoContable.FechaAsiento} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            'frmRangoFecha.Show vbModal
            frmFiltroReporte.strCodFondo = strCodFondo
            frmFiltroReporte.strCodAdministradora = gstrCodAdministradora
            frmFiltroReporte.chkOpcionFiltro(1).Enabled = True
            frmFiltroReporte.chkOpcionFiltro(1).Value = 1
            frmFiltroReporte.txtCodCuenta.Enabled = True
            frmFiltroReporte.cmdBusquedaCuenta.Enabled = True
            
            frmFiltroReporte.chkOpcionFiltro(2).Enabled = False
            frmFiltroReporte.chkOpcionFiltro(2).Value = 0
            frmFiltroReporte.txtNumAsiento.Enabled = False
        
            frmFiltroReporte.dtpFechaInicial = dtpFechaDesde
            frmFiltroReporte.dtpFechaFinal = dtpFechaHasta
            
            
            frmFiltroReporte.Show vbModal

            If gstrSelFrml <> "0" Then
                Set adoRegistro = New ADODB.Recordset

                '*** Se Realizó Cierre anteriormente ? ***

                adoComm.CommandText = "{ call up_GNValidaCierreRealizado('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)) & "','" & _
                    Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)))) & "') }"

                Set adoRegistro = adoComm.Execute
                If adoRegistro.EOF Then
                    MsgBox "El Cierre del Día " & Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10) & " No fué realizado.", vbCritical, Me.Caption

                    adoRegistro.Close: Set adoRegistro = Nothing
                    Exit Sub
                End If
                adoRegistro.Close: Set adoRegistro = Nothing

                Set frmReporte = New frmVisorReporte
                'Dim strCuenta As String

                ReDim aReportParamS(8)
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
                aReportParamS(3) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))
                aReportParamS(4) = gstrCodMonedaReporte 'strCodMoneda
                
                aReportParamS(5) = gstrCodClaseTipoCambioFondo 'Codigo_Listar_Todos
                aReportParamS(6) = gstrValorTipoCambioCierre   '"0000000000"
                aReportParamS(7) = Codigo_Listar_Todos
                
                If Trim(gstrCodCuenta) = Valor_Caracter Or gstrCodCuenta = "0000000000" Then
                    aReportParamS(8) = "%" 'gstrCodCuenta '"0000000000"
                Else
                    aReportParamS(8) = gstrCodCuenta
                End If
                            
                'gstrNameRepo = "LibroMayorAnalitico"
                
                If chkSimulacion.Value Then
                    gstrNameRepo = "SLibroMayor"
                Else
                    gstrNameRepo = "HistLibroMayor1"
                End If

                


            End If
    End Select

    If gstrSelFrml = "0" Then Exit Sub
    
    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show 'vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
                
End Sub



Private Function ValidaCuadreContable() As Boolean

    Dim curMontoDebe        As Currency, curMontoHaber      As Currency
    Dim curMontoContable    As Currency
    
    ValidaCuadreContable = False
    
    adoRegistroAux.MoveFirst
    
    curMontoContable = 0
    
    Do While Not adoRegistroAux.EOF
        If adoRegistroAux.Fields("IndDebeHaber") = Codigo_Tipo_Naturaleza_Debe Then
            curMontoDebe = CCur(adoRegistroAux.Fields("MontoContable"))
        Else
            curMontoHaber = CCur(adoRegistroAux.Fields("MontoContable"))
        End If
        curMontoContable = curMontoContable + CCur(adoRegistroAux.Fields("MontoContable"))
                
        adoRegistroAux.MoveNext
    Loop
    
    If curMontoContable <> 0 Then Exit Function
    
    ValidaCuadreContable = True
    
End Function

Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset
    Dim strSQL      As String
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = adoRegistro("FechaCuota")
            gdblTipoCambio = adoRegistro("ValorTipoCambio")
            gstrCodMoneda = adoRegistro("CodMoneda")
            dtpFechaDesde.Value = gdatFechaActual
            dtpFechaHasta.Value = dtpFechaDesde.Value
            
            gstrPeriodoActual = Format(Year(gdatFechaActual), "0000")
            gstrMesActual = Format(Month(gdatFechaActual), "00")
            
            'Carga las monedas contables de los fondos
            strSQL = "{ call up_ACSelDatosParametro(70,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
            CargarControlLista strSQL, cboMonedaContable, arrMonedaContable(), Sel_Todos
        
            If cboMonedaContable.ListCount > 0 Then cboMonedaContable.ListIndex = 0
            
            txtTipoCambioMovimiento.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, gstrCodMoneda, Codigo_Moneda_Local))
            If CDbl(txtTipoCambioMovimiento.Text) = 0 Then txtTipoCambioMovimiento.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, gdatFechaActual), gstrCodMoneda, Codigo_Moneda_Local))
            gdblTipoCambio = CDbl(txtTipoCambioMovimiento.Text)
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        
        
        
            Call Buscar
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub


Private Sub cboModulo_Click()

    strCodModulo = ""
    If cboModulo.ListIndex < 0 Then Exit Sub
    
    strCodModulo = Trim(arrModulo(cboModulo.ListIndex))
    
End Sub


Private Sub cboMoneda_Click()

    strCodMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
    strCodMonedaParEvaluacion = strCodMoneda & Codigo_Moneda_Local
    
    If strCodMoneda <> Codigo_Moneda_Local Then
        strCodMonedaParPorDefecto = ObtenerMonedaParPorDefecto(gstrCodClaseTipoCambioOperacionFondo, strCodMonedaParEvaluacion)
        txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, Mid(strCodMonedaParPorDefecto, 1, 2), Mid(strCodMonedaParPorDefecto, 3, 2)))
    Else
        strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
        txtTipoCambio.Text = "1"
    End If
    
    lblDescripTC.Caption = "(" + Trim(ObtenerCodSignoMoneda(Mid(strCodMonedaParPorDefecto, 3, 2))) + "/" + Trim(ObtenerCodSignoMoneda(Mid(strCodMonedaParPorDefecto, 1, 2))) + ")"
    
    If strCodMonedaMovimiento <> Codigo_Moneda_Local Then
        txtTipoCambioMovimiento.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, Mid(strCodMonedaParPorDefecto, 1, 2), Mid(strCodMonedaParPorDefecto, 3, 2)))
        txtTipoCambioMovimiento.Enabled = True
    Else
        txtTipoCambioMovimiento.Text = "1"
        txtTipoCambioMovimiento.Enabled = True 'False
    End If
    Call txtTipoCambioMovimiento_KeyPress(vbKeyReturn)
    
'    txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaAsiento.Value, Codigo_Moneda_Local, strCodMoneda))
'    If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaAsiento.Value), Codigo_Moneda_Local, strCodMoneda))
    
End Sub


Private Sub cboMonedaContable_Click()
    
    Dim intRegistro As Integer
    
    strCodMonedaContable = Valor_Caracter
    If cboMonedaContable.ListIndex < 0 Then Exit Sub
    
    strCodMonedaContable = arrMonedaContable(cboMonedaContable.ListIndex)
    
'    If strCodMonedaContable = Valor_Caracter Then
'        chkMovContable.Value = vbUnchecked
'        Call chkMovContable_Click
'    Else
'        chkMovContable.Value = vbChecked
'        Call chkMovContable_Click
'    End If
    
    strCodMonedaContable = IIf(strCodMonedaContable = Valor_Caracter, Codigo_Moneda_Local, strCodMonedaContable)
    
    If strIndSoloMovimientoContable = Valor_Indicador Then
        intRegistro = ObtenerItemLista(arrMonedaMovimiento(), strCodMonedaContable)
        If intRegistro >= 0 Then cboMonedaMovimiento.ListIndex = intRegistro
    End If
    
End Sub

Private Sub cboMonedaMovimiento_Click()

    Dim dblValorTC As Double

    strCodMonedaMovimiento = Valor_Caracter
    If cboMonedaMovimiento.ListIndex < 0 Then Exit Sub
    
    strCodMonedaMovimiento = Trim(arrMonedaMovimiento(cboMonedaMovimiento.ListIndex))
    
    strCodMonedaParEvaluacion = strCodMonedaMovimiento & Codigo_Moneda_Local
    
    If strCodMonedaMovimiento <> Codigo_Moneda_Local Then
        strCodMonedaParPorDefecto = ObtenerMonedaParPorDefecto(gstrCodClaseTipoCambioOperacionFondo, strCodMonedaParEvaluacion)
    Else
        strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
    End If
    
    If strCodMonedaParPorDefecto = "0000" Then strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
    
    If strCodMoneda <> Codigo_Moneda_Local Then 'strCodMoneda <> Codigo_Moneda_Local Then
        txtTipoCambioMovimiento.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, Mid(strCodMonedaParPorDefecto, 1, 2), Mid(strCodMonedaParPorDefecto, 3, 2)))
        txtTipoCambioMovimiento.Enabled = True
        lblDescripTC.Caption = "(" + Trim(ObtenerCodSignoMoneda(Mid(strCodMonedaParPorDefecto, 3, 2))) + "/" + Trim(ObtenerCodSignoMoneda(Mid(strCodMonedaParPorDefecto, 1, 2))) + ")"
    Else
        txtTipoCambioMovimiento.Text = "1"
        txtTipoCambioMovimiento.Enabled = True 'False
        lblDescripTC.Caption = Valor_Caracter
    End If
    
    If strIndSoloMovimientoContable = Valor_Caracter Then
'        txtMontoContable.Text = txtMontoMovimiento.Text
'    Else
        dblValorTC = CDbl(txtTipoCambioMovimiento.Text)
        dblValorTC = ObtenerTipoCambioArbitraje(dblValorTC, strCodMonedaParEvaluacion, strCodMonedaParPorDefecto)
        txtMontoContable.Text = CStr(CDbl(txtMontoMovimiento.Text) * dblValorTC)
    End If
    
End Sub


Private Sub cboNaturaleza_Click()

    strCodNaturaleza = Valor_Caracter
    If cboNaturaleza.ListIndex < 0 Then Exit Sub
    
    strCodNaturaleza = Trim(arrNaturaleza(cboNaturaleza.ListIndex))
    
    If strIndSoloMovimientoContable = Valor_Caracter Then
        txtMontoMovimiento.Text = Abs(CDbl(txtMontoMovimiento.Text))
        If strCodNaturaleza = Codigo_Tipo_Naturaleza_Haber Then
            txtMontoMovimiento.Text = CStr(Abs(CDbl(txtMontoMovimiento.Text)) * -1)
        End If
    Else
        txtMontoContable.Text = Abs(CDbl(txtMontoContable.Text))
        If strCodNaturaleza = Codigo_Tipo_Naturaleza_Haber Then
            txtMontoContable.Text = CStr(Abs(CDbl(txtMontoContable.Text)) * -1)
        End If
    End If
    
End Sub




Private Sub cboTipoPersonaContraparte_Click()
    
    strTipoPersonaContraparte = Valor_Caracter
    
    If cboTipoPersonaContraparte.ListIndex < 0 Then Exit Sub
    
    strTipoPersonaContraparte = arrTipoPersonaContraparte(cboTipoPersonaContraparte.ListIndex)

    txtPersonaContraparte.Text = ""
    
    If cboTipoPersonaContraparte.ListIndex > 0 Then
        cmdBusqueda(3).Enabled = True
    Else
        cmdBusqueda(3).Enabled = False
    End If

End Sub

'Private Sub cboTipoAuxiliar_Click()
'
'
'    strTipoAuxiliar = ""
'
'    If cboTipoAuxiliar.ListIndex < 0 Then Exit Sub
'
'    strTipoAuxiliar = Trim(arrTipoAuxiliar(cboTipoAuxiliar.ListIndex))
'    'If strTipoAuxiliar = Valor_Caracter Then strTipoAuxiliar = "00"
'
'    strCodAuxiliar = ""
'    strDescripAuxiliar = ""
'    txtDescripAuxiliar.Text = ""
'
'End Sub

Private Sub cboTipoDocumento_Click()

    strTipoDocumento = Valor_Caracter
    If cboTipoDocumento.ListIndex < 0 Then Exit Sub
    
    strTipoDocumento = arrTipoDocumento(cboTipoDocumento.ListIndex)

End Sub

Private Sub cboTipoDocumentoDet_Click()
    strTipoDocumentoDet = Valor_Caracter
    
    If cboTipoDocumentoDet.ListIndex < 0 Then Exit Sub
    
    strTipoDocumentoDet = arrTipoDocumentoDet(cboTipoDocumentoDet.ListIndex)
End Sub



Private Sub chkContracuenta_Click()

    If chkContracuenta.Value = vbChecked Then
        cmdContracuenta.Enabled = True
        If strCodContracuenta <> Valor_Caracter Then
            lblContracuenta.Caption = strCodContracuenta + " / " + strCodFileContracuenta + "-" + strCodAnaliticaContracuenta
        Else
            lblContracuenta.Caption = Valor_Caracter
        End If
    Else
        cmdContracuenta.Enabled = False
        lblContracuenta.Caption = Valor_Caracter
    End If

End Sub



Private Sub chkMovContable_Click()

'    If chkMovContable.Value = vbChecked Then
'        txtMontoMovimiento.Text = "0"
'        txtMontoMovimiento.Enabled = False
'    Else
'        txtMontoMovimiento.Enabled = True
'    End If
    Dim intRegistro As Integer

    If chkMovContable.Value = vbChecked Then
        strIndSoloMovimientoContable = Valor_Indicador
        txtMontoMovimiento.Text = "0"
        txtMontoMovimiento.Enabled = False
        txtMontoContable.Enabled = True
        
        lblDescrip(32).Visible = True
        cboMonedaContable.Visible = True
        
        If cboMonedaContable.ListIndex > 0 Then
            intRegistro = ObtenerItemLista(arrMonedaContable(), strCodMonedaContable)
            If intRegistro >= 0 Then cboMonedaContable.ListIndex = intRegistro
        Else
            intRegistro = ObtenerItemLista(arrMonedaContable(), Codigo_Moneda_Local)
            If intRegistro >= 0 Then cboMonedaContable.ListIndex = intRegistro
        End If
        
    Else
        strIndSoloMovimientoContable = Valor_Caracter
        txtMontoMovimiento.Enabled = True
        txtMontoContable.Enabled = False
    
        lblDescrip(32).Visible = False
        cboMonedaContable.Visible = False
        cboMonedaContable.ListIndex = 0
    
    End If

End Sub

Private Sub chkSimulacion_Click()

    If chkSimulacion.Value = vbChecked Then
        strTipoProceso = "1"
    End If

    If chkSimulacion.Value = vbUnchecked Then
        strTipoProceso = "0"
    End If


End Sub

Private Sub cmdActualizar_Click()

    Dim strCriterio As String
    
    'VALIDAR QUE EXISTA REGISTRO
    If adoRegistroAux.RecordCount = 0 Then
        MsgBox "No puede editar un movimiento si no existen registros en el detalle del asiento!", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If adoRegistroAux.EOF Then
        MsgBox "Debe seleccionar un movimiento para editar!", vbInformation, Me.Caption
        Exit Sub
    End If
    
    adoRegistroAux.Fields("NumAsiento") = lblNumAsiento.Caption
    'adoRegistroAux.Fields("SecMovimiento") = 0 'numSecMovimiento
    adoRegistroAux.Fields("DescripMovimiento") = txtDescripMovimiento.Text
    adoRegistroAux.Fields("CodCuenta") = Trim(txtCodCuenta.Text)
    adoRegistroAux.Fields("CodFile") = Trim(txtCodFile.Text)
    adoRegistroAux.Fields("CodAnalitica") = Trim(txtCodAnalitica.Text)
    adoRegistroAux.Fields("DescripAnalitica") = Trim(txtCodFile.Text) + "-" + Trim(txtCodAnalitica.Text)
    adoRegistroAux.Fields("IndDebeHaber") = strCodNaturaleza
    adoRegistroAux.Fields("CodMonedaMovimiento") = strCodMonedaMovimiento
    adoRegistroAux.Fields("CodSignoMoneda") = ObtenerCodSignoMoneda(strCodMonedaMovimiento)
    
    If strCodNaturaleza = "D" Then
        adoRegistroAux.Fields("MontoDebe") = CDbl(Trim(txtMontoMovimiento.Text))
        adoRegistroAux.Fields("MontoHaber") = 0
    Else
        adoRegistroAux.Fields("MontoHaber") = CDbl(Trim(txtMontoMovimiento.Text))
        adoRegistroAux.Fields("MontoDebe") = 0
    End If
    
    
    adoRegistroAux.Fields("CodMonedaContable") = strCodMonedaContable
    
    adoRegistroAux.Fields("MontoContable") = CDbl(txtMontoContable.Value)
    
    adoRegistroAux.Fields("ValorTipoCambio") = CDbl(txtTipoCambioMovimiento.Text)
    
'    adoRegistroAux.Fields("TipoAuxiliar") = strTipoAuxiliar
'    adoRegistroAux.Fields("CodAuxiliar") = strCodAuxiliar
    
    'nuevo
    adoRegistroAux.Fields("TipoDocumento") = strTipoDocumentoDet 'arrTipoDocumentoDet(cboTipoDocumentoDet.ListIndex)
    adoRegistroAux.Fields("NumDocumento") = txtNumDocumentoDet.Text
    
    If cboTipoPersonaContraparte.ListIndex <> -1 Then
        adoRegistroAux.Fields("TipoPersonaContraparte") = arrTipoPersonaContraparte(cboTipoPersonaContraparte.ListIndex)
    Else
        adoRegistroAux.Fields("TipoPersonaContraparte") = Valor_Caracter
    End If
    
    adoRegistroAux.Fields("CodPersonaContraparte") = strCodPersonaContraparte
    adoRegistroAux.Fields("DescripPersonaContraparte") = strDescripPersonaContraparte
    
    adoRegistroAux.Fields("IndSoloMovimientoContable") = strIndSoloMovimientoContable

    If chkContracuenta.Value = vbChecked Then
        adoRegistroAux.Fields("IndContracuenta") = Valor_Indicador
        
        adoRegistroAux.Fields("CodContracuenta") = strCodContracuenta
        adoRegistroAux.Fields("CodFileContracuenta") = strCodFileContracuenta
        adoRegistroAux.Fields("CodAnaliticaContracuenta") = strCodAnaliticaContracuenta
        adoRegistroAux.Fields("DescripContracuenta") = strDescripContracuenta
        adoRegistroAux.Fields("DescripFileAnaliticaContracuenta") = strDescripFileAnaliticaContracuenta
    Else
        adoRegistroAux.Fields("IndContracuenta") = Valor_Caracter
        
        adoRegistroAux.Fields("CodContracuenta") = Valor_Caracter
        adoRegistroAux.Fields("CodFileContracuenta") = Valor_Caracter
        adoRegistroAux.Fields("CodAnaliticaContracuenta") = Valor_Caracter
        adoRegistroAux.Fields("DescripContracuenta") = Valor_Caracter
        adoRegistroAux.Fields("DescripFileAnaliticaContracuenta") = Valor_Caracter
    End If
    
'    adoRegistroAux.Fields("IndUltimoMovimiento") = strIndUltimoMoviiento
'    adoRegistroAux.Fields("IndSoloMovimientoContable") = strIndSoloMovimientoContable
    
    
    If adoRegistroAux.Fields("CodMonedaMovimiento") <> Codigo_Moneda_Local Then
        strCriterio = "CodMonedaOrigen='" & Mid(strCodMonedaParPorDefecto, 1, 2) & "'" & _
                      " AND CodMonedaCambio = '" & Mid(strCodMonedaParPorDefecto, 3, 2) & "'" & _
                      " AND ValorTipoCambio = " & CDbl(txtTipoCambioMovimiento.Text)
        If Not FindRecordset(adoRegistroAuxTC, strCriterio) Then
            adoRegistroAuxTC.AddNew
            adoRegistroAuxTC.Fields("CodMonedaOrigen") = Mid(strCodMonedaParPorDefecto, 1, 2) 'strCodMoneda
            adoRegistroAuxTC.Fields("CodMonedaCambio") = Mid(strCodMonedaParPorDefecto, 3, 2) 'strCodMonedaCuenta
            adoRegistroAuxTC.Fields("ValorTipoCambio") = CDbl(txtTipoCambioMovimiento.Text)
        End If
    End If
    
    Call TotalizarMovimientos

End Sub

Private Sub cmdAgregar_Click()

    Dim adoRegistro As ADODB.Recordset
    Dim intSecuencial As Integer
    Dim dblBookmark As Double
    Dim strCriterio As String
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    If strEstado = Reg_Adicion Or strEstado = Reg_Edicion Then
        If TodoOkMovimiento() Then
           
            adoRegistroAux.AddNew
            adoRegistroAux.Fields("NumAsiento") = lblNumAsiento.Caption
            adoRegistroAux.Fields("SecMovimiento") = 0 'numSecMovimiento
            adoRegistroAux.Fields("DescripMovimiento") = txtDescripMovimiento.Text
            adoRegistroAux.Fields("CodCuenta") = Trim(txtCodCuenta.Text)
            adoRegistroAux.Fields("CodFile") = Trim(txtCodFile.Text)
            adoRegistroAux.Fields("CodAnalitica") = Trim(txtCodAnalitica.Text)
            adoRegistroAux.Fields("DescripAnalitica") = Trim(txtCodFile.Text) + "-" + Trim(txtCodAnalitica.Text)
            adoRegistroAux.Fields("IndDebeHaber") = strCodNaturaleza
            adoRegistroAux.Fields("CodMonedaMovimiento") = strCodMonedaMovimiento
            adoRegistroAux.Fields("CodSignoMoneda") = ObtenerCodSignoMoneda(strCodMonedaMovimiento)
            
            If strCodNaturaleza = "D" Then
                adoRegistroAux.Fields("MontoDebe") = CDbl(Trim(txtMontoMovimiento.Text))
                adoRegistroAux.Fields("MontoHaber") = 0
            Else
                adoRegistroAux.Fields("MontoHaber") = CDbl(Trim(txtMontoMovimiento.Text))
                adoRegistroAux.Fields("MontoDebe") = 0
            End If
            
            adoRegistroAux.Fields("CodMonedaContable") = strCodMonedaContable
            
            adoRegistroAux.Fields("MontoContable") = CDbl(txtMontoContable.Value)
            
            adoRegistroAux.Fields("ValorTipoCambio") = CDbl(txtTipoCambioMovimiento.Text)
            
'            adoRegistroAux.Fields("TipoAuxiliar") = strTipoAuxiliar
'            adoRegistroAux.Fields("CodAuxiliar") = strCodAuxiliar
            
            '**************BMM NUEVOS CAMBIOS *******
            adoRegistroAux.Fields("TipoDocumento") = strTipoDocumentoDet
            adoRegistroAux.Fields("NumDocumento") = txtNumDocumentoDet.Text
            
            If cboTipoPersonaContraparte.ListIndex <> -1 Then
                adoRegistroAux.Fields("TipoPersonaContraparte") = arrTipoPersonaContraparte(cboTipoPersonaContraparte.ListIndex)
            Else
                adoRegistroAux.Fields("TipoPersonaContraparte") = Valor_Caracter
            End If
            
            adoRegistroAux.Fields("CodPersonaContraparte") = strCodPersonaContraparte
            adoRegistroAux.Fields("DescripPersonaContraparte") = strDescripPersonaContraparte
            
            adoRegistroAux.Fields("IndSoloMovimientoContable") = strIndSoloMovimientoContable
            
            If chkContracuenta.Value = vbChecked Then
                adoRegistroAux.Fields("CodContracuenta") = strCodContracuenta
                adoRegistroAux.Fields("CodFileContracuenta") = strCodFileContracuenta
                adoRegistroAux.Fields("CodAnaliticaContracuenta") = strCodAnaliticaContracuenta
                adoRegistroAux.Fields("DescripContracuenta") = strDescripContracuenta
                adoRegistroAux.Fields("DescripFileAnaliticaContracuenta") = strDescripFileAnaliticaContracuenta
                adoRegistroAux.Fields("IndContracuenta") = Valor_Indicador
            Else
                adoRegistroAux.Fields("CodContracuenta") = Valor_Caracter
                adoRegistroAux.Fields("CodFileContracuenta") = Valor_Caracter
                adoRegistroAux.Fields("CodAnaliticaContracuenta") = Valor_Caracter
                adoRegistroAux.Fields("DescripContracuenta") = Valor_Caracter
                adoRegistroAux.Fields("DescripFileAnaliticaContracuenta") = Valor_Caracter
                adoRegistroAux.Fields("IndContracuenta") = Valor_Caracter
            End If
            
            
            If adoRegistroAux.Fields("CodMonedaMovimiento") <> Codigo_Moneda_Local Then
                strCriterio = "CodMonedaOrigen='" & Mid(strCodMonedaParPorDefecto, 1, 2) & "'" & _
                              " AND CodMonedaCambio = '" & Mid(strCodMonedaParPorDefecto, 3, 2) & "'" & _
                              " AND ValorTipoCambio = " & CDbl(txtTipoCambioMovimiento.Text)
                If Not FindRecordset(adoRegistroAuxTC, strCriterio) Then
                    adoRegistroAuxTC.AddNew
                    adoRegistroAuxTC.Fields("CodMonedaOrigen") = Mid(strCodMonedaParPorDefecto, 1, 2) 'strCodMoneda
                    adoRegistroAuxTC.Fields("CodMonedaCambio") = Mid(strCodMonedaParPorDefecto, 3, 2) 'strCodMonedaCuenta
                    adoRegistroAuxTC.Fields("ValorTipoCambio") = CDbl(txtTipoCambioMovimiento.Text)
                End If
            End If
            
            adoRegistroAux.Update
            
            dblBookmark = adoRegistroAux.Bookmark
            
            tdgMovimiento.Refresh
            
            Call NumerarRegistros
            
            adoRegistroAux.Bookmark = dblBookmark
            
            Call TotalizarMovimientos
            
            adoRegistroAux.Bookmark = dblBookmark
            
            cmdQuitar.Enabled = True
            
            Call LimpiarDatosMovimiento
        
        End If
    End If
    
End Sub
Private Sub NumerarRegistros()

    Dim n As Long
    
    n = 1
    
    If Not adoRegistroAux.EOF And Not adoRegistroAux.BOF Then
        adoRegistroAux.MoveFirst
    End If
    
    While Not adoRegistroAux.EOF
        adoRegistroAux.Fields("SecMovimiento") = n
        adoRegistroAux.Update
        n = n + 1
        adoRegistroAux.MoveNext
    Wend


End Sub

Private Sub LimpiarDatosMovimiento()
    
    txtDescripCuenta.Text = Valor_Caracter
    txtCodCuenta.Text = Valor_Caracter
    txtCodFile.Text = Valor_Caracter
    txtCodAnalitica.Text = Valor_Caracter
    txtMontoMovimiento.Text = "0"
    txtMontoContable.Text = "0.00"
    txtTipoCambioMovimiento.Text = "0"
    txtDescripMovimiento.Text = Valor_Caracter
    
    cboMonedaMovimiento.ListIndex = -1
    cboNaturaleza.ListIndex = -1
        
End Sub

Private Sub cmdBusqueda_Click(Index As Integer)

   Dim sSql As String
   
   
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
        
        strTipoFile = "04"
        
        Select Case Index
        
            Case 0
            
                
                frmBus.Caption = " Relación de Cuentas Contables"
                .sSql = "SELECT CodCuenta,DescripCuenta,TipoFile,IndAuxiliar,TipoAuxiliar FROM PlanContable "
                .sSql = .sSql & " WHERE IndMovimiento='" & Valor_Indicador & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND NumVersion = dbo.uf_CNObtenerPlanContableVigente('" & gstrCodAdministradora & "') ORDER BY CodCuenta"
                .OutputColumns = "1,2,3,4,5"
                .HiddenColumns = "3,4,5"
                
            Case 1
        
                frmBus.Caption = " Relación de File Analiticas"
                .sSql = "{ call up_CNSelFileAnalitico('" & strCodFondo & "','" & gstrCodAdministradora & "','" & strCodCuenta & "','" & strTipoFile & "') }"
                .OutputColumns = "1,2,3,4,5"
                .HiddenColumns = ""
                
            
            Case 2
        
'                If cboTipoAuxiliar.ListIndex = -1 Then
'                    MsgBox "Seleccione primero el Tipo de Auxiliar!", vbInformation + vbOKOnly, Me.Caption
'                    Exit Sub
'                End If
                
'                frmBus.Caption = " Relación de Auxiliares Contables"
'                .sSql = "{ call up_CNSelAuxiliarContable('" & strTipoAuxiliar & "') }"
'                .OutputColumns = "1,2,3"
'                .HiddenColumns = "3"

                frmBus.Caption = " Relación de Auxiliares"
                .sSql = "{ call up_CNSelAuxiliarContable('" & strCodFondo & "','" & gstrCodAdministradora & "','" & strCodCuenta & "','" & "03" & "') }"
                .OutputColumns = "1,2,3,4,5"
                .HiddenColumns = ""
                
            Case 3
            
                frmBus.Caption = " Relacion de Personas"
                
                '** OBTENGO EL TIPO DEL COMBO SELECCIONADO **
                strTipoPersonaContraparte = arrTipoPersonaContraparte(cboTipoPersonaContraparte.ListIndex)
                                
                .sSql = "SELECT CodPersona,DescripPersona FROM InstitucionPersona " & _
                        "WHERE TipoPersona='" + strTipoPersonaContraparte + "'"
                .OutputColumns = "1,2"
                '.HiddenColumns = "1"
                'ME QUEDE ACA
        
        End Select
                
        Screen.MousePointer = vbHourglass
                
        .BuscarTabla
        
        Screen.MousePointer = vbNormal
        frmBus.Show 1
       
        If .iParams.Count = 0 Then Exit Sub
        
        If .iParams(1).Valor <> "" Then
            
        
            Select Case Index
            
                Case 0
                
'                    strCodCustodio = .iParams(1).Valor  '.sCodigo
'                    txtDescripCustodio.Text = .iParams(2).Valor '.sDescripcion
                    strTipoFile = Trim(.iParams(3).Valor)
                    strIndAuxiliar = Trim(.iParams(4).Valor)
                    strTipoAuxiliar = Trim(.iParams(5).Valor)
                    
                    strCodCuenta = Trim(.iParams(1).Valor)
                    strDescripCuenta = Trim(.iParams(2).Valor)
                    
                    txtCodCuenta.Text = strCodCuenta
                    
                    txtDescripCuenta.Text = strCodCuenta & " - " & strDescripCuenta
                    
                    txtDescripFileAnalitica.Text = ""
                    strCodFile = ""
                    strCodAnalitica = ""
                    strDescripFileAnalitica = ""
                    
                    txtDescripAuxiliar.Text = strCodFile
                    txtCodFile.Text = strCodAnalitica
                    txtCodAnalitica.Text = strDescripFileAnalitica
                    
                    txtDescripMovimiento = strDescripCuenta
                    
                    'txtCodFile.Text = ""
                    'txtCodAnalitica.Text = ""
                    
                    If strIndAuxiliar = Valor_Indicador Then
                    'este se cambia y comenta
                        'cmdBusqueda(2).Enabled = True
                        cmdBusqueda(2).Enabled = True
'                        If strTipoAuxiliar = "00" Then 'Todos
'                            cboTipoAuxiliar.ListIndex = -1
'                            cboTipoAuxiliar.Locked = False
'                        Else
'                            intRegistro = ObtenerItemLista(arrTipoAuxiliar(), strTipoAuxiliar)
'                            If intRegistro >= 0 Then cboTipoAuxiliar.ListIndex = intRegistro
'                            cboTipoAuxiliar.Locked = True
'                        End If
                    Else
                        'este se cambia y comenta
                        'cmdBusqueda(2).Enabled = False
                        cmdBusqueda(2).Enabled = False
                        txtDescripAuxiliar.Text = ""
                        strTipoAuxiliar = ""
                        strCodAuxiliar = ""
                    End If
                    
                    If strTipoFile = Valor_Caracter Then
                        'esto se cambia y comenta
                        'cmdBusqueda(1).Enabled = False
                        cmdBusqueda(1).Enabled = False
                    Else
                        'esto se cambia y comenta
                        'cmdBusqueda(1).Enabled = True
                         cmdBusqueda(1).Enabled = True
'                        intRegistro = ObtenerItemLista(arrTipoFile(), strTipoFile)
'                        If intRegistro >= 0 Then cboTipoFile.ListIndex = intRegistro
                    End If
                    
                Case 1
            
                    strCodFile = Trim(.iParams(1).Valor)
                    strCodAnalitica = Trim(.iParams(2).Valor)
                    strDescripFileAnalitica = Trim(.iParams(3).Valor)
                    strCodMoneda = Trim(.iParams(4).Valor)
                        
                    txtCodFile.Text = strCodFile
                    txtCodAnalitica.Text = strCodAnalitica
                
                    If strTipoFile = Valor_File_Generico Then
                        txtCodAnalitica.Enabled = True
                        txtDescripFileAnalitica.Text = "Analítica Genérica"
                    Else
                        txtDescripFileAnalitica.Text = strCodFile & "-" & strCodAnalitica & " - " & strDescripFileAnalitica
                        txtCodAnalitica.Enabled = True 'False
                    End If
                                        
                    cboMonedaMovimiento.ListIndex = -1
                    intRegistro = ObtenerItemLista(arrMonedaMovimiento(), strCodMoneda)
                    If intRegistro >= 0 Then cboMonedaMovimiento.ListIndex = intRegistro
                
                Case 2
            
                     strCodAuxiliar = Trim(.iParams(1).Valor)
                     strDescripAuxiliar = Trim(.iParams(2).Valor)
                     
                     txtDescripAuxiliar.Text = strDescripAuxiliar
                     
                     strCodMoneda = Trim(.iParams(3).Valor)
                     strCodFile = Trim(.iParams(4).Valor)
                     strCodAnalitica = Trim(.iParams(5).Valor)
                        
                     txtCodFile.Text = strCodFile
                     txtCodAnalitica.Text = strCodAnalitica
                     txtDescripAuxiliar.Text = strCodAuxiliar & " - " & strDescripAuxiliar
                     
                Case 3
                        
                     strCodPersonaContraparte = Trim(.iParams(1).Valor)
                     txtPersonaContraparte.Text = Trim(.iParams(2).Valor)
                     strDescripPersonaContraparte = Trim(.iParams(2).Valor)
            
            End Select
        
        End If
            
       
    End With
    
    Set frmBus = Nothing

End Sub

Private Sub cmdContracuenta_Click()

    Dim frmContracuenta As frmContracuenta
    
    Set frmContracuenta = New frmContracuenta
    
    frmContracuenta.strCodFondo = strCodFondo
    frmContracuenta.strTipoFileContracuenta = strTipoFileContracuenta
    frmContracuenta.strCodContracuenta = strCodContracuenta
    frmContracuenta.strCodFileContracuenta = strCodFileContracuenta
    frmContracuenta.strCodAnaliticaContracuenta = strCodAnaliticaContracuenta
    frmContracuenta.strDescripContracuenta = strDescripContracuenta 'REVISAR
    frmContracuenta.strDescripFileAnaliticaContracuenta = strDescripFileAnaliticaContracuenta 'REVISAR
    
    
    frmContracuenta.Show 1
    
    If frmContracuenta.blnOK Then
        strTipoFileContracuenta = frmContracuenta.strTipoFileContracuenta
        strCodContracuenta = frmContracuenta.strCodContracuenta
        strCodFileContracuenta = frmContracuenta.strCodFileContracuenta
        strCodAnaliticaContracuenta = frmContracuenta.strCodAnaliticaContracuenta
        strDescripContracuenta = frmContracuenta.strDescripContracuenta
        strDescripFileAnaliticaContracuenta = frmContracuenta.strDescripFileAnaliticaContracuenta
        lblContracuenta.Caption = frmContracuenta.strCodContracuenta + " / " + frmContracuenta.strCodFileAnaliticaContracuenta
    Else
        lblContracuenta.Caption = strCodContracuenta + " / " + strCodFileContracuenta + "-" + strCodAnaliticaContracuenta
    End If
    
    Set frmContracuenta = Nothing


End Sub

Private Sub cmdQuitar_Click()

    Dim dblBookmark As Double
    Dim strCriterio As String
    
    If adoRegistroAux.RecordCount > 0 Then
    
        dblBookmark = adoRegistroAux.Bookmark
    
        If adoRegistroAux.Fields("CodMonedaMovimiento") <> Codigo_Moneda_Local Then
            strCriterio = "CodMonedaOrigen='" & Mid(strCodMonedaParPorDefecto, 1, 2) & "'" & _
                          " AND CodMonedaCambio = '" & Mid(strCodMonedaParPorDefecto, 3, 2) & "'" & _
                          " AND ValorTipoCambio = " & CDbl(txtTipoCambioMovimiento.Text)
            If FindRecordset(adoRegistroAuxTC, strCriterio) Then
                adoRegistroAuxTC.Delete adAffectCurrent
            End If
        End If
    
        adoRegistroAux.Delete adAffectCurrent
        
        If adoRegistroAux.EOF Then
            adoRegistroAux.MovePrevious
            tdgMovimiento.MovePrevious
        End If
            
        adoRegistroAux.Update
        
        If adoRegistroAux.RecordCount = 0 Then cmdQuitar.Enabled = False

        If adoRegistroAux.RecordCount > 0 And Not adoRegistroAux.BOF And Not adoRegistroAux.EOF And dblBookmark > 1 Then adoRegistroAux.Bookmark = dblBookmark - 1
        
        Call NumerarRegistros
        
        If adoRegistroAux.RecordCount > 0 And Not adoRegistroAux.BOF And Not adoRegistroAux.EOF Then adoRegistroAux.Bookmark = dblBookmark - 1
        
        Call TotalizarMovimientos
        
        If adoRegistroAux.RecordCount > 0 And Not adoRegistroAux.BOF And Not adoRegistroAux.EOF Then adoRegistroAux.Bookmark = dblBookmark - 1
   
        tdgMovimiento.Refresh
    
    End If
    
    
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
Private Sub CargarListas()

    Dim strSQL  As String
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    '*** Empresas Módulos del Sistema ***
    strSQL = "SELECT CodModulo CODIGO, DescripModulo DESCRIP FROM ModuloSistema WHERE EstadoModulo='01' ORDER BY DescripModulo"
    CargarControlLista strSQL, cboModulo, arrModulo(), Valor_Caracter

    '*** Moneda Asiento y del Movimiento ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Valor_Caracter
    CargarControlLista strSQL, cboMonedaMovimiento, arrMonedaMovimiento(), Valor_Caracter
    
    '*** Naturaleza ***
    strSQL = "SELECT ValorParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='NATCTA' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboNaturaleza, arrNaturaleza(), Valor_Caracter

    '*** Tipo de File ***
'    strSQL = "SELECT ValorParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='TIPFIL' ORDER BY DescripParametro"
'    CargarControlLista strSQL, cboTipoFile, arrTipoFile(), Valor_Caracter

    '*** Tipo de Auxiliar ***
'    strSQL = "SELECT ValorParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='TIPAUX' ORDER BY DescripParametro"
'    CargarControlLista strSQL, cboTipoAuxiliar, arrTipoAuxiliar(), Valor_Caracter

    '*** Tipo de Comprobante Sunat ***
    strSQL = "SELECT CodTipoComprobantePago CODIGO,DescripTipoComprobantePago DESCRIP From TipoComprobantePago ORDER BY DescripTipoComprobantePago"
    CargarControlLista strSQL, cboTipoDocumento, arrTipoDocumento(), Sel_Defecto

    '***BMM NUEVOS CAMBIOS***
    
    '*** Tipo de Referencia de Documento ***
'    strSql = "SELECT ValorParametro AS CODIGO,DescripParametro AS DESCRIP " & _
'            "FROM AuxiliarParametro WHERE CodTipoParametro='TIPREF'"
'    CargarControlLista strSql, cboTipoRef, arrTipoRef(), Sel_Defecto
    
    '*** Tipo de Persona ***'
    strSQL = "SELECT CodParametro AS CODIGO,DescripParametro AS DESCRIP FROM AuxiliarParametro " & _
                "WHERE CodTipoParametro='TIPPER' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoPersonaContraparte, arrTipoPersonaContraparte(), Sel_Defecto
    
    '*** Tipo de Comprobante Sunat para Detalle ***
    
    strSQL = "SELECT CodTipoComprobantePago CODIGO,DescripTipoComprobantePago DESCRIP From TipoComprobantePago ORDER BY DescripTipoComprobantePago"
    CargarControlLista strSQL, cboTipoDocumentoDet, arrTipoDocumentoDet(), Sel_Defecto
      
    '*************************

    '*** Digitador ***
    
    '*** Verificador ***
        
End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    tabAsiento.Tab = 0
    chkSimulacion.Value = vbUnchecked
    strEstado = Reg_Defecto
    
    Call InicializarVariables
    
    Call ConfiguraRecordsetAuxiliar
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 9
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 40
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 8
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 18
               
   ' MsgBox gstrTempPath & "Layout.grx"
            
    tdgMovimiento.LayoutFileName = gstrTempPath & "Layout.grx"
    'tdgMovimiento.Layouts.Add "TestLayout"
            
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub
Private Sub InicializarVariables()

   
    numSecMovimiento = 1
    strTipoProceso = "0"

    strCodCuenta = Valor_Caracter
    strCodFile = Valor_Caracter
    strCodAnalitica = Valor_Caracter
    strTipoFile = Valor_Caracter
    strDescripCuenta = Valor_Caracter
    strDescripFileAnalitica = Valor_Caracter
    strCodMonedaContable = Codigo_Moneda_Local
    
    strTipoPersonaContraparte = Valor_Caracter
    strCodPersonaContraparte = Valor_Caracter
    strDescripPersonaContraparte = Valor_Caracter
    
    strCodContracuenta = Valor_Caracter
    strCodFileContracuenta = Valor_Caracter
    strCodAnaliticaContracuenta = Valor_Caracter
    strDescripFileAnaliticaContracuenta = Valor_Caracter
    strTipoFileContracuenta = Valor_Caracter

End Sub
Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    Set frmAsientoContable = Nothing
    
End Sub

Private Sub ExtornarComprobante(strpCodFondo As String, strpNumAsiento As String)

'    Dim strSQL As String, adoresultaux1 As New Recordset, adoresultaux4 As New Recordset
'    Dim WComCon As RCabasicon, WMovCon() As RDetasicon
'    Dim Sec_com As Long, Res As Integer, s_OldCom$, NewNroCom$, nSec%
'    Dim s_DscNewCom$
'
'    On Error GoTo Ctrl_Error
'
'    '* (2) Extorno Contable
'    '** Leer comprobante origen
'    With adoComm
'        .CommandText = "SELECT * FROM FMCOMPRO WHERE COD_FOND='" & s_ParCodFon$ & "' AND NRO_COMP='" & s_ParNroCom$ & "' "
'        Set adoresultaux4 = .Execute
'
'        If Not adoresultaux4.EOF Then
'                .CommandText = "SELECT FLG_CIER FROM FMCUOTAS WHERE COD_FOND='" & strCodFond & "' AND "
'                .CommandText = .CommandText & "FCH_CUOT='" & adoresultaux4("FCH_COMP") & "'"
'                Set adoresultaux1 = .Execute
'
'                If Not adoresultaux1.EOF Then
'                    If adoresultaux1("FLG_CIER") = "X" Then
'                        MsgBox "El comprobante contable no puede ser extornado por haber sido mayorizado en un cierre contable.", vbCritical, Me.Caption
'                        adoresultaux1.Close: Set adoresultaux1 = Nothing
'                        gblnRollBack = True
'                        Exit Sub
'                    End If
'                End If
'                adoresultaux1.Close: Set adoresultaux1 = Nothing
'        End If
'
'        Sec_com = 1
'        .CommandText = "SELECT NRO_ULTI_SOLI FROM FMPARAME WHERE COD_FOND='" & s_ParCodFon$ & "' AND COD_PARA='COM'"
'        Set adoresultaux1 = .Execute
'
'        If Not IsNull(adoresultaux1("NRO_ULTI_SOLI")) Then Sec_com = CLng(adoresultaux1("NRO_ULTI_SOLI")) + 1
'        adoresultaux1.Close: Set adoresultaux1 = Nothing
'
'    End With
'
'    WComCon.COD_FOND = adoresultaux4("COD_FOND")
'    WComCon.COD_MONC = adoresultaux4("COD_MONC")
'    WComCon.CNT_MOVI = adoresultaux4("CNT_MOVI")
'    WComCon.COD_MONE = adoresultaux4("COD_MONE")
'    WComCon.FLG_AUTO = adoresultaux4("FLG_AUTO")
'    WComCon.FLG_CONT = adoresultaux4("FLG_CONT")
'    WComCon.GEN_COMP = adoresultaux4("GEN_COMP")
'    WComCon.HOR_COMP = adoresultaux4("HOR_COMP")
'    WComCon.NRO_DOCU = ""
'    WComCon.NRO_OPER = adoresultaux4("NRO_OPER")
'    WComCon.PER_DIGI = gstrLogin
'    WComCon.PER_REVI = ""
'    WComCon.STA_COMP = ""
'    WComCon.SUB_SIST = "C"
'    WComCon.TIP_CAMB = adoresultaux4("TIP_CAMB")
'    WComCon.TIP_COMP = ""
'    WComCon.TIP_DOCU = ""
'    '** Variable
'    WComCon.VAL_COMP = adoresultaux4("VAL_COMP")
'    WComCon.NRO_COMP = Format(Sec_com, "00000000")
'    WComCon.FCH_COMP = FmtFec(gstrFechaAct, "win", "yyyymmdd", Res)
'    WComCon.FCH_CONT = FmtFec(gstrFechaAct, "win", "yyyymmdd", Res)
'    WComCon.MES_CONT = Mid$(WComCon.FCH_COMP, 5, 2)
'    WComCon.prd_cont = Mid$(WComCon.FCH_COMP, 1, 4)
'    WComCon.DSL_COMP = "Ext Comprobante (" & adoresultaux4("NRO_COMP") & ">>" & adoresultaux4("FCH_COMP") & ")"
'    WComCon.GLO_COMP = WComCon.DSL_COMP
'    s_OldCom = WComCon.NRO_COMP
'    NewNroCom = Format(Sec_com, "00000000")
'    WComCon.NRO_COMP = NewNroCom
'    s_DscNewCom$ = "(Ext)" & adoresultaux4("DSL_COMP")
'    adoresultaux4.Close
'
'    adoresultaux4.CursorLocation = adUseClient
'    adoresultaux4.CursorType = adOpenStatic
'
'    strSQL = "SELECT * FROM FMMOVCON WHERE COD_FOND='" & s_ParCodFon$ & "' AND NRO_COMP='" & s_ParNroCom$ & "' "
'    adoComm.CommandText = strSQL
'    adoresultaux4.Open adoComm.CommandText, adoConn, , , adCmdText
'    'Set adoresultaux4 = adoComm.Execute
'
'    If adoresultaux4.EOF Then
'        MsgBox "El Sistema no puede encontrar el comprobante contable.", vbExclamation
'        Exit Sub
'    Else
'        adoresultaux4.MoveLast
'        ReDim WMovCon(adoresultaux4.RecordCount)
'        adoresultaux4.MoveFirst
'    End If
'    nSec = 0
'    Do While Not adoresultaux4.EOF
'       nSec = nSec + 1
'       LIniRDetAsiCon WMovCon(nSec)
'       WMovCon(nSec).SEC_MOVI = adoresultaux4("SEC_MOVI")
'       WMovCon(nSec).COD_FOND = adoresultaux4("COD_FOND")
'       WMovCon(nSec).COD_MONE = adoresultaux4("COD_MONE")
'       WMovCon(nSec).FLG_PROC = "X"
'       WMovCon(nSec).STA_MOVI = "X"
'       WMovCon(nSec).TIP_GENR = "P"
'       WMovCon(nSec).CTA_AMAR = ""
'       WMovCon(nSec).CTA_AUTO = ""
'       WMovCon(nSec).CTA_ORIG = ""
'       WMovCon(nSec).COD_FILE = adoresultaux4("COD_FILE")
'       WMovCon(nSec).COD_ANAL = adoresultaux4("COD_ANAL")
'       WMovCon(nSec).FCH_MOVI = FmtFec(gstrFechaAct, "win", "yyyymmdd", Res)
'       WMovCon(nSec).prd_cont = Mid$(WMovCon(nSec).FCH_MOVI, 1, 4)
'       WMovCon(nSec).MES_COMP = Mid$(WMovCon(nSec).FCH_MOVI, 5, 2)
'       WMovCon(nSec).NRO_COMP = WComCon.NRO_COMP
'       WMovCon(nSec).FLG_DEHA = IIf(adoresultaux4("FLG_DEHA") = "D", "H", "D")
'       WMovCon(nSec).COD_CTA = adoresultaux4("COD_CTA")
'       WMovCon(nSec).DSC_MOVI = "Extorno(" & adoresultaux4("DSC_MOVI") & ")"
'       WMovCon(nSec).COD_FILE = adoresultaux4("COD_FILE")
'       WMovCon(nSec).COD_ANAL = adoresultaux4("COD_ANAL")
'       WMovCon(nSec).VAL_MOVN = (adoresultaux4("VAL_MOVN") * -1)
'       WMovCon(nSec).VAL_MOVX = (adoresultaux4("VAL_MOVX") * -1)
'       WMovCon(nSec).VAL_CONT = (adoresultaux4("VAL_CONT") * -1)
'       adoresultaux4.MoveNext
'    Loop
'
'    WComCon.CNT_MOVI = nSec
'    LGraAsi WComCon, WMovCon() 'Grabar el asiento
'    Call UpdNewNro(s_ParCodFon$, "COM", NewNroCom)
'
''    strsql = "UPDATE fmCompro SET "
''    strsql = strsql & " DSL_COMP='" & s_DscNewCom$ & "',"
''    strsql = strsql & " GLO_COMP='" & s_DscNewCom$ & "'"
''    strsql = strsql & " WHERE COD_FOND='" & s_ParCodFon$ & "'"
''    strsql = strsql & " AND NRO_COMP='" & s_ParNroCom$ & "'"
''    adoConn.Execute strsql
'    Exit Sub
'
'Ctrl_Error:
'    gblnRollBack = True
'    MsgBox "Error " & Err.Number & " => " & Err.Description, vbCritical
'    Exit Sub
    
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


Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        Call Deshabilita
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabAsiento
            .TabEnabled(0) = False
            .Tab = 1
        End With
    End If
    
End Sub

Public Sub Eliminar()
    
    Dim strFechaGrabar As String
    Dim strNumAsiento As String
    Dim strPeriodoContable As String
    Dim strMesContable As String
    
    
    If strEstado <> Reg_Edicion Then
        If strEstado <> Reg_Consulta Then Exit Sub
    End If

    strNumAsiento = tdgConsulta.Columns("NumAsiento").Value
'    strPeriodoContable = tdgConsulta.Columns("PeriodoContable").Value
'    strMesContable = tdgConsulta.Columns("MesContable").Value
        
    'Validar que no se pueda modificar un asiento de otra fecha
    If Convertyyyymmdd(gdatFechaActual) <> Convertyyyymmdd(tdgConsulta.Columns("FechaAsiento").Value) Then
        MsgBox "No se Puede Anular el Comprobante Contable Nro. " & strNumAsiento & " porque Corresponde a Otra Fecha!", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    
    
    If strEstado <> Reg_Edicion Then
        If strEstado <> Reg_Consulta Then Exit Sub
    End If

    If MsgBox("Desea Anular el comprobante contable " & tdgConsulta.Columns(0).Value & " ?", vbYesNo + vbQuestion + vbDefaultButton2) <> vbYes Then
        Exit Sub
    End If

    On Error GoTo Ctrl_Error

    Me.MousePointer = vbHourglass
                                        
    With adoComm
        
        .CommandType = adCmdText
        
        strFechaGrabar = Convertyyyymmdd(gdatFechaActual) & Space(1) & Format(Time, "hh:ss")
    
        strNumAsiento = tdgConsulta.Columns("NumAsiento").Value
        
        '*** Cabecera ***
        .CommandText = "{ call up_ACProcAsientoContableAnulacion('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "') }"
        adoConn.Execute .CommandText
       
'ASIENTO DE EXTORNO -- POR EL MOMENTO DESHABILITADO
'        .CommandText = "{ call up_ACProcAsientoContableExtorno('" & _
'            strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
'            strFechaGrabar & "') }"
'        adoConn.Execute .CommandText
       
        Me.MousePointer = vbDefault
       
        Call Buscar
        
        MsgBox Mensaje_Proceso_Exitoso, vbExclamation
            
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"

                                                                       
    End With
    
    Exit Sub
    
Ctrl_Error:
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
    
    
End Sub

Public Sub Grabar()
    
    Dim objAsientoContableDetalleXML    As DOMDocument60
    Dim objTipoCambioReemplazoXML       As DOMDocument60
    Dim strAsientoContableDetalleXML    As String
    Dim strTipoCambioReemplazoXML       As String
    Dim strMsgError                     As String
    
    strTipoCambioReemplazoXML = XML_TipoCambioReemplazo
    
    If strEstado = Reg_Adicion Or strEstado = Reg_Edicion Then
        If TodoOK() Then
            Dim intCantRegistros    As Integer, intRegistro         As Integer
            'Dim adoRegistro         As ADODB.Recordset
            Dim strNumAsiento       As String, strFechaGrabar       As String
            
            Me.MousePointer = vbHourglass
                                                
            With adoComm
                
                strFechaGrabar = Convertyyyymmdd(dtpFechaAsiento.Value) & Space(1) & Format(Time, "hh:ss")
                strNumAsiento = Trim(lblNumAsiento.Caption)
                
                If strCodMoneda = Codigo_Moneda_Local Then
                   txtMontoAsiento.Text = lblTotalDebeMN.Caption
                Else
                  txtMontoAsiento.Text = lblTotalDebeME.Caption
                End If
                
                If strIndSoloMovimientoContable = Valor_Caracter Then
                    strCodMonedaContable = Valor_Caracter
                End If
                
                'On Error GoTo Ctrl_Error
                
                Call XMLADORecordset(objAsientoContableDetalleXML, "AsientoContableDetalle", "Detalle", adoRegistroAux, strMsgError)
                strAsientoContableDetalleXML = objAsientoContableDetalleXML.xml 'CrearXMLDetalle(objTipoCambioReemplazoXML)
                
                Call XMLADORecordset(objTipoCambioReemplazoXML, "TipoCambioReemplazo", "MonedaTipoCambio", adoRegistroAuxTC, strMsgError)
                strTipoCambioReemplazoXML = objTipoCambioReemplazoXML.xml 'CrearXMLDetalle(objTipoCambioReemplazoXML)
                               
                
                '*** Cabecera ***
                .CommandText = "{ call up_ACManAsientoContableXML('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
                    strFechaGrabar & "','" & _
                    gstrPeriodoActual & "','" & gstrMesActual & "','" & Tipo_Asiento_Diario & _
                    "','" & _
                    Trim(txtDescripAsiento.Text) & "','" & strCodMoneda & "','" & strCodMonedaContable & _
                    "','" & _
                    strTipoDocumento & "','" & _
                    Trim(txtNumDocumento.Text) & "'," & _
                    CDec(txtMontoAsiento.Text) & ",'" & Estado_Activo & "'," & _
                    intCantRegistros & ",'" & _
                    strFechaGrabar & "','" & _
                    strCodModulo & "','" & _
                    "'," & _
                    CDec(txtTipoCambio.Text) & ",'" & gstrLogin & _
                    "','" & _
                    "','" & _
                    Trim(txtDescripAsiento.Text) & "','" & _
                    "','" & _
                    "X','','" & strAsientoContableDetalleXML & "','" & strTipoCambioReemplazoXML & "','" & _
                    IIf(strEstado = Reg_Adicion, "I", "U") & "') }"
                adoConn.Execute .CommandText
                
                                                                                
            End With
            
            'Set adoRegistroAux = Nothing
            'Set adoRegistroAuxTC = Nothing
                
            Me.MousePointer = vbDefault
        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabAsiento
                .TabEnabled(0) = True
                .Tab = 0
            End With
            
            Call Buscar
        End If
    End If
    Exit Sub
    
Ctrl_Error:
'    adoComm.CommandText = "ROLLBACK TRAN ProcAsiento"
'    adoConn.Execute adoComm.CommandText
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
            
End Sub

Public Sub Adicionar()

    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar asiento..."
    
    If cboFondo.ListIndex = -1 Then
        MsgBox "Debe seleccionar un Fondo!", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabAsiento
        .TabEnabled(0) = False
        .Tab = 1
    End With
    Call Habilita
    
End Sub

Private Sub LlenarFormulario(strModo As String)
        
    Dim intRegistro As Integer
    Dim strSQL As String
    
    Select Case strModo
        Case Reg_Adicion
            
            Call InicializarVariables
            
            lblNumAsiento.Caption = "GENERADO" 'NumAleatorio(10)
            txtDescripAsiento.Enabled = True
            txtDescripAsiento.Text = Valor_Caracter
            
            cboModulo.Enabled = True
            intRegistro = ObtenerItemLista(arrModulo(), frmMainMdi.Tag)
            If intRegistro >= 0 Then cboModulo.ListIndex = intRegistro
            
            dtpFechaAsiento.Value = gdatFechaActual
            txtHoraAsiento.Text = Format(Time, "hh:mm")
            txtMontoAsiento.Text = "0"
            txtTipoCambioMovimiento.Enabled = True
            
            cboMoneda.Enabled = True
            intRegistro = ObtenerItemLista(arrMoneda(), gstrCodMoneda)
            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
            
            chkAjuste.Value = vbUnchecked
            
            chkMovContable.Value = vbUnchecked
            
            Call chkMovContable_Click
            
            txtDescripCuenta.Text = Valor_Caracter
            txtCodCuenta.Text = Valor_Caracter
            txtCodFile.Text = Valor_Caracter
            txtCodAnalitica.Text = Valor_Caracter
            txtMontoMovimiento.Text = "0"
            txtMontoContable.Text = "0.00"
            txtTipoCambioMovimiento.Text = "0"
            txtDescripMovimiento.Text = Valor_Caracter

            intRegistro = ObtenerItemLista(arrNaturaleza(), Codigo_Tipo_Naturaleza_Debe)
            If intRegistro >= 0 Then cboNaturaleza.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrMonedaMovimiento(), gstrCodMoneda)
            If intRegistro >= 0 Then cboMonedaMovimiento.ListIndex = intRegistro
            
            cboTipoPersonaContraparte.ListIndex = -1
            cboTipoDocumentoDet.ListIndex = -1
            
            txtPersonaContraparte.Text = Valor_Caracter
            txtNumDocumentoDet.Text = Valor_Caracter
            
            cboNaturaleza.ListIndex = -1
            cboMonedaMovimiento.ListIndex = -1
            
            lblContracuenta.Caption = Valor_Caracter
            chkContracuenta.Value = vbUnchecked
            
            lblTotalDebeME.Caption = "0"
            lblTotalHaberME.Caption = "0"
            lblTotalDebeMN.Caption = "0"
            lblTotalHaberMN.Caption = "0"
            
            txtDescripAsiento.SetFocus
                        
            Call ConfiguraRecordsetAuxiliarTC
                        
            Call CargarDetalleGrilla
            
            Call TotalizarMovimientos
                        
                        
        Case Reg_Edicion
            
            Dim adoRecordset As New ADODB.Recordset
            
            cboMonedaMovimiento.ListIndex = -1
                        
            strSQL = "SELECT FechaAsiento, dbo.uf_ACObtenerHoraFecha(FechaAsiento) AS HoraAsiento, " & _
                     "TipoDocumento, NumDocumento, DescripAsiento, " & _
                     "MontoAsiento, CodMoneda, ValorTipoCambio " & _
                     "FROM AsientoContable AC " & _
                     "WHERE " & _
                     "AC.CodFondo = '" & strCodFondo & "' AND " & _
                     "AC.CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
                     "AC.NumAsiento = '" & tdgConsulta.Columns("NumAsiento") & "'"
                     
            adoComm.CommandText = strSQL
                     
            Set adoRecordset = adoComm.Execute

            If Not adoRecordset.EOF Then
            
                lblNumAsiento.Caption = tdgConsulta.Columns("NumAsiento")
                txtDescripAsiento.Text = adoRecordset.Fields("DescripAsiento") 'tdgConsulta.Columns(2)
                'txtDescripAsiento.Enabled = False
                
                strTipoDocumento = adoRecordset.Fields("TipoDocumento")
                strNumDocumento = adoRecordset.Fields("NumDocumento")
                
                txtNumDocumento.Text = strNumDocumento
                
                intRegistro = ObtenerItemLista(arrTipoDocumento(), strTipoDocumento)
                If intRegistro >= 0 Then cboTipoDocumento.ListIndex = intRegistro
                
                cboModulo.Enabled = False
                
                intRegistro = ObtenerItemLista(arrModulo(), frmMainMdi.Tag)
                If intRegistro >= 0 Then cboModulo.ListIndex = intRegistro
                
                dtpFechaAsiento.Value = adoRecordset.Fields("FechaAsiento")
                txtHoraAsiento.Text = adoRecordset.Fields("HoraAsiento")
                
                txtMontoAsiento.Text = adoRecordset.Fields("MontoAsiento")
                
                txtTipoCambioMovimiento.Text = adoRecordset.Fields("ValorTipoCambio")
                txtTipoCambioMovimiento.Enabled = True
                
                cboMoneda.Enabled = False
                
                intRegistro = ObtenerItemLista(arrMoneda(), adoRecordset.Fields("CodMoneda"))
                If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
            
'                intRegistro = ObtenerItemLista(arrMonedaContable(), adoRecordset.Fields("CodMonedaContable"))
'                If intRegistro >= 0 Then cboMonedaContable.ListIndex = intRegistro
            
                Call ConfiguraRecordsetAuxiliarTC
            
                Call CargarDetalleGrilla
                
                Call TotalizarMovimientos
            
                'adoRegistroAux.MoveFirst
            
            Else
                MsgBox "El Sistema no puede encontrar el comprobante contable para consultar!", vbExclamation
                Exit Sub
            End If
    
                                                            
        End Select
        
        
    
End Sub
'Private Sub CargarDetalleGrilla()
'
'    Dim strSQL As String
'    Dim adoRecordset As New ADODB.Recordset
'
'    strSQL = "SELECT FechaAsiento, TipoDocumento, NumDocumento, DescripAsiento, " & _
'             "NumAsiento, SecMovimiento, FechaMovimiento, PeriodoContable, " & _
'             "MesContable, DescripMovimiento, IndDebeHaber, CodCuenta," & _
'             "CodMoneda, MontoMovimientoMN, MontoMovimientoME, MontoContable," & _
'             "CodFile , CodAnalitica, TipoAuxiliar, CodAuxiliar, IndUltimoMovimiento " & _
'             "FROM AsientoContableDetalle ACD " & _
'             "JOIN AsientoContable AC ON " & _
'             "(AC.NumAsiento = ACD.NumAsiento AND AC.CodFondo = ACD.CodFondo AND AC.CodAdministradora = ACD.CodAdministradora) " & _
'             "WHERE " & _
'             "ACD.CodFondo = '" & strCodFondo & "' AND " & _
'             "ACD.CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
'             "ACD.NumAsiento = '" & strNumAsiento & "'"
'
'    Set adoRecordset = adoComm.Execute
'
'
'
'
'End Sub


Private Sub CargarDetalleGrilla()
    
    Dim adoRegistro As ADODB.Recordset
    Dim adoField As ADODB.Field
    
    Dim strSQL As String
    
    Set adoRegistro = New ADODB.Recordset
        
    Call ConfiguraRecordsetAuxiliar
    
    If strEstado = Reg_Edicion Then
        
        strSQL = "{ call up_CNLstAsientoContableDetalle ('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                            Trim(lblNumAsiento.Caption) & "','" & strTipoProceso & "')}"

        With adoRegistro
        'With adoMovimiento
            .ActiveConnection = gstrConnectConsulta
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open strSQL
        
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    adoRegistroAux.AddNew
                    For Each adoField In adoRegistroAux.Fields
                        adoRegistroAux.Fields(adoField.Name) = adoRegistro.Fields(adoField.Name)
                    Next
                    adoRegistroAux.Update
                    adoRegistro.MoveNext
                    'adoMovimiento.MoveNext
                Loop
                adoRegistroAux.MoveFirst
            End If
            
        End With
    
    End If
    
    tdgMovimiento.DataSource = adoRegistroAux
    
    'If adoRegistroAux.RecordCount > 0 Then strEstado = Reg_Consulta
            
End Sub


Private Function TodoOK() As Boolean

    TodoOK = False
    
    If Trim(txtDescripAsiento.Text) = Valor_Caracter Then
        MsgBox "Descripción de asiento no ingresada", vbCritical, gstrNombreEmpresa
        txtDescripAsiento.SetFocus
        Exit Function
    End If
            
    If CDbl(txtTipoCambioMovimiento.Text) = 0 Then
        MsgBox "Tipo de Cambio no ingresado", vbCritical, gstrNombreEmpresa
        txtTipoCambioMovimiento.SetFocus
        Exit Function
    End If
    
'    If cboTipoDocumento.ListIndex < 0 Then
'        MsgBox "Seleccione el tipo de documento", vbCritical, gstrNombreEmpresa
'        cboTipoDocumento.SetFocus
'        Exit Function
'    End If
    
'    If Len(Trim(txtNumDocumento.Text)) = 0 Then
'        MsgBox "Número de documento no ingresado", vbCritical, gstrNombreEmpresa
'        txtNumDocumento.SetFocus
'        Exit Function
'    End If
    
    If cboMoneda.ListIndex < 0 Then
        MsgBox "Seleccione la moneda", vbCritical, gstrNombreEmpresa
        cboMoneda.SetFocus
        Exit Function
    End If
    
    If adoRegistroAux.EOF And adoRegistroAux.BOF Then
        MsgBox "Comprobante sin registros", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
    
    '*** Validar cuadre del asiento y generar movimiento por diferencia ***
    If Not ValidaCuadreContable() Then
        MsgBox "Comprobante descuadrado", vbCritical, gstrNombreEmpresa
        Exit Function
    End If

    '*** Si todo paso OK ***
    TodoOK = True
  
End Function


Private Function TodoOkMovimiento() As Boolean

    TodoOkMovimiento = False
    
    If Trim(txtDescripAsiento.Text) = Valor_Caracter Then
        MsgBox "Descripción de asiento no ingresada.", vbCritical, gstrNombreEmpresa
        txtDescripAsiento.SetFocus
        Exit Function
    End If
                
    If CDbl(txtTipoCambioMovimiento.Text) = 0 Then
        MsgBox "Tipo de Cambio no ingresado.", vbCritical, gstrNombreEmpresa
        txtTipoCambioMovimiento.SetFocus
        Exit Function
    End If
    
    If cboMoneda.ListIndex < Valor_Numero Then
        MsgBox "Seleccione la Moneda del Asiento.", vbCritical, gstrNombreEmpresa
        cboMoneda.SetFocus
        Exit Function
    End If
    
    If Not ValidarCuentaContable(txtCodCuenta.Text, gstrCodAdministradora) Then
        MsgBox "La cuenta contable no existe.", vbCritical, gstrNombreEmpresa
        txtCodCuenta.SetFocus
        Exit Function
    End If
    
    If Trim(txtCodFile.Text) = Valor_Caracter Then
        MsgBox "Código de File no ingresado.", vbCritical, gstrNombreEmpresa
        txtCodFile.SetFocus
        Exit Function
    End If
    
    If Trim(txtCodAnalitica.Text) = Valor_Caracter Then
        MsgBox "Código de Analítica no ingresado.", vbCritical, gstrNombreEmpresa
        txtCodAnalitica.SetFocus
        Exit Function
    End If
    
    If cboMonedaMovimiento.ListIndex < Valor_Numero Then
        MsgBox "Seleccione la Moneda del Movimiento.", vbCritical, gstrNombreEmpresa
        cboMonedaMovimiento.SetFocus
        Exit Function
    End If
    
    If cboNaturaleza.ListIndex < Valor_Numero Then
        MsgBox "Seleccione la Naturaleza del Movimiento.", vbCritical, gstrNombreEmpresa
        cboNaturaleza.SetFocus
        Exit Function
    End If
    
    If CDbl(txtMontoMovimiento.Text) = 0 And strIndSoloMovimientoContable = Valor_Caracter Then
        MsgBox "Monto de Movimiento no ingresado.", vbCritical, gstrNombreEmpresa
        txtMontoMovimiento.SetFocus
        Exit Function
    End If
    
    If Trim(txtDescripMovimiento.Text) = Valor_Caracter Then
        MsgBox "Descripción del movimiento no ingresada.", vbCritical, gstrNombreEmpresa
        txtDescripMovimiento.SetFocus
        Exit Function
    End If
    
'ACR: este comentario es a pedido de Gino Elizagaray -- 07-10-2010
'    If Not ValidarFile(txtCodFile.Text) Then
'        MsgBox "File no existe o no está vigente...", vbCritical
'        txtCodCuenta.SetFocus
'        Exit Function
'    End If
    
'    If Not ValidarAnalitica(txtCodFile.Text, txtCodAnalitica.Text, strCodFondo) Then
'        MsgBox "Analitica no existe o no está vigente...", vbCritical
'        txtCodAnalitica.SetFocus
'        Exit Function
'    End If
'ACR: este comentario es a pedido de Gino Elizagaray -- 07-10-2010
        
    '*** Si todo pasó OK ***
    TodoOkMovimiento = True
  
End Function


Private Sub lblTotalDebeME_Change()

    Call FormatoMillarEtiqueta(lblTotalDebeME, Decimales_Monto)
    
End Sub

Private Sub lblTotalDebeMN_Change()

    Call FormatoMillarEtiqueta(lblTotalDebeMN, Decimales_Monto)
    
End Sub

Private Sub lblTotalHaberME_Change()

    Call FormatoMillarEtiqueta(lblTotalHaberME, Decimales_Monto)
    
End Sub

Private Sub lblTotalHaberMN_Change()

    Call FormatoMillarEtiqueta(lblTotalHaberMN, Decimales_Monto)
    
End Sub

Private Sub tabAsiento_Click(PreviousTab As Integer)

    Select Case tabAsiento.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabAsiento.Tab = 0
        
    End Select
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 4 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
    If ColIndex = 5 Then
        Call DarFormatoValor(Value, Decimales_TipoCambio)
    End If
    
End Sub


Private Sub tdgConsulta_HeadClick(ByVal ColIndex As Integer)

    Static numColindex As Integer

    tdgConsulta.Splits(0).Columns(numColindex).HeadingStyle.ForegroundPicture = Null

    Call OrdenarDBGrid(ColIndex, adoConsulta, tdgConsulta)
    
    numColindex = ColIndex

End Sub

Private Sub tdgMovimiento_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If tdgMovimiento.Columns(ColIndex).DataField = "MontoDebe" Or _
       tdgMovimiento.Columns(ColIndex).DataField = "MontoHaber" Or _
       tdgMovimiento.Columns(ColIndex).DataField = "MontoContable" Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
   
End Sub


Private Sub tdgMovimiento_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    
    If adoRegistroAux.EOF Then Exit Sub 'And adoRegistroAux.BOF
    
    lblNumAsiento.Caption = adoRegistroAux.Fields("NumAsiento")
    txtDescripMovimiento.Text = adoRegistroAux.Fields("DescripMovimiento")
    txtCodCuenta.Text = Trim(adoRegistroAux.Fields("CodCuenta"))
    txtCodFile.Text = adoRegistroAux.Fields("CodFile")
    txtCodAnalitica.Text = adoRegistroAux.Fields("CodAnalitica")
    txtDescripCuenta.Text = Trim(txtCodCuenta.Text) & " - " & ObtenerDescripcionCuenta(Trim(txtCodCuenta.Text))
    
    intRegistro = ObtenerItemLista(arrNaturaleza(), adoRegistroAux.Fields("IndDebeHaber"))
    If intRegistro >= 0 Then cboNaturaleza.ListIndex = intRegistro

    intRegistro = ObtenerItemLista(arrMonedaMovimiento(), adoRegistroAux.Fields("CodMonedaMovimiento"))
    If intRegistro >= 0 Then cboMonedaMovimiento.ListIndex = intRegistro
    
    If strCodNaturaleza = "D" Then
        txtMontoMovimiento.Text = adoRegistroAux.Fields("MontoDebe")
    Else
        txtMontoMovimiento.Text = adoRegistroAux.Fields("MontoHaber")
    End If
    
    intRegistro = ObtenerItemLista(arrMonedaContable(), adoRegistroAux.Fields("CodMonedaContable"))
    If intRegistro >= 0 Then cboMonedaContable.ListIndex = intRegistro
    
    txtMontoContable.Text = adoRegistroAux.Fields("MontoContable")
    
    intRegistro = ObtenerItemLista(arrTipoDocumentoDet(), adoRegistroAux.Fields("TipoDocumento"))
    If intRegistro >= 0 Then cboTipoDocumentoDet.ListIndex = intRegistro
    
    txtNumDocumentoDet.Text = adoRegistroAux.Fields("NumDocumento")
    
    intRegistro = ObtenerItemLista(arrTipoPersonaContraparte(), adoRegistroAux.Fields("TipoPersonaContraparte"))
    If intRegistro >= 0 Then cboTipoPersonaContraparte.ListIndex = intRegistro
    
    strCodPersonaContraparte = adoRegistroAux.Fields("CodPersonaContraparte")
    strDescripPersonaContraparte = adoRegistroAux.Fields("DescripPersonaContraparte")
    txtPersonaContraparte.Text = strDescripPersonaContraparte
    
    txtTipoCambioMovimiento.Text = CStr(adoRegistroAux.Fields("ValorTipoCambio"))
    
    If adoRegistroAux.Fields("IndSoloMovimientoContable") = Valor_Indicador Then
        chkMovContable.Value = vbChecked
        Call chkMovContable_Click
    Else
        chkMovContable.Value = vbUnchecked
        Call chkMovContable_Click
    End If
    
    If adoRegistroAux.Fields("IndContracuenta") = Valor_Indicador Then
        chkContracuenta.Value = vbChecked
    Else
        chkContracuenta.Value = vbUnchecked
    End If
    
    If chkContracuenta.Value = vbChecked Then
        strCodContracuenta = adoRegistroAux.Fields("CodContracuenta")
        strCodFileContracuenta = adoRegistroAux.Fields("CodFileContracuenta")
        strCodAnaliticaContracuenta = adoRegistroAux.Fields("CodAnaliticaContracuenta")
        strDescripContracuenta = adoRegistroAux.Fields("DescripContracuenta")
        strDescripFileAnaliticaContracuenta = adoRegistroAux.Fields("DescripFileAnaliticaContracuenta")
        strTipoFileContracuenta = adoRegistroAux.Fields("TipoFileContracuenta")
        
        lblContracuenta.Caption = strCodContracuenta + " / " + strCodFileContracuenta + "-" + strCodAnaliticaContracuenta
    Else
        strCodContracuenta = Valor_Caracter
        strCodFileContracuenta = Valor_Caracter
        strCodAnaliticaContracuenta = Valor_Caracter
        strDescripContracuenta = Valor_Caracter
        strDescripContracuenta = Valor_Caracter
        strDescripFileAnaliticaContracuenta = Valor_Caracter
        strTipoFileContracuenta = Valor_Caracter
        
        lblContracuenta.Caption = Valor_Caracter
    End If

End Sub

Private Sub txtCodAnalitica_LostFocus()

    txtCodAnalitica.Text = Right(String(8, "0") & Trim(txtCodAnalitica.Text), 8)
            
    strCodAnalitica = txtCodAnalitica.Text
            
End Sub


Private Sub txtCodCuenta_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N")
    
End Sub


Private Sub txtCodFile_LostFocus()

    txtCodFile.Text = Right(String(3, "0") & Trim(txtCodFile.Text), 3)
    
    strCodFile = txtCodFile.Text
    
End Sub


Private Sub txtMontoAsiento_Change()

    Call FormatoCajaTexto(txtMontoAsiento, Decimales_Monto)
    
End Sub

Private Sub txtMontoMovimiento_Change()

    'Call FormatoCajaTexto(txtMontoMovimiento, Decimales_Monto)
    Call Calcular
    
End Sub
Private Sub Calcular()

    If Not IsNumeric(txtTipoCambioMovimiento.Text) Or Not IsNumeric(txtMontoMovimiento.Value) Then Exit Sub

    If strIndSoloMovimientoContable = Valor_Caracter Then
        If strCodMonedaMovimiento = Codigo_Moneda_Local Then
            txtMontoContable.Text = CStr(txtMontoMovimiento.Value)
        Else
            txtMontoContable.Text = CStr(txtMontoMovimiento.Value * CDbl(txtTipoCambioMovimiento.Text))
        End If
    End If
    
End Sub



Private Sub txtMontoMovimiento_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call txtMontoMovimiento_Change
    End If

    
End Sub


Private Sub txtMontoMovimiento_LostFocus()

    txtMontoMovimiento.Text = Abs(txtMontoMovimiento.Value)
    If strCodNaturaleza = Codigo_Tipo_Naturaleza_Haber Then
        txtMontoMovimiento.Text = Abs(txtMontoMovimiento.Value) * -1
    End If
        
End Sub

Private Sub txtTipoCambioMovimiento_Change()

    'Call FormatoCajaTexto(txtTipoCambio, Decimales_TipoCambio)
    Call Calcular

End Sub


Private Sub txtTipoCambioMovimiento_KeyPress(KeyAscii As Integer)

    'Call ValidaCajaTexto(KeyAscii, "M", txtTipoCambio, Decimales_TipoCambio)
    If KeyAscii = 13 Then
        Call txtTipoCambioMovimiento_Change
    End If
    
End Sub
Private Sub ConfiguraRecordsetAuxiliarTC()

    Set adoRegistroAuxTC = New ADODB.Recordset

    With adoRegistroAuxTC
       .CursorLocation = adUseClient
       .Fields.Append "CodMonedaOrigen", adChar, 2
       .Fields.Append "CodMonedaCambio", adChar, 2
       .Fields.Append "ValorTipoCambio", adDecimal
       .Fields.Item("ValorTipoCambio").Precision = 20
       .Fields.Item("ValorTipoCambio").NumericScale = 12
       .LockType = adLockBatchOptimistic
    End With

    adoRegistroAuxTC.Open

End Sub
Private Sub ConfiguraRecordsetAuxiliar()

    Set adoRegistroAux = New ADODB.Recordset

    With adoRegistroAux
       .CursorLocation = adUseClient
       .Fields.Append "NumAsiento", adVarChar, 10
       .Fields.Append "SecMovimiento", adInteger, 4
       .Fields.Append "DescripMovimiento", adVarChar, 100
       .Fields.Append "CodCuenta", adVarChar, 10
       .Fields.Append "CodFile", adVarChar, 3
       .Fields.Append "CodAnalitica", adVarChar, 8
       .Fields.Append "DescripAnalitica", adVarChar, 12
       .Fields.Append "IndDebeHaber", adChar, 1
       .Fields.Append "CodMonedaMovimiento", adVarChar, 2
       .Fields.Append "CodSignoMoneda", adVarChar, 3
       .Fields.Append "MontoDebe", adDecimal, 19
       .Fields.Append "MontoHaber", adDecimal, 19
       .Fields.Append "CodMonedaContable", adVarChar, 2
       .Fields.Append "MontoContable", adDecimal, 19 'SOLES
       .Fields.Append "ValorTipoCambio", adDecimal, 20
       .Fields.Append "TipoDocumento", adChar, 2
       .Fields.Append "NumDocumento", adVarChar, 20
       .Fields.Append "TipoPersonaContraparte", adChar, 2
       .Fields.Append "CodPersonaContraparte", adVarChar, 8
       .Fields.Append "DescripPersonaContraparte", adVarChar, 100
       .Fields.Append "IndContracuenta", adChar, 1
       .Fields.Append "CodContracuenta", adVarChar, 10
       .Fields.Append "CodFileContracuenta", adVarChar, 3
       .Fields.Append "CodAnaliticaContracuenta", adVarChar, 8
       .Fields.Append "DescripContracuenta", adVarChar, 100
       .Fields.Append "DescripFileAnaliticaContracuenta", adVarChar, 100
       .Fields.Append "TipoFileContracuenta", adChar, 2
'       .Fields.Append "IndUltimoMovimiento", adVarChar, 1
       .Fields.Append "IndSoloMovimientoContable", adVarChar, 1
       .LockType = adLockBatchOptimistic
    End With

    With adoRegistroAux.Fields.Item("MontoDebe")
        .Precision = 19
        .NumericScale = 2
    End With
    
    With adoRegistroAux.Fields.Item("MontoHaber")
        .Precision = 19
        .NumericScale = 2
    End With
    
    With adoRegistroAux.Fields.Item("MontoContable")
        .Precision = 19
        .NumericScale = 2
    End With
    
    With adoRegistroAux.Fields.Item("ValorTipoCambio")
        .Precision = 20
        .NumericScale = 12
    End With
    
    
    adoRegistroAux.Open

End Sub

