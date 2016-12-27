VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmProvisionCastigos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Porcentajes de Provisiones de Castigos"
   ClientHeight    =   8655
   ClientLeft      =   1140
   ClientTop       =   960
   ClientWidth     =   10350
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
   ScaleHeight     =   8655
   ScaleWidth      =   10350
   Begin TabDlg.SSTab tabAsiento 
      Height          =   7695
      Left            =   60
      TabIndex        =   11
      Top             =   60
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   13573
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmProvisionCastigos.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblDescrip(0)"
      Tab(0).Control(1)=   "lblDescrip(1)"
      Tab(0).Control(2)=   "lblDescrip(2)"
      Tab(0).Control(3)=   "lblDescrip(3)"
      Tab(0).Control(4)=   "lblDescrip(4)"
      Tab(0).Control(5)=   "tdgConsulta"
      Tab(0).Control(6)=   "dtpFechaHasta"
      Tab(0).Control(7)=   "dtpFechaDesde"
      Tab(0).Control(8)=   "cboFondo"
      Tab(0).Control(9)=   "chkIndVigente"
      Tab(0).Control(10)=   "cboTipoInstrumento"
      Tab(0).Control(11)=   "cboClaseInstrumento"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmProvisionCastigos.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraResumen"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -73170
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1650
         Width           =   6900
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -73170
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1140
         Width           =   6900
      End
      Begin VB.CheckBox chkIndVigente 
         Caption         =   "Solo Vigentes"
         Height          =   255
         Left            =   -74580
         TabIndex        =   21
         ToolTipText     =   "Marcar para ver los movimientos de la simulación"
         Top             =   2820
         Width           =   1815
      End
      Begin VB.Frame fraResumen 
         ForeColor       =   &H00800000&
         Height          =   6075
         Left            =   180
         TabIndex        =   15
         Top             =   450
         Width           =   9825
         Begin VB.CommandButton cmdQuitar 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   390
            Picture         =   "frmProvisionCastigos.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Quitar detalle"
            Top             =   5340
            Width           =   465
         End
         Begin VB.CommandButton cmdAgregar 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   390
            Picture         =   "frmProvisionCastigos.frx":028A
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Agregar detalle"
            Top             =   4890
            Width           =   465
         End
         Begin VB.CommandButton cmdActualizar 
            Height          =   435
            Left            =   390
            Picture         =   "frmProvisionCastigos.frx":0537
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Actualizar Detalle"
            Top             =   4440
            Width           =   465
         End
         Begin VB.TextBox txtDescripCastigo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1650
            MaxLength       =   40
            TabIndex        =   30
            Top             =   1950
            Width           =   4785
         End
         Begin VB.ComboBox cboClaseInstrumentoNew 
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
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1560
            Width           =   6900
         End
         Begin VB.ComboBox cboTipoInstrumentoNew 
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
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1110
            Width           =   6900
         End
         Begin VB.ComboBox cboFondoNew 
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
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   270
            Width           =   6900
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
            TabIndex        =   18
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
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   8130
            Width           =   1065
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
            Index           =   2
            Left            =   12780
            TabIndex        =   24
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   8220
            Width           =   375
         End
         Begin VB.TextBox txtDescripAuxiliar 
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
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   8370
            Width           =   9495
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
            TabIndex        =   9
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
            TabIndex        =   8
            Top             =   8400
            Width           =   1635
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
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   8310
            Width           =   4575
         End
         Begin VB.TextBox txtMontoAsiento 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1980
            TabIndex        =   5
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
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   8160
            Width           =   1395
         End
         Begin MSComCtl2.DTPicker dtpFechaDesdeNew 
            Height          =   315
            Left            =   1650
            TabIndex        =   4
            Top             =   2400
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
            Format          =   175898625
            CurrentDate     =   38068
         End
         Begin TrueOleDBGrid60.TDBGrid tdgDetalle 
            Bindings        =   "frmProvisionCastigos.frx":07F2
            Height          =   1365
            Left            =   930
            OleObjectBlob   =   "frmProvisionCastigos.frx":080E
            TabIndex        =   22
            Top             =   4440
            Width           =   8565
         End
         Begin MSComCtl2.DTPicker dtpFechaHastaNew 
            Height          =   315
            Left            =   5000
            TabIndex        =   31
            Top             =   2400
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
            Format          =   175898625
            CurrentDate     =   38068
         End
         Begin TAMControls.TAMTextBox txtValorPorcentaje 
            Height          =   315
            Left            =   2580
            TabIndex        =   45
            Top             =   3860
            Width           =   1485
            _ExtentX        =   2619
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
            Container       =   "frmProvisionCastigos.frx":3C6D
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
         Begin TAMControls.TAMTextBox txtDiaDesde 
            Height          =   315
            Left            =   2580
            TabIndex        =   46
            Top             =   2960
            Width           =   800
            _ExtentX        =   1402
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
            Container       =   "frmProvisionCastigos.frx":3C89
            Text            =   "0"
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtDiaHasta 
            Height          =   315
            Left            =   2580
            TabIndex        =   47
            Top             =   3410
            Width           =   800
            _ExtentX        =   1402
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
            Container       =   "frmProvisionCastigos.frx":3CA5
            Text            =   "0"
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   999999999
         End
         Begin VB.Label lblDescrip 
            Caption         =   "%"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   17
            Left            =   4110
            TabIndex        =   44
            Top             =   3915
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Porcentaje Provisión Flat"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   16
            Left            =   360
            TabIndex        =   43
            Top             =   3920
            Width           =   2175
         End
         Begin VB.Label lblDescrip 
            Caption         =   "(días)"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   15
            Left            =   3500
            TabIndex        =   42
            Top             =   3470
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            Caption         =   "(días)"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   14
            Left            =   3500
            TabIndex        =   41
            Top             =   3020
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Provisión Hasta"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   13
            Left            =   360
            TabIndex        =   40
            Top             =   3470
            Width           =   1695
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Provisión Desde"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   12
            Left            =   360
            TabIndex        =   39
            Top             =   3020
            Width           =   1695
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha Final"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   11
            Left            =   3750
            TabIndex        =   38
            Top             =   2450
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha Inicial"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   10
            Left            =   360
            TabIndex        =   37
            Top             =   2450
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Descripción"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   9
            Left            =   360
            TabIndex        =   36
            Top             =   2000
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Código"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   8
            Left            =   360
            TabIndex        =   35
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Clase"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   7
            Left            =   360
            TabIndex        =   34
            Top             =   1620
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Instrumento"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   6
            Left            =   360
            TabIndex        =   33
            Top             =   1170
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fondo"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   5
            Left            =   360
            TabIndex        =   32
            Top             =   300
            Width           =   615
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
            TabIndex        =   10
            Top             =   8415
            Width           =   1155
         End
         Begin VB.Label lblCodCastigo 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "000"
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
            Left            =   1650
            TabIndex        =   3
            Top             =   690
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
         Left            =   -73170
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   660
         Width           =   6900
      End
      Begin MSComCtl2.DTPicker dtpFechaDesde 
         Height          =   315
         Left            =   -73170
         TabIndex        =   1
         Top             =   2220
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
         Format          =   175898625
         CurrentDate     =   38068
      End
      Begin MSComCtl2.DTPicker dtpFechaHasta 
         Height          =   315
         Left            =   -69200
         TabIndex        =   2
         Top             =   2220
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
         Format          =   175898625
         CurrentDate     =   38068
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmProvisionCastigos.frx":3CC1
         Height          =   4095
         Left            =   -74580
         OleObjectBlob   =   "frmProvisionCastigos.frx":3CDB
         TabIndex        =   20
         Top             =   3480
         Width           =   9315
      End
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   7290
         TabIndex        =   48
         Top             =   6720
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
      Begin VB.Label lblDescrip 
         Caption         =   "Clase"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   4
         Left            =   -74580
         TabIndex        =   26
         Top             =   1690
         Width           =   615
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Instrumento"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   -74580
         TabIndex        =   25
         Top             =   1170
         Width           =   1335
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fecha Final"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   -70500
         TabIndex        =   14
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fecha Inicial"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   -74580
         TabIndex        =   13
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fondo"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   -74580
         TabIndex        =   12
         Top             =   700
         Width           =   615
      End
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   720
      TabIndex        =   49
      Top             =   7830
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   1296
      Buttons         =   4
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Eliminar"
      Tag2            =   "4"
      ToolTipText2    =   "Eliminar"
      Caption3        =   "&Buscar"
      Tag3            =   "5"
      ToolTipText3    =   "Buscar"
      UserControlWidth=   5700
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   8430
      TabIndex        =   50
      Top             =   7830
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
End
Attribute VB_Name = "frmProvisionCastigos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim strSQL                          As String
Dim arrFondo()                      As String, arrMoneda()              As String
Dim arrMonedaMovimiento()           As String, arrNaturaleza()          As String
Dim arrModulo()                     As String, arrTipoFile()            As String
Dim arrTipoAuxiliar()               As String, arrTipoDocumento()       As String
Dim arrTipoPersonaContraparte()     As String, arrTipoDocumentoDet()    As String
Dim arrMonedaContable()             As String
Dim arrTipoInstrumento()            As String, arrClaseInstrumento()    As String
Dim arrFondoNew()                   As String, arrTipoInstrumentoNew()  As String
Dim arrClaseInstrumentoNew()        As String

Dim strCodFondo                 As String, strCodMoneda                         As String
Dim strCodMonedaMovimiento      As String, strCodNaturaleza                     As String
Dim strCodModulo                As String, strEstado                            As String
Dim strCodTipoInstrumento       As String, strCodClaseInstrumento               As String
Dim strCodTipoInstrumentoNew    As String, strCodFileNew                        As String
Dim strCodFondoNew              As String, strCodCastigo                        As String
Dim strCodClaseInstrumentoNew   As String

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
Dim strIndVigente               As String

Public Sub Buscar()

    Dim strSQL          As String
    Dim strFechaDesde   As String, strFechaHasta As String
    Dim datFecha        As Date
    
    Set adoConsulta = New ADODB.Recordset
    
    strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
    datFecha = DateAdd("d", 1, dtpFechaHasta.Value)
    strFechaHasta = Convertyyyymmdd(datFecha)
    
    strSQL = "SELECT SecCastigo, DiaDesde, DiaHasta, ValorPorcentaje " & _
            "FROM ProvisionCastigo PC JOIN ProvisionCastigoDetalle PCD ON " & _
            "(PC.CodFondo=PCD.CodFondo AND PC.CodFile=PCD.CodFile AND PC.CodDetalleFile=PCD.CodDetalleFile AND PC.CodCastigo=PCD.CodCastigo) " & _
            "WHERE PC.CodFondo='" & strCodFondo & "' AND PC.CodFile='" & strCodFile & "' AND PC.CodDetalleFile='" & strCodClaseInstrumento & "' "
               
    If chkIndVigente.Value = vbChecked Then
        strSQL = strSQL & " AND (PC.IndVigente = 'X')"
    End If
                        
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
                            
    'lblTotalDebeME.Caption = "0"
    'lblTotalHaberME.Caption = "0"
    'lblTotalDebeMN.Caption = "0"
    'lblTotalHaberMN.Caption = "0"
    
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
    
'    lblTotalDebeME.Caption = CStr(dblAcumuladoDebeME)
'    lblTotalHaberME.Caption = CStr(dblAcumuladoHaberME)
'    lblTotalDebeMN.Caption = CStr(dblAcumuladoDebeMN)
'    lblTotalHaberMN.Caption = CStr(dblAcumuladoHaberMN)
    
            
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
'
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Diario General Analítico"
'
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Mayor General Analítico"
    
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

'    Select Case Index
            
                
'         Case 1
'
'
'            gstrNumAsiento = Valor_Caracter
'
'            If tabAsiento.Tab = 1 And (strEstado = Reg_Edicion Or strEstado = Reg_Consulta) Then
'                '*** Comprobante seleccionado ***
''                strSeleccionRegistro = "{AsientoContable.NumAsiento} = '" & Trim(lblNumAsiento.Caption) & "'"
''                strSeleccionRegistro = strSeleccionRegistro & " AND {AsientoContableDetalle.NumAsiento} = '" & Trim(lblNumAsiento.Caption) & "'"
''                strSeleccionRegistro = strSeleccionRegistro & " AND {AsientoContableDetalle.FechaMovimiento} = '" & Convertyyyymmdd(dtpFechaAsiento.Value) & "'"
''                gstrSelFrml = strSeleccionRegistro
''                gstrNumAsiento = CStr(CLng(Trim(lblNumAsiento.Caption)))
'                gstrCodMonedaReporte = Codigo_Moneda_Local
'                gstrSelFrml = "1"
'            Else
'                '*** Lista de comprobantes por rango de fecha ***
'                strSeleccionRegistro = "{AsientoContable.FechaAsiento} IN 'Fch1' TO 'Fch2'"
'                gstrSelFrml = strSeleccionRegistro
'                'frmRangoFecha.Show vbModal
''                frmFiltroReporte.strCodFondo = strCodFondo
''                frmFiltroReporte.strCodAdministradora = gstrCodAdministradora
''
''                frmFiltroReporte.chkOpcionFiltro(1).Enabled = False
''                frmFiltroReporte.chkOpcionFiltro(1).Value = 0
''                frmFiltroReporte.txtCodCuenta.Enabled = False
''                frmFiltroReporte.cmdBusquedaCuenta.Enabled = False
''
''                frmFiltroReporte.chkOpcionFiltro(2).Enabled = True
''                frmFiltroReporte.chkOpcionFiltro(2).Value = 0
''                frmFiltroReporte.txtNumAsiento.Enabled = False
''
''                frmFiltroReporte.Show vbModal
'
'            End If
'
'            If gstrSelFrml <> "0" Then
'                Set frmReporte = New frmVisorReporte
''INICIO: REVISAR NUEVA VERSION SPECTRUM 1_5
''                ReDim aReportParamS(9)
''FIN: REVISAR NUEVA VERSION SPECTRUM 1_5
'                ReDim aReportParamS(8)
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
'
'                If tabAsiento.Tab = 1 And (strEstado = Reg_Edicion Or strEstado = Reg_Consulta) Then
''                    aReportParamF(1) = dtpFechaAsiento
''                    aReportParamF(2) = dtpFechaAsiento
'                Else
'                    aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
'                    aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
'                End If
'
'                aReportParamF(3) = Format(Time(), "hh:mm:ss")
'                aReportParamF(4) = Trim(cboFondo.Text)
'                aReportParamF(5) = gstrNombreEmpresa & Space(1)
'
'                'SP
'                aReportParamS(0) = strCodFondo
'                aReportParamS(1) = gstrCodAdministradora
'
'                If tabAsiento.Tab = 1 And (strEstado = Reg_Edicion Or strEstado = Reg_Consulta) Then
''                    aReportParamS(2) = Convertyyyymmdd(dtpFechaAsiento.Value)
''                    aReportParamS(3) = Convertyyyymmdd(dtpFechaAsiento.Value)
'                Else
'                    aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
'                    aReportParamS(3) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))
'                End If
'
'                aReportParamS(4) = gstrCodMonedaReporte 'strCodMoneda
'                aReportParamS(5) = gstrCodClaseTipoCambioFondo 'Codigo_Listar_Todos
'                aReportParamS(6) = gstrValorTipoCambioCierre   '"0000000000"
'                'aReportParamS(7) = Codigo_Listar_Todos
'
'                If Trim(gstrNumAsiento) <> Valor_Caracter Then
'                    aReportParamS(7) = Codigo_Listar_Individual '"I"
'                    aReportParamS(8) = gstrNumAsiento
'                Else
'                    aReportParamS(7) = Codigo_Listar_Todos '"T"
'                    aReportParamS(8) = "%"
'                End If
'
''INICIO: REVISAR NUEVA VERSION SPECTRUM 1_5
''                If chkIndVigente.Value Then
''                    aReportParamS(9) = "1"
''                Else
''                    aReportParamS(9) = "0"
''                End If
''
''                gstrNameRepo = "LibroDiarioAnalitico"
''FIN: REVISAR NUEVA VERSION SPECTRUM 1_5
'
'                If chkIndVigente.Value Then
'                    gstrNameRepo = "SLibroDiario"
'                Else
'                    gstrNameRepo = "LibroDiarioMM"
'                End If
'
'
'            End If
'
'
'        Case 2
'
'            '*** Lista de comprobantes por rango de fecha ***
'            strSeleccionRegistro = "{AsientoContable.FechaAsiento} IN 'Fch1' TO 'Fch2'"
'            gstrSelFrml = strSeleccionRegistro
'            'frmRangoFecha.Show vbModal
''            frmFiltroReporte.strCodFondo = strCodFondo
''            frmFiltroReporte.strCodAdministradora = gstrCodAdministradora
''            frmFiltroReporte.chkOpcionFiltro(1).Enabled = True
''            frmFiltroReporte.chkOpcionFiltro(1).Value = 1
''            frmFiltroReporte.txtCodCuenta.Enabled = True
''            frmFiltroReporte.cmdBusquedaCuenta.Enabled = True
''
''            frmFiltroReporte.chkOpcionFiltro(2).Enabled = False
''            frmFiltroReporte.chkOpcionFiltro(2).Value = 0
''            frmFiltroReporte.txtNumAsiento.Enabled = False
''
''            frmFiltroReporte.Show vbModal
'
'            If gstrSelFrml <> "0" Then
'                Set adoRegistro = New ADODB.Recordset
'
'                '*** Se Realizó Cierre anteriormente ? ***
'
'                adoComm.CommandText = "{ call up_GNValidaCierreRealizado('" & _
'                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
'                    Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)) & "','" & _
'                    Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)))) & "') }"
'
'                Set adoRegistro = adoComm.Execute
'                If adoRegistro.EOF Then
'                    MsgBox "El Cierre del Día " & Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10) & " No fué realizado.", vbCritical, Me.Caption
'
'                    adoRegistro.Close: Set adoRegistro = Nothing
'                    Exit Sub
'                End If
'                adoRegistro.Close: Set adoRegistro = Nothing
'
'                Set frmReporte = New frmVisorReporte
'                'Dim strCuenta As String
'
'                ReDim aReportParamS(8)
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
'                aReportParamS(3) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))
'                aReportParamS(4) = gstrCodMonedaReporte 'strCodMoneda
'
'                aReportParamS(5) = gstrCodClaseTipoCambioFondo 'Codigo_Listar_Todos
'                aReportParamS(6) = gstrValorTipoCambioCierre   '"0000000000"
'                aReportParamS(7) = Codigo_Listar_Todos
'
'                If Trim(gstrCodCuenta) = Valor_Caracter Or gstrCodCuenta = "0000000000" Then
'                    aReportParamS(8) = "%" 'gstrCodCuenta '"0000000000"
'                Else
'                    aReportParamS(8) = gstrCodCuenta
'                End If
'
'                'gstrNameRepo = "LibroMayorAnalitico"
'
'                gstrNameRepo = "HistLibroMayor1"
'
'
'            End If
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

Private Sub cboClaseInstrumento_Click()
    
    strCodClaseInstrumento = Valor_Caracter
    If cboClaseInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodClaseInstrumento = Trim(arrClaseInstrumento(cboClaseInstrumento.ListIndex))
    
    Call Buscar
    
End Sub

Private Sub cboClaseInstrumentoNew_Click()

    strCodClaseInstrumentoNew = Valor_Caracter
    If cboClaseInstrumentoNew.ListIndex < 0 Then Exit Sub
    
    strCodClaseInstrumentoNew = Trim(arrClaseInstrumentoNew(cboClaseInstrumentoNew.ListIndex))

End Sub

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
'            strSQL = "{ call up_ACSelDatosParametro(70,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
'            CargarControlLista strSQL, cboMonedaContable, arrMonedaContable(), Sel_Todos
        
'            If cboMonedaContable.ListCount > 0 Then cboMonedaContable.ListIndex = 0
            
'            txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, gstrCodMoneda, Codigo_Moneda_Local))
'            If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, gdatFechaActual), gstrCodMoneda, Codigo_Moneda_Local))
'            gdblTipoCambio = CDbl(txtTipoCambio.Text)
'
'            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        
        
        
'            Call Buscar
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & _
            "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & _
            "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_CortoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' ORDER BY CODIGO"
    CargarControlLista strSQL, cboTipoInstrumento, arrTipoInstrumento(), Sel_Defecto
        
    cboTipoInstrumento.ListIndex = 0
    
End Sub


Private Sub cboModulo_Click()

'    strCodModulo = ""
'    If cboModulo.ListIndex < 0 Then Exit Sub
'
'    strCodModulo = Trim(arrModulo(cboModulo.ListIndex))
    
End Sub


Private Sub cboFondoNew_Click()

Dim adoRegistro As ADODB.Recordset
    
    strCodFondoNew = Valor_Caracter
    If cboFondoNew.ListIndex < 0 Then Exit Sub
    
    strCodFondoNew = Trim(arrFondoNew(cboFondoNew.ListIndex))
    
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
        
'            Call Buscar
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & _
            "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & _
            "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_CortoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' ORDER BY CODIGO"
    CargarControlLista strSQL, cboTipoInstrumentoNew, arrTipoInstrumentoNew(), Sel_Defecto
        
    cboTipoInstrumentoNew.ListIndex = 0

End Sub

Private Sub cboMoneda_Click()

    strCodMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
    strCodMonedaParEvaluacion = strCodMoneda & Codigo_Moneda_Local
    
    If strCodMoneda <> Codigo_Moneda_Local Then
        strCodMonedaParPorDefecto = ObtenerMonedaParPorDefecto(gstrCodClaseTipoCambioOperacionFondo, strCodMonedaParEvaluacion)
    Else
        strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
    End If
    
    'lblDescripTC.Caption = "(" + Trim(ObtenerCodSignoMoneda(Mid(strCodMonedaParPorDefecto, 3, 2))) + "/" + Trim(ObtenerCodSignoMoneda(Mid(strCodMonedaParPorDefecto, 1, 2))) + ")"
    
    If strCodMoneda <> Codigo_Moneda_Local Then
'        txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, Mid(strCodMonedaParPorDefecto, 1, 2), Mid(strCodMonedaParPorDefecto, 3, 2)))
'        txtTipoCambio.Enabled = True
    Else
'        txtTipoCambio.Text = "1"
'        txtTipoCambio.Enabled = True 'False
    End If
'    Call txtTipoCambio_KeyPress(vbKeyReturn)
    
'    txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaAsiento.Value, Codigo_Moneda_Local, strCodMoneda))
'    If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaAsiento.Value), Codigo_Moneda_Local, strCodMoneda))
    
End Sub


Private Sub cboMonedaContable_Click()
    
    Dim intRegistro As Integer
    
    strCodMonedaContable = Valor_Caracter
'    If cboMonedaContable.ListIndex < 0 Then Exit Sub
    
'    strCodMonedaContable = arrMonedaContable(cboMonedaContable.ListIndex)
    
'    If strCodMonedaContable = Valor_Caracter Then
'        chkMovContable.Value = vbUnchecked
'        Call chkMovContable_Click
'    Else
'        chkMovContable.Value = vbChecked
'        Call chkMovContable_Click
'    End If

    If strIndSoloMovimientoContable = Valor_Indicador Then
        intRegistro = ObtenerItemLista(arrMonedaMovimiento(), strCodMonedaContable)
        'If intRegistro >= 0 Then cboMonedaMovimiento.ListIndex = intRegistro
    End If
    
End Sub

Private Sub cboMonedaMovimiento_Click()
'
'    Dim dblValorTC As Double
'
'    strCodMonedaMovimiento = Valor_Caracter
'    If cboMonedaMovimiento.ListIndex < 0 Then Exit Sub
'
'    strCodMonedaMovimiento = Trim(arrMonedaMovimiento(cboMonedaMovimiento.ListIndex))
'
'    strCodMonedaParEvaluacion = strCodMonedaMovimiento & Codigo_Moneda_Local
'
'    If strCodMonedaMovimiento <> Codigo_Moneda_Local Then
'        strCodMonedaParPorDefecto = ObtenerMonedaParPorDefecto(gstrCodClaseTipoCambioOperacionFondo, strCodMonedaParEvaluacion)
'    Else
'        strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
'    End If
'
'    If strCodMonedaParPorDefecto = "0000" Then strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
'
'    If strCodMoneda <> Codigo_Moneda_Local Then
'        txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, Mid(strCodMonedaParPorDefecto, 1, 2), Mid(strCodMonedaParPorDefecto, 3, 2)))
'        txtTipoCambio.Enabled = True
'        lblDescripTC.Caption = "(" + Trim(ObtenerCodSignoMoneda(Mid(strCodMonedaParPorDefecto, 3, 2))) + "/" + Trim(ObtenerCodSignoMoneda(Mid(strCodMonedaParPorDefecto, 1, 2))) + ")"
'    Else
'        txtTipoCambio.Text = "1"
'        txtTipoCambio.Enabled = True 'False
'        lblDescripTC.Caption = Valor_Caracter
'    End If
'
'    If strIndSoloMovimientoContable = Valor_Caracter Then
''        txtMontoContable.Text = txtMontoMovimiento.Text
''    Else
'        dblValorTC = CDbl(txtTipoCambio.Text)
''        dblValorTC = ObtenerTipoCambioArbitraje(dblValorTC, strCodMonedaParEvaluacion, strCodMonedaParPorDefecto)
'        txtMontoContable.Text = CStr(CDbl(txtMontoMovimiento.Text) * dblValorTC)
'    End If
    
End Sub


Private Sub cboNaturaleza_Click()

'    strCodNaturaleza = Valor_Caracter
'    If cboNaturaleza.ListIndex < 0 Then Exit Sub
'
'    strCodNaturaleza = Trim(arrNaturaleza(cboNaturaleza.ListIndex))
'
'    If strIndSoloMovimientoContable = Valor_Caracter Then
'        txtMontoMovimiento.Text = Abs(CDbl(txtMontoMovimiento.Text))
'        If strCodNaturaleza = Codigo_Tipo_Naturaleza_Haber Then
'            txtMontoMovimiento.Text = CStr(Abs(CDbl(txtMontoMovimiento.Text)) * -1)
'        End If
'    Else
'        txtMontoContable.Text = Abs(CDbl(txtMontoContable.Text))
'        If strCodNaturaleza = Codigo_Tipo_Naturaleza_Haber Then
'            txtMontoContable.Text = CStr(Abs(CDbl(txtMontoContable.Text)) * -1)
'        End If
'    End If
    
End Sub

Private Sub cboTipoInstrumento_Click()
    
    Dim adoRegistro As ADODB.Recordset
    Dim strFecha    As String
    
    strCodTipoInstrumento = Valor_Caracter
    'strIndPacto = Valor_Caracter: strIndNegociable = Valor_Caracter
    If cboTipoInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumento = Trim(arrTipoInstrumento(cboTipoInstrumento.ListIndex))
    
    'Asignar nemónico
    'txtNemonico.Text = GenerarNemonico(strCodTipoInstrumentoOrden, strCodClaseInstrumento, strCodEmisor, txtNumDocDscto.Text)
    'txtNemonico.Text = GenerarNemonico(strCodTipoInstrumentoOrden, strCodClaseInstrumento, strCodEmisor, txtNumDocDsctoCancel.Text)
    'txtDescripOrden.Text = Trim(cboTipoInstrumentoOrden.Text) & " - " & txtNemonico.Text
    'txtDescripOrdenCancel.Text = Trim(cboTipoInstrumentoOrden.Text) & " - " & txtNemonicoCancel.Text
            
    'lblAnalitica.Caption = strCodTipoInstrumentoOrden & " - ????????"
    strCodFile = strCodTipoInstrumento
    
    '*** Clase de Instrumento ***
    strSQL = "SELECT CodDetalleFile CODIGO,DescripDetalleFile DESCRIP FROM InversionDetalleFile WHERE CodFile='" & strCodTipoInstrumento & "' AND IndVigente='X' ORDER BY DescripDetalleFile"
    CargarControlLista strSQL, cboClaseInstrumento, arrClaseInstrumento(), Sel_Defecto
    
    If cboClaseInstrumento.ListCount > 0 Then
        cboClaseInstrumento.ListIndex = 0
        cboClaseInstrumento.Enabled = True
    End If
    
End Sub

Private Sub cboTipoPersonaContraparte_Click()
'
'    strTipoPersonaContraparte = Valor_Caracter
'
'    If cboTipoPersonaContraparte.ListIndex < 0 Then Exit Sub
'
'    strTipoPersonaContraparte = arrTipoPersonaContraparte(cboTipoPersonaContraparte.ListIndex)
'
'    txtPersonaContraparte.Text = ""
'
'    If cboTipoPersonaContraparte.ListIndex > 0 Then
'        cmdBusqueda(3).Enabled = True
'    Else
'        cmdBusqueda(3).Enabled = False
'    End If

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

'    strTipoDocumento = Valor_Caracter
'    If cboTipoDocumento.ListIndex < 0 Then Exit Sub
'
'    strTipoDocumento = arrTipoDocumento(cboTipoDocumento.ListIndex)

End Sub

Private Sub cboTipoDocumentoDet_Click()
'    strTipoDocumentoDet = Valor_Caracter
'
'    If cboTipoDocumentoDet.ListIndex < 0 Then Exit Sub
'
'    strTipoDocumentoDet = arrTipoDocumentoDet(cboTipoDocumentoDet.ListIndex)
End Sub



Private Sub chkContracuenta_Click()

'    If chkContracuenta.Value = vbChecked Then
'        cmdContracuenta.Enabled = True
'        If strCodContracuenta <> Valor_Caracter Then
'            lblContracuenta.Caption = strCodContracuenta + " / " + strCodFileContracuenta + "-" + strCodAnaliticaContracuenta
'        Else
'            lblContracuenta.Caption = Valor_Caracter
'        End If
'    Else
'        cmdContracuenta.Enabled = False
'        lblContracuenta.Caption = Valor_Caracter
'    End If

End Sub



Private Sub chkMovContable_Click()

'    If chkMovContable.Value = vbChecked Then
'        txtMontoMovimiento.Text = "0"
'        txtMontoMovimiento.Enabled = False
'    Else
'        txtMontoMovimiento.Enabled = True
'    End If
'    Dim intRegistro As Integer
'
'    If chkMovContable.Value = vbChecked Then
'        strIndSoloMovimientoContable = Valor_Indicador
'        txtMontoMovimiento.Text = "0"
'        txtMontoMovimiento.Enabled = False
'        txtMontoContable.Enabled = True
'
'        lblDescrip(32).Visible = True
'        cboMonedaContable.Visible = True
'
'        If cboMonedaContable.ListIndex > 0 Then
'            intRegistro = ObtenerItemLista(arrMonedaContable(), strCodMonedaContable)
'            If intRegistro >= 0 Then cboMonedaContable.ListIndex = intRegistro
'        Else
'            intRegistro = ObtenerItemLista(arrMonedaContable(), Codigo_Moneda_Local)
'            If intRegistro >= 0 Then cboMonedaContable.ListIndex = intRegistro
'        End If
'
'    Else
'        strIndSoloMovimientoContable = Valor_Caracter
'        txtMontoMovimiento.Enabled = True
'        txtMontoContable.Enabled = False
'
'        lblDescrip(32).Visible = False
'        cboMonedaContable.Visible = False
'        cboMonedaContable.ListIndex = 0
'
'    End If

End Sub

Private Sub cboTipoInstrumentoNew_Click()
    
Dim adoRegistro As ADODB.Recordset
    Dim strFecha    As String
    
    strCodTipoInstrumentoNew = Valor_Caracter

    If cboTipoInstrumentoNew.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumentoNew = Trim(arrTipoInstrumentoNew(cboTipoInstrumentoNew.ListIndex))
    
    strCodFileNew = strCodTipoInstrumentoNew
    
    '*** Clase de Instrumento ***
    strSQL = "SELECT CodDetalleFile CODIGO,DescripDetalleFile DESCRIP FROM InversionDetalleFile WHERE CodFile='" & strCodTipoInstrumentoNew & "' AND IndVigente='X' ORDER BY DescripDetalleFile"
    CargarControlLista strSQL, cboClaseInstrumentoNew, arrClaseInstrumentoNew(), Sel_Defecto
    
    If cboClaseInstrumentoNew.ListCount > 0 Then
        cboClaseInstrumentoNew.ListIndex = 0
        cboClaseInstrumentoNew.Enabled = True
    End If
    
End Sub

Private Sub chkIndVigente_Click()

'    If chkIndVigente.Value = vbChecked Then
'        'strTipoProceso = "1"
'        strIndVigente = Valor_Indicador
'    End If
'
'    If chkIndVigente.Value = vbUnchecked Then
'        'strTipoProceso = "0"
'        strIndVigente = Valor_Caracter
'    End If

    Call Buscar

End Sub


Private Sub cmdActualizar_Click()

'    Dim strCriterio As String
'
'    'VALIDAR QUE EXISTA REGISTRO
'    If adoRegistroAux.RecordCount = 0 Then
'        MsgBox "No puede editar un movimiento si no existen registros en el detalle del asiento!", vbInformation, Me.Caption
'        Exit Sub
'    End If
'
'    If adoRegistroAux.EOF Then
'        MsgBox "Debe seleccionar un movimiento para editar!", vbInformation, Me.Caption
'        Exit Sub
'    End If
'
'    adoRegistroAux.Fields("NumAsiento") = lblNumAsiento.Caption
'    'adoRegistroAux.Fields("SecMovimiento") = 0 'numSecMovimiento
'    'adoRegistroAux.Fields("DescripMovimiento") = txtDescripMovimiento.Text
'    adoRegistroAux.Fields("CodCuenta") = Trim(txtCodCuenta.Text)
'    adoRegistroAux.Fields("CodFile") = Trim(txtCodFile.Text)
'    adoRegistroAux.Fields("CodAnalitica") = Trim(txtCodAnalitica.Text)
'    adoRegistroAux.Fields("DescripAnalitica") = Trim(txtCodFile.Text) + "-" + Trim(txtCodAnalitica.Text)
'    adoRegistroAux.Fields("IndDebeHaber") = strCodNaturaleza
'    adoRegistroAux.Fields("CodMonedaMovimiento") = strCodMonedaMovimiento
'    adoRegistroAux.Fields("CodSignoMoneda") = ObtenerCodSignoMoneda(strCodMonedaMovimiento)
'
'    If strCodNaturaleza = "D" Then
'        adoRegistroAux.Fields("MontoDebe") = CDbl(Trim(txtMontoMovimiento.Text))
'        adoRegistroAux.Fields("MontoHaber") = 0
'    Else
'        adoRegistroAux.Fields("MontoHaber") = CDbl(Trim(txtMontoMovimiento.Text))
'        adoRegistroAux.Fields("MontoDebe") = 0
'    End If
'
'
'    adoRegistroAux.Fields("CodMonedaContable") = strCodMonedaContable
'
'    adoRegistroAux.Fields("MontoContable") = CDbl(txtMontoContable.Value)
'
'    adoRegistroAux.Fields("ValorTipoCambio") = CDbl(txtTipoCambio.Text)
'
''    adoRegistroAux.Fields("TipoAuxiliar") = strTipoAuxiliar
''    adoRegistroAux.Fields("CodAuxiliar") = strCodAuxiliar
'
'    'nuevo
'    adoRegistroAux.Fields("TipoDocumento") = strTipoDocumentoDet 'arrTipoDocumentoDet(cboTipoDocumentoDet.ListIndex)
'    adoRegistroAux.Fields("NumDocumento") = txtNumDocumentoDet.Text
'
'    If cboTipoPersonaContraparte.ListIndex <> -1 Then
'        adoRegistroAux.Fields("TipoPersonaContraparte") = arrTipoPersonaContraparte(cboTipoPersonaContraparte.ListIndex)
'    Else
'        adoRegistroAux.Fields("TipoPersonaContraparte") = Valor_Caracter
'    End If
'
'    adoRegistroAux.Fields("CodPersonaContraparte") = strCodPersonaContraparte
'    adoRegistroAux.Fields("DescripPersonaContraparte") = strDescripPersonaContraparte
'
'    adoRegistroAux.Fields("IndSoloMovimientoContable") = strIndSoloMovimientoContable
'
'    If chkContracuenta.Value = vbChecked Then
'        adoRegistroAux.Fields("IndContracuenta") = Valor_Indicador
'
'        adoRegistroAux.Fields("CodContracuenta") = strCodContracuenta
'        adoRegistroAux.Fields("CodFileContracuenta") = strCodFileContracuenta
'        adoRegistroAux.Fields("CodAnaliticaContracuenta") = strCodAnaliticaContracuenta
'        adoRegistroAux.Fields("DescripContracuenta") = strDescripContracuenta
'        adoRegistroAux.Fields("DescripFileAnaliticaContracuenta") = strDescripFileAnaliticaContracuenta
'    Else
'        adoRegistroAux.Fields("IndContracuenta") = Valor_Caracter
'
'        adoRegistroAux.Fields("CodContracuenta") = Valor_Caracter
'        adoRegistroAux.Fields("CodFileContracuenta") = Valor_Caracter
'        adoRegistroAux.Fields("CodAnaliticaContracuenta") = Valor_Caracter
'        adoRegistroAux.Fields("DescripContracuenta") = Valor_Caracter
'        adoRegistroAux.Fields("DescripFileAnaliticaContracuenta") = Valor_Caracter
'    End If
'
''    adoRegistroAux.Fields("IndUltimoMovimiento") = strIndUltimoMoviiento
''    adoRegistroAux.Fields("IndSoloMovimientoContable") = strIndSoloMovimientoContable
'
'
'    If adoRegistroAux.Fields("CodMonedaMovimiento") <> Codigo_Moneda_Local Then
'        strCriterio = "CodMonedaOrigen='" & Mid(strCodMonedaParPorDefecto, 1, 2) & "'" & _
'                      " AND CodMonedaCambio = '" & Mid(strCodMonedaParPorDefecto, 3, 2) & "'" & _
'                      " AND ValorTipoCambio = " & CDbl(txtTipoCambio.Text)
''        If Not FindRecordset(adoRegistroAuxTC, strCriterio) Then
''            adoRegistroAuxTC.AddNew
''            adoRegistroAuxTC.Fields("CodMonedaOrigen") = Mid(strCodMonedaParPorDefecto, 1, 2) 'strCodMoneda
''            adoRegistroAuxTC.Fields("CodMonedaCambio") = Mid(strCodMonedaParPorDefecto, 3, 2) 'strCodMonedaCuenta
''            adoRegistroAuxTC.Fields("ValorTipoCambio") = CDbl(txtTipoCambio.Text)
''        End If
'    End If
'
'    Call TotalizarMovimientos

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
            adoRegistroAux.Fields("SecCastigo") = 0
            adoRegistroAux.Fields("DiaDesde") = txtDiaDesde.Value
            adoRegistroAux.Fields("DiaHasta") = txtDiaHasta.Value
            adoRegistroAux.Fields("ValorPorcentaje") = txtValorPorcentaje.Value
            
            adoRegistroAux.Update
            
            dblBookmark = adoRegistroAux.Bookmark
            
            tdgDetalle.Refresh
            
            Call NumerarRegistros
            
            adoRegistroAux.Bookmark = dblBookmark
            
'            adoRegistroAux.Bookmark = dblBookmark
            
            cmdQuitar.Enabled = True
                        
            txtDiaDesde.Text = Valor_Caracter
            txtDiaHasta.Text = Valor_Caracter
            txtValorPorcentaje.Text = Valor_Caracter
            
            'Call LimpiarDatos
           
'            adoRegistroAux.AddNew
'            adoRegistroAux.Fields("NumAsiento") = lblNumAsiento.Caption
'            adoRegistroAux.Fields("SecMovimiento") = 0 'numSecMovimiento
'            'adoRegistroAux.Fields("DescripMovimiento") = txtDescripMovimiento.Text
'            adoRegistroAux.Fields("CodCuenta") = Trim(txtCodCuenta.Text)
'            adoRegistroAux.Fields("CodFile") = Trim(txtCodFile.Text)
'            adoRegistroAux.Fields("CodAnalitica") = Trim(txtCodAnalitica.Text)
'            adoRegistroAux.Fields("DescripAnalitica") = Trim(txtCodFile.Text) + "-" + Trim(txtCodAnalitica.Text)
'            adoRegistroAux.Fields("IndDebeHaber") = strCodNaturaleza
'            adoRegistroAux.Fields("CodMonedaMovimiento") = strCodMonedaMovimiento
'            adoRegistroAux.Fields("CodSignoMoneda") = ObtenerCodSignoMoneda(strCodMonedaMovimiento)
'
'            If strCodNaturaleza = "D" Then
'                adoRegistroAux.Fields("MontoDebe") = CDbl(Trim(txtMontoMovimiento.Text))
'                adoRegistroAux.Fields("MontoHaber") = 0
'            Else
'                adoRegistroAux.Fields("MontoHaber") = CDbl(Trim(txtMontoMovimiento.Text))
'                adoRegistroAux.Fields("MontoDebe") = 0
'            End If
'
'            adoRegistroAux.Fields("CodMonedaContable") = strCodMonedaContable
'
'            adoRegistroAux.Fields("MontoContable") = CDbl(txtMontoContable.Value)
'
'            adoRegistroAux.Fields("ValorTipoCambio") = CDbl(txtTipoCambio.Text)
'
''            adoRegistroAux.Fields("TipoAuxiliar") = strTipoAuxiliar
''            adoRegistroAux.Fields("CodAuxiliar") = strCodAuxiliar
'
'            '**************BMM NUEVOS CAMBIOS *******
'            adoRegistroAux.Fields("TipoDocumento") = strTipoDocumentoDet
'            adoRegistroAux.Fields("NumDocumento") = txtNumDocumentoDet.Text
'
'            If cboTipoPersonaContraparte.ListIndex <> -1 Then
'                adoRegistroAux.Fields("TipoPersonaContraparte") = arrTipoPersonaContraparte(cboTipoPersonaContraparte.ListIndex)
'            Else
'                adoRegistroAux.Fields("TipoPersonaContraparte") = Valor_Caracter
'            End If
'
'            adoRegistroAux.Fields("CodPersonaContraparte") = strCodPersonaContraparte
'            adoRegistroAux.Fields("DescripPersonaContraparte") = strDescripPersonaContraparte
'
'            adoRegistroAux.Fields("IndSoloMovimientoContable") = strIndSoloMovimientoContable
'
'            If chkContracuenta.Value = vbChecked Then
'                adoRegistroAux.Fields("CodContracuenta") = strCodContracuenta
'                adoRegistroAux.Fields("CodFileContracuenta") = strCodFileContracuenta
'                adoRegistroAux.Fields("CodAnaliticaContracuenta") = strCodAnaliticaContracuenta
'                adoRegistroAux.Fields("DescripContracuenta") = strDescripContracuenta
'                adoRegistroAux.Fields("DescripFileAnaliticaContracuenta") = strDescripFileAnaliticaContracuenta
'                adoRegistroAux.Fields("IndContracuenta") = Valor_Indicador
'            Else
'                adoRegistroAux.Fields("CodContracuenta") = Valor_Caracter
'                adoRegistroAux.Fields("CodFileContracuenta") = Valor_Caracter
'                adoRegistroAux.Fields("CodAnaliticaContracuenta") = Valor_Caracter
'                adoRegistroAux.Fields("DescripContracuenta") = Valor_Caracter
'                adoRegistroAux.Fields("DescripFileAnaliticaContracuenta") = Valor_Caracter
'                adoRegistroAux.Fields("IndContracuenta") = Valor_Caracter
'            End If
'
'
'            If adoRegistroAux.Fields("CodMonedaMovimiento") <> Codigo_Moneda_Local Then
'                strCriterio = "CodMonedaOrigen='" & Mid(strCodMonedaParPorDefecto, 1, 2) & "'" & _
'                              " AND CodMonedaCambio = '" & Mid(strCodMonedaParPorDefecto, 3, 2) & "'" & _
'                              " AND ValorTipoCambio = " & CDbl(txtTipoCambio.Text)
''                If Not FindRecordset(adoRegistroAuxTC, strCriterio) Then
''                    adoRegistroAuxTC.AddNew
''                    adoRegistroAuxTC.Fields("CodMonedaOrigen") = Mid(strCodMonedaParPorDefecto, 1, 2) 'strCodMoneda
''                    adoRegistroAuxTC.Fields("CodMonedaCambio") = Mid(strCodMonedaParPorDefecto, 3, 2) 'strCodMonedaCuenta
''                    adoRegistroAuxTC.Fields("ValorTipoCambio") = CDbl(txtTipoCambio.Text)
''                End If
'            End If
'
'            adoRegistroAux.Update
'
'            dblBookmark = adoRegistroAux.Bookmark
'
'            tdgDetalle.Refresh
'
'            Call NumerarRegistros
'
'            adoRegistroAux.Bookmark = dblBookmark
'
'            Call TotalizarMovimientos
'
'            adoRegistroAux.Bookmark = dblBookmark
'
'            cmdQuitar.Enabled = True
'            Call LimpiarDatos
        
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
        adoRegistroAux.Fields("SecCastigo") = n
        adoRegistroAux.Update
        n = n + 1
        adoRegistroAux.MoveNext
    Wend


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
                
                frmBus.Caption = " Relación de Auxiliares Contables"
                .sSql = "{ call up_CNSelAuxiliarContable('" & strTipoAuxiliar & "') }"
                .OutputColumns = "1,2,3"
                .HiddenColumns = "3"
                
            Case 3
            
                frmBus.Caption = " Relacion de Personas"
                
                '** OBTENGO EL TIPO DEL COMBO SELECCIONADO **
                'strTipoPersonaContraparte = arrTipoPersonaContraparte(cboTipoPersonaContraparte.ListIndex)
                                
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
                    
                    'txtCodCuenta.Text = strCodCuenta
                    
                    'txtDescripCuenta.Text = strCodCuenta & " - " & strDescripCuenta
                    
                    'txtDescripFileAnalitica.Text = ""
                    strCodFile = ""
                    strCodAnalitica = ""
                    strDescripFileAnalitica = ""
                    
                    
                    'txtDescripMovimiento = strDescripCuenta
                    
                    'txtCodFile.Text = ""
                    'txtCodAnalitica.Text = ""
                    
                    If strIndAuxiliar = Valor_Indicador Then
                    'este se cambia y comenta
                        'cmdBusqueda(2).Enabled = True
                        cmdBusqueda(1).Enabled = True
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
                        cmdBusqueda(1).Enabled = False
                        txtDescripAuxiliar.Text = ""
                        strTipoAuxiliar = ""
                        strCodAuxiliar = ""
                    End If
                    
                    If strTipoFile = Valor_Caracter Then
                        'esto se cambia y comenta
                        'cmdBusqueda(1).Enabled = False
                        cmdBusqueda(2).Enabled = False
                    Else
                        'esto se cambia y comenta
                        'cmdBusqueda(1).Enabled = True
                         cmdBusqueda(2).Enabled = True
'                        intRegistro = ObtenerItemLista(arrTipoFile(), strTipoFile)
'                        If intRegistro >= 0 Then cboTipoFile.ListIndex = intRegistro
                    End If
                    
                Case 1
            
                    strCodFile = Trim(.iParams(1).Valor)
                    strCodAnalitica = Trim(.iParams(2).Valor)
                    strDescripFileAnalitica = Trim(.iParams(3).Valor)
                    strCodMoneda = Trim(.iParams(4).Valor)
                        
                    'txtCodFile.Text = strCodFile
                    'txtCodAnalitica.Text = strCodAnalitica
                
                    'If strTipoFile = Valor_File_Generico Then
'                        txtCodAnalitica.Enabled = True
'                        txtDescripFileAnalitica.Text = "Analítica Genérica"
                   ' Else
'                        txtDescripFileAnalitica.Text = strCodFile & "-" & strCodAnalitica & " - " & strDescripFileAnalitica
'                        txtCodAnalitica.Enabled = True 'False
                    'End If
                                        
'                    cboMonedaMovimiento.ListIndex = -1
'                    intRegistro = ObtenerItemLista(arrMonedaMovimiento(), strCodMoneda)
'                    If intRegistro >= 0 Then cboMonedaMovimiento.ListIndex = intRegistro
                
                Case 2
            
                     strCodAuxiliar = Trim(.iParams(1).Valor)
                     strDescripAuxiliar = Trim(.iParams(2).Valor)
                     
                     txtDescripAuxiliar.Text = strDescripAuxiliar
                     
                Case 3
                        
                     strCodPersonaContraparte = Trim(.iParams(1).Valor)
'                     txtPersonaContraparte.Text = Trim(.iParams(2).Valor)
                     strDescripPersonaContraparte = Trim(.iParams(2).Valor)
            
            End Select
        
        End If
            
       
    End With
    
    Set frmBus = Nothing

End Sub

Private Sub cmdContracuenta_Click()

'    Dim frmContracuenta As frmContracuenta
'
'    Set frmContracuenta = New frmContracuenta
'
'    frmContracuenta.strCodFondo = strCodFondo
'    frmContracuenta.strTipoFileContracuenta = strTipoFileContracuenta
'    frmContracuenta.strCodContracuenta = strCodContracuenta
'    frmContracuenta.strCodFileContracuenta = strCodFileContracuenta
'    frmContracuenta.strCodAnaliticaContracuenta = strCodAnaliticaContracuenta
'    frmContracuenta.strDescripContracuenta = strDescripContracuenta 'REVISAR
'    frmContracuenta.strDescripFileAnaliticaContracuenta = strDescripFileAnaliticaContracuenta 'REVISAR
'
'
'    frmContracuenta.Show 1
'
'    If frmContracuenta.blnOK Then
'        strTipoFileContracuenta = frmContracuenta.strTipoFileContracuenta
'        strCodContracuenta = frmContracuenta.strCodContracuenta
'        strCodFileContracuenta = frmContracuenta.strCodFileContracuenta
'        strCodAnaliticaContracuenta = frmContracuenta.strCodAnaliticaContracuenta
'        strDescripContracuenta = frmContracuenta.strDescripContracuenta
'        strDescripFileAnaliticaContracuenta = frmContracuenta.strDescripFileAnaliticaContracuenta
'        lblContracuenta.Caption = frmContracuenta.strCodContracuenta + " / " + frmContracuenta.strCodFileAnaliticaContracuenta
'    Else
'        lblContracuenta.Caption = strCodContracuenta + " / " + strCodFileContracuenta + "-" + strCodAnaliticaContracuenta
'    End If
'
'    Set frmContracuenta = Nothing


End Sub

Private Sub cmdQuitar_Click()

    Dim dblBookmark As Double
    Dim strCriterio As String
    
    If adoRegistroAux.RecordCount > 0 Then
    
        dblBookmark = adoRegistroAux.Bookmark
    
'        If adoRegistroAux.Fields("CodMonedaMovimiento") <> Codigo_Moneda_Local Then
'            strCriterio = "CodMonedaOrigen='" & Mid(strCodMonedaParPorDefecto, 1, 2) & "'" & _
'                          " AND CodMonedaCambio = '" & Mid(strCodMonedaParPorDefecto, 3, 2) & "'" & _
'                          " AND ValorTipoCambio = " & CDbl(txtTipoCambio.Text)
'            'If FindRecordset(adoRegistroAuxTC, strCriterio) Then
'                adoRegistroAuxTC.Delete adAffectCurrent
'            'End If
'        End If
'
        adoRegistroAux.Delete adAffectCurrent

        If adoRegistroAux.EOF Then
            adoRegistroAux.MovePrevious
            tdgDetalle.MovePrevious
        End If

        adoRegistroAux.Update

        If adoRegistroAux.RecordCount = 0 Then cmdQuitar.Enabled = False

        If adoRegistroAux.RecordCount > 0 And Not adoRegistroAux.BOF And Not adoRegistroAux.EOF And dblBookmark > 1 Then adoRegistroAux.Bookmark = dblBookmark - 1

        Call NumerarRegistros

        If adoRegistroAux.RecordCount > 0 And Not adoRegistroAux.BOF And Not adoRegistroAux.EOF Then adoRegistroAux.Bookmark = dblBookmark - 1

        If adoRegistroAux.RecordCount > 0 And Not adoRegistroAux.BOF And Not adoRegistroAux.EOF Then adoRegistroAux.Bookmark = dblBookmark - 1

        tdgDetalle.Refresh
    
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
    Call OcultarComponentes
    
    
'    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
'    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
            
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
 

End Sub

Private Sub OcultarComponentes()
    
    lblDescrip(1).Visible = False
    lblDescrip(2).Visible = False
    dtpFechaDesde.Visible = False
    dtpFechaHasta.Visible = False
    
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    Dim elemento As Object
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
    For Each elemento In Me.Controls
    
        If TypeOf elemento Is TDBGrid Then
            'Call FormatoGrilla(elemento)
        End If
    
    Next
            
End Sub
Private Sub CargarListas()
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0

    '*** Naturaleza ***
    strSQL = "SELECT ValorParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='NATCTA' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboNaturaleza, arrNaturaleza(), Valor_Caracter
        
End Sub

Private Sub CargarListasNew()
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondoNew, arrFondoNew(), Valor_Caracter
    
    If cboFondoNew.ListCount > 0 Then cboFondoNew.ListIndex = 0
    
    '*** Naturaleza ***
    strSQL = "SELECT ValorParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='NATCTA' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboNaturaleza, arrNaturaleza(), Valor_Caracter

End Sub

Private Sub InicializarValores()

    '*** Valores Iniciales ***
    tabAsiento.Tab = 0
    chkIndVigente.Value = vbUnchecked
    strEstado = Reg_Defecto
    
    Call InicializarVariables
    
    Call ConfiguraRecordsetAuxiliar
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 20
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 20
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 20
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 20
    
    chkIndVigente.Value = vbChecked
            
    Set cmdOpcion.FormularioActivo = Me
    cmdOpcion.Button(1).Visible = False
    cmdOpcion.Button(2).Visible = False
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub
Private Sub InicializarVariables()
   
    numSecMovimiento = 1
    strTipoProceso = "0"
    
    strIndVigente = Valor_Caracter

    strCodCuenta = Valor_Caracter
    'strCodFile = Valor_Caracter
    strCodAnalitica = Valor_Caracter
    strTipoFile = Valor_Caracter
    strDescripCuenta = Valor_Caracter
    strDescripFileAnalitica = Valor_Caracter
    
    strTipoPersonaContraparte = Valor_Caracter
    strCodPersonaContraparte = Valor_Caracter
    strDescripPersonaContraparte = Valor_Caracter
    
    strCodContracuenta = Valor_Caracter
    strCodFileContracuenta = Valor_Caracter
    strCodAnaliticaContracuenta = Valor_Caracter
    strDescripFileAnaliticaContracuenta = Valor_Caracter
    strTipoFileContracuenta = Valor_Caracter
    
    dtpFechaDesde.Value = gdatFechaActual
    dtpFechaHasta.Value = gdatFechaActual
    
    

End Sub
Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    Set frmProvisionCastigos = Nothing
    
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
    
    Dim objTipoCambioReemplazoXML       As DOMDocument60
    Dim strMsgError                     As String
    
    Dim objProvisionCastigoDetalleXML   As DOMDocument60
    Dim strProvisionCastigoDetalleXML   As String
    
    If strEstado = Reg_Adicion Or strEstado = Reg_Edicion Then
    
        If TodoOK() Then
            Dim intCantRegistros    As Integer, intRegistro         As Integer
            Dim adoRegistro         As ADODB.Recordset
            Dim strNumAsiento       As String, strFechaGrabar       As String
            
            Dim strFechaDesdeNew    As String, strFechaHastaNew     As String
            
            Me.MousePointer = vbHourglass
                                                
            With adoComm
                
                strFechaDesdeNew = Convertyyyymmdd(dtpFechaDesdeNew.Value) & Space(1) & Format(Time, "hh:ss")
                strFechaHastaNew = Convertyyyymmdd(dtpFechaHastaNew.Value) & Space(1) & Format(Time, "hh:ss")
                strCodCastigo = Trim(lblCodCastigo.Caption)

                Call XMLADORecordset(objProvisionCastigoDetalleXML, "ProvisionCastigoDetalle", "Detalle", adoRegistroAux, strMsgError)
                strProvisionCastigoDetalleXML = objProvisionCastigoDetalleXML.xml
                               
                '*** Cabecera ***
                .CommandText = "{ call up_ACManProvisionCastigoXML('" & _
                    strCodFondoNew & "','" & gstrCodAdministradora & "','" & strCodFileNew & "','" & _
                    strCodClaseInstrumentoNew & "','" & _
                    strCodCastigo & "','" & Trim(txtDescripCastigo.Text) & "','" & _
                    strFechaDesdeNew & "','" & strFechaHastaNew & "','" & _
                    "X','" & strProvisionCastigoDetalleXML & "','" & _
                    IIf(strEstado = Reg_Adicion, "I", "U") & "') }"
                adoConn.Execute .CommandText
                                                                                
            End With
            
            Set adoRegistroAux = Nothing
                
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

    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Castigos..."
    
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
    
    Select Case strModo
        Case Reg_Adicion
                        
            Call CargarListasNew
            
            Call InicializarVariables
            
            lblCodCastigo.Caption = "GENERADO" 'NumAleatorio(10)
            
            txtDescripCastigo.Text = Valor_Caracter
            txtDiaDesde.Text = Valor_Caracter
            txtDiaHasta.Text = Valor_Caracter
            txtValorPorcentaje.Text = Valor_Caracter
            
            cboMoneda.Enabled = True
            intRegistro = ObtenerItemLista(arrMoneda(), gstrCodMoneda)
            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
            
            Call chkMovContable_Click

            intRegistro = ObtenerItemLista(arrNaturaleza(), Codigo_Tipo_Naturaleza_Debe)
            If intRegistro >= 0 Then cboNaturaleza.ListIndex = intRegistro

            dtpFechaDesdeNew.Value = gdatFechaActual
            dtpFechaHastaNew.Value = "2999/12/31"
                        
            Call ConfiguraRecordsetAuxiliarTC
                        
            Call CargarDetalleGrilla
            
            Call TotalizarMovimientos
                        
        Case Reg_Edicion
            
            Dim adoRecordset As New ADODB.Recordset
            
'            cboMonedaMovimiento.ListIndex = -1
                        
'            strSQL = "SELECT FechaAsiento, dbo.uf_ACObtenerHoraFecha(FechaAsiento) AS HoraAsiento, " & _
'                     "TipoDocumento, NumDocumento, DescripAsiento, " & _
'                     "MontoAsiento, CodMoneda, ValorTipoCambio " & _
'                     "FROM AsientoContable AC " & _
'                     "WHERE " & _
'                     "AC.CodFondo = '" & strCodFondo & "' AND " & _
'                     "AC.CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
'                     "AC.NumAsiento = '" & tdgConsulta.Columns("NumAsiento") & "'"
'
'            adoComm.CommandText = strSQL
                     
'            Set adoRecordset = adoComm.Execute
'
'            If Not adoRecordset.EOF Then
            
'                lblNumAsiento.Caption = tdgConsulta.Columns("NumAsiento")
'                txtDescripAsiento.Text = adoRecordset.Fields("DescripAsiento") 'tdgConsulta.Columns(2)
                'txtDescripAsiento.Enabled = False
                
'                strTipoDocumento = adoRecordset.Fields("TipoDocumento")
'                strNumDocumento = adoRecordset.Fields("NumDocumento")
                
''                txtNumDocumento.Text = strNumDocumento
                
'                intRegistro = ObtenerItemLista(arrTipoDocumento(), strTipoDocumento)
'                If intRegistro >= 0 Then cboTipoDocumento.ListIndex = intRegistro
                
'                cboModulo.Enabled = False
                
'                intRegistro = ObtenerItemLista(arrModulo(), frmMainMdi.Tag)
'                If intRegistro >= 0 Then cboModulo.ListIndex = intRegistro
                
'                dtpFechaAsiento.Value = adoRecordset.Fields("FechaAsiento")
'                txtHoraAsiento.Text = adoRecordset.Fields("HoraAsiento")
'
'                txtMontoAsiento.Text = adoRecordset.Fields("MontoAsiento")
'
'                txtTipoCambio.Text = adoRecordset.Fields("ValorTipoCambio")
'                txtTipoCambio.Enabled = True
                
'                cboMoneda.Enabled = False
'
'                intRegistro = ObtenerItemLista(arrMoneda(), adoRecordset.Fields("CodMoneda"))
'                If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
'
'                intRegistro = ObtenerItemLista(arrMonedaContable(), adoRecordset.Fields("CodMonedaContable"))
'                If intRegistro >= 0 Then cboMonedaContable.ListIndex = intRegistro
'
'                Call ConfiguraRecordsetAuxiliarTC
'
'                Call CargarDetalleGrilla
'
'                Call TotalizarMovimientos
'
                'adoRegistroAux.MoveFirst
            
'            Else
'                MsgBox "El Sistema no puede encontrar el comprobante contable para consultar!", vbExclamation
'                Exit Sub
'            End If
    
                                                            
        End Select
        
        
    
End Sub

Private Sub CargarDetalleGrilla()
    
    Dim adoRegistro As ADODB.Recordset
    Dim adoField As ADODB.Field
    
    Dim strSQL As String
    
    Set adoRegistro = New ADODB.Recordset
        
    Call ConfiguraRecordsetAuxiliar
    
    If strEstado = Reg_Edicion Then
        
'        strSQL = "{ call up_CNLstAsientoContableDetalle ('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
'                            Trim(lblNumAsiento.Caption) & "','" & strTipoProceso & "')}"

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
    
    tdgDetalle.DataSource = adoRegistroAux
    
    'If adoRegistroAux.RecordCount > 0 Then strEstado = Reg_Consulta
            
End Sub


Private Function TodoOK() As Boolean

    TodoOK = False
    
'    If Trim(txtDescripAsiento.Text) = Valor_Caracter Then
'        MsgBox "Descripción de asiento no ingresada", vbCritical, gstrNombreEmpresa
'        txtDescripAsiento.SetFocus
'        Exit Function
'    End If
'
'    If CDbl(txtTipoCambio.Text) = 0 Then
'        MsgBox "Tipo de Cambio no ingresado", vbCritical, gstrNombreEmpresa
'        txtTipoCambio.SetFocus
'        Exit Function
'    End If
    
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
    
'    If cboMoneda.ListIndex < 0 Then
'        MsgBox "Seleccione la moneda", vbCritical, gstrNombreEmpresa
'        cboMoneda.SetFocus
'        Exit Function
'    End If
'
'    If adoRegistroAux.EOF And adoRegistroAux.BOF Then
'        MsgBox "Comprobante sin registros", vbCritical, gstrNombreEmpresa
'        Exit Function
'    End If
'
'    '*** Validar cuadre del asiento y generar movimiento por diferencia ***
'    If Not ValidaCuadreContable() Then
'        MsgBox "Comprobante descuadrado", vbCritical, gstrNombreEmpresa
'        Exit Function
'    End If

    '*** Si todo paso OK ***
    TodoOK = True
  
End Function


Private Function TodoOkMovimiento() As Boolean

    TodoOkMovimiento = False
    
    'Si estoy ingresando el 1er registro pasa defrente
    If adoRegistroAux.EOF Then
        TodoOkMovimiento = True
        Exit Function
    End If
    
    While Not adoRegistroAux.EOF
        adoRegistroAux.MoveNext
    Wend
    
    'Comparare con el ultimo registro ingresado
    adoRegistroAux.MovePrevious
    
    If Not (adoRegistroAux.Fields("DiaHasta") = CInt(Trim(txtDiaDesde.Text)) - 1) Then
        MsgBox "El día inicial debe ser una dia mas del ultimo rango ingresado", vbCritical, gstrNombreEmpresa
        txtDiaDesde.SetFocus
        Exit Function
    End If
    
    If (CInt(txtDiaDesde.Text) >= CInt(txtDiaHasta.Text)) Then
        MsgBox "El día inicial debe ser menor que el dia final", vbCritical, gstrNombreEmpresa
        txtDiaDesde.SetFocus
        Exit Function
    End If
        
    '*** Si todo pasó OK ***
    TodoOkMovimiento = True
  
End Function


Private Sub lblTotalDebeME_Change()

    'Call FormatoMillarEtiqueta(lblTotalDebeME, Decimales_Monto)
    
End Sub

Private Sub lblTotalDebeMN_Change()

'    Call FormatoMillarEtiqueta(lblTotalDebeMN, Decimales_Monto)
    
End Sub

Private Sub lblTotalHaberME_Change()

'    Call FormatoMillarEtiqueta(lblTotalHaberME, Decimales_Monto)
    
End Sub

Private Sub lblTotalHaberMN_Change()

'    Call FormatoMillarEtiqueta(lblTotalHaberMN, Decimales_Monto)
    
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

Private Sub tdgDetalle_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If tdgDetalle.Columns(ColIndex).DataField = "MontoDebe" Or _
       tdgDetalle.Columns(ColIndex).DataField = "MontoHaber" Or _
       tdgDetalle.Columns(ColIndex).DataField = "MontoContable" Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
   
End Sub


Private Sub tdgDetalle_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    
'    If adoRegistroAux.EOF Then Exit Sub 'And adoRegistroAux.BOF
'
'    lblNumAsiento.Caption = adoRegistroAux.Fields("NumAsiento")
'    'txtDescripMovimiento.Text = adoRegistroAux.Fields("DescripMovimiento")
'    txtCodCuenta.Text = Trim(adoRegistroAux.Fields("CodCuenta"))
'    txtCodFile.Text = adoRegistroAux.Fields("CodFile")
'    txtCodAnalitica.Text = adoRegistroAux.Fields("CodAnalitica")
'    'txtDescripCuenta.Text = Trim(txtCodCuenta.Text) & " - " & ObtenerDescripcionCuenta(Trim(txtCodCuenta.Text))
'
'    intRegistro = ObtenerItemLista(arrNaturaleza(), adoRegistroAux.Fields("IndDebeHaber"))
'    If intRegistro >= 0 Then cboNaturaleza.ListIndex = intRegistro
'
'    intRegistro = ObtenerItemLista(arrMonedaMovimiento(), adoRegistroAux.Fields("CodMonedaMovimiento"))
'    If intRegistro >= 0 Then cboMonedaMovimiento.ListIndex = intRegistro
'
'    If strCodNaturaleza = "D" Then
'        txtMontoMovimiento.Text = adoRegistroAux.Fields("MontoDebe")
'    Else
'        txtMontoMovimiento.Text = adoRegistroAux.Fields("MontoHaber")
'    End If
'
'    intRegistro = ObtenerItemLista(arrMonedaContable(), adoRegistroAux.Fields("CodMonedaContable"))
'    If intRegistro >= 0 Then cboMonedaContable.ListIndex = intRegistro
'
'    txtMontoContable.Text = adoRegistroAux.Fields("MontoContable")
'
'    intRegistro = ObtenerItemLista(arrTipoDocumentoDet(), adoRegistroAux.Fields("TipoDocumento"))
'    If intRegistro >= 0 Then cboTipoDocumentoDet.ListIndex = intRegistro
'
'    txtNumDocumentoDet.Text = adoRegistroAux.Fields("NumDocumento")
'
'    intRegistro = ObtenerItemLista(arrTipoPersonaContraparte(), adoRegistroAux.Fields("TipoPersonaContraparte"))
'    If intRegistro >= 0 Then cboTipoPersonaContraparte.ListIndex = intRegistro
'
'    strCodPersonaContraparte = adoRegistroAux.Fields("CodPersonaContraparte")
'    strDescripPersonaContraparte = adoRegistroAux.Fields("DescripPersonaContraparte")
'    txtPersonaContraparte.Text = strDescripPersonaContraparte
'
'    txtTipoCambio.Text = CStr(adoRegistroAux.Fields("ValorTipoCambio"))
'
'    If adoRegistroAux.Fields("IndSoloMovimientoContable") = Valor_Indicador Then
'        chkMovContable.Value = vbChecked
'        Call chkMovContable_Click
'    Else
'        chkMovContable.Value = vbUnchecked
'        Call chkMovContable_Click
'    End If
'
'    If adoRegistroAux.Fields("IndContracuenta") = Valor_Indicador Then
'        chkContracuenta.Value = vbChecked
'    Else
'        chkContracuenta.Value = vbUnchecked
'    End If
'
'    If chkContracuenta.Value = vbChecked Then
'        strCodContracuenta = adoRegistroAux.Fields("CodContracuenta")
'        strCodFileContracuenta = adoRegistroAux.Fields("CodFileContracuenta")
'        strCodAnaliticaContracuenta = adoRegistroAux.Fields("CodAnaliticaContracuenta")
'        strDescripContracuenta = adoRegistroAux.Fields("DescripContracuenta")
'        strDescripFileAnaliticaContracuenta = adoRegistroAux.Fields("DescripFileAnaliticaContracuenta")
'        strTipoFileContracuenta = adoRegistroAux.Fields("TipoFileContracuenta")
'
'        lblContracuenta.Caption = strCodContracuenta + " / " + strCodFileContracuenta + "-" + strCodAnaliticaContracuenta
'    Else
'        strCodContracuenta = Valor_Caracter
'        strCodFileContracuenta = Valor_Caracter
'        strCodAnaliticaContracuenta = Valor_Caracter
'        strDescripContracuenta = Valor_Caracter
'        strDescripContracuenta = Valor_Caracter
'        strDescripFileAnaliticaContracuenta = Valor_Caracter
'        strTipoFileContracuenta = Valor_Caracter
'
'        lblContracuenta.Caption = Valor_Caracter
'    End If

End Sub

Private Sub txtCodAnalitica_LostFocus()

'    txtCodAnalitica.Text = Right(String(8, "0") & Trim(txtCodAnalitica.Text), 8)
'
'    strCodAnalitica = txtCodAnalitica.Text
            
End Sub


Private Sub txtCodCuenta_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N")
    
End Sub


Private Sub txtCodFile_LostFocus()

'    txtCodFile.Text = Right(String(3, "0") & Trim(txtCodFile.Text), 3)
'
'    strCodFile = txtCodFile.Text
    
End Sub


Private Sub txtMontoAsiento_Change()

    Call FormatoCajaTexto(txtMontoAsiento, Decimales_Monto)
    
End Sub

Private Sub txtMontoMovimiento_Change()

    'Call FormatoCajaTexto(txtMontoMovimiento, Decimales_Monto)
    Call Calcular
    
End Sub
Private Sub Calcular()

'    If Not IsNumeric(txtTipoCambio.Text) Or Not IsNumeric(txtMontoMovimiento.Value) Then Exit Sub
'
'    If strIndSoloMovimientoContable = Valor_Caracter Then
'        If strCodMonedaMovimiento = Codigo_Moneda_Local Then
'            txtMontoContable.Text = CStr(txtMontoMovimiento.Value)
'        Else
'            txtMontoContable.Text = CStr(txtMontoMovimiento.Value * CDbl(txtTipoCambio.Text))
'        End If
'    End If
    
End Sub



Private Sub txtMontoMovimiento_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call txtMontoMovimiento_Change
    End If

    
End Sub


Private Sub txtMontoMovimiento_LostFocus()

'    txtMontoMovimiento.Text = Abs(txtMontoMovimiento.Value)
'    If strCodNaturaleza = Codigo_Tipo_Naturaleza_Haber Then
'        txtMontoMovimiento.Text = Abs(txtMontoMovimiento.Value) * -1
'    End If
'
End Sub

Private Sub txtTipoCambio_Change()

    'Call FormatoCajaTexto(txtTipoCambio, Decimales_TipoCambio)
    Call Calcular

End Sub


Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)

    'Call ValidaCajaTexto(KeyAscii, "M", txtTipoCambio, Decimales_TipoCambio)
    If KeyAscii = 13 Then
        Call txtTipoCambio_Change
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
       .Fields.Append "SecCastigo", adInteger
       .Fields.Append "DiaDesde", adInteger
       .Fields.Append "DiaHasta", adInteger
       .Fields.Append "ValorPorcentaje", adDecimal, 9
       .LockType = adLockBatchOptimistic
    End With


    With adoRegistroAux.Fields.Item("ValorPorcentaje")
        .Precision = 6
        .NumericScale = 2
    End With
    
    adoRegistroAux.Open

End Sub

