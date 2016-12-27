VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmOrdenPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes de Pago"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
   Icon            =   "frmOrdenPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9930
   ScaleWidth      =   11580
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   9480
      TabIndex        =   3
      Top             =   9120
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
      TabIndex        =   4
      Top             =   9120
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
   Begin VB.CommandButton cmdAgregar 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Agregar detalle"
      Top             =   9960
      Width           =   375
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Quitar detalle"
      Top             =   10320
      Width           =   375
   End
   Begin TabDlg.SSTab tabCatalogo 
      Height          =   8865
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   15637
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
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
      TabPicture(0)   =   "frmOrdenPago.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraBusqueda"
      Tab(0).Control(1)=   "tdgConsulta"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmOrdenPago.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraDetalle"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAccion"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   6600
         TabIndex        =   1
         Top             =   8040
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
      Begin VB.Frame fraBusqueda 
         Caption         =   "Criterios de Búsqueda"
         Height          =   1545
         Left            =   -74640
         TabIndex        =   10
         Top             =   750
         Width           =   10485
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   510
            Width           =   7725
         End
         Begin MSComCtl2.DTPicker dtpFechaInicial 
            Height          =   315
            Left            =   5790
            TabIndex        =   33
            Top             =   1020
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            Format          =   50397185
            CurrentDate     =   41387
         End
         Begin MSComCtl2.DTPicker dtpFechaFinal 
            Height          =   315
            Left            =   7950
            TabIndex        =   35
            Top             =   1020
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            Format          =   50397185
            CurrentDate     =   41387
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Al"
            BeginProperty Font 
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
            Left            =   7620
            TabIndex        =   36
            Top             =   1080
            Width           =   180
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Del"
            BeginProperty Font 
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
            Left            =   5340
            TabIndex        =   34
            Top             =   1080
            Width           =   300
         End
         Begin VB.Label lblDescrip 
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
            Height          =   285
            Index           =   0
            Left            =   360
            TabIndex        =   12
            Top             =   540
            Width           =   735
         End
      End
      Begin VB.Frame fraDetalle 
         Height          =   7335
         Left            =   330
         TabIndex        =   2
         Top             =   660
         Width           =   10755
         Begin VB.CommandButton cmdActualizar 
            Caption         =   "A"
            Height          =   375
            Left            =   300
            TabIndex        =   41
            Top             =   5220
            Width           =   375
         End
         Begin VB.CommandButton Command2 
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
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Quitar detalle"
            Top             =   6060
            Width           =   375
         End
         Begin VB.CommandButton Command1 
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
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Agregar detalle"
            Top             =   5640
            Width           =   375
         End
         Begin VB.ComboBox cboEstado 
            Height          =   315
            Left            =   8250
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   2010
            Width           =   1815
         End
         Begin VB.ComboBox cboGasto 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   2760
            Width           =   8115
         End
         Begin MSComCtl2.DTPicker dtpFechaOrden 
            Height          =   315
            Left            =   1950
            TabIndex        =   18
            Top             =   600
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            Format          =   50397185
            CurrentDate     =   41387
         End
         Begin VB.CommandButton cmdProveedor 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10050
            TabIndex        =   14
            ToolTipText     =   "Buscar Proveedor"
            Top             =   1050
            Width           =   375
         End
         Begin TAMControls.TAMTextBox txtPorcenComision 
            Height          =   315
            Left            =   1980
            TabIndex        =   17
            Top             =   4140
            Width           =   1785
            _ExtentX        =   3149
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
            Container       =   "frmOrdenPago.frx":0044
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
         Begin TAMControls.TAMTextBox TAMTextBox1 
            Height          =   315
            Left            =   1980
            TabIndex        =   24
            Top             =   4590
            Width           =   1785
            _ExtentX        =   3149
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
            Container       =   "frmOrdenPago.frx":0060
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
         Begin TrueOleDBGrid60.TDBGrid TDBGrid1 
            Height          =   1815
            Left            =   870
            OleObjectBlob   =   "frmOrdenPago.frx":007C
            TabIndex        =   42
            Top             =   5190
            Width           =   9405
         End
         Begin TAMControls.TAMTextBox TAMTextBox2 
            Height          =   315
            Left            =   1950
            TabIndex        =   43
            Top             =   1500
            Width           =   1785
            _ExtentX        =   3149
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
            Container       =   "frmOrdenPago.frx":40C1
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
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "099-00000001"
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   8250
            TabIndex        =   47
            Top             =   1500
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "File / Analitica"
            BeginProperty Font 
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
            Left            =   6840
            TabIndex        =   46
            Top             =   1530
            Width           =   1260
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Provisión Total"
            BeginProperty Font 
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
            Left            =   330
            TabIndex        =   45
            Top             =   1530
            Width           =   1275
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Index           =   15
            Left            =   3750
            TabIndex        =   44
            Top             =   1560
            Width           =   390
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Estado Orden"
            BeginProperty Font 
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
            TabIndex        =   38
            Top             =   2040
            Width           =   1170
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "File / Analitica"
            BeginProperty Font 
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
            Left            =   6660
            TabIndex        =   32
            Top             =   3210
            Width           =   1260
         End
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "098-00000001"
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   8310
            TabIndex        =   31
            Top             =   3180
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Periodo Gasto"
            BeginProperty Font 
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
            TabIndex        =   30
            Top             =   3690
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1"
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   1980
            TabIndex        =   29
            Top             =   3660
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Num. Gasto"
            BeginProperty Font 
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
            TabIndex        =   28
            Top             =   3240
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "15"
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   1980
            TabIndex        =   27
            Top             =   3210
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Mto. Orden Pago"
            BeginProperty Font 
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
            TabIndex        =   26
            Top             =   4620
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Index           =   7
            Left            =   3780
            TabIndex        =   25
            Top             =   4620
            Width           =   390
         End
         Begin VB.Line Line3 
            X1              =   330
            X2              =   10260
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
            BeginProperty Font 
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
            Left            =   360
            TabIndex        =   23
            Top             =   2820
            Width           =   825
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Num. Orden"
            BeginProperty Font 
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
            Left            =   6840
            TabIndex        =   21
            Top             =   630
            Width           =   1020
         End
         Begin VB.Label lblNumOrden 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "GENERADO"
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   8250
            TabIndex        =   20
            Top             =   600
            Visible         =   0   'False
            Width           =   1800
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
            Index           =   4
            Left            =   360
            TabIndex        =   19
            Top             =   630
            Width           =   1110
         End
         Begin VB.Label lblDescripProveedor 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1950
            TabIndex        =   15
            Top             =   1050
            Width           =   8100
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Index           =   2
            Left            =   3780
            TabIndex        =   13
            Top             =   4200
            Width           =   390
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            BeginProperty Font 
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
            Left            =   330
            TabIndex        =   7
            Top             =   1110
            Width           =   885
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Mto. Provisión"
            BeginProperty Font 
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
            TabIndex        =   6
            Top             =   4170
            Width           =   1215
         End
         Begin VB.Label lblCodProveedor 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1980
            TabIndex        =   16
            Top             =   1050
            Visible         =   0   'False
            Width           =   1380
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmOrdenPago.frx":40DD
         Height          =   5415
         Left            =   -74640
         OleObjectBlob   =   "frmOrdenPago.frx":40F7
         TabIndex        =   5
         Top             =   2610
         Width           =   10515
      End
   End
End
Attribute VB_Name = "frmOrdenPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()      As String, arrSucursal()    As String
Dim arrAgencia()    As String

Dim strCodFondo     As String, strCodSucursal   As String
Dim strCodAgencia   As String

Dim strSQL          As String
'Dim adoRegistroAux  As ADODB.Recordset
Dim adoRegistro As ADODB.Recordset
Dim intRegistro     As String
'Dim vntTmp          As Variant
Dim strEstado       As String




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

Public Sub Adicionar()

    If cboFondo.ListIndex < 0 Then Exit Sub
    
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Datos de Catálogo..."
                    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabCatalogo
        .TabEnabled(0) = False
        .Tab = 1
    End With
    
End Sub

Public Sub Buscar()
   
   
    Set adoRegistro = New ADODB.Recordset
    
'    Set adoRegistroAux = New ADODB.Recordset
'
'
'    With adoRegistroAux
'       .CursorLocation = adUseClient
'       .Fields.Append "CodSucursal", adChar, 3
'       .Fields.Append "DescripSucursal", adVarChar, 25
'       .Fields.Append "CodAgencia", adChar, 6
'       .Fields.Append "DescripAgencia", adVarChar, 50
'       .Fields.Append "HoraInicio", adChar, 5
'       .Fields.Append "HoraTermino", adChar, 5
'       .CursorType = adOpenStatic
'       .LockType = adLockBatchOptimistic
'    End With
'
'    adoRegistroAux.Open

    
    strSQL = "{ call up_ACSelDatosParametro(54,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"

    
    With adoRegistro
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
        '.ActiveConnection = Nothing
        
'        If .RecordCount > 0 Then
'            .MoveFirst
'            Do While Not .EOF
'                adoRegistroAux.AddNew Array("CodSucursal", "DescripSucursal", "CodAgencia", "DescripAgencia", "HoraInicio", "HoraTermino"), Array(.Fields("CodSucursal").Value, .Fields("DescripSucursal").Value, .Fields("CodAgencia").Value, .Fields("DescripAgencia").Value, .Fields("HoraInicio").Value, .Fields("HoraTermino").Value)
'                adoRegistro.MoveNext
'            Loop
'            adoRegistroAux.MoveFirst
'        End If
        
    End With
    
    tdgConsulta.DataSource = adoRegistro
    
    'tdgConsulta.Refresh
    'tdgConsulta.ReBind
    
    If adoRegistro.RecordCount > 0 Then strEstado = Reg_Consulta
    
'    Me.Refresh
'
'    DoEvents
        
End Sub

Private Sub CargarListas()

    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
'    strSQL = "SELECT CodSucursal CODIGO, DescripSucursal DESCRIP FROM SucursalBancaria"
'    CargarControlLista strSQL, cboSucursal, arrSucursal(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    

    
End Sub
Public Sub Cancelar()

    fraDetalle.Enabled = True
    cmdOpcion.Visible = True
    With tabCatalogo
        .TabEnabled(0) = True
        .Tab = 0
    End With
    Call Buscar
    
    'strEstado = Reg_Consulta
    
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    
'    For intCont = 0 To (fraDetalle.Count - 1)
'        Call FormatoMarco(fraDetalle(intCont))
'    Next
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
End Sub

Public Sub Eliminar()


    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        strEstado = Reg_Eliminacion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabCatalogo
            .TabEnabled(0) = False
            .Tab = 1
        End With
    End If
    
    
End Sub


Public Sub Grabar()
    
    Dim intAccion               As Integer
    Dim lngNumError             As Integer
    Dim strCodComisionista      As String
    Dim dblPorcenComision       As Double
    
   ' On Error GoTo CtrlError

    If TodoOK() Then
        
        'strCodComisionista = lblCodComisionista.Caption
        'dblPorcenComision = txtPorcenComision.Value
    
        '*** Guardar ***
        With adoComm

            If strEstado = Reg_Eliminacion Then
                If MsgBox(Mensaje_Eliminacion, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            End If
            
            .CommandText = "{ call up_GNManFondoComisionista('" & strCodFondo & "','" & _
                gstrCodAdministradora & "','" & strCodComisionista & "'," & _
                dblPorcenComision & ",'" & IIf(strEstado = Reg_Adicion, "I", IIf(strEstado = Reg_Edicion, "U", "D")) & "') }"
            
            adoConn.Execute .CommandText
        
        End With
                                                                                                                    
        Me.MousePointer = vbDefault
                    
        If strEstado = Reg_Adicion Then
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
        End If
        
        If strEstado = Reg_Edicion Then
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
        End If
        
        If strEstado = Reg_Eliminacion Then
            MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation
        End If
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        cmdOpcion.Visible = True
        With tabCatalogo
            .TabEnabled(0) = True
            .Tab = 0
        End With
        
        Call Buscar
   
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

End Sub


Public Sub Imprimir()
    
    Call SubImprimir(1)
    
End Sub



Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabCatalogo
            .TabEnabled(0) = False
            .Tab = 1
        End With
        
    End If
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub






Private Sub cboFondo_Click()

    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Call Buscar

End Sub




Private Sub cmdAgregar_Click()

    
'    If TodoOK() Then
'        'If Not tdgDinamica.EOF Then
'            If adoRegistroAux.Supports(adAddNew) Then
'                adoRegistroAux.AddNew Array("CodSucursal", "DescripSucursal", "CodAgencia", "DescripAgencia", "HoraInicio", "HoraTermino"), Array(strCodSucursal, cboSucursal.List(cboSucursal.ListIndex), strCodAgencia, cboAgencia.List(cboAgencia.ListIndex), Format(dtpHoraInicio.Value, "hh:mm"), Format(dtpHoraTermino.Value, "hh:mm"))
'            End If
'        'End If
'    End If
'
'
'    tdgDinamica.Refresh
'
   
    
End Sub
Private Function TodoOK() As Boolean

    TodoOK = False
    
'    If Trim(lblCodComisionista.Caption) = Valor_Caracter Then
'        MsgBox "Debe seleccionar un comisionista!", vbOKOnly + vbExclamation, Me.Caption
'        Exit Function
'    End If
    
    If txtPorcenComision.Value <= 0 Then
        MsgBox "Debe ingresar una comisión valida!", vbOKOnly + vbExclamation, Me.Caption
        Exit Function
    End If

    TodoOK = True


End Function
        


Private Sub LlenarFormulario(strModo As String)


    
'    With adoRegistroAux
'        If .RecordCount > 0 Then
'            If .EOF Or .BOF Then
'               .MoveFirst
'            End If
'
'            'CodSucursal,CodAgencia,HoraInicio,HoraTermino
'
'            intRegistro = ObtenerItemLista(arrSucursal(), .Fields("CodSucursal").Value)
'            If intRegistro >= 0 Then cboSucursal.ListIndex = intRegistro
'
'            dtpHoraInicio.Value = .Fields("HoraInicio").Value
'            dtpHoraTermino.Value = .Fields("HoraTermino").Value
'
'        End If
'
'    End With


    Select Case strModo
        Case Reg_Adicion
'            lblDescripComisionista.Caption = Valor_Caracter
'            lblCodComisionista.Caption = Valor_Caracter
'            txtPorcenComision.Text = "0.0000"
        
        Case Reg_Edicion, Reg_Eliminacion
            
'            lblDescripComisionista.Caption = tdgConsulta.Columns("DescripComisionista").Value
'            lblCodComisionista.Caption = tdgConsulta.Columns("CodComisionista").Value
'            txtPorcenComision.Text = tdgConsulta.Columns("PorcenComision").Value
            
            If strModo = Reg_Eliminacion Then
                fraDetalle.Enabled = False
            End If
            
    
    End Select

End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabCatalogo.Tab = 0
    
    '*** Ancho por defecto de las columnas de la grilla ***
'    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 14
'    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 60
'    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 10
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String

    If tabCatalogo.Tab = 1 Then Exit Sub

    Select Case Index
        Case 1
            gstrNameRepo = "Catalogo"

            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(1)
            ReDim aReportParamFn(2)
            ReDim aReportParamF(2)

            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"

            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)

'            If strCodConcepto = Valor_Caracter Then
'                aReportParamS(0) = strCodConcepto
'                aReportParamS(1) = Codigo_Listar_Todos
'            Else
'                aReportParamS(0) = strCodConcepto
'                aReportParamS(1) = Codigo_Listar_Individual
'            End If

    End Select

    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub
Private Sub CargarReportes()

    
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Vista Activa"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    
End Sub



Private Sub cmdProveedor_Click()

  Dim sSql As String
   
    Screen.MousePointer = vbHourglass
   
    Dim frmBus As frmBuscar
    
    Set frmBus = New frmBuscar
    
    With frmBus.TBuscarRegistro1
           
        .ADOConexion = adoConn
        .ADOConexion.CommandTimeout = 0
        .iTipoGrilla = 2
        
        frmBus.Caption = "Relación de Proveedores"
        .sSql = "{ call up_ACSelDatos(45) }"
        
        .OutputColumns = "1,2,3,4,5,6"
        .HiddenColumns = "1,2,6"
        
        .BuscarTabla
        
        Screen.MousePointer = vbNormal
        frmBus.Show 1
       
        If .iParams.Count = 0 Then Exit Sub
        
        If .iParams(1).Valor <> "" Then
            lblDescripProveedor.Caption = .iParams(5).Valor
            lblCodProveedor.Caption = .iParams(1).Valor
        
            strSQL = "SELECT FCG.CodCuenta CODIGO,(RTRIM(DescripCuenta)) DESCRIP " & _
                "FROM FondoGastoDevengo " & _
                "FCG JOIN PlanContable PCG ON(PCG.CodCuenta=FCG.CodCuenta AND PCG.CodAdministradora=FCG.CodAdministradora) " & _
                "WHERE CodFondo='" & strCodFondo & "' AND FCG.CodAdministradora='" & gstrCodAdministradora & "' AND PCG.NumVersion = dbo.uf_CNObtenerPlanContableVigente('" & gstrCodAdministradora & "') " & _
                "ORDER BY DescripCuenta"
        
            'CargarControlLista strSQL, cboGasto, arrGasto(), Sel_Defecto
        
        End If
            
       
    End With
    
    Set frmBus = Nothing



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
    Call DarFormato
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)

    CentrarForm Me
        
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmOrdenPago = Nothing
    
End Sub


Private Sub tabCatalogo_Click(PreviousTab As Integer)

    Select Case tabCatalogo.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabCatalogo.Tab = 0
        
    End Select

End Sub


Private Sub tdgConsulta_DblClick()

    If strEstado <> Reg_Consulta Or tdgConsulta.Bookmark = 0 Then Exit Sub
    Call Accion(vModify)

End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 2 Then
        Call DarFormatoValor(Value, Decimales_Tasa)
    End If

End Sub
