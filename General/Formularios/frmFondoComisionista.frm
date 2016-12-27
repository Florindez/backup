VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmFondoComisionista 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comisionistas por Fondo"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16980
   Icon            =   "frmFondoComisionista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   16980
   Begin TabDlg.SSTab tabCatalogo 
      Height          =   8235
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   16905
      _ExtentX        =   29819
      _ExtentY        =   14526
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
      TabPicture(0)   =   "frmFondoComisionista.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(1)=   "fraBusqueda"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmFondoComisionista.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "cmdAcciones"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraDetalle"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fraDetalle 
         Caption         =   "Condiciones de Comisiones"
         Height          =   6795
         Left            =   300
         TabIndex        =   28
         Top             =   570
         Width           =   16215
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   2310
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Tag             =   "TIPO_TASA"
            Top             =   1230
            Width           =   2565
         End
         Begin VB.CommandButton cmdActualizar 
            Height          =   575
            Left            =   360
            Picture         =   "frmFondoComisionista.frx":0044
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Actualizar Detalle"
            Top             =   4980
            Width           =   495
         End
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
            Height          =   550
            Left            =   360
            Picture         =   "frmFondoComisionista.frx":02FF
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Quitar detalle"
            Top             =   6120
            Width           =   495
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
            Height          =   575
            Left            =   360
            Picture         =   "frmFondoComisionista.frx":0551
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Agregar detalle"
            Top             =   5550
            Width           =   495
         End
         Begin VB.CheckBox chkComisionistaActivo 
            Caption         =   "Condiciones Activas?"
            DataSource      =   "iv"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   11070
            TabIndex        =   4
            Top             =   1260
            Width           =   2295
         End
         Begin VB.Frame Frame5 
            Caption         =   "Pagos"
            Height          =   1605
            Left            =   8100
            TabIndex        =   52
            Top             =   3390
            Width           =   7755
            Begin VB.ComboBox cboTipoDesplazamiento 
               Height          =   315
               Left            =   1950
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   750
               Width           =   2595
            End
            Begin VB.ComboBox cboFrecuenciaPago 
               Height          =   315
               ItemData        =   "frmFondoComisionista.frx":07FE
               Left            =   2460
               List            =   "frmFondoComisionista.frx":0800
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   360
               Width           =   2085
            End
            Begin VB.CheckBox chkFinDeMes 
               Caption         =   "Fin de Mes"
               BeginProperty Font 
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
               Left            =   4740
               TabIndex        =   18
               Top             =   360
               Width           =   1515
            End
            Begin TAMControls.TAMTextBox txtPagoCada 
               Height          =   315
               Left            =   1950
               TabIndex        =   16
               Top             =   360
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
               Container       =   "frmFondoComisionista.frx":0802
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
               AutoSize        =   -1  'True
               Caption         =   "Desplazamiento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   54
               Top             =   810
               Width           =   1350
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cada"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   360
               TabIndex        =   53
               Top             =   420
               Width           =   450
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Devengo"
            Height          =   1605
            Left            =   360
            TabIndex        =   47
            Top             =   3390
            Width           =   7755
            Begin VB.ComboBox cboPeriodoDevengo 
               Height          =   315
               Left            =   1950
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   780
               Width           =   2565
            End
            Begin VB.ComboBox cboModalidadDevengo 
               Height          =   315
               Left            =   1950
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Tag             =   "MODO_DEVENGO"
               Top             =   360
               Width           =   2565
            End
            Begin VB.CommandButton cmdCondicionDevengo 
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
               Index           =   0
               Left            =   6870
               TabIndex        =   15
               ToolTipText     =   "Buscar Proveedor"
               Top             =   1170
               Width           =   315
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Periodo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   51
               Top             =   810
               Width           =   660
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Modalidad"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   50
               Top             =   390
               Width           =   885
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Formula"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   49
               Top             =   1200
               Width           =   675
            End
            Begin VB.Label lblFormulaDevengo 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1950
               TabIndex        =   48
               Top             =   1170
               Width           =   4905
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Comision"
            Height          =   1215
            Left            =   360
            TabIndex        =   40
            Top             =   2280
            Width           =   15495
            Begin VB.ComboBox cboPeriodoTasa 
               Height          =   315
               Left            =   7020
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   330
               Width           =   2565
            End
            Begin VB.ComboBox cboTipoTasa 
               Height          =   315
               Left            =   1950
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Tag             =   "TIPO_TASA"
               Top             =   750
               Width           =   2565
            End
            Begin VB.ComboBox cboBaseCalculo 
               Height          =   315
               Left            =   12120
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   300
               Width           =   2565
            End
            Begin VB.ComboBox cboPeriodoCapitalizacion 
               Height          =   315
               Left            =   7020
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Tag             =   "TIPO_CAPITALIZA"
               Top             =   750
               Width           =   2565
            End
            Begin TAMControls.TAMTextBox txtPorcenTasa 
               Height          =   315
               Left            =   1950
               TabIndex        =   8
               Top             =   330
               Width           =   2265
               _ExtentX        =   3995
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
               Container       =   "frmFondoComisionista.frx":081E
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
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Periodo Tasa"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   5460
               TabIndex        =   46
               Top             =   390
               Width           =   1140
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
               Index           =   4
               Left            =   390
               TabIndex        =   45
               Top             =   810
               Width           =   870
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "(%)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   4230
               TabIndex        =   44
               Top             =   360
               Width           =   270
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
               Index           =   3
               Left            =   360
               TabIndex        =   43
               Top             =   390
               Width           =   435
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Base Calculo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   10410
               TabIndex        =   42
               Top             =   330
               Width           =   1125
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Periodo Capitaliza"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   5460
               TabIndex        =   41
               Top             =   810
               Width           =   1545
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Vigencia"
            Height          =   735
            Left            =   360
            TabIndex        =   35
            Top             =   1620
            Width           =   15495
            Begin VB.CheckBox chkIndeterminado 
               Caption         =   "No determinado"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   10020
               TabIndex        =   7
               Top             =   360
               Width           =   1695
            End
            Begin MSComCtl2.DTPicker dtpFechaHasta 
               Height          =   315
               Left            =   8310
               TabIndex        =   6
               Top             =   300
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   556
               _Version        =   393216
               Format          =   175898625
               CurrentDate     =   42086
            End
            Begin MSComCtl2.DTPicker dtpFechaDesde 
               Height          =   315
               Left            =   1950
               TabIndex        =   5
               Top             =   300
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   556
               _Version        =   393216
               Format          =   175898625
               CurrentDate     =   42086
            End
            Begin VB.Label lblFechaDesde 
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
               Index           =   5
               Left            =   360
               TabIndex        =   37
               Top             =   360
               Width           =   555
            End
            Begin VB.Label lblFechaHasta 
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
               Index           =   0
               Left            =   6780
               TabIndex        =   36
               Top             =   330
               Width           =   510
            End
         End
         Begin VB.ComboBox cboTipoComisionista 
            Height          =   315
            Left            =   10140
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   420
            Width           =   3195
         End
         Begin VB.CommandButton cmdComisionista 
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
            Left            =   13020
            TabIndex        =   2
            ToolTipText     =   "Buscar Proveedor"
            Top             =   840
            Width           =   315
         End
         Begin TrueOleDBGrid60.TDBGrid tdgFondoComisionCondicion 
            Height          =   1785
            Left            =   960
            OleObjectBlob   =   "frmFondoComisionista.frx":083A
            TabIndex        =   29
            Top             =   4980
            Width           =   14895
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
            Index           =   14
            Left            =   390
            TabIndex        =   55
            Top             =   1260
            Width           =   690
         End
         Begin VB.Line Line1 
            X1              =   390
            X2              =   15870
            Y1              =   1590
            Y2              =   1590
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Codigo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            TabIndex        =   34
            Top             =   420
            Width           =   600
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Comisionista"
            BeginProperty Font 
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
            Left            =   8130
            TabIndex        =   33
            Top             =   450
            Width           =   1695
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisionista"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            TabIndex        =   32
            Top             =   840
            Width           =   1065
         End
         Begin VB.Label lblDescripComisionista 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2310
            TabIndex        =   31
            Top             =   810
            Width           =   10725
         End
         Begin VB.Label lblCodComisionista 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2310
            TabIndex        =   30
            Top             =   390
            Width           =   1620
         End
      End
      Begin VB.Frame fraBusqueda 
         Caption         =   "Criterios de Búsqueda"
         Height          =   1155
         Left            =   -74640
         TabIndex        =   24
         Top             =   660
         Width           =   16035
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   540
            Width           =   11145
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
            TabIndex        =   26
            Top             =   570
            Width           =   735
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmFondoComisionista.frx":C168
         Height          =   5655
         Left            =   -74640
         OleObjectBlob   =   "frmFondoComisionista.frx":C182
         TabIndex        =   27
         Top             =   2040
         Width           =   16005
      End
      Begin TAMControls2.ucBotonEdicion2 cmdAcciones 
         Height          =   735
         Left            =   13170
         TabIndex        =   23
         Top             =   7410
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
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   1050
      TabIndex        =   39
      Top             =   8280
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   1296
      Buttons         =   4
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Consultar"
      Tag1            =   "1"
      ToolTipText1    =   "Consultar"
      Caption2        =   "&Buscar"
      Tag2            =   "5"
      ToolTipText2    =   "Buscar"
      Caption3        =   "&Anular"
      Tag3            =   "4"
      ToolTipText3    =   "Anular"
      UserControlWidth=   5700
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   14670
      TabIndex        =   38
      Top             =   8280
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
End
Attribute VB_Name = "frmFondoComisionista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variables para manejo de Combos
Dim arrFondo()                      As String
Dim arrTipoTasa()                   As String, arrPeriodoTasa()             As String
Dim arrModalidadDevengo()           As String, arrPeriodoDevengo()          As String
Dim arrFrecuenciaPago()             As String, arrTipoDesplazamiento()      As String
Dim arrTipoComisionista()           As String, arrBaseCalculo()             As String
Dim arrPeriodoCapitalizacion()      As String, arrMoneda()                  As String

'Variables de los atributos de la PK
Dim strCodFondo                     As String, strTipoComisionista          As String
Dim strCodComisionista              As String

'Variables de los otros atributos
Dim strCodTipoTasa                  As String, strCodPeriodoTasa            As String
Dim strCodModalidadDevengo          As String, strCodPeriodoDevengo         As String
Dim strCodFrecuenciaPago            As String, strCodTipoDesplazamiento     As String
Dim strIndFinDeMes                  As String, strIndNoDeterminado          As String
Dim strCodBaseCalculo               As String, strCodMoneda                 As String
Dim strCodPeriodoCapitalizacion     As String

'Variables auxiliares
Dim strSQL                          As String
Dim intRegistro                     As String
Dim strEstado                       As String
Dim strCodVistaProceso              As String
Dim blnActualizaRecordset           As Boolean

Dim numSecuencial                   As Integer, strCodFormulaDevengo        As String

Dim arrCboControl()                 As String


'Variables de Conexion a Base de Datos
Dim adoRegistro                     As ADODB.Recordset
Dim adoFondoComisionistaCondicion   As ADODB.Recordset



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
    
    strSQL = "{ call up_ACSelDatosParametro(65,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
    
    With adoRegistro
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    tdgConsulta.DataSource = adoRegistro
    
    If adoRegistro.RecordCount > 0 Then strEstado = Reg_Consulta
        
End Sub

Private Sub CargarListas()

    
    '*** Fondos ***
    strSQL = "SELECT CodFondo CODIGO,DescripFondo DESCRIP FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND Estado='01' and CodFondo = '" & gstrCodFondoContable & "' ORDER BY DescripFondo"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter, True
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
        
    '*** Tipo de Comisionistas ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='CLSCOM' ORDER BY CodParametro"
    CargarControlLista strSQL, cboTipoComisionista, arrTipoComisionista(), Valor_Caracter, True
    
    '*** Monedas ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Valor_Caracter, True
    
    '*** Tipo de Tasas ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='NATTAS' ORDER BY CodParametro"
    CargarControlLista strSQL, cboTipoTasa, arrTipoTasa(), Valor_Caracter, True

    '*** Frecuencia de Pago ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPPRD' ORDER BY CodParametro"
    CargarControlLista strSQL, cboFrecuenciaPago, arrFrecuenciaPago(), Valor_Caracter, True
    
    '*** Tipo de Desplazamiento ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPDES' ORDER BY CodParametro"
    CargarControlLista strSQL, cboTipoDesplazamiento, arrTipoDesplazamiento(), Valor_Caracter, True
 
    '*** Base de Calculo ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='BASANU' ORDER BY CodParametro"
    CargarControlLista strSQL, cboBaseCalculo, arrBaseCalculo(), Valor_Caracter, True
    
    '*** Periodo de Capitalizacion ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' ORDER BY CodParametro"
    CargarControlLista strSQL, cboPeriodoCapitalizacion, arrPeriodoCapitalizacion(), Valor_Caracter, True
    
    
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
    
    Dim intAccion                           As Integer
    Dim lngNumError                         As Integer
    Dim strCodComisionista                  As String
    Dim strIndComisionistaActivo            As String

    Dim objFondoComisionistaCondicionXML    As DOMDocument60
    Dim strFondoComisionistaCondicionXML    As String
    Dim strMsgError                         As String
    Dim strListaCamposActualizables         As String
    
    If strEstado = Reg_Adicion Or strEstado = Reg_Edicion Then
        If TodoOK() Then
            Dim intCantRegistros    As Integer, intRegistro         As Integer
            Dim adoRegistro         As ADODB.Recordset
            Dim strNumAsiento       As String, strFechaGrabar       As String
            
            Me.MousePointer = vbHourglass
                                                
            With adoComm
                
                strCodComisionista = Trim(lblCodComisionista.Caption)
                
                strIndComisionistaActivo = IIf(chkComisionistaActivo.Value = vbChecked, Valor_Indicador, Valor_Caracter)
                
                'On Error GoTo Ctrl_Error
                strListaCamposActualizables = "NumSecuencial,FechaInicio,FechaFin,IndIndeterminado,PorcenTasa,CodTipoTasa,CodPeriodoTasa,CodPeriodoCapitalizacion,CodBaseCalculo,CodModalidadDevengo,CodPeriodoDevengo,CodFormulaDevengo,NumPeriodoPago,CodFrecuenciaPago,IndFinDeMes,CodTipoDesplazamiento"
                
                Call XMLADORecordset(objFondoComisionistaCondicionXML, "FondoComisionistaCondicion", "Condicion", adoFondoComisionistaCondicion, strMsgError, strListaCamposActualizables)
                strFondoComisionistaCondicionXML = objFondoComisionistaCondicionXML.xml
                
                '*** Cabecera ***
                .CommandText = "{ call up_ACManFondoComisionistaXML('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & strTipoComisionista & "','" & _
                    strCodComisionista & "','" & strCodMoneda & "','" & strCodVistaProceso & "','" & _
                    strIndComisionistaActivo & "','" & strFondoComisionistaCondicionXML & "','" & _
                    IIf(strEstado = Reg_Adicion, "I", "U") & "') }"
                adoConn.Execute .CommandText
                
                                                                                
            End With
            
            Set adoFondoComisionistaCondicion = Nothing
                
            Me.MousePointer = vbDefault
        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabCatalogo
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



Private Sub cboBaseCalculo_Click()

    strCodBaseCalculo = Valor_Caracter
    If cboBaseCalculo.ListIndex < 0 Then Exit Sub

    strCodBaseCalculo = arrBaseCalculo(cboBaseCalculo.ListIndex)

End Sub

Private Sub cboFondo_Click()

    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Call Buscar

End Sub


Private Function TodoOK() As Boolean

    TodoOK = False
    
    If Trim(lblCodComisionista.Caption) = Valor_Caracter Then
        MsgBox "Debe seleccionar un comisionista!", vbOKOnly + vbExclamation, Me.Caption
        Exit Function
    End If
    
    If cboTipoComisionista.ListIndex = -1 Then
        MsgBox "Debe seleccionar el Tipo Comisionista!", vbOKOnly + vbExclamation, Me.Caption
        Exit Function
    End If

    If Trim(cboMoneda.ListIndex) = -1 Then
        MsgBox "Debe seleccionar la moneda!", vbCritical, gstrNombreEmpresa
        If cboMoneda.Enabled Then cboMoneda.SetFocus
        Exit Function
    End If

    TodoOK = True


End Function
        
Private Sub InicializarVariables()

    strTipoComisionista = Valor_Caracter
    strCodVistaProceso = Valor_Caracter
    
    'Variables de los otros atributos
    strCodTipoTasa = Valor_Caracter
    strCodPeriodoTasa = Valor_Caracter
    strCodPeriodoCapitalizacion = Valor_Caracter
    strCodBaseCalculo = Valor_Caracter
    strCodModalidadDevengo = Valor_Caracter
    strCodPeriodoDevengo = Valor_Caracter
    strCodFrecuenciaPago = Valor_Caracter
    strCodTipoDesplazamiento = Valor_Caracter
    strIndFinDeMes = Valor_Caracter
    strIndNoDeterminado = Valor_Caracter
    
    
    'Variables auxiliares
    strSQL = Valor_Caracter
    intRegistro = 0
    numSecuencial = 0

End Sub

Private Sub LlenarFormulario(strModo As String)

    Select Case strModo
        Case Reg_Adicion
        
            Call InicializarVariables
            Call InicializarCampos
            
            cboTipoComisionista.Enabled = True
            cmdComisionista.Enabled = True
            cboMoneda.Enabled = True
            
            Call CargarDetalleGrilla
            
        Case Reg_Edicion ', Reg_Eliminacion
            
            Dim adoRecordset As New ADODB.Recordset
            
            strSQL = "SELECT TipoComisionista, CodComisionista, CodMoneda, IndComisionistaActivo " & _
                     "FROM FondoComisionista FC " & _
                     "WHERE " & _
                     "FC.CodFondo = '" & strCodFondo & "' AND " & _
                     "FC.CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
                     "FC.TipoComisionista = '" & tdgConsulta.Columns("TipoComisionista") & "' AND " & _
                     "FC.CodComisionista = '" & tdgConsulta.Columns("CODIGO") & "' AND " & _
                     "FC.CodMoneda = '" & tdgConsulta.Columns("CodMoneda") & "'"
                     
                     
            adoComm.CommandText = strSQL
                     
            Set adoRecordset = adoComm.Execute

            If Not adoRecordset.EOF Then
            
                Call InicializarVariables
                Call InicializarCampos
                
                lblCodComisionista.Caption = adoRecordset("CodComisionista").Value
                lblDescripComisionista.Caption = tdgConsulta.Columns("DESCRIPCION")
                
                intRegistro = ObtenerItemLista(arrTipoComisionista, adoRecordset("TipoComisionista").Value)
                If intRegistro >= 0 Then cboTipoComisionista.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrMoneda, adoRecordset("CodMoneda").Value)
                If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
                
                chkComisionistaActivo.Value = IIf(adoRecordset("IndComisionistaActivo").Value = Valor_Indicador, vbChecked, vbUnchecked)
                
                cboTipoComisionista.Enabled = False
                cmdComisionista.Enabled = False
                cboMoneda.Enabled = False
            
                Call CargarDetalleGrilla
            
            Else
                MsgBox "El Sistema no puede encontrar el comisionista seleccionado para consultar!", vbExclamation
                Exit Sub
            End If
            

            
    
    End Select

End Sub
Private Sub InicializarCampos()

    'Cabecera
    cboTipoComisionista.ListIndex = -1
    lblCodComisionista.Caption = Valor_Caracter
    lblDescripComisionista.Caption = Valor_Caracter

    'Vigencia
    dtpFechaDesde.Value = gdatFechaActual
    dtpFechaHasta.Value = gdatFechaActual
    
    'Tasa y Devengo
    cboMoneda.ListIndex = -1
    txtPorcenTasa.Text = "0.0000"
    cboTipoTasa.ListIndex = -1
    cboPeriodoTasa.ListIndex = -1
    cboBaseCalculo.ListIndex = ObtenerItemLista(arrBaseCalculo, Codigo_Base_Actual_360)
    cboModalidadDevengo.ListIndex = -1
    cboPeriodoDevengo.ListIndex = -1
    
    'Pagos
    txtPagoCada.Text = "1"
    cboFrecuenciaPago.ListIndex = ObtenerItemLista(arrFrecuenciaPago, Codigo_Tipo_Frecuencia_Mensual)
    chkFinDeMes.Value = vbChecked
    cboTipoDesplazamiento.ListIndex = ObtenerItemLista(arrTipoDesplazamiento, Tipo_Desplazamiento_Ningun_Desplazamiento)


End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabCatalogo.Tab = 0

    
    strIndFinDeMes = Valor_Caracter
    '*** Ancho por defecto de las columnas de la grilla ***
'    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 14
'    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 60
'    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 10
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAcciones.FormularioActivo = Me
    
    
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
'
'
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Vista Activa"
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
'
End Sub

Private Sub cboFrecuenciaPago_Click()
    
    strCodFrecuenciaPago = Valor_Caracter
    If cboFrecuenciaPago.ListIndex < 0 Then Exit Sub

    strCodFrecuenciaPago = arrFrecuenciaPago(cboFrecuenciaPago.ListIndex)

End Sub


Private Sub cboMoneda_Click()

    strCodMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))

End Sub


Private Sub cboPeriodoCapitalizacion_Click()

    strCodPeriodoCapitalizacion = Valor_Caracter
    If cboPeriodoCapitalizacion.ListIndex < 0 Then Exit Sub

    strCodPeriodoCapitalizacion = arrPeriodoCapitalizacion(cboPeriodoCapitalizacion.ListIndex)

End Sub

Private Sub cboPeriodoTasa_Click()

    strCodPeriodoTasa = Valor_Caracter
    If cboPeriodoTasa.ListIndex < 0 Then Exit Sub

    strCodPeriodoTasa = arrPeriodoTasa(cboPeriodoTasa.ListIndex)

End Sub



Private Sub cboPeriodoDevengo_Click()

    strCodPeriodoDevengo = Valor_Caracter
    If cboPeriodoDevengo.ListIndex < 0 Then Exit Sub

    strCodPeriodoDevengo = arrPeriodoDevengo(cboPeriodoDevengo.ListIndex)

End Sub

Private Sub cboTipoTasa_Click()

    strCodTipoTasa = Valor_Caracter
    If cboTipoTasa.ListIndex < 0 Then Exit Sub

    strCodTipoTasa = arrTipoTasa(cboTipoTasa.ListIndex)

    If strCodTipoTasa = Codigo_Tipo_Tasa_Flat Then
        '*** Periodo de Tasa ***
        strSQL = "SELECT '" & Valor_Caracter & "' AS CODIGO,'" & Sel_NoAplicable & "' AS DESCRIP"
        CargarControlLista strSQL, cboPeriodoTasa, arrPeriodoTasa(), Valor_Caracter, True
        
        cboPeriodoTasa.Visible = False
        lblDescrip(7).Visible = False
        
        '*** Periodo de Capitalizacion ***
        strSQL = "SELECT '" & Valor_Caracter & "' AS CODIGO,'" & Sel_NoAplicable & "' AS DESCRIP"
        CargarControlLista strSQL, cboPeriodoCapitalizacion, arrPeriodoCapitalizacion(), Valor_Caracter, True
        
        cboPeriodoCapitalizacion.Visible = False
        lblDescrip(12).Visible = False
        
        '*** Base de Calculo ***
        strSQL = "SELECT '" & Valor_Caracter & "' AS CODIGO,'" & Sel_NoAplicable & "' AS DESCRIP"
        CargarControlLista strSQL, cboBaseCalculo, arrBaseCalculo(), Valor_Caracter, True
        
        cboBaseCalculo.Visible = False
        lblDescrip(10).Visible = False
        
        '*** Tipo de Devengo ***: Soporta Alicuota (Lineal) y Valor Total
        strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='MODDEV' and Estado = '01' and CodParametro in ('01','03') ORDER BY CodParametro"
        CargarControlLista strSQL, cboModalidadDevengo, arrModalidadDevengo(), Valor_Caracter, True
        
        If strEstado = Reg_Adicion Then 'Se cargan valores por defecto solo cuando es nuevo
            cboPeriodoTasa.ListIndex = ObtenerItemLista(arrPeriodoTasa, Valor_Caracter)
            cboPeriodoCapitalizacion.ListIndex = ObtenerItemLista(arrPeriodoCapitalizacion, Valor_Caracter)
            cboBaseCalculo.ListIndex = ObtenerItemLista(arrBaseCalculo, Valor_Caracter)
            cboModalidadDevengo.ListIndex = ObtenerItemLista(arrModalidadDevengo, Codigo_Tipo_Devengo_Alicuota_Lineal)
        End If
        
    Else
        '*** Periodo de Tasa ***
        strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' ORDER BY CodParametro"
        CargarControlLista strSQL, cboPeriodoTasa, arrPeriodoTasa(), Valor_Caracter, True
           
        cboPeriodoTasa.Visible = True
        lblDescrip(7).Visible = True
           
        '*** Periodo de Capitalizacion ***
        strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' ORDER BY CodParametro"
        If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
            CargarControlLista strSQL, cboPeriodoCapitalizacion, arrPeriodoCapitalizacion(), Valor_Caracter, True
        ElseIf strCodTipoTasa = Codigo_Tipo_Tasa_Nominal Then
            CargarControlLista strSQL, cboPeriodoCapitalizacion, arrPeriodoCapitalizacion(), Sel_NoAplicable, True
        End If
        
        cboPeriodoCapitalizacion.Visible = True
        lblDescrip(12).Visible = True
    
        '*** Base de Calculo ***
        strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='BASANU' ORDER BY CodParametro"
        CargarControlLista strSQL, cboBaseCalculo, arrBaseCalculo(), Valor_Caracter, True
        
        cboBaseCalculo.Visible = True
        lblDescrip(10).Visible = True
    
        '*** Tipo de Devengo ***: Soporta Provision y Valor Total
        strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='MODDEV' and Estado = '01' and CodParametro in ('03','05') ORDER BY CodParametro"
        CargarControlLista strSQL, cboModalidadDevengo, arrModalidadDevengo(), Valor_Caracter, True
    
        If strEstado = Reg_Adicion Then 'Se cargan valores por defecto solo cuando es nuevo
            cboPeriodoTasa.ListIndex = -1
            cboPeriodoCapitalizacion.ListIndex = -1
            cboBaseCalculo.ListIndex = -1
            cboModalidadDevengo.ListIndex = ObtenerItemLista(arrModalidadDevengo, Codigo_Tipo_Devengo_Provision_Periodica)
        End If
        
    End If

End Sub

Private Sub cboTipoComisionista_Click()
    
    strTipoComisionista = Valor_Caracter
    If cboTipoComisionista.ListIndex < 0 Then Exit Sub

    strTipoComisionista = arrTipoComisionista(cboTipoComisionista.ListIndex)

    '*** Vistas Dinamica ***
'    strSQL = "SELECT CodVistaProceso CODIGO,DescripVistaProceso DESCRIP FROM VistaProceso " & _
'                "WHERE IndVigente='X'"
    
    If strTipoComisionista = Codigo_Tipo_Comisionista_Participe Then
        strCodVistaProceso = "019"
    ElseIf strTipoComisionista = Codigo_Tipo_Comisionista_Inversion Then
        strCodVistaProceso = "018"
    End If

End Sub




Private Sub cboTipoDesplazamiento_Click()

    strCodTipoDesplazamiento = Valor_Caracter
    If cboTipoDesplazamiento.ListIndex < 0 Then Exit Sub

    strCodTipoDesplazamiento = arrTipoDesplazamiento(cboTipoDesplazamiento.ListIndex)

End Sub

Private Sub cboModalidadDevengo_Click()

    strCodModalidadDevengo = Valor_Caracter
    If cboModalidadDevengo.ListIndex < 0 Then Exit Sub

    strCodModalidadDevengo = arrModalidadDevengo(cboModalidadDevengo.ListIndex)

    If strCodModalidadDevengo = Codigo_Tipo_Tasa_Flat Then
        '*** Periodo de Devengo ***
        strSQL = "SELECT '" & Valor_Caracter & "' AS CODIGO,'" & Sel_NoAplicable & "' AS DESCRIP"
        CargarControlLista strSQL, cboPeriodoDevengo, arrPeriodoDevengo(), Valor_Caracter, True
        
        cboPeriodoDevengo.Enabled = False
                    
        If strEstado = Reg_Adicion Then
            cboPeriodoDevengo.ListIndex = ObtenerItemLista(arrPeriodoDevengo, Valor_Caracter)
        End If
    Else
        '*** Periodo de Devengo ***
        strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' ORDER BY CodParametro"
        CargarControlLista strSQL, cboPeriodoDevengo, arrPeriodoDevengo(), Valor_Caracter, True
        
        cboPeriodoDevengo.Enabled = True
        
        If strEstado = Reg_Adicion Then
            cboPeriodoDevengo.ListIndex = ObtenerItemLista(arrPeriodoDevengo, Codigo_Tipo_Frecuencia_Diaria)
        End If
    End If

End Sub

Private Sub chkFinDeMes_Click()
    If chkFinDeMes.Value = 1 Then
        strIndFinDeMes = Valor_Indicador
    Else
        strIndFinDeMes = Valor_Caracter
    End If
End Sub

Private Sub chkIndeterminado_Click()
    
    If chkIndeterminado.Value = vbChecked Then
        dtpFechaHasta.Value = Convertddmmyyyy("99991231")
        dtpFechaHasta.Enabled = False
    Else
        dtpFechaHasta.Value = dtpFechaDesde.Value
        dtpFechaHasta.Enabled = True
    End If

End Sub











Private Sub cmdActualizar_Click()
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    If strEstado = Reg_Adicion Or strEstado = Reg_Edicion Then
        
        adoFondoComisionistaCondicion.Fields("FechaInicio") = dtpFechaDesde.Value
        adoFondoComisionistaCondicion.Fields("FechaFin") = dtpFechaHasta.Value
        adoFondoComisionistaCondicion.Fields("IndIndeterminado") = IIf(chkIndeterminado.Value = vbChecked, Valor_Indicador, Valor_Caracter)
        adoFondoComisionistaCondicion.Fields("PorcenTasa") = txtPorcenTasa.Value
        adoFondoComisionistaCondicion.Fields("CodTipoTasa") = strCodTipoTasa
        adoFondoComisionistaCondicion.Fields("DescripTipoTasa") = cboTipoTasa.List(cboTipoTasa.ListIndex) 'ObtenerItemLista(arrTipoTasa, strCodTipoTasa)
        adoFondoComisionistaCondicion.Fields("CodPeriodoTasa") = strCodPeriodoTasa
        adoFondoComisionistaCondicion.Fields("DescripPeriodoTasa") = cboPeriodoTasa.List(cboPeriodoTasa.ListIndex) 'ObtenerItemLista(arrPeriodoTasa, strCodPeriodoTasa)
        adoFondoComisionistaCondicion.Fields("CodPeriodoCapitalizacion") = strCodPeriodoCapitalizacion
        adoFondoComisionistaCondicion.Fields("DescripPeriodoCapitalizacion") = cboPeriodoCapitalizacion.List(cboPeriodoCapitalizacion.ListIndex) 'ObtenerItemLista(arrPeriodoCapitalizacion, strCodPeriodoCapitalizacion)
        adoFondoComisionistaCondicion.Fields("CodBaseCalculo") = strCodBaseCalculo
        adoFondoComisionistaCondicion.Fields("DescripBaseCalculo") = cboBaseCalculo.List(cboBaseCalculo.ListIndex) 'ObtenerItemLista(arrBaseCalculo, strCodBaseCalculo)
        adoFondoComisionistaCondicion.Fields("CodModalidadDevengo") = strCodModalidadDevengo
        adoFondoComisionistaCondicion.Fields("DescripModalidadDevengo") = cboModalidadDevengo.List(cboModalidadDevengo.ListIndex) 'ObtenerItemLista(arrModalidadDevengo, strCodModalidadDevengo)
        adoFondoComisionistaCondicion.Fields("CodPeriodoDevengo") = strCodPeriodoDevengo
        adoFondoComisionistaCondicion.Fields("DescripPeriodoDevengo") = Mid(cboPeriodoDevengo.List(cboPeriodoDevengo.ListIndex), 1, 30) 'ObtenerItemLista(arrPeriodoDevengo, strCodPeriodoDevengo)
        adoFondoComisionistaCondicion.Fields("CodFormulaDevengo") = strCodFormulaDevengo
        adoFondoComisionistaCondicion.Fields("DescripFormulaDevengo") = Mid(Trim(lblFormulaDevengo.Caption), 1, 100)
        adoFondoComisionistaCondicion.Fields("NumPeriodoPago") = txtPagoCada.Value
        adoFondoComisionistaCondicion.Fields("CodFrecuenciaPago") = strCodFrecuenciaPago
        adoFondoComisionistaCondicion.Fields("DescripFrecuenciaPago") = cboFrecuenciaPago.List(cboFrecuenciaPago.ListIndex) 'ObtenerItemLista(arrFrecuenciaPago, strCodFrecuenciaPago)
        adoFondoComisionistaCondicion.Fields("IndFinDeMes") = IIf(chkFinDeMes.Value = vbChecked, Valor_Indicador, Valor_Caracter)
        adoFondoComisionistaCondicion.Fields("CodTipoDesplazamiento") = strCodTipoDesplazamiento
        adoFondoComisionistaCondicion.Fields("DescripTipoDesplazamiento") = cboTipoDesplazamiento.List(cboTipoDesplazamiento.ListIndex) 'ObtenerItemLista(arrTipoDesplazamiento, strCodTipoDesplazamiento)
        
    End If

End Sub

Private Sub cmdAgregar_Click()
    Dim dblBookmark As Double
    
    If strEstado = Reg_Consulta Then Exit Sub
            
    If strEstado = Reg_Adicion Or strEstado = Reg_Edicion Then
        If TodoOkComisionistaCondicion() Then
           
            adoFondoComisionistaCondicion.AddNew
                        
            numSecuencial = numSecuencial + 1
            
            adoFondoComisionistaCondicion.Fields("NumSecuencial") = numSecuencial
            adoFondoComisionistaCondicion.Fields("FechaInicio") = dtpFechaDesde.Value
            adoFondoComisionistaCondicion.Fields("FechaFin") = dtpFechaHasta.Value
            adoFondoComisionistaCondicion.Fields("IndIndeterminado") = IIf(chkIndeterminado.Value = vbChecked, Valor_Indicador, Valor_Caracter)
            adoFondoComisionistaCondicion.Fields("PorcenTasa") = txtPorcenTasa.Value
            adoFondoComisionistaCondicion.Fields("CodTipoTasa") = strCodTipoTasa
            adoFondoComisionistaCondicion.Fields("DescripTipoTasa") = cboTipoTasa.List(cboTipoTasa.ListIndex) 'ObtenerItemLista(arrTipoTasa, strCodTipoTasa)
            adoFondoComisionistaCondicion.Fields("CodPeriodoTasa") = strCodPeriodoTasa
            adoFondoComisionistaCondicion.Fields("DescripPeriodoTasa") = cboPeriodoTasa.List(cboPeriodoTasa.ListIndex) 'ObtenerItemLista(arrPeriodoTasa, strCodPeriodoTasa)
            adoFondoComisionistaCondicion.Fields("CodPeriodoCapitalizacion") = strCodPeriodoCapitalizacion
            adoFondoComisionistaCondicion.Fields("DescripPeriodoCapitalizacion") = cboPeriodoCapitalizacion.List(cboPeriodoCapitalizacion.ListIndex) 'ObtenerItemLista(arrPeriodoCapitalizacion, strCodPeriodoCapitalizacion)
            adoFondoComisionistaCondicion.Fields("CodBaseCalculo") = strCodBaseCalculo
            adoFondoComisionistaCondicion.Fields("DescripBaseCalculo") = cboBaseCalculo.List(cboBaseCalculo.ListIndex) 'ObtenerItemLista(arrBaseCalculo, strCodBaseCalculo)
            adoFondoComisionistaCondicion.Fields("CodModalidadDevengo") = strCodModalidadDevengo
            adoFondoComisionistaCondicion.Fields("DescripModalidadDevengo") = cboModalidadDevengo.List(cboModalidadDevengo.ListIndex) 'ObtenerItemLista(arrModalidadDevengo, strCodModalidadDevengo)
            adoFondoComisionistaCondicion.Fields("CodPeriodoDevengo") = strCodPeriodoDevengo
            adoFondoComisionistaCondicion.Fields("DescripPeriodoDevengo") = Mid(cboPeriodoDevengo.List(cboPeriodoDevengo.ListIndex), 1, 30) 'ObtenerItemLista(arrPeriodoDevengo, strCodPeriodoDevengo)
            adoFondoComisionistaCondicion.Fields("CodFormulaDevengo") = strCodFormulaDevengo
            adoFondoComisionistaCondicion.Fields("DescripFormulaDevengo") = Mid(Trim(lblFormulaDevengo.Caption), 1, 100)
            adoFondoComisionistaCondicion.Fields("NumPeriodoPago") = txtPagoCada.Value
            adoFondoComisionistaCondicion.Fields("CodFrecuenciaPago") = strCodFrecuenciaPago
            adoFondoComisionistaCondicion.Fields("DescripFrecuenciaPago") = cboFrecuenciaPago.List(cboFrecuenciaPago.ListIndex) 'ObtenerItemLista(arrFrecuenciaPago, strCodFrecuenciaPago)
            adoFondoComisionistaCondicion.Fields("IndFinDeMes") = IIf(chkFinDeMes.Value = vbChecked, Valor_Indicador, Valor_Caracter)
            adoFondoComisionistaCondicion.Fields("CodTipoDesplazamiento") = strCodTipoDesplazamiento
            adoFondoComisionistaCondicion.Fields("DescripTipoDesplazamiento") = cboTipoDesplazamiento.List(cboTipoDesplazamiento.ListIndex) 'ObtenerItemLista(arrTipoDesplazamiento, strCodTipoDesplazamiento)
            adoFondoComisionistaCondicion.Fields("IndCondicionVigente") = Valor_Caracter
            adoFondoComisionistaCondicion.Fields("EstadoRegistro") = Valor_Caracter
                        
            adoFondoComisionistaCondicion.Update
                        
            tdgFondoComisionCondicion.Refresh
                        
            cmdQuitar.Enabled = True
        
        End If
    End If


End Sub
Private Function TodoOkComisionistaCondicion()

    TodoOkComisionistaCondicion = False
                  
    If Trim(cboTipoTasa.ListIndex) = -1 Then
        MsgBox "Tipo de tasa no ingresada", vbCritical, gstrNombreEmpresa
        If cboTipoTasa.Enabled Then cboTipoTasa.SetFocus
        Exit Function
    End If
    
    If Trim(cboPeriodoTasa.ListIndex) = -1 Then
        MsgBox "Periodo de tasa no ingresada", vbCritical, gstrNombreEmpresa
        If cboPeriodoTasa.Enabled Then cboPeriodoTasa.SetFocus
        Exit Function
    End If
    
    If Trim(cboPeriodoCapitalizacion.ListIndex) = -1 Then
        MsgBox "Periodo de capitalización de tasa no ingresado", vbCritical, gstrNombreEmpresa
        If cboPeriodoCapitalizacion.Enabled Then cboPeriodoCapitalizacion.SetFocus
        Exit Function
    End If
        
    If Trim(cboBaseCalculo.ListIndex) = -1 Then
        MsgBox "Base de calculo no ingresada", vbCritical, gstrNombreEmpresa
        If cboBaseCalculo.Enabled Then cboBaseCalculo.SetFocus
        Exit Function
    End If
        
    If Trim(cboModalidadDevengo.ListIndex) = -1 Then
        MsgBox "Modalidad de devengo no ingresada", vbCritical, gstrNombreEmpresa
        If cboModalidadDevengo.Enabled Then cboModalidadDevengo.SetFocus
        Exit Function
    End If
        
    If Trim(cboPeriodoDevengo.ListIndex) = -1 Then
        MsgBox "Periodo de devengo no ingresado", vbCritical, gstrNombreEmpresa
        If cboPeriodoDevengo.Enabled Then cboPeriodoDevengo.SetFocus
        Exit Function
    End If
        
    If Not IsNumeric(txtPagoCada.Text) Or CInt(txtPagoCada.Text) = 0 Then
        MsgBox "Numero de periodos para pago invalido!", vbCritical, gstrNombreEmpresa
        If txtPagoCada.Enabled Then txtPagoCada.SetFocus
        Exit Function
    End If
        
    If Trim(cboFrecuenciaPago.ListIndex) = -1 Then
        MsgBox "Frecuencia de pago no ingresada", vbCritical, gstrNombreEmpresa
        If cboFrecuenciaPago.Enabled Then cboFrecuenciaPago.SetFocus
        Exit Function
    End If
        
    If Trim(cboTipoDesplazamiento.ListIndex) = -1 Then
        MsgBox "Tipo de desplazmiento no ingresado", vbCritical, gstrNombreEmpresa
        If cboTipoDesplazamiento.Enabled Then cboTipoDesplazamiento.SetFocus
        Exit Function
    End If
        
    '*** Si todo pasó OK ***
    TodoOkComisionistaCondicion = True

End Function

Private Sub LimpiarDatos()

    lblFormulaDevengo.Caption = Valor_Caracter

End Sub

Private Sub cmdCondicionDevengo_Click(Index As Integer)

    Dim objParametroFormulaXML  As DOMDocument60
    Dim strParametroFormulaXML  As String
    Dim sSql As String
    Dim frmBus As frmBuscar
    Dim strMsgError As String
    
    If cboTipoComisionista.ListCount > 0 Then
        If cboTipoComisionista.ListIndex = -1 Then
            MsgBox "Debe seleccionar un tipo de comisionista", vbCritical
            cboTipoComisionista.SetFocus
            Exit Sub
        End If
    End If
    
    'Cargar las Variables del Formulario
    Call XMLFormularioControlTag(objParametroFormulaXML, "ParametroFormula", "Parametro", Me, strMsgError)
    
    strParametroFormulaXML = objParametroFormulaXML.xml
    
    Screen.MousePointer = vbHourglass
    
    Set frmBus = New frmBuscar
    
    With frmBus.TBuscarRegistro1
           
        .ADOConexion = adoConn
        .ADOConexion.CommandTimeout = 0
        .iTipoGrilla = 2
        
        frmBus.Caption = "Formulas Financieras"
        .sSql = "{ call up_ACSelFormulaExpresion('" & strParametroFormulaXML & "') }"
        .OutputColumns = "1,2"
        .HiddenColumns = ""
        
        .BuscarTabla
        
        Screen.MousePointer = vbNormal
        frmBus.Show 1
       
        If .iParams.Count = 0 Then Exit Sub
        
        'If .sCodigo <> "" Then
        If .iParams(1).Valor <> "" Then
            strCodFormulaDevengo = .iParams(1).Valor  '.sCodigo
            lblFormulaDevengo.Caption = .iParams(2).Valor '.sDescripcion
        End If
          
       
    End With
    
    Set frmBus = Nothing
    
    
'    gstrTextoAdministradorFormula = Valor_Caracter
'
'    Call CargarAdministradorFormulas
'
'    Select Case Index
'
'    Case 0
'        If Trim(lblFormulaDevengo.Caption) <> Valor_Caracter Then
'            gstrTextoAdministradorFormula = Trim(lblFormulaDevengo.Caption)
'            frmAdministradorFormulas.AdministradorFormulas1.TextoFormula = Trim(lblFormulaDevengo.Caption)
'        End If
'
'        frmAdministradorFormulas.Show vbModal
'
'        lblFormulaDevengo.Caption = gstrTextoAdministradorFormula
'
'    End Select

End Sub
Public Sub CargarAdministradorFormulas()
    Me.MousePointer = vbHourglass
    frmAdministradorFormulas.AdministradorFormulas1.CargarVariables gstrConnectNET, "up_ACLstVariablesVistaProceso", strCodVistaProceso, Tipo_Campo_Output, Valor_Caracter
    frmAdministradorFormulas.AdministradorFormulas1.CargarFunciones gstrConnectNET, "up_CNLstVistaUsuarioFuncion", strCodVistaProceso
    frmAdministradorFormulas.AdministradorFormulas1.CargarOperadoresConCadena "+|-"
    Me.MousePointer = vbDefault
End Sub
Private Sub cmdComisionista_Click()

  Dim sSql As String
   
    Screen.MousePointer = vbHourglass
   
    Dim frmBus As frmBuscar
    
    Set frmBus = New frmBuscar
    
    With frmBus.TBuscarRegistro1
           
        .ADOConexion = adoConn
        .ADOConexion.CommandTimeout = 0
        .iTipoGrilla = 2
        
        frmBus.Caption = " Relación de Comisionistas"
        .sSql = "{ call up_ACSelDatos(57) }"
        
        .OutputColumns = "1,2,3,4,5,6"
        .HiddenColumns = "1,2,6"
        
        .BuscarTabla
        
        Screen.MousePointer = vbNormal
        frmBus.Show 1
       
        If .iParams.Count = 0 Then Exit Sub
        
        If .iParams(1).Valor <> "" Then
            lblDescripComisionista.Caption = .iParams(5).Valor
            lblCodComisionista.Caption = .iParams(1).Valor
        End If
            
       
    End With
    
    Set frmBus = Nothing



End Sub


Private Sub cmdQuitar_Click()
    Dim dblBookmark As Double
    Dim numSecuencialActual As Integer

    If adoFondoComisionistaCondicion.RecordCount > 0 Then
    
        dblBookmark = adoFondoComisionistaCondicion.Bookmark
    
        numSecuencialActual = adoFondoComisionistaCondicion.Fields("NumSecuencial").Value
    
        If numSecuencial <= numSecuencialActual Then
            If adoFondoComisionistaCondicion.RecordCount > 1 Then
                numSecuencial = numSecuencial - 1
            Else
                numSecuencial = 0
            End If
        End If
        
        adoFondoComisionistaCondicion.Delete adAffectCurrent
        
        If adoFondoComisionistaCondicion.EOF Then
            adoFondoComisionistaCondicion.MovePrevious
            tdgFondoComisionCondicion.MovePrevious
        End If
            
        adoFondoComisionistaCondicion.Update
        
        If adoFondoComisionistaCondicion.RecordCount = 0 Then cmdQuitar.Enabled = False

        If adoFondoComisionistaCondicion.RecordCount > 0 And Not adoFondoComisionistaCondicion.BOF And Not adoFondoComisionistaCondicion.EOF And numSecuencial <= numSecuencialActual Then adoFondoComisionistaCondicion.Bookmark = dblBookmark - 1

        If adoFondoComisionistaCondicion.RecordCount > 0 And Not adoFondoComisionistaCondicion.BOF And Not adoFondoComisionistaCondicion.EOF And numSecuencial > numSecuencialActual Then adoFondoComisionistaCondicion.Bookmark = dblBookmark + 1
   
        tdgFondoComisionCondicion.Refresh
    
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
    
    Call ValidarPermisoUsoControl(Trim(gstrLogin), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
        
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
 

    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmFondoComisionista = Nothing
    
End Sub

Private Sub CargarDetalleGrilla()
    
    Dim adoRegistro As ADODB.Recordset
    Dim adoField As ADODB.Field
    
    Dim strSQL As String

    Call ConfiguraFondoComisionistaCondicion
    
    If strEstado = Reg_Edicion Then
    
        Set adoRegistro = New ADODB.Recordset
    
        strSQL = "{ call up_CNLstFondoComisionistaCondicion ('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                            strTipoComisionista & "','" & Trim(lblCodComisionista.Caption) & "','" & strCodMoneda & "')}"
        With adoRegistro
            .ActiveConnection = gstrConnectConsulta
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open strSQL
        
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    adoFondoComisionistaCondicion.AddNew
                    numSecuencial = numSecuencial + 1
                    For Each adoField In adoFondoComisionistaCondicion.Fields
                        adoFondoComisionistaCondicion.Fields(adoField.Name) = adoRegistro.Fields(adoField.Name)
                    Next
                    adoFondoComisionistaCondicion.Update
                    adoRegistro.MoveNext
                Loop
                adoFondoComisionistaCondicion.MoveFirst
            
                Call CargarDetalleGrillaRegistro
            
            End If
        End With
        
    End If
    
    tdgFondoComisionCondicion.DataSource = adoFondoComisionistaCondicion
            
    tdgFondoComisionCondicion.Refresh
    
End Sub

Private Sub CargarDetalleGrillaRegistro()

    Dim intRegistro As Integer

    If adoFondoComisionistaCondicion("EstadoRegistro") = "C" Then
        If adoFondoComisionistaCondicion("IndCondicionVigente") = Valor_Indicador Then
            cmdActualizar.Enabled = True
            cmdQuitar.Enabled = False
        Else
            cmdActualizar.Enabled = False
            cmdQuitar.Enabled = False
        End If
    Else
        cmdActualizar.Enabled = True
        cmdQuitar.Enabled = True
    End If

    'Vigencia
    dtpFechaDesde.Value = adoFondoComisionistaCondicion("FechaInicio").Value
    dtpFechaHasta.Value = adoFondoComisionistaCondicion("FechaFin").Value
    chkIndeterminado.Value = IIf(adoFondoComisionistaCondicion("IndIndeterminado").Value = Valor_Indicador, vbChecked, vbUnchecked)
    
    'Tasa y Devengo
    txtPorcenTasa.Text = adoFondoComisionistaCondicion("PorcenTasa").Value
    
    intRegistro = ObtenerItemLista(arrTipoTasa, adoFondoComisionistaCondicion("CodTipoTasa").Value)
    If intRegistro >= 0 Then cboTipoTasa.ListIndex = intRegistro
    
    intRegistro = ObtenerItemLista(arrPeriodoTasa, adoFondoComisionistaCondicion("CodPeriodoTasa").Value)
    If intRegistro >= 0 Then cboPeriodoTasa.ListIndex = intRegistro
    
    intRegistro = ObtenerItemLista(arrPeriodoCapitalizacion, adoFondoComisionistaCondicion("CodPeriodoCapitalizacion").Value)
    If intRegistro >= 0 Then cboPeriodoCapitalizacion.ListIndex = intRegistro
    
    intRegistro = ObtenerItemLista(arrBaseCalculo, adoFondoComisionistaCondicion("CodBaseCalculo").Value)
    If intRegistro >= 0 Then cboBaseCalculo.ListIndex = intRegistro
    
    intRegistro = ObtenerItemLista(arrModalidadDevengo, adoFondoComisionistaCondicion("CodModalidadDevengo").Value)
    If intRegistro >= 0 Then cboModalidadDevengo.ListIndex = intRegistro
    
    intRegistro = ObtenerItemLista(arrPeriodoDevengo, adoFondoComisionistaCondicion("CodPeriodoDevengo").Value)
    If intRegistro >= 0 Then cboPeriodoDevengo.ListIndex = intRegistro
    
    lblFormulaDevengo.Caption = adoFondoComisionistaCondicion("DescripFormulaDevengo").Value
    strCodFormulaDevengo = adoFondoComisionistaCondicion("CodFormulaDevengo").Value
    
    'Pagos
    txtPagoCada.Text = adoFondoComisionistaCondicion("NumPeriodoPago").Value
    
    intRegistro = ObtenerItemLista(arrFrecuenciaPago, adoFondoComisionistaCondicion("CodFrecuenciaPago").Value)
    If intRegistro >= 0 Then cboFrecuenciaPago.ListIndex = intRegistro

    chkFinDeMes.Value = IIf(adoFondoComisionistaCondicion("IndFinDeMes").Value = Valor_Indicador, vbChecked, vbUnchecked)
    
    intRegistro = ObtenerItemLista(arrTipoDesplazamiento, adoFondoComisionistaCondicion("CodTipoDesplazamiento").Value)
    If intRegistro >= 0 Then cboTipoDesplazamiento.ListIndex = intRegistro


End Sub


Private Sub tabCatalogo_Click(PreviousTab As Integer)

    Select Case tabCatalogo.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vQuery)
            If strEstado = Reg_Defecto Then tabCatalogo.Tab = 0
        

    End Select
    

End Sub


Private Sub tdgConsulta_DblClick()

    If strEstado <> Reg_Consulta Or tdgConsulta.Bookmark = 0 Then Exit Sub
    Call Accion(vQuery)

End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 2 Then
        Call DarFormatoValor(Value, Decimales_Tasa)
    End If

End Sub
'Public Sub CargarControlListaLocal(ByVal strSentencia As String, _
'                                   ByVal CtrlNombre As Control, _
'                                   ByRef arrControl() As String, _
'                                   ByVal strValor As String)
'
'    Dim adoBusqueda As ADODB.Recordset
'    Dim intCont     As Long
'
'    Set adoBusqueda = New ADODB.Recordset
'
'    adoComm.CommandText = strSentencia
'    Set adoBusqueda = adoComm.Execute
'
'    CtrlNombre.Clear
'    intCont = 0
'    ReDim arrControl(intCont)
'
'    Do Until adoBusqueda.EOF
'        CtrlNombre.AddItem adoBusqueda("DESCRIP")
'        ReDim Preserve arrControl(intCont)
'        arrControl(intCont) = adoBusqueda("CODIGO")
'        adoBusqueda.MoveNext
'        intCont = intCont + 1
'    Loop
'
'    adoBusqueda.Close: Set adoBusqueda = Nothing

'End Sub



Private Sub ConfiguraFondoComisionistaCondicion()

    Set adoFondoComisionistaCondicion = New ADODB.Recordset
    
    With adoFondoComisionistaCondicion
        .CursorLocation = adUseClient
        .Fields.Append "NumSecuencial", adInteger
        .Fields.Append "FechaInicio", adDate, 10
        .Fields.Append "FechaFin", adDate, 10
        .Fields.Append "IndIndeterminado", adChar, 1
        .Fields.Append "PorcenTasa", adDecimal
        .Fields.Item("PorcenTasa").Precision = 19
        .Fields.Item("PorcenTasa").NumericScale = 6
        .Fields.Append "CodTipoTasa", adVarChar, 2
        .Fields.Append "DescripTipoTasa", adVarChar, 30
        .Fields.Append "CodPeriodoTasa", adVarChar, 2
        .Fields.Append "DescripPeriodoTasa", adVarChar, 30
        .Fields.Append "CodPeriodoCapitalizacion", adVarChar, 2
        .Fields.Append "DescripPeriodoCapitalizacion", adVarChar, 30
        .Fields.Append "CodBaseCalculo", adVarChar, 2
        .Fields.Append "DescripBaseCalculo", adVarChar, 30
        .Fields.Append "CodModalidadDevengo", adVarChar, 2
        .Fields.Append "DescripModalidadDevengo", adVarChar, 30
        .Fields.Append "CodPeriodoDevengo", adVarChar, 2
        .Fields.Append "DescripPeriodoDevengo", adVarChar, 30
        .Fields.Append "CodFormulaDevengo", adVarChar, 10
        .Fields.Append "DescripFormulaDevengo", adVarChar, 100
        .Fields.Append "NumPeriodoPago", adInteger
        .Fields.Append "CodFrecuenciaPago", adVarChar, 2
        .Fields.Append "DescripFrecuenciaPago", adVarChar, 30
        .Fields.Append "IndFinDeMes", adChar, 1
        .Fields.Append "CodTipoDesplazamiento", adVarChar, 2
        .Fields.Append "DescripTipoDesplazamiento", adVarChar, 30
        .Fields.Append "IndCondicionVigente", adChar, 1
        .Fields.Append "EstadoRegistro", adChar, 1
        .LockType = adLockBatchOptimistic
    End With
    adoFondoComisionistaCondicion.Open

End Sub


Private Sub tdgFondoComisionCondicion_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If adoFondoComisionistaCondicion.EOF Or adoFondoComisionistaCondicion.RecordCount = 0 Then Exit Sub 'And adoRegistroAux.BOF
        
    Call CargarDetalleGrillaRegistro

End Sub
