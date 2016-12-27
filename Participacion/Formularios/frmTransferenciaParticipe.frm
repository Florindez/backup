VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmTransferenciaParticipe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transferencia de Certificados"
   ClientHeight    =   9660
   ClientLeft      =   1245
   ClientTop       =   1635
   ClientWidth     =   14775
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmTransferenciaParticipe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   14775
   ShowInTaskbar   =   0   'False
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   12840
      TabIndex        =   1
      Top             =   8880
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      Visible0        =   0   'False
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   8880
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   1296
      Buttons         =   4
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      Visible0        =   0   'False
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      Visible1        =   0   'False
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Eliminar"
      Tag2            =   "4"
      Visible2        =   0   'False
      ToolTipText2    =   "Eliminar"
      Caption3        =   "&Buscar"
      Tag3            =   "5"
      Visible3        =   0   'False
      ToolTipText3    =   "Buscar"
      UserControlWidth=   5700
   End
   Begin TabDlg.SSTab tabTransferencia 
      Height          =   8730
      Left            =   0
      TabIndex        =   11
      Top             =   30
      Width           =   14715
      _ExtentX        =   25956
      _ExtentY        =   15399
      _Version        =   393216
      Style           =   1
      Tabs            =   2
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
      TabPicture(0)   =   "frmTransferenciaParticipe.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraSolicitud"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Operación"
      TabPicture(1)   =   "frmTransferenciaParticipe.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraDatosGenerales"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraDatosTransferente"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraDatosSolicitud"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdAccion"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   11250
         TabIndex        =   10
         Top             =   7800
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
      Begin VB.Frame Frame1 
         Caption         =   "Criterios de búsqueda"
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
         Left            =   -74640
         TabIndex        =   51
         Top             =   540
         Width           =   14055
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   360
            Width           =   5835
         End
         Begin VB.ComboBox cboSucursal 
            Height          =   315
            Left            =   330
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   1350
            Width           =   3435
         End
         Begin VB.ComboBox cboAgencia 
            Height          =   315
            Left            =   3900
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   1350
            Width           =   3435
         End
         Begin VB.ComboBox cboPromotor 
            Height          =   315
            Left            =   7500
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   1350
            Width           =   3435
         End
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   315
            Left            =   8250
            TabIndex        =   56
            Top             =   330
            Width           =   1545
            _ExtentX        =   2725
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
            Format          =   178126849
            CurrentDate     =   38068
         End
         Begin MSComCtl2.DTPicker dtpFechaHasta 
            Height          =   315
            Left            =   8250
            TabIndex        =   57
            Top             =   705
            Width           =   1545
            _ExtentX        =   2725
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
            Format          =   178126849
            CurrentDate     =   38068
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Operador"
            BeginProperty Font 
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
            Index           =   2
            Left            =   7500
            TabIndex        =   63
            Top             =   1095
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Agencia"
            BeginProperty Font 
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
            Index           =   1
            Left            =   3960
            TabIndex        =   62
            Top             =   1095
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Sucursal"
            BeginProperty Font 
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
            Left            =   330
            TabIndex        =   61
            Top             =   1095
            Width           =   1215
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
            Index           =   22
            Left            =   330
            TabIndex        =   60
            Top             =   375
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
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
            Height          =   285
            Index           =   24
            Left            =   7500
            TabIndex        =   59
            Top             =   345
            Width           =   615
         End
         Begin VB.Label lblDescrip 
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
            Height          =   285
            Index           =   13
            Left            =   7500
            TabIndex        =   58
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.Frame fraDatosSolicitud 
         Caption         =   "Datos de quien(es) recibe(n) la Transferencia"
         Height          =   3180
         Left            =   360
         TabIndex        =   30
         Top             =   4470
         Width           =   13935
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
            Left            =   270
            Picture         =   "frmTransferenciaParticipe.frx":0044
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Agregar detalle"
            Top             =   1950
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
            Left            =   270
            Picture         =   "frmTransferenciaParticipe.frx":02F1
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Quitar detalle"
            Top             =   2520
            Width           =   495
         End
         Begin VB.CommandButton cmdBusqueda 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   6750
            TabIndex        =   6
            ToolTipText     =   "Búsqueda de Partícipe"
            Top             =   690
            Width           =   375
         End
         Begin VB.TextBox txtNumPapeleta 
            Enabled         =   0   'False
            Height          =   285
            Left            =   9480
            TabIndex        =   8
            Top             =   1005
            Width           =   2775
         End
         Begin VB.ComboBox cboFondoTransferido 
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   360
            Visible         =   0   'False
            Width           =   5085
         End
         Begin VB.ComboBox cboTipoDocumentoTransferido 
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   1005
            Width           =   3920
         End
         Begin VB.TextBox txtNumDocumentoTransferido 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            TabIndex        =   40
            Top             =   1365
            Width           =   3920
         End
         Begin VB.ComboBox cboFormaIngreso 
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "frmTransferenciaParticipe.frx":0543
            Left            =   9480
            List            =   "frmTransferenciaParticipe.frx":0545
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1365
            Width           =   2775
         End
         Begin VB.TextBox txtCuotasNuevoCertificado 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9480
            TabIndex        =   7
            Top             =   690
            Width           =   2775
         End
         Begin TrueOleDBGrid60.TDBGrid tdgTransferido 
            Bindings        =   "frmTransferenciaParticipe.frx":0547
            Height          =   1200
            Left            =   840
            OleObjectBlob   =   "frmTransferenciaParticipe.frx":0564
            TabIndex        =   36
            Top             =   1950
            Width           =   12945
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Num. Papeleta"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   7320
            TabIndex        =   48
            Top             =   1020
            Width           =   1050
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Partícipe"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   46
            Top             =   375
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Num.Documento"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   19
            Left            =   240
            TabIndex        =   45
            Top             =   1395
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Documento"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   20
            Left            =   240
            TabIndex        =   44
            Top             =   1050
            Width           =   1455
         End
         Begin VB.Label lblDescripParticipeTransferido 
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   240
            TabIndex        =   43
            Top             =   690
            Width           =   6495
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Forma de Ingreso"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   7320
            TabIndex        =   37
            Top             =   1395
            Width           =   1230
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nuevo(s)  Certificado(s)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   840
            TabIndex        =   34
            Top             =   1710
            Width           =   1650
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Valor Cuota"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   12
            Left            =   7335
            TabIndex        =   33
            Top             =   375
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cuotas Nuevo Certificado"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   7320
            TabIndex        =   32
            Top             =   705
            Width           =   1815
         End
         Begin VB.Label lblValorCuota 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   9480
            TabIndex        =   31
            Top             =   360
            Visible         =   0   'False
            Width           =   2775
         End
      End
      Begin VB.Frame fraDatosTransferente 
         Caption         =   "Datos de quien Transfiere"
         Height          =   2505
         Left            =   360
         TabIndex        =   22
         Top             =   1980
         Width           =   13935
         Begin VB.CommandButton cmdBusqueda 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   6690
            TabIndex        =   5
            ToolTipText     =   "Búsqueda de Contrato"
            Top             =   690
            Width           =   375
         End
         Begin VB.ComboBox cboTipoDocumentoTransferente 
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1050
            Width           =   3920
         End
         Begin VB.TextBox txtNumDocumentoTransferente 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            TabIndex        =   23
            Top             =   1455
            Width           =   3920
         End
         Begin TrueOleDBGrid60.TDBGrid tdgTransferente 
            Bindings        =   "frmTransferenciaParticipe.frx":903F
            Height          =   2130
            Left            =   7380
            OleObjectBlob   =   "frmTransferenciaParticipe.frx":905D
            TabIndex        =   38
            Top             =   330
            Width           =   6405
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Partícipe"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   25
            Left            =   240
            TabIndex        =   50
            Top             =   300
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Certificados Vigentes"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   7380
            TabIndex        =   39
            Top             =   150
            Width           =   1485
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Num.Documento"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   5
            Left            =   240
            TabIndex        =   29
            Top             =   1485
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Documento"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   18
            Left            =   240
            TabIndex        =   28
            Top             =   1065
            Width           =   1455
         End
         Begin VB.Label lblDescripParticipeTransferente 
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   240
            TabIndex        =   27
            Top             =   690
            Width           =   6435
         End
         Begin VB.Label lblDescripTipoParticipeTransferente 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2040
            TabIndex        =   26
            Top             =   1845
            Width           =   3920
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Partícipe"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   4
            Left            =   240
            TabIndex        =   25
            Top             =   1860
            Width           =   1575
         End
      End
      Begin VB.Frame fraDatosGenerales 
         Caption         =   "Datos Generales"
         Height          =   1455
         Left            =   360
         TabIndex        =   16
         Top             =   480
         Width           =   13935
         Begin VB.TextBox txtEspecificarOtro 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1830
            TabIndex        =   65
            Top             =   660
            Width           =   3915
         End
         Begin VB.ComboBox cboFondoTransferente 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1020
            Width           =   3945
         End
         Begin VB.ComboBox cboEjecutivo 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   7350
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   630
            Width           =   3750
         End
         Begin VB.ComboBox cboTipoOperacion 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   300
            Width           =   3915
         End
         Begin VB.Timer tmrHora 
            Enabled         =   0   'False
            Interval        =   60000
            Left            =   9000
            Top             =   200
         End
         Begin MSComCtl2.DTPicker dtpHoraSolicitud 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   4
            EndProperty
            Height          =   285
            Left            =   10200
            TabIndex        =   35
            Top             =   300
            Width           =   855
            _ExtentX        =   1508
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
            CustomFormat    =   "HH:mm"
            Format          =   178126851
            UpDown          =   -1  'True
            CurrentDate     =   38831
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Especificar"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   26
            Left            =   240
            TabIndex        =   64
            Top             =   660
            Width           =   1395
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fondo"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   14
            Left            =   240
            TabIndex        =   49
            Top             =   1035
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Hora"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   17
            Left            =   9480
            TabIndex        =   21
            Top             =   315
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   16
            Left            =   6120
            TabIndex        =   20
            Top             =   315
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Operador"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   6120
            TabIndex        =   19
            Top             =   675
            Width           =   660
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Transferencia"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   23
            Left            =   240
            TabIndex        =   18
            Top             =   315
            Width           =   1575
         End
         Begin VB.Label lblFechaSolicitud 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "dd/mm/yyyy"
            Height          =   285
            Left            =   7320
            TabIndex        =   17
            Top             =   300
            Width           =   1575
         End
      End
      Begin VB.Frame fraSolicitud 
         Height          =   5685
         Left            =   -74640
         TabIndex        =   12
         Top             =   2370
         Width           =   14025
         Begin VB.ListBox lstLeyenda 
            Height          =   255
            Left            =   9120
            TabIndex        =   13
            Top             =   180
            Visible         =   0   'False
            Width           =   1200
         End
         Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
            Bindings        =   "frmTransferenciaParticipe.frx":E949
            Height          =   4725
            Left            =   330
            OleObjectBlob   =   "frmTransferenciaParticipe.frx":E963
            TabIndex        =   47
            Top             =   600
            Width           =   13005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Parcial (0)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   2040
            TabIndex        =   15
            Top             =   240
            Width           =   705
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Total (0)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   21
            Left            =   360
            TabIndex        =   14
            Top             =   240
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frmTransferenciaParticipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()                                  As String, arrFondoTransferente()           As String
Dim arrFondoTransferido()                       As String, arrTipoOperacion()               As String
Dim arrSucursal()                               As String, arrEjecutivo()                   As String
Dim arrAgencia()                                As String, arrPromotor()                    As String
Dim arrLeyendaTransferencia()                   As String, arrFormaIngreso()                As String

Dim strCodFondo                                 As String, strCodFondoTransferente          As String
Dim strCodFondoTransferido                      As String, strCodTipoOperacion              As String
Dim strCodSucursal                              As String, strCodEjecutivo                  As String
Dim strCodAgencia                               As String, strCodPromotor                   As String
Dim strCodSucursalTransferencia                 As String, strCodAgenciaTransferencia       As String
Dim strCodClaseOperacion                        As String, strCodTipoTransferencia          As String
Dim strCodMonedaFondoTransferente               As String, strCodMonedaFondoTransferido     As String
Dim strEstado                                   As String, strCodFormaIngreso               As String
Dim strHoraCorte                                As String, strCodTipoValuacion              As String
Dim strCodComision                              As String
Dim dblValorCuotaTransferente                   As Double, dblMontoMinSuscripcionInicial    As Double
Dim dblCantCuotaMinSuscripcionInicial           As Double, dblCantMinCuotaSuscripcion       As Double
Dim dblMontoMinSuscripcion                      As Double
Dim blnSeleccion                                As Boolean, blnValorConocido                As Boolean
Dim adoConsulta                                 As ADODB.Recordset
Dim indSortAsc                                  As Boolean, indSortDesc                     As Boolean
Dim strCodParticipeBusqueda                     As String, numSecuencial                    As Long

Dim strCodParticipeTransferente                 As String
Dim strTipoDocumentoTransferente                As String
Dim strNumDocumentoTransferente                 As String
Dim strCodTipoDocumentoTransferente             As String
Dim strDescripTitularSolicitanteTransferente    As String
Dim strCodClienteTransferente                   As String
Dim adoTransferente                             As ADODB.Recordset

Dim strCodParticipeTransferido                  As String
Dim strTipoDocumentoTransferido                 As String
Dim strNumDocumentoTransferido                  As String
Dim strCodTipoDocumentoTransferido              As String
Dim strDescripTitularSolicitanteTransferido     As String
Dim strCodClienteTransferido                    As String
Dim adoTransferido                              As ADODB.Recordset

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vModify
            Call Modificar
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

Public Sub Eliminar()

End Sub

Public Sub Adicionar()
        
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Operación de Transferencia..."
    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabTransferencia
        .TabEnabled(0) = False
        .Tab = 1
    End With
    Call Deshabilita
    
End Sub

Private Sub Deshabilita()

    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim strSql As String
    Dim intRegistro As Integer
    Dim adoRegistro As ADODB.Recordset
    strCodTipoOperacion = "03"
    
    Select Case strModo
        Case Reg_Adicion
            
            strCodParticipeTransferido = Valor_Caracter
            
            cboTipoOperacion.ListIndex = -1
            If cboTipoOperacion.ListCount > 0 Then cboTipoOperacion.ListIndex = 0
                        
            txtNumPapeleta.Text = Valor_Caracter
            lblFechaSolicitud.Caption = CStr(gdatFechaActual)
            
            dtpHoraSolicitud.Value = ObtenerHoraServidor
            
            cboEjecutivo.ListIndex = -1
            intRegistro = ObtenerItemLista(arrEjecutivo(), gstrCodPromotor)
            If intRegistro >= 0 Then cboEjecutivo.ListIndex = intRegistro
            
            cboFondoTransferente.ListIndex = -1
            If cboFondoTransferente.ListCount > 0 Then cboFondoTransferente.ListIndex = 0
            
            cboFondoTransferido.ListIndex = -1
            If cboFondoTransferido.ListCount > 0 Then cboFondoTransferido.ListIndex = 0
            
            cboTipoDocumentoTransferente.ListIndex = -1
            If cboTipoDocumentoTransferente.ListCount > 0 Then cboTipoDocumentoTransferente.ListIndex = 0
            
            cboTipoDocumentoTransferido.ListIndex = -1
            If cboTipoDocumentoTransferido.ListCount > 0 Then cboTipoDocumentoTransferido.ListIndex = 0
                        
            txtNumDocumentoTransferente.Text = ""
            txtNumDocumentoTransferido.Text = ""
            lblDescripTipoParticipeTransferente.Caption = ""
            lblDescripParticipeTransferido.Caption = ""
                                                                        
            'Call ObtenerCertificados
            
            '*** Limpiamos tabla ***
            adoComm.CommandText = "DELETE CertificadoTransferidoTmp"
            adoConn.Execute adoComm.CommandText
            
            '*** Limpiamos tabla transferente***
            adoComm.CommandText = "DELETE CertificadoTransferenteTmp"
            adoConn.Execute adoComm.CommandText
            
            'Call ObtenerCertificadoTransferido
            
            txtCuotasNuevoCertificado.Text = "0"
            lblValorCuota.Caption = "0"
            
            cboFormaIngreso.ListIndex = -1
            intRegistro = ObtenerItemLista(arrFormaIngreso(), "02")
            If intRegistro >= 0 Then cboFormaIngreso.ListIndex = intRegistro
            cboFormaIngreso.Enabled = False
            blnSeleccion = False
                                        
            strSql = "{ call up_ACSelDatos(41) }"
            CargarControlLista strSql, cboEjecutivo, arrEjecutivo(), Sel_Defecto
    
            If cboEjecutivo.ListCount > 0 Then cboEjecutivo.ListIndex = 0
            intRegistro = ObtenerItemLista(arrEjecutivo(), strCodEjecutivo)
            If intRegistro >= 0 Then cboEjecutivo.ListIndex = intRegistro
                                
            adoComm.CommandText = "{ call up_ACSelDatosParametro(64,'" & strCodFondo & "') }"
            Set adoRegistro = adoComm.Execute

            If Not adoRegistro.EOF Then
                If IsNull(Mid(adoRegistro("NumPapeleta"), 1, 15)) Then
                    txtNumPapeleta.Text = Format(1, "000000000000000")
                Else
                    txtNumPapeleta.Text = adoRegistro("NumPapeleta")
                End If
            Else
                txtNumPapeleta.Text = Format(1, "000000000000000")
            End If
            
            adoRegistro.Close: Set adoRegistro = Nothing
                                        
            numSecuencial = 0

                                        
            cboEjecutivo.SetFocus
                        
        Case Reg_Edicion
    
    End Select
    
    tmrHora.Enabled = True
    
End Sub
Public Sub Ayuda()

End Sub

Public Sub Buscar()

    Dim strFechaDesde       As String, strFechaHasta        As String
    Dim datFechaSiguiente   As Date
    Dim strSql              As String
    
    Set adoConsulta = New ADODB.Recordset
                                                                                    
    Me.MousePointer = vbHourglass
    
    strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
    datFechaSiguiente = DateAdd("d", 1, dtpFechaHasta.Value)
    strFechaHasta = Convertyyyymmdd(datFechaSiguiente)
                
    If cboSucursal.ListIndex > 0 And cboAgencia.ListIndex > 0 And cboPromotor.ListIndex > 0 Then
    
        strSql = "{ call up_ACSelDatosParametro(59,'" & strCodFondo & "','" & _
            gstrCodAdministradora & "','" & strFechaDesde & "','" & strFechaHasta & "','" & _
            Codigo_Operacion_Transferencia & "','" & strCodSucursal & "','" & strCodAgencia & "','" & _
            strCodPromotor & "') }"
            
    ElseIf cboSucursal.ListIndex > 0 And cboAgencia.ListIndex > 0 Then
    
        strSql = "{ call up_ACSelDatosParametro(58,'" & strCodFondo & "','" & _
            gstrCodAdministradora & "','" & strFechaDesde & "','" & strFechaHasta & "','" & _
            Codigo_Operacion_Transferencia & "','" & strCodSucursal & "','" & strCodAgencia & "') }"
    
    ElseIf cboSucursal.ListIndex > 0 Then
    
        strSql = "{ call up_ACSelDatosParametro(57,'" & strCodFondo & "','" & _
            gstrCodAdministradora & "','" & strFechaDesde & "','" & strFechaHasta & "','" & _
            Codigo_Operacion_Transferencia & "','" & strCodSucursal & "') }"
    
    Else
        strSql = "{ call up_ACSelDatosParametro(56,'" & strCodFondo & "','" & _
            gstrCodAdministradora & "','" & strFechaDesde & "','" & strFechaHasta & "','" & _
            Codigo_Operacion_Transferencia & "') }"
        'MsgBox strSQL, vbCritical
            
    End If
    
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSql
    End With
        
    tdgConsulta.DataSource = adoConsulta
    
    If adoConsulta.RecordCount > 0 Then
        strEstado = Reg_Consulta
        
        Dim intNumTotales As Integer, intNumParciales  As Integer

        intNumTotales = 0: intNumParciales = 0
        With adoConsulta
            .MoveFirst
            
            Do While Not .EOF
                If Right(.Fields("TipoOperacion"), 1) = "T" Then intNumTotales = intNumTotales + 1
                If Right(.Fields("TipoOperacion"), 1) = "P" Then intNumParciales = intNumParciales + 1

                .MoveNext
            Loop
        End With
        lblDescrip(21).Caption = "Total (" & CStr(intNumTotales) & ")"
        lblDescrip(11).Caption = "Parcial (" & CStr(intNumParciales) & ")"
    Else
        lblDescrip(21).Caption = "Total (0)"
        lblDescrip(11).Caption = "Parcial (0)"
    End If
            
    Me.MousePointer = vbDefault
    
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabTransferencia
        .TabEnabled(0) = True
        .Tab = 0
    End With
    gstrCodParticipeTransferente = Valor_Caracter
    gstrCodParticipeTransferido = Valor_Caracter
    tmrHora.Enabled = False
    
End Sub

Public Sub Grabar()

    Dim adoRegistro As ADODB.Recordset
    Dim adoRegistro2 As ADODB.Recordset
    Dim adoRegistro3 As ADODB.Recordset
    
    Dim intRegistro As Integer, intContador         As Integer
    Dim dblCuotas   As Double, dblCuotasAcumulado   As Double, valorCuotaCopia As Double, codParticipeCopia As String
    Dim fechaSuscripcionCopia As Date
    Dim strNumSolicitud As String
    Dim dblValorCuota   As Double
    
    
    Dim objTransferentesXML                 As DOMDocument60
    Dim objTransferidosXML                  As DOMDocument60
    Dim strTransferentesXML                 As String
    Dim strTransferidosXML                  As String
    Dim strMsgError                         As String
    Dim strListaCamposActualizables         As String
    
    
    Set adoRegistro2 = New ADODB.Recordset
       
    If strEstado = Reg_Consulta Then Exit Sub
    
    dblCuotas = 0
    'dblCuotasAcumulado = 0
        
    If strEstado = Reg_Adicion Then
        If TodoOK() Then

            strListaCamposActualizables = "CodParticipeTransferente,NumCertificadoTransferente,CantCuotas,CantCuotasPorTransferir,ValorCuota"
            
            Call XMLADORecordset(objTransferentesXML, "ParticipeSolicitud", "Transferente", adoTransferente, strMsgError, strListaCamposActualizables)
            strTransferentesXML = objTransferentesXML.xml

            strListaCamposActualizables = "CodParticipeTransferente,NumCertificadoTransferente,CantCuotas,ValorCuota,CodParticipeTransferido"
            
            Call XMLADORecordset(objTransferidosXML, "ParticipeSolicitud", "Transferido", adoTransferido, strMsgError, strListaCamposActualizables)
            strTransferidosXML = objTransferidosXML.xml

'            If blnSeleccion Then
'                strCodClaseOperacion = Codigo_Clase_TransferenciaParcial
'            Else
'                strCodClaseOperacion = Codigo_Clase_TransferenciaTotal
'            End If
            
            '*** Guardar Solicitud ***
            adoComm.CommandText = "{ call up_PRProcTransferenciaParticipe('" & _
                strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
                strNumSolicitud & "','" & _
                Convertyyyymmdd(CVDate(lblFechaSolicitud.Caption)) & Space(1) & Format(dtpHoraSolicitud.Value, "hh:mm") & "','" & _
                strCodTipoOperacion & "','" & _
                strCodParticipeTransferente & "','" & _
                Trim(txtNumPapeleta.Text) & "','" & _
                strCodSucursal & "','" & _
                strCodAgencia & "','" & _
                strCodEjecutivo & "','" & _
                strCodMonedaFondoTransferente & "'," & _
                dblValorCuota & "," & _
                dblCuotasAcumulado & ",'" & _
                strCodTipoTransferencia & "','" & _
                Trim(txtEspecificarOtro.Text) & "','" & _
                strTransferentesXML & "','" & _
                strTransferidosXML & "','" & _
                "I') }"

            adoConn.Execute adoComm.CommandText


'            Dim strSql              As String
'            Dim strNumCertificado   As String, strNumOperacion  As String
'            Dim strClaseCliente     As String, strNumSolicitud  As String
'            Dim strFechaInicio      As String, strFechaFin      As String
'            Dim contAgrupado        As Integer
'
'            adoComm.CommandText = "{ call up_ACSelDatosParametro(64,'" & strCodFondo & "') }"
'            Set adoRegistro = adoComm.Execute
'
'            If Not adoRegistro.EOF Then
'                If IsNull(Mid(adoRegistro("NumPapeleta"), 1, 15)) Then
'                    txtNumPapeleta.Text = Format(1, "000000000000000")
'                Else
'                    txtNumPapeleta.Text = adoRegistro("NumPapeleta")
'                End If
'            Else
'                txtNumPapeleta.Text = Format(1, "000000000000000")
'            End If
'
'            Me.MousePointer = vbHourglass
'
'            If blnSeleccion Then
'                strCodClaseOperacion = Codigo_Clase_TransferenciaParcial
'            Else
'                strCodClaseOperacion = Codigo_Clase_TransferenciaTotal
'            End If
'
'
'            If MsgBox("¿Desea agrupar todos los certificados en uno solo?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
'
'                adoTransferido.Recordset.MoveFirst
'                blnSeleccion = False
'
'                Do While Not adoTransferido.Recordset.EOF
'                    If adoTransferido.Recordset.Fields(5) = "02" Then   'que sume todos menos los generados por saldo restante
'                        dblCuotasAcumulado = dblCuotasAcumulado + CDbl(adoTransferido.Recordset.Fields("CantCuotas"))
'                    End If
'                    adoTransferido.Recordset.MoveNext
'                Loop
'
'                Set adoRegistro = New ADODB.Recordset
'                With adoComm
'
'                    intContador = 1
'                    contAgrupado = 0
'
'                    adoTransferido.Recordset.MoveFirst
'
'                    Do While Not adoTransferido.Recordset.EOF
'                        If adoTransferido.Recordset.Fields(5) = "02" Then
'
'                            If contAgrupado < 1 Then
'                                contAgrupado = contAgrupado + 1
'
'                                .CommandText = "SELECT ValorParametro FROM AuxiliarParametro " & _
'                                "WHERE CodParametro='" & strCodTipoDocumentoTransferente & "' AND CodTipoParametro='TIPIDE'"
'                                Set adoRegistro = .Execute
'
'                                If Not adoRegistro.EOF Then
'                                    strClaseCliente = Trim(adoRegistro("ValorParametro"))
'                                End If
'                                adoRegistro.Close
'
'
'                                adoComm.CommandText = "SELECT CodFondo, CodParticipe, FechaOperacion, NumOperacion, FechaSuscripcion, CantCuotas, CantCuotasPagadas, ValorCuota, TipoOperacion, ClaseOperacion, NumCertificado FROM ParticipeCertificado " & _
'                                          "WHERE NumOperacion='" & adoTransferido.Recordset.Fields("NumOpCopia") & "' AND CodFondo='" & strCodFondoTransferido & "'"
'                                Set adoRegistro2 = adoComm.Execute
'
'                                If Not adoRegistro2.EOF Then
'                                    If CDec(adoTransferido.Recordset.Fields("CantCuotas")) < adoRegistro2("CantCuotas") Then
'                                        strCodClaseOperacion = Codigo_Clase_TransferenciaParcial
'                                    Else
'                                        strCodClaseOperacion = Codigo_Clase_TransferenciaTotal
'                                    End If
'                                End If
'
'
'                                If gstrCodParticipeTransferente <> Trim(adoTransferido.Recordset.Fields("CodParticipe")) Then
'                                    .CommandType = adCmdStoredProc
'
'                                    '*** Obtener el número del parámetro **
'                                    .CommandText = "up_ACObtenerUltNumero"
'                                    .Parameters.Append .CreateParameter("CodFondo", adChar, adParamInput, 3, strCodFondo)
'                                    .Parameters.Append .CreateParameter("CodAdministradora", adChar, adParamInput, 3, gstrCodAdministradora)
'                                    .Parameters.Append .CreateParameter("CodParametro", adChar, adParamInput, 6, Valor_NumSolicitud)
'                                    .Parameters.Append .CreateParameter("NuevoNumero", adChar, adParamOutput, 10, "")
'                                    .Execute
'
'                                    If Not .Parameters("NuevoNumero") Then
'                                        strNumSolicitud = .Parameters("NuevoNumero").Value
'                                        .Parameters.Delete ("CodFondo")
'                                        .Parameters.Delete ("CodAdministradora")
'                                        .Parameters.Delete ("CodParametro")
'                                        .Parameters.Delete ("NuevoNumero")
'                                    End If
'
'                                    .CommandType = adCmdText
'
'                                    '*** Guardar Solicitud ***
'                                    .CommandText = "{ call up_PRManTransferenciaParticipe('" & _
'                                        strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                        strNumSolicitud & "','" & gstrCodParticipeTransferente & "','" & _
'                                        Trim(adoTransferido.Recordset.Fields("NumFolio")) & "','','" & _
'                                        strCodSucursal & "','" & strCodSucursal & "','" & _
'                                        strCodAgencia & "','" & strCodAgencia & "','" & _
'                                        strCodEjecutivo & "','" & strCodEjecutivo & "','" & _
'                                        strCodTipoOperacion & "','" & strCodClaseOperacion & "','" & _
'                                        Convertyyyymmdd(CVDate(lblFechaSolicitud.Caption)) & Space(1) & Format(dtpHoraSolicitud.Value, "hh:mm") & "','" & _
'                                        Convertyyyymmdd(CVDate(lblFechaSolicitud.Caption)) & "','" & _
'                                        strCodMonedaFondoTransferente & "',0,"
'
'                                    .CommandText = .CommandText & CDec(adoTransferido.Recordset.Fields("ValorCuota")) & "," & dblCuotasAcumulado & "," & _
'                                        "0,0,0,0,'" & _
'                                        "','','','','" & _
'                                        "','','','','" & _
'                                        "','','" & _
'                                        "X','','" & _
'                                        Convertyyyymmdd(CVDate(Valor_Fecha)) & "','X','','" & _
'                                        "',0,'','" & _
'                                        Estado_Solicitud_Procesada & "','" & _
'                                        gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
'                                        gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
'                                        strCodTipoOperacion & "','" & Trim(lblDescripParticipeTransferente.Caption) & "','" & strCodTipoTransferencia & "','" & Trim(txtEspecificarOtro.Text) & "','I') }"
'
'                                    adoConn.Execute .CommandText
'
'                                    '*** Guardar Detalle Solicitud ***
'                                    .CommandText = "{ call up_PRManTransferenciaParticipeDetalle('" & _
'                                        strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                        strNumSolicitud & "','1','" & gstrCodParticipeTransferente & "','" & _
'                                        "'," & dblCuotasAcumulado & ",'" & _
'                                        gstrCodParticipeTransferido & "','" & Trim(lblDescripParticipeTransferido.Caption) & "','" & _
'                                        "X','I') }"
'                                        'MsgBox .CommandText, vbCritical
'                                    adoConn.Execute .CommandText
'
'                                    '*** Actualizar Secuenciales ***
'                                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                        Valor_NumSolicitud & "','" & strNumSolicitud & "') }"
'                                    adoConn.Execute .CommandText
'                                End If
'
'                                .CommandType = adCmdStoredProc
'                                '*** Obtener el número del parámetro **
'                                .CommandText = "up_ACObtenerUltNumero"
'                                .Parameters.Append .CreateParameter("CodFondo", adChar, adParamInput, 3, strCodFondoTransferente)
'                                .Parameters.Append .CreateParameter("CodAdministradora", adChar, adParamInput, 3, gstrCodAdministradora)
'                                .Parameters.Append .CreateParameter("CodParametro", adChar, adParamInput, 6, Valor_NumOpeCertificado)
'                                .Parameters.Append .CreateParameter("NuevoNumero", adChar, adParamOutput, 10, Valor_Caracter)
'                                .Execute
'
'                                If Not .Parameters("NuevoNumero") Then
'                                    strNumOperacion = .Parameters("NuevoNumero").Value
'                                    .Parameters.Delete ("CodFondo")
'                                    .Parameters.Delete ("CodAdministradora")
'                                    .Parameters.Delete ("CodParametro")
'                                    .Parameters.Delete ("NuevoNumero")
'                                End If
'
'                                .CommandType = adCmdText
'
'                                '*** Guardar Operación ***
'                                .CommandText = "{ call up_GNAdicOperacionParticipe('" & _
'                                    strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                    gstrCodParticipeTransferente & "','" & strNumOperacion & "','" & _
'                                    Convertyyyymmdd(gdatFechaActual) & "','" & Trim(arrTipoOperacion(cboTipoOperacion.ListIndex)) & "','" & _
'                                    strCodClaseOperacion & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
'                                    strCodSucursalTransferencia & "','" & strCodAgenciaTransferencia & "','" & _
'                                    "','" & strCodEjecutivo & "','" & strClaseCliente & "','','" & _
'                                    strCodMonedaFondoTransferente & "',0," & dblCuotasAcumulado & "," & _
'                                    CDec(adoTransferido.Recordset.Fields("ValorCuota")) & ",'C','" & _
'                                    "','','X','X','','','','','','','','','','','" & _
'                                    Trim(adoTransferido.Recordset.Fields("NumFolio")) & "','','" & Convertyyyymmdd(Valor_Fecha) & "','" & _
'                                    "',0,0,0,'" & strNumSolicitud & "','X','','" & Estado_Activo & "','" & _
'                                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "') }"
'
'
'                                adoConn.Execute .CommandText
'
'                                .CommandText = "SELECT ValorParametro FROM AuxiliarParametro " & _
'                                "WHERE CodParametro='" & Trim(adoTransferido.Recordset.Fields("TipoIdentidad")) & "' AND CodTipoParametro='TIPIDE'"
'                                Set adoRegistro = .Execute
'
'                                If Not adoRegistro.EOF Then
'                                    strClaseCliente = Trim(adoRegistro("ValorParametro"))
'                                End If
'                                adoRegistro.Close
'
'                                .CommandType = adCmdStoredProc
'                                '*** Obtener el número del parámetro **
'                                .CommandText = "up_ACObtenerUltNumero"
'                                .Parameters.Append .CreateParameter("CodFondo", adChar, adParamInput, 3, strCodFondoTransferido)
'                                .Parameters.Append .CreateParameter("CodAdministradora", adChar, adParamInput, 3, gstrCodAdministradora)
'                                .Parameters.Append .CreateParameter("CodParametro", adChar, adParamInput, 6, Valor_NumCertificado)
'                                .Parameters.Append .CreateParameter("NuevoNumero", adChar, adParamOutput, 10, "")
'                                .Execute
'
'                                If Not .Parameters("NuevoNumero") Then
'                                    strNumCertificado = .Parameters("NuevoNumero").Value
'                                    .Parameters.Delete ("CodFondo")
'                                    .Parameters.Delete ("CodAdministradora")
'                                    .Parameters.Delete ("CodParametro")
'                                    .Parameters.Delete ("NuevoNumero")
'                                End If
'
'                                .CommandType = adCmdText
'
'                                '*** Guardar Detalle Operación ***
'                                .CommandText = "{ call up_GNAdicOperacionParticipeDetalle('" & _
'                                    strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                    gstrCodParticipeTransferente & "','" & strNumOperacion & "'," & _
'                                    intContador & ",'C','" & Trim(adoTransferido.Recordset.Fields("CodParticipe")) & "','" & _
'                                    strNumCertificado & "','" & Convertyyyymmdd(gdatFechaActual) & "'," & _
'                                    dblCuotasAcumulado & ",0," & _
'                                    CDec(adoTransferido.Recordset.Fields("ValorCuota")) & "," & _
'                                    "0,0,0,0,0,0,'') }"
'                                adoConn.Execute .CommandText
'
'                                '*** Guardar Certificados del Transferido***
'                                .CommandText = "{ call up_GNAdicCertificadoParticipe('" & _
'                                    strCodFondoTransferido & "','" & gstrCodAdministradora & "','" & _
'                                    Trim(adoTransferido.Recordset.Fields("CodParticipe")) & "','" & strNumCertificado & "','" & _
'                                    Convertyyyymmdd(gdatFechaActual) & "','" & _
'                                    Convertyyyymmdd(CVDate(adoTransferido.Recordset.Fields("FechaSuscripcion"))) & "','" & _
'                                    Trim(arrTipoOperacion(cboTipoOperacion.ListIndex)) & "','" & strCodClaseOperacion & "','" & _
'                                    strNumOperacion & "','C'," & _
'                                    dblCuotasAcumulado & "," & _
'                                    dblCuotasAcumulado & "," & _
'                                    CDec(adoTransferido.Recordset.Fields("ValorCuota")) & ",'" & _
'                                    strCodMonedaFondoTransferido & "','" & strClaseCliente & "','" & _
'                                    strCodEjecutivo & "','X','X','" & gstrLogin & "') }"
'                                adoConn.Execute .CommandText
'
'                                strFechaInicio = Convertyyyymmdd(gdatFechaActual)
'                                strFechaFin = Convertyyyymmdd(DateAdd("d", 1, gdatFechaActual))
'
'                                '*** Actualizar datos adicionales ***
'                                .CommandText = "UPDATE ParticipeCertificado SET " & _
'                                    "OrigenOperacion='" & adoRegistro2("NumOperacion") & "'," & _
'                                    "TipoOrigenOperacion='" & adoRegistro2("TipoOperacion") & "',ClaseOrigenOperacion='" & adoRegistro2("ClaseOperacion") & "'," & _
'                                    "NumCertificadoEliminado='" & adoRegistro2("NumOperacion") & "' " & _
'                                    "WHERE CodFondo='" & strCodFondoTransferido & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                                    "(FechaOperacion>='" & strFechaInicio & "' AND FechaOperacion<'" & strFechaFin & "') AND " & _
'                                    "NumCertificado='" & strNumCertificado & "' AND " & _
'                                    "CodParticipe='" & Trim(adoTransferido.Recordset.Fields("CodParticipe")) & "'"
'                                adoConn.Execute .CommandText
'
'                                '*** Actualizar Secuenciales ***
'                                .CommandText = "{ call up_ACActUltNumero('" & strCodFondoTransferido & "','" & gstrCodAdministradora & "','" & _
'                                    Valor_NumCertificado & "','" & strNumCertificado & "') }"
'                                adoConn.Execute .CommandText
'
'                                .CommandText = "{ call up_ACActUltNumero('" & strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                    Valor_NumOpeCertificado & "','" & strNumOperacion & "') }"
'                                adoConn.Execute .CommandText
'
'                                '*** Actualizar Forma de Ingreso en el Contrato ***
'                                .CommandText = "UPDATE ParticipeContrato SET " & _
'                                    "TipoIngreso='" & Trim(adoTransferido.Recordset.Fields("TipoIngreso")) & "'," & _
'                                    "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & Convertyyyymmdd(gdatFechaActual) & "' " & _
'                                    "WHERE CodParticipe='" & Trim(adoTransferido.Recordset.Fields("CodParticipe")) & "'"
'                                adoConn.Execute .CommandText
'
'                                strFechaInicio = Convertyyyymmdd(adoRegistro2("FechaOperacion")) 'Convertyyyymmdd(tdgTransferente.Columns(4))
'                                strFechaFin = Convertyyyymmdd(DateAdd("d", 1, adoRegistro2("FechaOperacion"))) 'Convertyyyymmdd(DateAdd("d", 1, tdgTransferente.Columns(4)))
'
'                                adoComm.CommandText = "UPDATE ParticipeCertificado SET " & _
'                                        "IndVigente='',NumFinOperacion='" & strNumOperacion & "',FechaRedencion='" & Convertyyyymmdd(gdatFechaActual) & "'," & _
'                                        "TipoFinOperacion='" & strCodTipoOperacion & "',ClaseFinOperacion='" & strCodClaseOperacion & "'," & _
'                                        "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & Convertyyyymmdd(gdatFechaActual) & "' " & _
'                                        "WHERE CodFondo='" & strCodFondoTransferente & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                                        "(FechaOperacion>='" & strFechaInicio & "' AND FechaOperacion<'" & strFechaFin & "') AND " & _
'                                        "NumCertificado='" & adoRegistro2("NumCertificado") & "' AND " & _
'                                        "CodParticipe='" & adoRegistro2("CodParticipe") & "'"
'                                adoConn.Execute .CommandText
'
'                            Else 'cambiamos a NO VIGENTE los certificados de la lista
'                                adoComm.CommandText = "SELECT CodFondo, CodParticipe, FechaOperacion, NumOperacion, FechaSuscripcion, CantCuotas, CantCuotasPagadas, ValorCuota, TipoOperacion, ClaseOperacion, NumCertificado FROM ParticipeCertificado " & _
'                                          "WHERE NumOperacion='" & adoTransferido.Recordset.Fields("NumOpCopia") & "' AND CodFondo='" & strCodFondoTransferido & "'"
'                                Set adoRegistro2 = adoComm.Execute
'
'                                If Not adoRegistro2.EOF Then
'                                    If CDec(adoTransferido.Recordset.Fields("CantCuotas")) < adoRegistro2("CantCuotas") Then
'                                        strCodClaseOperacion = Codigo_Clase_TransferenciaParcial
'                                    Else
'                                        strCodClaseOperacion = Codigo_Clase_TransferenciaTotal
'                                    End If
'                                End If
'
'                                strFechaInicio = Convertyyyymmdd(adoRegistro2("FechaOperacion")) 'Convertyyyymmdd(tdgTransferente.Columns(4))
'                                strFechaFin = Convertyyyymmdd(DateAdd("d", 1, adoRegistro2("FechaOperacion"))) 'Convertyyyymmdd(DateAdd("d", 1, tdgTransferente.Columns(4)))
'
'                                adoComm.CommandText = "UPDATE ParticipeCertificado SET " & _
'                                        "IndVigente='',NumFinOperacion='" & strNumOperacion & "',FechaRedencion='" & Convertyyyymmdd(gdatFechaActual) & "'," & _
'                                        "TipoFinOperacion='" & strCodTipoOperacion & "',ClaseFinOperacion='" & strCodClaseOperacion & "'," & _
'                                        "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & Convertyyyymmdd(gdatFechaActual) & "' " & _
'                                        "WHERE CodFondo='" & strCodFondoTransferente & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                                        "(FechaOperacion>='" & strFechaInicio & "' AND FechaOperacion<'" & strFechaFin & "') AND " & _
'                                        "NumCertificado='" & adoRegistro2("NumCertificado") & "' AND " & _
'                                        "CodParticipe='" & adoRegistro2("CodParticipe") & "'"
'                                adoConn.Execute .CommandText
'                                'Next
'
'                                'adoRegistro3.Close: Set adoRegistro3 = Nothing
'                            End If
'
'                        ElseIf adoTransferido.Recordset.Fields(5) = "01" Then 'es el certificado generado por el saldo que sobró -- TipoFormaIngreso='01'
'
'                                .CommandText = "SELECT ValorParametro FROM AuxiliarParametro " & _
'                                "WHERE CodParametro='" & strCodTipoDocumentoTransferente & "' AND CodTipoParametro='TIPIDE'"
'                                Set adoRegistro = .Execute
'
'                                If Not adoRegistro.EOF Then
'                                    strClaseCliente = Trim(adoRegistro("ValorParametro"))
'                                End If
'                                adoRegistro.Close
'
'
'                                adoComm.CommandText = "SELECT CodFondo, CodParticipe, FechaOperacion, NumOperacion, FechaSuscripcion, CantCuotas, CantCuotasPagadas, ValorCuota, TipoOperacion, ClaseOperacion, NumCertificado FROM ParticipeCertificado " & _
'                                          "WHERE NumOperacion='" & adoTransferido.Recordset.Fields("NumOpCopia") & "' AND CodFondo='" & strCodFondoTransferido & "'"
'                                Set adoRegistro2 = adoComm.Execute
'
'                                If Not adoRegistro2.EOF Then
'                                    If CDec(adoTransferido.Recordset.Fields("CantCuotas")) < adoRegistro2("CantCuotas") Then
'                                        strCodClaseOperacion = Codigo_Clase_TransferenciaParcial
'                                    Else
'                                        strCodClaseOperacion = Codigo_Clase_TransferenciaTotal
'                                    End If
'                                End If
'
'
'                                If gstrCodParticipeTransferente <> Trim(adoTransferido.Recordset.Fields("CodParticipe")) Then
'                                    .CommandType = adCmdStoredProc
'
'                                    '*** Obtener el número del parámetro **
'                                    .CommandText = "up_ACObtenerUltNumero"
'                                    .Parameters.Append .CreateParameter("CodFondo", adChar, adParamInput, 3, strCodFondo)
'                                    .Parameters.Append .CreateParameter("CodAdministradora", adChar, adParamInput, 3, gstrCodAdministradora)
'                                    .Parameters.Append .CreateParameter("CodParametro", adChar, adParamInput, 6, Valor_NumSolicitud)
'                                    .Parameters.Append .CreateParameter("NuevoNumero", adChar, adParamOutput, 10, "")
'                                    .Execute
'
'                                    If Not .Parameters("NuevoNumero") Then
'                                        strNumSolicitud = .Parameters("NuevoNumero").Value
'                                        .Parameters.Delete ("CodFondo")
'                                        .Parameters.Delete ("CodAdministradora")
'                                        .Parameters.Delete ("CodParametro")
'                                        .Parameters.Delete ("NuevoNumero")
'                                    End If
'
'                                    .CommandType = adCmdText
'
'                                    '*** Guardar Solicitud ***
'                                    .CommandText = "{ call up_PRManTransferenciaParticipe('" & _
'                                        strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                        strNumSolicitud & "','" & gstrCodParticipeTransferente & "','" & _
'                                        Trim(adoTransferido.Recordset.Fields("NumFolio")) & "','','" & _
'                                        strCodSucursal & "','" & strCodSucursal & "','" & _
'                                        strCodAgencia & "','" & strCodAgencia & "','" & _
'                                        strCodEjecutivo & "','" & strCodEjecutivo & "','" & _
'                                        Codigo_FormaIngreso_Suscripcion & "','" & "03" & "','" & _
'                                        Convertyyyymmdd(CVDate(lblFechaSolicitud.Caption)) & Space(1) & Format(dtpHoraSolicitud.Value, "hh:mm") & "','" & _
'                                        Convertyyyymmdd(CVDate(lblFechaSolicitud.Caption)) & "','" & _
'                                        strCodMonedaFondoTransferente & "',0,"
'
'                                    .CommandText = .CommandText & CDec(adoTransferido.Recordset.Fields("ValorCuota")) & "," & CDec(adoTransferido.Recordset.Fields("CantCuotas")) & "," & _
'                                        "0,0,0,0,'" & _
'                                        "','','','','" & _
'                                        "','','','','" & _
'                                        "','','" & _
'                                        "X','','" & _
'                                        Convertyyyymmdd(CVDate(Valor_Fecha)) & "','X','','" & _
'                                        "',0,'','" & _
'                                        Estado_Solicitud_Procesada & "','" & _
'                                        gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
'                                        gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
'                                        strCodTipoOperacion & "','" & Trim(lblDescripParticipeTransferente.Caption) & "','" & strCodTipoTransferencia & "','" & Trim(txtEspecificarOtro.Text) & "','I') }"
'                                    adoConn.Execute .CommandText
'
'                                    '*** Guardar Detalle Solicitud ***
'                                    .CommandText = "{ call up_PRManTransferenciaParticipeDetalle('" & _
'                                        strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                        strNumSolicitud & "','1','" & gstrCodParticipeTransferente & "','" & _
'                                        "'," & dblCuotasAcumulado & ",'" & _
'                                        gstrCodParticipeTransferido & "','" & Trim(lblDescripParticipeTransferido.Caption) & "','" & _
'                                        "X','I') }"
'                                    adoConn.Execute .CommandText
'
'                                    '*** Actualizar Secuenciales ***
'                                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                        Valor_NumSolicitud & "','" & strNumSolicitud & "') }"
'                                    adoConn.Execute .CommandText
'                                End If
'
'                                .CommandType = adCmdStoredProc
'                                '*** Obtener el número del parámetro **
'                                .CommandText = "up_ACObtenerUltNumero"
'                                .Parameters.Append .CreateParameter("CodFondo", adChar, adParamInput, 3, strCodFondoTransferente)
'                                .Parameters.Append .CreateParameter("CodAdministradora", adChar, adParamInput, 3, gstrCodAdministradora)
'                                .Parameters.Append .CreateParameter("CodParametro", adChar, adParamInput, 6, Valor_NumOpeCertificado)
'                                .Parameters.Append .CreateParameter("NuevoNumero", adChar, adParamOutput, 10, Valor_Caracter)
'                                .Execute
'
'                                If Not .Parameters("NuevoNumero") Then
'                                    strNumOperacion = .Parameters("NuevoNumero").Value
'                                    .Parameters.Delete ("CodFondo")
'                                    .Parameters.Delete ("CodAdministradora")
'                                    .Parameters.Delete ("CodParametro")
'                                    .Parameters.Delete ("NuevoNumero")
'                                End If
'
'                                .CommandType = adCmdText
'
'                                '*** Guardar Operación ***
'                                .CommandText = "{ call up_GNAdicOperacionParticipe('" & _
'                                    strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                    gstrCodParticipeTransferente & "','" & strNumOperacion & "','" & _
'                                    Convertyyyymmdd(gdatFechaActual) & "','" & Codigo_FormaIngreso_Suscripcion & "','" & _
'                                    "01" & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
'                                    strCodSucursalTransferencia & "','" & strCodAgenciaTransferencia & "','" & _
'                                    "','" & strCodEjecutivo & "','" & strClaseCliente & "','','" & _
'                                    strCodMonedaFondoTransferente & "',0," & CDec(adoTransferido.Recordset.Fields("CantCuotas")) & "," & _
'                                    CDec(adoTransferido.Recordset.Fields("ValorCuota")) & ",'C','" & _
'                                    "','','X','X','','','','','','','','','','','" & _
'                                    Trim(adoTransferido.Recordset.Fields("NumFolio")) & "','','" & Convertyyyymmdd(Valor_Fecha) & "','" & _
'                                    "',0,0,0,'" & strNumSolicitud & "','X','','" & Estado_Activo & "','" & _
'                                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "') }"
'                                adoConn.Execute .CommandText
'
'                                .CommandText = "SELECT ValorParametro FROM AuxiliarParametro " & _
'                                "WHERE CodParametro='" & Trim(adoTransferido.Recordset.Fields("TipoIdentidad")) & "' AND CodTipoParametro='TIPIDE'"
'                                Set adoRegistro = .Execute
'
'                                If Not adoRegistro.EOF Then
'                                    strClaseCliente = Trim(adoRegistro("ValorParametro"))
'                                End If
'                                adoRegistro.Close
'
'                                .CommandType = adCmdStoredProc
'                                '*** Obtener el número del parámetro **
'                                .CommandText = "up_ACObtenerUltNumero"
'                                .Parameters.Append .CreateParameter("CodFondo", adChar, adParamInput, 3, strCodFondoTransferido)
'                                .Parameters.Append .CreateParameter("CodAdministradora", adChar, adParamInput, 3, gstrCodAdministradora)
'                                .Parameters.Append .CreateParameter("CodParametro", adChar, adParamInput, 6, Valor_NumCertificado)
'                                .Parameters.Append .CreateParameter("NuevoNumero", adChar, adParamOutput, 10, "")
'                                .Execute
'
'                                If Not .Parameters("NuevoNumero") Then
'                                    strNumCertificado = .Parameters("NuevoNumero").Value
'                                    .Parameters.Delete ("CodFondo")
'                                    .Parameters.Delete ("CodAdministradora")
'                                    .Parameters.Delete ("CodParametro")
'                                    .Parameters.Delete ("NuevoNumero")
'                                End If
'
'                                .CommandType = adCmdText
'
'                                '*** Guardar Detalle Operación ***
'                                .CommandText = "{ call up_GNAdicOperacionParticipeDetalle('" & _
'                                    strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                    gstrCodParticipeTransferente & "','" & strNumOperacion & "'," & _
'                                    intContador & ",'C','" & Trim(adoTransferido.Recordset.Fields("CodParticipe")) & "','" & _
'                                    strNumCertificado & "','" & Convertyyyymmdd(gdatFechaActual) & "'," & _
'                                    CDec(adoTransferido.Recordset.Fields("CantCuotas")) & ",0," & _
'                                    CDec(adoTransferido.Recordset.Fields("ValorCuota")) & "," & _
'                                    "0,0,0,0,0,0,'') }"
'                                adoConn.Execute .CommandText
'
'                                '*** Guardar Certificados del Transferido***
'                                .CommandText = "{ call up_GNAdicCertificadoParticipe('" & _
'                                    strCodFondoTransferido & "','" & gstrCodAdministradora & "','" & _
'                                    Trim(adoTransferido.Recordset.Fields("CodParticipe")) & "','" & strNumCertificado & "','" & _
'                                    Convertyyyymmdd(gdatFechaActual) & "','" & _
'                                    Convertyyyymmdd(CVDate(adoTransferido.Recordset.Fields("FechaSuscripcion"))) & "','" & _
'                                    Codigo_FormaIngreso_Suscripcion & "','" & "01" & "','" & _
'                                    strNumOperacion & "','C'," & _
'                                    CDec(adoTransferido.Recordset.Fields("CantCuotas")) & "," & _
'                                    CDec(adoTransferido.Recordset.Fields("CantCuotas")) & "," & _
'                                    CDec(adoTransferido.Recordset.Fields("ValorCuota")) & ",'" & _
'                                    strCodMonedaFondoTransferido & "','" & strClaseCliente & "','" & _
'                                    strCodEjecutivo & "','X','X','" & gstrLogin & "') }"
'                                adoConn.Execute .CommandText
'
'                                strFechaInicio = Convertyyyymmdd(gdatFechaActual)
'                                strFechaFin = Convertyyyymmdd(DateAdd("d", 1, gdatFechaActual))
'
'                                '*** Actualizar datos adicionales ***
'                                .CommandText = "UPDATE ParticipeCertificado SET " & _
'                                    "OrigenOperacion='" & adoRegistro2("NumOperacion") & "'," & _
'                                    "TipoOrigenOperacion='" & adoRegistro2("TipoOperacion") & "',ClaseOrigenOperacion='" & adoRegistro2("ClaseOperacion") & "'," & _
'                                    "NumCertificadoEliminado='" & adoRegistro2("NumOperacion") & "' " & _
'                                    "WHERE CodFondo='" & strCodFondoTransferido & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                                    "(FechaOperacion>='" & strFechaInicio & "' AND FechaOperacion<'" & strFechaFin & "') AND " & _
'                                    "NumCertificado='" & strNumCertificado & "' AND " & _
'                                    "CodParticipe='" & Trim(adoTransferido.Recordset.Fields("CodParticipe")) & "'"
'                                adoConn.Execute .CommandText
'
'                                '*** Actualizar Secuenciales ***
'                                .CommandText = "{ call up_ACActUltNumero('" & strCodFondoTransferido & "','" & gstrCodAdministradora & "','" & _
'                                    Valor_NumCertificado & "','" & strNumCertificado & "') }"
'                                adoConn.Execute .CommandText
'
'                                .CommandText = "{ call up_ACActUltNumero('" & strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                    Valor_NumOpeCertificado & "','" & strNumOperacion & "') }"
'                                adoConn.Execute .CommandText
'
'                                '*** Actualizar Forma de Ingreso en el Contrato ***
'                                .CommandText = "UPDATE ParticipeContrato SET " & _
'                                    "TipoIngreso='" & Trim(adoTransferido.Recordset.Fields("TipoIngreso")) & "'," & _
'                                    "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & Convertyyyymmdd(gdatFechaActual) & "' " & _
'                                    "WHERE CodParticipe='" & Trim(adoTransferido.Recordset.Fields("CodParticipe")) & "'"
'                                adoConn.Execute .CommandText
'
'
'                                Set adoRegistro3 = New ADODB.Recordset
'
'                                adoComm.CommandText = "SELECT COUNT(*) NumRegistros FROM CertificadoTransferidoTmp"
'                                Set adoRegistro3 = adoComm.Execute
'
'                                If Not adoRegistro3.EOF Then
'                                    intContador = adoRegistro3("NumRegistros")
'                                End If
'
'                                strFechaInicio = Convertyyyymmdd(adoRegistro2("FechaOperacion")) 'Convertyyyymmdd(tdgTransferente.Columns(4))
'                                strFechaFin = Convertyyyymmdd(DateAdd("d", 1, adoRegistro2("FechaOperacion"))) 'Convertyyyymmdd(DateAdd("d", 1, tdgTransferente.Columns(4)))
'
'                                .CommandText = "UPDATE ParticipeCertificado SET " & _
'                                        "IndVigente='',NumFinOperacion='" & strNumOperacion & "',FechaRedencion='" & Convertyyyymmdd(gdatFechaActual) & "'," & _
'                                        "TipoFinOperacion='" & strCodTipoOperacion & "',ClaseFinOperacion='" & strCodClaseOperacion & "'," & _
'                                        "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & Convertyyyymmdd(gdatFechaActual) & "' " & _
'                                        "WHERE CodFondo='" & strCodFondoTransferente & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                                        "(FechaOperacion>='" & strFechaInicio & "' AND FechaOperacion<'" & strFechaFin & "') AND " & _
'                                        "NumCertificado='" & adoRegistro2("NumCertificado") & "' AND " & _
'                                        "CodParticipe='" & adoRegistro2("CodParticipe") & "'"
'                                adoConn.Execute .CommandText
'
'                                adoRegistro3.Close: Set adoRegistro3 = Nothing
'
'
'                        End If
'
'                        adoTransferido.Recordset.MoveNext
'
'                    Loop
'
'                End With
'
'            'Si no desea agruparlos...'
'            Else
'
'                Set adoRegistro = New ADODB.Recordset
'                With adoComm
'
'                    intContador = 1
'
'                    adoTransferido.Recordset.MoveFirst
'
'                    Do While Not adoTransferido.Recordset.EOF
'                        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro " & _
'                            "WHERE CodParametro='" & strCodTipoDocumentoTransferente & "' AND CodTipoParametro='TIPIDE'"
'                        Set adoRegistro = .Execute
'
'                        If Not adoRegistro.EOF Then
'                            strClaseCliente = Trim(adoRegistro("ValorParametro"))
'                        End If
'                        adoRegistro.Close
'
'
'                        adoComm.CommandText = "SELECT CodFondo, CodParticipe, FechaOperacion, NumOperacion, FechaSuscripcion, CantCuotas, CantCuotasPagadas, ValorCuota, TipoOperacion, ClaseOperacion, NumCertificado FROM ParticipeCertificado " & _
'                                  "WHERE NumOperacion='" & adoTransferido.Recordset.Fields("NumOpCopia") & "' AND CodFondo='" & strCodFondoTransferido & "'"
'                        Set adoRegistro2 = adoComm.Execute
'
'                        If Not adoRegistro2.EOF Then
'                            If CDec(adoTransferido.Recordset.Fields("CantCuotas")) < adoRegistro2("CantCuotas") Then
'                                strCodClaseOperacion = Codigo_Clase_TransferenciaParcial
'                            Else
'                                strCodClaseOperacion = Codigo_Clase_TransferenciaTotal
'                            End If
'                        End If
'
'
'                        If gstrCodParticipeTransferente <> Trim(adoTransferido.Recordset.Fields("CodParticipe")) Then
'                            .CommandType = adCmdStoredProc
'
'                            '*** Obtener el número del parámetro **
'                            .CommandText = "up_ACObtenerUltNumero"
'                            .Parameters.Append .CreateParameter("CodFondo", adChar, adParamInput, 3, strCodFondo)
'                            .Parameters.Append .CreateParameter("CodAdministradora", adChar, adParamInput, 3, gstrCodAdministradora)
'                            .Parameters.Append .CreateParameter("CodParametro", adChar, adParamInput, 6, Valor_NumSolicitud)
'                            .Parameters.Append .CreateParameter("NuevoNumero", adChar, adParamOutput, 10, "")
'                            .Execute
'
'                            If Not .Parameters("NuevoNumero") Then
'                                strNumSolicitud = .Parameters("NuevoNumero").Value
'                                .Parameters.Delete ("CodFondo")
'                                .Parameters.Delete ("CodAdministradora")
'                                .Parameters.Delete ("CodParametro")
'                                .Parameters.Delete ("NuevoNumero")
'                            End If
'
'                            If adoTransferido.Recordset.Fields(5) = "02" Then
'
'                                .CommandType = adCmdText
'
'                                '*** Guardar Solicitud ***
'                                .CommandText = "{ call up_PRManTransferenciaParticipe('" & _
'                                    strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                    strNumSolicitud & "','" & gstrCodParticipeTransferente & "','" & _
'                                    Trim(adoTransferido.Recordset.Fields("NumFolio")) & "','','" & _
'                                    strCodSucursal & "','" & strCodSucursal & "','" & _
'                                    strCodAgencia & "','" & strCodAgencia & "','" & _
'                                    strCodEjecutivo & "','" & strCodEjecutivo & "','" & _
'                                    strCodTipoOperacion & "','" & strCodClaseOperacion & "','" & _
'                                    Convertyyyymmdd(CVDate(lblFechaSolicitud.Caption)) & Space(1) & Format(dtpHoraSolicitud.Value, "hh:mm") & "','" & _
'                                    Convertyyyymmdd(CVDate(lblFechaSolicitud.Caption)) & "','" & _
'                                    strCodMonedaFondoTransferente & "',0,"
'
'                                .CommandText = .CommandText & CDec(adoTransferido.Recordset.Fields("ValorCuota")) & "," & CDec(adoTransferido.Recordset.Fields("CantCuotas")) & "," & _
'                                    "0,0,0,0,'" & _
'                                    "','','','','" & _
'                                    "','','','','" & _
'                                    "','','" & _
'                                    "X','','" & _
'                                    Convertyyyymmdd(CVDate(Valor_Fecha)) & "','X','','" & _
'                                    "',0,'','" & _
'                                    Estado_Solicitud_Procesada & "','" & _
'                                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
'                                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
'                                    strCodTipoOperacion & "','" & Trim(lblDescripParticipeTransferente.Caption) & "','" & strCodTipoTransferencia & "','" & Trim(txtEspecificarOtro.Text) & "','I') }"
'                                adoConn.Execute .CommandText
'
'                                '*** Guardar Detalle Solicitud ***
'                                .CommandText = "{ call up_PRManTransferenciaParticipeDetalle('" & _
'                                    strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                    strNumSolicitud & "','1','" & gstrCodParticipeTransferente & "','" & _
'                                    "'," & dblCuotasAcumulado & ",'" & _
'                                    gstrCodParticipeTransferido & "','" & Trim(lblDescripParticipeTransferido.Caption) & "','" & _
'                                    "X','I') }"
'                                adoConn.Execute .CommandText
'
'                                '*** Actualizar Secuenciales ***
'                                .CommandText = "{ call up_ACActUltNumero('" & strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                    Valor_NumSolicitud & "','" & strNumSolicitud & "') }"
'                                adoConn.Execute .CommandText
'
'                            ElseIf adoTransferido.Recordset.Fields(5) = "01" Then
'
'                                .CommandType = adCmdText
'
'                                '*** Guardar Solicitud ***
'                                .CommandText = "{ call up_PRManTransferenciaParticipe('" & _
'                                    strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                    strNumSolicitud & "','" & gstrCodParticipeTransferente & "','" & _
'                                    Trim(adoTransferido.Recordset.Fields("NumFolio")) & "','','" & _
'                                    strCodSucursal & "','" & strCodSucursal & "','" & _
'                                    strCodAgencia & "','" & strCodAgencia & "','" & _
'                                    strCodEjecutivo & "','" & strCodEjecutivo & "','" & _
'                                    Codigo_FormaIngreso_Suscripcion & "','" & "03" & "','" & _
'                                    Convertyyyymmdd(CVDate(lblFechaSolicitud.Caption)) & Space(1) & Format(dtpHoraSolicitud.Value, "hh:mm") & "','" & _
'                                    Convertyyyymmdd(CVDate(lblFechaSolicitud.Caption)) & "','" & _
'                                    strCodMonedaFondoTransferente & "',0,"
'
'                                .CommandText = .CommandText & CDec(adoTransferido.Recordset.Fields("ValorCuota")) & "," & CDec(adoTransferido.Recordset.Fields("CantCuotas")) & "," & _
'                                    "0,0,0,0,'" & _
'                                    "','','','','" & _
'                                    "','','','','" & _
'                                    "','','" & _
'                                    "X','','" & _
'                                    Convertyyyymmdd(CVDate(Valor_Fecha)) & "','X','','" & _
'                                    "',0,'','" & _
'                                    Estado_Solicitud_Procesada & "','" & _
'                                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
'                                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
'                                    strCodTipoOperacion & "','" & Trim(lblDescripParticipeTransferente.Caption) & "','" & strCodTipoTransferencia & "','" & Trim(txtEspecificarOtro.Text) & "','I') }"
'                                adoConn.Execute .CommandText
'
'                                '*** Guardar Detalle Solicitud ***
'                                .CommandText = "{ call up_PRManTransferenciaParticipeDetalle('" & _
'                                    strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                    strNumSolicitud & "','1','" & gstrCodParticipeTransferente & "','" & _
'                                    "'," & dblCuotasAcumulado & ",'" & _
'                                    gstrCodParticipeTransferido & "','" & Trim(lblDescripParticipeTransferido.Caption) & "','" & _
'                                    "X','I') }"
'                                adoConn.Execute .CommandText
'                                'CDec(adoTransferido.Recordset.Fields("CantCuotas"))
'                                '*** Actualizar Secuenciales ***
'                                .CommandText = "{ call up_ACActUltNumero('" & strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                    Valor_NumSolicitud & "','" & strNumSolicitud & "') }"
'                                adoConn.Execute .CommandText
'
'                            End If
'
'
'                        End If
'
'                        .CommandType = adCmdStoredProc
'                        '*** Obtener el número del parámetro **
'                        .CommandText = "up_ACObtenerUltNumero"
'                        .Parameters.Append .CreateParameter("CodFondo", adChar, adParamInput, 3, strCodFondoTransferente)
'                        .Parameters.Append .CreateParameter("CodAdministradora", adChar, adParamInput, 3, gstrCodAdministradora)
'                        .Parameters.Append .CreateParameter("CodParametro", adChar, adParamInput, 6, Valor_NumOpeCertificado)
'                        .Parameters.Append .CreateParameter("NuevoNumero", adChar, adParamOutput, 10, Valor_Caracter)
'                        .Execute
'
'                        If Not .Parameters("NuevoNumero") Then
'                            strNumOperacion = .Parameters("NuevoNumero").Value
'                            .Parameters.Delete ("CodFondo")
'                            .Parameters.Delete ("CodAdministradora")
'                            .Parameters.Delete ("CodParametro")
'                            .Parameters.Delete ("NuevoNumero")
'                        End If
'
'                        If adoTransferido.Recordset.Fields(5) = "02" Then
'
'                            .CommandType = adCmdText
'
'                            '*** Guardar Operación ***
'                            .CommandText = "{ call up_GNAdicOperacionParticipe('" & _
'                                strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                gstrCodParticipeTransferente & "','" & strNumOperacion & "','" & _
'                                Convertyyyymmdd(gdatFechaActual) & "','" & Trim(arrTipoOperacion(cboTipoOperacion.ListIndex)) & "','" & _
'                                strCodClaseOperacion & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
'                                strCodSucursalTransferencia & "','" & strCodAgenciaTransferencia & "','" & _
'                                "','" & strCodEjecutivo & "','" & strClaseCliente & "','','" & _
'                                strCodMonedaFondoTransferente & "',0," & CDec(adoTransferido.Recordset.Fields("CantCuotas")) & "," & _
'                                CDec(adoTransferido.Recordset.Fields("ValorCuota")) & ",'C','" & _
'                                "','','X','X','','','','','','','','','','','" & _
'                                Trim(adoTransferido.Recordset.Fields("NumFolio")) & "','','" & Convertyyyymmdd(Valor_Fecha) & "','" & _
'                                "',0,0,0,'" & strNumSolicitud & "','X','','" & Estado_Activo & "','" & _
'                                gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "') }"
'                            adoConn.Execute .CommandText
'
'                        ElseIf adoTransferido.Recordset.Fields(5) = "01" Then
'
'                            .CommandType = adCmdText
'
'                            '*** Guardar Operación ***
'                            .CommandText = "{ call up_GNAdicOperacionParticipe('" & _
'                                strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                                gstrCodParticipeTransferente & "','" & strNumOperacion & "','" & _
'                                Convertyyyymmdd(gdatFechaActual) & "','" & Codigo_FormaIngreso_Suscripcion & "','" & _
'                                "01" & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
'                                strCodSucursalTransferencia & "','" & strCodAgenciaTransferencia & "','" & _
'                                "','" & strCodEjecutivo & "','" & strClaseCliente & "','','" & _
'                                strCodMonedaFondoTransferente & "',0," & CDec(adoTransferido.Recordset.Fields("CantCuotas")) & "," & _
'                                CDec(adoTransferido.Recordset.Fields("ValorCuota")) & ",'C','" & _
'                                "','','X','X','','','','','','','','','','','" & _
'                                Trim(adoTransferido.Recordset.Fields("NumFolio")) & "','','" & Convertyyyymmdd(Valor_Fecha) & "','" & _
'                                "',0,0,0,'" & strNumSolicitud & "','X','','" & Estado_Activo & "','" & _
'                                gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "') }"
'                            adoConn.Execute .CommandText
'
'                        End If
'
'                        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro " & _
'                        "WHERE CodParametro='" & Trim(adoTransferido.Recordset.Fields("TipoIdentidad")) & "' AND CodTipoParametro='TIPIDE'"
'                        Set adoRegistro = .Execute
'
'                        If Not adoRegistro.EOF Then
'                            strClaseCliente = Trim(adoRegistro("ValorParametro"))
'                        End If
'                        adoRegistro.Close
'
'                        .CommandType = adCmdStoredProc
'                        '*** Obtener el número del parámetro **
'                        .CommandText = "up_ACObtenerUltNumero"
'                        .Parameters.Append .CreateParameter("CodFondo", adChar, adParamInput, 3, strCodFondoTransferido)
'                        .Parameters.Append .CreateParameter("CodAdministradora", adChar, adParamInput, 3, gstrCodAdministradora)
'                        .Parameters.Append .CreateParameter("CodParametro", adChar, adParamInput, 6, Valor_NumCertificado)
'                        .Parameters.Append .CreateParameter("NuevoNumero", adChar, adParamOutput, 10, "")
'                        .Execute
'
'                        If Not .Parameters("NuevoNumero") Then
'                            strNumCertificado = .Parameters("NuevoNumero").Value
'                            .Parameters.Delete ("CodFondo")
'                            .Parameters.Delete ("CodAdministradora")
'                            .Parameters.Delete ("CodParametro")
'                            .Parameters.Delete ("NuevoNumero")
'                        End If
'
'                        .CommandType = adCmdText
'
'                        '*** Guardar Detalle Operación ***
'                        .CommandText = "{ call up_GNAdicOperacionParticipeDetalle('" & _
'                            strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                            gstrCodParticipeTransferente & "','" & strNumOperacion & "'," & _
'                            intContador & ",'C','" & Trim(adoTransferido.Recordset.Fields("CodParticipe")) & "','" & _
'                            strNumCertificado & "','" & Convertyyyymmdd(gdatFechaActual) & "'," & _
'                            CDec(adoTransferido.Recordset.Fields("CantCuotas")) & ",0," & _
'                            CDec(adoTransferido.Recordset.Fields("ValorCuota")) & "," & _
'                            "0,0,0,0,0,0,'') }"
'                        adoConn.Execute .CommandText
'
'                        If adoTransferido.Recordset.Fields(5) = "02" Then
'
'                            '*** Guardar Certificados del Transferido***
'                            .CommandText = "{ call up_GNAdicCertificadoParticipe('" & _
'                                strCodFondoTransferido & "','" & gstrCodAdministradora & "','" & _
'                                Trim(adoTransferido.Recordset.Fields("CodParticipe")) & "','" & strNumCertificado & "','" & _
'                                Convertyyyymmdd(gdatFechaActual) & "','" & _
'                                Convertyyyymmdd(CVDate(adoTransferido.Recordset.Fields("FechaSuscripcion"))) & "','" & _
'                                Trim(arrTipoOperacion(cboTipoOperacion.ListIndex)) & "','" & strCodClaseOperacion & "','" & _
'                                strNumOperacion & "','C'," & _
'                                CDec(adoTransferido.Recordset.Fields("CantCuotas")) & "," & _
'                                CDec(adoTransferido.Recordset.Fields("CantCuotas")) & "," & _
'                                CDec(adoTransferido.Recordset.Fields("ValorCuota")) & ",'" & _
'                                strCodMonedaFondoTransferido & "','" & strClaseCliente & "','" & _
'                                strCodEjecutivo & "','X','X','" & gstrLogin & "') }"
'                            adoConn.Execute .CommandText
'
'                        ElseIf adoTransferido.Recordset.Fields(5) = "01" Then
'
'                            '*** Guardar Certificados del Transferido***
'                            .CommandText = "{ call up_GNAdicCertificadoParticipe('" & _
'                                strCodFondoTransferido & "','" & gstrCodAdministradora & "','" & _
'                                Trim(adoTransferido.Recordset.Fields("CodParticipe")) & "','" & strNumCertificado & "','" & _
'                                Convertyyyymmdd(gdatFechaActual) & "','" & _
'                                Convertyyyymmdd(CVDate(adoTransferido.Recordset.Fields("FechaSuscripcion"))) & "','" & _
'                                Codigo_FormaIngreso_Suscripcion & "','" & "01" & "','" & _
'                                strNumOperacion & "','C'," & _
'                                CDec(adoTransferido.Recordset.Fields("CantCuotas")) & "," & _
'                                CDec(adoTransferido.Recordset.Fields("CantCuotas")) & "," & _
'                                CDec(adoTransferido.Recordset.Fields("ValorCuota")) & ",'" & _
'                                strCodMonedaFondoTransferido & "','" & strClaseCliente & "','" & _
'                                strCodEjecutivo & "','X','X','" & gstrLogin & "') }"
'                            adoConn.Execute .CommandText
'
'                        End If
'
'                        strFechaInicio = Convertyyyymmdd(gdatFechaActual)
'                        strFechaFin = Convertyyyymmdd(DateAdd("d", 1, gdatFechaActual))
'
'                        '*** Actualizar datos adicionales ***
'                        .CommandText = "UPDATE ParticipeCertificado SET " & _
'                            "OrigenOperacion='" & adoRegistro2("NumOperacion") & "'," & _
'                            "TipoOrigenOperacion='" & adoRegistro2("TipoOperacion") & "',ClaseOrigenOperacion='" & adoRegistro2("ClaseOperacion") & "'," & _
'                            "NumCertificadoEliminado='" & adoRegistro2("NumOperacion") & "' " & _
'                            "WHERE CodFondo='" & strCodFondoTransferido & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                            "(FechaOperacion>='" & strFechaInicio & "' AND FechaOperacion<'" & strFechaFin & "') AND " & _
'                            "NumCertificado='" & strNumCertificado & "' AND " & _
'                            "CodParticipe='" & Trim(adoTransferido.Recordset.Fields("CodParticipe")) & "'"
'                        adoConn.Execute .CommandText
'
'                        '*** Actualizar Secuenciales ***
'                        .CommandText = "{ call up_ACActUltNumero('" & strCodFondoTransferido & "','" & gstrCodAdministradora & "','" & _
'                            Valor_NumCertificado & "','" & strNumCertificado & "') }"
'                        adoConn.Execute .CommandText
'
'                        .CommandText = "{ call up_ACActUltNumero('" & strCodFondoTransferente & "','" & gstrCodAdministradora & "','" & _
'                            Valor_NumOpeCertificado & "','" & strNumOperacion & "') }"
'                        adoConn.Execute .CommandText
'
'                        '*** Actualizar Forma de Ingreso en el Contrato ***
'                        .CommandText = "UPDATE ParticipeContrato SET " & _
'                            "TipoIngreso='" & Trim(adoTransferido.Recordset.Fields("TipoIngreso")) & "'," & _
'                            "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & Convertyyyymmdd(gdatFechaActual) & "' " & _
'                            "WHERE CodParticipe='" & Trim(adoTransferido.Recordset.Fields("CodParticipe")) & "'"
'                        adoConn.Execute .CommandText
'
'                        Set adoRegistro3 = New ADODB.Recordset
'
'                        adoComm.CommandText = "SELECT COUNT(*) NumRegistros FROM CertificadoTransferidoTmp"
'                        Set adoRegistro3 = adoComm.Execute
'
'                        If Not adoRegistro3.EOF Then
'                            intContador = adoRegistro3("NumRegistros")
'                        End If
'
'                        dblCuotas = CDec(adoTransferido.Recordset.Fields("CantCuotas"))
'                        dblCuotasAcumulado = dblCuotasAcumulado + dblCuotas
'
'                        strFechaInicio = Convertyyyymmdd(adoRegistro2("FechaOperacion")) 'Convertyyyymmdd(tdgTransferente.Columns(4))
'                        strFechaFin = Convertyyyymmdd(DateAdd("d", 1, adoRegistro2("FechaOperacion"))) 'Convertyyyymmdd(DateAdd("d", 1, tdgTransferente.Columns(4)))
'
'                        .CommandText = "UPDATE ParticipeCertificado SET " & _
'                                "IndVigente='',NumFinOperacion='" & strNumOperacion & "',FechaRedencion='" & Convertyyyymmdd(gdatFechaActual) & "'," & _
'                                "TipoFinOperacion='" & strCodTipoOperacion & "',ClaseFinOperacion='" & strCodClaseOperacion & "'," & _
'                                "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & Convertyyyymmdd(gdatFechaActual) & "' " & _
'                                "WHERE CodFondo='" & strCodFondoTransferente & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                                "(FechaOperacion>='" & strFechaInicio & "' AND FechaOperacion<'" & strFechaFin & "') AND " & _
'                                "NumCertificado='" & adoRegistro2("NumCertificado") & "' AND " & _
'                                "CodParticipe='" & adoRegistro2("CodParticipe") & "'"
'                        adoConn.Execute .CommandText
'
'                        adoRegistro3.Close: Set adoRegistro3 = Nothing
'
'                        adoTransferido.Recordset.MoveNext
'                    Loop
'
'                End With
'
'            End If
    
            Me.MousePointer = vbDefault
            
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
                                                                            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"

            cmdOpcion.Visible = True
            With tabTransferencia
                .TabEnabled(0) = True
                .Tab = 0
            End With
                        
            Call Buscar
        
        End If
    End If

End Sub

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String

    If tabTransferencia.Tab = 1 Then Exit Sub
    Dim proceder As Integer
    proceder = 1
    
    Select Case Index
        Case 1
        
            'If adoConsulta.RecordCount > 0 Then
                gstrNameRepo = "ParticipeTransferencia"
                            
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(7)
                ReDim aReportParamFn(5)
                ReDim aReportParamF(5)
    
                strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
                strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value))
                
                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "Hora"
                aReportParamFn(2) = "NombreEmpresa"
                aReportParamFn(3) = "Fondo"
                aReportParamFn(4) = "FechaDesde"
                aReportParamFn(5) = "FechaHasta"
                
                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Format(Time(), "hh:mm:ss")
                aReportParamF(2) = gstrNombreEmpresa & Space(1)
                aReportParamF(3) = Trim(cboFondo.Text)
                aReportParamF(4) = CStr(dtpFechaDesde.Value)
                aReportParamF(5) = CStr(dtpFechaHasta.Value)
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Codigo_Operacion_Transferencia
                aReportParamS(3) = strFechaDesde
                aReportParamS(4) = strFechaHasta
                aReportParamS(5) = strCodSucursal
                aReportParamS(6) = strCodAgencia
                aReportParamS(7) = strCodPromotor
                
                proceder = 1
'            Else
'                MsgBox "No existe información para los criterios seleccionados", vbCritical
'                proceder = 0
'            End If
            
    Case 2
            If Not adoConsulta.EOF Then
                gstrNameRepo = "SolicitudTransferencia"
                            
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(3)
                ReDim aReportParamFn(3)
                ReDim aReportParamF(3)
    
                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "Hora"
                aReportParamFn(2) = "NombreEmpresa"
                aReportParamFn(3) = "Fondo"
                
                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Format(Time(), "hh:mm:ss")
                aReportParamF(2) = gstrNombreEmpresa & Space(1)
                aReportParamF(3) = Trim(cboFondo.Text)
                
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = adoConsulta.Fields("CodParticipe")
                aReportParamS(3) = adoConsulta.Fields("NumOperacion")
            Else
                MsgBox "Debe Seleccionar una Solicitud de Transferencia para ver el Reporte", vbCritical
                proceder = 0
            End If
            
    Case 3
            If Not adoConsulta.EOF Then
                gstrNameRepo = "SolicitudTransferenciaNuevo"
                            
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(3)
                ReDim aReportParamFn(3)
                ReDim aReportParamF(3)
    
                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "Hora"
                aReportParamFn(2) = "NombreEmpresa"
                aReportParamFn(3) = "Fondo"
                
                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Format(Time(), "hh:mm:ss")
                aReportParamF(2) = gstrNombreEmpresa & Space(1)
                aReportParamF(3) = Trim(cboFondo.Text)
                
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = adoConsulta.Fields("CodParticipe")
                aReportParamS(3) = adoConsulta.Fields("NumOperacion")
            Else
                MsgBox "Debe Seleccionar una Solicitud de Transferencia para ver el Reporte", vbCritical
                proceder = 0
            End If
    End Select
    
    If proceder = 1 Then
        gstrSelFrml = ""
        frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"
    
        Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())
    
        frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
        frmReporte.Show vbModal
    
        Set frmReporte = Nothing
    
        Screen.MousePointer = vbNormal
    End If
    
End Sub

Private Function TodoOK() As Boolean

    Dim dblCuotasAcumulado  As Double, dblCuotas    As Double
    Dim intRegistro         As Integer
    Dim numOpC              As String
    Dim adoRegistro         As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
    'dblCuotasAcumulado = 0
    
    TodoOK = False
    
'    If adoTransferido.Recordset.RecordCount = 0 Then
'        MsgBox "No existen certificados a transferir", vbCritical, Me.Caption
'        Exit Function
'    End If
'
'    adoTransferido.Recordset.MoveFirst
'    blnSeleccion = False
'
'    Do While Not adoTransferido.Recordset.EOF
'        intRegistro = CInt(adoTransferido.Recordset.Fields("NumSecuencial")) + 1
'        numOpC = adoTransferido.Recordset.Fields("NumOpCopia")
'
'        adoComm.CommandText = "SELECT CodFondo, CodParticipe, FechaOperacion, NumOperacion, FechaSuscripcion, CantCuotas, CantCuotasPagadas, ValorCuota, NumCertificado FROM ParticipeCertificado " & _
'                              "WHERE NumOperacion='" & numOpC & "' AND CodFondo='" & strCodFondoTransferido & "'"
'        Set adoRegistro = adoComm.Execute
'
'        'dblCuotasAcumulado = dblCuotasAcumulado + CDbl(adoTransferido.Recordset.Fields("CantCuotas"))
'
'        If Not adoRegistro.EOF Then
'            If adoRegistro("CantCuotas") > CDbl(adoTransferido.Recordset.Fields("CantCuotas")) Then
''                If MsgBox("La cantidad de cuotas ha transferir no está completa para el certificado Nº " & adoRegistro("NumCertificado") & vbNewLine & vbNewLine & _
''                "Desea completarla como un nuevo certificado para el Transferente", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Function
'
'                If MsgBox("El certificado Nº " & adoRegistro("NumCertificado") & " tiene un saldo pendiente." & vbNewLine & vbNewLine & _
'                "¿Desea crear un nuevo certificado para quien transfiere?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Function
'
'                dblCuotas = adoRegistro("CantCuotas") - CDbl(adoTransferido.Recordset.Fields("CantCuotas"))
'
'                adoComm.CommandText = "INSERT INTO CertificadoTransferidoTmp VALUES('" & _
'                    strCodFondoTransferido & "','" & _
'                    gstrCodParticipeTransferente & "','" & Convertyyyymmdd(adoRegistro("FechaOperacion")) & "'," & _
'                    intRegistro & ",'" & Convertyyyymmdd(adoRegistro("FechaSuscripcion")) & "'," & _
'                    CDec(adoRegistro("ValorCuota")) & "," & dblCuotas & ",'" & Codigo_FormaIngreso_Suscripcion & "','" & _
'                    strCodTipoDocumentoTransferente & "','" & _
'                    Trim(txtNumPapeleta.Text) & "','" & numOpC & "')"
'                adoConn.Execute adoComm.CommandText
'
'                blnSeleccion = True
'            End If
'        End If
'
'        If gstrCodParticipeTransferente = Trim(adoTransferido.Recordset.Fields("CodParticipe")) Then blnSeleccion = True
'
'        adoRegistro.Close: Set adoRegistro = Nothing
'        adoTransferido.Recordset.MoveNext
'    Loop
'
'    Call ObtenerCertificadoTransferido
    
    '*** Si todo paso OK ***
    TodoOK = True

End Function

Public Sub Imprimir()

End Sub

Public Sub Modificar()

'    cmdOpcion.Visible = False
    With tabTransferencia
'        .TabEnabled(0) = False
'        .Tab = 1
        .Tab = 0
    End With
    
End Sub

Public Sub ObtenerCertificados()

    Dim strSql  As String
        
    Set adoTransferente = New ADODB.Recordset
        
    strSql = "SELECT FechaSuscripcion,ValorCuota,CantCuotas,NumCertificado,FechaOperacion,NumOperacion,TipoOperacion,ClaseOperacion FROM ParticipeCertificado " & _
        "WHERE CodParticipe='" & gstrCodParticipeTransferente & "' AND CodFondo='" & strCodFondoTransferente & "' AND " & _
        "CodAdministradora='" & gstrCodAdministradora & "' AND IndVigente='X' AND IndBloqueo='' " & _
        "ORDER BY FechaSuscripcion"
    
    With adoTransferente
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSql
    End With

    tdgConsulta.DataSource = adoTransferente
        
    tdgTransferente.Refresh
    
    '*** Llenamos la tabla temporal CertificadoTransferenteTmp con la lista de certificados vigentes del participe seleccionado***
    adoComm.CommandText = "INSERT INTO CertificadoTransferenteTmp " & strSql
    adoConn.Execute adoComm.CommandText
                
End Sub

Private Sub ObtenerCertificadoTransferido()

'    Dim strSql  As String
'
'    strSql = "SELECT FechaSuscripcion,ValorCuota,CantCuotas,NumSecuencial,CT.CodParticipe,CT.TipoIngreso," & _
'        "CodFondo,CT.TipoIdentidad,FechaOperacion,DescripParticipe,DescripParametro DescripFormaIngreso,NumFolio,NumOpCopia " & _
'        "FROM CertificadoTransferidoTmp CT JOIN ParticipeContrato PC ON(PC.CodParticipe=CT.CodParticipe) " & _
'        "JOIN AuxiliarParametro AP ON(AP.CodParametro=CT.TipoIngreso AND CodTipoParametro='FORING') " & _
'        "ORDER BY NumSecuencial"
'
'    With adoTransferido
'        .ConnectionString = gstrConnectConsulta
'        .RecordSource = strSql
'        .Refresh
'    End With
'
'    tdgTransferido.Refresh
                
End Sub

Public Sub ObtenerCertificadosTmp()

'    Dim strSql  As String
'
'    strSql = "SELECT FechaSuscripcion,ValorCuota,CantCuotas,NumCertificado,FechaOperacion,NumOperacion,TipoOperacion,ClaseOperacion FROM CertificadoTransferenteTmp " '& _
'        '"WHERE FechaSuscripcion ='" & Convertyyyymmdd(tdgTransferente.Columns(0)) & "' AND ValorCuota='" & CDec(tdgTransferente.Columns(1)) & "' AND CantCuotas='" & _
'        'CDec(tdgTransferente.Columns(2)) & "' AND FechaOperacion='" & Convertyyyymmdd(tdgTransferente.Columns(4)) & "' " & _
'        '"ORDER BY FechaSuscripcion"
'
'    With adoTransferente
'        .ConnectionString = gstrConnectConsulta
'        .RecordSource = strSql
'        .Refresh
'    End With
'
'    tdgTransferente.Refresh
                
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Function ValidaIngresoParticipe() As Boolean

    Dim adoRegistro As ADODB.Recordset
    
    ValidaIngresoParticipe = False
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT COUNT(NumCertificado) NumRegistros FROM ParticipeCertificado " & _
            "WHERE CodParticipe='" & gstrCodParticipeTransferido & "' AND CodFondo='" & _
            strCodFondoTransferido & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
            "IndVigente='X'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            If IsNull(adoRegistro("NumRegistros")) Then
                adoRegistro.Close: Set adoRegistro = Nothing
                Exit Function
            Else
                If adoRegistro("NumRegistros") = 0 Then
                    adoRegistro.Close: Set adoRegistro = Nothing
                    Exit Function
                End If
            End If
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    ValidaIngresoParticipe = True

End Function

Private Sub cboAgencia_Click()

    Dim strSql As String, intRegistro   As Integer
    
    strCodAgencia = Valor_Caracter
    If cboAgencia.ListIndex < 0 Then Exit Sub
    
    strCodAgencia = Trim(arrAgencia(cboAgencia.ListIndex))
    
    'strSQL = "{ call up_ACSelDatosParametro(11,'" & strCodAgencia & "') }"
    'CargarControlLista strSQL, cboPromotor, arrPromotor(), Sel_Todos
    
    'If cboPromotor.ListCount > -1 Then cboPromotor.ListIndex = 0
    'intRegistro = ObtenerItemLista(arrPromotor(), gstrCodPromotor)
    'If intRegistro >= 0 Then cboPromotor.ListIndex = intRegistro
    
End Sub

Private Sub cboEjecutivo_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodEjecutivo = ""
    If cboEjecutivo.ListIndex < 0 Then Exit Sub
    
    strCodEjecutivo = Trim(arrEjecutivo(cboEjecutivo.ListIndex))
    
    If cboFondo.ListIndex > 0 And cboTipoOperacion.ListIndex > 0 And Trim(txtNumPapeleta.Text) <> "" Then
        Call Habilita
    Else
        Call Deshabilita
    End If
    
    Set adoRegistro = New ADODB.Recordset
    
    adoComm.CommandText = "SELECT CodSucursal,CodAgencia FROM InstitucionPersona " & _
    "WHERE TipoPersona='01' AND CodPersona='" & strCodEjecutivo & "'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        strCodSucursalTransferencia = Trim(adoRegistro("CodSucursal"))
        strCodAgenciaTransferencia = Trim(adoRegistro("CodAgencia"))
    End If
    adoRegistro.Close: Set adoRegistro = Nothing
    
End Sub

Private Sub cboFondo_Click()

    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
End Sub

Private Sub cboFondoTransferente_Click()

    Dim adoRegistro As ADODB.Recordset
    Dim intRegistro As Integer
    
    strCodFondoTransferente = Valor_Caracter
    If cboFondoTransferente.ListIndex < 0 Then Exit Sub
    
    strCodFondoTransferente = Trim(arrFondoTransferente(cboFondoTransferente.ListIndex))
    
    intRegistro = ObtenerItemLista(arrFondoTransferido(), strCodFondoTransferente)
    If intRegistro >= 0 Then cboFondoTransferido.ListIndex = intRegistro
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondoTransferente & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            lblFechaSolicitud.Caption = CStr(adoRegistro("FechaCuota"))
            dblValorCuotaTransferente = CDbl(adoRegistro("ValorCuotaInicial"))
            lblValorCuota.Caption = CStr(dblValorCuotaTransferente)
            strCodMonedaFondoTransferente = Trim(adoRegistro("CodMoneda"))
                        
            gdatFechaActual = adoRegistro("FechaCuota")
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            
            Call ObtenerParametrosFondo
                                                            
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
            
End Sub

Private Sub ObtenerParametrosFondo()

    Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
                
    With adoComm
        '*** Hora de Corte ***
        .CommandText = "{ call up_ACSelDatosParametro(24,'" & strCodFondoTransferente & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strHoraCorte = adoRegistro("HoraCorte")
            strCodTipoValuacion = adoRegistro("TipoValuacion")
        End If
        adoRegistro.Close
        
        If dtpHoraSolicitud.Value > strHoraCorte Then
            blnValorConocido = False
            
            
            lblValorCuota.Caption = "0"
        Else
            blnValorConocido = True
            
            
            If strCodTipoValuacion <> Codigo_Asignacion_TMenos1 Then
                
                
            End If
        End If
        
        If strCodTipoValuacion <> Codigo_Asignacion_TMenos1 Then lblValorCuota.Caption = "0"
                            
        '*** Obtener Código de Comisión ***
        .CommandText = "SELECT CodParametro FROM AuxiliarParametro WHERE CodTipoParametro='COMFON' AND " & _
            "ValorParametro='" & strCodTipoOperacion & "'"
        Set adoRegistro = .Execute
        
        strCodComision = Valor_Caracter
        If Not adoRegistro.EOF Then
            strCodComision = Trim(adoRegistro("CodParametro"))
        End If
        adoRegistro.Close
                        
        '*** Valores Minimos y Máximos ***
        .CommandText = "{ call up_ACSelDatosParametro(25,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            dblCantCuotaMinSuscripcionInicial = CDbl(adoRegistro("CantCuotaMinSuscripcionInicial"))
            dblMontoMinSuscripcionInicial = CDbl(adoRegistro("MontoMinSuscripcionInicial"))
            dblCantMinCuotaSuscripcion = CDbl(adoRegistro("CantMinCuotaSuscripcion"))
            dblMontoMinSuscripcion = CDbl(adoRegistro("MontoMinSuscripcion"))
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub
Private Sub cboFondoTransferido_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondoTransferido = Valor_Caracter
    If cboFondoTransferido.ListIndex < 0 Then Exit Sub
    
    strCodFondoTransferido = Trim(arrFondoTransferido(cboFondoTransferido.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondoTransferido & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodMonedaFondoTransferido = Trim(adoRegistro("CodMoneda"))
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub

Private Sub cboFormaIngreso_Click()

    strCodFormaIngreso = Valor_Caracter
    If cboFormaIngreso.ListIndex < 0 Then Exit Sub
    
    strCodFormaIngreso = Trim(arrFormaIngreso(cboFormaIngreso.ListIndex))
        
End Sub

Private Sub cboPromotor_Click()

    strCodPromotor = Valor_Caracter
    If cboPromotor.ListIndex < 0 Then Exit Sub
    
    strCodPromotor = Trim(arrPromotor(cboPromotor.ListIndex))
    
End Sub

Private Sub cboSucursal_Click()

    Dim strSql As String, intRegistro   As Integer
    
    strCodSucursal = Valor_Caracter
    If cboSucursal.ListIndex < 0 Then Exit Sub
    
    strCodSucursal = Trim(arrSucursal(cboSucursal.ListIndex))
    
    strSql = "{ call up_ACSelDatosParametro(10,'" & strCodSucursal & "') }"
    CargarControlLista strSql, cboAgencia, arrAgencia(), Sel_Todos
    
    If cboAgencia.ListCount > -1 Then cboAgencia.ListIndex = 0
    intRegistro = ObtenerItemLista(arrAgencia(), gstrCodAgencia)
    If intRegistro >= 0 Then cboAgencia.ListIndex = intRegistro
    
End Sub

Private Sub cboTipoDocumentoTransferente_Click()

    strCodTipoDocumentoTransferente = Valor_Caracter
    If cboTipoDocumentoTransferente.ListIndex < 0 Then Exit Sub
    
    strCodTipoDocumentoTransferente = Trim(garrTipoDocumentoTransferente(cboTipoDocumentoTransferente.ListIndex))
    
End Sub

Private Sub cboTipoDocumentoTransferido_Click()

    strCodTipoDocumentoTransferido = Valor_Caracter
    If cboTipoDocumentoTransferido.ListIndex < 0 Then Exit Sub
    
    strCodTipoDocumentoTransferido = Trim(garrTipoDocumentoTransferido(cboTipoDocumentoTransferido.ListIndex))
    
End Sub

Private Sub cboTipoOperacion_Click()

    strCodTipoTransferencia = Valor_Caracter
    If cboTipoOperacion.ListIndex < 0 Then Exit Sub
    
    strCodTipoTransferencia = Trim(arrTipoOperacion(cboTipoOperacion.ListIndex))
    
    If strCodTipoTransferencia = "05" Then
        txtEspecificarOtro.Enabled = True
        txtEspecificarOtro.Text = ""
        Call ColorControlHabilitado(txtEspecificarOtro)
    Else
        txtEspecificarOtro.Enabled = False
        txtEspecificarOtro.Text = ""
        Call ColorControlDeshabilitado(txtEspecificarOtro)
    End If
End Sub

Private Sub cmdAgregar_Click()

    Dim dblCuotas       As Double, dblCuotasAcumulado   As Double
    Dim intRegistro     As Integer, intContador         As Integer
    Dim intNumRegistros As Integer
    Dim intRegistroT    As Integer, intContadorT        As Integer
    Dim dblCuotasTexto   As Double
         

    strCodFormaIngreso = "02"
    
    intContador = tdgTransferente.SelBookmarks.Count - 1
    
    If tdgTransferente.SelBookmarks.Count = 0 Then
        MsgBox "Seleccione el Certificado del Transferente", vbCritical, Me.Caption
        Exit Sub
    End If
        
    If CDbl(txtCuotasNuevoCertificado.Text) <= 0 Then Exit Sub
    
    If strCodParticipeTransferente <> strCodParticipeTransferido Then
        If Trim(txtNumPapeleta.Text) = Valor_Caracter Then
            MsgBox "Ingrese el Número de Papeleta de la operación", vbCritical, Me.Caption
            txtNumPapeleta.SetFocus
            Exit Sub
        End If
    Else
        txtNumPapeleta.Text = Valor_Caracter
    End If
    
    For intRegistroT = 0 To intContador
        tdgTransferente.Row = tdgTransferente.SelBookmarks(intRegistroT) - 1
        tdgTransferente.Refresh
                                         
        dblCuotas = CDbl(txtCuotasNuevoCertificado.Text)
        
        If dblCuotas > CDbl(tdgTransferente.Columns("CantCuotasPorTransferir")) Then
            MsgBox "No puede transferir una cantidad mayor a la indicada en el certificado seleccionado.", vbCritical, Me.Caption
            Exit Sub
        End If
        
        adoTransferido.AddNew
                        
        numSecuencial = numSecuencial + 1
        
        adoTransferido.Fields("CodFondoTransferido").Value = strCodFondoTransferente
        adoTransferido.Fields("CodAdministradoraTransferido").Value = gstrCodAdministradora
        adoTransferido.Fields("CodParticipeTransferido").Value = strCodParticipeTransferido
        adoTransferido.Fields("DescripParticipeTransferido").Value = lblDescripParticipeTransferido.Caption
        adoTransferido.Fields("NumSecuencial").Value = numSecuencial
        adoTransferido.Fields("FechaOperacion").Value = adoTransferente.Fields("FechaOperacion")
        adoTransferido.Fields("NumFolio").Value = Trim(txtNumPapeleta.Text)
        adoTransferido.Fields("TipoFormaIngreso").Value = strCodFormaIngreso
        adoTransferido.Fields("DescripFormaIngreso").Value = cboFormaIngreso.List(cboFormaIngreso.ListIndex)
        adoTransferido.Fields("CodFondoTransferente").Value = strCodFondoTransferente
        adoTransferido.Fields("CodAdministradoraTransferente").Value = gstrCodAdministradora
        adoTransferido.Fields("CodParticipeTransferente").Value = strCodParticipeTransferente
        adoTransferido.Fields("DescripParticipeTransferente").Value = lblDescripParticipeTransferente.Caption
        adoTransferido.Fields("NumCertificadoTransferente").Value = adoTransferente.Fields("NumCertificadoTransferente")
        adoTransferido.Fields("FechaSuscripcion").Value = adoTransferente.Fields("FechaSuscripcion")
        adoTransferido.Fields("ValorCuota").Value = adoTransferente.Fields("ValorCuota")
        adoTransferido.Fields("CantCuotas").Value = dblCuotas
        adoTransferido.Fields("NumOperacionOrigen").Value = adoTransferente.Fields("NumOperacion")
        
        adoTransferido.Update
                        
        tdgTransferido.Refresh
                        
        adoTransferente.Fields("CantCuotasPorTransferir").Value = adoTransferente.Fields("CantCuotasPorTransferir").Value - dblCuotas
                        
        dblCuotasAcumulado = dblCuotasAcumulado + dblCuotas
                        
        adoTransferente.Update

        tdgTransferente.Refresh

        cmdQuitar.Enabled = True
        
'        With adoComm
'
'            If adoTransferido.EOF And adoTransferido.BOF Then
'
'                If dblCuotas <= 0 Then Exit Sub
'
'                intRegistro = 1
'
'                .CommandText = "INSERT INTO CertificadoTransferidoTmp VALUES('" & _
'                    strCodFondoTransferido & "','" & _
'                    gstrCodParticipeTransferido & "','" & Convertyyyymmdd(tdgTransferente.Columns("FechaOperacion")) & "'," & _
'                    intRegistro & ",'" & Convertyyyymmdd(tdgTransferente.Columns("FechaSuscripcion")) & "'," & _
'                    CDec(tdgTransferente.Columns("ValorCuota")) & "," & dblCuotas & ",'" & _
'                    strCodFormaIngreso & "','" & strCodTipoDocumentoTransferido & "','" & _
'                    Trim(txtNumPapeleta.Text) & "','" & tdgTransferente.Columns("NumOperacion") & "')"
'                adoConn.Execute adoComm.CommandText
'
'                If dblCuotas = CDec(tdgTransferente.Columns("CantCuotas")) Then
'                    .CommandText = "DELETE FROM CertificadoTransferenteTmp WHERE FechaSuscripcion='" & _
'                        Convertyyyymmdd(tdgTransferente.Columns(0)) & "' AND ValorCuota='" & _
'                        CDec(tdgTransferente.Columns(1)) & "' AND CantCuotas='" & _
'                        CDec(tdgTransferente.Columns(2)) & "' AND FechaOperacion='" & _
'                        Convertyyyymmdd(tdgTransferente.Columns(4)) & "'"
'                        adoConn.Execute adoComm.CommandText
'                Else 'ya fue validado si dblCuotas es mayor por lo que a este else entrarán los que sean menor
'                    .CommandText = "UPDATE CertificadoTransferenteTmp SET CantCuotas=CantCuotas-" & _
'                        dblCuotas & " WHERE FechaSuscripcion='" & _
'                        Convertyyyymmdd(tdgTransferente.Columns(0)) & "' AND ValorCuota='" & _
'                        CDec(tdgTransferente.Columns(1)) & "' AND CantCuotas='" & _
'                        CDec(tdgTransferente.Columns(2)) & "' AND FechaOperacion='" & _
'                        Convertyyyymmdd(tdgTransferente.Columns(4)) & "'"
'                        adoConn.Execute adoComm.CommandText
'                End If
'
'
'                 Call ObtenerCertificadosTmp
'
'            Else
'
'                adoTransferido.MoveFirst
'
'                Do While Not adoTransferido.EOF
'
'                    intRegistro = CInt(adoTransferido.Fields("NumSecuencial")) + 1
'                    dblCuotasAcumulado = dblCuotasAcumulado + CDbl(adoTransferido.Fields("CantCuotas"))
'                    adoTransferido.MoveNext
'
'                Loop
'
'
'                .CommandText = "INSERT INTO CertificadoTransferidoTmp VALUES('" & _
'                    strCodFondoTransferido & "','" & _
'                    gstrCodParticipeTransferido & "','" & Convertyyyymmdd(tdgTransferente.Columns(4)) & "'," & _
'                    intRegistro & ",'" & Convertyyyymmdd(tdgTransferente.Columns(0)) & "'," & _
'                    CDec(tdgTransferente.Columns(1)) & "," & dblCuotas & ",'" & _
'                    strCodFormaIngreso & "','" & strCodTipoDocumentoTransferido & "','" & _
'                    Trim(txtNumPapeleta.Text) & "','" & tdgTransferente.Columns(7) & "')"
'                adoConn.Execute adoComm.CommandText
'
'                If dblCuotas = CDec(tdgTransferente.Columns(2)) Then
'                    .CommandText = "DELETE FROM CertificadoTransferenteTmp WHERE FechaSuscripcion='" & _
'                        Convertyyyymmdd(tdgTransferente.Columns(0)) & "' AND ValorCuota='" & _
'                        CDec(tdgTransferente.Columns(1)) & "' AND CantCuotas='" & _
'                        CDec(tdgTransferente.Columns(2)) & "' AND FechaOperacion='" & _
'                        Convertyyyymmdd(tdgTransferente.Columns(4)) & "'"
'                        adoConn.Execute adoComm.CommandText
'                Else 'menor
'                    .CommandText = "UPDATE CertificadoTransferenteTmp SET CantCuotas=CantCuotas-" & _
'                        dblCuotas & " WHERE FechaSuscripcion='" & _
'                        Convertyyyymmdd(tdgTransferente.Columns(0)) & "' AND ValorCuota='" & _
'                        CDec(tdgTransferente.Columns(1)) & "' AND CantCuotas='" & _
'                        CDec(tdgTransferente.Columns(2)) & "' AND FechaOperacion='" & _
'                        Convertyyyymmdd(tdgTransferente.Columns(4)) & "'"
'                        adoConn.Execute adoComm.CommandText
'                End If
'
'                Call ObtenerCertificadosTmp
'
'            End If
'        End With
    Next
 
    'txtCuotasNuevoCertificado.Text = "0"
    'dblCuotas = 0
    
    'If cboTipoDocumentoTransferido.ListCount > 0 Then cboTipoDocumentoTransferido.ListIndex = 0
    'txtNumDocumentoTransferido.Text = Valor_Caracter
    'txtNumPapeleta.Text = Valor_Caracter
    'gstrCodParticipeTransferido = Valor_Caracter
    'lblDescripParticipeTransferido.Caption = Valor_Caracter
                
    'Call ObtenerCertificadoTransferido
    
    'cboFormaIngreso.ListIndex = -1
    'intRegistro = ObtenerItemLista(arrFormaIngreso(), "02")
    'If intRegistro >= 0 Then cboFormaIngreso.ListIndex = intRegistro
    'cboFormaIngreso.Enabled = False
    'blnSeleccion = False
    
End Sub

Private Sub cmdBusqueda_Click(Index As Integer)

    'gstrFormulario = "frmTransferenciaParticipe"
    'Select Case index
     '   Case 0
      '      frmTransferenciaParticipe.Tag = 0
       ' Case 1
           ' frmTransferenciaParticipe.Tag = 1
    'End Select
    
    'frmBusquedaParticipeP.Show vbModal
    
    
    'gstrFormulario = Me.Name
    'frmBusquedaCliente.Show vbModal
    
    
'    gstrFormulario = "frmTransferenciaParticipe"
'    frmBusquedaParticipeP.Show vbModal
'
'    If gstrCodParticipe <> Valor_Caracter Then
'        strCodParticipeBusqueda = gstrCodParticipeTransferente
'    End If
    
'''''''''''''''NUEVO
    
    Dim sSql As String
    Dim intRegistro As Integer
   
    Screen.MousePointer = vbHourglass
   
    Dim frmBus As frmBuscar
    
    Set frmBus = New frmBuscar
    
    With frmBus.TBuscarRegistro1
           
        .ADOConexion = adoConn
        .ADOConexion.CommandTimeout = 0

        .iTipoGrilla = 2
        
        frmBus.Caption = " Relación de Participes"
        .sSql = "{ call up_ACSelDatos(30) }"
        
        .OutputColumns = "1,2,3,4,5,6,7,8,9,10,11"
        .HiddenColumns = "1,2,5,6,7,10,11"
        
        .BuscarTabla
        
        Screen.MousePointer = vbNormal
        frmBus.Show 1
       
        If .iParams.Count = 0 Then Exit Sub
        
        If .iParams(1).Valor <> "" Then
            
            If Index = 0 Then
                
                strCodParticipeTransferente = Trim(.iParams(1).Valor)
                
                strCodTipoDocumentoTransferente = Trim(.iParams(6).Valor)
                strTipoDocumentoTransferente = Trim(.iParams(2).Valor)
                
                strNumDocumentoTransferente = Trim(.iParams(3).Valor)
                
                cboTipoDocumentoTransferente.ListIndex = ObtenerItemLista(garrTipoDocumentoTransferente(), strCodTipoDocumentoTransferente)
                txtNumDocumentoTransferente.Text = strNumDocumentoTransferente
    
                lblDescripParticipeTransferente.Caption = Trim(.iParams(4).Valor)
                'lblDescripTipoParticipe.Caption = Trim(.iParams(7).Valor)
                
                strDescripTitularSolicitanteTransferente = Trim(.iParams(8).Valor)
                strCodClienteTransferente = Trim(.iParams(9).Valor)
                
                'Cargar certificados
                Call CargarDetalleGrillaTransferente
                'Call CargarCertificadosTransferente
            
            End If
            
            If Index = 1 Then
                
                If Trim(.iParams(1).Valor) <> strCodParticipeTransferido Then
                    If strCodParticipeTransferido <> Valor_Caracter Then
                        If MsgBox("Ha seleccionado un participe transferido diferente al actual. Desea iniciar una nueva transferencia? Los datos de la transferencia actual se eliminarán.", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                            Call CargarDetalleGrillaTransferido
                        End If
                    Else
                        Call CargarDetalleGrillaTransferido
                    End If
                End If
                
                strCodParticipeTransferido = Trim(.iParams(1).Valor)
                
                strCodTipoDocumentoTransferido = Trim(.iParams(6).Valor)
                strTipoDocumentoTransferido = Trim(.iParams(2).Valor)
                
                strNumDocumentoTransferido = Trim(.iParams(3).Valor)
                
                cboTipoDocumentoTransferido.ListIndex = ObtenerItemLista(garrTipoDocumentoTransferente(), strCodTipoDocumentoTransferido)
                txtNumDocumentoTransferido.Text = strNumDocumentoTransferido
    
                lblDescripParticipeTransferido.Caption = Trim(.iParams(4).Valor)
                'lblDescripTipoParticipe.Caption = Trim(.iParams(7).Valor)
                
                strDescripTitularSolicitanteTransferido = Trim(.iParams(8).Valor)
                strCodClienteTransferido = Trim(.iParams(9).Valor)
            
            End If
        
        Else
        
            If Index = 0 Then
            
                strCodParticipeTransferente = Valor_Caracter
                strCodTipoDocumentoTransferente = Valor_Caracter
                strTipoDocumentoTransferente = Valor_Caracter
                strNumDocumentoTransferente = Valor_Caracter
                
                cboTipoDocumentoTransferente.ListIndex = ObtenerItemLista(garrTipoDocumentoTransferente(), strCodTipoDocumentoTransferente)
                txtNumDocumentoTransferente.Text = strNumDocumentoTransferente
    
                lblDescripParticipeTransferente.Caption = Valor_Caracter
                
                strDescripTitularSolicitanteTransferente = Valor_Caracter
                strCodClienteTransferente = Valor_Caracter
            
            End If
        
            If Index = 1 Then
            
                strCodParticipeTransferido = Valor_Caracter
                strCodTipoDocumentoTransferido = Valor_Caracter
                strTipoDocumentoTransferido = Valor_Caracter
                strNumDocumentoTransferido = Valor_Caracter
                
                cboTipoDocumentoTransferido.ListIndex = ObtenerItemLista(garrTipoDocumentoTransferido(), strCodTipoDocumentoTransferido)
                txtNumDocumentoTransferido.Text = strNumDocumentoTransferido
    
                lblDescripParticipeTransferido.Caption = Valor_Caracter
                
                strDescripTitularSolicitanteTransferido = Valor_Caracter
                strCodClienteTransferido = Valor_Caracter
            
                Call CargarDetalleGrillaTransferido
            
            End If
        
        
        End If
            
       
    End With
    
    Set frmBus = Nothing
    
    
End Sub
Private Sub CargarCertificadosTransferente()

    Dim strSql  As String
        
    Set adoTransferente = New ADODB.Recordset
        
    strSql = "SELECT FechaSuscripcion,ValorCuota,CantCuotas,NumCertificado,FechaOperacion,NumOperacion,TipoOperacion,ClaseOperacion FROM ParticipeCertificado " & _
        "WHERE CodParticipe='" & strCodParticipeTransferente & "' AND CodFondo='" & strCodFondoTransferente & "' AND " & _
        "CodAdministradora='" & gstrCodAdministradora & "' AND IndVigente='X' AND IndBloqueo='' " & _
        "ORDER BY FechaSuscripcion"
    
    With adoTransferente
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSql
    End With

    tdgTransferente.DataSource = adoTransferente
        
    tdgTransferente.Refresh

End Sub
Private Sub cmdQuitar_Click()
    
'    Dim intRegistro     As Integer, intContador         As Integer
'    Dim strFechaDesde   As String, strFechaHasta        As String
'
'    intContador = tdgTransferido.SelBookmarks.Count - 1
'
'    For intRegistro = 0 To intContador
'        tdgTransferido.Row = tdgTransferido.SelBookmarks(intRegistro) - 1
'        tdgTransferido.Refresh
'
'        strFechaDesde = Convertyyyymmdd(CVDate(tdgTransferido.Columns(7)))
'        strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, CVDate(tdgTransferido.Columns(7))))
'
'        adoComm.CommandText = "DELETE CertificadoTransferidoTmp WHERE " & _
'            "(FechaOperacion>='" & strFechaDesde & "' AND FechaOperacion<'" & strFechaHasta & "') AND " & _
'            "ValorCuota=" & CDec(tdgTransferido.Columns(2)) & " AND CantCuotas=" & _
'            CDec(tdgTransferido.Columns(3)) & " AND NumSecuencial=" & CInt(tdgTransferido.Columns(10)) & " AND " & _
'            "CodParticipe='" & Trim(tdgTransferido.Columns(5)) & "'"
'        adoConn.Execute adoComm.CommandText
'
'        adoComm.CommandText = "INSERT INTO CertificadoTransferenteTmp VALUES ('" & _
'            Convertyyyymmdd(CVDate(tdgTransferido.Columns(4))) & "','" & _
'            tdgTransferido.Columns(5) & "','" & tdgTransferido.Columns(6) & "','','" & _
'            Convertyyyymmdd(CVDate(tdgTransferido.Columns(2))) & "','','','" & "')"
'        adoConn.Execute adoComm.CommandText
'
'    Next
'
'    Call ObtenerCertificadoTransferido
    
    
    Dim dblBookmark As Double
    Dim numSecuencialActual As Integer

    If adoTransferido.RecordCount > 0 Then
    
        dblBookmark = adoTransferido.Bookmark
    
        numSecuencialActual = adoTransferido.Fields("NumSecuencial").Value
    
        If numSecuencial <= numSecuencialActual Then
            If adoTransferido.RecordCount > 1 Then
                numSecuencial = numSecuencial - 1
            Else
                numSecuencial = 0
            End If
        End If
        
        adoTransferente.Fields("CantCuotasPorTransferir").Value = adoTransferente.Fields("CantCuotasPorTransferir").Value + adoTransferido.Fields("CantCuotas").Value
        
        'dblCuotasAcumulado = dblCuotasAcumulado - adoTransferido.Fields("CantCuotas").Value
        
        adoTransferente.Update

        tdgTransferente.Refresh
        
        adoTransferido.Delete adAffectCurrent
        
        If adoTransferido.EOF Then
            adoTransferido.MovePrevious
            tdgTransferido.MovePrevious
        End If
            
        adoTransferido.Update
        
        If adoTransferido.RecordCount = 0 Then cmdQuitar.Enabled = False

        If adoTransferido.RecordCount > 0 And Not adoTransferido.BOF And Not adoTransferido.EOF And numSecuencial <= numSecuencialActual Then adoTransferido.Bookmark = dblBookmark - 1

        If adoTransferido.RecordCount > 0 And Not adoTransferido.BOF And Not adoTransferido.EOF And numSecuencial > numSecuencialActual Then adoTransferido.Bookmark = dblBookmark + 1
   
        tdgTransferido.Refresh
    
    End If
    
    
    
End Sub

Private Sub Form_Activate()

    Call CargarReportes
    
End Sub

Private Sub Form_Deactivate()

    ReDim garrTipoDocumento(0)
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
    
'    Dim adoRecord As New Recordset
'
'    gstrLlamaSoli = ""
'
'    Call LimpiaScr
'    gstrflgbusqtrans = "F"
'    tim_horSoli.Text = Format(Time, "HH:MM:SS")
'
'    strSentencia = "Sp_INF_OpeTran '16'"
'    Call LCmbLoad(strSentencia, Cmb_Cod_fond, aMapCod_Fond(), "")
'    Cmb_Cod_fond.ListIndex = -1
'
'    strSentencia = "Sp_INF_OpeTran '17'"
'    Call LCmbLoad(strSentencia, Cmb_Cod_prom, aMapCod_Prom(), "")
'
'    If gstrCodProm <> "" Then
'        Cmb_Cod_prom.ListIndex = LBsqIteArr(aMapCod_Prom(), gstrCodProm)
'    End If
'
'    ReDim aHeaDet(5)
'    ReDim aWidDet(5)
'    aHeaDet(1) = "Cnt. Cuotas": aHeaDet(2) = "Custodia": aHeaDet(3) = "Partícipe": aHeaDet(4) = "CodPar": aHeaDet(4) = "IdxPar"
'    aWidDet(1) = 1000: aWidDet(2) = 390: aWidDet(3) = 2000: aWidDet(4) = 5: aWidDet(5) = 5
'
'    If Trim(gstrLlamaSoli) <> "" Then
'        adoComm.CommandText = "Sp_INF_OpeTran '18'"
'        Set adoRecord = adoComm.Execute
'        If Not adoRecord.EOF Then
'            Txt_CodUnico.Text = adoRecord!COD_UNICO
'        End If
'        adoRecord.Close: Set adoRecord = Nothing
'        DoEvents
'    End If
'
'   '*** Extrae el Tipo de Cambio de la Tabla FMCUOTAS ***
'   adoComm.CommandText = "Sp_INF_OpeTran '19', '" & gstrFechaAct & "'"
'   Set adoRecord = adoComm.Execute
'   dblTipCamb = Format(adoRecord!VAL_TCMB, "0.0000")
'   TipCambio.Text = dblTipCamb
'   adoRecord.Close: Set adoRecord = Nothing
'
'
'    '*** Configurar la grilla
'    ReDim aGrdCnf(1 To 19)
'
'    aGrdCnf(1).TitDes = "Folio"
'    aGrdCnf(1).DatNom = "NRO_FOLI"
'    aGrdCnf(1).DatAnc = 130 * 9
'
'    aGrdCnf(2).TitDes = "Tip.Soli."
'    aGrdCnf(2).DatNom = "TIP_SOLI"
'    aGrdCnf(2).DatAnc = 130 * 5
'
'    aGrdCnf(3).TitDes = "Partícipe"
'    aGrdCnf(3).DatNom = "DSC_PAR2"
'    aGrdCnf(3).DatAnc = 130 * 20
'
'    aGrdCnf(4).TitDes = "Certificado"
'    aGrdCnf(4).DatNom = "NRO_CERT"
'    aGrdCnf(4).DatAnc = 130 * 7
'
'    aGrdCnf(5).TitDes = "Tasa"
'    aGrdCnf(5).DatNom = "TAS_OPER"
'    aGrdCnf(5).DatAnc = 130 * 4
'    aGrdCnf(5).DatJus = 1
'    aGrdCnf(5).DatFmt = "D"
'
'    aGrdCnf(6).TitDes = "Cuotas"
'    aGrdCnf(6).DatNom = "CNT_CUOT"
'    aGrdCnf(6).DatAnc = 130 * 10
'    aGrdCnf(6).DatJus = 1
'    aGrdCnf(6).DatFmt = "C"
'
'    aGrdCnf(7).TitDes = "Monto"
'    aGrdCnf(7).DatNom = "MTO_SUSC"
'    aGrdCnf(7).DatAnc = 130 * 10
'    aGrdCnf(7).DatFmt = "D"
'    aGrdCnf(7).DatJus = 1
'
'    aGrdCnf(8).TitDes = "Cobro/Pago - Tipo Cta. - Nro.Cta."
'    aGrdCnf(8).DatNom = "TIP_PAGO"
'    aGrdCnf(8).DatJus = vbLeftJustify
'    aGrdCnf(8).DatAnc = 130 * 20
'
'    aGrdCnf(9).TitDes = "Custodia"
'    aGrdCnf(9).DatNom = "FLG_CUST"
'    aGrdCnf(9).DatAnc = 130 * 6
'    aGrdCnf(9).DatJus = 2
'
'    aGrdCnf(10).TitDes = "Liquidado"
'    aGrdCnf(10).DatNom = "FLG_CONF"
'    aGrdCnf(10).DatAnc = 130 * 6
'    aGrdCnf(10).DatJus = 2
'
'    aGrdCnf(11).TitDes = "Fecha"
'    aGrdCnf(11).DatNom = "FCH_SOLI"
'    aGrdCnf(11).DatAnc = 130 * 10
'    aGrdCnf(11).DatFmt = "F"
'    aGrdCnf(11).DatJus = 2
'
'    aGrdCnf(12).TitDes = "Hora"
'    aGrdCnf(12).DatNom = "HOR_SOLI"
'    aGrdCnf(12).DatAnc = 130 * 10
'    aGrdCnf(12).DatJus = 2
'
'    aGrdCnf(13).TitDes = "Cheque"
'    aGrdCnf(13).DatNom = "NRO_CHEQ"
'    aGrdCnf(13).DatAnc = 130 * 15
'
'    aGrdCnf(14).TitDes = "Retención"
'    aGrdCnf(14).DatNom = "FCH_FRET"
'    aGrdCnf(14).DatAnc = 130 * 10
'    aGrdCnf(14).DatFmt = "F"
'
'    aGrdCnf(15).TitDes = "Nro.Soli."
'    aGrdCnf(15).DatNom = "NRO_SOLI"
'    aGrdCnf(15).DatAnc = 130 * 6
'
'    aGrdCnf(16).TitDes = "Estado"
'    aGrdCnf(16).DatNom = "FLG_SOLI"
'    aGrdCnf(16).DatAnc = 130 * 6
'    aGrdCnf(16).DatJus = 2
'
'    aGrdCnf(17).TitDes = "Agencia Operador"
'    aGrdCnf(17).DatNom = "DSC_AGE_OPER"
'    aGrdCnf(17).DatAnc = 130 * 25
'
'    aGrdCnf(18).TitDes = "Operador"
'    aGrdCnf(18).DatNom = "DSC_PERS"
'    aGrdCnf(18).DatAnc = 130 * 25
'
'    aGrdCnf(19).TitDes = "Agencia Sectorización"
'    aGrdCnf(19).DatNom = "DSC_AGE_SECT"
'    aGrdCnf(19).DatAnc = 130 * 25

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
Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Lista de Transferencias"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Imprimir Solicitud de Transferencia"
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Text = "Imprimir Formato de Transferencia"
    
End Sub
Private Sub CargarListas()

    Dim intRegistro As Integer
    Dim strSql      As String
    
    '*** Fondos ***
    strSql = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSql, cboFondo, arrFondo(), Valor_Caracter
    CargarControlLista strSql, cboFondoTransferente, arrFondoTransferente(), Sel_Defecto
    CargarControlLista strSql, cboFondoTransferido, arrFondoTransferido(), Sel_Defecto
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    '*** Tipo Solicitud/Operación ***
    strSql = "SELECT CodTipoSolicitud CODIGO,DescripTipoSolicitud DESCRIP FROM TipoSolicitud WHERE CodCorto='M' or CodCorto='O' ORDER BY DescripTipoSolicitud"
    CargarControlLista strSql, cboTipoOperacion, arrTipoOperacion(), Sel_Todos
                                
    '*** Sucursal ***
    strSql = "{ call up_ACSelDatos(15) }"
    CargarControlLista strSql, cboSucursal, arrSucursal(), Sel_Todos
    
    If cboSucursal.ListCount > 0 Then cboSucursal.ListIndex = 0
    intRegistro = ObtenerItemLista(arrSucursal(), gstrCodSucursal)
    If intRegistro >= 0 Then cboSucursal.ListIndex = 0
            
    '*** Operador ***
    'strSQL = "SELECT CodPersona CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='01' ORDER BY DescripPersona"
    'CargarControlLista strSQL, cboEjecutivo, arrEjecutivo(), ""
    
    strSql = "{ call up_ACSelDatos(41) }"
    CargarControlLista strSql, cboPromotor, arrPromotor(), Sel_Todos
    
    If cboPromotor.ListCount > 0 Then cboPromotor.ListIndex = 0
    intRegistro = ObtenerItemLista(arrPromotor(), gstrCodPromotor)
    If intRegistro >= 0 Then cboPromotor.ListIndex = 0
    'If intRegistro >= 0 Then cboPromotor.ListIndex = intRegistro
    
    
    
    '*** Tipo Documento Identidad ***
    strSql = "{ call up_ACSelDatos(11) }"
    CargarControlLista strSql, cboTipoDocumentoTransferente, garrTipoDocumentoTransferente(), Sel_Defecto
    CargarControlLista strSql, cboTipoDocumentoTransferido, garrTipoDocumentoTransferido(), Sel_Defecto
    
    '*** Tipo Forma Ingreso Calidad Partícipe ***
    strSql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='FORING' ORDER BY DescripParametro"
    CargarControlLista strSql, cboFormaIngreso, arrFormaIngreso(), Sel_Defecto
    
End Sub
Private Sub InicializarValores()

    Dim adoRegistro As ADODB.Recordset
    Dim intCont     As Integer
    
    strEstado = Reg_Defecto
    blnValorConocido = False
    tabTransferencia.Tab = 0

    dtpFechaDesde.Value = gdatFechaActual
    dtpFechaHasta.Value = gdatFechaActual
    
    Set adoRegistro = New ADODB.Recordset
    
    intCont = 0
    ReDim arrLeyendaTransferencia(intCont)
    
    With adoComm
        '*** Transferencia ***
        .CommandText = "SELECT RTRIM(TSD.CodCorto) + space(1) + '=' + space(1) + RTRIM(TS.DescripTipoSolicitud) + space(1) + RTRIM(TSD.DescripDetalleTipoSolicitud)  ValorLeyenda " & _
            "FROM TipoSolicitud TS JOIN TipoSolicitudDetalle TSD " & _
            "ON(TSD.CodTipoSolicitud=TS.CodTipoSolicitud AND " & _
            "TS.CodTipoSolicitud='" & Codigo_Operacion_Transferencia & "') " & _
            "WHERE TS.CodCorto='T'"
        Set adoRegistro = .Execute
        
        Do Until adoRegistro.EOF
            ReDim Preserve arrLeyendaTransferencia(intCont)
            arrLeyendaTransferencia(intCont) = adoRegistro("ValorLeyenda")
            
            adoRegistro.MoveNext
            intCont = intCont + 1
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
            
    End With
    
    '*** Verificando Nivel de Acceso de Usuario ***
'    strNivAcceso = AccesoForm(gstrNomOpc, gstrNumInd)

    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set frmTransferenciaParticipe = Nothing
    gstrCodParticipeTransferente = Valor_Caracter
    gstrCodParticipeTransferido = Valor_Caracter
    tmrHora.Enabled = False
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción..."
    
End Sub

Private Sub lblDescrip_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim intContador As Integer
    
    If Index = 21 Then
        intContador = UBound(arrLeyendaTransferencia)
        
        lstLeyenda.AddItem "Leyenda :"
        lstLeyenda.AddItem ""
        
        For intContador = 0 To (UBound(arrLeyendaTransferencia))
            lstLeyenda.AddItem arrLeyendaTransferencia(intContador)
        Next
        
        lstLeyenda.Height = lblDescrip(Index).Height * (intContador + 2)
        lstLeyenda.Left = lblDescrip(Index).Left
        lstLeyenda.Top = lblDescrip(Index).Top + lblDescrip(Index).Height
        lstLeyenda.Width = 3300
        lstLeyenda.Visible = True
    End If
    
    If Index = 28 Then
        intContador = UBound(arrLeyendaTransferencia)
        
        lstLeyenda.AddItem "Leyenda :"
        lstLeyenda.AddItem ""
        
        For intContador = 0 To (UBound(arrLeyendaTransferencia))
            lstLeyenda.AddItem arrLeyendaTransferencia(intContador)
        Next
        
        lstLeyenda.Height = lblDescrip(Index).Height * (intContador + 2)
        lstLeyenda.Left = lblDescrip(Index).Left
        lstLeyenda.Top = lblDescrip(Index).Top + lblDescrip(Index).Height
        lstLeyenda.Width = 3800
        lstLeyenda.Visible = True
    End If
End Sub

Private Sub lblDescrip_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Index = 21 Or Index = 28 Then
        lstLeyenda.Clear
        lstLeyenda.Visible = False
    End If
    
End Sub

Private Sub lblValorCuota_Change()

    Call FormatoMillarEtiqueta(lblValorCuota, Decimales_CantCuota)
    
End Sub

Private Sub tabTransferencia_Click(PreviousTab As Integer)

    Select Case tabTransferencia.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabTransferencia.Tab = 0
        
    End Select
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 4 Then
        Call DarFormatoValor(Value, Decimales_ValorCuota)
    End If
            
    If ColIndex = 5 Then
        Call DarFormatoValor(Value, Decimales_CantCuota)
    End If
    
End Sub

Private Sub tdgTransferente_Click()

    'blnSeleccion = True
    
End Sub

Private Sub tdgTransferente_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 2 Then
        Call DarFormatoValor(Value, Decimales_ValorCuota)
    End If
            
    If ColIndex = 3 Or ColIndex = 4 Then
        Call DarFormatoValor(Value, Decimales_CantCuota)
    End If
    
End Sub

Private Sub tdgTransferente_SelChange(Cancel As Integer)

'    Dim dblCuotas   As Double, dblCuotasAcumulado   As Double
'    Dim intRegistro As Integer, intContador         As Integer
    
    'If blnSeleccion Then
'    intContador = tdgTransferente.SelBookmarks.Count - 1
    Dim intRegistro As Integer, intContador        As Integer
    Static strFechaSuscripcion As String
    Dim blnSeleccion As Boolean, numRowInicial As Long
    
    intContador = tdgTransferente.SelBookmarks.Count - 1
    
    If tdgTransferente.SelBookmarks.Count = 0 Then
        blnSeleccion = False
        strFechaSuscripcion = Valor_Caracter
    End If
    
    If tdgTransferente.SelBookmarks.Count = 1 Then
        blnSeleccion = False
        strFechaSuscripcion = Convertyyyymmdd(adoTransferente.Fields("FechaSuscripcion").Value)
    End If
    
    If tdgTransferente.SelBookmarks.Count > 1 Then
        blnSeleccion = True
    End If

    If blnSeleccion Then
        numRowInicial = tdgTransferente.Row
        For intRegistro = 0 To intContador
            tdgTransferente.Row = tdgTransferente.SelBookmarks(intRegistro) - 1
            tdgTransferente.Refresh
            If strFechaSuscripcion <> Convertyyyymmdd(tdgTransferente.Columns("FechaSuscripcion").Value) Then
                MsgBox "Sólo se puede transferir más de un certificado a la vez si estos tienen la misma fecha de suscripción; de lo contrario sólo se podrá transferir uno cada vez!", vbCritical, Me.Caption
                tdgTransferente.Row = numRowInicial
                Cancel = True
                Exit Sub
            End If
        Next
    End If
                
                
                
                
                
'        If gstrCodParticipeTransferente = gstrCodParticipeTransferido Then
'            If tdgTransferente.SelBookmarks.Count > 1 Then
'                MsgBox "Solo se puede seleccionar un certificado", vbCritical, Me.Caption
'                tdgTransferente.ReBind
'                Exit Sub
'            End If
'        End If
    
'    For intRegistro = 0 To intContador
'        tdgTransferente.Row = tdgTransferente.SelBookmarks(intRegistro) - 1
'        tdgTransferente.Refresh
'
'        dblCuotas = CDbl(tdgTransferente.Columns(2))
'
'        dblCuotasAcumulado = dblCuotasAcumulado + dblCuotas
'
'    Next
    'End If
    
End Sub

Private Sub tdgTransferido_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 2 Then
        Call DarFormatoValor(Value, Decimales_ValorCuota)
    End If
            
    If ColIndex = 3 Then
        Call DarFormatoValor(Value, Decimales_CantCuota)
    End If
    
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

Private Sub tmrHora_Timer()

    dtpHoraSolicitud.Value = ObtenerHoraServidor
    
End Sub

Private Sub Calcular()

    Dim dblMonto            As Double, dblCuota     As Double
    Dim dblMontoComision    As Double, dblMontoIgv  As Double
    Dim dblComision         As Double
    
    dblComision = ObtenerComisionParticipacion(strCodComision, strCodFondo, gstrCodAdministradora)
    
'    If blnCuota Then '*** CalcularMonto ***
'        If blnValorConocido Then
'            dblMonto = CDbl(txtCuotas.Text) * CDbl(lblValorCuota.Caption)
'            txtMonto.Text = CStr(dblMonto)
'        Else
'            txtMonto.Text = "0"
'        End If
'    Else '*** Calcular Cantidad de Cuotas ***
'        If blnValorConocido Then
'            dblMonto = CDbl(txtMonto.Text)
'            dblCuota = dblMonto / CDbl(lblValorCuota)
'            txtCuotas.Text = CStr(dblCuota)
'            If strCodTipoValuacion <> Codigo_Asignacion_TMenos1 Then txtCuotas.Text = "0"
'        Else
'            txtCuotas.Text = "0"
'        End If
'    End If
        
End Sub

Private Sub txtCuotasNuevoCertificado_Change()

    Call FormatoCajaTexto(txtCuotasNuevoCertificado, Decimales_CantCuota)
    
End Sub

Private Sub txtCuotasNuevoCertificado_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtCuotasNuevoCertificado, Decimales_CantCuota)
    
End Sub

Private Sub txtNumDocumentoTransferente_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call ObtenerDatosParticipe(0)
        Call ObtenerCertificados
    End If
    
End Sub

Private Sub ObtenerDatosParticipe(Index As Integer)

    Dim adoRegistro As ADODB.Recordset

    Set adoRegistro = New ADODB.Recordset
    adoRegistro.CursorLocation = adUseClient
    adoRegistro.CursorType = adOpenStatic

    If Index = 0 Then
        adoComm.CommandText = "SELECT PC.CodParticipe,AP1.DescripParametro TipoIdentidad,PCD.NumIdentidad,DescripParticipe,FechaIngreso,PCD.TipoIdentidad CodIdentidad,PC.TipoMancomuno, AP2.DescripParametro DescripMancomuno " & _
        "FROM ParticipeContratoDetalle PCD JOIN ParticipeContrato PC " & _
        "ON(PCD.CodParticipe=PC.CodParticipe AND PCD.TipoIdentidad='" & strCodTipoDocumentoTransferente & "' AND PCD.NumIdentidad='" & Trim(txtNumDocumentoTransferente.Text) & "') " & _
        "JOIN AuxiliarParametro AP1 ON(AP1.CodParametro=PCD.TipoIdentidad AND CodTipoParametro='TIPIDE') " & _
        "JOIN AuxiliarParametro AP2 ON(AP2.CodParametro=PC.TipoMancomuno AND AP2.CodTipoParametro='TIPMAN')"
        adoRegistro.Open adoComm.CommandText, adoConn
    
        If Not adoRegistro.EOF Then
            If adoRegistro.RecordCount > 1 Then
                gstrFormulario = "frmTransferenciaParticipe"
                frmTransferenciaParticipe.Tag = 0
                frmBusquedaParticipeP.optCriterio(1).Value = vbChecked
                frmBusquedaParticipeP.txtNumDocumento = Trim(txtNumDocumentoTransferente.Text)
                Call frmBusquedaParticipeP.Buscar
                frmBusquedaParticipeP.Show vbModal
            Else
                gstrCodParticipeTransferente = Trim(adoRegistro("CodParticipe"))
                lblDescripTipoParticipeTransferente.Caption = Trim(adoRegistro("DescripMancomuno"))
                lblDescripParticipeTransferente.Caption = Trim(adoRegistro("DescripParticipe"))
            End If
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    Else
        adoComm.CommandText = "SELECT PC.CodParticipe,AP1.DescripParametro TipoIdentidad,PCD.NumIdentidad,DescripParticipe,FechaIngreso,PCD.TipoIdentidad CodIdentidad,PC.TipoMancomuno, AP2.DescripParametro DescripMancomuno " & _
        "FROM ParticipeContratoDetalle PCD JOIN ParticipeContrato PC " & _
        "ON(PCD.CodParticipe=PC.CodParticipe AND PCD.TipoIdentidad='" & strCodTipoDocumentoTransferido & "' AND PCD.NumIdentidad='" & Trim(txtNumDocumentoTransferido.Text) & "') " & _
        "JOIN AuxiliarParametro AP1 ON(AP1.CodParametro=PCD.TipoIdentidad AND CodTipoParametro='TIPIDE') " & _
        "JOIN AuxiliarParametro AP2 ON(AP2.CodParametro=PC.TipoMancomuno AND AP2.CodTipoParametro='TIPMAN')"
        adoRegistro.Open adoComm.CommandText, adoConn
    
        If Not adoRegistro.EOF Then
            If adoRegistro.RecordCount > 1 Then
                gstrFormulario = "frmTransferenciaParticipe"
                frmTransferenciaParticipe.Tag = 1
                frmBusquedaParticipeP.optCriterio(1).Value = vbChecked
                frmBusquedaParticipeP.txtNumDocumento = Trim(txtNumDocumentoTransferido.Text)
                Call frmBusquedaParticipeP.Buscar
                frmBusquedaParticipeP.Show vbModal
            Else
                gstrCodParticipeTransferido = Trim(adoRegistro("CodParticipe"))
                lblDescripParticipeTransferido.Caption = Trim(adoRegistro("DescripParticipe"))
            End If
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End If

End Sub

Private Sub txtNumDocumentoTransferido_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call ObtenerDatosParticipe(1)
    End If
    
End Sub


Private Sub txtNumPapeleta_LostFocus()

    txtNumPapeleta.Text = Format(txtNumPapeleta.Text, "000000000000000")
    
    If cboFondo.ListIndex > 0 And cboEjecutivo.ListIndex > -1 And cboTipoOperacion.ListIndex > 0 Then
        Call Habilita
    Else
        Call Deshabilita
    End If
    
    If strEstado = Reg_Adicion Then
        If gstrCodParticipeTransferente <> gstrCodParticipeTransferido Then
            If Not ValidarNumFolio(Trim(txtNumPapeleta.Text), strCodTipoOperacion, strCodFondo, Me) Then Exit Sub
            
            If adoTransferido.RecordCount > 0 Then
                adoTransferido.MoveFirst
                
                Do While Not adoTransferido.EOF
                    
                    If Trim(txtNumPapeleta.Text) = Trim(adoTransferido.Fields("NumFolio")) Then
                        MsgBox "El Número de Papeleta ya ha sido utilizado.", vbCritical, gstrNombreEmpresa
                        txtNumPapeleta.SetFocus
                        txtNumPapeleta.SelStart = 0
                        txtNumPapeleta.SelLength = Len(txtNumPapeleta.Text)
                    
                        Exit Do
                    End If
                    
                    adoTransferido.MoveNext
                Loop
            End If
        End If
    End If
    
End Sub

Private Sub Habilita()
   
End Sub
Private Sub CargarDetalleGrillaTransferente()
    
    Dim adoRegistro As ADODB.Recordset
    Dim adoField As ADODB.Field
    Dim numSecuencial As Long
    Dim strSql As String

    Call ConfiguraDetalleGrillaTransferente
    
    Set adoRegistro = New ADODB.Recordset
    
    strSql = "SELECT CodParticipe as CodParticipeTransferente, FechaSuscripcion,ValorCuota,CantCuotas,CantCuotas as CantCuotasPorTransferir, NumCertificado as NumCertificadoTransferente,FechaOperacion,NumOperacion,TipoOperacion,ClaseOperacion FROM ParticipeCertificado " & _
        "WHERE CodParticipe='" & strCodParticipeTransferente & "' AND CodFondo='" & strCodFondoTransferente & "' AND " & _
        "CodAdministradora='" & gstrCodAdministradora & "' AND IndVigente='X' AND IndBloqueo='' " & _
        "ORDER BY FechaSuscripcion"
    
    With adoRegistro
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSql
    
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                adoTransferente.AddNew
                numSecuencial = numSecuencial + 1
                For Each adoField In adoTransferente.Fields
                    adoTransferente.Fields(adoField.Name) = adoRegistro.Fields(adoField.Name)
                Next
                adoTransferente.Update
                adoRegistro.MoveNext
            Loop
            adoTransferente.MoveFirst
        
            'Call CargarDetalleGrillaRegistro
        
        End If
    End With
    
    tdgTransferente.DataSource = adoTransferente
            
    tdgTransferente.Refresh
    
End Sub
Private Sub CargarDetalleGrillaTransferido()
    
    Dim adoRegistro As ADODB.Recordset
    Dim adoField As ADODB.Field
    'Dim numSecuencial As Long
    Dim strSql As String

    Call ConfiguraDetalleGrillaTransferido
    
'    Set adoRegistro = New ADODB.Recordset
'
'    strSQL = "SELECT '' AS CodFondoTransferido, '' AS CodAdministradoraTransferido, '' AS CodParticipeTransferido, " & _
'             "'' AS DescripParticipeTransferido, 0 AS NumSecuencial, '19000101' AS FechaOperacion, '' AS NumFolio, '' AS TipoFormaIngreso, '' AS DescripFormaIngreso, " & _
'             "'' AS CodFondoTransferente, '' AS CodAdministradoraTransferente, '' AS CodParticipeTransferente, '' AS DescripParticipeTransferente, " & _
'             "'' AS NumCertificadoTransferente, '19000101' AS FechaSuscripcion, 0.00 AS ValorCuota, 0.00 AS CantCuotas, '' AS NumOperacionOrigen " & _
'             "WHERE 1=2"
'
'    With adoRegistro
'        .ActiveConnection = gstrConnectConsulta
'        .CursorLocation = adUseClient
'        .CursorType = adOpenStatic
'        .LockType = adLockBatchOptimistic
'        .Open strSQL
'
'        If .RecordCount > 0 Then
'            .MoveFirst
'            Do While Not .EOF
'                adoTransferido.AddNew
'                numSecuencial = numSecuencial + 1
'                For Each adoField In adoTransferido.Fields
'                    adoTransferido.Fields(adoField.Name) = adoRegistro.Fields(adoField.Name)
'                Next
'                adoTransferido.Update
'                adoRegistro.MoveNext
'            Loop
'            adoTransferido.MoveFirst
'        End If
'    End With
    
    tdgTransferido.DataSource = adoTransferido
            
    tdgTransferido.Refresh
    
End Sub

Private Sub ConfiguraDetalleGrillaTransferente()

    Set adoTransferente = New ADODB.Recordset
    
    With adoTransferente
        .CursorLocation = adUseClient
        .Fields.Append "CodParticipeTransferente", adVarChar, 20
        .Fields.Append "FechaSuscripcion", adDate, 10
        .Fields.Append "ValorCuota", adDecimal
        .Fields.Item("ValorCuota").Precision = 19
        .Fields.Item("ValorCuota").NumericScale = 5
        .Fields.Append "CantCuotas", adDecimal
        .Fields.Item("CantCuotas").Precision = 19
        .Fields.Item("CantCuotas").NumericScale = 5
        .Fields.Append "CantCuotasPorTransferir", adDecimal
        .Fields.Item("CantCuotasPorTransferir").Precision = 19
        .Fields.Item("CantCuotasPorTransferir").NumericScale = 5
        .Fields.Append "NumCertificadoTransferente", adVarChar, 10
        .Fields.Append "FechaOperacion", adDate, 10
        .Fields.Append "NumOperacion", adVarChar, 10
        .Fields.Append "TipoOperacion", adVarChar, 2
        .Fields.Append "ClaseOperacion", adVarChar, 2
        .LockType = adLockBatchOptimistic
    End With
    
    adoTransferente.Open

End Sub
Private Sub ConfiguraDetalleGrillaTransferido()

    Set adoTransferido = New ADODB.Recordset
    
'CodFondoTransferido
'CodAdministradoraTransferido
'CodParticipeTransferido
'DescripParticipeTransferido
'NumSecuencial
'FechaOperacion
'NumFolio
'TipoFormaIngreso
'DescripFormaIngreso
'CodFondoTransferente
'CodAdministradoraTransferente
'CodParticipeTransferente
'DescripParticipeTransferente
'NumCertificadoTransferente
'FechaSuscripcion
'ValorCuota
'CantCuotas
'NumOperacionOrigen
    
        'CodParticipeTransferente , NumCertificadoTransferente, CantCuotas, ValorCuota, CodParticipeTransferido
        
    
    With adoTransferido
        .CursorLocation = adUseClient
        .Fields.Append "CodFondoTransferido", adVarChar, 3
        .Fields.Append "CodAdministradoraTransferido", adVarChar, 3
        .Fields.Append "CodParticipeTransferido", adVarChar, 20
        .Fields.Append "DescripParticipeTransferido", adVarChar, 100
        .Fields.Append "NumSecuencial", adInteger, 3
        .Fields.Append "FechaOperacion", adDate, 10
        .Fields.Append "NumFolio", adVarChar, 15
        .Fields.Append "TipoFormaIngreso", adVarChar, 2
        .Fields.Append "DescripFormaIngreso", adVarChar, 50
        .Fields.Append "CodFondoTransferente", adVarChar, 3
        .Fields.Append "CodAdministradoraTransferente", adVarChar, 3
        .Fields.Append "CodParticipeTransferente", adVarChar, 20
        .Fields.Append "DescripParticipeTransferente", adVarChar, 100
        .Fields.Append "NumCertificadoTransferente", adVarChar, 10
        .Fields.Append "FechaSuscripcion", adDate, 10
        .Fields.Append "ValorCuota", adDecimal
        .Fields.Item("ValorCuota").Precision = 19
        .Fields.Item("ValorCuota").NumericScale = 5
        .Fields.Append "CantCuotas", adDecimal
        .Fields.Item("CantCuotas").Precision = 19
        .Fields.Item("CantCuotas").NumericScale = 5
        .Fields.Append "NumOperacionOrigen", adVarChar, 10
        .LockType = adLockBatchOptimistic
    End With
    
    adoTransferido.Open

End Sub

Private Sub CargarDetalleGrillaRegistro()

'    Dim intRegistro As Integer
'
'    If adoFondoComisionistaCondicion("EstadoRegistro") = "C" Then
'        If adoFondoComisionistaCondicion("IndCondicionVigente") = Valor_Indicador Then
'            cmdActualizar.Enabled = True
'            cmdQuitar.Enabled = False
'        Else
'            cmdActualizar.Enabled = False
'            cmdQuitar.Enabled = False
'        End If
'    Else
'        cmdActualizar.Enabled = True
'        cmdQuitar.Enabled = True
'    End If
'
'    'Vigencia
'    dtpFechaDesde.Value = adoFondoComisionistaCondicion("FechaInicio").Value
'    dtpFechaHasta.Value = adoFondoComisionistaCondicion("FechaFin").Value
'    chkIndeterminado.Value = IIf(adoFondoComisionistaCondicion("IndIndeterminado").Value = Valor_Indicador, vbChecked, vbUnchecked)
'
'    'Tasa y Devengo
'    txtPorcenTasa.Text = adoFondoComisionistaCondicion("PorcenTasa").Value
'
'    intRegistro = ObtenerItemLista(arrTipoTasa, adoFondoComisionistaCondicion("CodTipoTasa").Value)
'    If intRegistro >= 0 Then cboTipoTasa.ListIndex = intRegistro
'
'    intRegistro = ObtenerItemLista(arrPeriodoTasa, adoFondoComisionistaCondicion("CodPeriodoTasa").Value)
'    If intRegistro >= 0 Then cboPeriodoTasa.ListIndex = intRegistro
'
'    intRegistro = ObtenerItemLista(arrPeriodoCapitalizacion, adoFondoComisionistaCondicion("CodPeriodoCapitalizacion").Value)
'    If intRegistro >= 0 Then cboPeriodoCapitalizacion.ListIndex = intRegistro
'
'    intRegistro = ObtenerItemLista(arrBaseCalculo, adoFondoComisionistaCondicion("CodBaseCalculo").Value)
'    If intRegistro >= 0 Then cboBaseCalculo.ListIndex = intRegistro
'
'    intRegistro = ObtenerItemLista(arrModalidadDevengo, adoFondoComisionistaCondicion("CodModalidadDevengo").Value)
'    If intRegistro >= 0 Then cboModalidadDevengo.ListIndex = intRegistro
'
'    intRegistro = ObtenerItemLista(arrPeriodoDevengo, adoFondoComisionistaCondicion("CodPeriodoDevengo").Value)
'    If intRegistro >= 0 Then cboPeriodoDevengo.ListIndex = intRegistro
'
'    lblFormulaDevengo.Caption = adoFondoComisionistaCondicion("DescripFormulaDevengo").Value
'    strCodFormulaDevengo = adoFondoComisionistaCondicion("CodFormulaDevengo").Value
'
'    'Pagos
'    txtPagoCada.Text = adoFondoComisionistaCondicion("NumPeriodoPago").Value
'
'    intRegistro = ObtenerItemLista(arrFrecuenciaPago, adoFondoComisionistaCondicion("CodFrecuenciaPago").Value)
'    If intRegistro >= 0 Then cboFrecuenciaPago.ListIndex = intRegistro
'
'    chkFinDeMes.Value = IIf(adoFondoComisionistaCondicion("IndFinDeMes").Value = Valor_Indicador, vbChecked, vbUnchecked)
'
'    intRegistro = ObtenerItemLista(arrTipoDesplazamiento, adoFondoComisionistaCondicion("CodTipoDesplazamiento").Value)
'    If intRegistro >= 0 Then cboTipoDesplazamiento.ListIndex = intRegistro
'

End Sub

