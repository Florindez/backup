VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Begin VB.Form frmCronograma 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Condiciones Financieras"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdVistaPrevia 
      Caption         =   "&Vista Previa"
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
      Left            =   6810
      Picture         =   "frmCronograma.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   1260
      Width           =   1200
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
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
      Left            =   2250
      Picture         =   "frmCronograma.frx":0BE2
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   7320
      Width           =   1200
   End
   Begin VB.CheckBox chkigv 
      Caption         =   "Cálculo con IGV"
      BeginProperty Font 
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
      Left            =   2910
      TabIndex        =   62
      Top             =   870
      Value           =   1  'Checked
      Width           =   1965
   End
   Begin VB.ComboBox cbPeriodoCapitalizacion 
      Height          =   315
      ItemData        =   "frmCronograma.frx":198C
      Left            =   1410
      List            =   "frmCronograma.frx":198E
      Style           =   2  'Dropdown List
      TabIndex        =   65
      Top             =   2010
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
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
      Left            =   690
      Picture         =   "frmCronograma.frx":1990
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   7320
      Width           =   1200
   End
   Begin VB.ComboBox cbFormaCalculo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1830
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Frame gbTramos 
      Caption         =   "Especificación de tramos"
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
      Height          =   2985
      Left            =   4200
      TabIndex        =   47
      Top             =   5250
      Width           =   3855
      Begin VB.OptionButton optCuota 
         Caption         =   "Por cuota"
         BeginProperty Font 
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
         Left            =   2040
         TabIndex        =   50
         Top             =   570
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optAmortizacion 
         Caption         =   "Por amortización"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   49
         Top             =   570
         Width           =   1785
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dgvTramos 
         Height          =   1905
         Left            =   120
         OleObjectBlob   =   "frmCronograma.frx":1F84
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   840
         Width           =   3480
      End
      Begin TAMControls.TAMTextBox txtCantidadTramos 
         Height          =   315
         Left            =   2520
         TabIndex        =   58
         Top             =   210
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   8.25
         Container       =   "frmCronograma.frx":3FAE
         Apariencia      =   1
         Borde           =   1
      End
      Begin VB.Label lblCantidadTramos 
         Caption         =   "Cantidad de tramos"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   48
         Top             =   300
         Width           =   1845
      End
   End
   Begin VB.Frame gbDesembolsos 
      Caption         =   "Desembolsos múltiples"
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
      Height          =   2820
      Left            =   4200
      TabIndex        =   45
      Top             =   2400
      Width           =   3855
      Begin DXDBGRIDLibCtl.dxDBGrid dgvDesembolsos 
         Height          =   1815
         Left            =   120
         OleObjectBlob   =   "frmCronograma.frx":3FCA
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   570
         Width           =   3495
      End
      Begin TAMControls.TAMTextBox txtCantidadDesembolsos 
         Height          =   315
         Left            =   2520
         TabIndex        =   56
         Top             =   210
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   8.25
         Container       =   "frmCronograma.frx":5C68
         Apariencia      =   1
         Borde           =   1
      End
      Begin VB.Label lblCantidadDesembolsos 
         Caption         =   "Cantidad de desembolsos:"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   46
         Top             =   300
         Width           =   2325
      End
   End
   Begin VB.CheckBox chkDesembolsosMultiples 
      Caption         =   "Desembolsos múltiples"
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
      Left            =   4260
      TabIndex        =   44
      Top             =   2070
      Width           =   2295
   End
   Begin VB.Frame gbParametrosCuponera 
      Caption         =   "Parámetros de calendario de cuotas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   60
      TabIndex        =   29
      Top             =   2400
      Width           =   4035
      Begin VB.ComboBox cbTipoDia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   3840
         Width           =   2505
      End
      Begin VB.ComboBox cbDesplazamientoPago 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   3480
         Width           =   3825
      End
      Begin VB.ComboBox cbDesplazamientoCorte 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   2850
         Width           =   3825
      End
      Begin VB.CheckBox chkCorteAFinPeriodo 
         Caption         =   "Corte a fin de periodo de cuota"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   37
         Top             =   1800
         Visible         =   0   'False
         Width           =   3165
      End
      Begin MSComCtl2.DTPicker dtpCortePrimerCupon 
         Height          =   315
         Left            =   2340
         TabIndex        =   36
         Top             =   1440
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   175570945
         CurrentDate     =   40413
      End
      Begin VB.CheckBox chkCortePrimer 
         Caption         =   "Corte primera cuota:"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   35
         Top             =   1440
         Width           =   2355
      End
      Begin MSComCtl2.DTPicker dtpAPartir 
         Height          =   315
         Left            =   2340
         TabIndex        =   34
         Top             =   1080
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   175570945
         CurrentDate     =   40413
      End
      Begin VB.CheckBox chkAPartir 
         Caption         =   "A partir de la fecha"
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
         Left            =   120
         TabIndex        =   33
         Top             =   1110
         Width           =   2055
      End
      Begin VB.ComboBox cbUnidadesPeriodo 
         Height          =   315
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   720
         Width           =   1965
      End
      Begin VB.ComboBox cbPeriodoCupon 
         Height          =   315
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   360
         Width           =   1965
      End
      Begin TAMControls.TAMTextBox txtUnidadesPeriodo 
         Height          =   315
         Left            =   960
         TabIndex        =   59
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   8.25
         Container       =   "frmCronograma.frx":5C84
         Estilo          =   3
         Apariencia      =   1
         Borde           =   1
      End
      Begin TAMControls.TAMTextBox txtDiasMinimosCobroInteres 
         Height          =   315
         Left            =   3150
         TabIndex        =   67
         Top             =   2130
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   8.25
         Container       =   "frmCronograma.frx":5CA0
         Apariencia      =   1
         Borde           =   1
      End
      Begin VB.Label lblCada 
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
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   780
         Width           =   495
      End
      Begin VB.Label lblDiasMinimosCobroInteres 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Días mínimos de cobro de Interés"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   66
         Top             =   2190
         Width           =   2895
      End
      Begin VB.Label lblTipoDia 
         Caption         =   "Tipo de día:"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   42
         Top             =   3900
         Width           =   1425
      End
      Begin VB.Label lblDesplazamientoPago 
         Caption         =   "Desplazamiento de la fecha de pago:"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   40
         Top             =   3240
         Width           =   3225
      End
      Begin VB.Label lblDesplazamientoCorte 
         Caption         =   "Desplazamiento de la fecha de corte:"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   38
         Top             =   2580
         Width           =   3375
      End
      Begin VB.Label lblPeriodoCupon 
         Caption         =   "Periodo de cuota:"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   30
         Top             =   420
         Width           =   1635
      End
   End
   Begin VB.Frame gbAjuste 
      Caption         =   "Parámetros del ajuste"
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
      Height          =   1935
      Left            =   11700
      TabIndex        =   20
      Top             =   6540
      Width           =   4035
      Begin VB.ComboBox cbFinIndice 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1440
         Width           =   2235
      End
      Begin VB.ComboBox cbInicioIndice 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1080
         Width           =   2235
      End
      Begin VB.ComboBox cbClaseAjuste 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   720
         Width           =   2235
      End
      Begin VB.ComboBox cbTipoAjuste 
         Height          =   315
         Left            =   1710
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   360
         Width           =   2235
      End
      Begin VB.Label lblFinIndice 
         Caption         =   "Fin Indice"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   24
         Top             =   1500
         Width           =   945
      End
      Begin VB.Label lblInicioIndice 
         Caption         =   "Inicio indice"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   23
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label lblClaseAjuste 
         Caption         =   "Clase de ajuste"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   22
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label lblTipoAjuste 
         Caption         =   "Tipo Ajuste"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   21
         Top             =   420
         Width           =   1155
      End
   End
   Begin VB.ComboBox cbPeriodoTasa 
      Height          =   315
      ItemData        =   "frmCronograma.frx":5CBC
      Left            =   1410
      List            =   "frmCronograma.frx":5CBE
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1620
      Width           =   1755
   End
   Begin VB.ComboBox cbTipoTasa 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1410
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1230
      Width           =   1755
   End
   Begin VB.ComboBox cbAmortizacion 
      Height          =   315
      Left            =   4830
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1620
      Width           =   1755
   End
   Begin VB.ComboBox cbBaseCalculo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4830
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1230
      Width           =   1755
   End
   Begin VB.ComboBox cbTipoCupon 
      Height          =   315
      Left            =   13350
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   6030
      Width           =   2175
   End
   Begin VB.CheckBox chkConAjuste 
      Caption         =   "Con ajuste"
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
      Left            =   11910
      TabIndex        =   14
      Top             =   5310
      Width           =   1425
   End
   Begin VB.CheckBox chkCuponCero 
      Caption         =   "Instrumento con cupón cero"
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
      Left            =   11910
      TabIndex        =   13
      Top             =   5610
      Width           =   2745
   End
   Begin VB.CheckBox chkNumeroCuotas 
      Caption         =   "Número de cuotas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   5490
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   1875
   End
   Begin MSComCtl2.DTPicker dtpVencimiento 
      Height          =   315
      Left            =   6690
      TabIndex        =   0
      Top             =   480
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      Format          =   175570945
      CurrentDate     =   40413
   End
   Begin MSComCtl2.DTPicker dtpEmision 
      Height          =   315
      Left            =   6690
      TabIndex        =   1
      Top             =   90
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      Format          =   175570945
      CurrentDate     =   40413
   End
   Begin TAMControls.TAMTextBox txtTasa 
      Height          =   315
      Left            =   1410
      TabIndex        =   57
      Top             =   840
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8.25
      Locked          =   -1  'True
      Container       =   "frmCronograma.frx":5CC0
      Apariencia      =   1
      Borde           =   1
   End
   Begin TAMControls.TAMTextBox txtValorNominal 
      Height          =   315
      Left            =   1410
      TabIndex        =   60
      Top             =   450
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8.25
      Locked          =   -1  'True
      Container       =   "frmCronograma.frx":5CDC
      Apariencia      =   1
      Borde           =   1
   End
   Begin TAMControls.TAMTextBox txtCuotas 
      Height          =   315
      Left            =   4830
      TabIndex        =   61
      Top             =   450
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8.25
      Locked          =   -1  'True
      Container       =   "frmCronograma.frx":5CF8
      Apariencia      =   1
      Borde           =   1
   End
   Begin TAMControls.TAMTextBox txtBeneficiario 
      Height          =   285
      Left            =   1410
      TabIndex        =   55
      Top             =   90
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8.25
      Locked          =   -1  'True
      Container       =   "frmCronograma.frx":5D14
      Apariencia      =   1
      Borde           =   1
   End
   Begin VB.Label lblNumCuotas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número de Cuotas:"
      BeginProperty Font 
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
      Left            =   3210
      TabIndex        =   71
      Top             =   510
      Width           =   1785
   End
   Begin VB.Label Label1 
      Caption         =   "Capitalización:"
      BeginProperty Font 
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
      Left            =   120
      TabIndex        =   64
      Top             =   2070
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblFormaCalculo 
      Caption         =   "Forma de cálculo"
      BeginProperty Font 
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
      Left            =   180
      TabIndex        =   51
      Top             =   6780
      Width           =   1695
   End
   Begin VB.Label lblBeneficiario 
      Caption         =   "Beneficiario:"
      BeginProperty Font 
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
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblValorNominal 
      Caption         =   "Valor Nominal:"
      BeginProperty Font 
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
      Left            =   120
      TabIndex        =   11
      Top             =   510
      Width           =   1245
   End
   Begin VB.Label lblTipoCupon 
      Caption         =   "Tipo de cupón"
      BeginProperty Font 
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
      Left            =   11940
      TabIndex        =   10
      Top             =   6090
      Width           =   1395
   End
   Begin VB.Label lblEmision 
      Caption         =   "Emisión:"
      BeginProperty Font 
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
      Left            =   5520
      TabIndex        =   9
      Top             =   120
      Width           =   765
   End
   Begin VB.Label lblVencimiento 
      Caption         =   "Vencimiento:"
      BeginProperty Font 
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
      Left            =   5520
      TabIndex        =   8
      Top             =   510
      Width           =   1125
   End
   Begin VB.Label lblBaseCalculo 
      Caption         =   "Base de cálculo:"
      BeginProperty Font 
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
      Left            =   3300
      TabIndex        =   7
      Top             =   1290
      Width           =   1485
   End
   Begin VB.Label lblAmortizacion 
      Caption         =   "Amortización:"
      BeginProperty Font 
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
      Left            =   3300
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblTasa 
      Caption         =   "Tasa:"
      BeginProperty Font 
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
      Left            =   120
      TabIndex        =   5
      Top             =   900
      Width           =   555
   End
   Begin VB.Label lblTipoTasa 
      Caption         =   "Tipo de tasa:"
      BeginProperty Font 
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
      Left            =   120
      TabIndex        =   4
      Top             =   1290
      Width           =   1335
   End
   Begin VB.Label lblPeriodoTasa 
      Caption         =   "Periodo:"
      BeginProperty Font 
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
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   885
   End
End
Attribute VB_Name = "frmCronograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fechaEmision     As Date
Private fechavencimiento As Date
Private valorNominal     As Double
Public beneficiario      As String
Public codigoUnico       As String
Public codSolicitud      As String
Private Codfile          As String
Private codAnalitica     As String
Public flagVisor         As Boolean
Public codFondo          As String
Public codAdministradora As String
Private nemotecnico      As String

Dim strCodSubDetalleFile As String

Dim flagDesembolsos      As Boolean
Dim flagTramos           As Boolean
Dim cti_igv              As Boolean

Dim numDesembolsos       As Integer
Dim Tasa                 As Double
Dim tasaDiaria           As Double
Dim cantCupones          As Integer
Dim listaCupones()

Dim arrBaseCalculo()         As String
Dim arrPeriodo()             As String
Dim arrPeriodoCupon()        As String
Dim arrUnidadPeriodo()       As String
Dim arrDesplazamientoCorte() As String
Dim arrDesplazamientoPago()  As String
Dim arrTipoVac()             As String
Dim arrInicioIndice()        As String
Dim arrFinIndice()           As String
Dim arrTipoAjuste()          As String
Dim arrTipoTasa()            As String
Dim arrTipoAmortizacion()    As String
Dim arrTipoCupon()           As String
Dim arrTipoDia()             As String
Dim arrFormaCalculo()        As String

Dim strCodDesplazamientoCorte As String
Dim strCodDesplazamientoPago As String

Dim adoRegistroDesembolso    As ADODB.Recordset
Dim adoRegistroTramo         As ADODB.Recordset

Public Sub setFondo(cf As String)
    codFondo = cf
End Sub

Private Sub cbDesplazamientoCorte_Click()
    strCodDesplazamientoCorte = arrDesplazamientoCorte(cbDesplazamientoCorte.ListIndex)
    cantCupones = calcularCantidadCupones()
    txtCuotas.Text = Str$(cantCupones)
End Sub

Private Sub cbDesplazamientoPago_Click()
    strCodDesplazamientoPago = arrDesplazamientoPago(cbDesplazamientoPago.ListIndex)
    cantCupones = calcularCantidadCupones()
    txtCuotas.Text = Str$(cantCupones)
End Sub

Private Sub cmdGuardar_Click()
    
    If MsgBox("Si graba las condiciones, ya no se podrán modificar. ¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
        Call Grabar(False)
    End If
    
End Sub

Private Sub cmdSalir_Click()
    Call Salir
End Sub

Private Sub cmdVistaPrevia_Click()
    cti_igv = chkigv.Value
    Call Grabar(True)
    'Call GeneraCuponera(0, codAnalitica, codigoUnico, codFondo, gstrCodAdministradora, codSolicitud, strCodSubDetalleFile, False, cti_igv)
End Sub

Private Sub dtpEmision_Change()
    fechaEmision = dtpEmision.Value
End Sub

Private Sub dtpVencimiento_Change()
    fechavencimiento = dtpVencimiento.Value
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    cantCupones = calcularCantidadCupones()
    txtCuotas.Text = Trim$(Str$(cantCupones))

    If KeyCode = 27 Then
        Unload Me
    End If

End Sub

Private Sub Salir()
    Unload Me
End Sub

Private Sub Form_Load()
        
    dtpEmision.Value = gdatFechaActual
    dtpVencimiento.Value = gdatFechaActual
    fechaEmision = dtpEmision.Value
    fechavencimiento = dtpVencimiento.Value
    
    txtBeneficiario.Text = beneficiario
    txtValorNominal.Text = valorNominal
    flagDesembolsos = False
    flagTramos = False
    
    Call CargarListas
    ConfGrid dgvDesembolsos, True, False, False, False
    ConfGrid dgvTramos, True, False, False, False
    txtCuotas.Text = Str$(calcularCantidadCupones)
    txtUnidadesPeriodo.Text = 1
    
    'Call stub
    flagDesembolsos = True
    flagTramos = True
    
    Dim adoTituloInversion As ADODB.Recordset
    
    adoComm.CommandText = "select II.CodFile, II.CodSubDetalleFile, II.CodAnalitica, II.CodFondo, II.CodAdministradora, II.Nemotecnico, II.DescripTitulo, " & _
                            " II.FechaEmision, II.FechaVencimiento, ISL.MontoAprobado as ValorNominal, II.BaseAnual, ISL.TasaInteres, ISL.TipoTasa  " & _
                            " from InstrumentoInversion II join InversionSolicitud ISL on (II.CodFondo = ISL.CodFondo and II.CodAdministradora = ISL.CodAdministradora  " & _
                            " and II.CodFile = ISL.CodFile and II.CodAnalitica = ISL.CodAnalitica)  " & _
                            " where II.CodTitulo = '" & codigoUnico & "' and ISL.NumSolicitud = '" & codSolicitud & "' and II.CodFondo = '" & gstrCodFondoContable & _
                            " ' and II.CodAdministradora = '" & gstrCodAdministradora & "'"
                            
    Set adoTituloInversion = adoComm.Execute
    adoTituloInversion.MoveFirst
    
    Codfile = adoTituloInversion.Fields.Item("CodFile").Value
    strCodSubDetalleFile = adoTituloInversion.Fields.Item("CodSubDetalleFile").Value
    codAnalitica = adoTituloInversion.Fields.Item("CodAnalitica").Value
    codFondo = adoTituloInversion.Fields.Item("CodFondo").Value
    codAdministradora = adoTituloInversion.Fields.Item("CodAdministradora").Value
    nemotecnico = adoTituloInversion.Fields.Item("Nemotecnico").Value
    'beneficiario = adoTituloInversion.Fields.Item("DescripTitulo").Value
    
    fechaEmision = adoTituloInversion.Fields.Item("FechaEmision").Value
    fechavencimiento = adoTituloInversion.Fields.Item("Fechavencimiento").Value
    dtpEmision.Value = fechaEmision
    dtpVencimiento.Value = fechavencimiento
    
    valorNominal = adoTituloInversion.Fields.Item("ValorNominal").Value
    txtValorNominal.Text = valorNominal
    Dim i As Integer

    cbBaseCalculo.ListIndex = ObtenerItemLista(arrBaseCalculo(), adoTituloInversion.Fields.Item("BaseAnual").Value)
    cbTipoTasa.ListIndex = ObtenerItemLista(arrTipoTasa(), adoTituloInversion.Fields.Item("TipoTasa").Value)
    
    Tasa = adoTituloInversion.Fields.Item("TasaInteres").Value
    txtTasa.Text = Tasa
    
    'JJCC: Carga por defecto los datos de la grilla de los desembolsos.
    gbTramos.Enabled = False
    chkDesembolsosMultiples.Enabled = False
    chkDesembolsosMultiples.Value = 0
    Call chkDesembolsosMultiples_Click
    txtCantidadDesembolsos.Text = "1"
    txtCantidadTramos.Text = "0"
    CentrarForm Me

    If flagVisor Then
        Call cargarCondicionesFinancieras
    End If
        
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
End Sub

Private Sub cargarCondicionesFinancieras()

    Dim adoRegistroCondicionesFinancieras As New ADODB.Recordset

    Dim comm                              As ADODB.Command
    Set comm = New ADODB.Command
    
    ReDim listaCupones(cantCupones)
    'Revisa si es operacion al descuento o con tasa de interes
    adoComm.CommandText = "SELECT CodSubDetalleFile, TasaInteres FROM InversionSolicitud WHERE CodFile in ('016','021') and CodTitulo = '" & codigoUnico & "'"
    Set adoRegistroCondicionesFinancieras = adoComm.Execute
    strCodSubDetalleFile = adoRegistroCondicionesFinancieras.Fields("CodSubDetalleFile").Value

    If adoRegistroCondicionesFinancieras.Fields("CodSubDetalleFile").Value = "001" Then
        Me.Caption = Me.Caption & " (Operación Con Tasa de Interés)"
        '       JJCC :Si la operación es con tasa de interés, se permitirá especificar sin incluye o no igv.
        chkigv.Enabled = True
        chkigv.Value = 1
    Else
        Me.Caption = Me.Caption & " (Operación Al Descuento)"
        '       JJCC: Si es al descuento, no se tomará en cuenta la tasa de igv.
        chkigv.Enabled = False
        chkigv.Value = 0
    End If
    
      txtTasa.Text = Str$(adoRegistroCondicionesFinancieras("TasaInteres").Value)
    
    'query de la tabla instrumentoInversionCondicionesFinancieras
    adoComm.CommandText = "SELECT IndNumCuotas, NumCuotas, FechaEmision, FechaVencimiento, ValorNominal," & _
                            "Tasa, TipoTasa, PeriodoTasa, TipoCupon, BaseCalculo, TipoAmortizacion, PeriodoCupon, " & _
                            "IndPeriodoPersonalizable, CantUnidadesPeriodo, UnidadPeriodo, DesplazamientoCorte, DesplazamientoPago, IndCortePrimerCupon, " & _
                            "FechaPrimerCorte, IndFechaAPartir, FechaAPartir, DiasMinimosCobroInteres, IndDesembolsosMultiples, CantDesembolsos, CantTramos, " & _
                            "TipoTramo from InstrumentoInversionCondicionesFinancieras where CodTitulo = '" & codigoUnico & "'"
    Set adoRegistroCondicionesFinancieras = adoComm.Execute
   
    'Procesamiento de adoRegistroCondicionesFinancieras
    If Not adoRegistroCondicionesFinancieras.EOF Then
        If adoRegistroCondicionesFinancieras.Fields.Item("IndNumCuotas") = "0" Then
            chkNumeroCuotas.Value = 0
        Else
            chkNumeroCuotas.Value = 1
        End If
        
        txtCuotas.Text = adoRegistroCondicionesFinancieras.Fields.Item("NumCuotas")
        dtpEmision.Value = adoRegistroCondicionesFinancieras.Fields.Item("FechaEmision")
        dtpVencimiento.Value = adoRegistroCondicionesFinancieras.Fields.Item("FechaVencimiento")
        txtValorNominal.Text = adoRegistroCondicionesFinancieras.Fields.Item("ValorNominal")
        txtTasa.Text = adoRegistroCondicionesFinancieras.Fields.Item("Tasa")

        If adoRegistroCondicionesFinancieras.Fields.Item("TipoTasa") = "01" Then
            cbTipoTasa.ListIndex = 0
        ElseIf adoRegistroCondicionesFinancieras.Fields.Item("TipoTasa") = "02" Then
            cbTipoTasa.ListIndex = 1
        End If

        cbPeriodoTasa.ListIndex = CInt(adoRegistroCondicionesFinancieras.Fields.Item("PeriodoTasa")) - 1
        cbTipoCupon.ListIndex = CInt(adoRegistroCondicionesFinancieras.Fields.Item("TipoCupon")) - 1
        cbBaseCalculo.ListIndex = CInt(adoRegistroCondicionesFinancieras.Fields.Item("BaseCalculo"))

        'ajuste de indexbasecalculo
        If cbBaseCalculo.ListIndex = 1 Then
            cbBaseCalculo.ListIndex = 4
        ElseIf cbBaseCalculo.ListIndex > 3 Then
            cbBaseCalculo.ListIndex = cbBaseCalculo.ListIndex - 4
        End If
        
        cbUnidadesPeriodo.ListIndex = adoRegistroCondicionesFinancieras.Fields.Item("UnidadPeriodo") - 1
        cbAmortizacion.ListIndex = CInt(adoRegistroCondicionesFinancieras.Fields.Item("TipoAmortizacion")) - 1
        cbPeriodoCupon.ListIndex = CInt(adoRegistroCondicionesFinancieras.Fields.Item("PeriodoCupon")) - 1
        'indPeriodoPersonalizable = CInt(adoRegistroCondicionesFinancieras.Fields.Item("IndPeriodoPersonalizable"))
        txtUnidadesPeriodo.Text = adoRegistroCondicionesFinancieras.Fields.Item("CantUnidadesPeriodo")
        
        cbDesplazamientoCorte.ListIndex = CInt(adoRegistroCondicionesFinancieras.Fields.Item("DesplazamientoCorte"))
        cbDesplazamientoPago.ListIndex = CInt(adoRegistroCondicionesFinancieras.Fields.Item("DesplazamientoPago"))
        chkCortePrimer.Value = adoRegistroCondicionesFinancieras.Fields.Item("IndCortePrimerCupon")
        dtpCortePrimerCupon.Value = adoRegistroCondicionesFinancieras.Fields.Item("FechaPrimerCorte")
        chkAPartir.Value = CInt(adoRegistroCondicionesFinancieras.Fields.Item("IndFechaAPartir"))
        dtpAPartir.Value = adoRegistroCondicionesFinancieras.Fields.Item("FechaAPartir")
        txtDiasMinimosCobroInteres.Text = adoRegistroCondicionesFinancieras.Fields.Item("DiasMinimosCobroInteres")
        
        chkDesembolsosMultiples.Value = adoRegistroCondicionesFinancieras.Fields.Item("IndDesembolsosMultiples")
        txtCantidadDesembolsos.Text = adoRegistroCondicionesFinancieras.Fields.Item("CantDesembolsos")
        txtCantidadTramos.Text = adoRegistroCondicionesFinancieras.Fields.Item("CantTramos")
        
        If adoRegistroCondicionesFinancieras.Fields.Item("TipoTramo") = "cuota" Then
            optAmortizacion.Value = False
            optCuota.Value = True
        Else
            optAmortizacion.Value = True
            optCuota.Value = False
        End If
        
        Set dgvDesembolsos.DataSource = Nothing

        With dgvDesembolsos
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = gstrConnectConsulta
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.Active = False
            .Dataset.ADODataset.CommandText = "SELECT NumDesembolso as colNum, FechaDesembolso as colFecha, ValorDesembolso as colValor  From InstrumentoInversionCalendarioDesembolso where CodTitulo =  '" & codigoUnico & "'"
            .Dataset.DisableControls
            .Dataset.Active = True
            .KeyField = "colNum"
        End With
              
        Set dgvTramos.DataSource = Nothing

        With dgvTramos
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = gstrConnectConsulta
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.Active = False
            .Dataset.ADODataset.CommandText = "SELECT  NumTramo as colNum, InicioTramo as colInicio, FinTramo as colFin, Valor as colValor From InstrumentoInversionCalendarioTramo where CodTitulo =  '" & codigoUnico & "'"
            .Dataset.DisableControls
            .Dataset.Active = True
            .KeyField = "colNum"
        End With
        
        Call bloquearCampos

    End If

End Sub

Private Sub bloquearCampos()
    chkigv.Enabled = False 'JJCC
    chkConAjuste.Enabled = False
    chkCuponCero.Enabled = False
    chkNumeroCuotas.Enabled = False
    cbTipoCupon.Enabled = False
    cbBaseCalculo.Enabled = False
    cbPeriodoTasa.Enabled = False
    cbTipoTasa.Enabled = False
    cbAmortizacion.Enabled = False
    txtTasa.Locked = True
    chkDesembolsosMultiples.Enabled = False
    gbParametrosCuponera.Enabled = False
    gbAjuste.Enabled = False
    cmdGuardar.Visible = False
    gbDesembolsos.Enabled = False
    dtpEmision.Enabled = False
    dtpVencimiento.Enabled = False
    cmdVistaPrevia.Visible = False
End Sub

Public Sub CargarControlListaLocal(ByVal strSentencia As String, _
                                   ByVal CtrlNombre As Control, _
                                   ByRef arrControl() As String, _
                                   ByVal strValor As String)

    Dim adoBusqueda As ADODB.Recordset
    Dim intCont     As Long
    
    Set adoBusqueda = New ADODB.Recordset
    
    adoComm.CommandText = strSentencia
    Set adoBusqueda = adoComm.Execute

    CtrlNombre.Clear
    intCont = 0
    ReDim arrControl(intCont)
    
    Do Until adoBusqueda.EOF
        CtrlNombre.AddItem adoBusqueda("DESCRIP")
        ReDim Preserve arrControl(intCont)
        arrControl(intCont) = adoBusqueda("CODIGO")
        adoBusqueda.MoveNext
        intCont = intCont + 1
    Loop
   
    adoBusqueda.Close: Set adoBusqueda = Nothing

End Sub

Private Sub CargarListas()
    Dim strSQL As String
    
    '-----base de calculo--------
    strSQL = "{ call up_ACSelDatos(45) }"
    CargarControlListaLocal strSQL, cbBaseCalculo, arrBaseCalculo(), Sel_Defecto
    cbBaseCalculo.ListIndex = 0
    '-----tipo de frecuencia
    strSQL = "{ call up_ACSelDatos(52) }"
    CargarControlListaLocal strSQL, cbPeriodoTasa, arrPeriodo(), Sel_Defecto
    cbPeriodoTasa.ListIndex = 0
    '-----periodo de cupón
    strSQL = "{ call up_ACSelDatos(50) }"
    CargarControlListaLocal strSQL, cbPeriodoCupon, arrPeriodoCupon(), Sel_Defecto
    cbPeriodoCupon.ListIndex = 0
    '-----unidad de periodo
    strSQL = "{ call up_ACSelDatos(46) }"
    CargarControlListaLocal strSQL, cbUnidadesPeriodo, arrUnidadPeriodo(), Sel_Defecto
    cbUnidadesPeriodo.ListIndex = 0
    '-----desplazamiento
    strSQL = "{ call up_ACSelDatos(47) }"
    CargarControlListaLocal strSQL, cbDesplazamientoCorte, arrDesplazamientoCorte(), Sel_Defecto
    CargarControlListaLocal strSQL, cbDesplazamientoPago, arrDesplazamientoPago(), Sel_Defecto
    cbDesplazamientoCorte.ListIndex = 0
    cbDesplazamientoPago.ListIndex = 0
    '-----inicio indice
    strSQL = "{ call up_ACSelDatos(48) }"
    CargarControlListaLocal strSQL, cbInicioIndice, arrInicioIndice(), Sel_Defecto
    cbInicioIndice.ListIndex = 0
    '-----tipo vac
    strSQL = "{ call up_ACSelDatos(54) }"
    CargarControlListaLocal strSQL, cbFinIndice, arrFinIndice(), Sel_Defecto
    cbFinIndice.ListIndex = 0

    '-----tipoajuste
    strSQL = "{ call up_ACSelDatos(49) }"
    CargarControlListaLocal strSQL, cbTipoAjuste, arrTipoAjuste(), Sel_Defecto
    cbTipoAjuste.ListIndex = 0
    '-----tipoajuste
    strSQL = "{ call up_ACSelDatos(51) }"
    CargarControlListaLocal strSQL, cbTipoTasa, arrTipoTasa(), Sel_Defecto
    cbTipoTasa.ListIndex = 0
    '-----modalidad de amortizacion
    strSQL = "{ call up_ACSelDatos(53) }"
    CargarControlListaLocal strSQL, cbAmortizacion, arrTipoAmortizacion(), Sel_Defecto
    cbAmortizacion.ListIndex = 0
    '-----tipo cupon
    strSQL = "{ call up_ACSelDatos(4) }"
    CargarControlListaLocal strSQL, cbTipoCupon, arrTipoCupon(), Sel_Defecto
    cbTipoCupon.ListIndex = 0
    '-----tipo dia
    strSQL = "{ call up_ACSelDatos(55) }"
    CargarControlListaLocal strSQL, cbTipoDia, arrTipoDia(), Sel_Defecto
    cbTipoDia.ListIndex = 0
    '-----forma de calculo
    strSQL = "{ call up_ACSelDatos(56) }"
    CargarControlListaLocal strSQL, cbFormaCalculo, arrFormaCalculo(), Sel_Defecto
    cbFormaCalculo.ListIndex = 0
End Sub

Private Sub cbUnidadesPeriodo_Click()
    cantCupones = calcularCantidadCupones()
    txtCuotas.Text = Str$(cantCupones)
    
    txtDiasMinimosCobroInteres_Change
End Sub

Private Sub chkCorteAFinPeriodo_Click()
    cantCupones = calcularCantidadCupones()
    txtCuotas.Text = Str$(cantCupones)
End Sub

Private Sub dgvDesembolsos_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, _
                                        ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    dgvDesembolsos.Columns.Item(1).DisableEditor = False
    dgvDesembolsos.Columns.Item(1).ReadOnly = False
    dgvDesembolsos.Columns.Item(2).DisableEditor = False
    dgvDesembolsos.Columns.Item(2).ReadOnly = False

    If adoRegistroDesembolso.Fields.Item(0).Value = 1 Then
        dgvDesembolsos.Columns.Item(1).DisableEditor = True
    Else
        dgvDesembolsos.Columns.Item(1).DisableEditor = False
    End If

    If adoRegistroDesembolso.Fields.Item(0).Value = dgvDesembolsos.Count Then
        dgvDesembolsos.Columns.Item(2).DisableEditor = True
    Else
        dgvDesembolsos.Columns.Item(2).DisableEditor = False
    End If

End Sub

Private Sub dgvDesembolsos_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Dim cellx As Integer
    Dim celly As Integer
    Dim i As Integer
    celly = adoRegistroDesembolso.Fields.Item(0) - 1
    cellx = dgvDesembolsos.Columns.FocusedIndex

    If dgvDesembolsos.Dataset.State = dsEdit Then
        dgvDesembolsos.Dataset.Post
    End If

    If cellx = 1 Then
        If (adoRegistroDesembolso.Fields.Item(0) = 1) And (adoRegistroDesembolso.Fields.Item(1).Value <> fechaEmision) Then
            adoRegistroDesembolso.Fields.Item(1).Value = fechaEmision
        Else

            If (adoRegistroDesembolso.Fields.Item(1).Value < fechaEmision) Or (adoRegistroDesembolso.Fields.Item(1).Value > fechavencimiento) Then
                adoRegistroDesembolso.Fields.Item(1).Value = fechaEmision
            End If
        End If
    End If

    If cellx = 2 Then
        Dim sumaDesembolso As Double
        sumaDesembolso = 0

        If adoRegistroDesembolso.Fields.Item(cellx).Value < 0 Then
            adoRegistroDesembolso.Fields.Item(cellx).Value = 0
        End If

        adoRegistroDesembolso.MoveFirst

        If dgvDesembolsos.Count > 1 Then

            For i = 0 To dgvDesembolsos.Count - 2
                sumaDesembolso = sumaDesembolso + adoRegistroDesembolso.Fields.Item(cellx).Value
                adoRegistroDesembolso.MoveNext
            Next

        Else
            sumaDesembolso = 0
        End If

        adoRegistroDesembolso.MoveLast
        adoRegistroDesembolso.Fields.Item(2).Value = valorNominal - sumaDesembolso
    End If

End Sub

Private Sub dgvTramos_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, _
                                   ByVal Node As DXDBGRIDLibCtl.IdxGridNode)

    If flagTramos Then
        If dgvTramos.Count > 1 Then
            If adoRegistroTramo.Fields.Item(0).Value >= dgvTramos.Count - 1 Then
                If adoRegistroTramo.Fields.Item(0).Value = dgvTramos.Count Then
                    dgvTramos.Columns.Item(3).DisableEditor = True
                Else
                    dgvTramos.Columns.Item(3).DisableEditor = False
                End If

                dgvTramos.Columns.Item(2).DisableEditor = True
            Else
                dgvTramos.Columns.Item(2).DisableEditor = False
            End If
        End If
    End If

End Sub

Private Sub dgvTramos_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Dim cellx As Integer
    Dim celly As Integer
    Dim i As Integer
    Dim ylimit As Integer
    
    
    cellx = dgvTramos.Columns.FocusedIndex
    celly = adoRegistroTramo.Fields.Item(0).Value - 1
    ylimit = dgvTramos.Count - 1
    flagTramos = False
    
    If dgvTramos.Dataset.State = dsEdit Then
        dgvTramos.Dataset.Post
    End If

    If cellx = 2 Then
        If adoRegistroTramo.Fields.Item(cellx).Value <= 0 Then
            adoRegistroTramo.Fields.Item(cellx).Value = adoRegistroTramo.Fields.Item(1).Value
        End If

        If adoRegistroTramo.Fields.Item(cellx).Value < adoRegistroTramo.Fields.Item(cellx - 1).Value Then
            adoRegistroTramo.Fields.Item(cellx).Value = adoRegistroTramo.Fields.Item(cellx - 1).Value
        End If

        If adoRegistroTramo.Fields.Item(cellx).Value > (cantCupones - (dgvTramos.Count - (celly + 1))) Then
            adoRegistroTramo.Fields.Item(cellx).Value = cantCupones - (dgvTramos.Count - (celly + 1))
        End If

        Dim Valor As Integer
        Valor = adoRegistroTramo.Fields.Item(cellx).Value
        adoRegistroTramo.MoveNext

        For i = celly + 1 To ylimit
            Valor = Valor + 1
            adoRegistroTramo.Fields.Item(1).Value = Valor

            If (i < (ylimit - 1)) Then
                adoRegistroTramo.Fields.Item(2).Value = Valor
            End If

            adoRegistroTramo.MoveNext
        Next

        adoRegistroTramo.MoveFirst
        While adoRegistroTramo.Fields.Item(0).Value - 1 < celly
            adoRegistroTramo.MoveNext
        Wend

        If celly = (ylimit - 1) Then
            adoRegistroTramo.Fields.Item(cellx).Value = cantCupones - 1
        End If

        adoRegistroTramo.MoveLast
        adoRegistroTramo.Fields.Item(2).Value = cantCupones
        adoRegistroTramo.Fields.Item(1).Value = cantCupones
    End If

    flagTramos = True
End Sub

Private Sub cbAmortizacion_Click()

    If flagDesembolsos Then
        If cbAmortizacion.ListIndex > 2 Then
            gbTramos.Enabled = True
            chkDesembolsosMultiples.Enabled = True
            Call chkDesembolsosMultiples_Click
            txtCantidadDesembolsos.Text = "1"
        Else
            gbTramos.Enabled = False
            chkDesembolsosMultiples.Enabled = False
            chkDesembolsosMultiples.Value = 0
            Call chkDesembolsosMultiples_Click
            txtCantidadDesembolsos.Text = "1"
        End If
    End If

End Sub

Private Sub cbPeriodoCupon_Click()

    If cbUnidadesPeriodo.ListIndex = 6 And txtUnidadesPeriodo.Value > 999 Then
        txtUnidadesPeriodo.Text = 999
    End If
    If cbUnidadesPeriodo.ListIndex = 5 And txtUnidadesPeriodo.Value > 999 Then
        txtUnidadesPeriodo.Text = 999
    End If
    
    If cbUnidadesPeriodo.ListIndex <= 4 And txtUnidadesPeriodo.Value > 99 Then
        txtUnidadesPeriodo.Text = 99
    End If
  
    If cbUnidadesPeriodo.ListIndex = 0 And txtUnidadesPeriodo.Value > 50 Then
        txtUnidadesPeriodo.Text = 50
    End If
    
    If txtUnidadesPeriodo.Value = 0 Or txtUnidadesPeriodo.Text = "" Then
        txtUnidadesPeriodo.Text = 1
    End If
    
    
    If cbPeriodoCupon.ListIndex > 6 Then
        lblCada.Visible = True
        txtUnidadesPeriodo.Visible = True
        cbUnidadesPeriodo.Visible = True
    Else
        txtUnidadesPeriodo.Visible = False
        cbUnidadesPeriodo.Visible = False
        lblCada.Visible = False
    End If

    cantCupones = calcularCantidadCupones()
    txtCuotas.Text = Str$(cantCupones)
    
    txtDiasMinimosCobroInteres_Change
End Sub

Private Sub chkApartir_Click()

    If chkAPartir.Value = 1 Then
        dtpAPartir.Enabled = True
    Else
        dtpAPartir.Enabled = False
    End If

    cantCupones = calcularCantidadCupones()
    txtCuotas.Text = Str$(cantCupones)
End Sub

Private Sub chkConAjuste_Click()

    If chkConAjuste.Value = 1 Then
        gbAjuste.Enabled = True
    Else
        gbAjuste.Enabled = False
    End If

End Sub

Private Sub chkCortePrimer_Click()

    If chkCortePrimer.Value = 1 Then
        dtpCortePrimerCupon.Enabled = True
    Else
        dtpCortePrimerCupon.Enabled = False
    End If

    cantCupones = calcularCantidadCupones()
    txtCuotas.Text = Str$(cantCupones)
End Sub

Private Sub chkDesembolsosMultiples_Click()

    If (chkDesembolsosMultiples.Value = 1) Then
        gbDesembolsos.Enabled = True
    Else
        gbDesembolsos.Enabled = False
        txtCantidadDesembolsos.Text = "1"
    End If

End Sub

'JJCC: Evento para especificar si la operación incluye o no igv
Private Sub chkigv_Click()

    If (chkigv.Value = 1) Then
        cti_igv = True
    Else
        cti_igv = False
    End If
    
End Sub

Private Sub chkNumeroCuotas_Click()

    If chkNumeroCuotas.Value = 1 Then
        txtCuotas.Enabled = True
    Else
        txtCuotas.Enabled = False
    End If

End Sub

Private Function calcularCantidadCupones() As Integer
    If (chkNumeroCuotas.Value <> 1) Then
        Dim flag As Boolean
        flag = True
        Dim numCup As Integer
        numCup = 0
        'FI fechaInicio, FC fechaCorte, FV fechaVencimiento
        Dim FC, FI, FV As Date
        FC = gdatFechaActual
        FV = dtpVencimiento.Value

        If (chkCuponCero.Value = 1) Then
            calcularCantidadCupones = 1
        Else

            If (chkAPartir.Value = 1) Then
                FI = dtpAPartir.Value
            Else
                FI = dtpEmision.Value
            End If

            While (flag)
                numCup = numCup + 1

                If (numCup = 1) And (chkCortePrimer.Value = 1) Then
                    FC = dtpCortePrimerCupon.Value
                Else

                    If (chkCorteAFinPeriodo.Value = 1) Then
                        If (numCup > 1) Then
                            FC = ultimaFechaPeriodo(DateAdd("d", 1, FI), cbPeriodoCupon.ListIndex)
                        Else
                            FC = ultimaFechaPeriodo(FI, cbPeriodoCupon.ListIndex)
                        End If

                    Else

                        If cbBaseCalculo.ListIndex < 2 Then 'caso 360

                            Select Case cbPeriodoCupon.ListIndex

                                Case 0
                                    FC = DateAdd("d", 360, FI)

                                Case 1
                                    FC = DateAdd("d", 180, FI)

                                Case 2
                                    FC = DateAdd("d", 90, FI)

                                Case 3
                                    FC = DateAdd("d", 60, FI)

                                Case 4
                                    FC = DateAdd("d", 30, FI)

                                Case 5
                                    FC = DateAdd("d", 15, FI)

                                Case 6
                                    FC = DateAdd("d", 1, FI)

                                Case 7

                                    Select Case cbUnidadesPeriodo.ListIndex

                                        Case 0
                                            FC = DateAdd("d", (CInt(txtUnidadesPeriodo.Text) * 360), FI)

                                        Case 1
                                            FC = DateAdd("d", (CInt(txtUnidadesPeriodo.Text) * 180), FI)

                                        Case 2
                                            FC = DateAdd("d", (CInt(txtUnidadesPeriodo.Text) * 90), FI)

                                        Case 3
                                            FC = DateAdd("d", (CInt(txtUnidadesPeriodo.Text) * 60), FI)

                                        Case 4
                                            FC = DateAdd("d", (CInt(txtUnidadesPeriodo.Text) * 30), FI)

                                        Case 5
                                            FC = DateAdd("d", (CInt(txtUnidadesPeriodo.Text) * 15), FI)

                                        Case 6
                                            FC = DateAdd("d", CInt(txtUnidadesPeriodo.Text), FI)
                                    End Select
                            End Select

                        Else

                            Select Case cbPeriodoCupon.ListIndex

                                Case 0
                                    FC = DateAdd("yyyy", 1, FI)

                                Case 1
                                    FC = DateAdd("m", 6, FI)

                                Case 2
                                    FC = DateAdd("m", 3, FI)

                                Case 3
                                    FC = DateAdd("m", 2, FI)

                                Case 4
                                    FC = DateAdd("m", 1, FI)

                                Case 5
                                    FC = DateAdd("d", 15, FI)

                                Case 6
                                    FC = DateAdd("d", 1, FI)

                                Case 7

                                    Select Case cbUnidadesPeriodo.ListIndex

                                        Case 0
                                            FC = DateAdd("yyyy", (CInt(txtUnidadesPeriodo.Text)), FI)

                                        Case 1
                                            FC = DateAdd("m", (CInt(txtUnidadesPeriodo.Text) * 6), FI)

                                        Case 2
                                            FC = DateAdd("m", (CInt(txtUnidadesPeriodo.Text) * 3), FI)

                                        Case 3
                                            FC = DateAdd("m", (CInt(txtUnidadesPeriodo.Text) * 2), FI)

                                        Case 4
                                            FC = DateAdd("m", (CInt(txtUnidadesPeriodo.Text)), FI)

                                        Case 5
                                            FC = DateAdd("d", (CInt(txtUnidadesPeriodo.Text) * 15), FI)

                                        Case 6
                                            FC = DateAdd("d", (CInt(txtUnidadesPeriodo.Text)), FI)
                                    End Select
                            End Select

                        End If
                    End If
                    
                    If FC = CDate("28/12/2011") Then
                        FC = FC
                    End If
                    
                End If
                    
                FC = desplazamientoDiaLaborable(FC, strCodDesplazamientoCorte)
                FI = FC
                
                If FC >= FV Then
                    FC = FV
                    flag = False
                End If

            Wend
            calcularCantidadCupones = numCup
        End If
    End If
   
End Function

Private Function TodoOK() As Boolean
    Dim result As Boolean
    
    result = True
    If arrTipoAmortizacion(cbAmortizacion.ListIndex) = "04" And CInt(txtCantidadTramos.Text) < 1 Then
        result = result And False
        MsgBox "Debe registrar al menos un tramo de cuotas!", vbCritical, "Faltan datos"
    End If
    
    TodoOK = result
End Function

Private Sub Grabar(ByVal blnVistaPrevia As Boolean)
    Dim numDesembolso As Integer
    Dim numTramo      As Integer
    
    'Se limpia informacion proveniente de las vistas previas anteriores
    adoComm.CommandText = "DELETE FROM InstrumentoInversionCondicionesFinancieras WHERE CodTitulo = '" & codigoUnico & "'"
    adoConn.Execute adoComm.CommandText
    
    adoComm.CommandText = "DELETE FROM InversionOperacionCalendarioCuota WHERE CodTitulo = '" & codigoUnico & "' and CodFondo = '" & gstrCodFondoContable & "'"
    adoConn.Execute adoComm.CommandText

    If Not TodoOK() Then Exit Sub
    
    'Grabar condiciones financieras
    adoComm.CommandText = "INSERT INTO InstrumentoInversionCondicionesFinancieras (CodTitulo, IndNumCuotas, NumCuotas, FechaEmision, FechaVencimiento, " & _
       "ValorNominal, Tasa, TipoTasa, PeriodoTasa, TipoCupon, BaseCalculo, TipoAmortizacion, PeriodoCupon, IndPeriodoPersonalizable, " & _
       "CantUnidadesPeriodo, UnidadPeriodo, DesplazamientoCorte, DesplazamientoPago, IndCortePrimerCupon, FechaPrimerCorte,  " & _
       "IndFechaAPartir, FechaAPartir, DiasMinimosCobroInteres, IndDesembolsosMultiples, CantDesembolsos, CantTramos, TipoTramo) " & _
       "VALUES ('" & codigoUnico & "'," & chkNumeroCuotas.Value & "," & cantCupones & ",'" & Format$(fechaEmision, "yyyyMMdd") & "','" & _
       Format$(fechavencimiento, "yyyyMMdd") & "'," & valorNominal & "," & _
       txtTasa.Value & ",'" & Format$(arrTipoTasa(cbTipoTasa.ListIndex), "00") & "','" & Format$(arrPeriodo(cbPeriodoTasa.ListIndex), "00") & _
       "','" & Format$(arrTipoCupon(cbTipoCupon.ListIndex), "00") & "','" & Format$(arrBaseCalculo(cbBaseCalculo.ListIndex), "00") & "','" & _
       Format$(arrTipoAmortizacion(cbAmortizacion.ListIndex), "00") & "','" & Format$(arrPeriodoCupon(cbPeriodoCupon.ListIndex), "00") & "',"
                        
    If cbPeriodoCupon.ListIndex = 7 Then 'Periodo de cupón personalizable
        adoComm.CommandText = adoComm.CommandText & "1," & txtUnidadesPeriodo.Value & ",'"
    Else
        adoComm.CommandText = adoComm.CommandText & "0,1,'"
    End If

    adoComm.CommandText = adoComm.CommandText & Format$(arrUnidadPeriodo(cbUnidadesPeriodo.ListIndex), "00") & "','" & _
                        strCodDesplazamientoCorte & "','" & strCodDesplazamientoPago & "'," & chkCortePrimer.Value & ",'" & _
                        Format$(dtpCortePrimerCupon.Value, "yyyyMMdd") & "'," & chkAPartir.Value & ",'" & Format$(dtpAPartir.Value, "yyyyMMdd") & _
                        "'," & txtDiasMinimosCobroInteres.Value & "," & chkDesembolsosMultiples.Value & "," & txtCantidadDesembolsos.Value & "," & _
                        txtCantidadTramos.Value & ","

    If optCuota.Value = True Then
        adoComm.CommandText = adoComm.CommandText & "'cuota')"
    Else
        adoComm.CommandText = adoComm.CommandText & "'amort')"
    End If

    adoConn.Execute adoComm.CommandText

    If cbAmortizacion.ListIndex = 3 Then
        
        If chkDesembolsosMultiples.Value = vbChecked Then
            'Grabacion de la data de los desembolsos
            If Not adoRegistroDesembolso.EOF Then
                adoComm.CommandText = "DELETE FROM InstrumentoInversionCalendarioDesembolso WHERE CodTitulo = '" & codigoUnico & "'"
                adoConn.Execute adoComm.CommandText
                adoRegistroDesembolso.MoveFirst
    
                For numDesembolso = 0 To txtCantidadDesembolsos.Value - 1
                    adoComm.CommandText = "INSERT INTO InstrumentoInversionCalendarioDesembolso (CodTitulo, NumDesembolso, ValorDesembolso, FechaDesembolso, EstadoDesembolso) VALUES ('" & codigoUnico & "'," & (numDesembolso + 1) & "," & adoRegistroDesembolso.Fields(2).Value & ",'" & Convertyyyymmdd(adoRegistroDesembolso.Fields(1).Value) & "','01')"
                    adoConn.Execute adoComm.CommandText
                    adoRegistroDesembolso.MoveNext
                Next
    
            End If
        End If
        'grabacion de la data de los tramos
        If Not adoRegistroTramo.EOF Then
            adoComm.CommandText = "DELETE FROM InstrumentoInversionCalendarioTramo WHERE CodTitulo = '" & codigoUnico & "'"
            adoConn.Execute adoComm.CommandText
            adoRegistroTramo.MoveFirst

            For numTramo = 0 To txtCantidadTramos.Value - 1
                adoComm.CommandText = "INSERT INTO InstrumentoInversionCalendarioTramo (CodTitulo, NumTramo, InicioTramo, FinTramo, Valor) VALUES ('" & codigoUnico & "'," & (numTramo + 1) & "," & adoRegistroTramo.Fields(1).Value & "," & adoRegistroTramo.Fields(2).Value & "," & adoRegistroTramo.Fields(3).Value & ")"
                adoConn.Execute adoComm.CommandText
                adoRegistroTramo.MoveNext
            Next

        End If
    End If
    
    If Not blnVistaPrevia Then
        MsgBox "Las condiciones financieras se grabaron satisfactoriamente.", vbInformation, Me.Caption
    End If
    
    cti_igv = chkigv.Value
    
    Call GeneraCuponera(0, codAnalitica, codigoUnico, codFondo, gstrCodAdministradora, codSolicitud, strCodSubDetalleFile, False, cti_igv)
    'Call GeneraCuponera(0, codigoUnico, codFondo, gstrCodAdministradora, CodFile, codAnalitica, codSolicitud, strCodSubDetalleFile, False, cti_igv)
    
    frmVisorCronograma.codigoUnico = codigoUnico
    frmVisorCronograma.strNumSolicitud = codSolicitud

    If (chkDesembolsosMultiples.Value = 0) Then 'And txtCantidadDesembolsos.Value > 1 Then
        frmVisorCronograma.desembMultiple = False
    Else
        frmVisorCronograma.desembMultiple = True
    End If

    frmVisorCronograma.Source = 0
    frmVisorCronograma.Show
    
    If blnVistaPrevia Then
        adoComm.CommandText = "DELETE FROM InstrumentoInversionCondicionesFinancieras WHERE CodTitulo = '" & codigoUnico & "'"
        adoConn.Execute adoComm.CommandText
    End If

    If Not blnVistaPrevia Then
        Unload Me
    End If
    
End Sub

Private Sub txtCantidadDesembolsos_Change()

    If Not flagVisor Then
        Dim i                   As Integer
        Dim cantidadDesembolsos As Integer
        cantidadDesembolsos = txtCantidadDesembolsos.Value

        If cantidadDesembolsos > 0 Then
            Set adoRegistroDesembolso = New ADODB.Recordset

            With adoRegistroDesembolso
                .CursorLocation = adUseClient
                .Fields.Append "colNum", adInteger, 999
                .Fields.Append "colFecha", adDate
                .Fields.Append "colValor", adDouble
            End With

            adoRegistroDesembolso.Open
        
            For i = 0 To cantidadDesembolsos - 1
                adoRegistroDesembolso.AddNew Array("colNum", "colFecha", "colValor"), Array(i + 1, fechaEmision, "0")
            Next

            adoRegistroDesembolso.MoveLast
            adoRegistroDesembolso.Fields.Item(2).Value = valorNominal
            Dim strMsgError As String
            mostrarDatosGridSQL dgvDesembolsos, adoRegistroDesembolso, strMsgError, "colNum"
        Else
            Set dgvDesembolsos.DataSource = Nothing
        End If
    End If

End Sub

Private Sub txtCantidadTramos_Change()

    If Not flagVisor Then
        Dim i              As Integer
        Dim cantidadTramos As Integer
        
        If txtCantidadTramos.Text = Valor_Caracter Then
            txtCantidadTramos.Text = "0"
        End If
        
        cantidadTramos = txtCantidadTramos.Value

        If cantidadTramos > 0 Then
            If cantidadTramos < 1 Then
                txtCantidadTramos.Text = 1
            ElseIf cantidadTramos > txtCuotas.Value Then
                txtCantidadTramos.Text = txtCuotas.Value
            Else
                Set adoRegistroTramo = New ADODB.Recordset

                With adoRegistroTramo
                    .CursorLocation = adUseClient
                    .Fields.Append "colNum", adInteger, 999
                    .Fields.Append "colInicio", adInteger
                    .Fields.Append "colFin", adInteger
                    .Fields.Append "colValor", adDouble
                End With

                adoRegistroTramo.Open
            
                For i = 0 To cantidadTramos - 1
                    adoRegistroTramo.AddNew Array("colNum", "colInicio", "colFin", "colValor"), Array(i + 1, i + 1, i + 1, 0)
                Next

                Dim strMsgError As String
                mostrarDatosGridSQL dgvTramos, adoRegistroTramo, strMsgError, "colNum"
                
                adoRegistroTramo.MoveLast
                adoRegistroTramo.Fields.Item(1).Value = cantCupones
                adoRegistroTramo.Fields.Item(2).Value = cantCupones
        
                If dgvTramos.Dataset.RecordCount > 1 Then
                    adoRegistroTramo.MovePrevious
                    adoRegistroTramo.Fields.Item(2).Value = cantCupones - 1
                ElseIf dgvTramos.Dataset.RecordCount = 1 Then
                    adoRegistroTramo.MoveFirst
                    adoRegistroTramo.Fields.Item(1).Value = 1
                End If

                adoRegistroTramo.MoveLast
                
                dgvTramos.Columns.Item(1).DisableEditor = True
                dgvTramos.Columns.Item(2).DisableEditor = True
                
            End If

        Else
            Set dgvTramos.DataSource = Nothing
        End If
    End If

End Sub

Private Sub txtCuotas_Change()

    If txtCuotas.Value = 0 Then
        txtCuotas.Text = 1
    End If

End Sub

Private Sub txtDiasMinimosCobroInteres_Change()
    Dim dayNumber As Integer
    
     Select Case cbPeriodoCupon.ListIndex

        Case 0
            dayNumber = 360

        Case 1
            dayNumber = 180

        Case 2
            dayNumber = 90

        Case 3
            dayNumber = 60

        Case 4
            dayNumber = 30

        Case 5
            dayNumber = 15

        Case 6
            dayNumber = 1

        Case 7

            Select Case cbUnidadesPeriodo.ListIndex

                Case 0
                    dayNumber = (CInt(txtUnidadesPeriodo.Text) * 360)

                Case 1
                    dayNumber = (CInt(txtUnidadesPeriodo.Text) * 180)

                Case 2
                    dayNumber = (CInt(txtUnidadesPeriodo.Text) * 90)

                Case 3
                    dayNumber = (CInt(txtUnidadesPeriodo.Text) * 60)

                Case 4
                    dayNumber = (CInt(txtUnidadesPeriodo.Text) * 30)

                Case 5
                    dayNumber = (CInt(txtUnidadesPeriodo.Text) * 15)

                Case 6
                    dayNumber = (CInt(txtUnidadesPeriodo.Text))
            End Select
    End Select
    
    If txtDiasMinimosCobroInteres.Value <> Valor_Caracter Then
        If txtDiasMinimosCobroInteres.Value > dayNumber Then
            txtDiasMinimosCobroInteres.Text = dayNumber
        End If
    Else
        txtDiasMinimosCobroInteres.Text = "0"
    End If
End Sub


Private Sub txtUnidadesPeriodo_Change()
    If cbUnidadesPeriodo.ListIndex = 6 And txtUnidadesPeriodo.Value > 999 Then
        txtUnidadesPeriodo.Text = 999
    End If
    If cbUnidadesPeriodo.ListIndex = 5 And txtUnidadesPeriodo.Value > 999 Then
        txtUnidadesPeriodo.Text = 999
    End If
    
    If cbUnidadesPeriodo.ListIndex <= 4 And txtUnidadesPeriodo.Value > 99 Then
        txtUnidadesPeriodo.Text = 99
    End If
  
    If cbUnidadesPeriodo.ListIndex = 0 And txtUnidadesPeriodo.Value > 50 Then
        txtUnidadesPeriodo.Text = 50
    End If
    
    If txtUnidadesPeriodo.Value = 0 Or txtUnidadesPeriodo.Text = "" Then
        txtUnidadesPeriodo.Text = 1
    End If

    cantCupones = calcularCantidadCupones()
    txtCuotas.Text = cantCupones
End Sub

