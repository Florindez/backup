VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmOrdenDepositoBancario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes - Depósitos a Plazo"
   ClientHeight    =   8640
   ClientLeft      =   1500
   ClientTop       =   1680
   ClientWidth     =   14670
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
   Icon            =   "frmOrdenDepositoBancario.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8640
   ScaleWidth      =   14670
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   12720
      TabIndex        =   208
      Top             =   7830
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
      TabIndex        =   207
      Top             =   7830
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
   Begin TabDlg.SSTab tabRFCortoPlazo 
      Height          =   7725
      Left            =   0
      TabIndex        =   71
      Top             =   30
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   13626
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmOrdenDepositoBancario.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(1)=   "fraCriterio"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datos Orden Inversión"
      TabPicture(1)   =   "frmOrdenDepositoBancario.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblDescrip(85)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraResumen"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtObservacion"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fraDatosBasicos"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fraDatosTitulo"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "fraDatosNegociacion"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "fraComisionMontoFL1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdAccion"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Negociación"
      TabPicture(2)   =   "frmOrdenDepositoBancario.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraPosicion"
      Tab(2).Control(1)=   "fraComisionMontoFL2"
      Tab(2).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   11400
         TabIndex        =   206
         Top             =   6600
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
      Begin VB.Frame fraComisionMontoFL1 
         Caption         =   "Montos Totales"
         Height          =   2250
         Left            =   5010
         TabIndex        =   169
         Top             =   4020
         Width           =   6795
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
            Left            =   2745
            MaxLength       =   45
            TabIndex        =   179
            Top             =   2310
            Width           =   1340
         End
         Begin VB.CommandButton cmdCalculo 
            Caption         =   "#"
            Height          =   375
            Left            =   540
            TabIndex        =   178
            ToolTipText     =   "Calcular Valor al Vencimiento y TIRs de la orden"
            Top             =   1440
            Width           =   375
         End
         Begin VB.CheckBox chkAplicar 
            Caption         =   "Aplicar"
            Height          =   255
            Index           =   0
            Left            =   2100
            TabIndex        =   177
            Top             =   2430
            Width           =   975
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
            Left            =   11550
            MaxLength       =   45
            TabIndex        =   176
            Top             =   3360
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
            Left            =   11550
            MaxLength       =   45
            TabIndex        =   175
            Top             =   1290
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
            Left            =   11550
            MaxLength       =   45
            TabIndex        =   174
            Top             =   1605
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
            Left            =   11550
            MaxLength       =   45
            TabIndex        =   173
            Top             =   1935
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
            Left            =   11550
            MaxLength       =   45
            TabIndex        =   172
            Top             =   2265
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
            Index           =   0
            Left            =   11550
            MaxLength       =   45
            TabIndex        =   171
            Top             =   2595
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
            Index           =   0
            Left            =   9885
            MaxLength       =   45
            TabIndex        =   170
            Top             =   1290
            Width           =   1340
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
            Index           =   24
            Left            =   510
            TabIndex        =   204
            Top             =   2325
            Width           =   705
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Total"
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
            Left            =   2640
            TabIndex        =   203
            Top             =   735
            Width           =   360
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
            Index           =   37
            Left            =   2640
            TabIndex        =   202
            Top             =   3600
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
            TabIndex        =   201
            Top             =   240
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
            TabIndex        =   200
            Tag             =   "0.00"
            Top             =   1470
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
            Left            =   1110
            TabIndex        =   199
            Tag             =   "0.00"
            Top             =   1470
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
            Height          =   285
            Left            =   2655
            TabIndex        =   198
            Tag             =   "0.00"
            Top             =   1470
            Width           =   1335
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
            TabIndex        =   197
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
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   1
            X1              =   390
            X2              =   6330
            Y1              =   3165
            Y2              =   3165
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
            TabIndex        =   196
            Tag             =   "0.00"
            ToolTipText     =   "Monto Total de la Orden"
            Top             =   3585
            Width           =   2025
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
            Index           =   38
            Left            =   4470
            TabIndex        =   195
            Top             =   1200
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
            Index           =   39
            Left            =   1350
            TabIndex        =   194
            Top             =   1200
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
            Index           =   61
            Left            =   2910
            TabIndex        =   193
            Top             =   1200
            Width           =   660
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   2
            X1              =   360
            X2              =   6300
            Y1              =   3930
            Y2              =   3930
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
            Left            =   9885
            TabIndex        =   192
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Conasev"
            Top             =   2595
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
            Left            =   9885
            TabIndex        =   191
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Fondo Liquidación"
            Top             =   2265
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
            Left            =   9885
            TabIndex        =   190
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Cavali"
            Top             =   1935
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
            Index           =   0
            Left            =   9885
            TabIndex        =   189
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión BVL"
            Top             =   1605
            Width           =   1335
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
            Left            =   9885
            TabIndex        =   188
            Tag             =   "0.00"
            ToolTipText     =   "Porcentaje de Comisión IGV"
            Top             =   2925
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
            Index           =   0
            Left            =   11550
            TabIndex        =   187
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   2925
            Width           =   2025
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
            Index           =   36
            Left            =   9900
            TabIndex        =   186
            Top             =   3375
            Width           =   1020
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
            Left            =   7650
            TabIndex        =   185
            Top             =   2985
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
            Index           =   33
            Left            =   7650
            TabIndex        =   184
            Top             =   2655
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
            Index           =   32
            Left            =   7650
            TabIndex        =   183
            Top             =   2310
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
            Index           =   27
            Left            =   7650
            TabIndex        =   182
            Top             =   1980
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
            Index           =   26
            Left            =   7650
            TabIndex        =   181
            Top             =   1635
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
            Index           =   25
            Left            =   7650
            TabIndex        =   180
            Top             =   1305
            Width           =   990
         End
      End
      Begin VB.Frame fraDatosNegociacion 
         Caption         =   "Negociación"
         Height          =   2295
         Left            =   360
         TabIndex        =   148
         Top             =   3990
         Width           =   4365
         Begin VB.TextBox txtValorNominal 
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
            Left            =   1920
            MaxLength       =   45
            TabIndex        =   154
            Top             =   1455
            Width           =   1900
         End
         Begin VB.TextBox txtTasa 
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
            Left            =   1920
            MaxLength       =   45
            TabIndex        =   153
            Top             =   1110
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
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   152
            Top             =   735
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
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   151
            Top             =   360
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
            Left            =   6840
            MaxLength       =   45
            TabIndex        =   150
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox txtCantidad 
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
            Left            =   1920
            MaxLength       =   45
            TabIndex        =   149
            Top             =   1800
            Width           =   1900
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
            TabIndex        =   168
            Top             =   1470
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
            Left            =   6840
            TabIndex        =   167
            Tag             =   "0.00"
            ToolTipText     =   "Días de Plazo del Título de la Orden"
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lblFechaVencimiento 
            Alignment       =   2  'Center
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
            Left            =   6840
            TabIndex        =   166
            Tag             =   "0.00"
            ToolTipText     =   "Fecha de Vencimiento del Título de la Orden"
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label lblFechaEmision 
            Alignment       =   2  'Center
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
            Left            =   6840
            TabIndex        =   165
            Tag             =   "0.00"
            ToolTipText     =   "Fecha Emisión"
            Top             =   720
            Width           =   1815
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000015&
            X1              =   4395
            X2              =   4395
            Y1              =   360
            Y2              =   2145
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
            TabIndex        =   164
            Top             =   1125
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
            Left            =   375
            TabIndex        =   163
            Top             =   750
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
            Left            =   375
            TabIndex        =   162
            Top             =   375
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
            Left            =   4920
            TabIndex        =   161
            Top             =   1820
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
            TabIndex        =   160
            Top             =   1815
            Width           =   1095
         End
         Begin VB.Label lblFechaLiquidacion 
            Alignment       =   2  'Center
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
            Left            =   6840
            TabIndex        =   159
            Tag             =   "0.00"
            ToolTipText     =   "Fecha de Liquidación"
            Top             =   360
            Width           =   1815
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
            Index           =   18
            Left            =   4920
            TabIndex        =   158
            Top             =   380
            Width           =   810
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
            Index           =   47
            Left            =   4920
            TabIndex        =   157
            Top             =   740
            Width           =   540
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
            Left            =   4920
            TabIndex        =   156
            Top             =   1100
            Width           =   870
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
            Left            =   4920
            TabIndex        =   155
            Top             =   1460
            Width           =   870
         End
      End
      Begin VB.Frame fraCriterio 
         Caption         =   "Criterios de Búsqueda"
         Height          =   2535
         Left            =   -74640
         TabIndex        =   134
         Top             =   540
         Width           =   13935
         Begin VB.CommandButton cmdExportarExcel 
            Caption         =   "Excel"
            Height          =   735
            Left            =   10560
            Picture         =   "frmOrdenDepositoBancario.frx":0496
            Style           =   1  'Graphical
            TabIndex        =   205
            Top             =   1530
            Width           =   1200
         End
         Begin VB.CommandButton cmdEnviar 
            Caption         =   "En&viar"
            Height          =   735
            Left            =   12200
            Picture         =   "frmOrdenDepositoBancario.frx":0A9E
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Enviar a BackOffice"
            Top             =   1530
            Width           =   1200
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
            Top             =   510
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
            Top             =   1560
            Width           =   5145
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
            Top             =   990
            Width           =   5145
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenDesde 
            Height          =   285
            Left            =   9600
            TabIndex        =   3
            Top             =   600
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
            Format          =   293404673
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenHasta 
            Height          =   285
            Left            =   11955
            TabIndex        =   4
            Top             =   600
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
            Format          =   293404673
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionDesde 
            Height          =   285
            Left            =   9600
            TabIndex        =   5
            Top             =   1020
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
            Format          =   293404673
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionHasta 
            Height          =   285
            Left            =   11955
            TabIndex        =   6
            Top             =   1020
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
            Format          =   293404673
            CurrentDate     =   38785
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
            TabIndex        =   143
            Top             =   1035
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
            Index           =   45
            Left            =   8880
            TabIndex        =   142
            Top             =   1035
            Width           =   465
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
            TabIndex        =   141
            Top             =   1035
            Width           =   1305
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
            TabIndex        =   140
            Top             =   615
            Width           =   930
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
            Left            =   360
            TabIndex        =   139
            Top             =   525
            Width           =   450
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
            TabIndex        =   138
            Top             =   615
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
            Index           =   21
            Left            =   11280
            TabIndex        =   137
            Top             =   615
            Width           =   420
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
            Left            =   360
            TabIndex        =   136
            Top             =   1560
            Width           =   825
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
            Left            =   360
            TabIndex        =   135
            Top             =   1020
            Width           =   495
         End
      End
      Begin VB.Frame fraDatosTitulo 
         Caption         =   "Datos de la Orden"
         Height          =   1425
         Left            =   360
         TabIndex        =   126
         Top             =   2460
         Width           =   13935
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
            Left            =   11760
            MaxLength       =   15
            TabIndex        =   30
            Top             =   615
            Width           =   1575
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
            Height          =   285
            Left            =   1680
            MaxLength       =   45
            TabIndex        =   24
            Top             =   1005
            Width           =   5040
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   630
            Width           =   5085
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
            Height          =   285
            Left            =   8280
            TabIndex        =   26
            Top             =   615
            Width           =   1275
         End
         Begin MSComCtl2.DTPicker dtpFechaOrden 
            Height          =   285
            Left            =   1680
            TabIndex        =   21
            Top             =   300
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   293404673
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacion 
            Height          =   285
            Left            =   5160
            TabIndex        =   22
            Top             =   300
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   293404673
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaEmision 
            Height          =   285
            Left            =   8280
            TabIndex        =   25
            Top             =   285
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   293404673
            CurrentDate     =   38776
         End
         Begin MSComCtl2.UpDown updDiasPlazo 
            Height          =   285
            Left            =   9555
            TabIndex        =   27
            Top             =   615
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtDiasPlazo"
            BuddyDispid     =   196664
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
            Height          =   285
            Left            =   11760
            TabIndex        =   29
            Top             =   285
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   293404673
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaPago 
            Height          =   285
            Left            =   8280
            TabIndex        =   28
            Top             =   1005
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   293404673
            CurrentDate     =   38776
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
            Left            =   6960
            TabIndex        =   146
            Top             =   1035
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
            Index           =   30
            Left            =   10200
            TabIndex        =   144
            Top             =   630
            Width           =   720
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   360
            TabIndex        =   133
            Top             =   315
            Width           =   930
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
            Left            =   3960
            TabIndex        =   132
            Top             =   315
            Width           =   810
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
            TabIndex        =   131
            Top             =   1035
            Width           =   840
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
            Left            =   10200
            TabIndex        =   130
            Top             =   300
            Width           =   870
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
            Index           =   10
            Left            =   6960
            TabIndex        =   129
            Top             =   630
            Width           =   870
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
            TabIndex        =   128
            Top             =   660
            Width           =   585
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
            Left            =   6960
            TabIndex        =   127
            Top             =   300
            Width           =   540
         End
      End
      Begin VB.Frame fraDatosBasicos 
         Caption         =   "Datos Básicos"
         Height          =   1860
         Left            =   360
         TabIndex        =   115
         Top             =   540
         Width           =   13935
         Begin VB.CheckBox chkTitulo 
            Height          =   255
            Left            =   13530
            TabIndex        =   16
            ToolTipText     =   "Seleccionar Título"
            Top             =   2550
            Width           =   255
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
            Left            =   8940
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   2610
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
            TabIndex        =   20
            Top             =   1170
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1200
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   780
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   360
            Width           =   4185
         End
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
            Left            =   7140
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2970
            Width           =   4185
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
            Left            =   9315
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   690
            Width           =   4185
         End
         Begin VB.ComboBox cboAgente 
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
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2670
            Width           =   4185
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
            Left            =   7545
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   3390
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
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   3000
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
            Left            =   9315
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   270
            Width           =   4185
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Entidad Financiera"
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
            Left            =   6960
            TabIndex        =   125
            Top             =   315
            Width           =   1320
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
            Left            =   6960
            TabIndex        =   124
            Top             =   1200
            Width           =   1575
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
            TabIndex        =   123
            Top             =   1260
            Width           =   660
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
            TabIndex        =   122
            Top             =   840
            Width           =   825
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
            TabIndex        =   121
            Top             =   405
            Width           =   450
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
            Left            =   7620
            TabIndex        =   120
            Top             =   2640
            Width           =   390
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
            Left            =   5430
            TabIndex        =   119
            Top             =   3060
            Width           =   1590
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Mecanismo Negociación"
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
            Left            =   6960
            TabIndex        =   118
            Top             =   750
            Width           =   1755
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Agente"
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
            Left            =   270
            TabIndex        =   117
            Top             =   2700
            Width           =   510
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
            Index           =   5
            Left            =   5190
            TabIndex        =   116
            Top             =   3420
            Width           =   1530
         End
      End
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
         Height          =   675
         Left            =   360
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   66
         Top             =   6720
         Width           =   9960
      End
      Begin VB.Frame fraResumen 
         Caption         =   "Resumen Negociación"
         Height          =   405
         Left            =   90
         TabIndex        =   41
         Top             =   7860
         Width           =   13935
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
            TabIndex        =   42
            Tag             =   "0.00"
            Top             =   360
            Width           =   2025
         End
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
            Height          =   195
            Index           =   84
            Left            =   10200
            TabIndex        =   43
            Top             =   380
            Width           =   630
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000015&
            X1              =   9720
            X2              =   9720
            Y1              =   360
            Y2              =   2640
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
            TabIndex        =   44
            Top             =   720
            Width           =   1845
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
            TabIndex        =   45
            Top             =   720
            Width           =   2025
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
            TabIndex        =   46
            Tag             =   "0.00"
            Top             =   360
            Width           =   2025
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
            Height          =   195
            Index           =   16
            Left            =   5280
            TabIndex        =   47
            Top             =   380
            Width           =   870
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
            Height          =   195
            Index           =   42
            Left            =   480
            TabIndex        =   48
            Top             =   375
            Width           =   1095
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
            TabIndex        =   49
            Tag             =   "0.00"
            Top             =   360
            Width           =   2025
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000015&
            X1              =   4800
            X2              =   4800
            Y1              =   360
            Y2              =   2640
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
            TabIndex        =   50
            Tag             =   "0.00"
            Top             =   1365
            Width           =   2025
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
            TabIndex        =   51
            Tag             =   "0.00"
            Top             =   1035
            Width           =   2025
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
            Height          =   195
            Index           =   82
            Left            =   10200
            TabIndex        =   52
            Top             =   1380
            Width           =   570
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
            Height          =   195
            Index           =   78
            Left            =   10200
            TabIndex        =   53
            Top             =   1065
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
            Height          =   195
            Index           =   77
            Left            =   5280
            TabIndex        =   54
            Top             =   720
            Width           =   390
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
            Height          =   195
            Index           =   75
            Left            =   480
            TabIndex        =   55
            Top             =   720
            Width           =   600
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
            TabIndex        =   56
            Tag             =   "0.00"
            Top             =   2355
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
            Left            =   7320
            TabIndex        =   57
            Tag             =   "0.00"
            Top             =   2025
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
            Left            =   7320
            TabIndex        =   58
            Tag             =   "0.00"
            Top             =   1695
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
            Left            =   7320
            TabIndex        =   59
            Tag             =   "0.00"
            Top             =   1365
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
            TabIndex        =   60
            Tag             =   "0.00"
            Top             =   1035
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
            TabIndex        =   61
            Tag             =   "0.00"
            Top             =   2355
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
            Left            =   2400
            TabIndex        =   62
            Tag             =   "0.00"
            Top             =   2025
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
            Left            =   2400
            TabIndex        =   63
            Tag             =   "0.00"
            Top             =   1695
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
            Left            =   2400
            TabIndex        =   64
            Tag             =   "0.00"
            Top             =   1365
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
            Index           =   0
            Left            =   2400
            TabIndex        =   65
            Tag             =   "0.00"
            Top             =   1035
            Width           =   2025
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
            Height          =   195
            Index           =   74
            Left            =   5280
            TabIndex        =   67
            Top             =   2375
            Width           =   855
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
            Index           =   71
            Left            =   5280
            TabIndex        =   68
            Top             =   2045
            Width           =   1260
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
            Index           =   70
            Left            =   5280
            TabIndex        =   69
            Top             =   1715
            Width           =   795
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
            Index           =   69
            Left            =   5280
            TabIndex        =   70
            Top             =   1385
            Width           =   645
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
            Height          =   195
            Index           =   68
            Left            =   480
            TabIndex        =   72
            Top             =   2375
            Width           =   855
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
            Index           =   67
            Left            =   480
            TabIndex        =   73
            Top             =   2045
            Width           =   1260
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
            Index           =   66
            Left            =   480
            TabIndex        =   74
            Top             =   1715
            Width           =   795
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
            Index           =   65
            Left            =   480
            TabIndex        =   75
            Top             =   1385
            Width           =   645
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
            Height          =   195
            Index           =   64
            Left            =   480
            TabIndex        =   76
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
            Height          =   195
            Index           =   63
            Left            =   5280
            TabIndex        =   77
            Top             =   1055
            Width           =   450
         End
      End
      Begin VB.Frame fraComisionMontoFL2 
         Caption         =   "Comisiones y Montos - Plazo (FL2)"
         Height          =   2070
         Left            =   -67560
         TabIndex        =   89
         Top             =   2775
         Visible         =   0   'False
         Width           =   6735
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
            TabIndex        =   31
            Top             =   360
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
            Index           =   1
            Left            =   2610
            MaxLength       =   45
            TabIndex        =   33
            Top             =   1170
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
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   38
            Top             =   2490
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
            TabIndex        =   37
            Top             =   2160
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
            TabIndex        =   36
            Top             =   1830
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
            TabIndex        =   35
            Top             =   1500
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
            TabIndex        =   34
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
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   39
            Top             =   3255
            Width           =   2025
         End
         Begin VB.CommandButton Command1 
            Caption         =   "#"
            Height          =   375
            Left            =   480
            TabIndex        =   40
            ToolTipText     =   "Calcular Valor al Vencimiento y TIRs de la orden"
            Top             =   4185
            Width           =   375
         End
         Begin VB.CheckBox chkAplicar 
            Caption         =   "Aplicar"
            Height          =   255
            Index           =   1
            Left            =   390
            TabIndex        =   32
            Top             =   720
            Width           =   975
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
            TabIndex        =   114
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
            Index           =   56
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
            Index           =   57
            Left            =   390
            TabIndex        =   112
            Top             =   1530
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
            Index           =   58
            Left            =   390
            TabIndex        =   111
            Top             =   1875
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
            Index           =   59
            Left            =   390
            TabIndex        =   110
            Top             =   2205
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
            Index           =   60
            Left            =   390
            TabIndex        =   109
            Top             =   2550
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
            Index           =   62
            Left            =   390
            TabIndex        =   108
            Top             =   2880
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
            Index           =   72
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
            Index           =   73
            Left            =   2640
            TabIndex        =   106
            Top             =   3270
            Width           =   1020
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
            TabIndex        =   105
            Top             =   3600
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
            Index           =   1
            Left            =   4365
            TabIndex        =   104
            Top             =   240
            Width           =   1845
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
            TabIndex        =   103
            Tag             =   "0.00"
            Top             =   4290
            Width           =   2025
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
            TabIndex        =   102
            Tag             =   "0.00"
            Top             =   4290
            Width           =   1335
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
            TabIndex        =   101
            Tag             =   "0.00"
            Top             =   4290
            Width           =   1335
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
            TabIndex        =   100
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
            TabIndex        =   99
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   2820
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
            Index           =   1
            Left            =   2625
            TabIndex        =   98
            Tag             =   "0.00"
            ToolTipText     =   "Porcentaje de Comisión IGV"
            Top             =   2820
            Width           =   1335
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   4
            X1              =   390
            X2              =   6330
            Y1              =   3165
            Y2              =   3165
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
            TabIndex        =   97
            Tag             =   "0.00"
            ToolTipText     =   "Monto Total de la Orden"
            Top             =   3585
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
            Index           =   1
            Left            =   2625
            TabIndex        =   96
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión BVL"
            Top             =   1500
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
            TabIndex        =   95
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Cavali"
            Top             =   1830
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
            TabIndex        =   94
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Fondo Liquidación"
            Top             =   2160
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
            Index           =   1
            Left            =   2625
            TabIndex        =   93
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Conasev"
            Top             =   2490
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
            Index           =   79
            Left            =   4440
            TabIndex        =   92
            Top             =   4020
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
            Index           =   80
            Left            =   1320
            TabIndex        =   91
            Top             =   4020
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
            Index           =   81
            Left            =   2880
            TabIndex        =   90
            Top             =   4020
            Width           =   660
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   5
            X1              =   390
            X2              =   6330
            Y1              =   3930
            Y2              =   3930
         End
      End
      Begin VB.Frame fraPosicion 
         Caption         =   "Datos Posición"
         Height          =   2295
         Left            =   -65520
         TabIndex        =   78
         Top             =   480
         Width           =   4695
         Begin VB.Label lblTasaCupon 
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
            Left            =   2910
            TabIndex        =   147
            Tag             =   "0.00"
            Top             =   1080
            Width           =   1275
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
            TabIndex        =   88
            Tag             =   "0.00"
            ToolTipText     =   "Moneda del Título"
            Top             =   1800
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
            TabIndex        =   87
            Tag             =   "0.00"
            Top             =   1440
            Width           =   2025
         End
         Begin VB.Label lblBaseCalculo 
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
            TabIndex        =   86
            Tag             =   "0.00"
            Top             =   1080
            Width           =   705
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
            TabIndex        =   85
            Tag             =   "0.00"
            Top             =   720
            Width           =   2025
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
            TabIndex        =   84
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   54
            Left            =   480
            TabIndex        =   83
            Top             =   1820
            Width           =   585
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
            Index           =   53
            Left            =   480
            TabIndex        =   82
            Top             =   1460
            Width           =   1035
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
            TabIndex        =   81
            Top             =   1100
            Width           =   1020
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
            Index           =   49
            Left            =   480
            TabIndex        =   80
            Top             =   740
            Width           =   885
         End
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
            TabIndex        =   79
            Top             =   380
            Width           =   1050
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmOrdenDepositoBancario.frx":0FF9
         Height          =   3975
         Left            =   -74640
         OleObjectBlob   =   "frmOrdenDepositoBancario.frx":1013
         TabIndex        =   8
         Top             =   3180
         Width           =   13905
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
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
         Left            =   360
         TabIndex        =   145
         Top             =   6420
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmOrdenDepositoBancario"
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
Dim arrBaseAnual()          As String, arrTipoTasa()                As String
Dim arrOrigen()             As String, arrClaseInstrumento()        As String
Dim arrTitulo()             As String, arrSubClaseInstrumento()     As String
Dim arrAgente()             As String, arrConceptoCosto()           As String
Dim strCodFondo             As String, strCodFondoOrden             As String
Dim strCodTipoInstrumento   As String, strCodTipoInstrumentoOrden   As String
Dim strCodEstado            As String, strCodTipoOrden              As String
Dim strCodOperacion         As String, strCodNegociacion            As String
Dim strCodEmisor            As String, strCodMoneda                 As String
Dim strCodBaseAnual         As String, strCodTipoTasa               As String
Dim strCodOrigen            As String, strCodClaseInstrumento       As String
Dim strCodTitulo            As String, strCodSubClaseInstrumento    As String
Dim strCodAgente            As String, strIndPacto                  As String
Dim strIndNegociable        As String, strCodigosFile               As String
Dim strCodReportado         As String, strCodGarantia               As String
Dim strCodConcepto          As String
Dim strEstado               As String, strSQL                       As String

Dim strCodFile              As String, strCodAnalitica              As String
Dim strCodGrupo             As String, strCodCiiu                   As String
Dim strEstadoOrden          As String, strCodCategoria              As String
Dim strCodRiesgo            As String, strCodSubRiesgo              As String
Dim strCalcVcto             As String, strCodSector                 As String
Dim strCodTipoCostoBolsa    As String, strCodTipoCostoConasev       As String
Dim strCodTipoCostoFondo    As String, strCodTipoCavali             As String
Dim strIndCuponCero         As String, strCodIndiceFinal            As String
Dim strCodTipoAjuste        As String, strCodPeriodoPago            As String
Dim strCodIndiceInicial     As String
Dim dblTipoCambio           As Double, dblTasaCuponNormal           As Double
Dim dblComisionBolsa        As Double, dblComisionConasev           As Double
Dim dblComisionFondo        As Double, dblComisionCavali            As Double
Dim intBaseCalculo          As Integer
Public oExportacion As clsExportacion
Public indOk As Boolean
Dim adoExportacion As ADODB.Recordset
Dim adoConsulta             As ADODB.Recordset
Dim indSortAsc              As Boolean, indSortDesc                 As Boolean
Dim strCodCobroInteres As String
Dim strMontoInteres As Double

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
Public Sub Adicionar()
    
'    If Not EsDiaUtil(gdatFechaActual) Then
'        MsgBox "No se puede negociar en un día no útil !", vbCritical, Me.Caption
'        Exit Sub
'    End If
            
    If cboTipoInstrumento.ListCount > 1 Then
        frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Orden..."
                    
        strEstado = Reg_Adicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabRFCortoPlazo
            .TabEnabled(0) = False
            .TabEnabled(2) = False
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

'Private Sub CalcularTirBruta()
'
'    Dim dblTasaCalculada As Double
'
'    If CDbl(txtPrecioUnitario(0).Text) = 0 Then
'        MsgBox "Por favor ingrese el Precio.", vbCritical, Me.Caption
'        Exit Sub
'    End If
'
'    Me.MousePointer = vbHourglass
'
'    If CDbl(txtPrecioUnitario(0).Text) > 0 Then
'        ReDim Array_Monto(1): ReDim Array_Dias(1)
'
'        Array_Monto(0) = CDec(txtCantidad.Text) * -1 'CDec((CCur(lblSubTotal(0).Caption) + txtInteresCorrido(0).Text) * -1)
'        'Array_Dias(0) = dtpFechaLiquidacion.Value
'        Array_Dias(0) = dtpFechaEmision.Value
'
'        If strCodBaseAnual = Codigo_Base_Actual_Actual Or strCodBaseAnual = Codigo_Base_30_365 Or strCodBaseAnual = Codigo_Base_Actual_365 Then
'            If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'                dblTasaCalculada = ((1 + (CDbl(txtTasa.Text) / 100)) ^ (CDbl(txtDiasPlazo) / 365)) - 1
'            Else
'                dblTasaCalculada = (1 + ((CDbl(txtTasa.Text) / 100 / 365) * CDbl(txtDiasPlazo))) - 1
'            End If
'        Else
'            If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'                dblTasaCalculada = ((1 + (CDbl(txtTasa.Text) / 100)) ^ (CDbl(txtDiasPlazo) / 360)) - 1
'            Else
'                dblTasaCalculada = (1 + ((CDbl(txtTasa.Text) / 100 / 360) * CDbl(txtDiasPlazo))) - 1
'            End If
'        End If
'
'        If strCalcVcto = "D" Then
'            Array_Monto(1) = CDec(txtCantidad.Text)
'        Else
'            'Array_Monto(1) = CDbl(txtCantidad.Text) * (1 + dblTasaCalculada)
'            Array_Monto(1) = CDec((CCur(lblSubTotal(0).Caption) + txtInteresCorrido(0).Text))
'        End If
'
'        Array_Dias(1) = dtpFechaVencimiento.Value
'        lblTirBruta.Caption = CStr(TIR(Array_Monto(), Array_Dias(), 1) * 100)
'        lblTirBrutaResumen.Caption = lblTirBruta.Caption
'        If strCodTipoOrden = Codigo_Orden_Pacto Then lblTirBrutaResumen.Caption = "0"
'    End If
'    Me.MousePointer = vbDefault
'
'End Sub
Private Sub CalcularTirBruta()

    Dim dblTasaCalculada As Double

    If CDbl(txtPrecioUnitario(0).Text) = 0 Then
        MsgBox "Por favor ingrese el Precio.", vbCritical, Me.Caption
        Exit Sub
    End If

    Me.MousePointer = vbHourglass

    If CDbl(txtPrecioUnitario(0).Text) > 0 Then
        ReDim Array_Monto(1): ReDim Array_Dias(1)
        
'        Array_Monto(0) = CDec(txtCantidad.Text) * -1 'CDec((CCur(lblSubTotal(0).Caption) + txtInteresCorrido(0).Text) * -1)
        Array_Monto(0) = CDec(CCur(lblSubTotal(0).Caption) * -1)
'        Array_Dias(0) = dtpFechaLiquidacion.Value
        Array_Dias(0) = dtpFechaEmision.Value
        
        If strCodBaseAnual = Codigo_Base_Actual_Actual Or strCodBaseAnual = Codigo_Base_30_365 Or strCodBaseAnual = Codigo_Base_Actual_365 Then
            If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                dblTasaCalculada = ((1 + (CDbl(txtTasa.Text) / 100)) ^ (CDbl(txtDiasPlazo) / 365)) - 1
            Else
                dblTasaCalculada = (1 + ((CDbl(txtTasa.Text) / 100 / 365) * CDbl(txtDiasPlazo))) - 1
            End If
        Else
            If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                dblTasaCalculada = ((1 + (CDbl(txtTasa.Text) / 100)) ^ (CDbl(txtDiasPlazo) / 360)) - 1
            Else
                dblTasaCalculada = (1 + ((CDbl(txtTasa.Text) / 100 / 360) * CDbl(txtDiasPlazo))) - 1
            End If
        End If
    
        If strCalcVcto = "D" Then
            Array_Monto(1) = CCur(lblSubTotal(0).Caption)  'CDec(txtCantidad.Text)
        Else
            Array_Monto(1) = Round(CCur(lblSubTotal(0).Caption) * (1 + dblTasaCalculada), 2)
'            Array_Monto(1) = CDec((CCur(lblSubTotal(0).Caption) + txtInteresCorrido(0).Text))
        End If
        
        Array_Dias(1) = dtpFechaVencimiento.Value
        lblTirBruta.Caption = CStr(TIR(Array_Monto(), Array_Dias(), 1) * 100)
        lblTirBrutaResumen.Caption = lblTirBruta.Caption
        If strCodTipoOrden = Codigo_Orden_Pacto Then lblTirBrutaResumen.Caption = "0"
    End If
    Me.MousePointer = vbDefault

End Sub
Private Sub CalcularTirNeta()

    Dim dblTir As Double
    Dim dblTasaCalculada As Double
    Dim curComImp As Currency

    If CDbl(lblSubTotal(0).Caption) <= 0 Then
        MsgBox "Por favor ingrese los datos necesarios para hallar la TIR Neta", vbCritical, Me.Caption
        Exit Sub
    End If

    Me.MousePointer = vbHourglass
        
    curComImp = CCur(txtComisionAgente(0).Text) + CCur(txtComisionBolsa(0).Text) + CCur(txtComisionConasev(0).Text) + CCur(txtComisionCavali(0).Text) + CCur(txtComisionFondo(0).Text) + CCur(lblComisionIgv(0).Caption)
           
    ReDim Array_Monto(1): ReDim Array_Dias(1)

    If strCodTipoOrden = Codigo_Orden_Compra Or strCodTipoOrden = Codigo_Orden_Pacto Then
        Array_Monto(0) = CDec((CCur(lblSubTotal(0).Caption) + CCur(txtInteresCorrido(0).Text) + curComImp) * -1)
    ElseIf strCodTipoOrden = Codigo_Orden_Venta Or strCodTipoOrden = Codigo_Orden_Quiebre Then
        Array_Monto(0) = CCur(lblSubTotal(0).Caption) * -1 'CDec((CCur(lblSubTotal(0).Caption) + CCur(txtInteresCorrido(0).Text) + CCur(txtComisionAgente(0).Text) + CCur(txtComisionBolsa(0).Text) + CCur(txtComisionConasev(0).Text) + CCur(lblComisionIgv(0).Caption)) * -1)
    End If
    
    Array_Dias(0) = dtpFechaEmision.Value
    
    If strCodBaseAnual = Codigo_Base_Actual_Actual Or strCodBaseAnual = Codigo_Base_30_365 Or strCodBaseAnual = Codigo_Base_Actual_365 Then
        If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
            dblTasaCalculada = ((1 + (CDbl(txtTasa.Text) / 100)) ^ (CDbl(txtDiasPlazo) / 365)) - 1
        Else
            dblTasaCalculada = (1 + ((CDbl(txtTasa.Text) / 100 / 365) * CDbl(txtDiasPlazo))) - 1
        End If
    Else
        If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
            dblTasaCalculada = ((1 + (CDbl(txtTasa.Text) / 100)) ^ (CDbl(txtDiasPlazo) / 360)) - 1
        Else
            dblTasaCalculada = (1 + ((CDbl(txtTasa.Text) / 100 / 360) * CDbl(txtDiasPlazo))) - 1
        End If
    End If

    If strCalcVcto = "D" Then
        Array_Monto(1) = CCur(lblSubTotal(0).Caption)  'CDec(txtCantidad.Text)
    Else
        If strCodTipoOrden = Codigo_Orden_Compra Or strCodTipoOrden = Codigo_Orden_Pacto Then
            Array_Monto(1) = Round(CCur(lblSubTotal(0).Caption) * (1 + dblTasaCalculada), 2) 'ACR
            'Array_Monto(1) = CDec((CCur(lblSubTotal(0).Caption) + curComImp))       'ACR
        ElseIf strCodTipoOrden = Codigo_Orden_Venta Or strCodTipoOrden = Codigo_Orden_Quiebre Then '*** Venta ***
            'Array_Monto(1) = CDec(CDbl(txtCantidad.Text) * (1 + dblTasaCalculada)) 'ACR
            Array_Monto(1) = CDec((CCur(lblSubTotal(0).Caption) + CCur(txtInteresCorrido(0).Text) - curComImp))       'ACR
        End If
    End If
    Array_Dias(1) = dtpFechaVencimiento.Value
    
    lblTirNeta.Caption = CStr(TIR(Array_Monto(), Array_Dias(), 1) * 100)
'        dblTir = TIR(Array_Monto(), Array_Dias(), (10 / 100)) * 100

'        lblTirlblTirNetaNeta.Caption = CStr(dblTir)
    lblTirNetaResumen.Caption = lblTirNeta.Caption
'        lblTirNetaResumen.Caption = CStr(dblTir)
    If strCodTipoOrden = Codigo_Orden_Pacto Then lblTirNetaResumen.Caption = "0"
    Me.MousePointer = vbDefault

End Sub

Private Sub CalcularValorVencimiento()

    If DateDiff("d", dtpFechaEmision, dtpFechaVencimiento) < 0 Then
        MsgBox "La Fecha de vencimiento debe ser posterior a la Fecha de Emisión.", vbCritical, Me.Caption
        lblMontoVencimiento.Caption = "0"
    Else
        If strCalcVcto = "D" Then
            If Not IsNumeric(txtDiasPlazo.Text) Then txtDiasPlazo.Text = "0"
            lblMontoVencimiento.Caption = lblSubTotal(0).Caption
        Else
            Dim intNumDias30    As Integer
            
            '*** Hallar los días 30/360,30/365 ***
            intNumDias30 = Dias360(dtpFechaEmision.Value, dtpFechaVencimiento.Value, True)
            
            If Not IsNumeric(txtDiasPlazo.Text) Then txtDiasPlazo.Text = "0"
            lblMontoVencimiento.Caption = CStr(ValorVencimiento(CCur(lblSubTotal(0).Caption), CDbl(txtTasa.Text), intBaseCalculo, CInt(txtDiasPlazo.Text), intNumDias30, strCodTipoTasa, strCodBaseAnual)) 'ACR
            'lblMontoVencimiento.Caption = CStr(ValorVencimiento(CCur(txtCantidad.Text), CDbl(txtTasa.Text), intBaseCalculo, CInt(txtDiasPlazo.Text), intNumDias30, strCodTipoTasa, strCodBaseAnual)) 'ACR
            'lblMontoVencimiento.Caption = CStr(ValorVencimiento(CCur(txtCantidad.Text), dblTasaCuponNormal, intBaseCalculo, CInt(txtDiasPlazo.Text)))
        End If
    End If
    lblVencimientoResumen.Caption = lblMontoVencimiento.Caption

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
    Dim intRegistro As Integer
    
    Select Case strModo
        Case Reg_Adicion
            chkTitulo.Value = vbUnchecked
            
            intRegistro = ObtenerItemLista(arrFondoOrden(), strCodFondo)
            If intRegistro >= 0 Then cboFondoOrden.ListIndex = intRegistro
        
            cboAgente.ListIndex = -1
            If cboAgente.ListCount > 0 Then cboAgente.ListIndex = 0
            
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
                        
            intRegistro = ObtenerItemLista(arrOrigen(), Codigo_Negociacion_Local)
            If intRegistro >= 0 Then cboOrigen.ListIndex = intRegistro
            
            dtpFechaOrden.Value = gdatFechaActual
            dtpFechaLiquidacion.Value = dtpFechaOrden.Value
            lblFechaLiquidacion.Caption = CStr(dtpFechaOrden.Value)
            
            intRegistro = ObtenerItemLista(arrMoneda(), strCodMoneda)
            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
            
            txtDescripOrden.Text = Valor_Caracter
            txtNemonico.Text = Valor_Caracter
            txtObservacion.Text = Valor_Caracter
            txtPrecioUnitario(0).Text = "100"
            txtPrecioUnitario(1).Text = "0"
            txtValorNominal.Text = "1"
            txtCantidad.Text = "1"

            lblAnalitica.Caption = "??? - ????????"
            lblStockNominal.Caption = "0"
            lblClasificacion.Caption = Valor_Caracter

            dtpFechaEmision.Value = gdatFechaActual
            dtpFechaVencimiento.Value = dtpFechaEmision.Value
            dtpFechaPago.Value = dtpFechaVencimiento.Value
            lblFechaEmision.Caption = CStr(dtpFechaEmision.Value)
            lblFechaVencimiento.Caption = CStr(dtpFechaVencimiento.Value)
            lblVencimientoResumen.Caption = "0"
            
            txtDiasPlazo.Text = "0": lblDiasPlazo.Caption = "0"
            txtTasa.Text = "0"
            
            intRegistro = ObtenerItemLista(arrBaseAnual(), Codigo_Base_Actual_360)
            If cboBaseAnual.ListCount > 0 Then cboBaseAnual.ListIndex = intRegistro
            
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
                        
            lblFechaCupon.Caption = Valor_Caracter
            lblClasificacion.Caption = Valor_Caracter
            lblBaseCalculo.Caption = Valor_Caracter
            lblTasaCupon.Caption = Valor_Caracter
            lblStockNominal.Caption = "0"
            lblMoneda.Caption = Valor_Caracter
            lblCantidadResumen.Caption = "0"
                                                
            lblTirBrutaResumen.Caption = "0"
            lblTirNetaResumen.Caption = "0"
            
            cboFondoOrden.SetFocus
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
        
        strMensaje = "Se procederá a eliminar la ORDEN Número " & tdgConsulta.Columns(1) & vbNewLine & " por la " & _
            Trim(tdgConsulta.Columns(3)) & vbNewLine & vbNewLine & vbNewLine & "¿ Seguro de continuar ?"
        
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
                    Else
                        strCodAnalitica = NumAleatorio(8)
                        strCodTitulo = NumAleatorio(15)
                    End If
                Else
                    strIndTitulo = Valor_Indicador
                    strCodTitulo = strCodGarantia
                    strCodGarantia = Valor_Caracter
                    'strCodMoneda = lblMoneda.Tag
                    strFechaVencimiento = Convertyyyymmdd(Valor_Fecha)
                    strCodReportado = Valor_Caracter
                End If
                                                                                                    
                '.CommandText = "BEGIN TRAN ProcOrden"
                'adoConn.Execute .CommandText
                
'                .CommandText = "{ call up_IVAdicInversionOrden('" & strCodFondoOrden & "','" & _
'                    gstrCodAdministradora & "','','" & strFechaOrden & "','" & _
'                    strCodTitulo & "','" & Trim(txtNemonico.Text) & "','" & gstrPeriodoActual & "','" & gstrMesActual & "','" & _
'                    "','" & strEstadoOrden & "','" & strCodAnalitica & "','" & strCodFile & "','" & _
'                    strCodAnalitica & "','" & strCodClaseInstrumento & "','" & strCodSubClaseInstrumento & "','" & strCodTipoOrden & "','" & _
'                    strCodOperacion & "','" & strCodNegociacion & "','" & strCodOrigen & "','" & Trim(txtDescripOrden.Text) & "','" & strCodEmisor & "','" & _
'                    strCodAgente & "','" & strCodGarantia & "','','" & strFechaPago & "','" & strFechaVencimiento & "','" & strFechaLiquidacion & "','" & _
'                    strFechaEmision & "','" & strCodMoneda & "','" & strCodMoneda & "','" & strCodMoneda & "'," & CDec(txtCantidad.Text) & "," & _
'                    CDec(txtTipoCambio.Text) & "," & CDec(txtTipoCambio.Text) & "," & CDec(txtValorNominal.Text) & "," & _
'                    CDec(txtPrecioUnitario(0).Text) & "," & CDec(lblSubTotal(0).Caption) & "," & CDec(lblSubTotal(0).Caption) & "," & CDec(txtInteresCorrido(0).Text) & "," & _
'                    CDec(txtComisionAgente(0).Text) & "," & CDec(txtComisionCavali(0).Text) & "," & CDec(txtComisionConasev(0).Text) & "," & _
'                    CDec(txtComisionBolsa(0).Text) & "," & CDec(txtComisionFondo(0).Text) & ",0,0,0," & CDec(lblComisionIgv(0).Caption) & "," & CDec(lblMontoTotal(0).Caption) & "," & CDec(lblMontoTotal(0).Caption) & "," & _
'                    CDec(txtPrecioUnitario(1).Text) & "," & CDec(lblSubTotal(1).Caption) & "," & CDec(txtInteresCorrido(1).Text) & "," & CDec(txtComisionAgente(1).Text) & "," & _
'                    CDec(txtComisionCavali(1).Text) & "," & CDec(txtComisionConasev(1).Text) & "," & CDec(txtComisionBolsa(1).Text) & "," & _
'                    CDec(txtComisionFondo(1).Text) & ",0,0,0," & CDec(lblComisionIgv(1).Caption) & "," & CDec(lblMontoTotal(1).Caption) & "," & _
'                    CDec(lblMontoVencimiento.Caption) & "," & CInt(txtDiasPlazo.Text) & ",'','','','" & strCodReportado & "','" & strCodEmisor & "','" & strCodEmisor & "',0,'','','" & strIndTitulo & "','" & _
'                    strCodTipoTasa & "','" & strCodBaseAnual & "'," & CDec(txtTasa.Text) & "," & CDec(lblTirBrutaResumen.Caption) & "," & CDec(lblTirNetaResumen.Caption) & ",'" & _
'                    strCodRiesgo & "','" & strCodSubRiesgo & "','" & Trim(txtObservacion.Text) & "','" & gstrLogin & "') }"
'                adoConn.Execute .CommandText
                
                
                strCodCobroInteres = Modo_Cobro_Interes_Vencimiento
                strMontoInteres = CDec(lblMontoVencimiento.Caption) - CDec(txtValorNominal.Text)

                .CommandText = "{ call up_IVAdicInversionOrden('" & strCodFondoOrden & "','" & _
                    gstrCodAdministradora & "','','" & strFechaOrden & "','" & _
                    strCodTitulo & "','" & Trim(txtNemonico.Text) & "','" & gstrPeriodoActual & "','" & Mid(strFechaOrden, 5, 2) & "','" & _
                    "','" & strEstadoOrden & "','" & strCodAnalitica & "','" & strCodFile & "','" & _
                    strCodAnalitica & "','" & strCodClaseInstrumento & "','" & strCodSubClaseInstrumento & "','" & strCodTipoOrden & "','" & _
                    strCodOperacion & "','" & strCodNegociacion & "','" & strCodOrigen & "','" & Trim(txtDescripOrden.Text) & "','" & strCodEmisor & "','" & _
                    strCodAgente & "','" & strCodGarantia & "','',0,'" & strFechaPago & "','" & strFechaVencimiento & "','" & strFechaLiquidacion & "','" & _
                    strFechaEmision & "','" & strCodMoneda & "'," & CDec(txtValorNominal.Text) & ",'','" & strCodMoneda & "','" & strCodMoneda & "'," & CDec(txtCantidad.Text) & "," & _
                    CDec(txtTipoCambio.Text) & "," & CDec(txtTipoCambio.Text) & "," & CDec(txtValorNominal.Text) & ",100," & CDec(txtValorNominal.Text) & "," & _
                    CDec(txtPrecioUnitario(0).Text) & "," & CDec(txtPrecioUnitario(0).Text) & "," & CDec(lblSubTotal(0).Caption) & "," & CDec(txtInteresCorrido(0).Text) & "," & _
                    CDec(txtComisionAgente(0).Text) & "," & CDec(txtComisionCavali(0).Text) & "," & CDec(txtComisionConasev(0).Text) & "," & _
                    CDec(txtComisionBolsa(0).Text) & "," & CDec(txtComisionFondo(0).Text) & ",0,0,0," & CDec(lblComisionIgv(0).Caption) & "," & CDec(lblMontoTotal(0).Caption) & "," & _
                    CDec(txtPrecioUnitario(1).Text) & "," & CDec(txtPrecioUnitario(1).Text) & "," & CDec(lblSubTotal(1).Caption) & "," & CDec(txtInteresCorrido(1).Text) & "," & CDec(txtComisionAgente(1).Text) & "," & _
                    CDec(txtComisionCavali(1).Text) & "," & CDec(txtComisionConasev(1).Text) & "," & CDec(txtComisionBolsa(1).Text) & "," & _
                    CDec(txtComisionFondo(1).Text) & ",0,0,0," & CDec(lblComisionIgv(1).Caption) & "," & CDec(lblMontoTotal(1).Caption) & "," & _
                    CDec(lblMontoVencimiento.Caption) & "," & CInt(txtDiasPlazo.Text) & ",'','','','','','" & strCodReportado & "','" & strCodEmisor & "','" & strCodEmisor & "','','','',0,'','','" & strIndTitulo & "','" & _
                    strCodTipoTasa & "','" & strCodBaseAnual & "'," & CDec(txtTasa.Text) & ",'01','X','07','X'," & CDec(lblTirBrutaResumen.Caption) & "," & CDec(lblTirBrutaResumen.Caption) & "," & CDec(lblTirNetaResumen.Caption) & ",'" & _
                    strCodRiesgo & "','" & strCodSubRiesgo & "','" & Trim(txtObservacion.Text) & "','" & gstrLogin & "','" & gstrFechaActual & "','" & _
                    gstrLogin & "','" & gstrFechaActual & "','" & strCodTitulo & "','" & strCodCobroInteres & "'," & strMontoInteres & ",0,0,0,0,'01',0," & _
                    "0,0,0,0,0,0,0,0,0,0,'','','','','','','','','',''," & CDec(txtValorNominal.Text) & "," & CDec(txtCantidad.Text) & ",0,0) }"
                adoConn.Execute .CommandText
                
                
                '.CommandText = "COMMIT TRAN ProcOrden"
                'adoConn.Execute .CommandText
                                                                                                      
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

    'adoComm.CommandText = "ROLLBACK TRAN ProcOrden"
    'adoConn.Execute adoComm.CommandText
        
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
    
    If cboClaseInstrumento.ListIndex < 0 Then
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
            MsgBox "Debe seleccionar la Entidad Financiera.", vbCritical, Me.Caption
            If cboEmisor.Enabled Then cboEmisor.SetFocus
            Exit Function
        End If
        
        If strCodTipoOrden = Codigo_Orden_Compra Or strCodTipoOrden = Codigo_Orden_Pacto Then
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
        End If
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
    
    If CDbl(txtTasa.Text) = 0 Then
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
    
'    If CDbl(txtTipoCambio.Text) = 0 Then
'        MsgBox "Debe indicar el Tipo de Cambio.", vbCritical, Me.Caption
'        If txtTipoCambio.Enabled Then txtTipoCambio.SetFocus
'        Exit Function
'    End If
    
    '*** Validación de STOCK ***
    If strCodTipoOrden = Codigo_Orden_Venta Or strCodTipoOrden = Codigo_Orden_Quiebre Then
        If CCur(txtCantidad.Text) > CCur(lblStockNominal.Caption) Then
            MsgBox "Stock insuficiente para Registrar la Orden de Venta.", vbCritical, Me.Caption
            If txtValorNominal.Enabled Then txtValorNominal.SetFocus
            Exit Function
        End If
    Else
        If CCur(lblMontoVencimiento.Caption) = 0 Then
            MsgBox "Debe calcular el Valor al Vencimiento.", vbCritical, Me.Caption
            If cmdCalculo.Enabled Then cmdCalculo.SetFocus
            Exit Function
        End If
    End If
    
    If CDbl(lblTirBruta.Caption) = 0 Then
        MsgBox "Debe calcular la Tir Bruta.", vbCritical, Me.Caption
        If cmdCalculo.Enabled Then cmdCalculo.SetFocus
        Exit Function
    End If
    
    If CDbl(lblTirNeta.Caption) = 0 Then
        MsgBox "Debe calcular la Tir Neta.", vbCritical, Me.Caption
        If cmdCalculo.Enabled Then cmdCalculo.SetFocus
        Exit Function
    End If
    
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

Public Sub Salir()

    Unload Me
    
End Sub

Private Sub cboAgente_Click()

    strCodAgente = Valor_Caracter
    If cboAgente.ListIndex < 0 Then Exit Sub
    
    strCodAgente = Trim(arrAgente(cboAgente.ListIndex))
    
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
        Case Codigo_Base_Actual_Actual
            Dim adoRegistro     As ADODB.Recordset
            
            Set adoRegistro = New ADODB.Recordset
        
            adoComm.CommandText = "SELECT dbo.uf_ACValidaEsBisiesto(" & CInt(Right(lblFechaLiquidacion.Caption, 4)) & ") AS 'EsBisiesto'"
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                If adoRegistro("EsBisiesto") = 0 Then intBaseCalculo = 366
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
    End Select
    
    txtValorNominal_Change
    
    lblTirBruta.Caption = "0": lblTirNeta.Caption = "0"
    lblMontoVencimiento.Caption = "0"
    
End Sub


Private Sub cboClaseInstrumento_Click()

    strCodClaseInstrumento = Valor_Caracter
    If cboClaseInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodClaseInstrumento = Trim(arrClaseInstrumento(cboClaseInstrumento.ListIndex))
                
'    If strCodClaseInstrumento = "001" Then txtPrecioUnitario(0).Enabled = False
'    If strCodClaseInstrumento = "001" Then txtCantidad.Enabled = True
    If strCodClaseInstrumento = "001" Then strCalcVcto = "V"
'    If strCodClaseInstrumento = "002" Then txtPrecioUnitario(0).Enabled = True
'    If strCodClaseInstrumento = "002" Then txtCantidad.Enabled = False
    If strCodClaseInstrumento = "002" Then strCalcVcto = "D"

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
    
    strCodTitulo = Valor_Caracter
    strCodGrupo = Valor_Caracter
    strCodCiiu = Valor_Caracter
    strCodEmisor = Valor_Caracter
    'strCodAnalitica = Valor_Caracter
    lblAnalitica = strCodTipoInstrumentoOrden & "-??????": txtValorNominal.Text = "1"
    lblStockNominal = "0": strCodGrupo = Valor_Caracter
    
    If cboEmisor.ListIndex < 0 Then Exit Sub

    strCodEmisor = Left(Trim(arrEmisor(cboEmisor.ListIndex)), 8)
    strCodGrupo = Mid(Trim(arrEmisor(cboEmisor.ListIndex)), 9, 3)
    strCodCiiu = Right(Trim(arrEmisor(cboEmisor.ListIndex)), 4)

    '*** Validar Limites ***
    If strCodTipoInstrumentoOrden = Valor_Caracter Then Exit Sub
    If Not PosicionLimites() Then Exit Sub

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
                MsgBox "La Clasificación de Riesgo no está definida...", vbInformation, Me.Caption
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
        "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Valor_Caracter & "' AND IndInstrumento='X' AND IndVigente='X' AND IVF.CodEstructura='01' AND " & _
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
        "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Valor_Caracter & "' AND IndInstrumento='X' AND IndVigente='X' AND IVF.CodEstructura='01' AND " & _
        "CodFondo='" & strCodFondoOrden & "' AND CodAdministradora='" & gstrCodAdministradora & "' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumentoOrden, arrTipoInstrumentoOrden(), Sel_Defecto
        
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
            txtNemonico.Text = Trim(adoRegistro("DescripInicial")) & strFecha
            txtDescripOrden.Text = Trim(cboTipoInstrumentoOrden.Text) & " - " & Trim(txtNemonico.Text)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Tipo de Orden ***
    'strSQL = "SELECT IFTON.CodTipoOperacion CODIGO,DescripTipoOperacion DESCRIP " & _
    '    "FROM InversionFileTipoOperacionNegociacion IFTON JOIN TipoOperacionNegociacion TON ON(TON.CodTipoOperacion=IFTON.CodTipoOperacion)" & _
    '    "WHERE IFTON.CodFile='" & strCodTipoInstrumentoOrden & "' ORDER BY DescripTipoOperacion"
    
    strSQL = "SELECT IFTON.CodTipoOperacion CODIGO,TON.DescripParametro DESCRIP " & _
        "FROM InversionFileTipoOperacionNegociacion IFTON JOIN AuxiliarParametro TON ON(TON.CodParametro = IFTON.CodTipoOperacion)" & _
        " WHERE " & _
        " IFTON.CodFile = '" & strCodTipoInstrumentoOrden & "' AND " & _
        " TON.CodTipoParametro = 'OPECAJ' AND " & _
        " TON.ValorParametro = 'I' " & _
        " ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoOrden, arrTipoOrden(), Valor_Caracter

    If cboTipoOrden.ListCount > 0 Then cboTipoOrden.ListIndex = 0
    
    lblAnalitica.Caption = strCodTipoInstrumentoOrden & " - ????????"
    strCodFile = strCodTipoInstrumentoOrden

    '*** Clase de Instrumento ***
    strSQL = "SELECT CodDetalleFile CODIGO,DescripDetalleFile DESCRIP FROM InversionDetalleFile WHERE CodFile='" & strCodTipoInstrumentoOrden & "' AND IndVigente='X' ORDER BY CodDetalleFile"
    CargarControlLista strSQL, cboClaseInstrumento, arrClaseInstrumento(), Valor_Caracter 'Sel_Defecto
    
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
    
    txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaOrden.Value, Codigo_Moneda_Local, strCodMoneda))
    If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaOrden.Value), Codigo_Moneda_Local, strCodMoneda))
    
    If strCodMoneda = Codigo_Moneda_Local Then
        txtTipoCambio.Text = CStr(gdblTipoCambio)
        txtTipoCambio.Enabled = False
    Else
        txtTipoCambio.Enabled = True
    End If
    dblTipoCambio = CDbl(txtTipoCambio.Text)
        
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
            lblDescrip(6) = "Entidad Financiera"
            
            If chkTitulo.Value Then
                strSQL = "SELECT CodTitulo CODIGO,(Nemotecnico + ' ' + DescripTitulo)DESCRIP FROM InstrumentoInversion " & _
                    "WHERE CodFile='" & strCodFile & "' AND CodDetalleFile='" & strCodClaseInstrumento & "' AND IndVigente='X' ORDER BY DescripTitulo"
                CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
            
                If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0
            End If
            
            fraComisionMontoFL2.Visible = False

        Case Codigo_Orden_Venta, Codigo_Orden_Quiebre
            chkTitulo.Enabled = False
            cboTitulo.Visible = True: cboEmisor.Visible = False
            lblDescrip(6) = "Título"
            
            strSQL = "SELECT InstrumentoInversion.CodTitulo CODIGO," & _
                "(RTRIM(InstrumentoInversion.CodTitulo) + ' ' + RTRIM(InstrumentoInversion.Nemotecnico) + ' ' + RTRIM(InstrumentoInversion.DescripTitulo)) DESCRIP FROM InstrumentoInversion,InversionKardex " & _
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
    
    txtValorNominal_Change
    
    lblTirBruta.Caption = "0": lblTirNeta.Caption = "0"
    lblMontoVencimiento.Caption = "0"
    
End Sub


Private Sub cboTitulo_Click()

    Dim adoRegistro     As ADODB.Recordset
    Dim intRegistro     As Integer
    
    strCodGarantia = Valor_Caracter: txtDescripOrden = Valor_Caracter: strCodAnalitica = Valor_Caracter
    lblAnalitica = strCodTipoInstrumentoOrden & "-??????": txtValorNominal = 0
    lblStockNominal = "0"
    strCodEmisor = Valor_Caracter: strCodGrupo = Valor_Caracter
    If cboTitulo.ListIndex < 0 Then Exit Sub

    strCodGarantia = Trim(arrTitulo(cboTitulo.ListIndex))

    With adoComm
        Set adoRegistro = New ADODB.Recordset

        .CommandText = "SELECT INV.CodAnalitica,INV.ValorNominal,INV.CodMoneda,INV.CodEmisor," & _
            "INV.CodGrupo,INV.FechaEmision,INV.FechaVencimiento," & _
            "INV.TasaInteres,INV.CodRiesgo,INV.CodSubRiesgo,INV.CodTipoTasa," & _
            "INV.BaseAnual,INV.Nemotecnico,KAR.SaldoFinal " & _
            "FROM InstrumentoInversion INV " & _
            "JOIN InversionKardex KAR ON (INV.CodTitulo = KAR.CodTitulo) " & _
            "WHERE " & _
            "INV.CodTitulo = '" & strCodGarantia & "' AND " & _
            "KAR.CodFondo = '" & strCodFondoOrden & "' AND " & _
            "KAR.CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
            "KAR.IndUltimoMovimiento = 'X'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
            lblAnalitica.Caption = strCodFile & "-" & strCodAnalitica
            
            intRegistro = ObtenerItemLista(arrMoneda(), adoRegistro("CodMoneda"))
            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
                
            strCodEmisor = Trim(adoRegistro("CodEmisor")): strCodGrupo = Trim(adoRegistro("CodGrupo"))
                                                                        
            'Setea el combo de emisor (banco)
            intRegistro = ObtenerItemLista(arrEmisor(), strCodEmisor, 1, 8)
            If intRegistro >= 0 Then cboEmisor.ListIndex = intRegistro
            
            dtpFechaEmision.Value = adoRegistro("FechaEmision")
            dtpFechaVencimiento.Value = dtpFechaOrden.Value  'adoRegistro("FechaVencimiento")  'ACR:05/02/2009
            
            dtpFechaEmision_Change
            dtpFechaVencimiento_Change
            
            lblFechaEmision.Caption = dtpFechaEmision.Value
            lblFechaVencimiento.Caption = dtpFechaVencimiento.Value

            txtNemonico.Text = Trim(adoRegistro("Nemotecnico"))
            
            intRegistro = ObtenerItemLista(arrTipoTasa(), adoRegistro("CodTipoTasa"))
            If intRegistro >= 0 Then cboTipoTasa.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrBaseAnual(), adoRegistro("BaseAnual"))
            If intRegistro >= 0 Then cboBaseAnual.ListIndex = intRegistro
            
            txtTasa.Text = adoRegistro("TasaInteres")
            txtValorNominal.Text = CStr(adoRegistro("ValorNominal"))
                
            lblStockNominal.Caption = adoRegistro("SaldoFinal")
            txtCantidad.Text = adoRegistro("SaldoFinal")
            
            strCodEmisor = Trim(adoRegistro("CodEmisor")): strCodGrupo = Trim(adoRegistro("CodGrupo"))
            strCodRiesgo = Trim(adoRegistro("CodRiesgo"))
            strCodSubRiesgo = Trim(adoRegistro("CodSubRiesgo"))
            lblMoneda.Caption = ObtenerDescripcionMoneda(adoRegistro("CodMoneda"))
            
            lblTasaCupon.Caption = Trim(txtTasa.Text)
           
            If adoRegistro("BaseAnual") = Codigo_Base_30_360 Then
                lblBaseCalculo.Caption = "360/360"
            End If
            
            If adoRegistro("BaseAnual") = Codigo_Base_Actual_360 Then
                lblBaseCalculo.Caption = "Act/360"
            End If
            
            If adoRegistro("BaseAnual") = Codigo_Base_Actual_Actual Then
                lblBaseCalculo.Caption = "Act/Act"
            End If
            
            If adoRegistro("BaseAnual") = Codigo_Base_Actual_365 Then
                lblBaseCalculo.Caption = "Act/365"
            End If
            
            If adoRegistro("BaseAnual") = Codigo_Base_30_365 Then
                lblBaseCalculo.Caption = "360/365"
            End If
            
            cboMoneda.Enabled = False
            cboTipoTasa.Enabled = False
            cboBaseAnual.Enabled = False
            dtpFechaVencimiento.Enabled = False
            txtDiasPlazo.Enabled = False
            txtValorNominal.Enabled = False
            txtTasa.Enabled = True 'False 'ACR: 05/02/2009
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

    txtDescripOrden.Text = Trim(cboTipoInstrumentoOrden.Text) & " - " & Left(cboTitulo.Text, 15)
        
End Sub

Private Sub chkAplicar_Click(Index As Integer)

    If chkAplicar(Index).Value Then
        Call AplicarCostos(Index)
    Else
        Call IniciarComisiones
        Call CalculoTotal(Index)
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
        lblDescrip(6).Caption = "Entidad Financiera"
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

    'Call CalcularValorVencimiento 'ACR
    'Call CalcularPrecio
    Call CalcularTirBruta
    Call CalcularTirNeta
    Call CalcularValorVencimiento

    
End Sub

Private Sub cmdEnviar_Click()

    Dim strFechaDesde       As String, strFechaHasta        As String
    Dim intRegistro         As Integer, intContador         As Integer
    Dim datFecha            As Date
    
    If adoConsulta.RecordCount = 0 Then Exit Sub
    
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

Private Sub cmdExportarExcel_Click()
    Call ExportarExcel
End Sub

Private Sub dtpFechaEmision_Change()

    lblFechaEmision.Caption = CStr(dtpFechaEmision.Value)
    
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
    
    'txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaLiquidacion.Value, dtpFechaVencimiento.Value))
    txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaEmision.Value, dtpFechaVencimiento.Value))
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
    
    strSQL = "SELECT NumOrden,FechaOrden,FechaLiquidacion,CodTitulo,Nemotecnico,EstadoOrden,CodFile,CodAnalitica,TipoOrden,IOR.CodMoneda," & _
        "(RTRIM(DescripParametro) + SPACE(1) + DescripOrden) DescripOrden,CantOrden,ValorNominal,PrecioUnitarioMFL1,MontoTotalMFL1, CodSigno DescripMoneda " & _
        "FROM InversionOrden IOR JOIN AuxiliarParametro TON ON(TON.CodParametro=IOR.TipoOrden AND TON.CodTipoParametro='OPECAJ' AND TON.ValorParametro = 'I') " & _
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
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Papeleta de Inversión"
    
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
        
    '*** Tipo de Orden ***
    'strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro = 'OPECAJ' AND ValorParametro = 'I' ORDER BY DescripParametro"
    'CargarControlLista strSQL, cboTipoOrden, arrTipoOrden(), Valor_Caracter
    
    '*** Agente ***
    strSQL = "SELECT (CodPersona + CodGrupo + CodCiiu) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Agente & "' AND IndVigente='X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboAgente, arrAgente(), Sel_Defecto
    
    '*** Emisor - Banco ***
    strSQL = "SELECT (CodPersona + CodGrupo + CodCiiu) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' AND IndBanco='X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboEmisor, arrEmisor(), Sel_Defecto

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

    tabRFCortoPlazo.TabVisible(2) = False

    dtpFechaOrdenDesde.Value = gdatFechaActual
    dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    dtpFechaLiquidacionDesde.Value = Null
    dtpFechaLiquidacionHasta.Value = Null
    
    lblPorcenIgv(0).Caption = CStr(gdblTasaIgv)
    lblPorcenIgv(1).Caption = CStr(gdblTasaIgv)
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT CodFile FROM InversionFile " & _
            "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Valor_Caracter & "' AND IndInstrumento='X' AND IndVigente='X' AND CodEstructura='01' " & _
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
            If PreviousTab = 0 And strEstado = Reg_Consulta Then tabRFCortoPlazo.Tab = 0
            If strEstado = Reg_Defecto Then tabRFCortoPlazo.Tab = 0
            If tabRFCortoPlazo.Tab = 2 Then
                fraDatosNegociacion.Caption = "Negociación" & Space(1) & "-" & Space(1) & _
                    Trim(cboTipoOrden.Text) & Space(1) & Trim(Left(cboTitulo.Text, 15))
            End If
    End Select
    
End Sub

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
    txtValorNominal_Change
    
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtCantidad, Decimales_Monto)
    
End Sub

Private Sub txtComisionAgente_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionAgente(Index), Decimales_Monto)
    
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
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtComisionFondo_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionFondo(Index), Decimales_Monto)
    
    If chkAplicar(Index).Value Then
        'ActualizaPorcentaje txtComisionFondo, lblPorcenFondo
        ActualizaPorcentaje txtComisionFondo(Index), lblPorcenFondo(Index)
    End If
    
End Sub

Private Sub txtComisionFondo_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionFondo(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtDiasPlazo_Change()
       
    Call FormatoCajaTexto(txtDiasPlazo, 0)
    
    If IsNumeric(txtDiasPlazo.Text) Then
        'dtpFechaVencimiento.Value = DateAdd("d", txtDiasPlazo.Text, CVDate(dtpFechaLiquidacion.Value)) 'ACR
        dtpFechaVencimiento.Value = DateAdd("d", txtDiasPlazo.Text, CVDate(dtpFechaEmision.Value))
    Else
        dtpFechaVencimiento.Value = dtpFechaLiquidacion.Value
    End If

    Call CalculoTotal(0)
    lblDiasPlazo.Caption = CStr(txtDiasPlazo.Text)
    dtpFechaPago.Value = dtpFechaVencimiento.Value
    dtpFechaPago_Change
    lblFechaVencimiento.Caption = CStr(dtpFechaVencimiento.Value)
    lblFechaCupon.Caption = CStr(dtpFechaVencimiento.Value)
    
    If CInt(txtDiasPlazo.Text) > 0 Then tabRFCortoPlazo.TabEnabled(2) = True
    
End Sub


Private Sub CalculoTotal(Index As Integer)

    Dim curComImp As Currency, curMonTotal As Currency

    If Not IsNumeric(txtComisionAgente(Index).Text) And Not IsNumeric(txtComisionBolsa(Index).Text) And Not IsNumeric(txtComisionConasev(Index).Text) And Not IsNumeric(txtComisionCavali(Index).Text) And Not IsNumeric(txtComisionFondo(Index).Text) Then Exit Sub
    
    curComImp = CCur(CCur(txtComisionAgente(Index).Text) + CCur(txtComisionBolsa(Index).Text) + CCur(txtComisionConasev(Index).Text) + CCur(txtComisionCavali(Index).Text) + CCur(txtComisionFondo(Index).Text)) * CDbl(lblPorcenIgv(Index).Caption)
    lblComisionIgv(Index).Caption = CStr(curComImp)

    curComImp = CCur(txtComisionAgente(Index).Text) + CCur(txtComisionBolsa(Index).Text) + CCur(txtComisionConasev(Index).Text) + CCur(txtComisionCavali(Index).Text) + CCur(txtComisionFondo(Index).Text) + CCur(lblComisionIgv(Index).Caption)

    lblComisionesResumen(Index).Caption = CStr(curComImp)
            
    If strCodTipoOrden = Codigo_Orden_Compra Or strCodTipoOrden = Codigo_Orden_Pacto Then  '*** Compra ***
        If Index = 0 Then
            curMonTotal = CCur(lblSubTotal(Index).Caption) + curComImp
        Else
            curMonTotal = CCur(lblSubTotal(Index).Caption) - curComImp
        End If
    ElseIf strCodTipoOrden = Codigo_Orden_Venta Or Codigo_Orden_Quiebre Then '*** Venta ***
        curMonTotal = CCur(lblSubTotal(Index).Caption) - curComImp
    End If
    
    curMonTotal = curMonTotal + CCur(txtInteresCorrido(Index).Text)

    lblMontoTotal(Index).Caption = CStr(curMonTotal)
    
End Sub


Private Sub txtDiasPlazo_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N")
    
End Sub

Private Sub txtDiasPlazo_LostFocus()

'    cboEmisor_Click
    If CInt(txtDiasPlazo.Text) > 0 Then tabRFCortoPlazo.TabEnabled(2) = True
    
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

'    Dim intRes As Integer
'
'    If Not LEsDiaUtil(dtpFechaVencimiento) Then
'       dtpFechaVencimiento.Text = LProxDiaUtil(dtpFechaVencimiento.Text)
'    End If
'
'    If CVDate(dtpFechaEmision.Text) > CVDate(dtpFechaVencimiento.Text) Then
'       MsgBox "Fecha de Vencimiento debe ser posterior a la Fecha de Emisión", vbCritical
'       dtpFechaVencimiento.Text = dtpFechaEmision.Text
'       dtpFechaVencimiento.SetFocus
'    End If
'
'    txtDiasPlazo.Text = DateDiff("d", CVDate(dtpFechaEmision.Text), CVDate(dtpFechaVencimiento.Text))
'
'    'LCalIntCor
'    CalculoTotal
    
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

Private Sub txtNemonico_Change()

    txtDescripOrden.Text = Trim(cboTipoInstrumentoOrden.Text) & " - " & Trim(txtNemonico.Text)
    
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
            'ActualizaComision txtPorcenAgente, txtComisionAgente
            ActualizaComision txtPorcenAgente(Index), txtComisionAgente(Index)
        End If
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtPrecioUnitario_Change(Index As Integer)

    Call FormatoCajaTexto(txtPrecioUnitario(Index), Decimales_Precio)
    txtValorNominal_Change
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

Private Sub txtTasa_Change()

    Call FormatoCajaTexto(txtTasa, Decimales_Tasa)
    
    Call txtValorNominal_Change
    
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


Private Sub txtValorNominal_Change()

    Dim curCantidad         As Currency, curSubTotal            As Currency
    Dim dblPreUni           As Double, dblFactorDiarioNormal    As Double
    Dim curValorNominal     As Double
    
    Call FormatoCajaTexto(txtValorNominal, Decimales_Monto)
    
    If Trim(txtCantidad.Text) = Valor_Caracter Then Exit Sub
    
    If IsNumeric(txtCantidad.Text) Then
       curCantidad = CCur(txtCantidad.Text)
    Else
       curCantidad = 0
    End If
        
    lblCantidadResumen.Caption = CStr(curCantidad)
    
    If IsNumeric(txtValorNominal.Text) Then
       curValorNominal = CCur(txtValorNominal.Text)
    Else
       curValorNominal = 0
    End If
    
    
    If IsNumeric(txtPrecioUnitario(0).Text) Then dblPreUni = CDbl(txtPrecioUnitario(0).Text) * 0.01
    
    curSubTotal = Round(curCantidad * curValorNominal * dblPreUni, 2)
    lblSubTotal(0).Caption = curSubTotal
    
    Call CalculoTotal(0)
    
    'If CCur(txtCantidad.Text) > 0 And cboTitulo.ListIndex > 0 And strIndCuponCero = Valor_Caracter Then
    '   txtInteresCorrido(0).Text = CStr(CalculoInteresCorrido(strCodGarantia, CDbl(curCantidad), dtpFechaEmision.Value, dtpFechaLiquidacion.Value, strCodIndiceFinal, strCodTipoAjuste, strCodTipoTasa, strCodPeriodoPago, strCodIndiceInicial, strCodBaseAnual, intBaseCalculo))
    'ElseIf CCur(txtCantidad.Text) > 0 And Trim(txtTasa.Text) <> Valor_Caracter And Trim(txtDiasPlazo.Text) <> Valor_Caracter And strIndCuponCero = Valor_Caracter Then
    If CCur(txtCantidad.Text) > 0 And Trim(txtTasa.Text) <> Valor_Caracter And Trim(txtDiasPlazo.Text) <> Valor_Caracter And strIndCuponCero = Valor_Caracter Then
        If CCur(txtCantidad.Text) > 0 And CDbl(txtTasa.Text) > 0 And CInt(txtDiasPlazo.Text) And strIndCuponCero = Valor_Caracter Then
            '*** Calculando factores ***
            dblTasaCuponNormal = FactorAnualNormal(CDbl(txtTasa.Text), CInt(txtDiasPlazo.Text), intBaseCalculo, strCodTipoTasa, Valor_Indicador, Valor_Caracter, 0, CInt(txtDiasPlazo.Text), 1)
            dblFactorDiarioNormal = FactorDiarioNormal(dblTasaCuponNormal, CInt(txtDiasPlazo.Text), strCodTipoTasa, Valor_Indicador, CInt(txtDiasPlazo.Text))
    
            txtInteresCorrido(0).Text = CStr(CalculoInteresCorrido(strCodGarantia, CDbl(curCantidad), dtpFechaEmision.Value, dtpFechaLiquidacion.Value, strCodIndiceFinal, strCodTipoAjuste, strCodTipoTasa, strCodPeriodoPago, strCodIndiceInicial, strCodBaseAnual, intBaseCalculo, dblFactorDiarioNormal))
                                             'CalculoInteresCorrido(strpCodTitulo, dblpCantidad,      datpFechaEmision,      datpFechaLiquidacion,      strpCodIndiceFinal, strpTipoAjuste, strpTipoTasa,    strpPeriodoPago,   strpCodIndiceInicial,strpBaseCalculo, intpBase,        dblpFactorDiario)

        End If
    Else
       txtInteresCorrido(0).Text = "0"
    End If
    
End Sub

Private Sub txtValorNominal_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtValorNominal, Decimales_Monto)
    
End Sub

Private Sub ExportarExcel()
    
    Dim adoRegistro As ADODB.Recordset
    Dim execSQL As String
    Dim rutaExportacion As String
    
    Dim datFechaSiguiente As Date
    Dim strFechaLiquidacionHasta As String
    
    Set frmFormulario = frmOrdenDepositoBancario
    
    Set adoRegistro = New ADODB.Recordset
    
    'If TodoOK() Then
        
        Dim strNameProc As String
        
        gstrNameRepo = "OrdenDepositoBancario"
        
        strNameProc = ObtenerBaseReporte(gstrNameRepo)
        
        Dim arrParmS(6)
        
        arrParmS(0) = Trim(strCodFondo)
        arrParmS(1) = Trim(gstrCodAdministradora)
        
        If strCodTipoInstrumento <> Valor_Caracter Then
            arrParmS(2) = Trim(strCodTipoInstrumento)
        Else
            arrParmS(2) = "%"
        End If
        
        If IsNull(dtpFechaOrdenDesde.Value) And IsNull(dtpFechaOrdenHasta.Value) Then
            arrParmS(3) = Convertyyyymmdd(dtpFechaLiquidacionDesde.Value)
            datFechaSiguiente = DateAdd("d", 1, dtpFechaLiquidacionHasta.Value)
            strFechaLiquidacionHasta = Convertyyyymmdd(datFechaSiguiente)
            arrParmS(4) = strFechaLiquidacionHasta
            arrParmS(5) = "L"
        Else
            arrParmS(3) = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
            arrParmS(4) = Convertyyyymmdd(dtpFechaOrdenHasta.Value)
            arrParmS(5) = "O"
        End If
        
        If strCodEstado <> Valor_Caracter Then
            arrParmS(6) = strCodEstado
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
