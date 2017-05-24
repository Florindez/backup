VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMainMdi 
   BackColor       =   &H8000000C&
   Caption         =   "Módulo Inversión"
   ClientHeight    =   9795
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   16725
   Icon            =   "frmMainMdi.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Tag             =   "I"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbMdi 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   9420
      Width           =   16725
      _ExtentX        =   29501
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Picture         =   "frmMainMdi.frx":08CA
            Object.ToolTipText     =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   23839
            Text            =   "Acción"
            TextSave        =   "Acción"
            Object.ToolTipText     =   "Mensajes"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMdi 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   16725
      _ExtentX        =   29501
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlMdi"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Agregar"
            Description     =   "Boton 0"
            Object.ToolTipText     =   "&Nuevo"
            Object.Tag             =   "0"
            ImageKey        =   "NUEVO"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Modificar"
            Description     =   "Boton 1"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "1"
            ImageKey        =   "CONSULTAR"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Eliminar"
            Description     =   "Boton 2"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "4"
            ImageKey        =   "ELIMINAR"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Guardar"
            Object.ToolTipText     =   "Guardar"
            Object.Tag             =   "2"
            ImageKey        =   "GUARDAR"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sep0"
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   "6"
            ImageKey        =   "IMPRIMIR"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Refrescar"
            Object.ToolTipText     =   "Refrescar"
            ImageKey        =   "REFRESCAR"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sep1"
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Primero"
            Object.ToolTipText     =   "Primer Registro"
            ImageKey        =   "PRIMERO"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Anterior"
            Object.ToolTipText     =   "Registro Anterior "
            ImageKey        =   "ANTERIOR"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Siguiente Registro"
            ImageKey        =   "SIGUIENTE"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Ultimo Registro"
            ImageKey        =   "ULTIMO"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "sep3"
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar"
            Object.Tag             =   "5"
            ImageKey        =   "BUSCAR"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reportes"
            Object.ToolTipText     =   "Reportes"
            Object.Tag             =   "7"
            ImageKey        =   "REPORTES"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   10
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Repo1"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Repo2"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Repo3"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Repo4"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Repo5"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Repo6"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Repo7"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Repo8"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Repo9"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Repo10"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sep4"
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Bloquear"
            Object.ToolTipText     =   "Bloquear"
            ImageKey        =   "BLOQUEAR"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Ayuda"
            Object.ToolTipText     =   "Ayuda"
            ImageKey        =   "AYUDA"
         EndProperty
      EndProperty
      BorderStyle     =   1
      OLEDropMode     =   1
      Begin VB.TextBox txtEmpresa 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Empresa"
         Top             =   30
         Width           =   6000
      End
      Begin VB.TextBox txtUsuarioSistema 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Usuario"
         Top             =   30
         Width           =   5800
      End
      Begin VB.TextBox txtFechaSistema 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   15500
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Fecha del Sistema"
         Top             =   30
         Width           =   1455
      End
   End
   Begin MSComctlLib.ImageList imlMdi 
      Left            =   2760
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":0A4F
            Key             =   "NUEVO"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":2759
            Key             =   "GUARDAR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":3033
            Key             =   "BUSCAR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":3487
            Key             =   "CONSULTAR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":5191
            Key             =   "IMPRIMIR"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":6E9B
            Key             =   "ELIMINAR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":7CED
            Key             =   "BLOQUEAR"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":8B3F
            Key             =   "REPORTES"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":9991
            Key             =   "AYUDA"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":9CAB
            Key             =   "CANCELAR"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":9FC5
            Key             =   "PRIMERO"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":A89F
            Key             =   "ANTERIOR"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":B179
            Key             =   "SIGUIENTE"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":BA53
            Key             =   "ULTIMO"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":C32D
            Key             =   "REFRESCAR"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Registro"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuRegistro 
         Caption         =   "&Nuevo"
         Index           =   0
      End
      Begin VB.Menu mnuRegistro 
         Caption         =   "&Modificar"
         Index           =   1
      End
      Begin VB.Menu mnuRegistro 
         Caption         =   "&Eliminar"
         Index           =   2
      End
      Begin VB.Menu mnuRegistro 
         Caption         =   "&Guardar"
         Index           =   3
      End
      Begin VB.Menu mnuRegistro 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuRegistro 
         Caption         =   "&Imprimir"
         Index           =   5
      End
      Begin VB.Menu mnuRegistro 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuRegistro 
         Caption         =   "&Salir"
         Index           =   99
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Invertir"
      Index           =   1
      Begin VB.Menu mnuInvertir 
         Caption         =   "Renta Variable"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInvertir 
         Caption         =   "Renta Fija Corto Plazo"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInvertir 
         Caption         =   "Renta Fija Largo Plazo"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInvertir 
         Caption         =   "-"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInvertir 
         Caption         =   "Operación de Reporte"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInvertir 
         Caption         =   "-"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInvertir 
         Caption         =   "Depósitos a Plazo"
         Index           =   6
      End
      Begin VB.Menu mnuInvertir 
         Caption         =   "Certificados de Depósitos"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInvertir 
         Caption         =   "-"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInvertir 
         Caption         =   "Forward"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInvertir 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuInvertir 
         Caption         =   "Desembolso de Acreencias"
         Index           =   11
      End
      Begin VB.Menu mnuInvertir 
         Caption         =   "Cancelación de Acreencias"
         Index           =   12
      End
      Begin VB.Menu mnuInvertir 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuInvertir 
         Caption         =   "Solicitud de descuento de Contratos"
         Index           =   14
      End
      Begin VB.Menu mnuInvertir 
         Caption         =   "Bandeja de Inversiones"
         Index           =   15
      End
      Begin VB.Menu mnuInvertir 
         Caption         =   "Desembolso de descuento de Contratos"
         Index           =   16
      End
      Begin VB.Menu mnuInvertir 
         Caption         =   "Cancelación de descuento de Contratos"
         Index           =   17
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Mantenimiento"
      Index           =   2
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "&Títulos"
         Index           =   0
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "&Proyecto Inmobiliario "
         Index           =   1
         Visible         =   0   'False
         Begin VB.Menu mnuProyectoInmo 
            Caption         =   "Definicion/Planificacion"
            Index           =   0
         End
         Begin VB.Menu mnuProyectoInmo 
            Caption         =   "Constitucion"
            Index           =   1
         End
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "Tasas"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "Limites"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "Costos de Negociación"
         Index           =   5
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "Precios / TIR"
         Index           =   7
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "Eventos Corporativos"
         Index           =   8
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Procesos"
      Index           =   3
      Begin VB.Menu mnuProcesos 
         Caption         =   "Confirmaciones"
         Index           =   0
         Begin VB.Menu mnuConfirmaciones 
            Caption         =   "Ordenes"
            Index           =   0
         End
         Begin VB.Menu mnuConfirmaciones 
            Caption         =   "Eventos Corporativos"
            Index           =   1
         End
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Operaciones"
         Index           =   1
         Begin VB.Menu mnuOperaciones 
            Caption         =   "Quiebre de Depósitos"
            Index           =   0
         End
         Begin VB.Menu mnuOperaciones 
            Caption         =   "Carga Masiva de Operaciones"
            Index           =   1
         End
         Begin VB.Menu mnuOperaciones 
            Caption         =   "PrePago Renta Fija"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuOperaciones 
            Caption         =   "-"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuOperaciones 
            Caption         =   "Cobertura con Monedas"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnuOperaciones 
            Caption         =   "Cobro de Cupón Renta Fija"
            Index           =   5
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "-"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Importación"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "-"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Vencimientos"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Cálculo de TIR"
         Index           =   7
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Informes"
      Index           =   4
      Begin VB.Menu mnuInformes 
         Caption         =   "Control Diario"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInformes 
         Caption         =   "Kardex Cartera"
         Index           =   1
      End
      Begin VB.Menu mnuInformes 
         Caption         =   "-"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInformes 
         Caption         =   "Reglamento"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInformes 
         Caption         =   "Análisis de Cartera"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInformes 
         Caption         =   "-"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInformes 
         Caption         =   "Limites"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInformes 
         Caption         =   "-"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInformes 
         Caption         =   "Otros"
         Index           =   8
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Financiamiento"
      Index           =   5
      Begin VB.Menu mnuFinanciamiento 
         Caption         =   "Ordenes de Financiamiento"
         Index           =   0
         Begin VB.Menu mnuOrdenFinanciamiento 
            Caption         =   "Registro de Préstamos"
            Index           =   0
         End
         Begin VB.Menu mnuOrdenFinanciamiento 
            Caption         =   "Pago de Préstamos"
            Index           =   1
         End
      End
      Begin VB.Menu mnuFinanciamiento 
         Caption         =   "Confirmación de Ordenes de Financiamiento"
         Index           =   1
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Herramientas"
      Index           =   6
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Calculadora"
         Index           =   0
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Evaluación de TIR"
         Index           =   2
         Visible         =   0   'False
         Begin VB.Menu mnuEvaluacion 
            Caption         =   "Bonos"
            Index           =   0
         End
         Begin VB.Menu mnuEvaluacion 
            Caption         =   "Cartera"
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Ver"
      Index           =   7
      Begin VB.Menu mnuVer 
         Caption         =   "Barra de Estado"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuVer 
         Caption         =   "Barra de Herramientas"
         Checked         =   -1  'True
         Index           =   1
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "Ve&ntana"
      Index           =   8
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Ayuda"
      Index           =   9
      Begin VB.Menu mnuAyuda 
         Caption         =   "Acerca del Módulo de Inversión"
         Index           =   0
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "Menu PopUp"
      Index           =   10
      Visible         =   0   'False
      Begin VB.Menu mnuEmergente 
         Caption         =   "Cambio de Fondo"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMainMdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuopc62_Click(Index As Integer)

    Dim vntRetVal As Variant
    
    '*** Ejecuta Calculadora ***
    vntRetVal = Shell("CALC.EXE", vbNormalFocus)
    
End Sub

Private Sub MDIForm_Load()

    App.Title = "Inversion"
    
    Me.Show
    Sleep 0&
    
    If App.PrevInstance Then
        MsgBox "La aplicación " & App.Title & " ya se está ejecutando...", vbCritical, "Control de Instancias"
        Unload Me
        End
    End If
    
    frmMainMdi.BackColor = RGB(102, 180, 255)
    
    frmAcceso.Show vbModal
    Sleep 0&
    
     '---/// Acceso al modulo de seguridad
    Call ValidarPermisoUsoMenu(gstrLoginUS, Me, Trim(App.Title), Separador_Codigo_Objeto)
    
    If gboolMostrarSelectAdministradora Then frmSeleccionaAdministradora.Show vbModal
    
    Call OcultarReportes
    
    If gstrLoginUS = "admin" Or gstrLoginUS = "sa" Then
        Me.Caption = gstrServer & "\\" & gstrDataBase & "\\" & Me.Caption
    End If

End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 1 Then
    
        PopupMenu mnuGeneral(10)
    
    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    On Error GoTo CtrlError
    adoConn.Close: Set adoConn = Nothing
    
CtrlError:
    End
    
End Sub

Private Sub mnuAyuda_Click(Index As Integer)

    Select Case Index
        Case 0: frmAcercade.Show vbModal
    End Select
    
End Sub

Private Sub mnuConfirmaciones_Click(Index As Integer)

    Select Case Index
        Case 0: strNombreObjeto = "frmConfirmacionOrden"
                gstrNombreObjetoMenuPulsado = mnuConfirmaciones.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmConfirmacionOrden.Show
                End If
        
        Case 1: strNombreObjeto = "frmProEventoCorporativo"
                gstrNombreObjetoMenuPulsado = mnuConfirmaciones.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmProEventoCorporativo.Show
                End If
        
    End Select
    
End Sub



Private Sub mnuEmergente_Click(Index As Integer)
    
    Select Case Index
        
        Case 0: frmSeleccionaAdministradora.Show
        
    End Select
    
End Sub

Private Sub mnuFinanciamiento_Click(Index As Integer)
    Select Case Index
    
        Case 1: strNombreObjeto = "frmConfirmacionOrdenFinanciamiento"
                gstrNombreObjetoMenuPulsado = mnuFinanciamiento.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLogin), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmConfirmacionOrdenFinanciamiento.Show
                End If

    
    End Select
            
End Sub

Private Sub mnuGeneral_Click(Index As Integer)
         
    Select Case Index
        Case 99
            Unload Me
    End Select
    
End Sub

Private Sub mnuHerramientas_Click(Index As Integer)

    Select Case Index
        Case 0
            Dim lngValorRetorno As Long
    
            '*** Ejecuta Calculadora ***
            lngValorRetorno = Shell("CALC.EXE", vbNormalFocus)
    End Select
End Sub

Private Sub mnuInformes_Click(Index As Integer)

    Select Case Index
        Case 1
            
                strNombreObjeto = "frmKardexInversion"
                gstrNombreObjetoMenuPulsado = mnuInformes.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmKardexInversion.Show
                End If
        
    End Select
    
End Sub

Private Sub mnuInvertir_Click(Index As Integer)
    
    Dim strCodObjeto As String, strNombreObjeto As String

    Select Case Index
        Case 0: strNombreObjeto = "frmOrdenRentaVariable"
                gstrNombreObjetoMenuPulsado = mnuInvertir.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmOrdenRentaVariable.Show
                End If
        
        Case 1: strNombreObjeto = "frmOrdenRentaFijaCortoPlazo"
                gstrNombreObjetoMenuPulsado = mnuInvertir.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmOrdenRentaFijaCortoPlazo.Show
                End If
        
        Case 2: strNombreObjeto = "frmOrdenRentaFijaLargoPlazo"
                gstrNombreObjetoMenuPulsado = mnuInvertir.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmOrdenRentaFijaLargoPlazo.Show
                End If

        Case 4: strNombreObjeto = "frmOrdenReporteRentaVariable"
                gstrNombreObjetoMenuPulsado = mnuInvertir.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmOrdenReporteRentaVariable.Show
                End If

        Case 6: strNombreObjeto = "frmOrdenDepositoBancario"
                gstrNombreObjetoMenuPulsado = mnuInvertir.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmOrdenDepositoBancario.Show
                End If

        Case 7: strNombreObjeto = "frmOrdenCDepositoBancario"
                gstrNombreObjetoMenuPulsado = mnuInvertir.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmOrdenCDepositoBancario.Show
                End If
        
        Case 9: strNombreObjeto = "frmCobertura"
                gstrNombreObjetoMenuPulsado = mnuInvertir.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmCobertura.Show
                End If

        Case 11: strNombreObjeto = "frmDesembolsoAcreencias"
                gstrNombreObjetoMenuPulsado = mnuInvertir.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLogin), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmDesembolsoAcreencias.Show
                End If
        Case 12: strNombreObjeto = "frmCancelacionAcreencias"
                gstrNombreObjetoMenuPulsado = mnuInvertir.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLogin), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmCancelacionAcreencias.Show
                End If

        Case 14: strNombreObjeto = "frmSolicitudDescuentoContratos"
                gstrNombreObjetoMenuPulsado = mnuInvertir.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLogin), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmSolicitudDescuentoContratos.Show
                End If
        Case 15: strNombreObjeto = "frmBandejaInversion"
                gstrNombreObjetoMenuPulsado = mnuInvertir.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLogin), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmBandejaInversion.Show
                End If

        Case 16: strNombreObjeto = "frmOrdenDescuentoContratos"
                gstrNombreObjetoMenuPulsado = mnuInvertir.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLogin), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmOrdenDescuentoContratos.Show
                End If

        Case 17: strNombreObjeto = "frmCancelacionDescuentoContratos"
                gstrNombreObjetoMenuPulsado = mnuInvertir.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLogin), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmCancelacionDescuentoContratos.Show
                End If




    End Select
    
End Sub

Private Sub mnuMantenimiento_Click(Index As Integer)
    
    Dim strCodObjeto As String, strNombreObjeto As String

    Select Case Index
        Case 0: strNombreObjeto = "frmTitulosInversion"
                gstrNombreObjetoMenuPulsado = mnuMantenimiento.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmTitulosInversion.Show
                End If

        Case 3: strNombreObjeto = "frmLimite"
                gstrNombreObjetoMenuPulsado = mnuMantenimiento.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmLimite.Show
                End If
        
        Case 5: strNombreObjeto = "frmCostoNegociacion"
                gstrNombreObjetoMenuPulsado = mnuMantenimiento.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmCostoNegociacion.Show
                End If
                
        
        Case 7: strNombreObjeto = "frmPrecioTir"
                gstrNombreObjetoMenuPulsado = mnuMantenimiento.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmPrecioTir.Show
                End If
        
        Case 8: strNombreObjeto = "frmEventoCorporativo"
                gstrNombreObjetoMenuPulsado = mnuMantenimiento.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmEventoCorporativo.Show
                End If

    End Select
    
End Sub

Private Sub mnuOrdenes_Click(Index As Integer)

    Select Case Index
        Case 0
            frmOrdenRentaVariable.Show
        Case 1
            frmOrdenRentaFijaCortoPlazo.Show
        Case 2
            frmOrdenRentaFijaLargoPlazo.Show
        Case 4
            frmOrdenReporteRentaVariable.Show
        Case 6
            frmOrdenCDepositoBancario.Show
    End Select
    
End Sub

Private Sub mnuOperaciones_Click(Index As Integer)
    
    Dim strCodObjeto As String, strNombreObjeto As String

    Select Case Index
        Case 0: strNombreObjeto = "frmQuiebreDeposito"
                gstrNombreObjetoMenuPulsado = mnuOperaciones.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmQuiebreDeposito.Show
                End If
    
        Case 1: strNombreObjeto = "frmCargaOperaciones"
                gstrNombreObjetoMenuPulsado = mnuOperaciones.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmCargaOperaciones.Show
                End If
        
    End Select

End Sub

Private Sub mnuOrdenFinanciamiento_Click(Index As Integer)
    Select Case Index
    
        Case 0: strNombreObjeto = "frmOrdenPrestamo"
                gstrNombreObjetoMenuPulsado = mnuOrdenFinanciamiento.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLogin), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmOrdenPrestamo.Show
                End If
        Case 1: strNombreObjeto = "frmCancelacionPrestamos"
                gstrNombreObjetoMenuPulsado = mnuOrdenFinanciamiento.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLogin), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmCancelacionPrestamos.Show
                End If


    
    End Select
        
End Sub

Private Sub mnuProyectoInmo_Click(Index As Integer)
    Select Case Index
    
        Case 0: strNombreObjeto = "frmProyectoInmobiliarioDefinicion"
                gstrNombreObjetoMenuPulsado = mnuProyectoInmo.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmProyectoInmobiliarioDefinicion.Show
                End If
                
        Case 1: strNombreObjeto = "frmProyectoInmobiliarioConstitucion"
                gstrNombreObjetoMenuPulsado = mnuProyectoInmo.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmProyectoInmobiliarioConstitucion.Show
                End If
                
    End Select
End Sub

Private Sub mnuRegistro_Click(Index As Integer)

    If Index = 99 Then Unload Me
    
    If frmMainMdi.ActiveForm Is Nothing Then
        Exit Sub
    End If
    
    Select Case Index
        Case 0: frmMainMdi.ActiveForm.Adicionar
        Case 1: frmMainMdi.ActiveForm.Modificar
        Case 2: frmMainMdi.ActiveForm.Eliminar
        Case 3: frmMainMdi.ActiveForm.Grabar
        Case 5: frmMainMdi.ActiveForm.Imprimir
    End Select
    
End Sub

Private Sub mnuVer_Click(Index As Integer)

    Select Case Index
        Case 0
            stbMdi.Visible = Not stbMdi.Visible
        Case 1
            tlbMdi.Visible = Not tlbMdi.Visible
    End Select
    frmMainMdi.mnuVer(Index).Checked = Not frmMainMdi.mnuVer(Index).Checked
    
End Sub

Private Sub tlbMdi_ButtonClick(ByVal Button As MSComctlLib.IButton)

    If frmMainMdi.ActiveForm Is Nothing Then
        If Button.Key = "Bloquear" Then
            frmBloqueo.Show vbModal
        End If
        Exit Sub
    End If
    
    Select Case Trim(Button.Key)
    
        Case "Agregar": frmMainMdi.ActiveForm.Adicionar
        Case "Modificar": frmMainMdi.ActiveForm.Modificar
        Case "Guardar": frmMainMdi.ActiveForm.Grabar
        Case "Eliminar": frmMainMdi.ActiveForm.Eliminar
'        Case "Refrescar": frmMainMdi.ActiveForm.Refrescar
        Case "Buscar": frmMainMdi.ActiveForm.Buscar
        Case "Imprimir": frmMainMdi.ActiveForm.Imprimir
'        Case "Primero": frmMainMdi.ActiveForm.ucButNav1.cmdNav_Click (0)
'        Case "Anterior": frmMainMdi.ActiveForm.ucButNav1.cmdNav_Click (1)
'        Case "Siguiente": frmMainMdi.ActiveForm.ucButNav1.cmdNav_Click (2)
'        Case "Ultimo": frmMainMdi.ActiveForm.ucButNav1.cmdNav_Click (3)
'        Case "Bloquear": frmBloqueo.Show vbModal
'        Case "Ayuda": frmMainMdi.ActiveForm.Ayuda
    
    End Select
    
End Sub

Private Sub tlbMdi_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.IButtonMenu)

    If frmMainMdi.ActiveForm Is Nothing Then Exit Sub
    
    Select Case Trim(ButtonMenu.Key)
        Case "Repo1": frmMainMdi.ActiveForm.SubImprimir (1)
        Case "Repo2": frmMainMdi.ActiveForm.SubImprimir (2)
        Case "Repo3": frmMainMdi.ActiveForm.SubImprimir (3)
        Case "Repo4": frmMainMdi.ActiveForm.SubImprimir (4)
        Case "Repo5": frmMainMdi.ActiveForm.SubImprimir (5)
        Case "Repo6": frmMainMdi.ActiveForm.SubImprimir (6)
        Case "Repo7": frmMainMdi.ActiveForm.SubImprimir (7)
        Case "Repo8": frmMainMdi.ActiveForm.SubImprimir (8)
        Case "Repo9": frmMainMdi.ActiveForm.SubImprimir (9)
        Case "Repo10": frmMainMdi.ActiveForm.SubImprimir (10)
    End Select
    
End Sub


