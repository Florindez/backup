VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMainMdi 
   BackColor       =   &H8000000C&
   Caption         =   "Módulo General"
   ClientHeight    =   9810
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15960
   Icon            =   "frmMainMdi.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Tag             =   "G"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbMdi 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   9435
      Width           =   15960
      _ExtentX        =   28152
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
            Picture         =   "frmMainMdi.frx":030A
            Object.ToolTipText     =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   22490
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
      Width           =   15960
      _ExtentX        =   28152
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
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Usuario"
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
         Width           =   4800
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
      Left            =   2580
      Top             =   1860
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
            Picture         =   "frmMainMdi.frx":048F
            Key             =   "NUEVO"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":2199
            Key             =   "GUARDAR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":2A73
            Key             =   "BUSCAR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":2EC7
            Key             =   "CONSULTAR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":4BD1
            Key             =   "IMPRIMIR"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":68DB
            Key             =   "ELIMINAR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":772D
            Key             =   "BLOQUEAR"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":857F
            Key             =   "REPORTES"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":93D1
            Key             =   "AYUDA"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":96EB
            Key             =   "CANCELAR"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":9A05
            Key             =   "PRIMERO"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":A2DF
            Key             =   "ANTERIOR"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":ABB9
            Key             =   "SIGUIENTE"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":B493
            Key             =   "ULTIMO"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMdi.frx":BD6D
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
      Caption         =   "&Mantenimiento"
      Index           =   2
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "&Fondos"
         Index           =   0
         Begin VB.Menu mnuFondo 
            Caption         =   "&Definición"
            Index           =   0
         End
         Begin VB.Menu mnuFondo 
            Caption         =   "&Horas  de Control"
            Index           =   1
         End
         Begin VB.Menu mnuFondo 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuFondo 
            Caption         =   "Autorización"
            Index           =   3
            Begin VB.Menu mnuAutorizacion 
               Caption         =   "Gastos"
               Index           =   0
            End
            Begin VB.Menu mnuAutorizacion 
               Caption         =   "Ingresos"
               Index           =   1
            End
            Begin VB.Menu mnuAutorizacion 
               Caption         =   "Inversiones"
               Index           =   2
            End
            Begin VB.Menu mnuAutorizacion 
               Caption         =   "Activos Fijos"
               Index           =   3
            End
         End
         Begin VB.Menu mnuFondo 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuFondo 
            Caption         =   "&Limites de Inversión"
            Index           =   5
         End
         Begin VB.Menu mnuFondo 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuFondo 
            Caption         =   "C&omisionistas"
            Index           =   7
         End
         Begin VB.Menu mnuFondo 
            Caption         =   "&Comisiones"
            Index           =   8
            Begin VB.Menu mnuComisiones 
               Caption         =   "SAFI"
               Index           =   0
            End
            Begin VB.Menu mnuComisiones 
               Caption         =   "Participes"
               Index           =   1
            End
         End
         Begin VB.Menu mnuFondo 
            Caption         =   "&Pago de Cuotas"
            Index           =   9
         End
         Begin VB.Menu mnuFondo 
            Caption         =   "&Gastos"
            Index           =   10
         End
         Begin VB.Menu mnuFondo 
            Caption         =   "&Ingresos"
            Index           =   11
         End
         Begin VB.Menu mnuFondo 
            Caption         =   "&Activos Fijos"
            Index           =   12
         End
         Begin VB.Menu mnuFondo 
            Caption         =   "Porcenta&jes de Castigo"
            Index           =   13
         End
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "Instituciones"
         Index           =   1
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "Personas Relacionadas"
         Index           =   2
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "Tasas"
         Index           =   4
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "Calendario"
         Index           =   5
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "Tipo de cambio"
         Index           =   6
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "Catálogos"
         Index           =   8
         Begin VB.Menu mnuCatalogo 
            Caption         =   "General"
            Index           =   0
         End
         Begin VB.Menu mnuCatalogo 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuCatalogo 
            Caption         =   "Estructura de Limites"
            Index           =   2
         End
         Begin VB.Menu mnuCatalogo 
            Caption         =   "Monedas"
            Index           =   3
         End
         Begin VB.Menu mnuCatalogo 
            Caption         =   "Comisiones"
            Index           =   4
         End
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "Tablas del Sistema"
         Index           =   9
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "Parámetros Globales"
         Index           =   10
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "Configuracion de Vistas"
         Enabled         =   0   'False
         Index           =   11
         Visible         =   0   'False
         Begin VB.Menu mnuVistas 
            Caption         =   "Variables de Usuario"
            Index           =   0
         End
         Begin VB.Menu mnuVistas 
            Caption         =   "Vistas Procesos"
            Index           =   1
         End
         Begin VB.Menu mnuVistas 
            Caption         =   "Vistas Usuarios"
            Index           =   2
         End
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "Administrador de Fórmulas"
         Index           =   12
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "Feriados"
         Index           =   13
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Procesos"
      Index           =   3
      Begin VB.Menu mnuProcesos 
         Caption         =   "Cierre Partícipes"
         Index           =   0
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Cierre "
         Index           =   1
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Cierre Mensual"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Backup y Restore de Base de Datos"
         Index           =   3
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Reproceso de Cierre Diario"
         Index           =   4
         Visible         =   0   'False
         Begin VB.Menu mnuReproceso 
            Caption         =   "Apertura de Fecha"
            Index           =   0
         End
         Begin VB.Menu mnuReproceso 
            Caption         =   "Reproceso"
            Index           =   1
         End
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Distribución de Utilidades de Participes"
         Index           =   5
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Cierre de Utilidades de Partícipes"
         Index           =   6
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Cambio de Fecha del Sistema"
         Index           =   7
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Reapertura de Periodo Contable"
         Index           =   8
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Configuracion de Cambio de Fecha"
         Index           =   9
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Seguridad"
      Index           =   4
      Visible         =   0   'False
      Begin VB.Menu mnuSeguridad 
         Caption         =   "Perfiles"
         Index           =   0
      End
      Begin VB.Menu mnuSeguridad 
         Caption         =   "Usuarios"
         Index           =   1
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Herramientas"
      Index           =   5
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Calculadora"
         Index           =   0
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Conf. Comprobante Cobro"
         Index           =   1
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Ver"
      Index           =   6
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
      Index           =   7
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Ayuda"
      Index           =   8
      Begin VB.Menu mnuAyuda 
         Caption         =   "Acerca del Módulo General"
         Index           =   0
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "Menu PopUp"
      Index           =   9
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
Private Sub MDIForm_Load()

    App.Title = "General"
    
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
    
        '---/// Integración de seguridad
    Call ValidarPermisoUsoMenu(gstrLoginUS, Me, Trim(App.Title), Separador_Codigo_Objeto)
    
    If gboolMostrarSelectAdministradora Then frmSeleccionaAdministradora.Show vbModal
    
    Call OcultarReportes
    
    If gstrLoginUS = "admin" Or gstrLoginUS = "sa" Then
        Me.Caption = gstrServer & "\\" & gstrDataBase & "\\" & Me.Caption
    End If
    
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 1 Then
    
        PopupMenu mnuGeneral(9)
    
    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    On Error GoTo CtrlError
    adoConn.Close: Set adoConn = Nothing
    
CtrlError:
    End
    
End Sub

Private Sub mnuAutorizacion_Click(Index As Integer)
    
    Dim strCodObjeto As String, strNombreObjeto As String

    Select Case Index
    
        Case 0: strNombreObjeto = "frmFondoConceptoGasto"
                gstrNombreObjetoMenuPulsado = mnuAutorizacion.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmFondoConceptoGasto.Show
                End If

        Case 1: strNombreObjeto = "frmFondoConceptoIngreso"
                gstrNombreObjetoMenuPulsado = mnuAutorizacion.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmFondoConceptoIngreso.Show
                End If
   
        Case 2: strNombreObjeto = "frmFondoInstrumentos"
                gstrNombreObjetoMenuPulsado = mnuAutorizacion.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmFondoInstrumentos.Show
                End If


        Case 3: strNombreObjeto = "frmFondoConceptoActivoFijo"
                gstrNombreObjetoMenuPulsado = mnuAutorizacion.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmFondoConceptoActivoFijo.Show
                End If
         
    End Select
    
End Sub

Private Sub mnuAyuda_Click(Index As Integer)

    Select Case Index
        Case 0: frmAcercade.Show vbModal
    End Select
    
End Sub

Private Sub mnuCatalogo_Click(Index As Integer)
    
    Dim strCodObjeto As String, strNombreObjeto As String

    Select Case Index
        Case 0: strNombreObjeto = "frmCatalogo"
                gstrNombreObjetoMenuPulsado = mnuCatalogo.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmCatalogo.Show
                End If

        Case 2: strNombreObjeto = "frmLimite"
                gstrNombreObjetoMenuPulsado = mnuCatalogo.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmLimite.Show
                End If

        Case 3: strNombreObjeto = "frmMoneda"
                gstrNombreObjetoMenuPulsado = mnuCatalogo.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmMoneda.Show
                End If

        Case 4: strNombreObjeto = "frmComision"
                gstrNombreObjetoMenuPulsado = mnuCatalogo.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmComision.Show
                End If

    End Select
    
End Sub

Private Sub mnuComisiones_Click(Index As Integer)
    
    Dim strCodObjeto As String, strNombreObjeto As String

    Select Case Index
        Case 1: strNombreObjeto = "frmFondoComisionOperacion"
                gstrNombreObjetoMenuPulsado = mnuComisiones.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmFondoComisionOperacion.Show
                End If

        Case 0: strNombreObjeto = "frmFondoComision"
                gstrNombreObjetoMenuPulsado = mnuComisiones.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmFondoComision.Show
                End If
 
    End Select
    
End Sub

Private Sub mnuEmergente_Click(Index As Integer)
    
    Select Case Index
        
        Case 0: frmSeleccionaAdministradora.Show
    
    End Select
    
End Sub

Private Sub mnuFondo_Click(Index As Integer)
    
    Dim strCodObjeto As String, strNombreObjeto As String
    
    Select Case Index
        Case 0: strNombreObjeto = "frmFondos"
                gstrNombreObjetoMenuPulsado = mnuFondo.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmFondos.Show
                End If

        Case 1: strNombreObjeto = "frmCambioHorarioFondo"
                gstrNombreObjetoMenuPulsado = mnuFondo.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmCambioHorarioFondo.Show
                End If
        
        Case 5: strNombreObjeto = "frmPoliticaInversion"
                gstrNombreObjetoMenuPulsado = mnuFondo.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmPoliticaInversion.Show
                End If

        Case 7: strNombreObjeto = "frmFondoComisionista"
                gstrNombreObjetoMenuPulsado = mnuFondo.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmFondoComisionista.Show
                End If
        
        Case 9: strNombreObjeto = "frmFondoPagoSuscripcion"
                gstrNombreObjetoMenuPulsado = mnuFondo.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmFondoPagoSuscripcion.Show
                End If

        Case 10: strNombreObjeto = "frmFondoGastos"
                gstrNombreObjetoMenuPulsado = mnuFondo.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmFondoGastos.Show
                End If
        
        Case 11:    strNombreObjeto = "frmFondoIngreso"
                    gstrNombreObjetoMenuPulsado = mnuFondo.Item(Index).Name + "(" + CStr(Index) + ")"
                    strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                            Separador_Codigo_Objeto + strNombreObjeto
                    If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                        Exit Sub
                    Else
                        frmFondoIngreso.Show
                    End If

        Case 12:    strNombreObjeto = "frmFondoActivoFijo"
                    gstrNombreObjetoMenuPulsado = mnuFondo.Item(Index).Name + "(" + CStr(Index) + ")"
                    strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                            Separador_Codigo_Objeto + strNombreObjeto
                    If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                        Exit Sub
                    Else
                        frmFondoActivoFijo.Show
                    End If
        
        Case 13:    strNombreObjeto = "frmProvisionCastigos"
                    gstrNombreObjetoMenuPulsado = mnuFondo.Item(Index).Name + "(" + CStr(Index) + ")"
                    strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                            Separador_Codigo_Objeto + strNombreObjeto
                    If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                        Exit Sub
                    Else
                        frmProvisionCastigos.Show
                    End If
    End Select
    
End Sub



Private Sub mnuHerramientas_Click(Index As Integer)

    Select Case Index
        Case 0
            Dim lngValorRetorno As Long
    
            '*** Ejecuta Calculadora ***
            lngValorRetorno = Shell("CALC.EXE", vbNormalFocus)
            
        Case 1
            frmConfComprobanteCobro.Show
    End Select
    
End Sub



Private Sub mnuMantenimiento_Click(Index As Integer)

    Dim strCodObjeto As String, strNombreObjeto As String
    
    Select Case Index
        Case 0  'Fondos
        Case 1: strNombreObjeto = "frmInstitucion"
                gstrNombreObjetoMenuPulsado = mnuMantenimiento.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmInstitucion.Show
                End If

        Case 2: strNombreObjeto = "frmRelacionados"
                gstrNombreObjetoMenuPulsado = mnuMantenimiento.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmRelacionados.Show
                End If

        Case 4: strNombreObjeto = "frmTasas"
                gstrNombreObjetoMenuPulsado = mnuMantenimiento.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmTasas.Show
                End If

        Case 5: strNombreObjeto = "frmCalendario"
                gstrNombreObjetoMenuPulsado = mnuMantenimiento.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
              If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmCalendario.Show
                End If

        Case 6: strNombreObjeto = "frmTipoCambio"
                gstrNombreObjetoMenuPulsado = mnuMantenimiento.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmTipoCambio.Show
                End If
        
        Case 8 '*** Catálogos ***
        Case 9: strNombreObjeto = "frmTablas"
                gstrNombreObjetoMenuPulsado = mnuMantenimiento.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
               If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmTablas.Show
                End If

        Case 10: strNombreObjeto = "frmFondoParametro"
                gstrNombreObjetoMenuPulsado = mnuMantenimiento.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmFondoParametro.Show
                End If

'        Case 11:
'            Dim oMTG As New OSIMTG.clsTablaGenerica
'            Screen.MousePointer = vbHourglass
'            Set oMTG.ActiveConnection = adoConn
'            oMTG.Administrar ("") ' ("TBL084")

        Case 12: strNombreObjeto = "frmFormulaMant"
                gstrNombreObjetoMenuPulsado = mnuMantenimiento.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmFormulaMant.Show
                End If
                
        Case 13: strNombreObjeto = "frmFeriados"
                gstrNombreObjetoMenuPulsado = mnuMantenimiento.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmFeriados.Show
                End If
        
        
    End Select
    
End Sub



Private Sub mnuProcesos_Click(Index As Integer)
    
    Dim strCodObjeto As String, strNombreObjeto As String
    
    Select Case Index
        Case 0: strNombreObjeto = "frmCierreParticipes"
                gstrNombreObjetoMenuPulsado = mnuProcesos.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmCierreParticipes.Show
                End If
        
        Case 1: strNombreObjeto = "frmCierreDiario"
                gstrNombreObjetoMenuPulsado = mnuProcesos.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmCierreDiario.Show
                End If
                
          Case 2: strNombreObjeto = "frmCierreMensual"
                gstrNombreObjetoMenuPulsado = mnuProcesos.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmCierreMensual.Show
                End If
        
        Case 3: strNombreObjeto = "frmBackupRestore"
                gstrNombreObjetoMenuPulsado = mnuProcesos.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmBackupRestore.Show
                End If

        Case 5: strNombreObjeto = "frmDistribucionUtilidades"
                gstrNombreObjetoMenuPulsado = mnuProcesos.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmDistribucionUtilidades.Show
                End If
        
        Case 6: strNombreObjeto = "frmCierreUtilidades"
                gstrNombreObjetoMenuPulsado = mnuProcesos.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmCierreUtilidades.Show
                End If
        
        Case 7: strNombreObjeto = "frmControlFecha"
                gstrNombreObjetoMenuPulsado = mnuProcesos.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmControlFecha.Show
                End If
        
        Case 8: strNombreObjeto = "frmPeriodoContableReApertura"
                gstrNombreObjetoMenuPulsado = mnuProcesos.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmPeriodoContableReApertura.Show
                End If

        Case 9: strNombreObjeto = "frmCron"
                gstrNombreObjetoMenuPulsado = mnuProcesos.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmCron.Show
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

Private Sub mnuReproceso_Click(Index As Integer)
    
    Dim strCodObjeto As String, strNombreObjeto As String
    
    Select Case Index
        Case 0: strNombreObjeto = "frmAsignaFechaReproceso"
                gstrNombreObjetoMenuPulsado = mnuReproceso.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
'                If Not ValidarPermisoAccesoObjeto(Trim(gstrLogin), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
'                    Exit Sub
'                Else
                    frmAsignaFechaReproceso.Show
'                End If

         Case 1: strNombreObjeto = "frmAsignaFechaReproceso"
                gstrNombreObjetoMenuPulsado = mnuReproceso.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
'                If Not ValidarPermisoAccesoObjeto(Trim(gstrLogin), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
'                    Exit Sub
'                Else
                    'frmCierreDiario.Tag = "R"
                    frmAsignaFechaReproceso.Show 'vbModal
'                End If
           
    End Select

End Sub

Private Sub mnuSeguridad_Click(Index As Integer)

    Select Case Index
        Case 0: frmPerfil.Show
        Case 1: frmUsuarios.Show
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

Private Sub mnuVistas_Click(Index As Integer)
    
    Dim strCodObjeto As String, strNombreObjeto As String
    
    Select Case Index
        Case 0: strNombreObjeto = "frmVariableUsuario"
                gstrNombreObjetoMenuPulsado = mnuVistas.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmVariableUsuario.Show
                End If
        
        Case 1: strNombreObjeto = "frmVistaProceso"
                gstrNombreObjetoMenuPulsado = mnuVistas.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmVistaProceso.Show
                End If
        
        Case 2: strNombreObjeto = "frmVistaUsuario"
                gstrNombreObjetoMenuPulsado = mnuVistas.Item(Index).Name + "(" + CStr(Index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmVistaUsuario.Show
                End If

    End Select
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
        'Case "Refrescar": frmMainMdi.ActiveForm.Refrescar
        Case "Buscar": frmMainMdi.ActiveForm.Buscar
        Case "Imprimir": frmMainMdi.ActiveForm.Imprimir
        'Case "Primero": frmMainMdi.ActiveForm.ucButNav1.cmdNav_Click (0)
        'Case "Anterior": frmMainMdi.ActiveForm.ucButNav1.cmdNav_Click (1)
        'Case "Siguiente": frmMainMdi.ActiveForm.ucButNav1.cmdNav_Click (2)
        'Case "Ultimo": frmMainMdi.ActiveForm.ucButNav1.cmdNav_Click (3)
        Case "Bloquear": frmBloqueo.Show vbModal
        Case "Ayuda": frmMainMdi.ActiveForm.Ayuda
    
    
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




