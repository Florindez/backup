VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMainMdi 
   BackColor       =   &H8000000C&
   Caption         =   "Módulo Tesorería"
   ClientHeight    =   9960
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   16380
   Icon            =   "frmMainMdi.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Tag             =   "T"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbMdi 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   9585
      Width           =   16380
      _ExtentX        =   28893
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
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
            Object.Width           =   23230
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
      Width           =   16380
      _ExtentX        =   28893
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
      Left            =   2400
      Top             =   1920
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
      Caption         =   "&Mantenimiento"
      Index           =   1
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "Tasas de Interés"
         Index           =   0
      End
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "Cuentas Bancarias del Fondo"
         Index           =   1
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Procesos"
      Index           =   2
      Begin VB.Menu mnuProcesos 
         Caption         =   "Confirmación de Solicitudes"
         Index           =   0
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Control Pago Parcial Suscripciones"
         Index           =   1
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "-"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Liquidación de Operaciones de Rescate"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Liquidación de Operaciones con Retención"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Liquidación de Ordenes de Cobro y Pago"
         Index           =   6
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Movimientos Bancarios"
         Index           =   10
         Begin VB.Menu mnuMovBancarios 
            Caption         =   "Abonos y Retiros"
            Index           =   1
            Visible         =   0   'False
            Begin VB.Menu mnuAbonoRetiro 
               Caption         =   "Clientes"
               Index           =   0
            End
            Begin VB.Menu mnuAbonoRetiro 
               Caption         =   "Proveedores"
               Index           =   1
            End
         End
         Begin VB.Menu mnuMovBancarios 
            Caption         =   "Gastos Bancarios"
            Index           =   2
         End
         Begin VB.Menu mnuMovBancarios 
            Caption         =   "Transferencias entre Cuentas Bancarias"
            Index           =   3
         End
         Begin VB.Menu mnuMovBancarios 
            Caption         =   "Movimientos Cambiarios"
            Index           =   4
         End
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Interfases Archivos Scotiabank"
         Index           =   11
         Begin VB.Menu mnuInterfases 
            Caption         =   "Importación Archivo"
            Index           =   1
         End
         Begin VB.Menu mnuInterfases 
            Caption         =   "Exportación Archivo"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Informes"
      Index           =   3
      Visible         =   0   'False
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Herramientas"
      Index           =   4
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Calculadora"
         Index           =   0
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Ver"
      Index           =   5
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
      Index           =   6
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&Ayuda"
      Index           =   7
      Begin VB.Menu mnuAyuda 
         Caption         =   "Acerca del Módulo de Tesorería"
         Index           =   0
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "MenuPopUp"
      Index           =   8
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
Option Explicit

Private Sub mnuopc62_Click(index As Integer)

    Dim vntRetVal As Variant
    
    '*** Ejecuta Calculadora ***
    vntRetVal = Shell("CALC.EXE", vbNormalFocus)
    
End Sub

Private Sub MDIForm_Load()

    App.Title = "Tesoreria"

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
    
    Call ValidarPermisoUsoMenu(gstrLoginUS, Me, Trim(App.Title), Separador_Codigo_Objeto)
    
    If gboolMostrarSelectAdministradora Then frmSeleccionaAdministradora.Show vbModal
    
    Call OcultarReportes
    
    If gstrLoginUS = "admin" Or gstrLoginUS = "sa" Then
        Me.Caption = gstrServer & "\\" & gstrDataBase & "\\" & Me.Caption
    End If

End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 1 Then
    
        PopupMenu mnuGeneral(8)
    
    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    On Error GoTo CtrlError
    adoConn.Close: Set adoConn = Nothing
    
CtrlError:
    End
    
End Sub

Private Sub mnuAbonoRetiro_Click(index As Integer)

    Dim strCodObjeto As String, strNombreObjeto As String
    
    Select Case index
       
        Case 0: strNombreObjeto = "frmAbonoRetiroCtaCliente"
                gstrNombreObjetoMenuPulsado = mnuAbonoRetiro.Item(index).Name + "(" + CStr(index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmAbonoRetiroCtaCliente.Show
                End If

        Case 1: strNombreObjeto = "frmAbonoRetiroCtaProveedor"
                gstrNombreObjetoMenuPulsado = mnuAbonoRetiro.Item(index).Name + "(" + CStr(index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmAbonoRetiroCtaProveedor.Show
                End If

      End Select

End Sub


Private Sub mnuAyuda_Click(index As Integer)

    Select Case index
        Case 0: frmAcercade.Show vbModal
    End Select
    
End Sub

Private Sub mnuEmergente_Click(index As Integer)
    
    Select Case index
        
         Case 0: frmSeleccionaAdministradora.Show
        
    End Select
    
End Sub

Private Sub mnuHerramientas_Click(index As Integer)

    Select Case index
        Case 0
            Dim lngValorRetorno As Long
    
            '*** Ejecuta Calculadora ***
            lngValorRetorno = Shell("CALC.EXE", vbNormalFocus)
    End Select
    
End Sub


Private Sub mnuInterfases_Click(index As Integer)

    Dim strCodObjeto As String, strNombreObjeto As String
    
    Select Case index

        Case 1: strNombreObjeto = "frmRecaudacionFile"
                gstrNombreObjetoMenuPulsado = mnuInterfases.Item(index).Name + "(" + CStr(index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmRecaudacionFile.Show
                End If

        Case 2: strNombreObjeto = "frmGeneraArchivo"
                gstrNombreObjetoMenuPulsado = mnuInterfases.Item(index).Name + "(" + CStr(index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmGeneraArchivo.Show
                End If

    End Select

End Sub

Private Sub mnuMantenimiento_Click(index As Integer)
    
    Dim strCodObjeto As String, strNombreObjeto As String
    
    Select Case index
        Case 0: strNombreObjeto = "frmTasaInteres"
                gstrNombreObjetoMenuPulsado = mnuMantenimiento.Item(index).Name + "(" + CStr(index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmTasaInteres.Show
                End If
            
        Case 1: strNombreObjeto = "frmCuentaFondo"
                gstrNombreObjetoMenuPulsado = mnuMantenimiento.Item(index).Name + "(" + CStr(index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmCuentaFondo.Show
                End If
                

                
    End Select
    
End Sub


Private Sub mnuMovBancarios_Click(index As Integer)


    Dim strCodObjeto As String, strNombreObjeto As String
    
    Select Case index

        Case 2: strNombreObjeto = "frmMovimientoBancario"
                gstrNombreObjetoMenuPulsado = mnuMovBancarios.Item(index).Name + "(" + CStr(index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmMovimientoBancario.Show
                End If
        
        Case 3: strNombreObjeto = "frmCuentaFondoMovimiento"
                gstrNombreObjetoMenuPulsado = mnuMovBancarios.Item(index).Name + "(" + CStr(index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                    Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmCuentaFondoMovimiento.Show
                End If
        
        Case 4: strNombreObjeto = "frmMovimientoCambiario"
                gstrNombreObjetoMenuPulsado = mnuMovBancarios.Item(index).Name + "(" + CStr(index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                    Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmMovimientoCambiario.Show
                End If
    
    End Select



End Sub

Private Sub mnuProcesos_Click(index As Integer)
    
    Dim strCodObjeto As String, strNombreObjeto As String
    
    Select Case index
        Case 0: strNombreObjeto = "frmConfirmacionSolicitud"
                gstrNombreObjetoMenuPulsado = mnuProcesos.Item(index).Name + "(" + CStr(index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmConfirmacionSolicitud.Show
                End If
    
        Case 1: strNombreObjeto = "frmPagoCuotaSuscripcion"
                gstrNombreObjetoMenuPulsado = mnuProcesos.Item(index).Name + "(" + CStr(index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmPagoCuotaSuscripcion.Show
                End If

        Case 6: strNombreObjeto = "frmLiquidacionOrdenCaja"
                gstrNombreObjetoMenuPulsado = mnuProcesos.Item(index).Name + "(" + CStr(index) + ")"
                strCodObjeto = Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + _
                                        Separador_Codigo_Objeto + strNombreObjeto
                If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), strCodObjeto, Codigo_Tipo_Objeto_Formulario) Then
                    Exit Sub
                Else
                    frmLiquidacionOrdenCaja.Show
                End If
 
        
    End Select
    
End Sub

Private Sub mnuRegistro_Click(index As Integer)

    If index = 99 Then Unload Me
    
    If frmMainMdi.ActiveForm Is Nothing Then
        Exit Sub
    End If
    
    Select Case index
        Case 0: frmMainMdi.ActiveForm.Adicionar
        Case 1: frmMainMdi.ActiveForm.Modificar
        Case 2: frmMainMdi.ActiveForm.Eliminar
        Case 3: frmMainMdi.ActiveForm.Grabar
        Case 5: frmMainMdi.ActiveForm.Imprimir
    End Select
    
End Sub

Private Sub mnuVer_Click(index As Integer)

    Select Case index
        Case 0
            stbMdi.Visible = Not stbMdi.Visible
        Case 1
            tlbMdi.Visible = Not tlbMdi.Visible
    End Select
    frmMainMdi.mnuVer(index).Checked = Not frmMainMdi.mnuVer(index).Checked
    
End Sub

Private Sub tlbMdi_ButtonClick(ByVal Button As MSComctlLib.IButton)

    If frmMainMdi.ActiveForm Is Nothing Then
        If Button.Key = "Bloquear" Then
            'frmBloqueo.Show vbModal
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
        'Case "Bloquear": frmBloqueo.Show vbModal
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


