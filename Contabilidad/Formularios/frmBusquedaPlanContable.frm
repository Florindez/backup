VERSION 5.00
Object = "{805EEBCA-BB5B-454E-83AB-BDD03888489E}#1.0#0"; "TAMNetControl.tlb"
Begin VB.Form frmBusquedaPlanContable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de Cuentas"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin TAMNetControlCtl.BusquedaCuentas BusquedaCuentas1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   7695
      Object.Visible         =   "True"
      Enabled         =   "True"
      ForegroundColor =   "-2147483630"
      BackgroundColor =   "-2147483633"
      BackColor       =   "Control"
      ForeColor       =   "ControlText"
      Location        =   "0, 4"
      Name            =   "BusquedaCuentas"
      Size            =   "513, 505"
      Object.TabIndex        =   "0"
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
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
      Left            =   4110
      Picture         =   "frmBusquedaPlanContable.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   1200
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "&Seleccionar"
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
      Left            =   2190
      Picture         =   "frmBusquedaPlanContable.frx":0582
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7680
      Width           =   1200
   End
End
Attribute VB_Name = "frmBusquedaPlanContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    BusquedaCuentas1.CargarArbolCuentas gstrConnectNET, "up_CNLstPlanContable", gstrCodAdministradora
End Sub

Private Sub cmdSeleccionar_Click()
    If BusquedaCuentas1.IdSeleccionado = Valor_Caracter Then
        MsgBox "Debe seleccionar una cuenta", vbOKOnly, Me.Caption
        Exit Sub
    Else
        gstrCodCuentaBusquedaPlanContable = BusquedaCuentas1.IdSeleccionado
        gstrDescripCuentaBusquedaPlanContable = BusquedaCuentas1.TextoSeleccionado
        Unload Me
    End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub


