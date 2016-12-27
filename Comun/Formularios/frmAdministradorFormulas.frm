VERSION 5.00
Object = "{805EEBCA-BB5B-454E-83AB-BDD03888489E}#1.0#0"; "TAMNetControl.tlb"
Begin VB.Form frmAdministradorFormulas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrador de Formulas"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   11685
   StartUpPosition =   2  'CenterScreen
   Begin TAMNetControlCtl.AdministradorFormulas AdministradorFormulas1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      Object.Visible         =   "True"
      Enabled         =   "True"
      ForegroundColor =   "-2147483630"
      BackgroundColor =   "-2147483633"
      TextoFormula    =   ""
      BackColor       =   "Control"
      ForeColor       =   "ControlText"
      Location        =   "8, 8"
      Name            =   "AdministradorFormulas"
      Size            =   "761, 401"
      Object.TabIndex        =   "0"
   End
   Begin VB.CommandButton cmdVerPlanCuentas 
      Caption         =   "Ver Plan de Cuentas"
      Height          =   735
      Left            =   2760
      Picture         =   "frmAdministradorFormulas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1680
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   8280
      Picture         =   "frmAdministradorFormulas.frx":0565
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   735
      Left            =   5640
      Picture         =   "frmAdministradorFormulas.frx":0AC7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   1200
   End
End
Attribute VB_Name = "frmAdministradorFormulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    gstrTextoAdministradorFormula = Trim(AdministradorFormulas1.TextoFormula)
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdVerPlanCuentas_Click()
    gstrCodCuentaBusquedaPlanContable = Valor_Caracter
    gstrDescripCuentaBusquedaPlanContable = Valor_Caracter
    frmBusquedaPlanContable.Show vbModal
    AdministradorFormulas1.TextoFormula = AdministradorFormulas1.TextoFormula + Valor_Caracter + gstrCodCuentaBusquedaPlanContable
End Sub

Private Sub Form_Load()
    gstrTextoAdministradorFormula = Trim(AdministradorFormulas1.TextoFormula)
End Sub
