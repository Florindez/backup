VERSION 5.00
Object = "{5D05676A-9290-11D7-9297-00047610EA23}#8.0#0"; "TBuscar.ocx"
Begin VB.Form frmBuscar 
   Caption         =   "Buscar Registro"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7035
   Icon            =   "frmBuscar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin TBuscar.TBuscarRegistro TBuscarRegistro1 
      Height          =   5595
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   9869
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents m_TBuscarRegistro As TBuscarRegistro
Attribute m_TBuscarRegistro.VB_VarHelpID = -1

Private Sub cmdAceptar_Click()

End Sub

Private Sub Form_Load()

   Set m_TBuscarRegistro = TBuscarRegistro1
      
   CentrarForm Me

End Sub

