VERSION 5.00
Begin VB.Form frmFiltroReporte2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filtro Adicional"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6045
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
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
      Left            =   1740
      Picture         =   "frmFiltroReporte2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1140
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
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
      Left            =   3270
      Picture         =   "frmFiltroReporte2.frx":0485
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1140
      Width           =   1200
   End
   Begin VB.Frame fraCliente 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   5865
      Begin VB.ComboBox cboCliente 
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
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   510
         Width           =   5415
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente a consultar"
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
         Left            =   180
         TabIndex        =   2
         Top             =   210
         Width           =   1605
      End
   End
   Begin VB.Frame fraDias 
      Height          =   2205
      Left            =   90
      TabIndex        =   3
      Top             =   0
      Width           =   2685
      Begin VB.OptionButton rb60 
         Caption         =   "A 60 días"
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   240
         TabIndex        =   8
         Top             =   1650
         Width           =   1155
      End
      Begin VB.OptionButton rb30 
         Caption         =   "A 30 días"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   240
         TabIndex        =   7
         Top             =   1350
         Width           =   1125
      End
      Begin VB.OptionButton rb15 
         Caption         =   "A 15 días"
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   1020
         Width           =   1305
      End
      Begin VB.OptionButton rb7 
         Caption         =   "A 7 días"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   240
         TabIndex        =   5
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label lblDiasVenc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Días hasta el vencimiento:"
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
         Left            =   180
         TabIndex        =   4
         Top             =   330
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmFiltroReporte2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strReporte As String

Public intDias As Integer
Public strCodEmisor As String

Public blnCancelado As Boolean

Private arrCliente() As String
Private strSQL  As String


Private Sub CargarListas()
    Dim adoRecord   As ADODB.Recordset
    Dim intRegistro As Integer

    strSQL = "SELECT (CodPersona) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboCliente, arrCliente(), Sel_Defecto
    
    If cboCliente.ListCount > 0 Then
        cboCliente.ListIndex = 0
    End If

End Sub


Private Sub cboCliente_Click()
    strCodEmisor = arrCliente(cboCliente.ListIndex)

End Sub

Private Sub cmdAceptar_Click()
    blnCancelado = False
    If strReporte = "OperacionesPorVencer" Then
        If rb7.Value Then
            intDias = 7
        ElseIf rb15.Value Then
            intDias = 15
        ElseIf rb30.Value Then
            intDias = 30
        ElseIf rb60.Value Then
            intDias = 60
        End If
        
    End If
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    blnCancelado = True
    'gstrSelFrml = "0"  '** Cancela
    Unload Me
End Sub

Private Sub Form_Activate()
    If strReporte = "OperacionesPorVencer" Then
        fraCliente.Visible = False
        fraDias.Visible = True
        
        cmdAceptar.Left = 60
        cmdAceptar.Top = 2280
        
        cmdCancelar.Left = 1590
        cmdCancelar.Top = 2280
        
        Me.Height = 3525
        Me.Width = 2955
        
    End If
    If strReporte = "HistOperaciones" Then
        fraCliente.Visible = True
        fraDias.Visible = False
        
        cmdAceptar.Left = 1470
        cmdAceptar.Top = 1140
        
        cmdCancelar.Left = 3270
        cmdCancelar.Top = 1140
        
        Me.Height = 2385
        Me.Width = 6135
        
    End If

End Sub

Private Sub Form_Load()
    
    Call CargarListas
    
    CentrarForm Me
End Sub
