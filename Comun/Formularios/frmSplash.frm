VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4590
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Width           =   7080
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   2595
         Left            =   210
         ScaleHeight     =   2595
         ScaleWidth      =   3390
         TabIndex        =   7
         Top             =   330
         Width           =   3390
      End
      Begin VB.Timer Timer1 
         Interval        =   150
         Left            =   4080
         Top             =   2280
      End
      Begin VB.Image Image1 
         Height          =   1275
         Left            =   5910
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   3135
         Width           =   1245
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H80000005&
         Caption         =   "Compañía  : TAM Consulting"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3585
         TabIndex        =   3
         Top             =   3300
         Width           =   3315
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H80000005&
         Caption         =   $"frmSplash.frx":8EFA
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   150
         TabIndex        =   2
         Top             =   3585
         Width           =   5880
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Versión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   4
         Top             =   2700
         Width           =   885
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Plataforma"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5580
         TabIndex        =   5
         Top             =   2340
         Width           =   1275
      End
      Begin VB.Label lblProductName 
         BackColor       =   &H80000005&
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3630
         TabIndex        =   6
         Top             =   900
         Width           =   3270
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Autorizado a"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mAlpha As Long

Private Const WS_EX_LAYERED As Long = &H80000
Private Const LWA_ALPHA As Long = &H2

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long


Private Const RDW_INVALIDATE = &H1
Private Const RDW_ERASE = &H4
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_FRAME = &H400

Private Declare Function RedrawWindow2 Lib "user32" Alias "RedrawWindow" (ByVal hwnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Sub Fade(ByVal Enabled As Boolean)
If Enabled = True Then
    Hide
    Timer1.Enabled = True
Else
    Transparentar False, 0
End If
End Sub

Private Sub Transparentar(ByVal Enabled As Boolean, ByVal Porcentaje As Long)
If Enabled = True Then
    Dim tAlpha As Long
    tAlpha = Val(Porcentaje)
    If tAlpha < 1 Or tAlpha > 100 Then
        tAlpha = 70
    End If
    Call SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(hwnd, 0, (255 * tAlpha) / 100, LWA_ALPHA)
Else
    Call SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) And Not WS_EX_LAYERED)
    Call RedrawWindow2(hwnd, 0&, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_FRAME Or RDW_ALLCHILDREN)
End If
End Sub

Private Sub Form_Click()
IniciarLogin
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then IniciarLogin
End Sub

Private Sub Form_Load()
Dim strRutaLogo As String

strRutaLogo = gstrImagePath & "Logotam.jpg" 'App.Path & "\Logo\Logo.jpg"

If Len(dir(strRutaLogo, vbArchive)) <> 0 Then
    Picture1.Picture = LoadPicture(strRutaLogo)
End If

lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision

lblProductName.Caption = "SPECTRUM FONDOS" 'App.Title

Fade True

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then IniciarLogin
End Sub

Private Sub Frame1_Click()
    IniciarLogin
End Sub

Private Sub Timer1_Timer()
Show
mAlpha = mAlpha + 10
If mAlpha <= 100 Then
    Transparentar True, mAlpha
ElseIf mAlpha = 200 Then
    Timer1.Enabled = False
    IniciarLogin
End If
End Sub

Private Sub IniciarLogin()
Dim rst As New ADODB.Recordset
Dim strMsgError As String

On Error GoTo err
Unload Me

'''abrirConexion strMsgError
'''If strMsgError <> "" Then GoTo err
'''
'''frmLogin.Show 1

'frmPrincipal.Show
'frmAcceso.Show
frmMainMdi.Show

Exit Sub
err:
If strMsgError = "" Then strMsgError = err.Description
MsgBox strMsgError, vbInformation, App.Title
End Sub



