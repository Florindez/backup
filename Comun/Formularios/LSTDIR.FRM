VERSION 5.00
Begin VB.Form frm_ListaDir 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Lista de Directorios"
   ClientHeight    =   4140
   ClientLeft      =   2445
   ClientTop       =   2580
   ClientWidth     =   4200
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4140
   ScaleWidth      =   4200
   Begin VB.DirListBox dir_Lista 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   2340
      Left            =   210
      TabIndex        =   4
      Top             =   810
      Width           =   2445
   End
   Begin VB.DriveListBox drv_Lista 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   180
      TabIndex        =   3
      Top             =   3630
      Width           =   2415
   End
   Begin VB.CommandButton cmd_Accion 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   2940
      TabIndex        =   2
      Top             =   210
      Width           =   1005
   End
   Begin VB.CommandButton cmd_Accion 
      Appearance      =   0  'Flat
      Caption         =   "Cancelar"
      Height          =   405
      Index           =   1
      Left            =   2940
      TabIndex        =   1
      Top             =   690
      Width           =   1005
   End
   Begin VB.TextBox txt_directorio 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      MaxLength       =   40
      TabIndex        =   0
      Top             =   450
      Width           =   2385
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Unidades:"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   210
      TabIndex        =   5
      Top             =   3330
      Width           =   1665
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Directorio:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   180
      Width           =   2385
   End
End
Attribute VB_Name = "frm_ListaDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Accion_Click(index As Integer)
Select Case index
  Case 0
     gs_FormName = Trim(dir_Lista.List(dir_Lista.ListIndex))
     Unload Me
  Case 1
     Unload Me

End Select

End Sub

Private Sub dir_Lista_Click()

Dim s%, t%
Dim s_Path As String
Dim s_LastDir As String  'ultimo directorio...
Dim s_FirstDir As String  'primer(eros) directorio(s)...
Dim s_ChainPto As String

s_ChainPto = "..."

s_Path = Trim(dir_Lista.List(dir_Lista.ListIndex))
s% = Len(s_Path)

'se analiza el ultimo directorio de la ruta
t% = s%
Do Until Mid(s_Path, t%, 1) = "\"
   t% = t% - 1
Loop

s_LastDir = Mid$(s_Path, t%, s% - t% + 1)

If Len(s_LastDir) > txt_directorio.MaxLength Then
   txt_directorio.Text = Mid$(s_LastDir, 1, txt_directorio.MaxLength - Len(s_ChainPto)) + s_ChainPto
Else
   s_FirstDir = Mid$(s_Path, 1, t% - 1)
   If Len(s_FirstDir) > txt_directorio.MaxLength - Len(s_LastDir) Then
      txt_directorio = Mid$(s_FirstDir, 1, txt_directorio.MaxLength - Len(s_LastDir) - Len(s_ChainPto)) + s_ChainPto + s_LastDir
   Else
      txt_directorio = s_Path
   End If
End If



End Sub

Private Sub drv_Lista_Change()

On Error GoTo Errorhandler

dir_Lista.Path = drv_Lista.Drive

Exit Sub

Errorhandler:

MsgBox Error$(err), 48, frm_Lista.Caption
Resume Next


End Sub

Private Sub Form_Load()

CentrarForm Me

End Sub

