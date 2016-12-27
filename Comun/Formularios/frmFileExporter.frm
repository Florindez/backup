VERSION 5.00
Object = "{05B9F8C4-05D2-11D1-A081-444553540000}#1.0#0"; "newex.ocx"
Begin VB.Form frmFileExplorer 
   BackColor       =   &H8000000B&
   Caption         =   "Explorador de archivos"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
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
      Left            =   6960
      Picture         =   "frmFileExporter.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2880
      Width           =   1200
   End
   Begin VB.CommandButton Command16 
      Caption         =   "InitialDir"
      Height          =   270
      Left            =   7665
      TabIndex        =   28
      Top             =   4995
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   26
      Top             =   3120
      Width           =   5055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6480
      TabIndex        =   24
      Text            =   "Text2"
      Top             =   5520
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command15 
      Caption         =   "On"
      Height          =   255
      Left            =   4080
      TabIndex        =   21
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Off"
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Off"
      Height          =   255
      Left            =   5040
      TabIndex        =   19
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      Caption         =   "On"
      Height          =   255
      Left            =   4080
      TabIndex        =   18
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command11 
      Caption         =   "On"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Off"
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Off"
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "On"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7680
      TabIndex        =   11
      Text            =   "C:\windows\"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "On"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Off"
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "On"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Off"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Off"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "On"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BrowseFrom"
      Height          =   255
      Left            =   7680
      TabIndex        =   1
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7680
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin NEWEXLib.ExplorerList ExplorerList1 
      Height          =   2535
      Left            =   4080
      TabIndex        =   30
      Top             =   120
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   4471
      _StockProps     =   161
      BackColor       =   -2147483643
      Appearance      =   1
      ShowMenu        =   0   'False
      Startexe        =   0   'False
   End
   Begin NEWEXLib.ExplorerTree ExplorerTree1 
      Height          =   2535
      Left            =   240
      TabIndex        =   31
      Top             =   120
      Width           =   3615
      _Version        =   65536
      _ExtentX        =   6376
      _ExtentY        =   4471
      _StockProps     =   161
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label Label9 
      Caption         =   "Filename"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Selected Dir :"
      Height          =   255
      Left            =   5520
      TabIndex        =   25
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "3D View"
      Height          =   255
      Left            =   4200
      TabIndex        =   23
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Border"
      Height          =   255
      Left            =   4200
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Border"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "3D View"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "LinesAtRoot"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "TreeHasLines"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "TreeHasButtons"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmFileExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSeleccionar_Click()
    Select Case Index
        Case 0
           gs_FormName = Trim(Text3.Text)
           Unload Me
        Case 1
           Unload Me
    End Select
End Sub

Private Sub Combo1_Click()
ExplorerList1.View = Combo1.ListIndex
End Sub

Private Sub Command1_Click()
ExplorerTree1.BrowseFrom = Text1.Text
End Sub

Private Sub Command10_Click()
ExplorerTree1.BorderStyle = 0
End Sub

Private Sub Command11_Click()
ExplorerTree1.BorderStyle = 1
End Sub

Private Sub Command12_Click()
ExplorerList1.BorderStyle = 1
End Sub

Private Sub Command13_Click()
ExplorerList1.BorderStyle = 0
End Sub

Private Sub Command14_Click()
ExplorerList1.Appearance = 0
End Sub

Private Sub Command15_Click()
ExplorerList1.Appearance = 1
End Sub

Private Sub Command16_Click()
ExplorerTree1.InitialDir = Text1.Text
End Sub

Private Sub Command2_Click()
ExplorerTree1.TreeHasButtons = True
End Sub

Private Sub Command3_Click()
ExplorerTree1.TreeHasButtons = False
End Sub

Private Sub Command4_Click()
ExplorerTree1.TreeHasLines = False
End Sub

Private Sub Command5_Click()
ExplorerTree1.TreeHasLines = True
End Sub

Private Sub Command6_Click()
ExplorerTree1.TreeLinesatRoot = False
End Sub

Private Sub Command7_Click()
ExplorerTree1.TreeLinesatRoot = True
End Sub

Private Sub Command8_Click()
ExplorerTree1.Appearance = 1
End Sub

Private Sub Command9_Click()
ExplorerTree1.Appearance = 0
End Sub

Private Sub ExplorerList1_FolderClick()
ExplorerTree1.FolderClick (ExplorerList1.FileName)
End Sub

Private Sub ExplorerList1_GetFileName()
Text3.Text = ExplorerList1.FileName
End Sub

Private Sub ExplorerTree1_OnDirChanged()
Text2.Text = ExplorerTree1.Path
End Sub

Private Sub ExplorerTree1_TreeDataChanged()
On Error Resume Next
ExplorerList1.TreeDatas = ExplorerTree1.TreeDatas

End Sub

Private Sub Form_Load()
ExplorerTree1.InitialDir = "c:\windows\system"

'Combo1.ListIndex = 2
End Sub

