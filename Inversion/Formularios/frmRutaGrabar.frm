VERSION 5.00
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmRutaGrabar 
   Caption         =   "Guardar"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frCargaPrecio 
      Height          =   3195
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   8445
      Begin TAMControls2.ucBotonEdicion2 cmdSalir 
         Height          =   735
         Left            =   6120
         TabIndex        =   9
         Top             =   2160
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1296
         Caption0        =   "&Salir"
         Tag0            =   "9"
         ToolTipText0    =   "Salir"
         UserControlWidth=   1200
      End
      Begin VB.Frame frmCarga 
         Caption         =   "Guardar en ... "
         Height          =   1785
         Left            =   300
         TabIndex        =   1
         Top             =   330
         Width           =   7845
         Begin VB.TextBox txtNombreArchivo 
            Height          =   315
            Left            =   1110
            TabIndex        =   8
            Top             =   1140
            Width           =   4365
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "Ok"
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
            Left            =   5880
            Picture         =   "frmRutaGrabar.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   960
            Width           =   1125
         End
         Begin VB.TextBox txtArchivo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1110
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   630
            Width           =   5985
         End
         Begin VB.CommandButton cmdBusqueda 
            Caption         =   "..."
            Height          =   315
            Left            =   7140
            TabIndex        =   2
            ToolTipText     =   "Búsqueda de Partícipe"
            Top             =   630
            Width           =   375
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   390
            TabIndex        =   7
            Top             =   1200
            Width           =   555
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Ruta "
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   390
            TabIndex        =   5
            Top             =   660
            Width           =   390
         End
      End
      Begin TAMControls.ucBotonEdicion cmdSalir2 
         Height          =   390
         Left            =   6150
         TabIndex        =   6
         Top             =   2250
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   688
         Caption0        =   "&Salir"
         Tag0            =   "9"
         ToolTipText0    =   "Salir"
         UserControlHeight=   390
         UserControlWidth=   1200
      End
   End
End
Attribute VB_Name = "frmRutaGrabar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBusqueda_Click()
'    gs_FormName = ""
'    frmFileExplorer.Show vbModal
'
'    If Trim(gs_FormName) <> "" Then txtArchivo.Text = gs_FormName


    frm_ListaDir.Show 1
    
    If gs_FormName = Valor_Caracter Then Exit Sub
    
    If gs_FormName <> "" And Mid(gs_FormName, Len(gs_FormName)) <> "\" Then
        gs_FormName = gs_FormName + "\"
        txtArchivo.Text = gs_FormName
    Else
        txtArchivo.Text = gs_FormName
    End If

End Sub

Private Sub Form_Load()
    Set cmdSalir.FormularioActivo = Me
        
    CentrarForm Me
End Sub

Private Sub cmdOk_Click()
    
    If TodoOk() Then
    
        frmFormulario.indOk = True
        gs_FormName = gs_FormName + Trim(txtNombreArchivo.Text)
        Unload Me
        DoEvents
    
    
    End If
    
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
    
        Case vExit
            Call Salir
            
    End Select
            
End Sub

Private Sub Salir()
    
    frmFormulario.indOk = False
    Unload Me

End Sub

Private Function TodoOk() As Boolean

    TodoOk = False

    If txtArchivo.Text = Valor_Caracter Then
        MsgBox "No ah especificado la ruta", vbCritical
        txtArchivo.SetFocus
        Exit Function
    End If
    
    If txtNombreArchivo.Text = Valor_Caracter Then
        MsgBox "No ah especificado el nombre del archivo", vbCritical
        txtNombreArchivo.SetFocus
        Exit Function
    End If
    
    TodoOk = True

End Function



