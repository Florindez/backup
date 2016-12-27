VERSION 5.00
Begin VB.Form frmCambioClave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambiar Contraseña"
   ClientHeight    =   2580
   ClientLeft      =   1065
   ClientTop       =   2085
   ClientWidth     =   4560
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
   Icon            =   "frmCambioClave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2580
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2730
      TabIndex        =   9
      Top             =   240
      Width           =   1545
   End
   Begin VB.CommandButton cmd_Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   735
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   1200
   End
   Begin VB.CommandButton cmd_Aceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   735
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   1200
   End
   Begin VB.TextBox txant 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2730
      MaxLength       =   100
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1185
      Width           =   1545
   End
   Begin VB.TextBox txant 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2730
      MaxLength       =   100
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   870
      Width           =   1545
   End
   Begin VB.TextBox txant 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   2730
      MaxLength       =   100
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   555
      Width           =   1545
   End
   Begin VB.Label lbl_DesCam 
      AutoSize        =   -1  'True
      Caption         =   "Usuario"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   300
      TabIndex        =   8
      Top             =   270
      Width           =   660
   End
   Begin VB.Label lbl_DesCam 
      AutoSize        =   -1  'True
      Caption         =   "Repetir Contraseña Nueva"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   300
      TabIndex        =   2
      Top             =   1185
      Width           =   2265
   End
   Begin VB.Label lbl_DesCam 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña Nueva"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   1
      Top             =   870
      Width           =   1590
   End
   Begin VB.Label lbl_DesCam 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña Anterior"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   555
      Width           =   1695
   End
End
Attribute VB_Name = "frmCambioClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CambiarClave()
    Dim gstrConnectCambioClave
    Dim adoConnCambioClave As ADODB.Connection
    Dim adoCommCambioClave As ADODB.Command
    
    Dim adoUsuario As ADODB.Recordset
    Dim strUsuario As String
    Dim strPassw As String
    
    On Error GoTo cmdConectar_Error
    

    'Conexion a Seguridad
                                 
    '*** SQLOLEDB - Base de Datos ***
    gstrConnectCambioClave = "User ID=" & gstrLoginSeguridad & ";Password=" & gstrPasswSeguridad & ";" & _
                        "Data Source=" & gstrServerSeguridad & ";" & _
                        "Initial Catalog=" & gstrDataBaseSeguridad & ";" & _
                        "Application Name=" & App.Title & ";" & _
                        "Auto Translate=False"
    
    Set adoConnCambioClave = New ADODB.Connection
    
    frmMainMdi.stbMdi.Panels(3).Text = "Conectando a la Base de Datos..."
    With adoConnCambioClave
        .Provider = "SQLOLEDB"
        .ConnectionString = gstrConnectCambioClave
        .Open
    End With
        
    Set adoCommCambioClave = New ADODB.Command
    adoCommCambioClave.CommandTimeout = 0
    Set adoCommCambioClave.ActiveConnection = adoConnCambioClave
    
     With adoCommCambioClave

        Set adoUsuario = New ADODB.Recordset
        
        '---/// Validacion de Usuario Seguridad
        .CommandText = "SELECT IdUsuario, dbo.uf_SELeerClaveEnch(Passw) Passw FROM UsuarioSistema WHERE IdUsuario = '" & txtUserName.Text & "'"
        Set adoUsuario = .Execute
        strUsuario = adoUsuario("IdUsuario")
        strPassw = adoUsuario("Passw")
        adoUsuario.Close: Set adoUsuario = Nothing
    
     
        If txant(0).Text <> strPassw Then
            MsgBox "Password anterior incorrecto,verifique!", vbCritical
            txant(0).Text = Valor_Caracter
            txant(0).SetFocus
            Exit Sub
        End If
        
        If txant(2).Text <> txant(1).Text Then
            MsgBox "Password confirmado no coincide con el nuevo password.", vbCritical
            txant(2).SetFocus
            Exit Sub
        End If
        
        .CommandText = "{ call up_SECambioClave('" & txtUserName.Text & "','" & txant(0).Text & "','" & txant(1).Text & "')}"
        Set adoUsuario = .Execute
        
    End With
    
    MsgBox "Password Cambiado.", vbExclamation, gstrNombreEmpresa
    
    If adoConnCambioClave.State = 1 Then
        adoConnCambioClave.Close:  Set adoConnCambioClave = Nothing
    End If

    Unload frmCambioClave
    frmAcceso.txtPassword = ""
    
    Exit Sub
    
cmdConectar_Error:
    MousePointer = vbDefault
    With err
        
        frmMainMdi.stbMdi.Panels(3).Text = "Error de conexión..."
        MsgBox "Error! " & .Description, vbCritical, Me.Caption
    End With
    
End Sub

Private Sub cmd_Aceptar_Click()
    Call CambiarClave

End Sub

Private Sub cmd_Cancelar_Click()

    Unload frmCambioClave
    
End Sub

Private Sub Form_Activate()
    txtUserName.SetFocus
End Sub

Private Sub Form_Load()

    Dim intRes As Integer
    
    Me.Refresh
    gstrFechaActual = Convertyyyymmdd(Date)
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
      
   Set frmCambioClave = Nothing
   
End Sub

