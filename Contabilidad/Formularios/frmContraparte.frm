VERSION 5.00
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmContracuenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contracuenta"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   13245
   StartUpPosition =   3  'Windows Default
   Begin TAMControls2.ucBotonEdicion2 cmdAccion 
      Height          =   735
      Left            =   9360
      TabIndex        =   1
      Top             =   1800
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1296
      Buttons         =   2
      Caption0        =   "&Aceptar"
      Tag0            =   "2"
      ToolTipText0    =   "Guardar"
      Caption1        =   "&Cancelar"
      Tag1            =   "8"
      ToolTipText1    =   "Cancelar"
      UserControlWidth=   2700
   End
   Begin VB.Frame fraContracuenta 
      Caption         =   "Contraparte"
      Height          =   1665
      Left            =   -30
      TabIndex        =   0
      Top             =   60
      Width           =   13245
      Begin VB.CommandButton cmdBusqueda 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   12540
         TabIndex        =   7
         ToolTipText     =   "Buscar Cuenta Contable"
         Top             =   570
         Width           =   375
      End
      Begin VB.TextBox txtCodFile 
         Height          =   315
         Left            =   10290
         MaxLength       =   3
         TabIndex        =   6
         Top             =   930
         Width           =   555
      End
      Begin VB.TextBox txtCodAnalitica 
         Height          =   315
         Left            =   10920
         MaxLength       =   8
         TabIndex        =   5
         Top             =   930
         Width           =   1605
      End
      Begin VB.CommandButton cmdBusqueda 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   12540
         TabIndex        =   4
         ToolTipText     =   "Buscar Cuenta Contable"
         Top             =   930
         Width           =   375
      End
      Begin VB.TextBox txtDescripCuenta 
         Height          =   315
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   570
         Width           =   10815
      End
      Begin VB.TextBox txtDescripFileAnalitica 
         Height          =   315
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   930
         Width           =   8535
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Cuenta"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   17
         Left            =   360
         TabIndex        =   9
         Top             =   615
         Width           =   975
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Analítica"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   18
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmContracuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql                                   As String

Public strTipoFileContracuenta               As String
Public strCodFondo                           As String
Public strCodFileAnaliticaContracuenta       As String
Public strDescripContracuenta                As String
Public strDescripFileAnaliticaContracuenta   As String
Public strCodContracuenta                    As String
Public strCodFileContracuenta                As String
Public strCodAnaliticaContracuenta           As String
Public blnOK                                 As Boolean



Private Sub cmdBusqueda_Click(index As Integer)
   
    Dim frmBus As frmBuscar
    
    Set frmBus = New frmBuscar
    
    With frmBus.TBuscarRegistro1
           
        .ADOConexion = adoConn
        .ADOConexion.CommandTimeout = 0
        .iTipoGrilla = 2
        
        Select Case index
        
            Case 0
            
                
                frmBus.Caption = " Relación de Cuentas Contables"
                .sSql = "SELECT CodCuenta,DescripCuenta,TipoFile,IndAuxiliar,TipoAuxiliar FROM PlanContable "
                .sSql = .sSql & " WHERE IndMovimiento='" & Valor_Indicador & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND NumVersion = dbo.uf_CNObtenerPlanContableVigente('" & gstrCodAdministradora & "') ORDER BY CodCuenta"
                .OutputColumns = "1,2,3,4,5"
                .HiddenColumns = "3,4,5"
                
            Case 1
        
                frmBus.Caption = " Relación de File Analiticas"
                .sSql = "{ call up_CNSelFileAnalitico('" & strCodFondo & "','" & gstrCodAdministradora & "','" & strCodContracuenta & "','" & strTipoFileContracuenta & "') }"
                .OutputColumns = "1,2,3,4,5"
                .HiddenColumns = ""
                
                         
       
        End Select
                
        Screen.MousePointer = vbHourglass
                
        .BuscarTabla
        
        Screen.MousePointer = vbNormal
        frmBus.Show 1
       
        If .iParams.Count = 0 Then Exit Sub
        
        If .iParams(1).Valor <> "" Then
            
        
            Select Case index
            
                Case 0
                
                    strTipoFileContracuenta = Trim(.iParams(3).Valor)
'                    strIndAuxiliar = Trim(.iParams(4).Valor)
'                    strTipoAuxiliar = Trim(.iParams(5).Valor)
                    
                    strCodContracuenta = Trim(.iParams(1).Valor)
                    
                    txtDescripCuenta.Text = Trim(.iParams(1).Valor) & " - " & Trim(.iParams(2).Valor)
                    
                    strDescripContracuenta = txtDescripCuenta.Text
                    
                    
                    strCodFileContracuenta = Valor_Caracter
                    strCodAnaliticaContracuenta = Valor_Caracter
                    strDescripFileAnaliticaContracuenta = Valor_Caracter
                                     
                    txtCodFile.Text = strCodFileContracuenta
                    txtCodAnalitica.Text = strCodAnaliticaContracuenta
                    txtDescripFileAnalitica.Text = Valor_Caracter
                                     
                    If strTipoFileContracuenta = Valor_Caracter Then
                        cmdBusqueda(1).Enabled = False
                    Else
                        cmdBusqueda(1).Enabled = True
                    End If
                    
                Case 1
            
                    strCodFileContracuenta = Trim(.iParams(1).Valor)
                    strCodAnaliticaContracuenta = Trim(.iParams(2).Valor)
                    strDescripFileAnaliticaContracuenta = Trim(.iParams(3).Valor)
'                    strCodMoneda = Trim(.iParams(4).Valor)
                        
                    strCodFileAnaliticaContracuenta = strCodFileContracuenta + "-" + strCodAnaliticaContracuenta
                        
                    txtCodFile.Text = strCodFileContracuenta
                    txtCodAnalitica.Text = strCodAnaliticaContracuenta
                              
                    If strTipoFileContracuenta = Valor_File_Generico Then
                        txtCodAnalitica.Enabled = True
                        txtDescripFileAnalitica.Text = "Analítica Genérica"
                    Else
                        txtDescripFileAnalitica.Text = strCodFileAnaliticaContracuenta + " - " + strDescripFileAnaliticaContracuenta
                        txtCodAnalitica.Enabled = True
                    End If
            
                    strDescripFileAnaliticaContracuenta = txtDescripFileAnalitica.Text
                     
            
            
            End Select
        
        End If
            
       
    End With
    
    Set frmBus = Nothing

End Sub

Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    
    CentrarForm Me
  
End Sub
Sub InicializarValores()

    Dim intRegistro As Integer

    Set cmdAccion.FormularioActivo = Me

End Sub
Sub CargarListas()

    '*** Tipo de Persona ***'
    txtDescripCuenta.Text = strDescripContracuenta
    txtDescripFileAnalitica.Text = strDescripFileAnaliticaContracuenta
    txtCodFile.Text = strCodFileContracuenta
    txtCodAnalitica.Text = strCodAnaliticaContracuenta
    strCodFileAnaliticaContracuenta = strCodFileContracuenta + "-" + strCodAnaliticaContracuenta
    
    If strTipoFileContracuenta = Valor_Caracter Then
        cmdBusqueda(1).Enabled = False
    Else
        cmdBusqueda(1).Enabled = True
    End If

End Sub
Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vSave
            Call Grabar
        Case vCancel
            Call Cancelar
        
    End Select
    
End Sub
Public Sub Grabar()
    
    blnOK = True
    
    Unload Me
    
End Sub
Public Sub Cancelar()
    
    blnOK = False

    Unload Me
    
End Sub

Private Sub txtCodAnalitica_LostFocus()

    txtCodAnalitica.Text = Right(String(8, "0") & Trim(txtCodAnalitica.Text), 8)

    strCodAnaliticaContracuenta = txtCodAnalitica.Text
    strCodFileAnaliticaContracuenta = strCodFileContracuenta + "-" + strCodAnaliticaContracuenta

End Sub

Private Sub txtCodFile_LostFocus()

    txtCodFile.Text = Right(String(3, "0") & Trim(txtCodFile.Text), 3)

    strCodFileContracuenta = txtCodFile.Text
    strCodFileAnaliticaContracuenta = strCodFileContracuenta + "-" + strCodAnaliticaContracuenta

End Sub
