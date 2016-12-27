VERSION 5.00
Begin VB.Form frmCron 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Cambio de Fecha"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6510
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6510
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
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
      Left            =   5040
      Picture         =   "frmCron.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   1200
   End
   Begin VB.ComboBox cboFondo 
      Height          =   315
      Left            =   300
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   5925
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
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
      Left            =   3480
      Picture         =   "frmCron.frx":0582
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   1200
   End
   Begin VB.Frame fraControlfecha 
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
      Height          =   2925
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.CheckBox ChkCron 
         Caption         =   "Habilitar actualización automática de cambio de fecha"
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
         Height          =   375
         Left            =   690
         TabIndex        =   5
         Top             =   2280
         Width           =   4965
      End
      Begin VB.Label lblDescrip 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/01/1990"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   465
         Index           =   0
         Left            =   2760
         TabIndex        =   6
         Top             =   1740
         Width           =   2085
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fecha Actual"
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
         Height          =   255
         Index           =   1
         Left            =   1050
         TabIndex        =   2
         Top             =   1860
         Width           =   1515
      End
      Begin VB.Label lblFondo 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ARICSA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   300
         TabIndex        =   1
         Top             =   510
         Width           =   5925
      End
   End
End
Attribute VB_Name = "frmCron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFechaOriginal As String
Dim arrPeriodo() As String, arrFondo() As String
Dim strCodFondo As String

Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset
    
    Dim strSQL      As String, intRegistro As Integer
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    lblFondo.Caption = cboFondo.Text
    
    Set adoRegistro = New ADODB.Recordset
           
    With adoComm

        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            
            strFechaOriginal = adoRegistro("FechaCuota")
            lblDescrip(0).Caption = strFechaOriginal
            
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    Dim adoConsulta As ADODB.Recordset
    
    Set adoConsulta = New ADODB.Recordset
           
    With adoComm
        .CommandText = "SELECT IndCronActivo from FondoValorCuota where CodFondo = '" & strCodFondo & "' AND CodAdministradora = '" & gstrCodAdministradora & "' AND FechaCuota = '" & Convertyyyymmdd(strFechaOriginal) & "'"
        Set adoConsulta = .Execute
            
        If adoConsulta("IndCronActivo") = " " Then
            ChkCron.Value = 0
        Else
            ChkCron.Value = 1
        End If
        
        adoConsulta.Close: Set adoConsulta = Nothing
    End With
    
    frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)

End Sub

Private Sub cmdAceptar_Click()
    
    If MsgBox("¿Desea conservar los cambios?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
        If ChkCron.Value = 1 Then
            adoComm.CommandText = "UPDATE FondoValorCuota SET IndCronActivo = 'X' where CodFondo = '" & strCodFondo & "' AND CodAdministradora = '" & gstrCodAdministradora & "' AND FechaCuota = '" & strFechaOriginal & "'"
            adoConn.Execute adoComm.CommandText
        ElseIf ChkCron.Value = 0 Then
            adoComm.CommandText = "UPDATE FondoValorCuota SET IndCronActivo = '' where CodFondo = '" & strCodFondo & "' AND CodAdministradora = '" & gstrCodAdministradora & "' AND FechaCuota = '" & strFechaOriginal & "'"
            adoConn.Execute adoComm.CommandText
        End If
    End If
    
    frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
    
    Unload Me
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call DarFormato

    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)


End Sub

Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strFechaOriginal = Valor_Caracter
    CentrarForm Me
  
End Sub

Private Sub CargarListas()

    Dim strSQL  As String
    
    '*** Fondos Existentes***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
       
    Dim adoConsulta As ADODB.Recordset
    
    Set adoConsulta = New ADODB.Recordset
           
    With adoComm
        .CommandText = "SELECT IndCronActivo from FondoValorCuota where CodFondo = '" & strCodFondo & "' AND CodAdministradora = '" & gstrCodAdministradora & "' AND FechaCuota = '" & Convertyyyymmdd(strFechaOriginal) & "'"
        Set adoConsulta = .Execute
            
        If adoConsulta("IndCronActivo") = " " Then
            ChkCron.Value = 0
        Else
            ChkCron.Value = 1
        End If
        
        adoConsulta.Close: Set adoConsulta = Nothing
    End With
    
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbCenter)
    Next
            
End Sub
