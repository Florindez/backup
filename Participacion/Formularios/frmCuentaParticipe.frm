VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmCuentaParticipe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas del Contrato"
   ClientHeight    =   6000
   ClientLeft      =   1050
   ClientTop       =   4560
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9150
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   360
      TabIndex        =   13
      Top             =   5160
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1296
      Buttons         =   2
      Caption0        =   "Con&sultar"
      Tag0            =   "3"
      ToolTipText0    =   "Consultar"
      Caption1        =   "&Cerrar"
      Tag1            =   "9"
      ToolTipText1    =   "Cerrar"
      UserControlWidth=   2700
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   5760
      Top             =   5160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoConsulta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame fraCuentaContrato 
      Height          =   4815
      Left            =   200
      TabIndex        =   0
      Top             =   200
      Width           =   8775
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmCuentaParticipe.frx":0000
         Height          =   1335
         Left            =   840
         OleObjectBlob   =   "frmCuentaParticipe.frx":001A
         TabIndex        =   17
         Top             =   3240
         Width           =   7575
      End
      Begin VB.TextBox txtNombreTitular 
         Height          =   300
         Left            =   1830
         MaxLength       =   30
         TabIndex        =   16
         Top             =   2280
         Width           =   6540
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1491
         Width           =   3900
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "<"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   4020
         Width           =   375
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   ">"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtNumCuenta 
         Height          =   300
         Left            =   1830
         MaxLength       =   12
         TabIndex        =   4
         Top             =   1893
         Width           =   3900
      End
      Begin VB.ComboBox cboBanco 
         Height          =   315
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   687
         Width           =   3900
      End
      Begin VB.ComboBox cboTipoCuenta 
         Height          =   315
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1089
         Width           =   3900
      End
      Begin VB.CheckBox chkPredeterminado 
         Alignment       =   1  'Right Justify
         Caption         =   "Prederteminado"
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
         Height          =   240
         Left            =   255
         MaskColor       =   &H80000012&
         TabIndex        =   1
         Top             =   2715
         Width           =   1785
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Moneda"
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
         Height          =   270
         Index           =   5
         Left            =   300
         TabIndex        =   14
         Top             =   1500
         Width           =   1455
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Num. Cuenta"
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
         Height          =   270
         Index           =   3
         Left            =   300
         TabIndex        =   10
         Top             =   1890
         Width           =   1455
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Tipo de Cuenta"
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
         Height          =   270
         Index           =   2
         Left            =   300
         TabIndex        =   9
         Top             =   1100
         Width           =   1455
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Banco"
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
         Height          =   270
         Index           =   1
         Left            =   300
         TabIndex        =   8
         Top             =   700
         Width           =   1455
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Participe"
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
         Height          =   270
         Index           =   0
         Left            =   300
         TabIndex        =   7
         Top             =   320
         Width           =   1455
      End
      Begin VB.Label lblParticipe 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1830
         TabIndex        =   6
         Top             =   300
         Width           =   6540
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Títular Cuenta"
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
         Height          =   270
         Index           =   4
         Left            =   300
         TabIndex        =   5
         Top             =   2300
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmCuentaParticipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrBanco()      As String, arrTipoCuenta()  As String
Dim arrMoneda()     As String
Dim strCodBanco     As String, strCodTipoCuenta As String
Dim strCodMoneda    As String
Dim strEstado       As String
Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vModify
            Call Modificar
        Case vDelete
            Call Eliminar
        Case vSearch
            Call Buscar
        Case vReport
            Call Imprimir
        Case vSave
            Call Grabar
        Case vCancel
            Call Cancelar
        Case vExit
            Call Salir
        
    End Select
    
End Sub
Public Sub Abrir()

End Sub

Public Sub Adicionar()
     
End Sub

Public Sub Anterior()

End Sub

Public Sub Ayuda()

End Sub

Public Sub Buscar()

    Dim strSQL As String
                                                                                    
    Me.MousePointer = vbHourglass
                    
    strSQL = "{ call up_ACSelDatosParametro(16,'" & gstrCodParticipe & "') }"
    
    With adoConsulta
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With
    
    tdgConsulta.Refresh
    
    If adoConsulta.Recordset.RecordCount > 0 Then strEstado = Reg_Consulta
        
    Me.MousePointer = vbDefault
    
End Sub

Public Sub Cancelar()

    Call Salir
    
End Sub

Public Sub Eliminar()
                
End Sub


Public Sub Grabar()

                
End Sub

Public Sub Imprimir()

End Sub

Public Sub Modificar()

    Dim intRegistro As Integer
    
    If strEstado = Reg_Consulta Then
        
        intRegistro = ObtenerItemLista(arrBanco(), Trim(tdgConsulta.Columns(2)))
        If intRegistro >= 0 Then cboBanco.ListIndex = intRegistro
                        
        intRegistro = ObtenerItemLista(arrTipoCuenta(), Trim(tdgConsulta.Columns(4)))
        If intRegistro >= 0 Then cboTipoCuenta.ListIndex = intRegistro
        
        intRegistro = ObtenerItemLista(arrMoneda(), Trim(tdgConsulta.Columns(7)))
        If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
                
        txtNumCuenta.Text = Trim(tdgConsulta.Columns(5))
        
        If Trim(tdgConsulta.Columns(8)) = Valor_Caracter Then
            chkPredeterminado.Value = vbUnchecked
        Else
            chkPredeterminado.Value = vbChecked
        End If
        
        txtNombreTitular.Text = Trim(tdgConsulta.Columns(9))
                
    End If
    
End Sub

Public Sub Primero()

End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub Seguridad()

End Sub

Public Sub Siguiente()

End Sub

Public Sub Ultimo()

End Sub

Private Sub cboBanco_Click()

    strCodBanco = Valor_Caracter
    If cboBanco.ListIndex < 0 Then Exit Sub
    
    strCodBanco = Trim(arrBanco(cboBanco.ListIndex))
        
End Sub

Private Sub cboMoneda_Click()

    strCodMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
End Sub

Private Sub cboTipoCuenta_Click()

    strCodTipoCuenta = Valor_Caracter
    If cboTipoCuenta.ListIndex < 0 Then Exit Sub
    
    strCodTipoCuenta = Trim(arrTipoCuenta(cboTipoCuenta.ListIndex))
        
End Sub

Private Sub cmdAgregar_Click()

    On Error GoTo Error1            '/**/ HMC Habilitamos la rutina de Errores.

    Dim adoRegistro As ADODB.Recordset
    Dim intRegistro As Integer
    
    If TodoOK() Then
        Set adoRegistro = New ADODB.Recordset

        adoComm.CommandText = "{ call up_ACSelDatosParametro(17,'" & gstrCodParticipe & "','" & strCodBanco & "','" & strCodTipoCuenta & "','" & strCodMoneda & "','" & Trim(txtNumCuenta.Text) & "') }"
        Set adoRegistro = adoComm.Execute

        If Not adoRegistro.EOF Then
            MsgBox "Cuenta ya se encuentra registrada.", vbCritical, gstrNombreEmpresa
            Call InicializarValores
            adoRegistro.Close: Set adoRegistro = Nothing
            Exit Sub
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
                
        With adoComm
            .CommandText = "{ call up_PRManCuentaParticipe('"
            .CommandText = .CommandText & gstrCodParticipe & "',"
            .CommandText = .CommandText & "0,'"
            .CommandText = .CommandText & strCodTipoCuenta & "','"
            .CommandText = .CommandText & Trim(txtNumCuenta.Text) & "','"
            .CommandText = .CommandText & strCodBanco & "','"
            .CommandText = .CommandText & strCodMoneda & "','"
            .CommandText = .CommandText & Trim(txtNombreTitular.Text) & "','"
            If chkPredeterminado.Value Then
                .CommandText = .CommandText & "X','"
            Else
                .CommandText = .CommandText & "','"
            End If
            .CommandText = .CommandText & gstrLogin & "','"
            .CommandText = .CommandText & Convertyyyymmdd(gdatFechaActual) & "','"
            .CommandText = .CommandText & gstrLogin & "','"
            .CommandText = .CommandText & Convertyyyymmdd(gdatFechaActual) & "','"
            .CommandText = .CommandText & "I') }"
            
            adoConn.Execute .CommandText
                        
        End With
            
        Call Buscar
    End If
    Exit Sub

Error1:
    MsgBox DescripcionError & vbNewLine & DescripcionTecnica & err.Description, vbExclamation, TituloError ' Mostrar Error

End Sub

Private Function TodoOK() As Boolean

    TodoOK = False
    
    If cboBanco.ListIndex = 0 Then
        MsgBox "Seleccione el Banco.", vbCritical
        cboBanco.SetFocus
        Exit Function
    End If
    
    If cboTipoCuenta.ListIndex = 0 Then
        MsgBox "Seleccione el Tipo de Cuenta.", vbCritical
        cboTipoCuenta.SetFocus
        Exit Function
    End If
    
    If cboMoneda.ListIndex = 0 Then
        MsgBox "Seleccione la Moneda.", vbCritical
        cboMoneda.SetFocus
        Exit Function
    End If
    
    If Trim(txtNumCuenta.Text) = "" Then
        MsgBox "El Campo Número de Cuenta no es Válido!.", vbCritical
        txtNumCuenta.SetFocus
        Exit Function
    End If
    
    If Trim(txtNombreTitular.Text) = "" Then
        MsgBox "El Campo Nombre de Títular no es Válido!.", vbCritical
        txtNombreTitular.SetFocus
        Exit Function
    End If
                                                                
    '*** Si todo paso OK ***
    TodoOK = True

End Function

Private Sub cmdQuitar_Click()
    
    On Error GoTo Error1            '/**/ HMC Habilitamos la rutina de Errores.
    
    If tdgConsulta.Row <> -1 Then   '/**/
        With adoComm
            .CommandText = "DELETE CuentaParticipe "
            .CommandText = .CommandText & "WHERE CodParticipe='" & gstrCodParticipe & _
                            "' AND NumSecuencial=" & CInt(tdgConsulta.Columns(0))
            adoConn.Execute .CommandText
        End With
        Call Buscar
        Exit Sub
    
Error1:
    MsgBox DescripcionError & vbNewLine & DescripcionTecnica & err.Description, vbExclamation, TituloError ' Mostrar Error
    
    End If                          '/**/

End Sub


Private Sub Form_Deactivate()

    ReDim garrTipoDocumento(0)
    Call Salir
    
End Sub

Private Sub Form_Load()
    
    Call InicializarValores
    Call CargarListas
    Call CargarReportes
    Call Buscar
    Call DarFormato
    
    CentrarForm Me
    
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
            
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = Valor_Caracter
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = Valor_Caracter
    
End Sub

Private Sub CargarListas()

    Dim strSQL  As String
    
    '*** Tipo Documento Identidad  - Naturales ***
    strSQL = "{ call up_ACSelDatos(22) }"
    CargarControlLista strSQL, cboBanco, arrBanco(), Sel_Defecto
    
    If cboBanco.ListCount > 0 Then cboBanco.ListIndex = 0
        
    '*** Tipo de Cuenta ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='CTAFON' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoCuenta, arrTipoCuenta(), Sel_Defecto
    
    If cboTipoCuenta.ListCount > 0 Then cboTipoCuenta.ListIndex = 0
    
    '*** Tipo de Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Sel_Defecto
    
    If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = 0
        
End Sub
Private Sub InicializarValores()

    strEstado = Reg_Defecto
        
    cboBanco.ListIndex = -1
    If cboBanco.ListCount > 0 Then cboBanco.ListIndex = 0
    
    cboTipoCuenta.ListIndex = -1
    If cboTipoCuenta.ListCount > 0 Then cboTipoCuenta.ListIndex = 0
    
    cboMoneda.ListIndex = -1
    If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = 0
    
    txtNumCuenta.Text = Valor_Caracter
    txtNombreTitular.Text = Valor_Caracter
    
    chkPredeterminado.Value = vbUnchecked
    
    '*** Verificando Nivel de Acceso de Usuario ***
'    strNivAcceso = AccesoForm(gstrNomOpc, gstrNumInd)

    Set cmdOpcion.FormularioActivo = Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Unload Me
    
End Sub

Private Sub txtNumCuenta_KeyPress(KeyAscii As Integer)

    If KeyAscii >= 48 And KeyAscii <= 57 Then
        KeyAscii = KeyAscii
    ElseIf KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If
    
End Sub
