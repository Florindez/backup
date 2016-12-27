VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{442CAE95-1D41-47B1-BE83-6995DA3CE254}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmVariableUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Variables Reporte"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   10380
   Begin TabDlg.SSTab tabVariable 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   13361
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Variables"
      TabPicture(0)   =   "frmVariableUsuario.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdOpcion"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSalir"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Datos"
      TabPicture(1)   =   "frmVariableUsuario.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frVariable"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -68640
         TabIndex        =   10
         Top             =   6600
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   1296
         Buttons         =   2
         Caption0        =   "&Guardar"
         Tag0            =   "2"
         ToolTipText0    =   "Guardar"
         Caption1        =   "&Cancelar"
         Tag1            =   "8"
         ToolTipText1    =   "Cancelar"
         UserControlWidth=   2700
      End
      Begin TAMControls2.ucBotonEdicion2 cmdSalir 
         Height          =   735
         Left            =   8280
         TabIndex        =   2
         Top             =   6480
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1296
         Caption0        =   "&Salir"
         Tag0            =   "9"
         ToolTipText0    =   "Salir"
         UserControlWidth=   1200
      End
      Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
         Height          =   735
         Left            =   600
         TabIndex        =   3
         Top             =   6480
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   1296
         Buttons         =   3
         Caption0        =   "&Nuevo"
         Tag0            =   "0"
         ToolTipText0    =   "Nuevo"
         Caption1        =   "&Modificar"
         Tag1            =   "3"
         ToolTipText1    =   "Modificar"
         Caption2        =   "&Eliminar"
         Tag2            =   "4"
         ToolTipText2    =   "Eliminar"
         UserControlWidth=   4200
      End
      Begin VB.Frame frVariable 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5775
         Left            =   -74640
         TabIndex        =   4
         Top             =   540
         Width           =   9255
         Begin VB.TextBox txtTipoDato 
            Height          =   315
            Left            =   2520
            TabIndex        =   15
            Top             =   3090
            Width           =   4980
         End
         Begin VB.ComboBox cboTipoVariable 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1920
            Width           =   2190
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   315
            Left            =   2520
            TabIndex        =   6
            Top             =   1320
            Width           =   4980
         End
         Begin VB.TextBox txtIdVariable 
            Height          =   315
            Left            =   2520
            TabIndex        =   5
            Top             =   2520
            Width           =   4980
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Dato"
            Height          =   195
            Index           =   4
            Left            =   600
            TabIndex        =   14
            Top             =   3120
            Width           =   705
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo"
            Height          =   195
            Index           =   2
            Left            =   600
            TabIndex        =   13
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   11
            Top             =   1920
            Width           =   315
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion"
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   9
            Top             =   1320
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Id"
            Height          =   195
            Index           =   3
            Left            =   600
            TabIndex        =   8
            Top             =   2520
            Width           =   135
         End
         Begin VB.Label lblCodVariable 
            Height          =   255
            Left            =   2520
            TabIndex        =   7
            Top             =   720
            Width           =   1215
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmVariableUsuario.frx":0038
         Height          =   5415
         Left            =   120
         OleObjectBlob   =   "frmVariableUsuario.frx":0052
         TabIndex        =   1
         Top             =   720
         Width           =   9855
      End
   End
End
Attribute VB_Name = "frmVariableUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strEstado As String
Dim adoConsulta As ADODB.Recordset
Dim arrTipoVariable() As String
Dim strCodTipoVariable As String
Dim strTipoDato() As String
'**********************************BMM  19/02/2012*****************************

Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call CargarReportes
    Call Buscar
    Call DarFormato
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
    
End Sub

Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabVariable.Tab = 0
    tabVariable.TabEnabled(1) = False
    
    lblCodVariable.FontBold = True
    
    lblDescrip(4).Visible = False
    txtTipoDato.Visible = False

    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me

End Sub

Private Sub CargarListas()

    Dim strSQL  As String
    

    '*** Vistas de Usuario ***
    strSQL = "SELECT CodParametro AS CODIGO,DescripParametro AS DESCRIP " & _
                "From AuxiliarParametro WHERE CodTipoParametro='REPVAR' " & _
                "ORDER BY DescripParametro"
                
    CargarControlLista strSQL, cboTipoVariable, arrTipoVariable(), Sel_Defecto
    
    If cboTipoVariable.ListCount > 0 Then
        cboTipoVariable.ListIndex = 0
    End If

End Sub



Private Sub CargarReportes()

'   frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Diario General"
    
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Diario General (ME)"
    
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Mayor General"
    
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Text = "Mayor General (ME)"
    
    
End Sub

Public Sub Buscar()

    Dim strSQL As String
    
    Set adoConsulta = New ADODB.Recordset
           
    strSQL = "SELECT CodVariable,DescripVariable,IdVariable,DescripParametro,TipoVariable,TipoDato " & _
                "FROM VariableUsuario VU, AuxiliarParametro AP " & _
                "Where VU.TipoVariable = AP.CodParametro AND CodTipoParametro='REPVAR' " & _
                "AND IndVigente='X' ORDER BY DescripVariable"
                        
    strEstado = Reg_Defecto
    
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    tdgConsulta.DataSource = adoConsulta
    
    If adoConsulta.RecordCount > 0 Then strEstado = Reg_Consulta
    
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    Dim elemento As Object
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
    For Each elemento In Me.Controls
    
        If TypeOf elemento Is TDBGrid Then
            Call FormatoGrilla(elemento)
        End If
    
    Next
            
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vModify
            Call Modificar
        Case vDelete
            Call Eliminar
        Case vSave
            Call Grabar
        Case vCancel
            Call Cancelar
        Case vExit
            Call Salir
        
    End Select
    
End Sub

Public Sub Adicionar()
    If strEstado = Reg_Consulta Then
        strEstado = Reg_Adicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabVariable
            .TabEnabled(0) = False
            .Tab = 1
            .TabEnabled(1) = True
        End With
    End If
    
End Sub

Public Sub Modificar()
    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabVariable
            .TabEnabled(0) = False
            .Tab = 1
            .TabEnabled(1) = True
        End With
    End If
    
End Sub

Public Sub Grabar()
       
    
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    If strEstado = Reg_Adicion Then
    
    If MsgBox(Mensaje_Adicion, vbQuestion + vbYesNo + vbDefaultButton2, gstrNombreEmpresa) = vbNo Then Exit Sub
    
        If TodoOK() Then
                
                '*** Verificar Id Variable no se repita ***
                
                If VerificaIdDuplicada(Trim(txtIdVariable.Text)) Then
                
                    Me.MousePointer = vbHourglass
                    
                    With adoComm
                        
                        '*** Guardar ***
                        .CommandText = "{ call up_ACManVariableUsuario('" & _
                        lblCodVariable.Caption & "','" & Trim(txtIdVariable.Text) & "','" & txtDescripcion.Text & _
                        "','" & strCodTipoVariable & "','" & Trim(txtTipoDato.Text) & "','I') }"
                        
                        adoConn.Execute .CommandText
                        
                    End With
        
                    Me.MousePointer = vbDefault
                                
                    MsgBox Mensaje_Adicion_Exitosa, vbExclamation
                    
                    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
                    
                    cmdOpcion.Visible = True
                    With tabVariable
                        .TabEnabled(0) = True
                        .Tab = 0
                        .TabEnabled(1) = False
                    End With
                    
                    Call Limpiar
                    Call Buscar
                
                End If
                
        End If
    End If
    
    If strEstado = Reg_Edicion Then
    
    If MsgBox(Mensaje_Edicion, vbQuestion + vbYesNo + vbDefaultButton2, gstrNombreEmpresa) = vbNo Then Exit Sub
    
        If TodoOK() Then
                Me.MousePointer = vbHourglass
                
                '*** Guardar ***
                With adoComm
                    .CommandText = "{ call up_ACManVariableUsuario('" & _
                    lblCodVariable.Caption & "','" & Trim(txtIdVariable.Text) & "','" & txtDescripcion.Text & _
                    "','" & strCodTipoVariable & "','" & Trim(txtTipoDato.Text) & "','U') }"
                    
                    adoConn.Execute .CommandText
                    
                End With
    
                Me.MousePointer = vbDefault
                            
                MsgBox Mensaje_Edicion_Exitosa, vbExclamation
                
                frmMainMdi.stbMdi.Panels(3).Text = "Acción"
                
                cmdOpcion.Visible = True
                With tabVariable
                    .TabEnabled(0) = True
                    .Tab = 0
                    .TabEnabled(1) = False
                End With
                
                Call Limpiar
                Call Buscar
            End If
    End If

End Sub

Public Sub Eliminar()

    If MsgBox(Mensaje_Eliminacion, vbQuestion + vbYesNo + vbDefaultButton2, gstrNombreEmpresa) = vbNo Then Exit Sub
    
    If strEstado = Reg_Consulta Then
    
            Me.MousePointer = vbHourglass
                
                '*** Guardar ***
            With adoComm
                .CommandText = "{ call up_ACManVariableUsuario('" & _
                    tdgConsulta.Columns(0) & "','','','','','D') }"
                
                adoConn.Execute .CommandText
                
            End With
    
            Me.MousePointer = vbDefault
                        
            MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
                        
            Call Buscar
    
    End If

End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    
    With tabVariable
        .TabEnabled(0) = True
        .Tab = 0
        .TabEnabled(1) = False
    End With
    
    Call Limpiar
    Call Buscar
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub


Public Sub LlenarFormulario(ByVal strModo As String)

    Dim strCodVariable As String, strSQL As String
    
    Dim adoRegistro As ADODB.Recordset
    
    Dim intCont As Integer
    
    Select Case strModo
    
    Case Reg_Adicion
    
        frVariable.Caption = "Nueva Variable"
        frVariable.ForeColor = &H800000
        frVariable.FontBold = True
        frVariable.Font = "Arial"
        
        lblCodVariable.FontBold = True
          
        lblCodVariable.Caption = NuevoCodigo()
        txtDescripcion.SetFocus
        
        txtIdVariable.Locked = False
        txtIdVariable.BackColor = &H80000005
        
   
    Case Reg_Edicion
            
        strCodVariable = tdgConsulta.Columns(0)
        
        lblCodVariable.Caption = strCodVariable
        
        txtDescripcion.Text = tdgConsulta.Columns(1)
        
        frVariable.Caption = "Variable: " + tdgConsulta.Columns(2)
        frVariable.ForeColor = &H800000
        frVariable.FontBold = True
        frVariable.Font = "Arial"
        
        intRegistro = ObtenerItemLista(arrTipoVariable(), tdgConsulta.Columns(4))
        If intRegistro >= 0 Then cboTipoVariable.ListIndex = intRegistro
        
        If tdgConsulta.Columns(4) = "02" Then
        
            lblDescrip(4).Visible = True
            txtTipoDato.Visible = True
            txtTipoDato.Text = tdgConsulta.Columns(5)
            
        Else
        
            lblDescrip(4).Visible = False
            txtTipoDato.Visible = False
        
        End If
        
        cboTipoVariable.Enabled = False
              
        txtIdVariable.Text = tdgConsulta.Columns(2)
        txtIdVariable.Locked = True
        txtIdVariable.BackColor = &H8000000F
    
    End Select
    
End Sub

Private Sub tdgConsulta_DblClick()
    Call Modificar
End Sub

Private Sub tdgConsulta_HeadClick(ByVal ColIndex As Integer)
    
    Dim strColNameTDB  As String
    Static numColindex As Integer
    Static strPrevColumTDB As String
    '** agregar para que no se raye la seleccion de registro con ordenamiento
    strColNameTDB = tdgConsulta.Columns(ColIndex).DataField
    
    If strColNameTDB = strPrevColumTDB Then
        If indSortAsc Then
            indSortAsc = False
            indSortDesc = True
        Else
            indSortAsc = True
            indSortDesc = False
        End If
    Else
        indSortAsc = True
        indSortDesc = False
    End If
    '***

    tdgConsulta.Splits(0).Columns(numColindex).HeadingStyle.ForegroundPicture = Null

    Call OrdenarDBGrid(ColIndex, adoConsulta, tdgConsulta)
    
    numColindex = ColIndex
    
    '****
    strPrevColumTDB = strColNameTDB
    '***
    
End Sub

Private Sub cboTipoVariable_Click()

    strCodTipoVariable = arrTipoVariable(cboTipoVariable.ListIndex)
    
    If arrTipoVariable(cboTipoVariable.ListIndex) = "02" Then
        lblDescrip(4).Visible = True
        txtTipoDato.Visible = True
    Else
        lblDescrip(4).Visible = False
        txtTipoDato.Visible = False
    End If

End Sub

Private Function TodoOK() As Boolean
        
    TodoOK = False
            
    If Trim(txtDescripcion.Text) = Valor_Caracter Then
        MsgBox "La Descripcion de la Variable no puede estar en blanco", vbCritical
        txtDescripcion.SetFocus
        Exit Function
    End If
    
    If Trim(txtIdVariable.Text) = Valor_Caracter Then
        MsgBox "La Id la Variable no puede estar en blanco", vbCritical
        txtIdVariable.SetFocus
        Exit Function
    End If
    
    
    If cboTipoVariable.ListIndex <= 0 Then
    
        MsgBox "No ah seleccionado el tipo de Variable", vbCritical
        cboTipoVariable.SetFocus
        Exit Function
    
    End If
    
    If arrTipoVariable(cboTipoVariable.ListIndex) = "02" Then
    
        If Trim(txtTipoDato.Text) = "" Then
        
            MsgBox "No ah especificado un Tipo de dato", vbCritical
            txtTipoDato.SetFocus
            Exit Function
        End If
        '
        If ValidaTipoVariable = False Then
        
            MsgBox "El tipo de dato especificado no es un tipo de dato valido", vbCritical
            txtTipoDato.SetFocus
            Exit Function
        End If
     
     
    End If
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function

Private Sub Limpiar()

lblCodVariable.Caption = Valor_Caracter
txtDescripcion.Text = Valor_Caracter
txtIdVariable.Text = Valor_Caracter
cboTipoVariable.ListIndex = 0
txtDescripcion.SetFocus
cboTipoVariable.Enabled = True
txtIdVariable.BackColor = &H80000005

lblDescrip(4).Visible = False
txtTipoDato.Visible = False
txtTipoDato.Text = Valor_Caracter

End Sub

Public Function NuevoCodigo() As String

    Dim adoRegistro As ADODB.Recordset
    Dim strSQL As String
    Dim strCodigo As String, intCodigo As Integer
    
    NuevoCodigo = Valor_Caracter
    
    Set adoRegistro = New ADODB.Recordset

    With adoComm
    
        strSQL = "SELECT MAX(CodVariable) AS CodVariable FROM VariableUsuario"
        
        .CommandText = strSQL
        
        Set adoRegistro = .Execute
        
        Do Until adoRegistro.EOF
        
            strCodigo = adoRegistro("CodVariable")
            adoRegistro.MoveNext
        
        Loop
        
    End With
        
        intCodigo = CInt(strCodigo) + 1
        
        strCodigo = CStr(intCodigo)
        
        Select Case Len(strCodigo)
        
            Case 1
            
            strCodigo = "00" + strCodigo
            
            Case 2
            
            strCodigo = "0" + strCodigo
    
        End Select
    
    

    adoRegistro.Close: Set adoRegistro = Nothing
    
     
    NuevoCodigo = strCodigo

End Function

Public Function ValidaTipoVariable() As Boolean
    
    On Error GoTo CtrlErr
    
    ValidaTipoVariable = False
    
    Dim strTip As String
    strTip = Trim(txtTipoDato.Text)
    
    With adoComm
    
        .CommandText = "DECLARE @Valida " + strTip
    
    
    adoConn.Execute .CommandText
    
    
    End With
    
    
    
    ValidaTipoVariable = True
    Exit Function
    
CtrlErr:
    
    Exit Function
    

End Function

Public Function VerificaIdDuplicada(ByVal idVar As String) As Boolean

    VerificaIdDuplicada = False
    
    Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset

    With adoComm
    
        .CommandText = "SELECT IdVariable FROM VariableUsuario"
        
        Set adoRegistro = .Execute
    
    End With
    
    Do Until adoRegistro.EOF
        
        If adoRegistro("IdVariable") = idVar Then
        
            MsgBox "Ya existe una variable con este ID", vbCritical
            txtIdVariable.SetFocus
            Exit Function
        
        End If
        
        adoRegistro.MoveNext
    
    Loop
    
    VerificaIdDuplicada = True

End Function

'********************************************************************

