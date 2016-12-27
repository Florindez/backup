VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmRubros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rubros Reporte"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7035
   Begin TabDlg.SSTab tabRubro 
      Height          =   5715
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   10081
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
      TabCaption(0)   =   "Rubros"
      TabPicture(0)   =   "frmRubros.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos"
      TabPicture(1)   =   "frmRubros.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frRubros"
      Tab(1).Control(1)=   "cmdAccion"
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -71400
         TabIndex        =   6
         Top             =   4320
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
      Begin VB.Frame frRubros 
         Height          =   2325
         Left            =   -74940
         TabIndex        =   2
         Top             =   1140
         Width           =   6645
         Begin VB.TextBox txtDescripcion 
            Height          =   315
            Left            =   1560
            TabIndex        =   3
            Top             =   1110
            Width           =   4800
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   7
            Top             =   540
            Width           =   495
         End
         Begin VB.Label lblCodRubro 
            Height          =   255
            Left            =   1530
            TabIndex        =   5
            Top             =   540
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   4
            Top             =   1080
            Width           =   840
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmRubros.frx":0038
         Height          =   5145
         Left            =   150
         OleObjectBlob   =   "frmRubros.frx":0052
         TabIndex        =   1
         Top             =   450
         Width           =   6405
      End
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   990
      TabIndex        =   8
      Top             =   5940
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1296
      Buttons         =   2
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      ToolTipText1    =   "Modificar"
      UserControlWidth=   2700
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   4890
      TabIndex        =   9
      Top             =   5940
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
End
Attribute VB_Name = "frmRubros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strEstado As String
Dim adoConsulta As ADODB.Recordset
Dim indSortAsc As Boolean, indSortDesc As Boolean

'**********************************BMM  19/02/2012*****************************

Private Sub Form_Load()

    Call InicializarValores
'    Call CargarListas
'    Call CargarReportes
    Call Buscar
    Call DarFormato
    
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
    
End Sub

Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabRubro.Tab = 0
    tabRubro.TabEnabled(1) = False
    
    lblCodRubro.FontBold = True
            
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me

End Sub

Public Sub Buscar()

    Dim strSQL As String
    
    Set adoConsulta = New ADODB.Recordset
           
    strSQL = "SELECT CodRubroEstructura,DescripRubroEstructura " & _
                " FROM RubroEstructura WHERE IndVigente='X' "
                        
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
'        Case vDelete
'            Call Eliminar
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
        With tabRubro
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
        With tabRubro
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
                Me.MousePointer = vbHourglass
                
                '*** Guardar ***
                With adoComm
                    .CommandText = "{ call up_ACManRubroEstructura('" & _
                    lblCodRubro.Caption & "','" & txtDescripcion.Text & _
                    "','I') }"
                    
                    adoConn.Execute .CommandText
                    
                End With
    
                Me.MousePointer = vbDefault
                            
                MsgBox Mensaje_Adicion_Exitosa, vbExclamation
                
                frmMainMdi.stbMdi.Panels(3).Text = "Acción"
                
                cmdOpcion.Visible = True
                With tabRubro
                    .TabEnabled(0) = True
                    .Tab = 0
                    .TabEnabled(1) = False
                End With
                
                Call Limpiar
                Call Buscar
            End If
    End If
    
    If strEstado = Reg_Edicion Then
    
    If MsgBox(Mensaje_Edicion, vbQuestion + vbYesNo + vbDefaultButton2, gstrNombreEmpresa) = vbNo Then Exit Sub
    
        If TodoOK() Then
                Me.MousePointer = vbHourglass
                
                '*** Guardar ***
                With adoComm
                    .CommandText = "{ call up_ACManRubroEstructura('" & _
                    lblCodRubro.Caption & "','" & txtDescripcion.Text & _
                    "','U') }"
                    
                    adoConn.Execute .CommandText
                    
                End With
    
                Me.MousePointer = vbDefault
                            
                MsgBox Mensaje_Edicion_Exitosa, vbExclamation
                
                frmMainMdi.stbMdi.Panels(3).Text = "Acción"
                
                cmdOpcion.Visible = True
                With tabRubro
                    .TabEnabled(0) = True
                    .Tab = 0
                    .TabEnabled(1) = False
                End With
                
                Call Limpiar
                Call Buscar
            End If
    End If

End Sub

'Public Sub Eliminar()
'
'    If MsgBox(Mensaje_Eliminacion, vbQuestion + vbYesNo + vbDefaultButton2, gstrNombreEmpresa) = vbNo Then Exit Sub
'
'    If strEstado = Reg_Consulta Then
'
'            Me.MousePointer = vbHourglass
'
'                '*** Guardar ***
'            With adoComm
'                .CommandText = "{ call up_ACManVariableUsuario('" & _
'                    tdgConsulta.Columns(0) & "','','','','D') }"
'
'                adoConn.Execute .CommandText
'
'            End With
'
'            Me.MousePointer = vbDefault
'
'            MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation
'
'            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
'
'            Call Buscar
'
'    End If
'
'End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    
    With tabRubro
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

    Dim strCodRubro As String

    Dim strSQL As String
    
    Dim intCont As Integer
    
    Select Case strModo
    
    Case Reg_Adicion
    
        frRubros.Caption = "Nuevo Rubro"
        frRubros.ForeColor = &H800000
        frRubros.FontBold = True
        frRubros.Font = "Arial"
        
        lblCodRubro.FontBold = True
          
        lblCodRubro.Caption = NuevoCodigo()
        txtDescripcion.SetFocus
    
    Case Reg_Edicion
            
        strCodRubro = tdgConsulta.Columns(0)
        
        lblCodRubro.Caption = strCodRubro
       
        txtDescripcion.Text = tdgConsulta.Columns(1)
        
        frRubros.Caption = "Rubro: " + tdgConsulta.Columns(1)
        frRubros.ForeColor = &H800000
        frRubros.FontBold = True
        frRubros.Font = "Arial"
                  
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

Private Function TodoOK() As Boolean
        
    TodoOK = False
            
    If Trim(txtDescripcion.Text) = Valor_Caracter Then
        MsgBox "La Descripcion del Rubro no puede estar en blanco", vbCritical
        txtDescripcion.SetFocus
        Exit Function
    End If
     
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function

Private Sub Limpiar()

lblCodRubro.Caption = Valor_Caracter
txtDescripcion.Text = Valor_Caracter

End Sub

Public Function NuevoCodigo() As String

    Dim adoRegistro As ADODB.Recordset
    Dim strSQL As String
    Dim strCodigo As String, intCodigo As Integer
    
    NuevoCodigo = Valor_Caracter
    
    Set adoRegistro = New ADODB.Recordset

    With adoComm
    
        strSQL = "SELECT MAX(CodRubroEstructura) AS CodRubroEstructura FROM RubroEstructura" & _
                " WHERE CodRubroEstructura!='999'"
        
        .CommandText = strSQL
        
        Set adoRegistro = .Execute
        
        Do Until adoRegistro.EOF
        
            strCodigo = adoRegistro("CodRubroEstructura")
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



'********************************************************************


