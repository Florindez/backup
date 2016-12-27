VERSION 5.00
Object = "{BE4F3AC8-AEC9-101A-947B-00DD010F7B46}#1.0#0"; "MSOUTL32.OCX"
Begin VB.Form frmFormulas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formulas"
   ClientHeight    =   6960
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   13875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
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
      Left            =   10740
      TabIndex        =   11
      Top             =   6120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
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
      Left            =   3000
      Picture         =   "frmFormulas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6120
      Width           =   1200
   End
   Begin VB.CommandButton cmdValidar 
      Caption         =   "Validar"
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
      Left            =   5025
      Picture         =   "frmFormulas.frx":05F4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6120
      Width           =   1200
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "Seleccionar"
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
      Left            =   6990
      Picture         =   "frmFormulas.frx":0AE9
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
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
      Left            =   8940
      Picture         =   "frmFormulas.frx":0BB9
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Width           =   1200
   End
   Begin VB.Frame frFunciones 
      Caption         =   "Funciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   5430
      TabIndex        =   5
      Top             =   300
      Width           =   5535
      Begin VB.ListBox lstFunciones 
         Columns         =   1
         Height          =   2790
         Left            =   90
         TabIndex        =   6
         Top             =   240
         Width           =   5250
      End
   End
   Begin VB.Frame frOperadores 
      Caption         =   "Operadores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   11040
      TabIndex        =   3
      Top             =   300
      Width           =   2685
      Begin VB.ListBox lstOperadores 
         Columns         =   1
         Height          =   2790
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2400
      End
   End
   Begin VB.Frame frVariables 
      Caption         =   "Variables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   120
      TabIndex        =   2
      Top             =   300
      Width           =   5175
      Begin MSOutl.Outline otlVariable 
         Height          =   2760
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   4935
         _Version        =   65536
         _ExtentX        =   8705
         _ExtentY        =   4868
         _StockProps     =   77
         ForeColor       =   8388608
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Style           =   5
         PicturePlus     =   "frmFormulas.frx":111B
         PictureMinus    =   "frmFormulas.frx":1279
         PictureLeaf     =   "frmFormulas.frx":13D7
         PictureOpen     =   "frmFormulas.frx":14D1
         PictureClosed   =   "frmFormulas.frx":162F
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Editar Formula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   3660
      Width           =   13575
      Begin VB.TextBox txtFormula 
         Height          =   2025
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Text            =   "frmFormulas.frx":178D
         Top             =   240
         Width           =   13275
      End
   End
End
Attribute VB_Name = "frmFormulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aCodOperador() As String  '<<NO
Dim sCodOperador As String    '<<NO
Dim aCodFuncion()  '<<NO
Dim sCodFuncion As String '<<NO
Dim strSql As String
Dim nStart As Integer
'Dim MyParser As New clsParser
Dim strTipoConstantes As String

Dim arrVariables() As String, arrOperadores() As String, arrCodVistaFuncion() As String
Dim arrIdVariableFuncion() As String

Dim strCodGrupoVariableSel  As String
Dim strTipoVar As String

Private indAceptar As Boolean

'******VARIABLES DE NIVEL
Dim strNivelGrupo1, strNivelGrupo2, nstrModo, narrVariables() As String


Private Sub cmdAceptar_Click()

On Error GoTo cmdAceptar_ErrHandler

Dim result As Double
Dim strMsgError As String

    'result = MyParser.ParseExpression(txtFormula.Text, "", New ADODB.Recordset, strMsgError)
    If strMsgError <> "" Then GoTo cmdAceptar_ErrHandler
    indAceptar = True
    Me.Hide
    
    Exit Sub

cmdAceptar_ErrHandler:
'    If err.Number >= PERR_FIRST And _
'       err.Number <= PERR_LAST Then
'        ShowParseError
'    Else
'        MsgBox err.Description, vbCritical, "Unexpected Error"
'    End If
    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation
End Sub

Private Sub cmdCancelar_Click()

    If Trim(txtFormula.Text) = Valor_Caracter Then
        frmControlReporte.chkFiltro.Value = Unchecked
    End If
    
    indAceptar = False
    Erase narrVariables
    Erase arrCodVistaFuncion
    Me.Hide

End Sub

Private Sub cmdSeleccionar_Click()

    If otlVariable.ListIndex <> -1 Then
        Call otlVariable_DblClick
    ElseIf lstFunciones.ListIndex <> -1 Then
        Call lstFunciones_DblClick
    ElseIf lstOperadores.ListIndex <> -1 Then
        Call lstOperadores_DblClick
    Else
        MsgBox "Seleccione un elemento de las listas de Variables, Funciones u Operadores!", 48
    End If

End Sub

Private Sub cmdValidar_Click()

    Dim strMsgError As String

    On Error GoTo cmdValidar_ErrHandler

    Dim result As Double

    result = MyParser.ParseExpression(txtFormula.Text, "", New ADODB.Recordset, strMsgError)
    If strMsgError <> "" Then GoTo cmdValidar_ErrHandler
    'MsgBox "Result:  " & Format(Result, "#0.0#####")
        MsgBox "Formula OK!", 48
    Exit Sub

cmdValidar_ErrHandler:
'    If err.Number >= PERR_FIRST And _
'       err.Number <= PERR_LAST Then
'        ShowParseError
'    Else
'        MsgBox err.Description, vbCritical, "Unexpected Error"
'    End If
    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation
End Sub

Private Sub Command1_Click()
    Dim strMsgError As String

    On Error GoTo cmdValidar_ErrHandler

    Dim result As Double

    result = MyParser.ParseExpression(txtFormula.Text, "up_GNDatosFormula", New ADODB.Recordset, strMsgError)
    
    If strMsgError <> "" Then GoTo cmdValidar_ErrHandler
        MsgBox "Resultado:  " & Format(result, "#0.0#####")

    Exit Sub

cmdValidar_ErrHandler:
    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation
End Sub

Private Sub Form_Load()

    CentrarForm Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call EliminarVariables
End Sub

Private Sub lstFunciones_Click()

    If lstFunciones.ListIndex <> -1 Then
        gstrCodVistaUsuario = arrCodVistaFuncion(lstFunciones.ListIndex)
    End If
    
End Sub

Private Sub lstFunciones_DblClick()
    
    If lstFunciones.ListIndex <> -1 Then
        gstrCodVistaUsuario = arrCodVistaFuncion(lstFunciones.ListIndex)
    End If
    
    Dim nPos As Integer
    
    If lstFunciones.ListCount > 0 Then
    
        If nStart = 0 Then
            nStart = 1
        End If
    
        Dim strConcat As String
    
        If nStart = 0 Then
            nStart = 1
        End If
    
    
        strConcat = " "
    
        
        txtFormula.Text = Mid$(txtFormula.Text, 1, nStart) + arrIdVariableFuncion(lstFunciones.ListIndex) + strConcat
        
        txtFormula.SelStart = nStart + Len(arrIdVariableFuncion(lstFunciones.ListIndex)) + Len(strConcat)
        
        txtFormula.SetFocus
    
    End If

End Sub

Private Sub lstFunciones_GotFocus()

    otlVariable.ListIndex = -1
    lstOperadores.ListIndex = -1

End Sub

Private Sub lstOperadores_DblClick()

    Dim strConcat As String
    
    If nStart = 0 Then
        nStart = 1
    End If
    
    
    strConcat = " "
    
    
    txtFormula.Text = Mid$(txtFormula.Text, 1, nStart) + arrOperadores(lstOperadores.ListIndex) + strConcat
    
    txtFormula.SelStart = nStart + Len(arrOperadores(lstOperadores.ListIndex)) + Len(strConcat)
    
    txtFormula.SetFocus

End Sub

Private Sub lstOperadores_GotFocus()

    otlVariable.ListIndex = -1
    lstFunciones.ListIndex = -1

End Sub

Private Sub otlVariable_Click()

    If lstFunciones.ListIndex <> -1 Then
        lstFunciones.ListIndex = -1
    End If
    
    If otlVariable.Expand(otlVariable.ListIndex) = True Then
        otlVariable.Expand(otlVariable.ListIndex) = False
        If otlVariable.HasSubItems(otlVariable.ListIndex) = False Then
            otlVariable.PictureType(otlVariable.ListIndex) = 2
        Else
            otlVariable.PictureType(otlVariable.ListIndex) = 0
        End If
    Else
        otlVariable.Expand(otlVariable.ListIndex) = True
        If otlVariable.HasSubItems(otlVariable.ListIndex) Then
            otlVariable.PictureType(otlVariable.ListIndex) = 1
            strCodGrupoVariableSel = Format(otlVariable.ItemData(otlVariable.ListIndex), "00")
        End If
    End If


End Sub

Private Sub otlVariable_DblClick()

    Dim strConcat As String
    
    If nStart = 0 Then
       nStart = 1
    End If
    
    
    strConcat = " "
    
    
    If (otlVariable.ListIndex <> strNivelGrupo1 And otlVariable.ListIndex <> strNivelGrupo2) And (otlVariable.ListIndex <> 0) Then
    
    txtFormula.Text = Mid$(txtFormula.Text, 1, nStart) + arrVariables(otlVariable.ListIndex - 1) + strConcat
    
    txtFormula.SelStart = nStart + Len(arrVariables(otlVariable.ListIndex - 1)) + Len(strConcat)
    
    txtFormula.SetFocus
    
    End If

End Sub

Private Sub otlVariable_GotFocus()

    lstFunciones.ListIndex = -1
    lstOperadores.ListIndex = -1
    
End Sub

Private Sub txtFormula_GotFocus()

    nStart = Len(txtFormula.Text)
    
End Sub

Private Sub txtFormula_LostFocus()

    nStart = Len(txtFormula.Text)

End Sub

Private Sub ShowParseError()
    
    ' Show error details
    MsgBox "Error No. " & CStr(err.Number - PERR_FIRST + 1) & _
        " - " & err.Description & vbCrLf & _
        "Raised from: " & err.Source, _
        vbCritical, "Parse Error"
    
    ' Mark the position in the expression where the error
    ' was raised
    txtFormula.SelStart = MyParser.LastErrorPosition - 1
    txtFormula.SelLength = 1
    txtFormula.SetFocus

End Sub

Public Sub EliminarVariables()

    'MyParser.RemoveConstantAll

End Sub

Public Sub mostrarForm(ByRef strResultado As String, ByVal strParTipoConstantes As String, ByRef adoRegTemp As ADODB.Recordset, ByVal strModo As String, Optional ByVal strTipo As String)
    Dim strSql As String
    Dim strWhere As String
    
    indAceptar = False
    txtFormula.Text = strResultado
    txtFormula.SelStart = Len(strResultado)
    strTipoConstantes = strParTipoConstantes
    
    
    '''''''Funciones
    strSql = "SELECT Distinct VU.CodVistaUsuario AS CODIGO," & _
            "VU.DescripVista AS DescripVista,VUC.IdVariable AS IdVariable," & _
            "DescripVariable FROM VistaUsuarioCampo VUC " & _
            "JOIN VistaProcesoDetalle VPD ON (VUC.CodVistaUsuario=VPD.CodVistaUsuario) " & _
            "JOIN VistaUsuario VU ON (VUC.CodVistaUsuario=VU.CodVistaUsuario) " & _
            "JOIN VariableUsuario VRU ON (VUC.IdVariable=VRU.IdVariable) " & _
            "WHERE TipoCampo='02' AND VPD.CodVistaProceso='" + gstrCodVistaProceso + "' AND TipoVista='01' " & _
            "ORDER BY DescripVista"
    
    CargarLstFunciones strSql
    
    strTipoVar = strTipo
    
    nstrModo = strModo
    
    Call CargaCamposToArray(adoRegTemp)
    
    strWhere = ObtenerCondicion()
    
    
    'For i = 1 To UBound(narrVariables)
    '
    'narrVariables(i) = "'" + narrVariables(i) + "'"
    '
    'Next
    
    CargarVariables strWhere
    Me.Show 1
    
    If indAceptar Then strResultado = Trim(txtFormula.Text)
    
    Unload Me
End Sub

Private Sub CargarVariables(ByVal strCondicion As String)
    Dim adoRegistro As ADODB.Recordset
    Dim intContador As Integer, intRegistros As Integer
    Dim strWhere, strJoin As String
    Dim strTipoCampo As String
    
    strNivelGrupo1 = ""
    strNivelGrupo2 = ""
    
    Set adoRegistro = New ADODB.Recordset
    adoRegistro.CursorLocation = adUseClient
    adoRegistro.CursorType = adOpenStatic
    
    intContador = 1
    
    If nstrModo = "FORMULA" Then
    
        strTipoCampo = "02"
        
        frOperadores.Enabled = False
        lstOperadores.Enabled = False
        
        If strTipoVar = "99" Or strTipoVar = "98" Then
            
            frFunciones.Enabled = False
            lstFunciones.Enabled = False
            
        Else
        
            frFunciones.Enabled = True
            lstFunciones.Enabled = True
        
        End If
    
    ElseIf nstrModo = "FILTRO" Then
        
        frFunciones.Enabled = False
        lstFunciones.Enabled = False
        
        strTipoCampo = "03"
        
        ReDim arrOperadores(4)
        
        lstOperadores.AddItem "LIKE"
        arrOperadores(0) = "LIKE"
        lstOperadores.AddItem "AND"
        arrOperadores(1) = "AND"
        lstOperadores.AddItem "OR"
        arrOperadores(2) = "OR"
        lstOperadores.AddItem "IN"
        arrOperadores(3) = "IN"
            
    End If
    
    
    With adoComm
        
        .CommandText = "{ call up_CNObtenerListaVariableReporte ('" & _
                            gstrCodVistaProceso & "','" + strCondicion + "','" + strTipoCampo + "' ) }"
            
        adoRegistro.Open .CommandText, adoConn
      
        otlVariable.Clear
        otlVariable.List(0) = "[Datos de Reporte]"
        
        intRegistros = adoRegistro.RecordCount
        ReDim arrVariables(intRegistros)
        adoRegistro.MoveFirst

        Do While Not adoRegistro.EOF
            arrVariables(intContador - 1) = adoRegistro("IdVariable")
            
            If adoRegistro("IdVariable") = "VISTA_DATOS" Or adoRegistro("IdVariable") = "CAMPOS_FILTRO" Then
            strNivelGrupo1 = intContador
            End If
            
            If adoRegistro("IdVariable") = "VARIABLE_USUARIO" Then
            strNivelGrupo2 = intContador
            End If

            otlVariable.AddItem adoRegistro("IdVariable") + " - " + adoRegistro("DescripVariable")
            otlVariable.indent(intContador) = adoRegistro("NivelVariable")

            otlVariable.ItemData(intContador) = adoRegistro("CodGrupoVariable")

            intContador = intContador + 1
            adoRegistro.MoveNext
        Loop
        
        
        adoRegistro.Close: Set adoRegistro = Nothing
        
        'txtFormula.SelStart = 0
            
    For intContador = 1 To intRegistros
        If otlVariable.HasSubItems(intContador) = False Then
            otlVariable.PictureType(intContador) = 2
        End If
    Next
        
    End With
    
End Sub

Private Sub CargaCamposToArray(ByRef adoTemp As ADODB.Recordset)

    Dim intCont, intNum As Integer
    Dim dblBook As Double
    
    intCont = 0
    
    ReDim narrVariables(intCont)

       If Not adoTemp.BOF Then
                'dblBook = adoTemp.Bookmark
                adoTemp.MoveFirst

                Do Until adoTemp.EOF
                ReDim Preserve narrVariables(intCont)
                 
                 narrVariables(intCont) = adoTemp("CodVariableReporte")
                 
                 adoTemp.MoveNext
                 intCont = intCont + 1
                Loop

                'If Not adoTemp.BOF Then adoTemp.Bookmark = dblBook
            
         Else
                
                narrVariables(0) = ""
                
         End If
       
End Sub

Private Function ObtenerCondicion() As String

    Dim i As Integer
    
    For i = 0 To UBound(narrVariables)
    
        narrVariables(i) = "''" + narrVariables(i) + "''"
    
    Next
    
    ObtenerCondicion = Join(narrVariables, ",")


End Function

Private Sub CargarLstFunciones(ByVal strSql As String)

    Dim adoRegistro As ADODB.Recordset
    Dim intCont As Long
    
    Set adoRegistro = New ADODB.Recordset
    
    adoComm.CommandText = strSql
    Set adoRegistro = adoComm.Execute

    lstFunciones.Clear
    intCont = 0
    ReDim arrCodVistaFuncion(intCont)
    ReDim arrIdVariableFuncion(intCont)
          
    Do Until adoRegistro.EOF
        lstFunciones.AddItem adoRegistro("IdVariable") + " - " + adoRegistro("DescripVariable")
        ReDim Preserve arrCodVistaFuncion(intCont)
        ReDim Preserve arrIdVariableFuncion(intCont)
        
        arrCodVistaFuncion(intCont) = adoRegistro("CODIGO")
        arrIdVariableFuncion(intCont) = adoRegistro("IdVariable")
        adoRegistro.MoveNext
        intCont = intCont + 1
    Loop
   
    adoRegistro.Close: Set adoRegistro = Nothing

End Sub


