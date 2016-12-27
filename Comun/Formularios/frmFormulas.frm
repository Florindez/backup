VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFormulas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formulas"
   ClientHeight    =   8340
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   13545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
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
      Height          =   3525
      Left            =   120
      TabIndex        =   16
      Top             =   60
      Width           =   3615
      Begin VB.ListBox lstDatos 
         Columns         =   1
         Height          =   3180
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3405
      End
   End
   Begin VB.ComboBox cboFondoPrueba 
      Height          =   315
      Left            =   7920
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   7440
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Probar"
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7410
      Width           =   1200
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7440
      Width           =   1200
   End
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7440
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7440
      Width           =   1200
   End
   Begin VB.Frame Frame4 
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
      Height          =   3525
      Left            =   7530
      TabIndex        =   5
      Top             =   60
      Width           =   3015
      Begin VB.ListBox lstFunciones 
         Columns         =   1
         Height          =   3180
         Left            =   90
         TabIndex        =   6
         Top             =   240
         Width           =   2850
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   3525
      Left            =   10590
      TabIndex        =   3
      Top             =   60
      Width           =   2805
      Begin VB.ListBox lstOperadores 
         Columns         =   1
         Height          =   3180
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2640
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Constantes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   3840
      TabIndex        =   2
      Top             =   60
      Width           =   3615
      Begin VB.ListBox lstVariables 
         Columns         =   1
         Height          =   3180
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3405
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
      Height          =   3705
      Left            =   120
      TabIndex        =   0
      Top             =   3630
      Width           =   13245
      Begin VB.TextBox txtFormula 
         Height          =   3345
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   210
         Width           =   13035
      End
   End
   Begin MSComCtl2.DTPicker dtpFechaPrueba 
      Height          =   315
      Left            =   7920
      TabIndex        =   15
      Top             =   7800
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      _Version        =   393216
      CalendarForeColor=   8388608
      Format          =   146669569
      CurrentDate     =   2
   End
   Begin VB.Label lblDescrip 
      Caption         =   "Fecha"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   7080
      TabIndex        =   14
      Top             =   7860
      Width           =   615
   End
   Begin VB.Label lblDescrip 
      Caption         =   "Fondo"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   0
      Left            =   7080
      TabIndex        =   13
      Top             =   7500
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   5
      X1              =   6720
      X2              =   6720
      Y1              =   7440
      Y2              =   8280
   End
End
Attribute VB_Name = "frmFormulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aCodVariable()
Dim sCodVariable As String
Dim aCodDato()
Dim sCodDato As String
Dim aCodOperador()
Dim sCodOperador As String
Dim aCodFuncion()
Dim sCodFuncion As String
Dim strSQL As String
Dim nStart As Integer
Dim MyParser As New clsParser
Dim strTipoConstantes As String
Dim strTipoVariables As String
Dim strCodFormulaDatos As String

Dim arrFondoPrueba()    As String, strCodFondoPrueba    As String
Dim strFechaPrueba      As String

Private indAceptar As Boolean

Private Sub cboFondoPrueba_Click()
    strCodFondoPrueba = Valor_Caracter
    If cboFondoPrueba.ListIndex < 0 Then Exit Sub
    
    strCodFondoPrueba = Trim(arrFondoPrueba(cboFondoPrueba.ListIndex))
End Sub

Private Sub cmdAceptar_Click()

On Error GoTo cmdAceptar_ErrHandler

'Dim result As Double
Dim strMsgError As String

    'result = MyParser.ParseExpression(txtFormula.Text, "", New ADODB.Recordset, strMsgError)
    If strMsgError <> "" Then GoTo cmdAceptar_ErrHandler
    indAceptar = True
    Me.Hide
    
    Exit Sub

cmdAceptar_ErrHandler:

    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation
End Sub

Private Sub cmdCancelar_Click()

'Call EliminarVariables
indAceptar = False
Me.Hide

End Sub

Private Sub cmdSeleccionar_Click()

If lstFunciones.ListIndex <> -1 Then
    Call lstFunciones_DblClick
ElseIf lstDatos.ListIndex <> -1 Then
    Call lstDatos_DblClick
ElseIf lstVariables.ListIndex <> -1 Then
    Call lstVariables_DblClick
ElseIf lstOperadores.ListIndex <> -1 Then
    Call lstOperadores_DblClick
Else
    MsgBox "Seleccione un elemento de las listas de Variables, Funciones u Operadores!", 48
End If

End Sub

Private Sub cmdValidar_Click()
Dim strMsgError As String
Dim adoRegistro             As ADODB.Recordset
On Error GoTo cmdValidar_ErrHandler

Dim result As Double

Set adoRegistro = New ADODB.Recordset

adoComm.CommandText = "SELECT CodFondo,CodAdministradora,'" & strFechaPrueba & "' as FechaCierre,'0' as TipoCierre,0 as NumGasto,0 as NumPeriodo " & _
                      "from FondoValorCuota"
Set adoRegistro = adoComm.Execute

    result = MyParser.ParseExpression(txtFormula.Text, "", adoRegistro, strMsgError)
    If strMsgError <> "" Then GoTo cmdValidar_ErrHandler
    'MsgBox "Result:  " & Format(Result, "#0.0#####")
    MsgBox "Formula OK!", 48
    Call EliminarVariables
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
Dim adoRegistro             As ADODB.Recordset
On Error GoTo cmdValidar_ErrHandler

strFechaPrueba = Convertyyyymmdd(dtpFechaPrueba.Value)

Dim result As Double
Set adoRegistro = New ADODB.Recordset

adoComm.CommandText = "SELECT CodFondo,CodAdministradora,'" & strFechaPrueba & "' as FechaCierre,'0' as TipoCierre,0 as NumGasto,0 as NumPeriodo " & _
                      "from FondoValorCuota where CodFondo = '" & strCodFondoPrueba & "' and FechaCuota = '" & strFechaPrueba & "'"
Set adoRegistro = adoComm.Execute

    result = MyParser.ParseExpression(txtFormula.Text, strCodFormulaDatos, adoRegistro, strMsgError)
    If strMsgError <> "" Then GoTo cmdValidar_ErrHandler
    MsgBox "Resultado:  " & Format(result, "#0.0#####")
    Call EliminarVariables

    Exit Sub

cmdValidar_ErrHandler:
    If strMsgError = "" Then strMsgError = err.Description
    If strMsgError Like "El valor de BOF o EOF*" Then strMsgError = "No existen registros para la prueba solicitada. Revise la Fecha."
    MsgBox strMsgError, vbInformation
End Sub




Private Sub Form_Load()

CentrarForm Me

txtFormula.Text = Trim(gstrFormula)

txtFormula.SelStart = Len(txtFormula.Text)

'*** Fondo de Prueba ***
strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
CargarControlLista strSQL, cboFondoPrueba, arrFondoPrueba(), Valor_Caracter
    
If cboFondoPrueba.ListCount > 0 Then cboFondoPrueba.ListIndex = 0

'*** Fecha de Prueba ***
dtpFechaPrueba.Value = gdatFechaActual

End Sub

Private Sub Form_Unload(Cancel As Integer)
Call EliminarVariables
End Sub



Private Sub lstDatos_DblClick()
If nStart = 0 Then
   nStart = 1
End If

txtFormula.Text = Mid$(txtFormula.Text, 1, nStart) + aCodDato(lstDatos.ListIndex) + Mid$(txtFormula.Text, nStart + Len(txtFormula.SelText) + 1)

txtFormula.SelStart = nStart + Len(aCodDato(lstDatos.ListIndex))

txtFormula.SetFocus
End Sub

Private Sub lstDatos_LostFocus()

lstFunciones.ListIndex = -1
lstOperadores.ListIndex = -1
lstVariables.ListIndex = -1

End Sub

Private Sub lstFunciones_DblClick()

Dim nPos As Integer

If lstFunciones.ListCount > 0 Then

    If nStart = 0 Then
        nStart = 1
    End If
    
    txtFormula.Text = Mid$(txtFormula.Text, 1, nStart) + aCodFuncion(lstFunciones.ListIndex) + Mid$(txtFormula.Text, nStart + Len(txtFormula.SelText) + 1)
    
    nPos = InStr(aCodFuncion(lstFunciones.ListIndex), "(")
    
    If nPos = 0 Then
        txtFormula.SelStart = nStart + Len(aCodFuncion(lstFunciones.ListIndex))
    Else
        If nStart = 1 Then
           txtFormula.SelStart = nStart + nPos - 1
        Else
           txtFormula.SelStart = nStart + nPos
        End If
    End If
    
    txtFormula.SetFocus

End If

End Sub

Private Sub lstFunciones_GotFocus()

lstVariables.ListIndex = -1
lstOperadores.ListIndex = -1
lstDatos.ListIndex = -1

End Sub

Private Sub lstOperadores_DblClick()

If nStart = 0 Then
    nStart = 1
End If


txtFormula.Text = Mid$(txtFormula.Text, 1, nStart) + aCodOperador(lstOperadores.ListIndex) + Mid$(txtFormula.Text, nStart + Len(txtFormula.SelText) + 1)

txtFormula.SelStart = nStart + Len(aCodOperador(lstOperadores.ListIndex))

txtFormula.SetFocus

End Sub



Private Sub lstOperadores_GotFocus()

lstVariables.ListIndex = -1
lstFunciones.ListIndex = -1
lstDatos.ListIndex = -1

End Sub

Private Sub lstVariables_DblClick()

If nStart = 0 Then
   nStart = 1
End If

txtFormula.Text = Mid$(txtFormula.Text, 1, nStart) + aCodVariable(lstVariables.ListIndex) + Mid$(txtFormula.Text, nStart + Len(txtFormula.SelText) + 1)

txtFormula.SelStart = nStart + Len(aCodVariable(lstVariables.ListIndex))

txtFormula.SetFocus


End Sub

Private Sub lstVariables_GotFocus()

lstFunciones.ListIndex = -1
lstOperadores.ListIndex = -1
lstDatos.ListIndex = -1

End Sub

Private Sub txtFormula_GotFocus()

nStart = txtFormula.SelStart

End Sub

Private Sub txtFormula_LostFocus()

nStart = txtFormula.SelStart

End Sub

Public Sub AgregarVariables()

For i% = 0 To lstVariables.ListCount - 1
    lstVariables.ListIndex = i%
    MyParser.AddConstant CStr(aCodVariable(lstVariables.ListIndex)), 1  'se asume 1
'    ConstNames.Add CStr(aCodVariable(lstVariables.ListIndex))
Next i%


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

    MyParser.RemoveConstantAll


End Sub

Public Sub mostrarForm(ByRef strResultado As String, ByVal strParTipoConstantes As String, ByVal strDatosCodFormula)
indAceptar = False
txtFormula.Text = strResultado
strTipoConstantes = strParTipoConstantes
strTipoVariables = "02"
strCodFormulaDatos = strDatosCodFormula

''''''Operadores
strSQL = "select cod_operador as codigo, ' ' + txt_desc_sintaxis + '" & vbTab & "' + txt_desc as descrip from toperadores "
strSQL = strSQL + "order by txt_desc"
LCmbLoad strSQL, lstOperadores, aCodOperador(), ""

''''''Funciones
strSQL = "select txt_desc_param as codigo, ' ' + txt_desc_sintaxis as descrip from tfunciones "
strSQL = strSQL + "order by txt_desc"
LCmbLoad strSQL, lstFunciones, aCodFuncion(), ""

''''''Variables
strSQL = "select cod_variable as codigo, ' ' + cod_variable + '" & vbTab & "' + txt_desc as descrip from tcob_var where cod_tipo_var = '" & strTipoVariables & "' "
strSQL = strSQL + "order by txt_desc"
LCmbLoad strSQL, lstDatos, aCodDato(), ""

''''''Constantes
strSQL = "select cod_variable as codigo, ' ' + cod_variable + '" & vbTab & "' + txt_desc as descrip from tcob_var where cod_tipo_var = '" & strTipoConstantes & "' "
strSQL = strSQL + "order by txt_desc"
LCmbLoad strSQL, lstVariables, aCodVariable(), ""

If lstVariables.ListCount > 0 Then
   Call AgregarVariables
   lstVariables.ListIndex = 0
End If

Me.Show 1
If indAceptar Then strResultado = Trim(txtFormula.Text)
Unload Me
End Sub
