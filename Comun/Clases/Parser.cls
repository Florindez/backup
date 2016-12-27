VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' clsParser - Mathematical y Logical Expression Parser
' By Andres Caceres
'

Private Const PI = 3.14159265358979

' A generic error text to raise when there's no specific text
Private Const GENERIC_SYNTAX_ERR_MSG = "Syntax Error."

' Parser Error codes
' The values PERR_FIRST and PERR_LAST allow the client app
' to test whether the error is a parser error or VB error
' See the demo form
Public Enum ParserErrors
    PERR_FIRST = vbObjectError + 513
    PERR_SYNTAX_ERROR = PERR_FIRST
    PERR_DIVISION_BY_ZERO
    PERR_CLOSING_PARENTHESES_EXPECTED
    PERR_INVALID_CONST_NAME
    PERR_FUNCTION_DOES_NOT_EXIST
    PERR_RESERVED_WORD
    PERR_CONST_ALREADY_EXISTS
    PERR_CONST_DOES_NOT_EXIST
    PERR_LAST = PERR_CONST_DOES_NOT_EXIST
End Enum

' Tokens (Operators) supported by the parser.
Private Enum ParserTokens
    TOK_UNKNOWN
    TOK_FIRST
    TOK_ADD = TOK_FIRST
    TOK_SUBTRACT
    TOK_MULTIPLY
    TOK_DIVIDE
    TOK_POWER
    TOK_FACTORIAL
    TOK_PARAM
    TOK_NOT
    TOK_OR
    TOK_AND
    TOK_EQUAL
    TOK_NOT_EQUAL
    TOK_GREATER_THAN
    TOK_GREATER_OR_EQUAL_THAN
    TOK_LOWER_THAN
    TOK_LOWER_OR_EQUAL_THAN
    TOK_OPEN_PARENTHESES
    TOK_CLOSE_PARENTHESES
    TOK_LAST = TOK_CLOSE_PARENTHESES
End Enum

' This array holds the symbols used to represent operators.
' You may change them. For example, if you add a "not equal"
' operator, you may use '!=' or '<>' symbols for it
Private mTokenSymbols() As String

Private mExpression As String
' Current position where the parser is in the expression
Private mPosition As Long
Private mLastTokenLength As Long

' Holds user-defined and built-in constants
Private mConstants As Collection 'clsCollection

' Holds the VB Project name - used by error handling code
Private mProjectName As String


'Variables obtener constantes de la BD
Dim adoConsultaFormula             As ADODB.Recordset

' This function is the top-level parsing function, exposed
' to the client. Its sole logic is to check that there's no
' garbage at the end of the expression, since ParseNumExp
' and all the lower level function return when they
' run into something they don't identify - That's what runs
' the whole magic
Public Function ParseExpression(Expression As String, ByVal strCodFormulaDatos As String, ByVal rsDatos As ADODB.Recordset, ByRef strMsgError As String) As Double
On Error GoTo ParseExpression_ErrHandler

Dim Value As Double
Dim strParametros As String
Dim strSpDatosFormula As String

    mExpression = UCase(Expression)
    mPosition = 1
    
    'Call RemoveConstantAll
    
    If Trim(strCodFormulaDatos) <> "" Then
        strSpDatosFormula = traerCampo("FormulaDatos", "DescripFormulaDatos", "CodFormulaDatos", strCodFormulaDatos)
        strParametros = obtenerParametrosDatos(strCodFormulaDatos, rsDatos)
        obtenerDatosSistema strSpDatosFormula, strParametros, strMsgError
        If strMsgError <> "" Then GoTo ParseExpression_ErrHandler
    End If
        
    SkipSpaces
    Value = ParseLogExp
    SkipSpaces
    
    ' If ParseNumExp didn't parse the whole expression,
    ' it means there's some garbage at the end
    If mPosition <= Len(mExpression) Then
        err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG + " Character '" + Mid(mExpression, mPosition, 1) + "' Invalid."
    End If

    ParseExpression = Value
    
    If TypeName(adoConsultaFormula) = "Recordset" Then
        If adoConsultaFormula.State = 1 Then adoConsultaFormula.Close: Set adoConsultaFormula = Nothing
    End If
    Exit Function
    
ParseExpression_ErrHandler:
    ' The following call sets err.Source to the function
    ' name. If the error was raised by ParseNumExp, this
    ' function's name will be added to the existing
    ' err.Source, so the client can see exactly how the
    ' call stack looked like when the error occured
'    SetErrSource "ParseExpression"
'    err.Raise err.Number
    If TypeName(adoConsultaFormula) = "Recordset" Then
        If adoConsultaFormula.State = 1 Then adoConsultaFormula.Close: Set adoConsultaFormula = Nothing
    End If
    If strMsgError = "" Then strMsgError = err.Description
 End Function
 
 Private Function obtenerParametrosDatos(ByVal strCodFormulaDatos As String, ByVal rsDatos As ADODB.Recordset) As String
    Dim adoConsultaParametros             As New ADODB.Recordset
    Dim strParametros As String
 
    adoComm.CommandText = "SELECT FP.DescripFormulaParametros " & _
                          "FROM dbo.FormulaDatosParametros FDP " & _
                                "INNER JOIN dbo.FormulaParametros FP ON FDP.CodFormulaParametros = FP.CodFormulaParametros " & _
                                "WHERE CodFormulaDatos = '" & strCodFormulaDatos & "' ORDER BY FDP.intOrdenParametros ASC"
    Set adoConsultaParametros = adoComm.Execute
    
    strParametros = ""
    Do While Not adoConsultaParametros.EOF
        strParametros = strParametros & rsDatos(Right(adoConsultaParametros("DescripFormulaParametros"), Len(adoConsultaParametros("DescripFormulaParametros")) - 1)) & "|"
        adoConsultaParametros.MoveNext
    Loop
    adoConsultaParametros.Close: Set adoConsultaParametros = Nothing
    
    obtenerParametrosDatos = Left(strParametros, Len(strParametros) - 1)
 End Function

' This function handles -/+ binary operations
Private Function ParseNumExp() As Double
On Error GoTo ParseNumExp_ErrHandler

Dim Value As Double
Dim OtherValue As Double
Dim CurrToken As ParserTokens

    ' ParseTerm knows how to handle * and / operators,
    ' which must be executed first
    Value = ParseTerm
    
    ' While we didn't reach the expression's end,
    ' check for more +/- operators
    Do While mPosition <= Len(mExpression)
        
        ' GetToken just gives us a peek at the next token,
        ' It does not change the current position. We skip
        ' over the token ONLY IF WE CAN HANDLE IT in this
        ' function's scope
        CurrToken = GetToken

        If CurrToken = TOK_ADD Then
            ' We can handle the token, so let's skip over it
            ' and find the "other side" of the + operation
            SkipLastToken
            OtherValue = ParseTerm
            Value = Value + OtherValue
        ElseIf CurrToken = TOK_SUBTRACT Then
            SkipLastToken
            OtherValue = ParseTerm
            Value = Value - OtherValue
        ElseIf CurrToken = TOK_UNKNOWN Then
            err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG + " Character '" + Mid(mExpression, mPosition, 1) + "' Invalid."
        Else
            ' The operator is one not in the responsibility
            ' of this function - we can return up
            ParseNumExp = Value
            Exit Function
        End If
    Loop

    ParseNumExp = Value
    Exit Function
    
ParseNumExp_ErrHandler:
    SetErrSource "ParseNumExp"
    err.Raise err.Number
End Function

Private Function ParseTerm() As Double
On Error GoTo ParseTerm_ErrHandler

Dim Value As Double
Dim OtherValue As Double
Dim CurrToken As ParserTokens

    Value = ParseTermA
    
    ' While we didn't reach the expression's end,
    ' check for more * or / operators
    Do While mPosition <= Len(mExpression)
        
        CurrToken = GetToken

        If CurrToken = TOK_MULTIPLY Then
            SkipLastToken
            
            OtherValue = ParseTermA
            Value = Value * OtherValue
        ElseIf CurrToken = TOK_DIVIDE Then
            SkipLastToken
            
            OtherValue = ParseTermA
            If OtherValue = 0 Then
                err.Raise PERR_DIVISION_BY_ZERO, , _
                    "Division by Zero!"
            End If
            
            Value = Value / OtherValue
        ElseIf CurrToken = TOK_UNKNOWN Then
            err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG + " Character '" + Mid(mExpression, mPosition, 1) + "' Invalid."
        Else
            ParseTerm = Value
            Exit Function
        End If
    Loop

    ParseTerm = Value

    Exit Function
    
ParseTerm_ErrHandler:
    SetErrSource "ParseTerm"
    err.Raise err.Number
End Function

' This function reads a value that operators work on.
' The value can be a number, constant, function or a
' complete sub-expression (enclosed in parentheses (1+1) )
Private Function ParseValue() As Double
On Error GoTo ParseValue_ErrHandler

Dim Sign As Double
Dim CurrToken As ParserTokens
Dim Value As Double
Dim IsValue As Boolean

    Sign = 1

    CurrToken = GetToken
    If CurrToken = TOK_SUBTRACT Then
        ' We ran into an UNARY minus (like -1), so we
        ' have to multiply the next value with -1
        Sign = -1
        SkipLastToken
    ElseIf CurrToken = TOK_ADD Then
        ' Unary plus - no special meaning
        SkipLastToken
    End If

    CurrToken = GetToken
    If CurrToken = TOK_OPEN_PARENTHESES Then
        ' A sub-expression
        SkipLastToken
        ' Read the value of the sub-expression.
        ' When ParseLogExp runs into the closing parentheses,
        ' it will return (is the syntax is correct).
       
        Value = ParseLogExp
                
        CurrToken = GetToken
        If CurrToken = TOK_CLOSE_PARENTHESES Then
            SkipLastToken
        Else
            ' Where are those closing parentheses ?
            err.Raise PERR_CLOSING_PARENTHESES_EXPECTED, , "')' Expected"
        End If
    Else
        ' No sub-expression - It's an atom
        Value = ParseAtom
    End If

    ParseValue = Value * Sign
    Exit Function
    
ParseValue_ErrHandler:
    SetErrSource "ParseValue"
    err.Raise err.Number
End Function

' ParseAtom knows how to handle numbers, constants
' and functions
Private Function ParseAtom() As Double
On Error GoTo ParseAtom_ErrHandler

Dim CurrPosition As Long
Dim CurrToken As ParserTokens
Dim SymbolName As String
Dim ArgumentValue As Double
Dim DecimalPointFound As Boolean
Dim Value As Double
Dim IsValue As Boolean

    If mPosition > Len(mExpression) Then
        err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG + " Character '" + Mid(mExpression, mPosition - 1, 1) + "' Invalid."
    End If

    CurrPosition = mPosition
    
    ' We didn't recoginze a valid value yet
    IsValue = False
    
    ' Check if the atom is a number typed in explicitly
    If IsNumeric(Mid(mExpression, CurrPosition, 1)) Then
        IsValue = True
        
        CurrPosition = CurrPosition + 1
        DecimalPointFound = False
        
        ' Read the rest of the number
        Do While IsNumeric(Mid(mExpression, CurrPosition, 1)) Or _
            Mid(mExpression, CurrPosition, 1) = "."
            
            If Mid(mExpression, CurrPosition, 1) = "." Then
                If Not DecimalPointFound Then
                    DecimalPointFound = True
                Else
                    ' Can't have the decimal point twice!
                    err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG + " Character '" + Mid(mExpression, mPosition, 1) + "' Invalid."
                End If
            End If
            
            CurrPosition = CurrPosition + 1
        Loop
        
        Value = CDbl(Mid(mExpression, mPosition, CurrPosition - mPosition))
        
        mPosition = CurrPosition
        SkipSpaces
        
        'Verify if next token is Factorial and DecimalPointFound = true
        CurrToken = GetToken
        If CurrToken = TOK_FACTORIAL And DecimalPointFound = True Then
            'Can't solve factorial for non integers
            err.Raise PERR_SYNTAX_ERROR, , "Can't solve factorial for non integers."
        End If
    End If
    
    
    If Not IsValue Then
        ' Check if it's a constant/function name
        If IsLetter(Mid(mExpression, CurrPosition, 1)) Then
            CurrPosition = CurrPosition + 1
            
            ' Read the rest of the string. VB doesn't do
            ' "short-circuit" condition handling, so we have
            ' to put an If in the While loop
            Do While CurrPosition <= Len(mExpression)
                If IsValidSymbolCharacter(Mid(mExpression, CurrPosition, 1)) Then
                    CurrPosition = CurrPosition + 1
                Else
                    Exit Do
                End If
            Loop
                    
            SymbolName = Mid(mExpression, mPosition, CurrPosition - mPosition)
            mPosition = CurrPosition
            SkipSpaces
        
            ' If there are openning parentheses, it's a
            ' function call
            CurrToken = GetToken
            If CurrToken = TOK_OPEN_PARENTHESES Then
                SkipLastToken
                ' Get the argument to the function.
                ' Multi-argument functions are very
                ' easy to implement here.
                Value = CallBuiltinFunction(SymbolName)
                IsValue = True
            Else
                ' The symbol is supposed to be a constant
                ' name - check if it really exists
                If ConstExists(SymbolName) Then
                    Value = mConstants(SymbolName)
                    IsValue = True
                Else
                    err.Raise PERR_CONST_DOES_NOT_EXIST, , _
                        "Constant name " & SymbolName & " does not exist"
                End If
            End If
        End If
    End If

    If Not IsValue Then
        ' We didn't recognize the value
        err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG + " Character '" + Mid(mExpression, mPosition, 1) + "' Invalid."
    End If

    ParseAtom = Value
    Exit Function
    
ParseAtom_ErrHandler:
    SetErrSource "ParseAtom"
    err.Raise err.Number
End Function

Private Function GetToken() As ParserTokens
Dim CurrToken As ParserTokens
Dim i As ParserTokens
    
    If mPosition > Len(mExpression) Then
        GetToken = TOK_UNKNOWN
        Exit Function
    End If
    
    CurrToken = TOK_UNKNOWN
    mLastTokenLength = 0
    
    ' Iterate all known tokens and check if they match
    For i = TOK_FIRST To TOK_LAST
        If Mid(mExpression, mPosition, Len(mTokenSymbols(i))) = mTokenSymbols(i) Then
            CurrToken = i
            
            ' Save the token length so we can skip over it
            ' easily later
            mLastTokenLength = Len(mTokenSymbols(i))
            Exit For
        End If
    Next
    
    GetToken = CurrToken
End Function

Private Sub SkipLastToken()
    
    ' Skip over the last token, plus any spaces after it
    mPosition = mPosition + mLastTokenLength
    SkipSpaces

End Sub

'''''''''''''''''''''''''''''''
' Constants handling functions
'''''''''''''''''''''''''''''''

Private Function ConstExists(Name As String, Optional indAdicionaSiNoExiste As Boolean = True) As Boolean
Const ERR_KEY_NOT_FOUND = 5
Dim DummyValue As Double

    On Error Resume Next
    DummyValue = mConstants(Name)
    
    If err.Number = ERR_KEY_NOT_FOUND Then
        If indAdicionaSiNoExiste Then
            ConstExists = AddConstanteSistema(Name)
        Else
            ConstExists = False
        End If
    Else
        ConstExists = True
    End If
    
End Function

Private Sub obtenerDatosSistema(ByVal strSpDatosFormula As String, ByVal strParametros As String, ByRef strMsgError As String)
    Dim cn                  As ADODB.Connection
    Dim adoFormula          As ADODB.Command
    Dim intI As Integer
    Dim valoresParametros() As String
    
    On Error GoTo err
    
    Set cn = New ADODB.Connection
    Set adoFormula = New ADODB.Command
    
    cn.ConnectionString = gstrConnectConsulta
    cn.CursorLocation = adUseClient
    cn.Open
    adoFormula.CommandTimeout = 0
    adoFormula.ActiveConnection = cn
    
    Set adoConsultaFormula = New ADODB.Recordset
    
    valoresParametros = Split(strParametros, "|")
    
    adoFormula.CommandType = adCmdStoredProc
    adoFormula.CommandText = strSpDatosFormula
    For intI = 0 To UBound(valoresParametros)
        adoFormula.Parameters(intI + 1).Value = CStr(valoresParametros(intI))
    Next
    Set adoConsultaFormula = adoFormula.Execute
    
'    If cn.State = 1 Then cn.Close
'    Set cn = Nothing
'    Set adoFormula = Nothing
    
    Exit Sub
err:
'    If cn.State = 1 Then cn.Close
'    Set cn = Nothing
'    Set adoFormula = Nothing
    If strMsgError = "" Then strMsgError = err.Description
End Sub

Private Function AddConstanteSistema(ByVal Name As String) As Boolean
    Dim intI As Integer
    Dim indExiste As Boolean
    
    indExiste = False
    AddConstanteSistema = False
    If TypeName(adoConsultaFormula) = "Recordset" Then
        For intI = 0 To adoConsultaFormula.Fields.Count - 1
            If UCase(adoConsultaFormula.Fields(intI).Name) = Name Then
                indExiste = True
                Exit For
            End If
        Next
        If indExiste Then
            AddConstant Name, adoConsultaFormula(Name)
            AddConstanteSistema = True
        End If
    End If
End Function

Public Sub AddConstant(Name As String, Value As Double)
Dim i As ParserTokens
Dim TempName As String

    TempName = UCase(Trim(Name))
   
    ' Do all validity checks
    If Len(TempName) = 0 Then
        err.Raise PERR_INVALID_CONST_NAME, , "Constant name cannot be null"
    End If

    If Not IsLetter(Left(TempName, 1)) Then
        err.Raise PERR_INVALID_CONST_NAME, , "Constant name must begin with a letter"
    End If
            
    For i = 2 To Len(TempName)
        If Not IsValidSymbolCharacter(Mid(TempName, i, 1)) Then
            err.Raise PERR_INVALID_CONST_NAME, , "Invalid constant name"
        End If
    Next
    
    If ConstExists(TempName, False) Then
        err.Raise PERR_CONST_ALREADY_EXISTS, , "The constant already exists"
    End If
    
    If IsBuiltInFunction(TempName) Then
        err.Raise PERR_RESERVED_WORD, , "The name is a reserved word"
    End If
    
    mConstants.Add Value, TempName
    
End Sub

Public Function GetConstant(Name As String) As Double

    On Error Resume Next
    GetConstant = mConstants(Name)
    
End Function


Public Function ContantesCount() As Integer
ContantesCount = mConstants.Count
End Function

Public Sub RemoveConstant(Name As String)
Dim TempName As String

    TempName = UCase(Trim(Name))
    
    If ConstExists(TempName) Then
        mConstants.Remove TempName
    End If
    
End Sub

Public Sub RemoveConstantAll()

'    Dim a As Object
'
'    'Set a = mConstants.[_NewEnum]
'
'    For Each a In mConstants
'        a.Remove
'    Next
   
    Dim i As Integer
    
    For i = 1 To mConstants.Count
        mConstants.Remove 1
    Next i
   
   
End Sub

'''''''''''''''''''''''''''''''''
' Built-in function' functions...
'''''''''''''''''''''''''''''''''

' Check if a string name does stand for a supported built-in
' function - You may add as many as you like
Private Function IsBuiltInFunction(Name As String) As Boolean
Dim TempName As String

    TempName = UCase(Trim(Name))
    If TempName = "SIN" Or _
       TempName = "COS" Or _
       TempName = "ABS" Or _
       TempName = "ROUND" Or _
       TempName = "IIF" Then

        IsBuiltInFunction = True
    Else
        IsBuiltInFunction = False
    End If

End Function

' Execute the built-in function, and return its result
Private Function CallBuiltinFunction(Name As String) As Double
On Error GoTo CallBuiltinFunction_ErrHandler

Const DEGREES_TO_RADIANS = PI / 180
Dim TempName As String

    If Not IsBuiltInFunction(Name) Then
        err.Raise PERR_FUNCTION_DOES_NOT_EXIST, , _
            "Function " & Name & " Does not exist"
    End If
    
    TempName = UCase(Trim(Name))

    Select Case TempName
    Case "SIN"
         ' VB Trigonometric functions work with radians, so
         ' we need to convert the argument to radians
         CallBuiltinFunction = Sin(ParseParamList() * DEGREES_TO_RADIANS)
    Case "COS"
         CallBuiltinFunction = Cos(ParseParamList() * DEGREES_TO_RADIANS)
    Case "ABS"
         CallBuiltinFunction = Abs(ParseParamList())
    Case "ROUND"
         CallBuiltinFunction = Round(ParseParamList(), ParseParamList())
    Case "IIF"
         CallBuiltinFunction = IIf(ParseParamList(), ParseParamList(), ParseParamList())
    Case Else
        CallBuiltinFunction = 0
    
    End Select

    Exit Function
    
CallBuiltinFunction_ErrHandler:
    SetErrSource "CallBuiltinFunction"
    err.Raise err.Number
End Function

'''''''''''''''''''
' Helper functions
'''''''''''''''''''

Private Sub SkipSpaces()
    
    ' Skip spaces/tabs in the expression
    Do While mPosition <= Len(mExpression) And _
             (Mid(mExpression, mPosition, 1) = " " Or _
              Mid(mExpression, mPosition, 1) = vbTab)
        mPosition = mPosition + 1
    Loop

End Sub

' Check if a character is an english letter
Private Function IsLetter(Character As String) As Boolean
Dim CharAsciiCode As Long

    CharAsciiCode = Asc(UCase(Character))
    If (CharAsciiCode >= Asc("A") And _
        CharAsciiCode <= Asc("Z")) Then
        IsLetter = True
    Else
        IsLetter = False
    End If

End Function

' Check if a character is an english letter / a number /
' an underscore
Private Function IsValidSymbolCharacter(Character As String) As Boolean

    If IsLetter(Character) Or _
       IsNumeric(Character) Or _
       Character = "_" Then
            
        IsValidSymbolCharacter = True
    Else
        IsValidSymbolCharacter = False
    End If

End Function

'''''''''''''''''''''''''''
' Error handling functions
'''''''''''''''''''''''''''
Private Function GetProjectName() As String
On Error Resume Next

    ' Quite a way to get the project name...
    err.Raise 1, , " "
    GetProjectName = err.Source
    err.Clear

End Function

Private Sub SetErrSource(Name As String)
    
    If err.Source = mProjectName Then
        ' Error was "just raised", the supplied function name
        ' is the lowest function in the call stack
        err.Source = Name
    Else
        ' The error was propagated from a lower-level function.
        ' Add "this" function name to the call stack
        err.Source = Name & "->" & err.Source
    End If

End Sub

' This property can tell the programmer where the
' parser raised an error - Note that the value returned
' may not be what you expected... Experiment with
' syntax errors
Public Property Get LastErrorPosition() As Long
    LastErrorPosition = mPosition
End Property

'''''''''''''''''''''''''''
' Initialization
'''''''''''''''''''''''''''

Private Sub Class_Initialize()
    
    ' Initilalize token symbols
    ReDim mTokenSymbols(TOK_FIRST To TOK_LAST)

    mTokenSymbols(TOK_ADD) = "+"
    mTokenSymbols(TOK_SUBTRACT) = "-"
    mTokenSymbols(TOK_MULTIPLY) = "*"
    mTokenSymbols(TOK_DIVIDE) = "/"
    mTokenSymbols(TOK_POWER) = "^"
    mTokenSymbols(TOK_FACTORIAL) = "!"
    mTokenSymbols(TOK_PARAM) = ";"
    mTokenSymbols(TOK_NOT) = "NOT"
    mTokenSymbols(TOK_OR) = "OR"
    mTokenSymbols(TOK_AND) = "AND"
    mTokenSymbols(TOK_EQUAL) = "="
    mTokenSymbols(TOK_NOT_EQUAL) = "<>"
    mTokenSymbols(TOK_GREATER_THAN) = ">"
    mTokenSymbols(TOK_GREATER_OR_EQUAL_THAN) = ">="
    mTokenSymbols(TOK_LOWER_THAN) = "<"
    mTokenSymbols(TOK_LOWER_OR_EQUAL_THAN) = "<="
    mTokenSymbols(TOK_OPEN_PARENTHESES) = "("
    mTokenSymbols(TOK_CLOSE_PARENTHESES) = ")"

    ' Initilalize constants collection &
    ' add built-in constants
    Set mConstants = New Collection
    mConstants.Add PI, "PI"
    
    ' Get project name for proper error handling
    mProjectName = GetProjectName

End Sub

Public Function ParseTermA()
On Error GoTo ParseTermA_ErrHandler

Dim Value As Double
Dim OtherValue As Double
Dim CurrToken As ParserTokens

    Value = ParseTermB
    
    ' While we didn't reach the expression's end,
    ' check for more * or / operators
    Do While mPosition <= Len(mExpression)
        
        CurrToken = GetToken

        If CurrToken = TOK_POWER Then
            SkipLastToken
            'OtherValue = ParseValue
            OtherValue = ParseTermB
            Value = Value ^ OtherValue
        ElseIf CurrToken = TOK_UNKNOWN Then
            err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG + " Character '" + Mid(mExpression, mPosition, 1) + "' Invalid."
        Else
            ParseTermA = Value
            Exit Function
        End If
    Loop

    ParseTermA = Value

    Exit Function
    
ParseTermA_ErrHandler:
    SetErrSource "ParseTermA"
    err.Raise err.Number
End Function

Public Function ParseTermB()

On Error GoTo ParseTermB_ErrHandler

Dim Value As Double
Dim OtherValue As Double
Dim CurrToken As ParserTokens

    Value = ParseValue
    
    Do While mPosition <= Len(mExpression)
        
        CurrToken = GetToken

        If CurrToken = TOK_FACTORIAL Then
            Value = fact(CInt(Value))
            SkipLastToken
        ElseIf CurrToken = TOK_UNKNOWN Then
            err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG + " Character '" + Mid(mExpression, mPosition, 1) + "' Invalid."
        Else
            ParseTermB = Value
            Exit Function
        End If
    Loop

    ParseTermB = Value

    Exit Function
    
ParseTermB_ErrHandler:
    SetErrSource "ParseTermB"
    err.Raise err.Number

End Function
'Algo for factorial
Private Function fact(Num As Integer) As Double
    Dim B As Double
    B = 1
    For Num = 1 To Num
        B = B * Num 'I wish I could write it as b * = num :(
    Next
    fact = B

End Function

Public Function ParseParamList()
On Error GoTo ParseParamList_ErrHandler

Dim Value As Double
Dim OtherValue As Double
Dim CurrToken As ParserTokens
    
    Value = ParseLogExp
    
    If mPosition > Len(mExpression) Then
        err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG + " Character '" + Mid(mExpression, mPosition - 1, 1) + "' Invalid."
    End If
         
    CurrToken = GetToken

    If CurrToken = TOK_PARAM Then
        SkipLastToken
        
        ParseParamList = Value
        Exit Function

    ElseIf CurrToken = TOK_CLOSE_PARENTHESES Then
        SkipLastToken
        ParseParamList = Value
        Exit Function
    
    Else
        err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG + " Character '" + Mid(mExpression, mPosition, 1) + "' Invalid."
    End If

    ParseParamList = Value
    Exit Function
    
ParseParamList_ErrHandler:
    SetErrSource "ParseParamList"
    err.Raise err.Number

End Function

Public Function ParseLogExp()
On Error GoTo ParseLogExp_ErrHandler

Dim Value As Double
Dim OtherValue As Double
Dim CurrToken As ParserTokens

        
    Value = ParseProposition
    
    Do While mPosition <= Len(mExpression)
        
        ' GetToken just gives us a peek at the next token,
        ' It does not change the current position. We skip
        ' over the token ONLY IF WE CAN HANDLE IT in this
        ' function's scope
        CurrToken = GetToken

        If CurrToken = TOK_OR Then
            ' We can handle the token, so let's skip over it
            ' and find the "other side" of the + operation
            SkipLastToken
            OtherValue = ParseProposition
            Value = Value Or OtherValue
        ElseIf CurrToken = TOK_UNKNOWN Then
            err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG + " Character '" + Mid(mExpression, mPosition, 1) + "' Invalid."
        Else
            ' The operator is one not in the responsibility
            ' of this function - we can return up
            ParseLogExp = Value
            Exit Function
        End If
    Loop

    ParseLogExp = Value
    Exit Function
    
ParseLogExp_ErrHandler:
    SetErrSource "ParseLogExp"
    err.Raise err.Number

End Function

Public Function ParseProposition()
On Error GoTo ParseProposition_ErrHandler

Dim Value As Double
Dim OtherValue As Double
Dim CurrToken As ParserTokens

    Value = ParsePropositionA
    
    Do While mPosition <= Len(mExpression)
        
        ' GetToken just gives us a peek at the next token,
        ' It does not change the current position. We skip
        ' over the token ONLY IF WE CAN HANDLE IT in this
        ' function's scope
        CurrToken = GetToken

        If CurrToken = TOK_AND Then
            SkipLastToken
            OtherValue = ParsePropositionA
            Value = Value And OtherValue
        ElseIf CurrToken = TOK_UNKNOWN Then
            err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG + " Character '" + Mid(mExpression, mPosition, 1) + "' Invalid."
        Else
            ' The operator is one not in the responsibility
            ' of this function - we can return up
            ParseProposition = Value
            Exit Function
        End If
    Loop

    ParseProposition = Value
    Exit Function
    
ParseProposition_ErrHandler:
    SetErrSource "ParseProposition"
    err.Raise err.Number

End Function

Public Function ParsePropositionA()
On Error GoTo ParsePropositionA_ErrHandler

Dim Value As Double
Dim OtherValue As Double
Dim CurrToken As ParserTokens

    CurrToken = GetToken
    If CurrToken = TOK_NOT Then
        SkipLastToken
        Value = ParsePropositionB
        Value = Not Value
    Else
        Value = ParsePropositionB
    End If
        
    ParsePropositionA = Value
    Exit Function
    
ParsePropositionA_ErrHandler:
    SetErrSource "ParsePropositionA"
    err.Raise err.Number

End Function

Public Function ParsePropositionB()
On Error GoTo ParsePropositionB_ErrHandler

Dim Value As Double
Dim OtherValue As Double
Dim CurrToken As ParserTokens

    Value = ParseNumExp
    
    CurrToken = GetToken
    If CurrToken = TOK_EQUAL Then
        SkipLastToken
        OtherValue = ParseNumExp
        Value = IIf(Value = OtherValue, True, False)
       
    ElseIf CurrToken = TOK_NOT_EQUAL Then
        SkipLastToken
        OtherValue = ParseNumExp
        Value = IIf(Value <> OtherValue, True, False)
    
    ElseIf CurrToken = TOK_GREATER_THAN Then
        SkipLastToken
        OtherValue = ParseNumExp
        Value = IIf(Value > OtherValue, True, False)
    
    ElseIf CurrToken = TOK_GREATER_OR_EQUAL_THAN Then
        SkipLastToken
        OtherValue = ParseNumExp
        Value = IIf(Value >= OtherValue, True, False)
    
    ElseIf CurrToken = TOK_LOWER_THAN Then
        SkipLastToken
        OtherValue = ParseNumExp
        Value = IIf(Value < OtherValue, True, False)
        
    ElseIf CurrToken = TOK_LOWER_OR_EQUAL_THAN Then
        SkipLastToken
        OtherValue = ParseNumExp
        Value = IIf(Value <= OtherValue, True, False)
   
    Else
        ParsePropositionB = Value
        Exit Function
    
    End If
    
    ParsePropositionB = Value
    Exit Function
    
ParsePropositionB_ErrHandler:
    SetErrSource "ParsePropositionB"
    err.Raise err.Number
End Function


