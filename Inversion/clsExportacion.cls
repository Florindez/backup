VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExportacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'local variable(s) to hold property value(s)
Private m_Titulo As String                    'local copy
Private m_SubTitulo As String              'local copy
Private m_TDBGrilla As TDBGrid            'local copy
Private m_DTGrilla As DataGrid             'local copy
Private m_RS As ADODB.Recordset        'local copy
Private m_Posicion As Long

Private Type ExlCell
        Row As Long
        Col As Long
End Type

Public Property Set rs(ByRef vData As ADODB.Recordset)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Rs = Form1
    Set m_RS = vData
    
End Property

Public Property Get rs() As ADODB.Recordset
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Rs
    Set rs = m_RS
End Property

Public Property Set DTGrilla(ByVal vData As DataGrid)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DTGrilla = Form1
    Set m_DTGrilla = vData
End Property

Public Property Get DTGrilla() As DataGrid
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DTGrilla
    Set DTGrilla = m_DTGrilla
End Property

Public Property Set TDBGrilla(ByVal vData As TDBGrid)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.TDBGrilla = Form1
    Set m_TDBGrilla = vData
End Property

Public Property Get TDBGrilla() As TDBGrid
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TDBGrilla
    Set TDBGrilla = m_TDBGrilla
End Property


Public Function ExportaRecordSetExcel(ByRef rec As ADODB.Recordset, Optional ByVal opImprimirTitulo As String = "", Optional ByVal opSavePath As String = "") As Boolean
    
    On Error GoTo errSub
    
    Dim Excel       As Object
    Dim Libro       As Object
    Dim Hoja        As Object
    Dim arrData     As Variant
    Dim iRec        As Long
    Dim iCol        As Integer
    Dim iRow        As Integer
            
    rec.MoveFirst
    
    ' -- Crear los objetos para utilizar el Excel
    Set Excel = CreateObject("Excel.Application")
    Set Libro = Excel.Workbooks.Add
    
    ' -- Hacer referencia a la hoja
    Set Hoja = Libro.Worksheets(1)

    iCol = rec.Fields.Count
    
    If opImprimirTitulo = Valor_Caracter Then
    
        For iCol = 1 To rec.Fields.Count
        Hoja.Cells(1, iCol).Value = rec.Fields(iCol - 1).Name
        Next
    
        If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
        Hoja.Cells(2, 1).CopyFromRecordset rec
        Else

        arrData = rec.GetRows

        iRec = UBound(arrData, 2) + 1
        
        For iCol = 0 To rec.Fields.Count - 1
            For iRow = 0 To iRec - 1

                If IsDate(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = Format(arrData(iCol, iRow))

                ElseIf IsArray(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = "Array Field"
                End If
            Next iRow
        Next iCol
            
        ' -- Traspasa los datos a la hoja de Excel
        Hoja.Cells(2, 1).Resize(iRec, rec.Fields.Count).Value = GetData(arrData)
        
        End If
        
    Else
         
        For iCol = 1 To rec.Fields.Count
        Hoja.Cells(4, iCol).Value = rec.Fields(iCol - 1).Name
        
        Next
        
        If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
        Hoja.Cells(5, 1).CopyFromRecordset rec
        Else

        arrData = rec.GetRows

        iRec = UBound(arrData, 2) + 1
        
        For iCol = 0 To rec.Fields.Count - 1
            For iRow = 0 To iRec - 1

                If IsDate(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = Format(arrData(iCol, iRow))

                ElseIf IsArray(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = "Array Field"
                End If
            Next iRow
        Next iCol
            
        ' -- Traspasa los datos a la hoja de Excel
        Hoja.Cells(5, 1).Resize(iRec, rec.Fields.Count).Value = GetData(arrData)
        
        End If
        
    End If
         
    If opImprimirTitulo = Valor_Caracter Then
        
        Hoja.Range(Hoja.Cells(2, 1), _
        Hoja.Cells((2 + rec.RecordCount) - 1, 1 + rec.Fields.Count)).Select
        
        Excel.Selection.CurrentRegion.Columns.AutoFit

        Excel.Selection.CurrentRegion.Rows.AutoFit
    
            'Columnas en negrita
        Hoja.Range(Hoja.Cells(1, 1), _
        Hoja.Cells(1, 1 + rec.Fields.Count)).Font.Bold = True
        
        
       
    Else
    
        Hoja.Range(Hoja.Cells(5, 1), _
        Hoja.Cells((5 + rec.RecordCount) - 1, 1 + rec.Fields.Count)).Select
        
        Excel.Selection.CurrentRegion.Columns.AutoFit

        Excel.Selection.CurrentRegion.Rows.AutoFit
    
        'TITULO
        Hoja.Range(Hoja.Cells(1, 1), Hoja.Cells(1, iCol)).MergeCells = True
        Hoja.Range(Hoja.Cells(1, 1), Hoja.Cells(1, iCol)).Value = opImprimirTitulo
        
        Hoja.Range(Hoja.Cells(1, 1), Hoja.Cells(1, iCol)).HorizontalAlignment = xlHAlignCenter
        
        Hoja.Range(Hoja.Cells(1, 1), Hoja.Cells(1, iCol)).Font.Bold = True
        
        Hoja.Range(Hoja.Cells(1, 1), Hoja.Cells(1, iCol)).Font.Size = 15
    
        'Columnas en negrita
        Hoja.Range(Hoja.Cells(4, 1), _
        Hoja.Cells(4, 1 + rec.Fields.Count)).Font.Bold = True
        
    End If
        
'        stCell.Row = 1
'    stCell.Col = 1
    
    ' -- Cierra el recordset y la base de datos y los objetos ADO
    rec.Close
    
    Set rec = Nothing
    
    ' -- guardar el libro o mostrarlo
        If opSavePath = Valor_Caracter Then
            
            Excel.WindowState = xlMaximized
            
            Excel.Visible = True: Excel.UserControl = True
    
        Else
            
            Excel.WindowState = xlMaximized
            
            Libro.SaveAs opSavePath
            
            Libro.Close
             '-- Elimina las referencias Xls
            Set Hoja = Nothing
            Set Libro = Nothing
            Excel.Quit
            Set Excel = Nothing
        
        End If
   
    ExportaRecordSetExcel = True
    Exit Function
errSub:
    MsgBox err.Description, vbCritical, "Error"
    ExportaRecordSetExcel = False
End Function


'Public Function ExportaRecordSetExcel(ByVal Sql As String, sOutputPathXLS As String) As Boolean
'
'    On Error GoTo errSub
'
'    Dim rec         As New ADODB.Recordset
'    Dim Excel       As Object
'    Dim Libro       As Object
'    Dim Hoja        As Object
'    Dim arrData     As Variant
'    Dim iRec        As Long
'    Dim iCol        As Integer
'    Dim iRow        As Integer
'
'    'Me.Enabled = False
'
'    ' -- Abrir el Recordset pasándole la cadena sql
'    rec.Open Sql, adoConn
'
'    ' -- Crear los objetos para utilizar el Excel
'    Set Excel = CreateObject("Excel.Application")
'    Set Libro = Excel.Workbooks.Add
'
'    ' -- Hacer referencia a la hoja
'    Set Hoja = Libro.Worksheets(1)
'
'    iCol = rec.Fields.Count
'    For iCol = 1 To rec.Fields.Count
'        Hoja.Cells(1, iCol).Value = rec.Fields(iCol - 1).Name
'    Next
'
'    If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
'        Hoja.Cells(2, 1).CopyFromRecordset rec
'    Else
'
'        arrData = rec.GetRows
'
'        iRec = UBound(arrData, 2) + 1
'
'        For iCol = 0 To rec.Fields.Count - 1
'            For iRow = 0 To iRec - 1
'
'                If IsDate(arrData(iCol, iRow)) Then
'                    arrData(iCol, iRow) = Format(arrData(iCol, iRow))
'
'                ElseIf IsArray(arrData(iCol, iRow)) Then
'                    arrData(iCol, iRow) = "Array Field"
'                End If
'            Next iRow
'        Next iCol
'
'        ' -- Traspasa los datos a la hoja de Excel
'        Hoja.Cells(2, 1).Resize(iRec, rec.Fields.Count).Value = GetData(arrData)
'
'    End If
'
'    Excel.Selection.CurrentRegion.Columns.AutoFit
''    Excel.Selection.CurrentRegion.Columns.Font.Bold = True
'    Excel.Selection.CurrentRegion.Rows.AutoFit
'
'    'Columnas en negrita
'     Hoja.Range(Hoja.Cells(1, 1), _
'        Hoja.Cells(1, 1 + rec.Fields.Count)).Font.Bold = True
'
'
''        stCell.Row = 1
''    stCell.Col = 1
'
'    ' -- Cierra el recordset y la base de datos y los objetos ADO
'    rec.Close
'
'    Set rec = Nothing
'    ' -- guardar el libro
'    Libro.SaveAs sOutputPathXLS
'
'    Excel.Visible = True: Excel.UserControl = True
'    'Libro.Close
'    ' -- Elimina las referencias Xls
'    'Set Hoja = Nothing
'    'Set Libro = Nothing
'    'Excel.Quit
'    'Set Excel = Nothing
'
'    ExportaRecordSetExcel = True
'    Exit Function
'errSub:
'    MsgBox err.Description, vbCritical, "Error"
'    ExportaRecordSetExcel = False
'End Function

Private Function GetData(vValue As Variant) As Variant
    Dim X As Long, Y As Long, xMax As Long, yMax As Long, t As Variant
    
    xMax = UBound(vValue, 2): yMax = UBound(vValue, 1)
    
    ReDim t(xMax, yMax)
    For X = 0 To xMax
        For Y = 0 To yMax
            t(X, Y) = vValue(Y, X)
        Next Y
    Next X
    
    GetData = t
End Function

Public Function ExportaGrillaExcel() As Long
    
    Dim oExcel As Object                  ' Variable objeto EXCEL
    Dim objExlSht As Object             ' Variable objeto para el OLE automation
    Dim stCell As ExlCell                   ' Variable de tipo celda del objeto EXCEL
    Dim SomeArray() As Variant      ' Arreglo que almacenará los valores antes de la exportación
    Dim Row As Long                       ' Número de registros a exportar
    Dim Col As Long                        ' Número de columnas a exportar
    Dim iCol As Long                      ' Variable puntero de columna
    
    If rs Is Nothing Then Exit Function
    If DTGrilla Is Nothing And TDBGrilla Is Nothing Then Exit Function
    
    'Guardo la posición en el recordset
    If rs.AbsolutePosition = -1 Then Exit Function
    m_Posicion = rs.AbsolutePosition

    
    On Error GoTo ERRORNOEXCEL
    
    ' Creo objeto EXCEL
    Set oExcel = CreateObject("Excel.Application")
    ' Agrego una hoja de cálculo en el objeto EXCEL
    
    'On Error GoTo 0
   
    oExcel.Workbooks.Add
    ' Indico la hoja de cálculo activa
    Set objExlSht = oExcel.ActiveWorkbook.Sheets(1)
    ' Me ubico en la Celda 1, Columna 1 de la hoja de cálculo
    stCell.Row = 1
    stCell.Col = 1
    ' Ejecuto la exportación a EXCEL
    
    rs.MoveLast
    ReDim SomeArray(rs.RecordCount + 3, rs.Fields.Count)
     'Copy column headers to array
    
    Col = 0
    SomeArray(0, 0) = Trim(TITULO)
    SomeArray(1, 0) = Trim(SubTitulo)
    SomeArray(2, 0) = ""

    If TDBGrilla Is Nothing Then

        'Verifico si la grilla tiene splits
        If DTGrilla.Splits.Count = 1 Then
            'Agrego los títulos de las columnas
            For iCol = 0 To DTGrilla.Columns.Count - 1
                If DTGrilla.Columns(iCol).Visible Then
                    SomeArray(3, Col) = Trim(DTGrilla.Columns(iCol).Caption)
                    Col = Col + 1
                End If
            Next

            ' Redimensiono el array con las celdas necesarias
            ReDim Preserve SomeArray(rs.RecordCount + 3, Col - 1)

            Col = 0
            rs.MoveFirst
            For Row = 4 To rs.RecordCount + 3
                For iCol = 0 To DTGrilla.Columns.Count - 1
                    ' Si la columna de la grilla esta visible
                    If DTGrilla.Columns(iCol).Visible Then

                        ' Asigno un valor string
                        Select Case rs(iCol).Type
                            Case adBigInt, adDecimal, adDouble, adDecimal, adInteger, adNumeric, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt, adVarNumeric:
                                SomeArray(Row, Col) = DTGrilla.Columns(iCol).Text
                            Case Else: SomeArray(Row, Col) = "'" & Trim(CStr(DTGrilla.Columns(iCol).Text))
                        End Select

                        ' Valido los valores nulos que pudieran haber
                        If IsNull(SomeArray(Row, Col)) Then _
                        SomeArray(Row, Col) = ""
                        Col = Col + 1
                    End If
                Next
                rs.MoveNext
                Col = 0
            Next
        Else
            'La grilla tiene splits

            'Agrego los títulos de las columnas
            Dim iSplits As Integer
            For iSplits = 0 To DTGrilla.Splits.Count - 1
                For iCol = 0 To DTGrilla.Splits(iSplits).Columns.Count - 1
                    If DTGrilla.Splits(iSplits).Columns(iCol).Visible Then
                        SomeArray(3, Col) = Trim(DTGrilla.Splits(iSplits).Columns(iCol).Caption)
                        Col = Col + 1
                    End If
                Next
            Next

            ' Redimensiono el array con las celdas necesarias
            ReDim Preserve SomeArray(rs.RecordCount + 3, Col - 1)

            Col = 0
            rs.MoveFirst
            For Row = 4 To rs.RecordCount + 3
                For iSplits = 0 To DTGrilla.Splits.Count - 1
                    For iCol = 0 To DTGrilla.Splits(iSplits).Columns.Count - 1
                        ' Si la columna de la grilla esta visible
                        If DTGrilla.Splits(iSplits).Columns(iCol).Visible Then
                            ' Asigno un valor string
                            Select Case rs(iCol).Type
                                Case adBigInt, adDecimal, adDouble, adDecimal, adInteger, adNumeric, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt, adVarNumeric:
                                    SomeArray(Row, Col) = DTGrilla.Splits(iSplits).Columns(iCol).Text
                                Case Else: SomeArray(Row, Col) = "'" & Trim(CStr(DTGrilla.Splits(iSplits).Columns(iCol).Text))
                            End Select

                            ' Valido los valores nulos que pudieran haber
                            If IsNull(SomeArray(Row, Col)) Then _
                            SomeArray(Row, Col) = ""
                            Col = Col + 1
                        End If
                    Next
                Next
                rs.MoveNext
                Col = 0
            Next
            Col = 0
        End If

    Else
        'Verifico si la grilla tiene splits

        If TDBGrilla.Splits.Count = 1 Then
            'Agrego los títulos de las columnas
            For iCol = 0 To TDBGrilla.Columns.Count - 1
                If TDBGrilla.Columns(iCol).Visible Then
                    SomeArray(3, Col) = Trim(TDBGrilla.Columns(iCol).Caption)
                    Col = Col + 1
                End If
            Next

            ' Redimensiono el array con las celdas necesarias
            ReDim Preserve SomeArray(rs.RecordCount + 3, Col - 1)

            Col = 0
            rs.MoveFirst
            For Row = 4 To rs.RecordCount + 3
                For iCol = 0 To TDBGrilla.Columns.Count - 1
                    ' Si la columna de la grilla esta visible
                    If TDBGrilla.Columns(iCol).Visible Then

                        ' Asigno un valor string
                        Select Case rs(iCol).Type
                            Case adBigInt, adDecimal, adDouble, adDecimal, adInteger, adNumeric, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt, adVarNumeric:
                                SomeArray(Row, Col) = TDBGrilla.Columns(iCol).Text
                            Case Else: SomeArray(Row, Col) = "'" & Trim(CStr(TDBGrilla.Columns(iCol).Text))
                        End Select

                        ' Valido los valores nulos que pudieran haber
                        If IsNull(SomeArray(Row, Col)) Then _
                        SomeArray(Row, Col) = ""
                        Col = Col + 1
                    End If
                Next
                rs.MoveNext
'                TDBGrilla.Row = TDBGrilla.Row + 1
                Col = 0
            Next
        Else
            'La grilla tiene splits

            'Agrego los títulos de las columnas
            For iSplits = 0 To TDBGrilla.Splits.Count - 1
                For iCol = 0 To TDBGrilla.Splits(iSplits).Columns.Count - 1
                    If TDBGrilla.Splits(iSplits).Columns(iCol).Visible Then
                        SomeArray(3, Col) = Trim(TDBGrilla.Splits(iSplits).Columns(iCol).Caption)
                        Col = Col + 1
                    End If
                Next
            Next

            ' Redimensiono el array con las celdas necesarias
            ReDim Preserve SomeArray(rs.RecordCount + 3, Col - 1)
            
            Col = 0
            rs.MoveFirst
            For Row = 4 To rs.RecordCount + 3
                For iSplits = 0 To TDBGrilla.Splits.Count - 1
                    For iCol = 0 To TDBGrilla.Splits(iSplits).Columns.Count - 1
                        ' Si la columna de la grilla esta visible
                        If TDBGrilla.Splits(iSplits).Columns(iCol).Visible Then
                             
'                           Asigno un valor string
                            Select Case rs(iCol).Type
                                Case adBigInt, adDecimal, adDouble, adDecimal, adInteger, adNumeric, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt, adVarNumeric:
                                    SomeArray(Row, Col) = TDBGrilla.Splits(iSplits).Columns(iCol).Text
                                Case Else: SomeArray(Row, Col) = "'" & Trim(CStr(TDBGrilla.Splits(iSplits).Columns(iCol).Text))
                            End Select
                             
                            
                            ' Valido los valores nulos que pudieran haber
                            If IsNull(SomeArray(Row, Col)) Then _
                            SomeArray(Row, Col) = ""
                            Col = Col + 1
                        End If
                    Next
                Next
                rs.MoveNext
                Col = 0
            Next
            Col = 0
        End If
    End If
    
    'Restauro la posición que tenía en el recordset
    rs.AbsolutePosition = m_Posicion
    
    ' Selecciono toda la hoja
    oExcel.ActiveSheet.Cells.Select
    
    ' Pego en el rango de celdas seleccionadas los valores del array somearray
    objExlSht.Range(objExlSht.Cells(stCell.Row, stCell.Col), _
        objExlSht.Cells(stCell.Row + rs.RecordCount + 3, _
        stCell.Col + UBound(SomeArray, 2))).Value = SomeArray
     
    ' Tamaño de la fuente para el Título
    objExlSht.Range(objExlSht.Cells(stCell.Row, stCell.Col), _
        objExlSht.Cells(stCell.Row, stCell.Col)).Font.Size = 14
    
    ' Fuente para el Título
    objExlSht.Range(objExlSht.Cells(stCell.Row, stCell.Col), _
        objExlSht.Cells(stCell.Row, stCell.Col)).Font.Name = "Italic"
    
    ' Tamaño de la fuente para el subtítulo
    objExlSht.Range(objExlSht.Cells(stCell.Row + 1, stCell.Col), _
        objExlSht.Cells(stCell.Row + 1, stCell.Col)).Font.Size = 12
    
    ' Fuente para el Subtítulo
    objExlSht.Range(objExlSht.Cells(stCell.Row + 1, stCell.Col), _
        objExlSht.Cells(stCell.Row + 1, stCell.Col)).Font.Name = "Italic"
    
    ' Pongo en Negrita el Título de los Campos
    objExlSht.Range(objExlSht.Cells(stCell.Row, stCell.Col), _
        objExlSht.Cells(stCell.Row + 3, stCell.Col + UBound(SomeArray, 2))).Font.Bold = True
        
    ' Tamaño de la fuente para los Títulos de los Campos
    objExlSht.Range(objExlSht.Cells(stCell.Row + 3, stCell.Col), _
        objExlSht.Cells(stCell.Row + 3, stCell.Col + UBound(SomeArray, 2))).Font.Size = 10
        
    ' Selecciono toda la hoja
    oExcel.ActiveSheet.Cells.Select
    
    ' Formateo el tipo de celdas
    oExcel.Selection.NumberFormat = "#,##0.00"
    
    ' Activo el Autofit
    oExcel.Selection.Columns.AutoFit
    
    'Configuro la hoja para la impresión
    On Error Resume Next
    
    With oExcel.ActiveSheet
        If TITULO = "" Then
            .Name = "Hoja1"
        Else
            .Name = Left(Trim(TITULO), 31)
        End If
        .PageSetup.PrintTitleRows = "$1:$4"
        .PageSetup.PrintTitleColumns = ""
        .PageSetup.PrintArea = ""
        .PageSetup.LeftHeader = ""
        .PageSetup.CenterHeader = "&A"
        .PageSetup.RightHeader = ""
        .PageSetup.LeftFooter = ""
        .PageSetup.CenterFooter = "Página &P"
        .PageSetup.RightFooter = ""
        .PageSetup.LeftMargin = oExcel.Application.InchesToPoints(0.787401575)
        .PageSetup.RightMargin = oExcel.Application.InchesToPoints(0.787401575)
        .PageSetup.TopMargin = oExcel.Application.InchesToPoints(0.984251969)
        .PageSetup.BottomMargin = oExcel.Application.InchesToPoints(0.984251969)
        .PageSetup.HeaderMargin = oExcel.Application.InchesToPoints(0.511811024)
        .PageSetup.FooterMargin = oExcel.Application.InchesToPoints(0.511811024)
        .PageSetup.PrintHeadings = False
        .PageSetup.PrintGridlines = False
        .PageSetup.PrintNotes = False
        .PageSetup.CenterHorizontally = False
        .PageSetup.CenterVertically = False
        .PageSetup.Orientation = xlPortrait
        .PageSetup.Draft = False
        .PageSetup.PaperSize = xlPaperA4
        .PageSetup.FirstPageNumber = xlAutomatic
        .PageSetup.Order = xlDownThenOver
        .PageSetup.BlackAndWhite = False
        .PageSetup.Zoom = 100
    End With
    
    On Error GoTo 0
    
    ' Minimizo Excel
    oExcel.WindowState = xlMinimized
    
    'Hago visible Excel
    oExcel.Visible = True
    
    ' Grabo la hoja de cálculo
    oExcel.Application.Dialogs(xlDialogSaveAs).Show
    
    'Oculto Excel
    oExcel.Visible = False
    
    ' Cierro el objeto EXCEL y lo remuevo de memoria!
    objExlSht.Application.Quit
    
     Set oExcel = Nothing                ' Fin objeto EXCEL
     Set objExlSht = Nothing             ' Fin objeto para el OLE automation

    Screen.MousePointer = vbNormal
    
    Exit Function

ERRORNOEXCEL:
    Screen.MousePointer = vbNormal
    Select Case err.Number
           Case 0: Exit Function
'           Case Else: MsgBox "No se pudo iniciar instancia de Excel debido a un problema de conexión con el Excel, verifique si el Excel o de alguno de sus componentes se encuentran correctamente instalados", vbCritical, ""
            Case Else: MsgBox err.Description, vbCritical
            
    End Select
    Screen.MousePointer = vbNormal
End Function

Public Function ExportaCalendariosExcel(rsCabecera As ADODB.Recordset, rsDetalle As ADODB.Recordset) As Long
   Dim oExcel As Object                  ' Variable objeto EXCEL
   Dim objExlSht As Object               ' Variable objeto para el OLE automation
   Dim stCell As ExlCell                 ' Variable de tipo celda del objeto EXCEL
   Dim CabeceraArray() As Variant        ' Arreglo que almacenará los valores antes de la exportación
   Dim DetalleArray() As Variant         ' Arreglo que almacenará los valores antes de la exportación
   Dim Row As Long                       ' Número de registros a exportar
   Dim Col As Long                       ' Número de columnas a exportar
   Dim iCol As Long                      ' Variable puntero de columna
   Dim lContador As Long
    
   On Error GoTo ERRORNOEXCEL
   'Creo objeto EXCEL
   Set oExcel = CreateObject("Excel.Application")
   'Agrego una hoja de cálculo en el objeto EXCEL
   
   On Error GoTo 0
   oExcel.Workbooks.Add
   
   'Cabecera de los calendarios
   ReDim CabeceraArray(rsCabecera.RecordCount, 8)
   CabeceraArray(0, 0) = "Línea"
   CabeceraArray(0, 1) = "Préstamo"
   CabeceraArray(0, 2) = "IFI"
   CabeceraArray(0, 3) = "Beneficiario"
   CabeceraArray(0, 4) = "Fecha de Desembolso"
   CabeceraArray(0, 5) = "Moneda"
   CabeceraArray(0, 6) = "Monto Desembolsado"
   CabeceraArray(0, 7) = "Monto en VAC"
   CabeceraArray(0, 8) = "Indice VAC"
   
   rsCabecera.MoveFirst
   ReDim Preserve CabeceraArray(rsCabecera.RecordCount, 8)
   For Row = 1 To rsCabecera.RecordCount
      For iCol = 0 To rsCabecera.Fields.Count - 1
         If rsCabecera.Fields.Item(iCol).Name <> "CodigoLinea" Then
            CabeceraArray(Row, iCol - 1) = Trim(CStr(rsCabecera.Fields.Item(iCol).Value))
         End If
      Next
      rsCabecera.MoveNext
   Next
         
   'Detalle de los calendarios
   ReDim DetalleArray(rsDetalle.RecordCount, 9)
   DetalleArray(0, 0) = "Préstamo"
   DetalleArray(0, 1) = "Nº"
   DetalleArray(0, 2) = "Fecha de Vencimiento"
   DetalleArray(0, 3) = "Días"
   DetalleArray(0, 4) = "Moneda"
   DetalleArray(0, 5) = "Principal"
   DetalleArray(0, 6) = "Interés"
   DetalleArray(0, 7) = "Comisiones"
   DetalleArray(0, 8) = "Monto a Cobrar"
   DetalleArray(0, 9) = "Principal por Vencer"
   
   rsDetalle.MoveFirst
   ReDim Preserve DetalleArray(rsDetalle.RecordCount, 9)
   For Row = 1 To rsDetalle.RecordCount
      For iCol = 0 To rsDetalle.Fields.Count - 1
         DetalleArray(Row, iCol) = Trim(CStr(IIf(IsNull(rsDetalle.Fields.Item(iCol).Value), "", rsDetalle.Fields.Item(iCol).Value)))
      Next
      rsDetalle.MoveNext
    Next
      
   'Indico la hoja de cálculo activa
   Set objExlSht = oExcel.ActiveWorkbook.Sheets(1)
   oExcel.Sheets(1).Select
   'Pego en el rango de celdas seleccionadas los valores del array CabeceraArray
   oExcel.ActiveSheet.Range("A1", oExcel.ActiveSheet.Cells(UBound(CabeceraArray, 1) + 1, UBound(CabeceraArray, 2) + 1)).Value = CabeceraArray
   
   'Indico la hoja de cálculo activa
   Set objExlSht = oExcel.ActiveWorkbook.Sheets(2)
   oExcel.Sheets(2).Select
   'Pego en el rango de celdas seleccionadas los valores del array DetalleArray
   oExcel.ActiveSheet.Range("A1", oExcel.ActiveSheet.Cells(UBound(DetalleArray, 1) + 1, UBound(DetalleArray, 2) + 1)).Value = DetalleArray
   
   'Configuro la hoja para la impresión
   On Error Resume Next
   With oExcel.ActiveSheet
      .PageSetup.PrintTitleRows = "$1:$4"
      .PageSetup.PrintTitleColumns = ""
      .PageSetup.PrintArea = ""
      .PageSetup.LeftHeader = ""
      .PageSetup.CenterHeader = "&A"
      .PageSetup.RightHeader = ""
      .PageSetup.LeftFooter = ""
      .PageSetup.CenterFooter = "Página &P"
      .PageSetup.RightFooter = ""
      .PageSetup.LeftMargin = oExcel.Application.InchesToPoints(0.787401575)
      .PageSetup.RightMargin = oExcel.Application.InchesToPoints(0.787401575)
      .PageSetup.TopMargin = oExcel.Application.InchesToPoints(0.984251969)
      .PageSetup.BottomMargin = oExcel.Application.InchesToPoints(0.984251969)
      .PageSetup.HeaderMargin = oExcel.Application.InchesToPoints(0.511811024)
      .PageSetup.FooterMargin = oExcel.Application.InchesToPoints(0.511811024)
      .PageSetup.PrintHeadings = False
      .PageSetup.PrintGridlines = False
      .PageSetup.PrintNotes = False
      .PageSetup.CenterHorizontally = False
      .PageSetup.CenterVertically = False
      .PageSetup.Orientation = xlPortrait
      .PageSetup.Draft = False
      .PageSetup.PaperSize = xlPaperA4
      .PageSetup.FirstPageNumber = xlAutomatic
      .PageSetup.Order = xlDownThenOver
      .PageSetup.BlackAndWhite = False
      .PageSetup.Zoom = 100
   End With
    
   On Error GoTo 0
    
   'Grabo la hoja de cálculo
   oExcel.ActiveWorkbook.SaveAs FileName:=App.Path & "\HOJACALCULO.xls"

   'Cierro el objeto EXCEL y lo remuevo de memoria!
   objExlSht.Application.Quit
   Set objExlSht = Nothing
   
   Exit Function

ERRORNOEXCEL:
    Screen.MousePointer = vbNormal
    Select Case err.Number
           Case 0: Exit Function
           Case Else: MsgBox "No se pudo iniciar instancia de Excel debido a un problema de conexión, verifique si el Excel o de alguno de sus componentes se encuentran correctamente instalados", vbCritical
    End Select
    Screen.MousePointer = vbNormal
End Function

Public Property Let SubTitulo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubTitulo = 5
    m_SubTitulo = vData & " "
End Property

Public Property Get SubTitulo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubTitulo
    SubTitulo = m_SubTitulo
End Property

Public Property Let TITULO(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Titulo = 5
    m_Titulo = vData
End Property

Public Property Get TITULO() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Titulo
    TITULO = m_Titulo
End Property


