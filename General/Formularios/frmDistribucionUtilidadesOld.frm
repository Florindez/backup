VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmDistribucionUtilidadesOld 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distribución de Utilidades"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   13995
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   11520
      TabIndex        =   0
      Top             =   8280
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "&Cargar"
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
      Left            =   1080
      Picture         =   "frmDistribucionUtilidadesOld.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8280
      Width           =   1200
   End
   Begin VB.Frame fraCarga 
      Height          =   8175
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   13935
      Begin VB.Frame fraParametros 
         Caption         =   "Parametros de Distribución"
         Height          =   765
         Left            =   360
         TabIndex        =   12
         Top             =   2310
         Width           =   13185
         Begin VB.CommandButton cmdCalcularValorCuota 
            Caption         =   "Calcular Valor Cuota"
            Height          =   375
            Left            =   9660
            TabIndex        =   13
            Top             =   240
            Width           =   1995
         End
         Begin TAMControls.TAMTextBox txtValorCuotaActualizado 
            Height          =   315
            Left            =   5850
            TabIndex        =   14
            Top             =   300
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmDistribucionUtilidadesOld.frx":054B
            Text            =   "0.00000000"
            Decimales       =   8
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   8
            MaximoValor     =   999999999
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor Cuota Actualizado"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   3900
            TabIndex        =   15
            Top             =   330
            Width           =   1695
         End
      End
      Begin VB.Frame frmCarga 
         Caption         =   "Datos para Distribución"
         Height          =   1725
         Left            =   360
         TabIndex        =   2
         Top             =   390
         Width           =   13215
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   600
            Width           =   8565
         End
         Begin MSComCtl2.DTPicker dtpFechaRegistro 
            Height          =   345
            Left            =   1650
            TabIndex        =   4
            Top             =   1050
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   196870145
            CurrentDate     =   38790
         End
         Begin TAMControls.TAMTextBox txtValorCuota 
            Height          =   285
            Left            =   8430
            TabIndex        =   11
            Top             =   1050
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   503
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmDistribucionUtilidadesOld.frx":0567
            Text            =   "0.00000000"
            Decimales       =   8
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   8
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtValorTipoCambio 
            Height          =   315
            Left            =   4650
            TabIndex        =   16
            Top             =   1050
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmDistribucionUtilidadesOld.frx":0583
            Text            =   "0.00000000"
            Decimales       =   8
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   8
            MaximoValor     =   999999999
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   3570
            TabIndex        =   17
            Top             =   1080
            Width           =   885
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor Cuota Actual"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   6990
            TabIndex        =   10
            Top             =   1080
            Width           =   1320
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fondo"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   6
            Top             =   660
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   5
            Top             =   1110
            Width           =   450
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Height          =   4095
         Left            =   360
         OleObjectBlob   =   "frmDistribucionUtilidadesOld.frx":059F
         TabIndex        =   7
         Top             =   3300
         Width           =   13185
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsultaCO 
         Bindings        =   "frmDistribucionUtilidadesOld.frx":5DB3
         Height          =   645
         Left            =   450
         OleObjectBlob   =   "frmDistribucionUtilidadesOld.frx":5DCD
         TabIndex        =   8
         Top             =   8790
         Width           =   13215
      End
   End
End
Attribute VB_Name = "frmDistribucionUtilidadesOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Ficha de Mantenimiento de Precios/Tir"
Option Explicit

Dim strEstado                   As String, strSQL               As String
Dim blnSelec                    As Boolean
Dim dblPrecio                   As Double, dblPreProm           As Double
Dim dblTir                      As Double
Dim arrFondo()                  As String  'para el cbofondo.
Dim strCodFondo                 As String
Dim adoRegistro                 As ADODB.Recordset
Dim adoRegistroAux              As ADODB.Recordset
Dim adoClone                    As ADODB.Recordset
Dim adoClone1                   As ADODB.Recordset
Dim conexion                    As ADODB.Connection
Dim tstObservaciones            As New TrueOleDBGrid60.Style
Dim adoField                    As ADODB.Field
Dim adoConsulta                 As ADODB.Recordset
Dim indSortAsc                  As Boolean, indSortDesc         As Boolean
Dim strFechaCorte               As String
Dim strCodMoneda                As String
Dim dblValorCuotaNominal        As Double



'Public Sub SubImprimir(Index As Integer)
'
'    Dim frmReporte              As frmVisorReporte
'    Dim aReportParamS(), aReportParamF(), aReportParamFn()
'    Dim strFechaDesde           As String, strFechaHasta        As String
'    Dim strSeleccionRegistro    As String
'
'    'If tabPrecio.Tab = 1 Then Exit Sub
'
'    Select Case Index
'
'        Case 1
'
'            strSeleccionRegistro = "{InstrumentoPrecioTir.FechaCotizacion} IN 'Fch1' TO 'Fch2'"
'            gstrSelFrml = strSeleccionRegistro
'            frmRangoFecha.Show vbModal
'
'            If gstrSelFrml <> "0" Then
'
'            '/* Para validar al cerrar el Rango de Fechas */
'            If Mid(gstrSelFrml, 44, 4) = "Fch1" Then
'                Exit Sub
'            End If
'
'            Set frmReporte = New frmVisorReporte
'
'            ReDim aReportParamS(4)
'            ReDim aReportParamFn(5)
'            ReDim aReportParamF(5)
'
'            aReportParamFn(0) = "Usuario"
'            aReportParamFn(1) = "FechaDesde"
'            aReportParamFn(2) = "FechaHasta"
'            aReportParamFn(3) = "Hora"
'            aReportParamFn(4) = "NombreEmpresa"
'            aReportParamFn(5) = "Fondo"
'
'            aReportParamF(0) = gstrLogin
'            aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
'            aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
'            aReportParamF(3) = Format(Time(), "hh:mm:ss")
'            aReportParamF(4) = gstrNombreEmpresa & Space(1)
'            aReportParamF(5) = cboFondo.List(cboFondo.ListIndex)
'
'            aReportParamS(0) = Trim(arrFondo(cboFondo.ListIndex)) 'Trim(arrFondo(cboFondo.ListIndex)) 'Mid(strNemotecnicoVal, 1, 3)
'            aReportParamS(1) = "000" 'Mid(strNemotecnicoVal, 5, 3) 'ponemos la administradora x defecto
'            aReportParamS(2) = gstrCodAdministradora 'Mid(strNemotecnicoVal, 5, 3) 'ponemos la administradora x defecto
'            aReportParamS(3) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
'            aReportParamS(4) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))
'
'            gstrNameRepo = "ParticipeDistribucionUtilidad"
'
'    End Select
'
'    gstrSelFrml = Valor_Caracter
'    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"
'
'    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())
'
'    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
'    frmReporte.Show vbModal
'
'    Set frmReporte = Nothing
'
'    Screen.MousePointer = vbNormal
'
'End Sub

Public Sub Adicionar()

End Sub

Public Sub Cancelar()

    'cmdOpcion.Visible = True
'    With tabPrecio
'        .TabEnabled(0) = True
'        .Tab = 0
'    End With
'    Call Buscar
    
End Sub

Public Sub Eliminar()

End Sub

Public Sub Grabar()

'    Dim strFechaInicio  As String, strFechaFin  As String
'    Dim intRegistro     As Integer
'
'    If strEstado = Reg_Consulta Then Exit Sub
'
'    If strEstado = Reg_Edicion Then
'        'If TodoOK() Then
'            strFechaInicio = Convertyyyymmdd(dtpFechaRegistro.Value)
'            strFechaFin = Convertyyyymmdd(DateAdd("d", 1, dtpFechaRegistro.Value))
'
'            With adoComm
'                .CommandText = "UPDATE InstrumentoPrecioTir SET " & _
'                    "PrecioCierre=" & CDec(txtPrecioCierre.Text) & "," & _
'                    "TirCierre=" & CDec(txtTirCierre.Text) & "," & _
'                    "UsuarioEdicion='" & gstrLogin & "' " & _
'                    "WHERE CodTitulo='" & Trim(tdgConsulta.Columns(2).Value) & "' AND " & _
'                    "(FechaCotizacion>='" & strFechaInicio & "' AND FechaCotizacion<'" & strFechaFin & "')"
'                adoConn.Execute .CommandText, intRegistro
'
'                If intRegistro = 0 Then
'                    .CommandText = "UPDATE InstrumentoPrecioTir SET " & _
'                        "IndUltimoPrecio=''," & _
'                        "UsuarioEdicion='" & gstrLogin & "' " & _
'                        "WHERE CodTitulo='" & Trim(tdgConsulta.Columns(2).Value) & "' AND " & _
'                        "FechaCotizacion = (SELECT MAX(FechaCotizacion) FROM InstrumentoPrecioTir " & _
'                        "                   WHERE CodTitulo='" & Trim(tdgConsulta.Columns(2).Value) & "' AND " & _
'                        "                   FechaCotizacion < '" & strFechaInicio & "')"
'                    adoConn.Execute .CommandText
'
'                    .CommandText = "INSERT INTO InstrumentoPrecioTir " & _
'                     "(CodTitulo, FechaCotizacion, Nemotecnico," & _
'                     "CodFile, CodDetalleFile, CodAnalitica," & _
'                     "PrecioCierre, TirCierre, PrecioPromedio," & _
'                     "IndUltimoPrecio, UsuarioEdicion) " & _
'                     " VALUES ('" & _
'                     Trim(tdgConsulta.Columns(2).Value) & "','" & strFechaInicio & "','" & _
'                     Trim(lblNemotecnico.Caption) & "','" & strCodFile & "','" & _
'                     Trim(tdgConsulta.Columns(7).Value) & "','" & Trim(tdgConsulta.Columns(6).Value) & "'," & _
'                     CDec(txtPrecioCierre.Text) & "," & CDec(txtTirCierre.Text) & "," & _
'                     CDec(txtPrecioCierre.Text) & ",'X','" & gstrLogin & "')"
'                    adoConn.Execute .CommandText
'                End If
'
'            End With
'
'            Me.MousePointer = vbDefault
'
'            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
'
'            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
'
'            cmdOpcion.Visible = True
'            With tabPrecio
'                .TabEnabled(0) = True
'                .Tab = 0
'            End With
'
'            Call Buscar
'        'End If
'    End If

End Sub

Public Sub Imprimir()

End Sub

Public Sub Salir()

    Unload Me
    
End Sub






Private Sub cboFondo_Click()

    Dim adoConsulta As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    With adoComm
        '{*** Fecha Vigente, Moneda ***
        .CommandText = "{ call up_ACSelDatosParametro (23,'" & strCodFondo & "','" & gstrCodAdministradora & "','000') }"
        Set adoConsulta = .Execute
        
        If Not adoConsulta.EOF Then
            gdatFechaActual = CVDate(adoConsulta("FechaCuota"))
            gstrFechaActual = Convertyyyymmdd(gdatFechaActual)
            strFechaCorte = Convertyyyymmdd(DateAdd("d", -1, gdatFechaActual))
            dtpFechaRegistro.Value = gdatFechaActual
                          
            strCodMoneda = adoConsulta("CodMoneda")
                          
            txtValorCuota.Text = CStr(adoConsulta("ValorCuotaInicial"))
            dblValorCuotaNominal = adoConsulta("ValorCuotaNominal")
            txtValorCuotaActualizado.Text = CStr(dblValorCuotaNominal)
            
            txtValorTipoCambio.Text = CStr(2.586) 'CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaRegistro.Value, Codigo_Moneda_Local, strCodMoneda))
            
            If CDbl(txtValorTipoCambio.Text) = 0 Then txtValorTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaRegistro.Value), Codigo_Moneda_Local, strCodMoneda))
            
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            
            gstrPeriodoActual = Format(Year(gdatFechaActual), "0000")
            gstrMesActual = Format(Month(gdatFechaActual), "00")
            
            'ACTUALIZA PARAMETROS GLOBALES POR FONDO
            If Not CargarParametrosGlobales(strCodFondo) Then Exit Sub

        End If
        adoConsulta.Close: Set adoConsulta = Nothing
    
    End With
  
    Call Buscar
    

End Sub





'Private Sub cmdCargar_Click()
'
'    Call CargarPrecios_xInterfaz
'
'End Sub

'Private Sub CargarPrecios_xInterfaz()
'
'Dim objExcel As Excel.Application
'Dim xLibro As Excel.Workbook
'Dim Col As Integer, fila As Integer
'Dim precio As Double
'Dim fechaCarga As String
'Dim strNemotecnico, strCodTitulo, strMsgError As String
'Dim blnOpenExcel As Boolean
'
'Dim intColNemonicoElex, intColPrecioElex, intColPrecioAntElex, intColFechaAntElex  As Integer
'Dim lngFilaIniElex  As Long
'
'Dim intColNemonicoBloom, intColPrecioBloom As Integer
'Dim lngFilaIniBloom  As Long
'
'blnOpenExcel = False
'
'On Error GoTo CtrlError
'
'If MsgBox("Desea realizar la carga de precios de mercado de los instrumentos ?.", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
'   Me.Refresh: Exit Sub
'End If
'
'If Dir(Trim(txtArchivo.Text)) = "" Then
'    MsgBox "No se encontró el archivo con los precios de carga. Revise si está indicado correctamente. ", vbCritical
'    Exit Sub
'End If
'
'Screen.MousePointer = vbHourglass
'
'Set objExcel = New Excel.Application
'Set xLibro = objExcel.Workbooks.Open(Trim(txtArchivo.Text))   ' ("c:\precio.xls")
'objExcel.Visible = False
'blnOpenExcel = True
'
''-----------------------------------
''Valores pre-establecidos
''-----------------------------------
'
'' Seteos Elex
'intColNemonicoElex = 2
'intColPrecioElex = 7
'intColPrecioAntElex = 3
'intColFechaAntElex = 4
'lngFilaIniElex = 3
'
'' Seteos Bloomberg
'intColNemonicoBloom = 20
'intColPrecioBloom = 21
'lngFilaIniBloom = 2
'
''-----------------------------------
''Carga de precios de valores locales
''-----------------------------------
'
'With xLibro
'    With .Sheets(1)
'
'        fila = lngFilaIniElex
'        strNemotecnico = Trim(.Cells(fila, intColNemonicoElex))
'
'        'For fila = lngFilaInicial To 300
'        Do While Trim(strNemotecnico) <> ""
'
'            'Obtener el precio de la acción.
'            If Trim(.Cells(fila, intColPrecioElex)) = "------" Then
'                'Si no se encuentra el último valor se tomará el valor anterior.
'                If Trim(.Cells(fila, intColPrecioAntElex)) = "------" Then
'                    GoTo siga
'                Else
'                    precio = CDbl(.Cells(fila, intColPrecioAntElex))
'                    If Trim(.Cells(fila, intColFechaAntElex)) <> "-----" Then
'                    fechaCarga = Convertyyyymmdd(Trim(.Cells(fila, intColFechaAntElex)))
'
'                    End If
'                End If
'            Else
'                precio = CDbl(.Cells(fila, intColPrecioElex))
'                fechaCarga = gstrFechaActual
'            End If
'
'            With adoComm
'                .CommandText = "{ call up_IVActPrecioValores ('" & strNemotecnico & "'," & precio & ",'" & _
'                                  fechaCarga & "','" & gstrLogin & "' ) }"
'                adoConn.Execute .CommandText
'
'            End With
'
'siga:
'
'            fila = fila + 1
'            strNemotecnico = Trim(.Cells(fila, intColNemonicoElex))
'
'        Loop
'
'    End With
'
'End With
'
''---------------------------------------------------------------
''Ahora con el mismo archivo se cargan los precios del extranjero
''---------------------------------------------------------------
'
'With xLibro
'    With .Sheets(1)
'
'        fila = lngFilaIniBloom
'        strNemotecnico = Trim(.Cells(fila, intColNemonicoBloom))
'
'
'        Do While Trim(strNemotecnico) <> ""
'
'            'Obtener el precio de la acción.
'            If IsNumeric(.Cells(fila, intColPrecioBloom)) = True Then
'               precio = CDbl((.Cells(fila, intColPrecioBloom)))
'            Else
'                GoTo siga2
'            End If
'
'            With adoComm
'                .CommandText = "{ call up_IVActPrecioValores ('" & strNemotecnico & "'," & precio & ",'" & _
'                                  gstrFechaActual & "','" & gstrLogin & "' ) }"


Private Sub cmdCalcularValorCuota_Click()

    Dim adoRegistro                     As New ADODB.Recordset
    Dim dblValorCapital                 As Double
    Dim dblValorUtilidadRepartida       As Double
    Dim dblValorUtilidadReinvertida     As Double
    Dim dblValorUtilidadNoDistribuida   As Double
    Dim dblValorTotalResultados         As Double
    Dim dblValorCuotaActualizado        As Double
    Dim dblValorNuevoCapital            As Double
    Dim dblValorPatrimonioInicial       As Double
    Dim dblBookmark                     As Double
    
    
    If adoRegistroAux.RecordCount = 0 Then
        MsgBox "No existen registros para validar!", vbExclamation ''& cmdCommand.CommandText
        Exit Sub
    End If
    
    dblValorCapital = 0
    dblValorUtilidadRepartida = 0
    dblValorUtilidadReinvertida = 0
    dblValorUtilidadNoDistribuida = 0
    dblValorTotalResultados = 0
    dblValorPatrimonioInicial = 0
    
'    adoComm.CommandText = "SELECT SUM(SaldoInicialContable + SaldoParcialContable) AS ValorResultados FROM PartidaContableSaldos " & _
'                          "WHERE " & _
'                          "CodFondo = '" & strCodFondo & "' AND " & _
'                          "CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
'                          "FechaSaldo = '" & gstrFechaActual & "' AND " & _
'                          "CodCuenta LIKE '59%' AND " & _
'                          "CodMonedaContable = '" & Codigo_Moneda_Local & "'"

    adoComm.CommandText = "SELECT dbo.uf_ACObtenerResultadosEjercicio('" & strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaCorte & "') AS 'ValorResultados'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        dblValorTotalResultados = adoRegistro.Fields("ValorResultados") * -1
    End If
    adoRegistro.Close: Set adoRegistro = Nothing
    
    dblBookmark = adoRegistroAux.Bookmark

    adoRegistroAux.MoveFirst
    
    Do While Not adoRegistroAux.EOF

        dblValorCapital = dblValorCapital + Round((adoRegistroAux.Fields("CantCuotas") * adoRegistroAux.Fields("ValorCuota")), 2)
        
        If Trim(adoRegistroAux.Fields("IndReinvierte")) = Valor_Indicador Then
            dblValorUtilidadReinvertida = dblValorUtilidadReinvertida + adoRegistroAux.Fields("ValorUtilidad")
        End If
        
        If Trim(adoRegistroAux.Fields("IndReinvierte")) = Valor_Caracter Then
            dblValorUtilidadRepartida = dblValorUtilidadRepartida + adoRegistroAux.Fields("ValorUtilidad")
        End If

        adoRegistroAux.MoveNext
    Loop
            
    adoRegistroAux.Bookmark = dblBookmark
        
'    If strCodMoneda <> Codigo_Moneda_Local Then
'        dblValorCapital = Round(dblValorCapital * CDbl(txtValorTipoCambio.Text), 2)
'        dblValorUtilidadRepartida = Round(dblValorUtilidadRepartida * CDbl(txtValorTipoCambio.Text), 2)
'        dblValorUtilidadReinvertida = Round(dblValorUtilidadReinvertida * CDbl(txtValorTipoCambio.Text), 2)
'    End If
        
    'SE ASUME QUE SE DISTRIBUYE TODA LA UTILIDAD
    dblValorUtilidadNoDistribuida = 0 'dblValorTotalResultados - (dblValorUtilidadRepartida + dblValorUtilidadReinvertida)
        
    dblValorNuevoCapital = dblValorCapital - dblValorUtilidadRepartida + dblValorUtilidadReinvertida
        
    dblValorCuotaActualizado = Round(txtValorCuota.Value - ((dblValorUtilidadRepartida + dblValorUtilidadReinvertida) / dblValorTotalResultados) * (txtValorCuota.Value - dblValorCuotaNominal), 3)
    
    'Round(((dblValorNuevoCapital + dblValorUtilidadNoDistribuida) / dblValorPatrimonioInicial) * dblValorCuotaNominal, 5)

    txtValorCuotaActualizado.Text = CStr(dblValorCuotaActualizado)


End Sub

'                adoConn.Execute .CommandText
'
'            End With
'
'
'siga2:
'
'            fila = fila + 1
'            strNemotecnico = Trim(.Cells(fila, intColNemonicoBloom))
'
'        Loop
'
'    End With
'End With
'
''Cerrando el archivo excel
'xLibro.Close True
'Set xLibro = Nothing
'Set objExcel = Nothing
'
'Screen.MousePointer = vbNormal
'MsgBox "Finalizó exitosamente la carga de precios de mercado.", vbExclamation
'
'Call Buscar
'
'Exit Sub
'
'CtrlError:
'    If blnOpenExcel = True Then
'        xLibro.Close True
'        Set xLibro = Nothing
'        Set objExcel = Nothing
'    End If
'
'    Me.MousePointer = vbDefault
'    strMsgError = "Error " & Str(err.Number) & vbNewLine
'    strMsgError = strMsgError & err.Description
'    MsgBox strMsgError, vbCritical, "Error"
'
'
'End Sub
Private Sub cmdCargar_Click()
        
    Dim strFechaCarga As String

    Dim objSolicitudDistribucionUtilidadXML    As DOMDocument60
    Dim strSolicitudDistribucionUtilidadXML    As String
    Dim strMsgError                     As String

    If Not TodoOK() Then Exit Sub
    
    If MsgBox("Desea Proceder con la Carga de Operaciones del dia " & dtpFechaRegistro.Value & " ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    strFechaCarga = Convertyyyymmdd(dtpFechaRegistro.Value)

    Call XMLADORecordset(objSolicitudDistribucionUtilidadXML, "ParticipeSolicitud", "Solicitud", adoRegistroAux, strMsgError)
    strSolicitudDistribucionUtilidadXML = objSolicitudDistribucionUtilidadXML.xml

    'txtValorCuotaActualizado.Text = CStr(100#)

    With adoComm
        .CommandText = "{ call up_GNGenSolicitudDistribucionUtilidad ('" & strCodFondo & "','" & gstrCodAdministradora & "'," & _
                        "'" & strFechaCarga & "'," & CDbl(txtValorCuota.Value) & "," & CDbl(txtValorCuotaActualizado.Value) & ",'" & strSolicitudDistribucionUtilidadXML & "') }"
        adoConn.Execute .CommandText
    End With

    cmdCargar.Enabled = False
    
    MsgBox Mensaje_Carga_Exitosa, vbExclamation
    
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
    Exit Sub

CtrlError:
    Me.MousePointer = vbDefault

    MsgBox "Error al Leer El Archivo, Verifique que la estructura sea la correcta. "


End Sub

Private Function TodoOK()

    TodoOK = False

    If adoRegistroAux.RecordCount = 0 Then
        MsgBox "No existen registros para cargar!", vbExclamation ''& cmdCommand.CommandText
        Exit Function
    End If

    TodoOK = True


End Function

Private Sub Form_Activate()

    frmMainMdi.stbMdi.Panels(3).Text = Me.Caption
    Call CargarReportes
   
End Sub

Private Sub Form_Deactivate()

    Call OcultarReportes
    
End Sub

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
Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Distribucion de Utilidades"

End Sub
Private Sub CargarListas()
    
    Dim strSQL  As String
    
    '*** Fondo ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
        
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    'adoComm.CommandText = " SELECT ValorCuotaInicialReal FROM FondoSerieValorCuota WHERE IndAbierto='X' AND CodFondo= '" & gstrCodAdministradora & "'"
      '  Set adoRegistro = adoComm.Execute
        'Label1.Caption = adoRegistro("ValorCuotaInicialReal")
        
End Sub

Private Sub InicializarValores()
    
    strEstado = Reg_Defecto

    dtpFechaRegistro.Value = gdatFechaActual
    
    Set tstObservaciones = tdgConsulta.Styles.Add("Observaciones")
    tstObservaciones.Font.Bold = True
    
    ' Before modifying the grid's properties, make sure the grid is
 
    ' Create an additional split.
    Dim S As TrueOleDBGrid60.Split
    Set S = tdgConsulta.Splits.Add(0)
 
    ' Hide all columns in the leftmost split, Splits(0), except for columns 0 and 1.
    Dim c As TrueOleDBGrid60.Column
    Dim Cols As TrueOleDBGrid60.Columns
    Set Cols = tdgConsulta.Splits(0).Columns
    For Each c In Cols
        c.Visible = False
    Next c
    Cols(0).Visible = True
    Cols(1).Visible = True
 
    ' Configure Splits(0) to display exactly two columns, and disable resizing.
    With tdgConsulta.Splits(0)
        .SizeMode = dbgNumberOfColumns
        .Size = 2
        .AllowSizing = False
    End With
 
    ' Usually, if you fix columns 0 and 1 from scrolling  in a split, you will
    ' want to make them invisible in other splits.
    Set Cols = tdgConsulta.Splits(1).Columns
    Cols(0).Visible = False
    Cols(1).Visible = False
 
    ' Turn off the record selectors in Split 1.
    tdgConsulta.Splits(1).RecordSelectors = False
    
    Set cmdSalir.FormularioActivo = Me
    'Set cmdAccion.FormularioActivo = Me
    'Set cmdOpcion.FormularioActivo = Me
                
End Sub

Private Sub ConfiguraRecordsetAuxiliar()

    Set adoRegistroAux = New ADODB.Recordset
    
    With adoRegistroAux
       .CursorLocation = adUseClient
       .Fields.Append "CodParticipe", adChar, 20
       .Fields.Append "DescripParticipe", adVarChar, 200
       .Fields.Append "NumCertificado", adVarChar, 10
       .Fields.Append "CodMoneda", adChar, 2
       .Fields.Append "CodSigno", adChar, 3
       .Fields.Append "FechaIngreso", adDate, 10
       .Fields.Append "CantCuotas", adDecimal, 19
       .Fields.Append "ValorCuota", adDecimal, 19
       .Fields.Append "ValorUtilidadBruta", adDecimal, 19
       .Fields.Append "TasaImptoRenta", adDecimal, 19
       .Fields.Append "ValorImptoRenta", adDecimal, 19
       .Fields.Append "ValorUtilidadNeta", adDecimal, 19
       .Fields.Append "IndReinvierte", adChar, 1
       .Fields.Append "ClaseParticipe", adChar, 2
       .LockType = adLockBatchOptimistic
    End With

    With adoRegistroAux.Fields.Item("CantCuotas")
        .Precision = 19
        .NumericScale = 8
    End With

    With adoRegistroAux.Fields.Item("ValorCuota")
        .Precision = 19
        .NumericScale = 8
    End With

    With adoRegistroAux.Fields.Item("ValorUtilidadBruta")
        .Precision = 19
        .NumericScale = 2
    End With
    
    With adoRegistroAux.Fields.Item("ValorUtilidadNeta")
        .Precision = 19
        .NumericScale = 2
    End With
    
    With adoRegistroAux.Fields.Item("TasaImptoRenta")
        .Precision = 19
        .NumericScale = 8
    End With
    
    With adoRegistroAux.Fields.Item("ValorImptoRenta")
        .Precision = 19
        .NumericScale = 2
    End With
    
'    With adoRegistroAux
'       .CursorLocation = adUseClient
'       .Fields.Append "CodParticipe", adChar, 20
'       .Fields.Append "DescripParticipe", adVarChar, 100
'       .Fields.Append "NumCertificado", adVarChar, 20
'       .Fields.Append "CodMoneda", adChar, 2
'       .Fields.Append "CodSigno", adChar, 3
'       .Fields.Append "FechaIngreso", adDate, 10
'       .Fields.Append "CantCuotas", adDecimal, 19
'       .Fields.Append "ValorCuota", adDecimal, 19
'       .Fields.Append "ValorUtilidad", adDecimal, 19
'       .Fields.Append "IndReinvierte", adChar, 1
'       .LockType = adLockBatchOptimistic
'    End With
'
'    With adoRegistroAux.Fields.Item("CantCuotas")
'        .Precision = 19
'        .NumericScale = 5
'    End With
'
'    With adoRegistroAux.Fields.Item("ValorCuota")
'        .Precision = 19
'        .NumericScale = 8
'    End With
'
'    With adoRegistroAux.Fields.Item("ValorUtilidad")
'        .Precision = 19
'        .NumericScale = 2
'    End With
    
    adoRegistroAux.Open

End Sub

Public Sub Buscar()
            
    Dim strSQL As String
    
    Set adoRegistro = New ADODB.Recordset
    
    Call ConfiguraRecordsetAuxiliar
    
    strEstado = Reg_Defecto

    strSQL = "{ call up_GNLstParticipes('" & strCodFondo & "','" & gstrCodAdministradora & "','" & Convertyyyymmdd(dtpFechaRegistro.Value) & "'," & txtValorCuota.Value & ",100)}"

    With adoRegistro
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open strSQL

        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                adoRegistroAux.AddNew
                For Each adoField In adoRegistroAux.Fields
                    adoRegistroAux.Fields(adoField.Name) = adoRegistro.Fields(adoField.Name)
                Next
                adoRegistroAux.Update
                adoRegistro.MoveNext
                'adoMovimiento.MoveNext
            Loop
            adoRegistroAux.MoveFirst
        End If
        
    End With
    
    tdgConsulta.DataSource = adoRegistroAux
    
    tdgConsulta.Refresh
    
    If adoRegistroAux.RecordCount > 0 Then strEstado = Reg_Consulta
    
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
                
        Case vNew
            Call Adicionar
        Case vQuery
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

Public Sub Modificar()

'    If strEstado = Reg_Consulta Then
'        strEstado = Reg_Edicion
'        LlenarFormulario strEstado
'        cmdOpcion.Visible = False
'        With tabPrecio
'            .TabEnabled(0) = False
'            .Tab = 1
'        End With
'    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

'    Dim adoRegistro     As ADODB.Recordset
'
'    Select Case strModo
'        Case Reg_Edicion
'            lblDescripInstrumento.Caption = Trim(cboTipoInstrumento.Text)
'            lblFechaRegistro.Caption = CStr(dtpFechaRegistro.Value)
'            lblNemotecnico.Caption = CStr(tdgConsulta.Columns(1))
'            lblPrecioAnterior.Caption = CStr(tdgConsulta.Columns(3))
'            lblTirAnterior.Caption = CStr(tdgConsulta.Columns(4))
'            If Trim(tdgConsulta.Columns(0).Value) = Valor_Caracter Then
'                txtPrecioCierre.Text = "0"
'                txtTirCierre.Text = "0"
'            Else
'                If CVDate(tdgConsulta.Columns(0).Value) < dtpFechaRegistro.Value Then
'                    txtPrecioCierre.Text = "0"
'                    txtTirCierre.Text = "0"
'                Else
'                    txtPrecioCierre.Text = CStr(tdgConsulta.Columns(3))
'                    txtTirCierre.Text = CStr(tdgConsulta.Columns(4))
'                End If
'            End If
'
'            Set adoRegistro = New ADODB.Recordset
'
'            adoComm.CommandText = "SELECT IndPrecio,IndTir FROM InversionFile WHERE CodFile='" & strCodFile & "'"
'            Set adoRegistro = adoComm.Execute
'
'            If Not adoRegistro.EOF Then
'                txtPrecioCierre.Enabled = True
'                If Trim(adoRegistro("IndPrecio")) = Valor_Caracter Then txtPrecioCierre.Enabled = False
'                txtTirCierre.Enabled = True
'                If Trim(adoRegistro("IndTir")) = Valor_Caracter Then txtTirCierre.Enabled = False
'            End If
'            adoRegistro.Close: Set adoRegistro = Nothing
'
'    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmDistribucionUtilidadesOld = Nothing
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub







Private Sub tdgConsulta_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid60.StyleDisp)

'    Dim strTipoValidacion As String
'    Dim adoRegistro As New ADODB.Recordset
'
'    adoClone.Bookmark = Bookmark
'
'    Select Case Col
'        Case 2, 3, 4, 9 'TipoOperacion,Nemotecnico,Moneda,Broker
'
'            If Col = 2 Then 'TipoOperacion
'                strTipoValidacion = "01"
'            ElseIf Col = 3 Then 'Nemotecnico
'                strTipoValidacion = "02"
'            ElseIf Col = 4 Then 'Moneda
'                strTipoValidacion = "03"
'            ElseIf Col = 9 Then 'Broker
'                strTipoValidacion = "04"
'            End If
'
'            adoComm.CommandText = "SELECT dbo.uf_IVValidaDatoCargaOperacion('" & strTipoValidacion & "','" & tdgConsulta.Columns(Col).CellText(Bookmark) & "') AS 'ValidaDato'"
'            Set adoRegistro = adoComm.Execute
'
'            If Not adoRegistro.EOF Then
'                If Not adoRegistro("ValidaDato") Then
'                    'tdgConsulta.Columns("IndRegistroOK").CellText(Bookmark) = Valor_Caracter
'                    adoClone("IndRegistroOK").Value = Valor_Caracter
'                    CellStyle.Font.Bold = True
'                End If
'            End If
'            adoRegistro.Close: Set adoRegistro = Nothing
'
'    End Select


End Sub

Private Sub tdgConsulta_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
'
'    adoClone.Bookmark = Bookmark
'
'    If Trim(adoClone("IndRegistroOK").Value) = Valor_Caracter Then
'        RowStyle.ForeColor = vbWhite
'        RowStyle.BackColor = vbRed
'    End If

End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If tdgConsulta.Columns(ColIndex).DataField = "CantCuotas" Then
        Call DarFormatoValor(Value, Decimales_CantCuota)
    End If
    
    If tdgConsulta.Columns(ColIndex).DataField = "ValorCuota" Then
        Call DarFormatoValor(Value, Decimales_ValorCuota)
    End If
    
    If tdgConsulta.Columns(ColIndex).DataField = "ValorUtilidad" Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
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

    Call OrdenarDBGrid(ColIndex, adoRegistroAux, tdgConsulta)
    
    numColindex = ColIndex
    
    '****
    strPrevColumTDB = strColNameTDB
    '***
    
End Sub


Private Sub tdgConsulta_UnboundColumnFetch(Bookmark As Variant, ByVal Col As Integer, Value As Variant)

'    adoClone.Bookmark = Bookmark
'
'    If Col = 7 Then
'        Value = adoClone("ComisionSAB") + adoClone("ComisionBVL") + _
'                adoClone("ComisionConasev") + adoClone("ComisionCavali") + _
'                adoClone("FondoLiquidacion") + adoClone("FondoGarantia") + adoClone("IGV")
'    End If
    

End Sub
