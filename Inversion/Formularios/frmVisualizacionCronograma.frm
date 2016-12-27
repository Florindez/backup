VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Begin VB.Form frmVisorCronograma 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visualización de cronograma"
   ClientHeight    =   7290
   ClientLeft      =   5760
   ClientTop       =   3180
   ClientWidth     =   13125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   13125
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMoneda 
      Enabled         =   0   'False
      Height          =   285
      Left            =   11400
      TabIndex        =   19
      Top             =   1170
      Width           =   1455
   End
   Begin VB.TextBox txtTasaEfectiva 
      Enabled         =   0   'False
      Height          =   285
      Left            =   10320
      TabIndex        =   18
      Top             =   810
      Width           =   2535
   End
   Begin TAMControls.TAMTextBoxMultiline txtDetalleCupon 
      Height          =   5175
      Left            =   9960
      TabIndex        =   16
      Top             =   1560
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   9128
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
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      ForeColor       =   -2147483640
      Locked          =   -1  'True
      Container       =   "frmVisualizacionCronograma.frx":0000
      MultiLine       =   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpVencimiento 
      Height          =   285
      Left            =   11550
      TabIndex        =   14
      Top             =   450
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   213254145
      CurrentDate     =   40416
   End
   Begin MSComCtl2.DTPicker dtpEmision 
      Height          =   285
      Left            =   11550
      TabIndex        =   13
      Top             =   90
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   213254145
      CurrentDate     =   40416
   End
   Begin VB.TextBox txtPeriodo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Top             =   1170
      Width           =   1935
   End
   Begin VB.TextBox txtValorNominal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Top             =   810
      Width           =   1935
   End
   Begin VB.TextBox txtNombreValor 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Top             =   450
      Width           =   4215
   End
   Begin VB.TextBox txtCodigoUnico 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Top             =   90
      Width           =   1935
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11760
      TabIndex        =   1
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6840
      Width           =   1095
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dgvCronograma 
      Height          =   5175
      Left            =   120
      OleObjectBlob   =   "frmVisualizacionCronograma.frx":001C
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1560
      Width           =   9735
   End
   Begin VB.Label lblMoneda 
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
      Height          =   255
      Index           =   1
      Left            =   8640
      TabIndex        =   17
      Top             =   1230
      Width           =   1335
   End
   Begin VB.Label lblTasaEfectiva 
      Caption         =   "Tasa efectiva"
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
      Index           =   0
      Left            =   8640
      TabIndex        =   8
      Top             =   870
      Width           =   1515
   End
   Begin VB.Label lblFechaVencimiento 
      Caption         =   "Fecha de Vencimiento"
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
      Left            =   8640
      TabIndex        =   7
      Top             =   480
      Width           =   1995
   End
   Begin VB.Label lblFechaEmision 
      Caption         =   "Fecha de Emisión"
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
      Left            =   8640
      TabIndex        =   6
      Top             =   120
      Width           =   1785
   End
   Begin VB.Label lblPeriodo 
      Caption         =   "Periodo"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label lblValorNominal 
      Caption         =   "Monto Aprobado"
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
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblNombreValor 
      Caption         =   "Nombre del valor"
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
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblCodigoUnico 
      Caption         =   "Código Unico"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1245
   End
End
Attribute VB_Name = "frmVisorCronograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Source           As Integer
Private aSource         As Integer
Public desembMultiple   As Boolean
Public codigoUnico      As String
Public strNumSolicitud  As String
Private codOperacion    As String
Public nombreValor      As String
Public valorNominal     As Double
Public periodo          As String
Public fechaEmision     As Date
Public fechavencimiento As Date
Public tasaEfectiva     As Double
Public tipoTasa         As String
Public frecTasa         As String
Dim gdblTasaIgvTemp     As Double
Dim strCodFondo         As String

Private Sub cmdExportar_Click()
    Call SubImprimir
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub setFondo(cf As String)
    strCodFondo = cf
End Sub

Private Sub dgvCronograma_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, _
                                       ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
   
    If desembMultiple = True Then
        strSQL = "{ call up_IVSelInstrumentoInversionCalendario('" & strCodFondo & "','" & codOperacion & "',0,'" & dgvCronograma.Columns.Item(0).Value & "'," & Source & "," & gdblTasaIgvTemp & ",'" & txtMoneda.Text & "','" & txtPeriodo & "','" & txtTasaEfectiva & "') }"
        
        Dim adoDetalleCupon As ADODB.Recordset
        
        adoComm.CommandText = strSQL
        Set adoDetalleCupon = adoComm.Execute
        
        strdetallecupon = "Cupón Nº " & dgvCronograma.Columns.Item(0).Value
        
        If Not adoDetalleCupon.EOF Then
            adoDetalleCupon.MoveFirst
        End If
        
        While Not adoDetalleCupon.EOF
            'JJCC: Agregado el IGV al pequeño reporte situado a la derecha del grid del form
            'para el caso de desembolsos múltiples
            strdetallecupon = strdetallecupon & Chr(13) & Chr(10) & "Desembolso Nº " & Str(adoDetalleCupon.Fields.Item(0).Value) & Chr(13) & Chr(10) & "Saldo deudor: " & Str(adoDetalleCupon.Fields.Item(1).Value) & Chr(13) & Chr(10) & "Amortización : " & Str(adoDetalleCupon.Fields.Item(2).Value) & Chr(13) & Chr(10) & "Interés          : " & Str(adoDetalleCupon.Fields.Item(3).Value) & Chr(13) & Chr(10) & "IGV               : " & Str(adoDetalleCupon.Fields.Item(4).Value) & Chr(13) & Chr(10) & "============================="
            adoDetalleCupon.MoveNext
        Wend

        txtDetalleCupon.Text = strdetallecupon
    End If
    
End Sub

Private Sub Form_Load()

    'JJCC: Condición para verificar si se incluye o no en el cálculo.
    If csdf = "001" Then
        If cti_igv Then
            gdblTasaIgvTemp = gdblTasaIgv
        Else
            gdblTasaIgvTemp = 0
        End If

    Else
        gdblTasaIgvTemp = 0
    End If
    
    aSource = Source
    
    ConfGrid dgvCronograma, True, False, False, False
    
    If (Source < 2) Then
        codOperacion = codigoUnico
    Else
        codOperacion = strNumSolicitud
    End If
        
    Dim adoInfoInstrumento As ADODB.Recordset
    adoComm.CommandText = "SELECT II.DescripTitulo, II.FechaEmision, II.FechaVencimiento, ISOL.MontoAprobado, " & _
                        " II.TasaInteres, IICF.TipoTasa, IICF.PeriodoTasa, IICF.PeriodoCupon, ISOL.CodMoneda, IICF.IndDesembolsosMultiples " & _
                        " from InstrumentoInversion II, " & _
                        " InstrumentoInversionCondicionesFinancieras IICF, " & "InversionSolicitud ISOL where II.CodTitulo = '" & codigoUnico & _
                        "' and IICF.CodTitulo = '" & codigoUnico & "' and ISOL.CodTitulo = '" & codigoUnico & "'"
                        
    Set adoInfoInstrumento = adoComm.Execute

    If Not adoInfoInstrumento.EOF Then
        txtCodigoUnico.Text = codigoUnico
        txtNombreValor.Text = adoInfoInstrumento.Fields.Item("DescripTitulo")
        txtValorNominal.Text = adoInfoInstrumento.Fields.Item("MontoAprobado")
        
        If adoInfoInstrumento.Fields.Item("IndDesembolsosMultiples") = 1 Then
            desembMultiple = True
        Else
            desembMultiple = False
        End If
        
        fechaEmision = adoInfoInstrumento.Fields.Item("FechaEmision")
        fechavencimiento = adoInfoInstrumento.Fields.Item("FechaVencimiento")
        dtpEmision.Value = fechaEmision
        dtpVencimiento.Value = fechavencimiento
        
        'JJCC: Corregido case que no mostraba el periodo en el textbox.
        'Antes estaba tomando los índices (de 0 a 7), cuando ya se había
        'asignado el valor en el orden que están.
        Select Case CInt(adoInfoInstrumento.Fields.Item("PeriodoCupon"))

            Case 1: periodo = "Anual"

            Case 2: periodo = "Semestral"

            Case 3: periodo = "Trimestral"

            Case 4: periodo = "Bimestral"

            Case 5: periodo = "Mensual"

            Case 6: periodo = "Quincenal"

            Case 7: periodo = "Diario"

            Case 8: periodo = "Personalizado"
        End Select

        'Fin JJCC
        
        txtPeriodo.Text = periodo
        
        'JJCC: Concatenación de los datos de la Tasa
        tasaEfectiva = CDbl(adoInfoInstrumento.Fields.Item("TasaInteres"))
        
        Select Case CInt(adoInfoInstrumento.Fields.Item("TipoTasa"))

            Case 1: tipoTasa = "Efectiva"

            Case 2: tipoTasa = "Nominal"
        End Select
        
        Select Case CInt(adoInfoInstrumento.Fields.Item("PeriodoTasa"))

            Case 1: frecTasa = "Anual"

            Case 2: frecTasa = "Semestral"

            Case 3: frecTasa = "Trimestral"

            Case 4: frecTasa = "Bimestral"

            Case 5: frecTasa = "Mensual"

            Case 6: frecTasa = "Quincenal"

            Case 7: frecTasa = "Diaria"
        End Select
        
        'Se llena el textbox con el resultado de la concatenación
        txtTasaEfectiva.Text = tasaEfectiva & "% " & tipoTasa & " " & frecTasa
        'Fin JJCC.

        Select Case CInt(adoInfoInstrumento.Fields.Item("CodMoneda"))

            Case 1: txtMoneda.Text = "Soles"

            Case 2: txtMoneda.Text = "Dólares"
        End Select
                
        'Source = 4
        'codigoUnico = "001"
        
        'JJCC: Cambio en el Stored Procedure.
        strSQL = "{ call up_IVSelInstrumentoInversionCalendario('" & strCodFondo & "','" & codOperacion & "',0,''," & Source & "," & gdblTasaIgvTemp & ",'" & txtMoneda.Text & "','" & txtPeriodo & "','" & txtTasaEfectiva & "') }"
        'Fin JJCC.
        
        With dgvCronograma
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = gstrConnectConsulta
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.Active = False
            .Dataset.ADODataset.CommandText = strSQL
            .Dataset.DisableControls
            .Dataset.Active = True
            .KeyField = "NumCupon"
        End With

        Source = 1

    Else
        MsgBox "El calendario de pagos aún no ha sido generado.", vbInformation
        Unload frmVisorCronograma
    End If

End Sub

Public Sub SubImprimir()

    Dim frmReporte As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde        As String, strFechaHasta        As String
    Dim strSeleccionRegistro As String

    gstrNameRepo = "InversionSolicitudCalendario"
    
    Set frmReporte = New frmVisorReporte

    ReDim aReportParamS(8)
    ReDim aReportParamFn(1)
    ReDim aReportParamF(1)
                
    aReportParamFn(0) = "Usuario"
                
    aReportParamF(0) = gstrLogin
    aReportParamS(0) = strCodFondo
    aReportParamS(1) = codOperacion
    aReportParamS(2) = 0
    aReportParamS(3) = ""
    aReportParamS(4) = aSource
    aReportParamS(5) = gdblTasaIgvTemp
    aReportParamS(6) = txtMoneda.Text
    aReportParamS(7) = txtPeriodo.Text
    aReportParamS(8) = txtTasaEfectiva.Text
   
    If gstrSelFrml = "0" Then Exit Sub
    
    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub

