VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRecaudacionFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importacion de Recaudaciones"
   ClientHeight    =   5175
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5955
   Begin TabDlg.SSTab tabFiles 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmRecaudacionFile.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TDBImportaciones"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Importar"
      TabPicture(1)   =   "frmRecaudacionFile.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "btnImportar"
      Tab(1).Control(1)=   "cmdCancelar"
      Tab(1).Control(2)=   "fraDetalle"
      Tab(1).ControlCount=   3
      Begin VB.CommandButton btnImportar 
         Caption         =   "Importar"
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
         Left            =   -72120
         Picture         =   "frmRecaudacionFile.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4020
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
         Left            =   -70800
         Picture         =   "frmRecaudacionFile.frx":0583
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4020
         Width           =   1200
      End
      Begin VB.Frame fraTipoCambio 
         Caption         =   "Criterios de Búsqueda"
         Height          =   1815
         Left            =   -74760
         TabIndex        =   4
         Top             =   480
         Width           =   5280
         Begin VB.ComboBox cboAnio 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1125
            Width           =   1455
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1125
            Width           =   1455
         End
         Begin VB.ComboBox cboTipoTasa 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   360
            Width           =   3735
         End
         Begin VB.ComboBox cboClaseTasa 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   720
            Width           =   3735
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año"
            Height          =   195
            Index           =   0
            Left            =   2880
            TabIndex        =   12
            Top             =   1140
            Width           =   345
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mes"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   11
            Top             =   1140
            Width           =   360
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   10
            Top             =   375
            Width           =   390
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clase"
            Height          =   195
            Index           =   8
            Left            =   360
            TabIndex        =   9
            Top             =   720
            Width           =   480
         End
      End
      Begin VB.Frame fraDetalle 
         Height          =   3405
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   5280
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   3900
            Top             =   1140
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton btnExaminar 
            Caption         =   "&Examinar"
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
            Left            =   3840
            Picture         =   "frmRecaudacionFile.frx":0AE5
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   360
            Width           =   1200
         End
         Begin VB.TextBox txtFile 
            Height          =   375
            Left            =   210
            TabIndex        =   13
            Top             =   630
            Width           =   3525
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Seleccione el archivo a importar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   2
            Top             =   300
            Width           =   2760
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmRecaudacionFile.frx":101B
         Height          =   1875
         Left            =   -74760
         OleObjectBlob   =   "frmRecaudacionFile.frx":1035
         TabIndex        =   3
         Top             =   2415
         Width           =   5280
      End
      Begin TrueOleDBGrid60.TDBGrid TDBImportaciones 
         Bindings        =   "frmRecaudacionFile.frx":39C9
         Height          =   3765
         Left            =   210
         OleObjectBlob   =   "frmRecaudacionFile.frx":39E3
         TabIndex        =   17
         Top             =   510
         Width           =   5280
      End
   End
End
Attribute VB_Name = "frmRecaudacionFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strAnio             As String, arrAnio()            As String
Dim strMes              As String, arrMes()             As String
Dim strDiaInicial       As String, strDiaFinal          As String
Dim strFechaDesde       As String, strFechaHasta        As String
Dim strEstado           As String, strErrMsg            As String
Dim adoRegistroAux  As ADODB.Recordset

Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabFiles.Tab = 0
    strMes = "00"
    
    '*** Ancho por defecto de las columnas de la grilla ***
    TDBImportaciones.Columns(0).Width = TDBImportaciones.Width * 0.01 * 50
    TDBImportaciones.Columns(1).Width = TDBImportaciones.Width * 0.01 * 50
     
    
End Sub

Private Sub btnExaminar_Click()
    
    CommonDialog1.ShowOpen
    txtFile.Text = CommonDialog1.FileName

End Sub

Private Sub btnImportar_Click()
    Call ConfiguraRecordsetAuxiliar
    
    'Open "C:\Documents and Settings\soporte\Escritorio\modulodeparticipes\bws0709.txt" For Input As #1
    Open Trim(txtFile.Text) For Input As #1
    
    Dim adoRegistro As ADODB.Recordset
    Dim intSecuencial As Integer
    Dim dblBookmark As Double
    Dim NroDocumentosCobrados As Integer
    Dim Linea As String, Total As String
    Dim strSQL As String, strSQL1 As String, strSQL2 As String
    Dim strDetalleXML As String
    Dim objDetalleXML As DOMDocument60
    
    NroDocumentosCobrados = 0
    Do Until EOF(1)
    Line Input #1, Linea
        Select Case Mid(Linea, 1, 1)
        Case "H"
            NroDocumentosCobrados = Mid(Linea, 19, 7)
            strSQL1 = "{ call up_PRManRecaudacion("
            strSQL1 = strSQL1 & "'H'"
            strSQL1 = strSQL1 & ",'" & Mid(Linea, 2, 14) & "'"
            strSQL1 = strSQL1 & ",'" & Mid(Linea, 16, 3) & "'"
            strSQL1 = strSQL1 & ",'" & Mid(Linea, 19, 7) & "'"
            strSQL1 = strSQL1 & "," & Mid(Linea, 26, 13) & "." & Mid(Linea, 39, 2)
            strSQL1 = strSQL1 & "," & Mid(Linea, 41, 10) & "." & Mid(Linea, 51, 2)
            strSQL1 = strSQL1 & ",'" & Mid(Linea, 53, 8) & "'"
            
        Case "D"
    
            With adoRegistroAux
               .AddNew
               .Fields("TipoRegistro") = "D"
               .Fields("CodigoUsuario") = Mid(Linea, 19, 15)
               .Fields("NumeroRecibo") = Mid(Linea, 34, 15)
               .Fields("NombreUsuario") = Mid(Linea, 49, 20)
               .Fields("MonedaCobro") = Mid(Linea, 69, 4)
               .Fields("Importe1") = Mid(Linea, 73, 7) & "." & Mid(Linea, 80, 2)
               .Fields("Importe2") = Mid(Linea, 82, 7) & "." & Mid(Linea, 89, 2)
               .Fields("Importe3") = Mid(Linea, 91, 7) & "." & Mid(Linea, 98, 2)
               .Fields("Importe4") = Mid(Linea, 100, 7) & "." & Mid(Linea, 107, 2)
               .Fields("Importe5") = Mid(Linea, 109, 7) & "." & Mid(Linea, 116, 2)
               .Fields("Importe6") = Mid(Linea, 118, 7) & "." & Mid(Linea, 125, 2)
               .Fields("FechaVencimiento") = Mid(Linea, 127, 8)
               .Fields("FechaPago") = Mid(Linea, 135, 8)
               .Fields("TipoPago") = Mid(Linea, 143, 1)
               .Fields("MedioPago") = Mid(Linea, 144, 1)
               .Fields("NumeroOperacion") = Mid(Linea, 145, 13)
               .Fields("ReferenciaCobro") = Mid(Linea, 158, 20)
               .Update
            End With
                dblBookmark = adoRegistroAux.Bookmark
             
            'dblBookmark = adoRegistroAux.Bookmark
            'strSQL = strSQL & ",'D'"
            'strSQL = strSQL & ",'" & Mid(Linea, 19, 15) & "'"
            'strSQL = strSQL & ",'" & Mid(Linea, 34, 15) & "'"
            'strSQL = strSQL & ",'" & Mid(Linea, 49, 20) & "'"
            'strSQL = strSQL & ",'" & Mid(Linea, 69, 4) & "'"
            'strSQL = strSQL & "," & Mid(Linea, 73, 7) & "." & Mid(Linea, 80, 2)
            'strSQL = strSQL & "," & Mid(Linea, 82, 7) & "." & Mid(Linea, 89, 2)
            'strSQL = strSQL & "," & Mid(Linea, 91, 7) & "." & Mid(Linea, 98, 2)
            'strSQL = strSQL & "," & Mid(Linea, 100, 7) & "." & Mid(Linea, 107, 2)
            'strSQL = strSQL & "," & Mid(Linea, 109, 7) & "." & Mid(Linea, 116, 2)
            'strSQL = strSQL & "," & Mid(Linea, 118, 7) & "." & Mid(Linea, 125, 2)
            'strSQL = strSQL & ",'" & Mid(Linea, 127, 8) & "'"
            'strSQL = strSQL & ",'" & Mid(Linea, 135, 8) & "'"
            'strSQL = strSQL & ",'" & Mid(Linea, 143, 1) & "'"
            'strSQL = strSQL & ",'" & Mid(Linea, 144, 1) & "'"
            'strSQL = strSQL & ",'" & Mid(Linea, 145, 13) & "'"
            'strSQL = strSQL & ",'" & Mid(Linea, 158, 20) & "'"
            
        Case "T"
            strSQL2 = strSQL2 & ",'T'"
            strSQL2 = strSQL2 & ",'" & Mid(Linea, 19, 2) & "'"
            strSQL2 = strSQL2 & ",'" & Mid(Linea, 21, 30) & "'"
            strSQL2 = strSQL2 & ",'" & Mid(Linea, 51, 14) & "'"
            strSQL2 = strSQL2 & "," & Mid(Linea, 65, 13) & "." & Mid(Linea, 78, 2)
        End Select
    Loop
    Close #1
    Call XMLADORecordset(objDetalleXML, "Recaudaciones", "Detalle", adoRegistroAux, strErrMsg)
        strDetalleXML = objDetalleXML.xml
    strSQL = strSQL1 & ",'" & strDetalleXML & "'" & strSQL2 + ")}"
    
    If NroDocumentosCobrados > 0 Then
        'MsgBox strSQL, vbCritical
        adoConn.Execute strSQL
        MsgBox "Se realizó la importación sin errores"
    Else
        MsgBox "El file ingresado no tiene data", vbInformation
    End If

End Sub

Private Sub ConfiguraRecordsetAuxiliar()

    Set adoRegistroAux = New ADODB.Recordset

        With adoRegistroAux
           .CursorLocation = adUseClient
           .Fields.Append "TipoRegistro", adVarChar, 1
           .Fields.Append "CodigoUsuario", adVarChar, 15
           .Fields.Append "NumeroRecibo", adVarChar, 15
           .Fields.Append "NombreUsuario", adVarChar, 20
           .Fields.Append "MonedaCobro", adVarChar, 4
           .Fields.Append "Importe1", adVarChar, 10
           .Fields.Append "Importe2", adVarChar, 10
           .Fields.Append "Importe3", adVarChar, 10
           .Fields.Append "Importe4", adVarChar, 10
           .Fields.Append "Importe5", adVarChar, 10
           .Fields.Append "Importe6", adVarChar, 10
           .Fields.Append "FechaVencimiento", adVarChar, 8
           .Fields.Append "FechaPago", adVarChar, 8
           .Fields.Append "TipoPago", adVarChar, 1
           .Fields.Append "MedioPago", adVarChar, 1
           .Fields.Append "NumeroOperacion", adVarChar, 13
           .Fields.Append "ReferenciaCobro", adVarChar, 20
           .LockType = adLockBatchOptimistic
        End With
 
    adoRegistroAux.Open

End Sub
Private Sub CargarReportes()
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Listado de Ingreso de Recaudaciones"

'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Listado de Clientes Naturales"
'
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Text = "Listado de Clientes Jurídicos"
'
End Sub

Private Sub cmdCancelar_Click()

    'cmdOpcion.Visible = True
    With tabFiles
        .TabEnabled(0) = True
        .Tab = 0
    End With
    
End Sub

Private Sub Form_Activate()
Call CargarReportes
End Sub

Private Sub Form_Deactivate()
Call CargarReportes
End Sub

Private Sub Form_Load()

    Call InicializarValores
    'Call CargarListas
    Call CargarReportes
    'Call Buscar
    'Call DarFormato
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
        
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmRecaudacionFile = Nothing
        
End Sub

Public Sub OcultarReportes()

    With frmMainMdi.tlbMdi.Buttons("Reportes")
        .ButtonMenus("Repo1").Visible = False
    End With
    
End Sub

Public Sub SubImprimir(index As Integer)
    Dim strSeleccionRegistro    As String
    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    
    'If tabCliente.Tab = 1 Then Exit Sub
    
    Select Case index
        Case 1
            gstrNameRepo = "Recaudaciones"
                        
            '*** Lista de Clientes por rango de fecha ***
            strSeleccionRegistro = "{RecaudacionDetalle.FechaPago} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal
                
            If gstrSelFrml <> "0" Then
            
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(1)
            ReDim aReportParamFn(4)
            ReDim aReportParamF(4)
                        
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            aReportParamFn(3) = "FechaDel"
            aReportParamFn(4) = "FechaAl"
            
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
            aReportParamF(3) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
            aReportParamF(4) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                        
            aReportParamS(0) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
            aReportParamS(1) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
            End If
            
         End Select


    If gstrSelFrml = "0" Then Exit Sub
    
    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub
