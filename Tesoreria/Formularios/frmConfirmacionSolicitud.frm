VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmConfirmacionSolicitud 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Confirmación de Solicitudes"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   13425
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   11280
      TabIndex        =   12
      Top             =   6960
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin VB.Frame fraConfirmacion 
      Caption         =   "Confirmación"
      Height          =   6885
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   13395
      Begin VB.CommandButton cmdVerificar 
         Caption         =   "Verificar"
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
         Left            =   9960
         Picture         =   "frmConfirmacionSolicitud.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1680
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.ComboBox cboEstado 
         Height          =   315
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1860
         Width           =   2565
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Height          =   3525
         Left            =   360
         OleObjectBlob   =   "frmConfirmacionSolicitud.frx":05DB
         TabIndex        =   13
         Top             =   2460
         Width           =   12555
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   7830
         TabIndex        =   11
         Top             =   6150
         Width           =   2000
      End
      Begin VB.TextBox txtTotalSeleccionado 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   3750
         TabIndex        =   10
         Top             =   6150
         Width           =   2000
      End
      Begin VB.ComboBox cboTipoSolicitud 
         Height          =   315
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1005
         Width           =   2595
      End
      Begin VB.ComboBox cboFondo 
         Height          =   315
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   6285
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Confirmar"
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
         Left            =   11280
         Picture         =   "frmConfirmacionSolicitud.frx":7377
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Confirmar/Desconfirmar Solicitudes"
         Top             =   1680
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker dtpFechaConsulta 
         Height          =   315
         Left            =   2130
         TabIndex        =   3
         Top             =   1425
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   49938433
         CurrentDate     =   38068
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Estado"
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
         Height          =   285
         Index           =   5
         Left            =   360
         TabIndex        =   14
         Top             =   1890
         Width           =   975
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Tipo Solicitud"
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
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Monto Total"
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
         Height          =   285
         Index           =   4
         Left            =   6390
         TabIndex        =   7
         Top             =   6180
         Width           =   1455
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Monto Seleccionado"
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
         Height          =   285
         Index           =   3
         Left            =   1710
         TabIndex        =   6
         Top             =   6180
         Width           =   2055
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Ordenes Al"
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
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   1470
         Width           =   975
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fondo"
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
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   615
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmConfirmacionSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()      As String, arrTipoSolicitud()       As String
Dim arrEstado()     As String

Dim strCodFondo     As String, strCodTipoSolicitud      As String
Dim strCodMoneda    As String, strCodEstado             As String
Dim strEstado       As String, strSQL                   As String

Dim adoConsulta     As ADODB.Recordset
Dim adoRegistroAux  As ADODB.Recordset
Dim adoVerificacion As ADODB.Recordset
Dim indSortAsc      As Boolean, indSortDesc             As Boolean


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
Public Sub Abrir()

End Sub


Public Sub Adicionar()

End Sub

Public Sub Anterior()

End Sub

Public Sub Ayuda()

End Sub

Public Sub Buscar()

    Dim datFechaSiguiente   As Date
    Dim strFechaDesde       As String, strFechaHasta    As String
    Dim strSQL              As String
    
    strFechaDesde = Convertyyyymmdd(dtpFechaConsulta.Value)
'    datFechaSiguiente = DateAdd("d", 1, dtpFechaConsulta.Value)
'    strFechaHasta = Convertyyyymmdd(datFechaSiguiente)
    strFechaHasta = Convertyyyymmdd(dtpFechaConsulta.Value)
   
    Me.MousePointer = vbHourglass

            
    strSQL = "{ call up_TEListarParticipeSolicitud ('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
            strFechaDesde & "','" & strFechaHasta & "','" & strCodEstado & "','" & strCodTipoSolicitud & "') }"
    'MsgBox strSQL, vbCritical
    Set adoConsulta = New ADODB.Recordset
            
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    tdgConsulta.DataSource = adoConsulta

    If adoConsulta.RecordCount > 0 Then
        
        Dim adoRegistro As ADODB.Recordset
        
        strEstado = Reg_Consulta
        txtTotalSeleccionado.Text = "0"
        
        
        strSQL = "{ call up_SumParticipeSolicitud ('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
        strFechaDesde & "','" & strFechaHasta & "','" & strCodEstado & "','" & strCodTipoSolicitud & "') }"
        
        Set adoRegistro = New ADODB.Recordset
        
        
        With adoRegistro
            .ActiveConnection = gstrConnectConsulta
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open strSQL
        End With
        
        
'        With adoComm
''            .CommandText = "SELECT COALESCE(SUM(MontoNetoSolicitud),0) MontoTotal FROM ParticipeSolicitud " & _
''                "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
''                "(FechaSolicitud >='" & strFechaDesde & "' AND FechaSolicitud <'" & strFechaHasta & "') AND " & _
''                "EstadoSolicitud='" & strCodEstado & "' AND TipoSolicitud='" & strCodTipoSolicitud & "'"
''            'MsgBox .CommandText, vbCritical
'            Set adoRegistro = .Execute
            
            If Not adoRegistro.EOF Then
                txtTotal.Text = CStr(adoRegistro("MontoNetoSolicitud"))
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
'        End With
    Else
        txtTotal.Text = "0": txtTotalSeleccionado.Text = "0"
    End If
    
    Me.MousePointer = vbDefault

                        
End Sub


Public Sub Cancelar()

End Sub

Public Sub Eliminar()

End Sub

Public Sub Exportar()

End Sub

Public Sub Grabar()

End Sub


Public Sub Importar()

End Sub

Public Sub Imprimir()

End Sub

Public Sub Modificar()

End Sub

Public Sub Primero()

End Sub

Public Sub Refrescar()

End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub Seguridad()

End Sub

Public Sub Siguiente()

End Sub

Public Sub SubImprimir(index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim datFechaSiguiente       As Date
        
    Select Case index
        Case 1
            gstrNameRepo = "ConfirmacionSolicitud"
                        
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(6)
            ReDim aReportParamFn(3)
            ReDim aReportParamF(3)
                        
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            aReportParamFn(3) = "Fondo"
                        
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
            aReportParamF(3) = Trim(cboFondo.Text)
                        
            strFechaDesde = Convertyyyymmdd(dtpFechaConsulta.Value)
            datFechaSiguiente = DateAdd("d", 1, dtpFechaConsulta.Value)
            strFechaHasta = Convertyyyymmdd(datFechaSiguiente)
            
            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora
            aReportParamS(2) = strFechaDesde
            aReportParamS(3) = strFechaHasta
            aReportParamS(4) = strCodEstado
            aReportParamS(5) = strCodTipoSolicitud
            aReportParamS(6) = strCodMoneda
            
    End Select

    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub

Public Sub Ultimo()

End Sub


Private Sub cboEstado_Click()

    strCodEstado = Valor_Caracter
    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strCodEstado = Trim(arrEstado(cboEstado.ListIndex))
    
    If strCodEstado = Estado_Solicitud_Confirmada Then
        cmdProcesar.Caption = "&Desconfirmar"
    Else
        cmdProcesar.Caption = "&Confirmar"
    End If
    
    Call Buscar
    
End Sub


Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset
    Dim strSigno    As String
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            dtpFechaConsulta.Value = adoRegistro("FechaCuota")
            strCodMoneda = adoRegistro("CodMoneda")
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    strSigno = ObtenerSignoMoneda(strCodMoneda)
    tdgConsulta.Columns("MontoNetoSolicitud").Caption = "Monto" & Space(1) & strSigno
    lblDescrip(3).Caption = "Monto Seleccionado" & Space(1) & strSigno
    lblDescrip(4).Caption = "Monto Total" & Space(1) & strSigno
    
    Call Buscar
    
End Sub


Private Sub cboTipoSolicitud_Click()

    strCodTipoSolicitud = Valor_Caracter
    If cboTipoSolicitud.ListIndex < 0 Then Exit Sub
    
    strCodTipoSolicitud = Trim(arrTipoSolicitud(cboTipoSolicitud.ListIndex))
    
    Call Buscar
    
End Sub


Private Sub cmdProcesar_Click()
    
    Dim strFechaProceso             As String
    Dim intRegistro                 As Integer, intContador         As Integer
    Dim strParticipeSolicitudXML    As String
    Dim objParticipeSolicitudXML    As DOMDocument60
    Dim strMsgError                 As String
    
    If adoConsulta.RecordCount = 0 Then Exit Sub
    
    strFechaProceso = Convertyyyymmdd(dtpFechaConsulta.Value) & Space(1) & Format(Time, "hh:mm")

    intContador = tdgConsulta.SelBookmarks.Count - 1
    
    If intContador < 0 Then
        MsgBox "No se ha seleccionado ningún registro", vbCritical, Me.Caption
        Exit Sub
    End If
        
    Call ConfiguraRecordsetAuxiliar
    
    For intRegistro = 0 To intContador
               
        adoConsulta.MoveFirst
        
        adoConsulta.Move CLng(tdgConsulta.SelBookmarks(intRegistro) - 1), 0
                        
        tdgConsulta.Refresh
                        
        adoRegistroAux.AddNew
        
        adoRegistroAux.Fields("CodFondo") = strCodFondo
        adoRegistroAux.Fields("CodAdministradora") = gstrCodAdministradora
        adoRegistroAux.Fields("NumSolicitud") = tdgConsulta.Columns("NumSolicitud")
    
    Next
    
    Call XMLADORecordset(objParticipeSolicitudXML, "ParticipeSolicitud", "Solicitud", adoRegistroAux, strMsgError)
    strParticipeSolicitudXML = objParticipeSolicitudXML.xml

    adoComm.CommandText = "{ call up_TEProcAbonoParticipe('" & _
    strCodFondo & "','" & gstrCodAdministradora & "','" & _
    strFechaProceso & "','" & strCodEstado & "','" & _
    strParticipeSolicitudXML & "') }"
    'MsgBox adoComm.CommandText, vbCritical
    adoComm.Execute adoComm.CommandText
    
    If strCodEstado = Estado_Solicitud_Ingresada Then
        MsgBox Mensaje_Confirmacion_Exitoso, vbExclamation, gstrNombreEmpresa
    Else
        MsgBox Mensaje_Desconfirmacion_Exitoso, vbExclamation, gstrNombreEmpresa
    End If
    
    Call Buscar
    
End Sub


Private Sub cmdVerificar_Click()
    Dim strSQL              As String
    
    Set adoVerificacion = New ADODB.Recordset
    strSQL = "{ call up_ACSelDatos (43) }"
    With adoVerificacion
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    'adoComm.CommandText = "{ call up_ACSelDatos (43) }"
    'Set adoVerificacion = adoComm.Execute
    
   If adoVerificacion.RecordCount > 0 Then
        With adoVerificacion
            .MoveFirst
                
            Do While Not .EOF
           
                 With tdgConsulta
                    .MoveFirst
                        
                    Do While Not .EOF
                         If tdgConsulta.Columns(12).Text = adoVerificacion("NumContrato") Then
                            tdgConsulta.SelBookmarks.Add tdgConsulta.Bookmark
                            Exit Do
                         End If
                        .MoveNext
                    Loop
                End With
                .MoveNext
            Loop
        End With
        MsgBox "Se encontraron " & adoVerificacion.RecordCount & " coincidencias en los pagos", vbApplicationModal
    Else
        MsgBox "No se encontraron coincidencias en los pagos", vbApplicationModal
    End If
        
   
End Sub

Private Sub dtpFechaConsulta_Change()

    Call Buscar

End Sub

Private Sub Form_Activate()

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
          
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
          
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
    For intCont = 0 To (fraConfirmacion.Count - 1)
        Call FormatoMarco(fraConfirmacion(intCont))
    Next
            
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Confirmación de Solicitudes"
    
End Sub

Private Sub CargarListas()
        
    Dim intRegistro  As Integer
        
    '*** Tipo Solicitud ***
    strSQL = "SELECT CodTipoSolicitud CODIGO,DescripTipoSolicitud DESCRIP FROM TipoSolicitud WHERE CodCorto IN ('S','R','T') ORDER BY DescripTipoSolicitud DESC"
    CargarControlLista strSQL, cboTipoSolicitud, arrTipoSolicitud(), Valor_Caracter
    
    If cboTipoSolicitud.ListCount > 0 Then cboTipoSolicitud.ListIndex = 0
    intRegistro = ObtenerItemLista(arrTipoSolicitud(), Codigo_Operacion_Suscripcion)
    If intRegistro >= 0 Then cboTipoSolicitud.ListIndex = intRegistro
    
    '*** Estado ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='ESTSOL' AND ValorParametro='X' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Valor_Caracter
    
    If cboEstado.ListCount > 0 Then cboEstado.ListIndex = 0
    intRegistro = ObtenerItemLista(arrEstado(), Estado_Solicitud_Ingresada)
    If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
        
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
        
        
End Sub


Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(5).Width = tdgConsulta.Width * 0.01 * 35
    
    Set cmdSalir.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmConfirmacionSolicitud = Nothing
    
End Sub


Private Sub tdgConsulta_DblClick()
    
    On Error GoTo CtrlError '/**/ HMC Habilitamos la rutina de Errores.

    Dim adoRegistro         As ADODB.Recordset
    Dim intAccion           As Integer, lngNumError         As Long
    Dim strPagosParciales   As String

    Set adoRegistro = New ADODB.Recordset
    '*** Consultamos si estan habilitados los pagos parciales ***
    adoComm.CommandText = "SELECT CodAfirmacion FROM Fondo " & _
        "WHERE CodFondo ='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then                                     'HMC
        strPagosParciales = Trim(adoRegistro("CodAfirmacion"))
        '*** El Fondo NO acepta pagos parciales ***
        If strPagosParciales = Codigo_Respuesta_No Then Exit Sub
    End If                                                          'HMC
    adoRegistro.Close: Set adoRegistro = Nothing                    'HMC

    Dim intRegistro     As Integer
    
    gstrFormulario = Me.Name
    frmPagoCuotaSuscripcion.Show
    intRegistro = ObtenerItemLista(garrFondo(), strCodFondo)
    frmPagoCuotaSuscripcion.cboFondo.ListIndex = intRegistro
    intRegistro = ObtenerItemLista(garrParticipe(), tdgConsulta.Columns("CodParticipe").Value)
    frmPagoCuotaSuscripcion.cboParticipe.ListIndex = intRegistro
    frmPagoCuotaSuscripcion.cmdOpcion.Button(0).Enabled = False
    frmPagoCuotaSuscripcion.cmdOpcion.Button(1).Enabled = False
    frmPagoCuotaSuscripcion.Buscar
    
    Exit Sub
    
CtrlError:
    Me.MousePointer = vbDefault
    intAccion = ControlErrores
    Select Case intAccion
        Case 0: Resume
        Case 1: Resume Next
        Case 2: Exit Sub
        Case Else
            lngNumError = err.Number
            err.Raise Number:=lngNumError
            err.Clear
    End Select
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

Static numColindex As Integer

    If ColIndex = 6 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
    
   
    
End Sub


Private Sub tdgConsulta_SelChange(Cancel As Integer)

    Dim dblMonto    As Double, dblMontoAcumulado    As Double
    Dim intRegistro As Integer, intContador         As Integer
    
    intContador = tdgConsulta.SelBookmarks.Count - 1
     
    txtTotalSeleccionado.Text = "0"
    For intRegistro = 0 To intContador
       tdgConsulta.Row = tdgConsulta.SelBookmarks(intRegistro) - 1
       dblMonto = CDbl(tdgConsulta.Columns(6))
       
       dblMontoAcumulado = dblMontoAcumulado + dblMonto
    Next

    txtTotalSeleccionado.Text = CStr(dblMontoAcumulado)
    
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

Private Sub txtTotal_Change()

    Call FormatoCajaTexto(txtTotal, Decimales_Monto)
    
End Sub


Private Sub txtTotalSeleccionado_Change()

    Call FormatoCajaTexto(txtTotalSeleccionado, Decimales_Monto)
    
End Sub
Private Sub ConfiguraRecordsetAuxiliar()

    Set adoRegistroAux = New ADODB.Recordset

    With adoRegistroAux
       .CursorLocation = adUseClient
       .Fields.Append "CodFondo", adVarChar, 3
       .Fields.Append "CodAdministradora", adVarChar, 3
       .Fields.Append "NumSolicitud", adVarChar, 10
       .LockType = adLockBatchOptimistic
    End With

    adoRegistroAux.Open

End Sub


