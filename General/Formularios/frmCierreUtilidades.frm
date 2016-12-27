VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCierreUtilidades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre de Utilidades"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   6330
      Picture         =   "frmCierreUtilidades.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1830
      Width           =   1200
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   735
      Left            =   4830
      Picture         =   "frmCierreUtilidades.frx":0582
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1830
      Width           =   1200
   End
   Begin VB.Frame fraCierre 
      Caption         =   "Proceso de Cierre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7770
      Begin VB.TextBox txtTipoCambio 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1950
         TabIndex        =   3
         Top             =   1230
         Width           =   1420
      End
      Begin VB.ComboBox cboFondo 
         Height          =   315
         Left            =   1950
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   5325
      End
      Begin VB.ComboBox cboFondoSerie 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   5325
      End
      Begin MSComCtl2.DTPicker dtpFechaProceso 
         Height          =   315
         Left            =   1950
         TabIndex        =   4
         Top             =   795
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
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
         Format          =   176160769
         CurrentDate     =   38068
      End
      Begin MSComCtl2.DTPicker dtpFechaPago 
         Height          =   345
         Left            =   5850
         TabIndex        =   5
         Top             =   780
         Width           =   1425
         _ExtentX        =   2514
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
         Format          =   176160769
         CurrentDate     =   38068
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Tipo de Cambio"
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
         Index           =   2
         Left            =   270
         TabIndex        =   10
         Top             =   1245
         Width           =   1455
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fecha Pago Rescates"
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
         Left            =   3600
         TabIndex        =   9
         Top             =   825
         Width           =   2535
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fecha Proceso"
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
         Left            =   270
         TabIndex        =   8
         Top             =   825
         Width           =   1695
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
         Height          =   255
         Index           =   3
         Left            =   270
         TabIndex        =   7
         Top             =   390
         Width           =   1695
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Serie"
         Enabled         =   0   'False
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
         Index           =   4
         Left            =   690
         TabIndex        =   6
         Top             =   630
         Visible         =   0   'False
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmCierreUtilidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()          As String, arrFondoSerie() As String

Dim strCodFondo         As String, strCodMoneda         As String
Dim strSQL              As String, strCodFondoSerie     As String
Dim lngCantSolicitudes  As Long


Private Sub ValidarFechas()

    If EsDiaUtil(dtpFechaProceso.Value) Then
      dtpFechaProceso.Value = dtpFechaProceso
    Else
      dtpFechaProceso.Value = ProximoDiaUtil(dtpFechaProceso.Value)
    End If
    
    If EsDiaUtil(dtpFechaPago.Value) Then
      dtpFechaPago.Value = dtpFechaPago
    Else
      dtpFechaPago.Value = ProximoDiaUtil(dtpFechaPago)
    End If
    
End Sub

Private Sub cboFondo_Click()
    
    Dim adoRegistro As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
'    Cargamos las series del fondo
'    strSQL = "{ call up_ACSelDatosParametro(50,'" & gstrCodAdministradora & "','" & strCodFondo & "') }"
'    CargarControlLista strSQL, cboFondoSerie, arrFondoSerie(), Valor_Caracter
    
'    If cboFondoSerie.ListCount > 0 Then cboFondoSerie.ListIndex = 0

    Set adoRegistro = New ADODB.Recordset

    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            dtpFechaProceso.Value = adoRegistro("FechaCuota")
            dtpFechaPago.Value = DateAdd("d", -1, dtpFechaProceso.Value)

            'Call ValidarFechas
            strCodMoneda = adoRegistro("CodMoneda")

            txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, "02", dtpFechaPago.Value, Codigo_Moneda_Local, strCodMoneda))

            If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaProceso.Value), Codigo_Moneda_Local, strCodMoneda))

            gdatFechaActual = adoRegistro("FechaCuota")
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub


'Private Sub cboFondoSerie_Click()
'
'    Dim adoRegistro As ADODB.Recordset
'
'    strCodFondoSerie = ""
'    If cboFondoSerie.ListIndex < 0 Then Exit Sub
'
'    strCodFondoSerie = Trim(arrFondoSerie(cboFondoSerie.ListIndex))
'
'    Set adoRegistro = New ADODB.Recordset
'
'    With adoComm
'        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
'        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "','" & strCodFondoSerie & "') }"
'        Set adoRegistro = .Execute
'
'        If Not adoRegistro.EOF Then
'            dtpFechaProceso.Value = adoRegistro("FechaCuota")
'            dtpFechaPago.Value = DateAdd("d", -1, dtpFechaProceso.Value)
'
'            'Call ValidarFechas
'            strCodMoneda = adoRegistro("CodMoneda")
'
'            txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, "02", dtpFechaPago.Value, Codigo_Moneda_Local, strCodMoneda))
'
'            If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaProceso.Value), Codigo_Moneda_Local, strCodMoneda))
'
'            gdatFechaActual = adoRegistro("FechaCuota")
'            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
'        End If
'        adoRegistro.Close: Set adoRegistro = Nothing
'    End With
'
'
'End Sub

Private Sub cmdProcesar_Click()
    
  If TodoOK() Then
        If MsgBox("Iniciar proceso de " & CStr(lngCantSolicitudes) & " Solicitudes de Partícipes ?" & _
            vbNewLine & vbNewLine & "Fecha de Proceso :" & Space(1) & CStr(dtpFechaProceso.Value) & vbNewLine & _
            "Tipo de Cambio :" & Space(1) & Trim(txtTipoCambio.Text), vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

        '*** Iniciar Proceso ***
        Dim adoRegistro         As ADODB.Recordset
        Dim adoSolicitud        As ADODB.Recordset
        Dim adoCuentaFondo      As ADODB.Recordset
        Dim adoAuxiliar         As ADODB.Recordset
        Dim strFechaDesde       As String, strFechaHasta    As String
        Dim strCodAnalitica     As String
        Dim datFechaSiguiente   As Date
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción..."
        
        strFechaDesde = Convertyyyymmdd(dtpFechaProceso.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaProceso.Value)
        strFechaHasta = Convertyyyymmdd(datFechaSiguiente)
        
        Set adoRegistro = New ADODB.Recordset
    
        With adoComm
        
            .CommandTimeout = 0

            .CommandText = "SELECT CodFondo,CodAdministradora,DescripFondo,CodMoneda " & _
                           "FROM Fondo WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"

            Set adoRegistro = .Execute
        
                '*** Procesa Solicitud ***
                On Error GoTo Ctrl_Error
                 frmMainMdi.stbMdi.Panels(3).Text = "Procesando Solicitudes Fondo : " & Trim(adoRegistro("DescripFondo")) & "..."
                Me.Refresh
                
                .CommandText = "UPDATE ParticipeSolicitud " & _
                                "SET EstadoSolicitud='" & Estado_Solicitud_Confirmada & "' WHERE " & _
                                "FechaLiquidacion='" & strFechaDesde & "'  AND " & _
                                "EstadoSolicitud='" & Estado_Solicitud_Ingresada & "' AND " & _
                                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                                "TipoSolicitud = '07'"
                .Execute .CommandText

                .CommandText = "{ call up_GNProcDistribucionUtilidadParticipes('" & strCodFondo & "','" & gstrCodAdministradora & "'," & _
                    CDec(txtTipoCambio.Text) & ",'" & strFechaDesde & "','" & Convertyyyymmdd(dtpFechaPago.Value) & "', 'sa','" & Valor_Indicador & "','" & Codigo_Cierre_Definitivo & "') }"
                adoConn.Execute .CommandText
                
            MsgBox Mensaje_Proceso_Exitoso, vbExclamation
            frmMainMdi.stbMdi.Panels(3).Text = "Acción..."
        End With
    End If
    
    Exit Sub
    
Ctrl_Error:
    'adoComm.CommandText = "ROLLBACK TRAN CierreParticipes"
    'adoConn.Execute adoComm.CommandText
    'adoConn.RollbackTrans
    MsgBox adoConn.Errors.Item(0).Description & vbNewLine & vbNewLine & Mensaje_Proceso_NoExitoso, vbCritical
    frmMainMdi.stbMdi.Panels(3).Text = "Acción..."
    Exit Sub
        
End Sub

Private Function TodoOK() As Boolean
        
    Dim adoRegistro         As ADODB.Recordset
    Dim strFechaDesde       As String, strFechaHasta    As String
    Dim datFechaSiguiente   As Date
    
    TodoOK = False
        
    If cboFondo.ListCount = 0 Then
        MsgBox "No existen fondos definidos...", vbCritical, Me.Caption
        Exit Function
    End If
    
    If CDbl(txtTipoCambio.Text) <= 0 And strCodMoneda <> Codigo_Moneda_Local Then
        MsgBox "El tipo de cambio debe ser mayor a cero...", vbCritical, Me.Caption
        Exit Function
    End If
    
    strFechaDesde = Convertyyyymmdd(dtpFechaProceso.Value)
    datFechaSiguiente = DateAdd("d", 1, dtpFechaProceso.Value)
    strFechaHasta = Convertyyyymmdd(datFechaSiguiente)
    
    Set adoRegistro = New ADODB.Recordset
    '*** Confirmar Información ***
    With adoComm
        '*** Validar si existen solicitudes a procesar ***
        lngCantSolicitudes = 0
    
        .CommandText = "SELECT COUNT(NumSolicitud) CantSolicitud FROM ParticipeSolicitud WHERE " & _
            "FechaLiquidacion='" & strFechaDesde & "'  AND " & _
            "EstadoSolicitud='" & Estado_Solicitud_Ingresada & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND TipoSolicitud = '07'"
                    
        If cboFondo.ListIndex > 0 Then
            .CommandText = .CommandText & " AND CodFondo='" & strCodFondo & "'"
        End If
        
        If cboFondoSerie.ListCount > 0 Then
            .CommandText = .CommandText & " AND CodFondoSerie='" & strCodFondoSerie & "'"
        End If
        
        Set adoRegistro = .Execute
      
        If adoRegistro.EOF Then
            MsgBox "No hay Solicitudes pendientes de proceso...", vbInformation
                        
            adoRegistro.Close: Set adoRegistro = Nothing
            Exit Function
        Else
            If IsNull(adoRegistro("CantSolicitud")) Or CLng(adoRegistro("CantSolicitud")) = 0 Then
                MsgBox "No hay Solicitudes pendientes de proceso...", vbInformation
                                
                adoRegistro.Close: Set adoRegistro = Nothing
                Exit Function
            End If
            lngCantSolicitudes = adoRegistro("CantSolicitud")
        End If
        adoRegistro.Close: Set adoRegistro = Nothing

    End With
    
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function
Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call DarFormato
    
    Call ValidarPermisoUsoControl(Trim(gstrLogin), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)

    CentrarForm Me
          
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)

End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
            
End Sub
Private Sub CargarListas()
    
    '*** Fondos ***
'    strSQL = "{ call up_ACSelDatosParametro(29,'" & gstrCodAdministradora & "') }"
'    CargarControlLista strSQL, cboFondo, arrFondo(), Sel_Defecto
    
    '*** Fondos ***
    strSQL = "SELECT CodFondo CODIGO,DescripFondo DESCRIP FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND Estado='01' and CodFondo = '" & gstrCodFondoContable & "' ORDER BY DescripFondo"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
        
End Sub

Private Sub InicializarValores()

    dtpFechaProceso.Value = gdatFechaActual
    dtpFechaPago.Value = DateAdd("d", gintDiasPagoRescate, dtpFechaProceso.Value)
    txtTipoCambio.Text = "0"
    
    Call ValidarFechas
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmCierreParticipes = Nothing
        
End Sub

Private Sub txtTipoCambio_Change()

    Call FormatoCajaTexto(txtTipoCambio, Decimales_TipoCambio)
    
End Sub




