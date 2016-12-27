VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCierreParticipes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre de Partícipes"
   ClientHeight    =   2895
   ClientLeft      =   1635
   ClientTop       =   1785
   ClientWidth     =   10170
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2895
   ScaleWidth      =   10170
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   735
      Left            =   7380
      Picture         =   "frmCierreParticipes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2130
      Width           =   1200
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   8700
      Picture         =   "frmCierreParticipes.frx":0568
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2130
      Width           =   1200
   End
   Begin VB.Frame fraCierre 
      Caption         =   "Proceso de Cierre"
      Height          =   2025
      Left            =   15
      TabIndex        =   0
      Top             =   30
      Width           =   10110
      Begin VB.ComboBox cboFondo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   6795
      End
      Begin VB.TextBox txtTipoCambio 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Top             =   1440
         Width           =   1420
      End
      Begin MSComCtl2.DTPicker dtpFechaProceso 
         Height          =   285
         Left            =   3120
         TabIndex        =   1
         Top             =   740
         Width           =   1420
         _ExtentX        =   2514
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   177733633
         CurrentDate     =   38068
      End
      Begin MSComCtl2.DTPicker dtpFechaPago 
         Height          =   285
         Left            =   3120
         TabIndex        =   2
         Top             =   1080
         Width           =   1420
         _ExtentX        =   2514
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   177733633
         CurrentDate     =   38068
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fondo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   7
         Top             =   380
         Width           =   1695
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fecha de Proceso"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   5
         Top             =   760
         Width           =   1695
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fecha de Pago de Rescates"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   4
         Top             =   1100
         Width           =   2535
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Tipo de Cambio"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   3
         Top             =   1460
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmCierreParticipes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()          As String

Dim strCodFondo         As String, strCodMoneda         As String
Dim strSQL              As String
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
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            dtpFechaProceso.Value = adoRegistro("FechaCuota")
            dtpFechaPago.Value = DateAdd("d", gintDiasPagoRescate, dtpFechaProceso.Value)
            
            'Call ValidarFechas
            strCodMoneda = adoRegistro("CodMoneda")
            
            txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaProceso.Value, strCodMoneda, Codigo_Moneda_Local))
            
            'If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = 2.5886 'CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaProceso.Value), strCodMoneda, Codigo_Moneda_Local))
            
            gdatFechaActual = adoRegistro("FechaCuota")
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub


Private Sub cmdProcesar_Click()
    
   
    '*** Iniciar Proceso ***
    Dim adoRegistro         As ADODB.Recordset
    Dim adoSolicitud        As ADODB.Recordset
    Dim adoCuentaFondo      As ADODB.Recordset
    Dim strFechaProceso     As String

    frmMainMdi.stbMdi.Panels(3).Text = "Acción..."
    
    strFechaProceso = Convertyyyymmdd(dtpFechaProceso.Value)
    
   
    If TodoOK() Then

        If MsgBox("Iniciar proceso de " & CStr(lngCantSolicitudes) & " Solicitudes de Partícipes ?" & _
        vbNewLine & vbNewLine & "Fecha de Proceso :" & Space(1) & CStr(dtpFechaProceso.Value) & vbNewLine & _
        "Tipo de Cambio :" & Space(1) & Trim(txtTipoCambio.Text), vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
                    
        Set adoSolicitud = New ADODB.Recordset
        
        adoComm.CommandText = "SELECT CodBancoDestino, TipoSolicitud FROM ParticipeSolicitud " & _
            "WHERE FechaLiquidacion >= '" & strFechaProceso & "' AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & _
            gstrCodAdministradora & "' AND EstadoSolicitud='" & Estado_Solicitud_Confirmada & "' " & _
            "ORDER BY NumSolicitud ASC"
        
        Set adoSolicitud = adoComm.Execute

        If Not adoSolicitud.EOF Then
            Set adoCuentaFondo = New ADODB.Recordset

            adoComm.CommandText = "SELECT COUNT(*) CantCuenta FROM FondoCuenta WHERE CodFondo='" & strCodFondo & "' AND " & _
                "CodAdministradora='" & gstrCodAdministradora & "' AND CodBanco='" & adoSolicitud("CodBancoDestino") & "' AND " & _
                "TipoOperacion='" & adoSolicitud("TipoSolicitud") & "'"
                Set adoCuentaFondo = adoComm.Execute

            If IsNull(adoCuentaFondo("CantCuenta")) Or adoCuentaFondo("CantCuenta") = 0 And adoSolicitud("TipoSolicitud") <> Codigo_Operacion_Transferencia Then
                MsgBox "Falta información en la definición de la Cta. Corriente " & " para el fondo " & Trim(adoRegistro("DescripFondo")), vbCritical
                adoCuentaFondo.Close: Set adoCuentaFondo = Nothing
                Exit Sub
            End If
            adoCuentaFondo.Close: Set adoCuentaFondo = Nothing

            On Error GoTo Ctrl_Error
                        
            frmMainMdi.stbMdi.Panels(3).Text = "Procesando Solicitudes Fondo : " & Trim(cboFondo.List(cboFondo.ListIndex)) & "..."
            Me.Refresh
    
            adoComm.CommandText = "{ call up_GNProcCierreParticipes('" & strCodFondo & "','" & gstrCodAdministradora & "'," & _
                CDec(txtTipoCambio.Text) & ",'" & strFechaProceso & "','" & Convertyyyymmdd(dtpFechaPago.Value) & "','" & _
                gstrLogin & "') }"
            adoConn.Execute adoComm.CommandText

        End If
        adoSolicitud.Close: Set adoSolicitud = Nothing

        MsgBox Mensaje_Proceso_Exitoso, vbExclamation
        frmMainMdi.stbMdi.Panels(3).Text = "Acción..."
    
        Exit Sub
     
    End If
    
Ctrl_Error:
    
    If err.Number <> 0 Then
        MsgBox err.Number & " " & err.Description, vbCritical + vbOKOnly, Me.Caption
        Me.MousePointer = vbDefault
    End If
        
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
    
    If CDbl(txtTipoCambio.Text) <= 1 And strCodMoneda <> Codigo_Moneda_Local Then
        MsgBox "Ingrese el Tipo de Cambio para el dia...", vbCritical, Me.Caption
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
            "(FechaSolicitud>='" & strFechaDesde & "' AND FechaSolicitud<'" & strFechaHasta & "') AND " & _
            "EstadoSolicitud='" & Estado_Solicitud_Confirmada & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
            "CodFondo='" & strCodFondo & "'"
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
            
End Sub
Private Sub CargarListas()
    
'    '*** Fondos ***
'    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
'    CargarControlLista strSQL, cboFondo, arrFondo(), Sel_Defecto

   '*** Fondos ***
    strSQL = "SELECT CodFondo CODIGO,DescripFondo DESCRIP FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND Estado='01' and CodFondo = '" & gstrCodFondoContable & "' ORDER BY DescripFondo"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    cboFondo.ListIndex = 0
    
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


