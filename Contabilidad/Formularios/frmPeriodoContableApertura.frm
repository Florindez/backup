VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPeriodoContableApertura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proceso de Apertura de Periodo Contable"
   ClientHeight    =   3135
   ClientLeft      =   2970
   ClientTop       =   1110
   ClientWidth     =   8745
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
   Icon            =   "frmPeriodoContableApertura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3135
   ScaleWidth      =   8745
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   7110
      Picture         =   "frmPeriodoContableApertura.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2310
      Width           =   1200
   End
   Begin VB.CommandButton cmd_cierre 
      Caption         =   "&Procesar"
      Height          =   735
      Left            =   5700
      Picture         =   "frmPeriodoContableApertura.frx":09C4
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2310
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   8655
      Begin VB.CheckBox chkSimulacion 
         Caption         =   "Simular el Proceso de Apertura de Periodo Contable"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         ToolTipText     =   "Marcar para proceso de simulación"
         Top             =   1680
         Width           =   4965
      End
      Begin VB.TextBox txtTipoCambio 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   2250
         TabIndex        =   6
         Top             =   1230
         Width           =   1455
      End
      Begin VB.ComboBox cboFondo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2265
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   450
         Width           =   5955
      End
      Begin MSComCtl2.DTPicker dtpFechaApertura 
         Height          =   315
         Left            =   2265
         TabIndex        =   2
         Top             =   840
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   113377281
         CurrentDate     =   38068
      End
      Begin VB.Label lblFondo 
         Caption         =   "Fondo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblFondo 
         Caption         =   "Fecha de Apertura"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   885
         Width           =   1695
      End
      Begin VB.Label lblFondo 
         Caption         =   "Tipo de Cambio"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   1260
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmPeriodoContableApertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strAnoCie               As String
Dim strMesCie               As String
Dim strDiacie               As String
Dim strHorCie               As String
Dim arrFondo()              As String
Dim strCodFondo             As String, strFechaApertura               As String
Dim strFechaAnterior        As String, strFechaSiguiente            As String
Dim strFechaAnteAnterior    As String, strFechaSubSiguiente         As String
Dim strCodMoneda            As String, strCodModulo                 As String
Dim strSQL                  As String




Private Sub cmd_cierre_Click()

    Dim adoresult As New Recordset, strMen As String, res As Integer
    Dim datFecha As Date
    
    'On Error GoTo cmd_cierre_error
   
    If TodoOK() Then
        '*** Inicializar Variables de Trabajo ***
    '    pnlMsg.Caption = "Inicio del proceso..."
        Me.MousePointer = vbHourglass 'Reloj de Arena
        
    '    strFeccie = FmtFec(Dat_FecCie.Value, "win", "yyyymmdd", res)
        strAnoCie = Mid(strFechaApertura, 1, 4): strMesCie = Mid(strFechaApertura, 5, 2): strDiacie = Mid(strFechaApertura, 7, 2)
        strHorCie = Format(Now, "hh:mm")
        
        datFecha = dtpFechaApertura.Value
            
        '*** Pedir Confirmacion de Datos ***
        strMen = "Para el proceso de APERTURA DE PERIODO confirme lo siguiente : " & Chr$(10)
        strMen = strMen & " Fondo >> " & cboFondo.List(cboFondo.ListIndex) & Chr$(10)
        strMen = strMen & " Fecha >> " & dtpFechaApertura.Value & Chr$(10)
        strMen = strMen & " Tipo de cambio >> " & txtTipoCambio.Text & Chr$(10)
        strMen = strMen & "¿Seguro de continuar?"
        If MsgBox(strMen, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then  'No desea continuar
            GoTo cmd_cierre_fin
        End If
      
        frmMainMdi.stbMdi.Panels(3).Text = "Realizando Apertura de Nuevo Periodo..."
        
        adoConn.CommandTimeout = 10000

        adoComm.CommandText = "{ call up_GNProcAperturaPeriodoContable('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & datFecha & "'," & _
            CDbl(txtTipoCambio.Text) & ",'" & gstrLogin & "','"
        
        If chkSimulacion.Value Then
           adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Simulacion & "') }"
        Else
           adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Definitivo & "') }"
        End If
        adoConn.Execute adoComm.CommandText
        
        Sleep 0&
              
        Me.MousePointer = vbDefault 'Normal
              
        If chkSimulacion.Value Then
            MsgBox "Proceso de Simulación de Apertura de Periodo Contable culminado exitosamente", vbInformation
        Else
            MsgBox "Proceso de Apertura de Periodo Contable culminado exitosamente", vbInformation
        End If
        
    End If

cmd_cierre_fin:
    Me.MousePointer = vbDefault
'    pnlMsg.Caption = ""
    Exit Sub
    
cmd_cierre_error:
    strMen = "Error   : " & Str$(err) & Chr$(10)
    strMen = strMen & "Detalle : " & Error$ & Chr$(10)
    strMen = strMen & "SQL     : " & adoComm.CommandText
    MsgBox strMen, vbCritical
    Resume cmd_cierre_fin

End Sub

Private Sub cmd_salir_Click()

    Unload Me

End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub cboFondo_Click()

'    Dim adoAux As New Recordset, res As Integer
'
'    If cboFondo.ListIndex >= 0 Then
'        With adoComm
'            .CommandText = "SELECT FCH_FINA FROM FMPRDCON"
'            .CommandText = .CommandText & " WHERE COD_FOND='" & AcodFon(cboFondo.ListIndex) & "'"
'            .CommandText = .CommandText & " AND MES_CONT='99' AND FLG_CIER='X' ORDER BY FCH_FINA ASC"
'            Set adoAux = .Execute
'            If Not adoAux.EOF Then
''                Dat_FecCie.Value = FmtFec(adoAux!FCH_FINA, "yyyymmdd", "win", res)
'            End If
'            adoAux.Close: Set adoAux = Nothing
'        End With
'    End If

    Dim adoRegistro     As ADODB.Recordset
    Dim intRespuesta    As Integer
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        '.CommandText = "{ call up_ACSelDatosParametro(48,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        
        .CommandText = "{ call up_ACSelDatosParametro(48,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            dtpFechaApertura.Value = adoRegistro("FechaInicio")
                     
            strFechaApertura = Convertyyyymmdd(dtpFechaApertura.Value)
            gstrPeriodoActual = Format(Year(dtpFechaApertura.Value), "0000")
            gstrMesActual = Format(Month(dtpFechaApertura.Value), "00")
            gstrDiaActual = Format(Day(dtpFechaApertura.Value), "00")
            strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, dtpFechaApertura.Value))
            strFechaAnterior = Convertyyyymmdd(DateAdd("d", -1, dtpFechaApertura.Value))
            strFechaAnteAnterior = Convertyyyymmdd(DateAdd("d", -2, dtpFechaApertura.Value))
            strFechaSubSiguiente = Convertyyyymmdd(DateAdd("d", 2, dtpFechaApertura.Value))
    
            'Call ValidarFechas
            strCodMoneda = adoRegistro("CodMoneda")
      
            gdatFechaActual = adoRegistro("FechaInicio")
            gstrFechaActual = Convertyyyymmdd(gdatFechaActual)
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        Else
            MsgBox "No existe periodo contable para apertura en el fondo seleccionado!. Debe crear el periodo contable primero!", vbOKOnly + vbExclamation, Me.Caption
            adoRegistro.Close
            Exit Sub
        End If
        adoRegistro.Close
                        
        If Codigo_Moneda_Local <> strCodMoneda Then
            txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioFondo, gstrValorTipoCambioCierre, DateAdd("d", -1, dtpFechaApertura.Value), strCodMoneda, Codigo_Moneda_Local))
            If CDbl(txtTipoCambio.Text) = 0 Then
                MsgBox "Tipo de Cambio de Cierre NO REGISTRADO...", vbCritical, Me.Caption
                txtTipoCambio.Text = "0": Me.MousePointer = vbDefault
                Exit Sub
            End If
        Else
            txtTipoCambio.Text = "1"
        End If
            
    End With

End Sub


Private Sub Dat_FecCie_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Load()
  
    Dim strSQL As String
    
    Call InicializarValores
    Call CargarListas
    
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmPeriodoContableApertura = Nothing
    
End Sub


Private Sub Rea_TipCam_LostFocus()

    gdblTipoCambio = CDbl(txtTipoCambio.Text)
    
End Sub
Private Function TodoOK() As Boolean
                
    Dim adoConsulta As ADODB.Recordset
    Dim strMensaje  As String
    
    TodoOK = False
                
    If cboFondo.ListCount = 0 Then
        MsgBox "No existen fondos definidos...", vbCritical, Me.Caption
        Exit Function
    End If
    
    If cboFondo.ListIndex < 0 Then
        MsgBox "Seleccione el fondo...", vbCritical, Me.Caption
        cboFondo.SetFocus
        Exit Function
    End If
    
    If CDbl(txtTipoCambio.Text) <= 0 Then
        MsgBox "El Tipo de cambio para la fecha de cierre NO ESTA REGISTRADO...", vbCritical, Me.Caption
        Exit Function
    End If
    
                               
'    Set adoConsulta = New ADODB.Recordset
'    '*** Se Realizó Cierre anteriormente ? ***
'    adoComm.CommandText = "{ call up_GNValidaCierreRealizado('" & _
'        strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaApertura & "','" & _
'        strFechaSiguiente & "') }"
'    Set adoConsulta = adoComm.Execute
'
'    If Not adoConsulta.EOF Then
'        If Trim(adoConsulta("IndCierre")) = Valor_Caracter Then
'            MsgBox "El Cierre Diario del Día " & CStr(dtpFechaApertura.Value) & " aun no se ha realizado. No puede realizar el cierre anual.", vbCritical, Me.Caption
'            adoConsulta.Close: Set adoConsulta = Nothing
'            Exit Function
'        End If
'    End If
'    adoConsulta.Close

    '*** Verifica si se hizo apertura anual de periodo ***
    adoComm.CommandType = adCmdText
    adoComm.CommandText = "SELECT IndApertura FROM PeriodoContable " & _
        "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
        "MesContable='00' AND FechaInicio='" & strFechaApertura & "'"
    Set adoConsulta = adoComm.Execute
    
    If Not adoConsulta.EOF Then
        If Trim(adoConsulta("IndApertura")) = Valor_Indicador Then
            If MsgBox("La Apertura del Periodo ya fue realizada antes. Desea re-procesarla?", vbYesNo + vbExclamation, Me.Caption) = vbNo Then
                adoConsulta.Close: Set adoConsulta = Nothing
                Exit Function
            End If
        End If
    End If
    adoConsulta.Close
    

    '*** Cierre en fecha aún no abierta para el Fondo ***
'    adoComm.CommandText = "{ call up_GNValidaFechaNoAbierta('" & _
'        strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaApertura & "','" & _
'        strFechaSiguiente & "') }"
'    Set adoConsulta = adoComm.Execute
'
'    If Not adoConsulta.EOF Then
'        If Trim(adoConsulta("IndAbierto")) = Valor_Caracter Then
'            MsgBox "El Día " & CStr(dtpFechaApertura.Value) & " aún no ha sido abierto.", vbCritical, Me.Caption
'            adoConsulta.Close: Set adoConsulta = Nothing
'            Exit Function
'        End If
'    End If
'    adoConsulta.Close
        
    
   TodoOK = True
  
End Function
Private Sub CargarListas()
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    'If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
            
End Sub
Private Sub InicializarValores()

    dtpFechaApertura.Value = gdatFechaActual
    'dtpFechaEntrega.Value = DateAdd("d", gintDiasPagoRescate, dtpFechaApertura.Value)
    
    'Call ValidarFechas
    txtTipoCambio.Text = "0"
    
End Sub

Private Sub txtTipoCambio_Change()

    Call FormatoCajaTexto(txtTipoCambio, Decimales_TipoCambio)

End Sub
