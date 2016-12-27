VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmOperacionParticipe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operaciones de Partícipes"
   ClientHeight    =   7815
   ClientLeft      =   1425
   ClientTop       =   1605
   ClientWidth     =   12810
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7815
   ScaleWidth      =   12810
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   10440
      TabIndex        =   0
      Top             =   6960
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   600
      TabIndex        =   20
      Top             =   6960
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Buscar"
      Tag0            =   "5"
      ToolTipText0    =   "Buscar"
      UserControlWidth=   1200
   End
   Begin VB.Frame fraOperaciones 
      Caption         =   "Operaciones"
      Height          =   6615
      Left            =   140
      TabIndex        =   1
      Top             =   240
      Width           =   12420
      Begin VB.ComboBox cboTipoSolicitud 
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
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   795
         Width           =   4455
      End
      Begin VB.ComboBox cboFondoOperacion 
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
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   4455
      End
      Begin VB.ComboBox cboSucursal 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1620
         Width           =   3800
      End
      Begin VB.ComboBox cboAgencia 
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
         Left            =   4260
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1620
         Width           =   3800
      End
      Begin VB.ComboBox cboPromotor 
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
         Left            =   8320
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1620
         Width           =   3800
      End
      Begin VB.ListBox lstLeyenda 
         Height          =   255
         Left            =   9120
         TabIndex        =   2
         Top             =   2040
         Visible         =   0   'False
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker dtpFechaDesde 
         Height          =   315
         Left            =   9210
         TabIndex        =   11
         Top             =   360
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   48955393
         CurrentDate     =   38068
      End
      Begin MSComCtl2.DTPicker dtpFechaHasta 
         Height          =   315
         Left            =   9210
         TabIndex        =   12
         Top             =   795
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   48955393
         CurrentDate     =   38068
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmOperacionParticipe.frx":0000
         Height          =   3855
         Left            =   210
         OleObjectBlob   =   "frmOperacionParticipe.frx":001A
         TabIndex        =   3
         Top             =   2340
         Width           =   11895
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Tipo Operación"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   810
         Width           =   1335
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Operador"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   8340
         TabIndex        =   18
         Top             =   1275
         Width           =   1095
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Agencia"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   4290
         TabIndex        =   17
         Top             =   1275
         Width           =   1215
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Sucursal"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   1275
         Width           =   1215
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fondo"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   375
         Width           =   1335
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Desde"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   6
         Left            =   8340
         TabIndex        =   14
         Top             =   375
         Width           =   615
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Hasta"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   7
         Left            =   8340
         TabIndex        =   13
         Top             =   810
         Width           =   615
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Suscripciones (0)"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   2100
         Width           =   1815
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Rescates (0)"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   8
         Left            =   2400
         TabIndex        =   4
         Top             =   2100
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmOperacionParticipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondoOperacion()                 As String, arrTipoSolicitud()               As String
Dim arrSucursal()                       As String, arrAgencia()                     As String
Dim arrPromotor()                       As String, arrTipoOperacion()               As String
Dim arrFondo()                          As String, arrEjecutivo()                   As String
Dim arrTipoFormaPago()                  As String
Dim arrCuentaFondo()                    As String, arrNumCuenta()                   As String
Dim arrBanco()                          As String
Dim arrLeyendaSuscripcion()             As String, arrLeyendaRescate()              As String

Dim strCodFondoOperacion                As String, strCodTipoSolicitud              As String
Dim strCodSucursal                      As String, strCodAgencia                    As String
Dim strCodSucursalDestino               As String, strCodAgenciaDestino             As String
Dim strCodPromotor                      As String, strCodTipoOperacion              As String
Dim strCodClaseOperacion                As String, strCodTipoValuacion              As String
Dim strCodFondo                         As String, strCodEjecutivo                  As String
Dim strCodEjecutivoDestino              As String, strCodBancoDestino               As String
Dim strCodTipoDocumento                 As String, strCodTipoFormaPago              As String
Dim strCodCuentaFondo                   As String, strCodNumCuenta                  As String
Dim strCodBanco                         As String, strCodMonedaFondo                As String
Dim strEstado                           As String, strHoraCorte                     As String
Dim strSql                              As String

Dim dblTasaSuscripcion                  As Double, dblTasaRescate                   As Double
Dim dblCantCuotaMinSuscripcionInicial   As Double, dblMontoMinSuscripcionInicial    As Double
Dim dblCantMinCuotaSuscripcion          As Double, dblMontoMinSuscripcion           As Double
Dim dblPorcenMaxParticipe               As Double, dblValorCuota                    As Double
Dim dblCantCuotaInicio                  As Double, dblCantMaxCuotaFondo             As Double

Dim blnCuota                            As Boolean, blnMonto                        As Boolean
Dim blnValorConocido                    As Boolean, SeleccionLista                  As Integer

Dim adoConsulta                         As ADODB.Recordset
Dim indSortAsc                          As Boolean, indSortDesc                     As Boolean

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim strContinuar As String
    strContinuar = "1"
    
    Select Case Index
        Case 1
            gstrNameRepo = "ListaOperaciones"
                        
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(7)
            ReDim aReportParamFn(5)
            ReDim aReportParamF(5)

            strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
            strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value))
            
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            aReportParamFn(3) = "Fondo"
            aReportParamFn(4) = "FechaDel"
            aReportParamFn(5) = "FechaAl"
            
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
            aReportParamF(3) = Trim(cboFondoOperacion.Text)
            aReportParamF(4) = CStr(dtpFechaDesde.Value)
            aReportParamF(5) = CStr(dtpFechaHasta.Value)
                        
            aReportParamS(0) = strCodFondoOperacion
            aReportParamS(1) = gstrCodAdministradora
            aReportParamS(2) = strCodTipoSolicitud
            aReportParamS(3) = strFechaDesde
            aReportParamS(4) = strFechaHasta
            aReportParamS(5) = strCodSucursal
            aReportParamS(6) = strCodAgencia
            aReportParamS(7) = strCodPromotor
        Case 2
            If Not adoConsulta.EOF Then
                gstrNameRepo = "ParticipeOperacion"
                            
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(7)
                ReDim aReportParamFn(5)
                ReDim aReportParamF(5)
    
                strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
                strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value))
                
                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "Hora"
                aReportParamFn(2) = "NombreEmpresa"
                aReportParamFn(3) = "Fondo"
                aReportParamFn(4) = "FechaDesde"
                aReportParamFn(5) = "FechaHasta"
                
                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Format(Time(), "hh:mm:ss")
                aReportParamF(2) = gstrNombreEmpresa & Space(1)
                aReportParamF(3) = Trim(cboFondoOperacion.Text)
                aReportParamF(4) = CStr(dtpFechaDesde.Value)
                aReportParamF(5) = CStr(dtpFechaHasta.Value)
                            
                aReportParamS(0) = strCodFondoOperacion
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = strCodTipoSolicitud
                aReportParamS(3) = strFechaDesde
                aReportParamS(4) = strFechaHasta
                aReportParamS(5) = strCodSucursal
                aReportParamS(6) = strCodAgencia
                aReportParamS(7) = strCodPromotor
            Else
                MsgBox "Debe Seleccionar una Operacion de Participe para ver el Reporte", vbCritical
                strContinuar = "0"
            End If
         Case 3
         
            If Not adoConsulta.EOF Then
            'If SeleccionLista = 1 Then
                 
                gstrNameRepo = "RegistroOperaciones"
                            
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(3)
                ReDim aReportParamFn(5)
                ReDim aReportParamF(5)
                
                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "Hora"
                aReportParamFn(2) = "NombreEmpresa"
                aReportParamFn(3) = "Fondo"
                aReportParamFn(4) = "FechaDesde"
                aReportParamFn(5) = "FechaHasta"
                
                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Format(Time(), "hh:mm:ss")
                aReportParamF(2) = gstrNombreEmpresa & Space(1)
                aReportParamF(3) = Trim(cboFondoOperacion.Text)
                aReportParamF(4) = CStr(dtpFechaDesde.Value)
                aReportParamF(5) = CStr(dtpFechaHasta.Value)
                            
                aReportParamS(0) = strCodFondoOperacion
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = adoConsulta.Fields("CodParticipe")
                aReportParamS(3) = adoConsulta.Fields("NumOperacion")
            
            'Else
            '    strContinuar = 0
            'End If
            Else
                MsgBox "Debe Seleccionar una Operacion de Participe para ver el Reporte", vbCritical
                strContinuar = "0"
            End If
                
    End Select
     
        If strContinuar = "1" Then
        
            gstrSelFrml = ""
            frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"
        
            Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())
        
            frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
            frmReporte.Show vbModal
        
            Set frmReporte = Nothing
        
            Screen.MousePointer = vbNormal
        
        End If
End Sub
Public Sub Adicionar()
        
End Sub

Public Sub Buscar()
    SeleccionLista = 0
    Dim strFechaDesde       As String, strFechaHasta        As String
    Dim datFechaSiguiente   As Date
    Dim strSql              As String
                                                                                    
    Me.MousePointer = vbHourglass
    
    strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
    datFechaSiguiente = DateAdd("d", 1, dtpFechaHasta.Value)
    strFechaHasta = Convertyyyymmdd(datFechaSiguiente)
                
    If cboTipoSolicitud.ListIndex > 0 And cboSucursal.ListIndex > 0 And cboAgencia.ListIndex > 0 And cboPromotor.ListIndex > 0 Then
    
        strSql = "{ call up_ACSelDatosParametro(42,'" & strCodFondoOperacion & "','" & _
            gstrCodAdministradora & "','" & strFechaDesde & "','" & strFechaHasta & "','" & _
            strCodTipoSolicitud & "','" & strCodSucursal & "','" & strCodAgencia & "','" & _
            strCodPromotor & "') }"
            
    ElseIf cboTipoSolicitud.ListIndex > 0 And cboSucursal.ListIndex > 0 And cboAgencia.ListIndex > 0 Then
    
        strSql = "{ call up_ACSelDatosParametro(41,'" & strCodFondoOperacion & "','" & _
            gstrCodAdministradora & "','" & strFechaDesde & "','" & strFechaHasta & "','" & _
            strCodTipoSolicitud & "','" & strCodSucursal & "','" & strCodAgencia & "') }"
    
    ElseIf cboTipoSolicitud.ListIndex > 0 And cboSucursal.ListIndex > 0 Then
    
        strSql = "{ call up_ACSelDatosParametro(40,'" & strCodFondoOperacion & "','" & _
            gstrCodAdministradora & "','" & strFechaDesde & "','" & strFechaHasta & "','" & _
            strCodTipoSolicitud & "','" & strCodSucursal & "') }"
     
    ElseIf cboTipoSolicitud.ListIndex > 0 Then
    
        strSql = "{ call up_ACSelDatosParametro(39,'" & strCodFondoOperacion & "','" & _
            gstrCodAdministradora & "','" & strFechaDesde & "','" & strFechaHasta & "','" & _
            strCodTipoSolicitud & "','" & strCodSucursal & "') }"
     
    ElseIf cboTipoSolicitud.ListIndex <= 0 And cboSucursal.ListIndex > 0 And cboAgencia.ListIndex > 0 And cboPromotor.ListIndex > 0 Then
    
        strSql = "{ call up_ACSelDatosParametro(60,'" & strCodFondoOperacion & "','" & _
            gstrCodAdministradora & "','" & strFechaDesde & "','" & strFechaHasta & "','" & _
            strCodTipoSolicitud & "','" & strCodSucursal & "','" & strCodAgencia & "','" & _
            strCodPromotor & "') }"
            
    ElseIf cboTipoSolicitud.ListIndex <= 0 And cboSucursal.ListIndex > 0 And cboAgencia.ListIndex > 0 Then
    
        strSql = "{ call up_ACSelDatosParametro(61,'" & strCodFondoOperacion & "','" & _
            gstrCodAdministradora & "','" & strFechaDesde & "','" & strFechaHasta & "','" & _
            strCodTipoSolicitud & "','" & strCodSucursal & "','" & strCodAgencia & "') }"
    
    ElseIf cboTipoSolicitud.ListIndex <= 0 And cboSucursal.ListIndex > 0 Then
    
        strSql = "{ call up_ACSelDatosParametro(62,'" & strCodFondoOperacion & "','" & _
            gstrCodAdministradora & "','" & strFechaDesde & "','" & strFechaHasta & "','" & _
            strCodTipoSolicitud & "','" & strCodSucursal & "') }"
     
    Else
    
        strSql = "{ call up_ACSelDatosParametro(63,'" & strCodFondoOperacion & "','" & _
            gstrCodAdministradora & "','" & strFechaDesde & "','" & strFechaHasta & "','" & _
            strCodTipoSolicitud & "') }"
            
    End If
    
    Set adoConsulta = New ADODB.Recordset
    
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSql
    End With
        
    tdgConsulta.DataSource = adoConsulta
    
    If adoConsulta.RecordCount > 0 Then
        strEstado = Reg_Consulta
        
        Dim intNumSuscripciones As Integer, intNumRescates  As Integer

        intNumSuscripciones = 0: intNumRescates = 0
        With adoConsulta
            .MoveFirst
            
            Do While Not .EOF
                If Left(.Fields("TipoOperacion"), 1) = "S" Then intNumSuscripciones = intNumSuscripciones + 1
                If Left(.Fields("TipoOperacion"), 1) = "R" Then intNumRescates = intNumRescates + 1

                .MoveNext
            Loop
        End With
        lblDescrip(3).Caption = "Suscripciones (" & CStr(intNumSuscripciones) & ")"
        lblDescrip(8).Caption = "Rescates (" & CStr(intNumRescates) & ")"
    Else
        lblDescrip(3).Caption = "Suscripciones (0)"
        lblDescrip(8).Caption = "Rescates (0)"
    End If
            
    Me.MousePointer = vbDefault
    
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
        Case vExit
            Call Salir
        
    End Select
    
End Sub

Public Sub Eliminar()

    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        If MsgBox(Mensaje_Eliminacion, vbQuestion + vbYesNo, gstrNombreEmpresa) = vbYes Then
            adoComm.CommandText = "UPDATE ParticipeSolicitud SET EstadoSolicitud='" & Estado_Solicitud_Anulada & "' " & _
                "WHERE NumSolicitud='" & tdgConsulta.Columns(0) & "' AND CodParticipe='" & tdgConsulta.Columns(8) & "' AND " & _
                "CodFondo='" & strCodFondoOperacion & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            
            adoConn.Execute adoComm.CommandText

            Call Buscar
        End If
    End If

End Sub

Public Sub Grabar()
        
End Sub

Public Sub Imprimir()

End Sub

Public Sub Modificar()
   
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Private Sub cboAgencia_Click()

    Dim strSql As String
    
    strCodAgencia = ""
    If cboAgencia.ListIndex < 0 Then Exit Sub
    
    strCodAgencia = Trim(arrAgencia(cboAgencia.ListIndex))
    
    'strSQL = "{ call up_ACSelDatosParametro(11,'" & strCodAgencia & "') }"
    'CargarControlLista strSQL, cboPromotor, arrPromotor(), Sel_Todos
    
    'If cboPromotor.ListCount > -1 Then cboPromotor.ListIndex = 0
    
End Sub

Private Sub cboFondoOperacion_Click()

    strCodFondoOperacion = Valor_Caracter
    If cboFondoOperacion.ListIndex < 0 Then Exit Sub
    
    strCodFondoOperacion = Trim(arrFondoOperacion(cboFondoOperacion.ListIndex))
    
End Sub

Private Sub cboPromotor_Click()

    strCodPromotor = ""
    If cboPromotor.ListIndex < 0 Then Exit Sub
    
    strCodPromotor = Trim(arrPromotor(cboPromotor.ListIndex))
    
End Sub

Private Sub cboSucursal_Click()

    Dim strSql As String
    
    strCodSucursal = ""
    If cboSucursal.ListIndex < 0 Then Exit Sub
    
    strCodSucursal = Trim(arrSucursal(cboSucursal.ListIndex))
    
    strSql = "{ call up_ACSelDatosParametro(10,'" & strCodSucursal & "') }"
    CargarControlLista strSql, cboAgencia, arrAgencia(), Sel_Todos
    
    If cboAgencia.ListCount > -1 Then cboAgencia.ListIndex = 0
    
End Sub

Private Sub cboTipoSolicitud_Click()

    strCodTipoSolicitud = ""
    If cboTipoSolicitud.ListIndex < 0 Then Exit Sub
    
    strCodTipoSolicitud = Trim(arrTipoSolicitud(cboTipoSolicitud.ListIndex))
    
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

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Lista de Operaciones"
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Lista de Operaciones Detalladas"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Text = "Documento Registro Operación"
    
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

Private Sub CargarListas()

    Dim adoRegistro As ADODB.Recordset
    Dim intRegistro As Integer
    
    '*** Fondos ***
    strSql = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSql, cboFondoOperacion, arrFondoOperacion(), Valor_Caracter
    
    If cboFondoOperacion.ListCount > 0 Then cboFondoOperacion.ListIndex = 0
    
    '*** Tipo Solicitud/Operación ***
    strSql = "SELECT CodTipoSolicitud CODIGO,DescripTipoSolicitud DESCRIP FROM TipoSolicitud WHERE CodCorto<>'M' and CodCorto<>'O' ORDER BY DescripTipoSolicitud"
    CargarControlLista strSql, cboTipoSolicitud, arrTipoSolicitud(), Sel_Todos
    
    If cboTipoSolicitud.ListCount > 0 Then cboTipoSolicitud.ListIndex = 0
                        
    '*** Sucursal ***
    strSql = "{ call up_ACSelDatos(15) }"
    CargarControlLista strSql, cboSucursal, arrSucursal(), Sel_Todos
    
    If cboSucursal.ListCount > 0 Then cboSucursal.ListIndex = 0
    
    Set adoRegistro = New ADODB.Recordset
    
    adoComm.CommandText = "SELECT CodSucursal,CodAgencia FROM InstitucionPersona " & _
    "WHERE TipoPersona='01' AND CodPersona='" & gstrCodPromotor & "'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        intRegistro = ObtenerItemLista(arrSucursal(), Trim(adoRegistro("CodSucursal")))
        If intRegistro >= 0 Then cboSucursal.ListIndex = 0
        'If intRegistro >= 0 Then cboSucursal.ListIndex = intRegistro
        
        intRegistro = ObtenerItemLista(arrAgencia(), Trim(adoRegistro("CodAgencia")))
        If intRegistro >= 0 Then cboAgencia.ListIndex = intRegistro
        
        intRegistro = ObtenerItemLista(arrPromotor(), gstrCodPromotor)
        'If intRegistro >= 0 Then cboPromotor.ListIndex = intRegistro
        If intRegistro >= 0 Then cboPromotor.ListIndex = 0
    End If
    adoRegistro.Close: Set adoRegistro = Nothing
           
    
    strSql = "{ call up_ACSelDatos(41) }"
    CargarControlLista strSql, cboPromotor, arrPromotor(), Sel_Defecto
    
    If cboPromotor.ListCount > 0 Then cboPromotor.ListIndex = 0
    intRegistro = ObtenerItemLista(arrPromotor(), gstrCodPromotor)
    If intRegistro >= 0 Then cboPromotor.ListIndex = 0
    'If intRegistro >= 0 Then cboPromotor.ListIndex = intRegistro
End Sub

Private Sub InicializarValores()

    Dim adoRegistro As ADODB.Recordset
    Dim intCont     As Integer
    
    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    blnValorConocido = False

    dtpFechaDesde.Value = gdatFechaActual
    dtpFechaHasta.Value = gdatFechaActual
    
    Set adoRegistro = New ADODB.Recordset
    
    intCont = 0
    ReDim arrLeyendaSuscripcion(intCont)
    ReDim arrLeyendaRescate(intCont)
    
    With adoComm
        '*** Suscripción ***
        .CommandText = "SELECT RTRIM(TSD.CodCorto) + space(1) + '=' + space(1) + RTRIM(TS.DescripTipoSolicitud) + space(1) + RTRIM(TSD.DescripDetalleTipoSolicitud)  ValorLeyenda " & _
            "FROM TipoSolicitud TS JOIN TipoSolicitudDetalle TSD " & _
            "ON(TSD.CodTipoSolicitud=TS.CodTipoSolicitud AND " & _
            "TS.CodTipoSolicitud='" & Codigo_Operacion_Suscripcion & "') " & _
            "WHERE TS.CodCorto<>'T'"
        Set adoRegistro = .Execute
        
        Do Until adoRegistro.EOF
            ReDim Preserve arrLeyendaSuscripcion(intCont)
            arrLeyendaSuscripcion(intCont) = adoRegistro("ValorLeyenda")
            
            adoRegistro.MoveNext
            intCont = intCont + 1
        Loop
        adoRegistro.Close
        
        intCont = 0
        '*** Rescate ***
        .CommandText = "SELECT RTRIM(TSD.CodCorto) + space(1) + '=' + space(1) + RTRIM(TS.DescripTipoSolicitud) + space(1) + RTRIM(TSD.DescripDetalleTipoSolicitud)  ValorLeyenda " & _
            "FROM TipoSolicitud TS JOIN TipoSolicitudDetalle TSD " & _
            "ON(TSD.CodTipoSolicitud=TS.CodTipoSolicitud AND " & _
            "TS.CodTipoSolicitud='" & Codigo_Operacion_Rescate & "') " & _
            "WHERE TS.CodCorto<>'T'"
        Set adoRegistro = .Execute
        
        Do Until adoRegistro.EOF
            ReDim Preserve arrLeyendaRescate(intCont)
            arrLeyendaRescate(intCont) = adoRegistro("ValorLeyenda")
            
            adoRegistro.MoveNext
            intCont = intCont + 1
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
            
    End With
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 5
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 9
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 32
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 14
    tdgConsulta.Columns(5).Width = tdgConsulta.Width * 0.01 * 12
    
    '*** Verificando Nivel de Acceso de Usuario ***
'    strNivAcceso = AccesoForm(gstrNomOpc, gstrNumInd)

    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmOperacionParticipe = Nothing
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción..."
    
End Sub

Private Sub lblDescrip_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim intContador As Integer
    
    If Index = 3 Then
        intContador = UBound(arrLeyendaSuscripcion)
        
        lstLeyenda.AddItem "Leyenda :"
        lstLeyenda.AddItem ""
        
        For intContador = 0 To (UBound(arrLeyendaSuscripcion))
            lstLeyenda.AddItem arrLeyendaSuscripcion(intContador)
        Next
        
        lstLeyenda.Height = lblDescrip(Index).Height * (intContador + 2)
        lstLeyenda.Left = lblDescrip(Index).Left
        lstLeyenda.Top = lblDescrip(Index).Top + lblDescrip(Index).Height
        lstLeyenda.Width = 3300
        lstLeyenda.Visible = True
    End If
    
    If Index = 8 Then
        intContador = UBound(arrLeyendaRescate)
        
        lstLeyenda.AddItem "Leyenda :"
        lstLeyenda.AddItem ""
        
        For intContador = 0 To (UBound(arrLeyendaRescate))
            lstLeyenda.AddItem arrLeyendaRescate(intContador)
        Next
        
        lstLeyenda.Height = lblDescrip(Index).Height * (intContador + 2)
        lstLeyenda.Left = lblDescrip(Index).Left
        lstLeyenda.Top = lblDescrip(Index).Top + lblDescrip(Index).Height
        lstLeyenda.Width = 3800
        lstLeyenda.Visible = True
    End If
    
End Sub

Private Sub lblDescrip_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Index = 3 Or Index = 8 Then
        lstLeyenda.Clear
        lstLeyenda.Visible = False
    End If
    
End Sub

Private Sub tdgConsulta_Click()
    SeleccionLista = 1
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 4 Then
        Call DarFormatoValor(Value, Decimales_CantCuota)
    End If
    
    If ColIndex = 5 Then
        Call DarFormatoValor(Value, Decimales_ValorCuota)
    End If
    
    If ColIndex = 6 Then
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

    Call OrdenarDBGrid(ColIndex, adoConsulta, tdgConsulta)
    
    numColindex = ColIndex
    
    '****
    strPrevColumTDB = strColNameTDB
    '***
    
End Sub

