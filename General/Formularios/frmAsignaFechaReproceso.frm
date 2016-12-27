VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAsignaFechaReproceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Habilitar Fecha de Reproceso"
   ClientHeight    =   2550
   ClientLeft      =   1635
   ClientTop       =   1785
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2550
   ScaleWidth      =   7530
   Begin VB.ComboBox cboFondo 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   330
      Width           =   5835
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Cancelar"
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
      Index           =   1
      Left            =   4080
      Picture         =   "frmAsignaFechaReproceso.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1200
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Aceptar"
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
      Index           =   0
      Left            =   2400
      Picture         =   "frmAsignaFechaReproceso.frx":0562
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker dtpFechaTrabajo 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   1065
      Width           =   1365
      _ExtentX        =   2408
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
      Format          =   146669569
      CurrentDate     =   38068
   End
   Begin VB.Label lblDescrip 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   450
   End
   Begin VB.Label lblDescrip 
      AutoSize        =   -1  'True
      Caption         =   "Portafolio"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   660
   End
End
Attribute VB_Name = "frmAsignaFechaReproceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()      As String, strSQL       As String
Dim strCodFondo     As String

Private Sub CargarListas()

    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
End Sub

Private Sub cboFondo_Click()

    Dim adoRegistro     As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            dtpFechaTrabajo.MaxDate = "01/01/2100"
            dtpFechaTrabajo.Value = adoRegistro("FechaCuota")
            dtpFechaTrabajo.MaxDate = dtpFechaTrabajo.Value
            gdatFechaActual = adoRegistro("FechaCuota")
        
            gdatFechaActual = adoRegistro("FechaCuota")
            gdblTipoCambio = CDbl(adoRegistro("ValorTipoCambio"))
            gstrFechaActual = Convertyyyymmdd(adoRegistro("FechaCuota"))
            gstrCodMoneda = adoRegistro("CodMoneda")
            
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub


Private Sub cmdAccion_Click(Index As Integer)

    Select Case Index
        Case 0
            Dim adoRegistro             As ADODB.Recordset
            Dim strFechaInicio          As String, strFechaFin  As String
            Dim strFechaAnterior        As String
            Dim intNumReproceso         As Integer
            Dim strFechaHastaReproceso  As String
            
            strFechaAnterior = Convertyyyymmdd(DateAdd("d", -1, dtpFechaTrabajo.Value))
            strFechaInicio = Convertyyyymmdd(dtpFechaTrabajo.Value)
            strFechaFin = Convertyyyymmdd(DateAdd("d", 1, dtpFechaTrabajo.Value))
            strFechaHastaReproceso = Convertyyyymmdd(dtpFechaTrabajo.MaxDate)
            
            '*** Continuar si es menor a la fecha hábil ***
            If dtpFechaTrabajo.Value < dtpFechaTrabajo.MaxDate Or Year(dtpFechaTrabajo.Value) <> Year(DateAdd("d", 1, dtpFechaTrabajo.Value)) Or Year(dtpFechaTrabajo.Value) <> Year(gdatFechaActual) Then
                With adoComm
                
                    '*** Adicionar registro de reproceso ***
                    Set adoRegistro = New ADODB.Recordset
                    
                    adoComm.CommandText = "SELECT COUNT(*) SecuencialReproceso FROM FondoReproceso"
                    Set adoRegistro = adoComm.Execute
                    
                    If Not adoRegistro.EOF Then
                        intNumReproceso = adoRegistro("SecuencialReproceso") + 1
                    Else
                        intNumReproceso = 1
                    End If
                    adoRegistro.Close: Set adoRegistro = Nothing
            
                    .CommandText = "INSERT INTO FondoReproceso VALUES(" & intNumReproceso & ",'" & _
                        strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaInicio & "','" & _
                        Convertyyyymmdd(dtpFechaTrabajo.MaxDate) & "','" & Estado_Activo & "')"
                    adoConn.Execute .CommandText
                
                    '*** Actualizar Fecha Hábil de Reproceso ***
                    .CommandText = "UPDATE FondoValorCuota SET IndAbierto='',IndSuscripcionConocida='',IndSuscripcionDesconocida='' " & _
                        "WHERE IndAbierto='X' AND " & _
                        "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "'"
                    adoConn.Execute .CommandText

                    .CommandText = "UPDATE FondoValorCuota SET IndAbierto='X',IndCierre='',IndCalculado='X'," & _
                        "MontoGPDiarioPreCierre=0,MontoGPAcumuladoPreCierre=MontoGPAcumuladoPreCierre - MontoGPDiarioPreCierre," & _
                        "MontoGPDiarioCierre=0,MontoGPAcumuladoCierre=MontoGPAcumuladoCierre - MontoGPDiarioCierre, " & _
                        "ValorCuotaFinal = 0, CantCuotaFinal = 0, CantCuotaFinalPagada = 0, CantCuotaFinalPagadaReal = 0 " & _
                        " WHERE (FechaCuota >='" & strFechaInicio & "' AND FechaCuota <'" & strFechaFin & "') AND " & _
                        "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "'"
                    adoConn.Execute .CommandText
                    
'                    Set adoRegistro = New ADODB.Recordset
                
                    '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
                    .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
                    Set adoRegistro = .Execute

                    If Not adoRegistro.EOF Then
                        gdatFechaActual = adoRegistro("FechaCuota"): gdblTipoCambio = CDbl(adoRegistro("ValorTipoCambio"))
                        gstrFechaActual = Convertyyyymmdd(adoRegistro("FechaCuota"))
                        gstrCodMoneda = adoRegistro("CodMoneda")

                        frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
                    End If
                    adoRegistro.Close
                    
                    strFechaInicio = gstrFechaActual
                    strFechaFin = Convertyyyymmdd(DateAdd("d", 1, gdatFechaActual))
                    
'                    '*** Eliminar la información de partícipes ***
'                    .CommandText = "DELETE ParticipeOperacion " & _
'                        "WHERE FechaOperacion >='" & strFechaFin & "' AND " & _
'                        "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "'"
'                    adoConn.Execute .CommandText
'
'                    .CommandText = "DELETE ParticipeCertificado " & _
'                        "WHERE FechaOperacion >='" & strFechaFin & "' AND " & _
'                        "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "'"
'                    adoConn.Execute .CommandText
                    
'                    '*** Eliminar la información contable ***
'                    .CommandText = "DELETE PartidaContableSaldos " & _
'                        "WHERE FechaSaldo >='" & strFechaInicio & "' AND " & _
'                        "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "'"
'                    adoConn.Execute .CommandText
'
'                    .CommandText = "DELETE PartidaContablePreSaldos " & _
'                        "WHERE FechaSaldo >='" & strFechaInicio & "' AND " & _
'                        "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "'"
'                    adoConn.Execute .CommandText
                    
'                    '*** Anular los comprobantes ***
'                    .CommandText = "UPDATE AsientoContable SET EstadoAsiento='" & Estado_Eliminado & "' " & _
'                        "WHERE (FechaAsiento >='" & strFechaInicio & "' AND FechaAsiento <'" & strFechaFin & "') AND " & _
'                        "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "' AND " & _
'                        "TipoAsiento='" & Codigo_Tipo_Asiento_Cierre & "'"
'                    adoConn.Execute .CommandText
'
'                    .CommandText = "UPDATE AsientoContable SET EstadoAsiento='" & Estado_Eliminado & "' " & _
'                        "WHERE FechaAsiento >='" & strFechaFin & "' AND " & _
'                        "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "'"
'                    adoConn.Execute .CommandText
'
'                    '*** Actualizar las solicitudes de participación ***
'                    .CommandText = "UPDATE ParticipeSolicitud SET EstadoSolicitud='" & Estado_Solicitud_Confirmada & "' " & _
'                        "WHERE FechaSolicitud >='" & strFechaFin & "' AND " & _
'                        "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "'"
'                    adoConn.Execute .CommandText
'
'                    '*** Anular las ordenes de cobro/pago ***
'                    .CommandText = "UPDATE MovimientoFondo SET EstadoOrden='" & Estado_Caja_Anulado & "' " & _
'                        "WHERE FechaRegistro >='" & strFechaFin & "' AND ModuloOrigen <> 'I' AND " & _
'                        "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "'"
'                    adoConn.Execute .CommandText
                    
                    '*** Eliminar la información de valorizacion ***
                    .CommandText = "DELETE InversionValorizacion " & _
                        "WHERE FechaValorizacion >='" & strFechaInicio & "' AND " & _
                        "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "'"
                    adoConn.Execute .CommandText
                    
                    '*** Eliminar la información de detalle de valorizacion ***
                    .CommandText = "DELETE InversionValorizacionDiaria " & _
                        "WHERE FechaValorizacion >='" & strFechaInicio & "' AND " & _
                        "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "'"
                    adoConn.Execute .CommandText
                    
                    '*** Extorna saldos y anula comprobantes de pago cierre
                    .CommandText = "{ call up_GNProcExtornaAsientoReproceso('" & _
                        strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        strFechaInicio & "','" & strFechaHastaReproceso & "') }"
                    adoConn.Execute .CommandText
                    
                    .CommandText = "{ call up_GNProcExtornaOperacionParticipe('" & _
                        strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        strFechaInicio & "','" & strFechaHastaReproceso & "') }"
                    adoConn.Execute .CommandText
                    
'                    '*** Verificar valores vencidos ***
'                    .CommandText = "SELECT CodTitulo,CodAnalitica,CodFile,NumKardex " & _
'                        "FROM InversionKardex " & _
'                        "WHERE FechaOperacion >='" & strFechaFin & "' AND " & _
'                        "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "' AND " & _
'                        "IndUltimoMovimiento='X'"
'                    Set adoRegistro = .Execute
'
'                    Do While Not adoRegistro.EOF
'                        '*** Anular último movimiento del kardex ***
'                        .CommandText = "UPDATE InversionKardex SET IndUltimoMovimiento='',IndNoConfirma='X' " & _
'                            "WHERE FechaOperacion >='" & strFechaFin & "' AND " & _
'                            "CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                            "CodFondo='" & strCodFondo & "'"
'                        adoConn.Execute .CommandText
'
'                        '*** Actualizar último movimiento del kardex ***
'                        .CommandText = "UPDATE InversionKardex SET IndUltimoMovimiento='X' " & _
'                            "WHERE CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "' AND " & _
'                            "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "' AND " & _
'                            "NumKardex=(SELECT MAX(NumKardex) FROM InversionKardex IK " & _
'                            "WHERE IK.FechaOperacion <'" & strFechaFin & "' AND IK.CodTitulo=InversionKardex.CodTitulo AND " & _
'                            "IK.CodAdministradora = InversionKardex.CodAdministradora And IK.CodFondo = InversionKardex.CodFondo AND " & _
'                            "IK.IndNoConfirma='')"
'                        adoConn.Execute .CommandText
'
'                        '*** Actualizar vigencia del título valor ***
'                        .CommandText = "UPDATE InstrumentoInversion SET IndVigente='X' " & _
'                            "WHERE CodTitulo='" & adoRegistro("CodTitulo") & "'"
'                        adoConn.Execute .CommandText
'
'                        '*** Actualizar vencimiento de cupones ***
'                        .CommandText = "UPDATE InstrumentoInversionCalendario SET IndVencido='X',IndVigente='' " & _
'                            "WHERE FechaVencimiento <'" & strFechaFin & "' AND CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "'"
'                        adoConn.Execute .CommandText
'
'                        '*** Actualizar vigencia de cupones ***
'                        .CommandText = "UPDATE InstrumentoInversionCalendario SET IndVigente='X' " & _
'                            "WHERE CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "' AND IndVencido='' AND " & _
'                            "FechaInicio=(SELECT MIN(FechaInicio) FROM InstrumentoInversionCalendario IVC " & _
'                            "WHERE IVC.IndVencido='' AND IVC.CodTitulo=InstrumentoInversionCalendario.CodTitulo)"
'                        adoConn.Execute .CommandText
'
'
'                        adoRegistro.MoveNext
'                    Loop
'                    adoRegistro.Close
                    
'                    '*** Trasladar Saldos Iniciales de la fecha de reproceso ***
'                    .CommandText = "{ call up_GNProcTrasladoSaldosInicialesDiaSiguiente('" & _
'                        strCodFondo & "','" & gstrCodAdministradora & "','" & _
'                        strFechaAnterior & "','" & strFechaInicio & "') }"
'                    adoConn.Execute .CommandText
                    
                    
'                    '*** Actualizar Saldos del día de reproceso***
'                    .CommandText = "SELECT AC.PeriodoContable,AC.MesContable,CodCuenta,CodFile,CodAnalitica," & _
'                        "IndDebeHaber,ACD.CodMoneda,MontoMovimientoMN,MontoMovimientoME,MontoContable " & _
'                        "FROM AsientoContable AC JOIN AsientoContableDetalle ACD ON(ACD.FechaMovimiento=AC.FechaAsiento AND " & _
'                        "ACD.NumAsiento=AC.NumAsiento AND ACD.CodAdministradora=AC.CodAdministradora AND ACD.CodFondo=AC.CodFondo) " & _
'                        "WHERE (FechaAsiento >='" & strFechaInicio & "' AND FechaAsiento < '" & strFechaFin & "') AND " & _
'                        "AC.CodAdministradora='" & gstrCodAdministradora & "' AND AC.CodFondo='" & strCodFondo & "' AND " & _
'                        "EstadoAsiento='" & Estado_Activo & "'"
'                    Set adoRegistro = .Execute
'
'                    Do While Not adoRegistro.EOF
'                        .CommandText = "{ call up_ACGenPartidaContableSaldos('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
'                            adoRegistro("PeriodoContable") & "','" & adoRegistro("MesContable") & "','" & Trim(adoRegistro("CodCuenta")) & "','" & _
'                            adoRegistro("CodFile") & "','" & adoRegistro("CodAnalitica") & "','" & strFechaInicio & "','" & _
'                            strFechaFin & "'," & CDec(adoRegistro("MontoMovimientoMN")) & "," & CDec(adoRegistro("MontoMovimientoME")) & "," & _
'                            CDec(adoRegistro("MontoContable")) & ",'" & Trim(adoRegistro("IndDebeHaber")) & "','" & _
'                            Trim(adoRegistro("CodMoneda")) & "') }"
'                        adoConn.Execute .CommandText
'
'                        adoRegistro.MoveNext
'                    Loop
'                    adoRegistro.Close: Set adoRegistro = Nothing
                    Set adoRegistro = Nothing
                End With
            
                MsgBox Mensaje_Proceso_Exitoso, vbInformation, Me.Caption
            End If
            Unload Me
        Case 1: Unload Me
    End Select
    
End Sub


Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call DarFormato
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
    
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
            
End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    dtpFechaTrabajo.Value = gdatFechaActual
    strCodFondo = "000"
            
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    Set frmAsignaFechaReproceso = Nothing
    
End Sub


