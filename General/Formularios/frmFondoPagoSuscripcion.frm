VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmFondoPagoSuscripcion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago de Cuotas por Suscripción"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   8010
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   6120
      TabIndex        =   2
      Top             =   4560
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
      TabIndex        =   1
      Top             =   4560
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1296
      Buttons         =   2
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      ToolTipText1    =   "Modificar"
      UserControlWidth=   2700
   End
   Begin TabDlg.SSTab tabPagos 
      Height          =   4245
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7488
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmFondoPagoSuscripcion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCriterio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmFondoPagoSuscripcion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDatos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAccion"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -70680
         TabIndex        =   7
         Top             =   3360
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   1296
         Buttons         =   2
         Caption0        =   "&Guardar"
         Tag0            =   "2"
         ToolTipText0    =   "Guardar"
         Caption1        =   "&Cancelar"
         Tag1            =   "8"
         ToolTipText1    =   "Cancelar"
         UserControlWidth=   2700
      End
      Begin VB.Frame fraCriterio 
         Caption         =   "Criterios de Búsqueda"
         Height          =   945
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   6825
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   4815
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   14
            Top             =   360
            Width           =   450
         End
      End
      Begin VB.Frame fraDatos 
         Caption         =   "Datos"
         Height          =   2775
         Left            =   -74760
         TabIndex        =   9
         Top             =   480
         Width           =   6855
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   285
            Left            =   1800
            TabIndex        =   3
            Top             =   1560
            Width           =   1460
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   109903873
            CurrentDate     =   38949
         End
         Begin VB.TextBox txtMontoPago 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4920
            MaxLength       =   12
            TabIndex        =   6
            Text            =   " "
            Top             =   2160
            Width           =   1460
         End
         Begin VB.TextBox txtPorcenPago 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            MaxLength       =   40
            TabIndex        =   5
            Top             =   2160
            Width           =   1460
         End
         Begin MSComCtl2.DTPicker dtpFechaHasta 
            Height          =   285
            Left            =   4920
            TabIndex        =   4
            Top             =   1560
            Width           =   1460
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   109903873
            CurrentDate     =   38949
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   3720
            TabIndex        =   20
            Top             =   1560
            Width           =   420
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Pago (%)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   19
            Top             =   2180
            Width           =   630
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   3720
            TabIndex        =   18
            Top             =   2160
            Width           =   450
         End
         Begin VB.Label lblNumSecuencial 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   17
            Top             =   960
            Width           =   1020
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   16
            Top             =   500
            Width           =   540
         End
         Begin VB.Label lblDescripFondo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   15
            Top             =   480
            Width           =   4575
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Secuencial"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   11
            Top             =   975
            Width           =   795
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   10
            Top             =   1560
            Width           =   465
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmFondoPagoSuscripcion.frx":0038
         Height          =   2055
         Left            =   240
         OleObjectBlob   =   "frmFondoPagoSuscripcion.frx":0052
         TabIndex        =   12
         Top             =   1560
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmFondoPagoSuscripcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()          As String
Dim strCodFondo         As String
Dim strEstado           As String, strSQL               As String
Dim adoConsulta         As ADODB.Recordset
Dim indSortAsc          As Boolean, indSortDesc         As Boolean

Private Sub cboFondo_Click()

    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Call Buscar

End Sub

Private Sub dtpFechaDesde_Change()

    If Not EsDiaUtil(dtpFechaDesde.Value) Then
        MsgBox "La Fecha no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
        If dtpFechaDesde.Value >= gdatFechaActual Then
            dtpFechaDesde.Value = AnteriorDiaUtil(dtpFechaDesde.Value)
        Else
            dtpFechaDesde.Value = ProximoDiaUtil(dtpFechaDesde.Value)
        End If
    End If
    dtpFechaHasta.Value = dtpFechaDesde.Value
    
End Sub


Private Sub dtpFechaHasta_Change()

    If Not EsDiaUtil(dtpFechaHasta.Value) Then
        MsgBox "La Fecha no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
        If dtpFechaHasta.Value >= gdatFechaActual Then
            dtpFechaHasta.Value = AnteriorDiaUtil(dtpFechaHasta.Value)
        Else
            dtpFechaHasta.Value = ProximoDiaUtil(dtpFechaHasta.Value)
        End If
    End If
    
    If dtpFechaHasta.Value < dtpFechaDesde.Value Then
        dtpFechaHasta.Value = dtpFechaDesde.Value
    End If
    
End Sub


Private Sub Form_Activate()

    Call CargarReportes
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Pago de Cuotas por Suscripción"
    
End Sub
Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vModify
            Call Modificar
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

Public Sub Salir()

    Unload Me
    
End Sub
Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabPagos
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .Tab = 0
    End With
    strEstado = Reg_Consulta
    
End Sub
Public Sub Grabar()
                        
    Dim adoRegistro     As ADODB.Recordset
    Dim adoAuxiliar     As ADODB.Recordset
    Dim adoTemporal     As ADODB.Recordset
    Dim intAccion       As Integer, lngNumError     As Integer
    Dim dblPorcenTotal  As Double
    
    If strEstado = Reg_Consulta Then Exit Sub
            
    On Error GoTo CtrlError
    
    If strEstado = Reg_Adicion Then
        If TodoOK() Then
            frmMainMdi.stbMdi.Panels(3).Text = "Grabar Cronograma de Pagos de Cuotas..."
            
            dblPorcenTotal = 0
            If adoConsulta.RecordCount > 0 Then
                
                adoConsulta.MoveFirst
                                                    
                Do While Not adoConsulta.EOF
                    
                    If adoConsulta.Fields("NumSecuencial") <> CInt(lblNumSecuencial.Caption) Then
                        dblPorcenTotal = dblPorcenTotal + CDbl(adoConsulta.Fields("PorcenPago"))
                    End If
                    
                    adoConsulta.MoveNext
                Loop
            
            End If
            
            dblPorcenTotal = dblPorcenTotal + CDbl(txtPorcenPago.Text)
            
            If dblPorcenTotal > 100 Then
                MsgBox "El Porcentaje Total excede el 100%, verifique!", vbCritical, Me.Caption
                Exit Sub
            End If
            
            Me.MousePointer = vbHourglass
            
            With adoComm
                .CommandText = "{ call up_GNManFondoPagoSuscripcion('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    CInt(lblNumSecuencial.Caption) & "','" & _
                    Convertyyyymmdd(dtpFechaDesde.Value) & "','" & Convertyyyymmdd(dtpFechaHasta.Value) & "'," & _
                    CDec(txtPorcenPago.Text) & "," & CDec(txtMontoPago.Text) & ",'I') }"
                adoConn.Execute .CommandText
                
                '*** Pago Parcial de Cuotas de Suscripción ***
                Dim datFechaPagoParcial     As Date
                Dim curMontoSuscripcion     As Currency
                Dim dblCuotasSuscripcion    As Double
                Dim dblCuotasPagadas        As Double

                Set adoRegistro = New ADODB.Recordset
                                                
                '*** Adicionar en Cronograma de Pago de Partícipes ***
                .CommandText = "SELECT CodParticipe,NumSolicitud,NumOperacion,NumCertificado " & _
                    "FROM ParticipePagoSuscripcion " & _
                    "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "NumSecuencial=" & CInt(lblNumSecuencial.Caption) - 1
                Set adoRegistro = .Execute
                
                If Not adoRegistro.EOF Then
                    Do While Not adoRegistro.EOF
                        .CommandText = "{ call up_PRManParticipePagoSuscripcion('" & _
                            strCodFondo & "','" & gstrCodAdministradora & "','" & adoRegistro("CodParticipe") & "','" & _
                            adoRegistro("NumSolicitud") & "'," & CInt(lblNumSecuencial.Caption) & ",'" & _
                            Convertyyyymmdd(dtpFechaHasta.Value) & "'," & _
                            "0,'" & _
                            Convertyyyymmdd(CVDate(Valor_Fecha)) & "',0,'" & _
                            "I') }"
                        adoConn.Execute .CommandText
                        
                        .CommandText = "UPDATE ParticipePagoSuscripcion SET " & _
                            "NumOperacion='" & adoRegistro("NumOperacion") & "',NumCertificado='" & adoRegistro("NumCertificado") & "' " & _
                            "WHERE CodParticipe='" & adoAuxiliar("CodParticipe") & "' AND NumSolicitud='" & adoAuxiliar("NumSolicitud") & "' AND " & _
                            "NumSecuencial=" & CInt(lblNumSecuencial.Caption) & " AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                        adoConn.Execute .CommandText
                                                                
                        adoRegistro.MoveNext
                    Loop
                Else
                    adoRegistro.Close
                    
                    '*** Adicionar en Cronograma de Pago de Partícipes ***
                    .CommandText = "SELECT CodParticipe,NumSolicitud,NumOperacion,NumCertificado " & _
                        "FROM ParticipePagoSuscripcionTmp " & _
                        "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                        "NumSecuencial=" & CInt(lblNumSecuencial.Caption) - 1
                    Set adoRegistro = .Execute
                
                    Do While Not adoRegistro.EOF
                        .CommandText = "{ call up_PRManParticipePagoSuscripcionTmp('" & _
                            strCodFondo & "','" & gstrCodAdministradora & "','" & adoRegistro("CodParticipe") & "','" & _
                            adoRegistro("NumSolicitud") & "'," & CInt(lblNumSecuencial.Caption) & ",'" & _
                            Convertyyyymmdd(dtpFechaHasta.Value) & "'," & _
                            "0,'" & _
                            Convertyyyymmdd(CVDate(Valor_Fecha)) & "',0,'" & _
                            "I') }"
                        adoConn.Execute .CommandText
                        
                        .CommandText = "UPDATE ParticipePagoSuscripcionTmp SET " & _
                            "NumOperacion='" & adoRegistro("NumOperacion") & "',NumCertificado='" & adoRegistro("NumCertificado") & "' " & _
                            "WHERE CodParticipe='" & adoAuxiliar("CodParticipe") & "' AND NumSolicitud='" & adoAuxiliar("NumSolicitud") & "' AND " & _
                            "NumSecuencial=" & CInt(lblNumSecuencial.Caption) & " AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                        adoConn.Execute .CommandText
                                                                
                        adoRegistro.MoveNext
                    Loop
                End If
                adoRegistro.Close
                
                '*** Actualizar valores en Cronograma de Pago de Partícipes ***
                .CommandText = "SELECT NumSecuencial,FechaDesde,FechaHasta,PorcenPago,MontoPago FROM FondoPagoSuscripcion " & _
                    "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' " & _
                    "ORDER BY NumSecuencial"
                Set adoRegistro = .Execute

                Do While Not adoRegistro.EOF
                    Set adoAuxiliar = New ADODB.Recordset
                    
                    datFechaPagoParcial = CVDate(adoRegistro("FechaHasta"))

                    .CommandText = "SELECT CodParticipe,NumSolicitud,NumOperacion,NumCertificado " & _
                        "FROM ParticipePagoSuscripcion " & _
                        "WHERE NumSecuencial=" & adoRegistro("NumSecuencial") & " AND CodFondo='" & strCodFondo & "' AND " & _
                        "CodAdministradora='" & gstrCodAdministradora & "' AND NumOrdenCobroPago=''"
                    Set adoAuxiliar = .Execute
                    
                    If Not adoAuxiliar.EOF Then
                        Do While Not adoAuxiliar.EOF
                            
                            Set adoTemporal = New ADODB.Recordset
                            
                            .CommandText = "SELECT SUM(MontoTotal) MontoSuscripcion,SUM(CantCuotas) CuotasSuscripcion " & _
                                "FROM ParticipeOperacion " & _
                                "WHERE CodParticipe='" & adoAuxiliar("CodParticipe") & "' AND NumOperacion='" & adoAuxiliar("NumOperacion") & "' AND " & _
                                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                            Set adoTemporal = .Execute
                            
                            If Not adoTemporal.EOF Then
                                curMontoSuscripcion = adoTemporal("MontoSuscripcion")
                                dblCuotasSuscripcion = adoTemporal("CuotasSuscripcion")
                            End If
                            adoTemporal.Close: Set adoTemporal = Nothing
                        
                            If CDbl(adoRegistro("PorcenPago")) > 0 Then
                                dblCuotasPagadas = Round(dblCuotasSuscripcion * CDbl(adoRegistro("PorcenPago")) * 0.01, 5)
        
                                .CommandText = "UPDATE ParticipePagoSuscripcion SET " & _
                                    "MontoLiquidacion=" & curMontoSuscripcion * adoRegistro("PorcenPago") * 0.01 & "," & _
                                    "CantCuotasPagadas=" & dblCuotasPagadas & " " & _
                                    "WHERE CodParticipe='" & adoAuxiliar("CodParticipe") & "' AND NumSolicitud='" & adoAuxiliar("NumSolicitud") & "' AND " & _
                                    "NumSecuencial=" & adoRegistro("NumSecuencial") & " AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                            Else
                                dblCuotasPagadas = Round((CDbl(adoRegistro("MontoPago")) * CDbl(curMontoSuscripcion)) / dblCuotasSuscripcion, 5)
        
                                .CommandText = "UPDATE ParticipePagoSuscripcion SET " & _
                                    "MontoLiquidacion=" & CDec(adoRegistro("MontoPago")) & "," & _
                                    "CantCuotasPagadas=" & dblCuotasPagadas & " " & _
                                    "WHERE CodParticipe='" & adoAuxiliar("CodParticipe") & "' AND NumSolicitud='" & adoAuxiliar("NumSolicitud") & "' AND " & _
                                    "NumSecuencial=" & adoRegistro("NumSecuencial") & " AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                            End If
                            adoConn.Execute .CommandText
                            
                            adoAuxiliar.MoveNext
                        Loop
                    Else
                        adoAuxiliar.Close
                        
                        .CommandText = "SELECT CodParticipe,NumSolicitud,NumOperacion,NumCertificado " & _
                            "FROM ParticipePagoSuscripcionTmp " & _
                            "WHERE NumSecuencial=" & adoRegistro("NumSecuencial") & " AND CodFondo='" & strCodFondo & "' AND " & _
                            "CodAdministradora='" & gstrCodAdministradora & "' AND NumOrdenCobroPago=''"
                        Set adoAuxiliar = .Execute
                        
                        Do While Not adoAuxiliar.EOF
                            
                            Set adoTemporal = New ADODB.Recordset
                            
                            .CommandText = "SELECT SUM(MontoNetoSolictud) MontoSuscripcion,SUM(CantCuotas) CuotasSuscripcion " & _
                                "FROM ParticipeSolicitud " & _
                                "WHERE CodParticipe='" & adoAuxiliar("CodParticipe") & "' AND NumSolicitud='" & adoAuxiliar("NumSolicitud") & "' AND " & _
                                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                            Set adoTemporal = .Execute
                            
                            If Not adoTemporal.EOF Then
                                curMontoSuscripcion = adoTemporal("MontoSuscripcion")
                                dblCuotasSuscripcion = adoTemporal("CuotasSuscripcion")
                            End If
                            adoTemporal.Close: Set adoTemporal = Nothing
                        
                            If CDbl(adoRegistro("PorcenPago")) > 0 Then
                                dblCuotasPagadas = Round(dblCuotasSuscripcion * CDbl(adoRegistro("PorcenPago")) * 0.01, 5)
        
                                .CommandText = "UPDATE ParticipePagoSuscripcionTmp SET " & _
                                    "MontoLiquidacion=" & curMontoSuscripcion * adoRegistro("PorcenPago") * 0.01 & "," & _
                                    "CantCuotasPagadas=" & dblCuotasPagadas & " " & _
                                    "WHERE CodParticipe='" & adoAuxiliar("CodParticipe") & "' AND NumSolicitud='" & adoAuxiliar("NumSolicitud") & "' AND " & _
                                    "NumSecuencial=" & adoRegistro("NumSecuencial") & " AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                            Else
                                dblCuotasPagadas = Round((CDbl(adoRegistro("MontoPago")) * CDbl(curMontoSuscripcion)) / dblCuotasSuscripcion, 5)
        
                                .CommandText = "UPDATE ParticipePagoSuscripcionTmp SET " & _
                                    "MontoLiquidacion=" & CDec(adoRegistro("MontoPago")) & "," & _
                                    "CantCuotasPagadas=" & dblCuotasPagadas & " " & _
                                    "WHERE CodParticipe='" & adoAuxiliar("CodParticipe") & "' AND NumSolicitud='" & adoAuxiliar("NumSolicitud") & "' AND " & _
                                    "NumSecuencial=" & adoRegistro("NumSecuencial") & " AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                            End If
                            adoConn.Execute .CommandText
                            
                            adoAuxiliar.MoveNext
                        Loop
                    End If
                    adoAuxiliar.Close: Set adoAuxiliar = Nothing
                                        
                    adoRegistro.MoveNext
                Loop
                adoRegistro.Close: Set adoRegistro = Nothing
            End With
        
            Me.MousePointer = vbDefault
                            
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabPagos
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
    End If
    
    If strEstado = Reg_Edicion Then
        If TodoOK() Then
            frmMainMdi.stbMdi.Panels(3).Text = "Actualizar Cronograma de Pagos de Cuotas..."
            
            adoConsulta.MoveFirst
            dblPorcenTotal = 0
                                                
            Do While Not adoConsulta.EOF
                
                If adoConsulta.Fields("NumSecuencial") <> CInt(lblNumSecuencial.Caption) Then
                    dblPorcenTotal = dblPorcenTotal + CDbl(adoConsulta.Fields("PorcenPago"))
                End If
                
                adoConsulta.MoveNext
            Loop
            
            dblPorcenTotal = dblPorcenTotal + CDbl(txtPorcenPago.Text)
            
            If dblPorcenTotal > 100 Then
                MsgBox "El Porcentaje Total excede el 100%, verifique!", vbCritical, Me.Caption
                Exit Sub
            End If
            
            Me.MousePointer = vbHourglass
            
            With adoComm
                .CommandText = "SELECT FechaPago " & _
                    "FROM ParticipePagoSuscripcion " & _
                    "WHERE NumSecuencial=" & CInt(lblNumSecuencial.Caption) & " AND CodFondo='" & strCodFondo & "' AND " & _
                    "CodAdministradora='" & gstrCodAdministradora & "'"
                Set adoRegistro = .Execute
                
                If Not adoRegistro.EOF Then
                    If adoRegistro("FechaPago") <> Valor_Fecha Then
                        MsgBox "La Cuota ya fué pagada, no se puede modificar", vbCritical, Me.Caption
                        adoRegistro.Close: Set adoRegistro = Nothing
                        Me.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
                adoRegistro.Close
                    
                .CommandText = "{ call up_GNManFondoPagoSuscripcion('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    CInt(lblNumSecuencial.Caption) & "','" & _
                    Convertyyyymmdd(dtpFechaDesde.Value) & "','" & Convertyyyymmdd(dtpFechaHasta.Value) & "'," & _
                    CDec(txtPorcenPago.Text) & "," & CDec(txtMontoPago.Text) & ",'U') }"
                adoConn.Execute .CommandText
                
                '*** Actualizar valores en Cronograma de Pago de Partícipes ***
                .CommandText = "SELECT NumSecuencial,FechaDesde,FechaHasta,PorcenPago,MontoPago FROM FondoPagoSuscripcion " & _
                    "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' " & _
                    "ORDER BY NumSecuencial"
                Set adoRegistro = .Execute

                Do While Not adoRegistro.EOF
                    Set adoAuxiliar = New ADODB.Recordset
                    
                    datFechaPagoParcial = CVDate(adoRegistro("FechaHasta"))

                    .CommandText = "SELECT CodParticipe,NumSolicitud,NumOperacion,NumCertificado " & _
                        "FROM ParticipePagoSuscripcion " & _
                        "WHERE NumSecuencial=" & adoRegistro("NumSecuencial") & " AND CodFondo='" & strCodFondo & "' AND " & _
                        "CodAdministradora='" & gstrCodAdministradora & "' AND NumOrdenCobroPago=''"
                    Set adoAuxiliar = .Execute
                    
                    If Not adoAuxiliar.EOF Then
                        Do While Not adoAuxiliar.EOF
                            
                            Set adoTemporal = New ADODB.Recordset
                            
                            .CommandText = "SELECT SUM(MontoTotal) MontoSuscripcion,SUM(CantCuotas) CuotasSuscripcion " & _
                                "FROM ParticipeOperacion " & _
                                "WHERE CodParticipe='" & adoAuxiliar("CodParticipe") & "' AND NumOperacion='" & adoAuxiliar("NumOperacion") & "' AND " & _
                                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                            Set adoTemporal = .Execute
                            
                            If Not adoTemporal.EOF Then
                                curMontoSuscripcion = adoTemporal("MontoSuscripcion")
                                dblCuotasSuscripcion = adoTemporal("CuotasSuscripcion")
                            End If
                            adoTemporal.Close: Set adoTemporal = Nothing
                        
                            If CDbl(adoRegistro("PorcenPago")) > 0 Then
                                dblCuotasPagadas = Round(dblCuotasSuscripcion * CDbl(adoRegistro("PorcenPago")) * 0.01, 5)
        
                                .CommandText = "UPDATE ParticipePagoSuscripcion SET " & _
                                    "MontoLiquidacion=" & curMontoSuscripcion * adoRegistro("PorcenPago") * 0.01 & "," & _
                                    "CantCuotasPagadas=" & dblCuotasPagadas & " " & _
                                    "WHERE CodParticipe='" & adoAuxiliar("CodParticipe") & "' AND NumSolicitud='" & adoAuxiliar("NumSolicitud") & "' AND " & _
                                    "NumSecuencial=" & adoRegistro("NumSecuencial") & " AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                            Else
                                dblCuotasPagadas = Round((CDbl(adoRegistro("MontoPago")) * CDbl(curMontoSuscripcion)) / dblCuotasSuscripcion, 5)
        
                                .CommandText = "UPDATE ParticipePagoSuscripcion SET " & _
                                    "MontoLiquidacion=" & CDec(adoRegistro("MontoPago")) & "," & _
                                    "CantCuotasPagadas=" & dblCuotasPagadas & " " & _
                                    "WHERE CodParticipe='" & adoAuxiliar("CodParticipe") & "' AND NumSolicitud='" & adoAuxiliar("NumSolicitud") & "' AND " & _
                                    "NumSecuencial=" & adoRegistro("NumSecuencial") & " AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                            End If
                            adoConn.Execute .CommandText
                            
                            adoAuxiliar.MoveNext
                        Loop
                    Else
                        adoAuxiliar.Close
                        
                        .CommandText = "SELECT CodParticipe,NumSolicitud,NumOperacion,NumCertificado " & _
                            "FROM ParticipePagoSuscripcionTmp " & _
                            "WHERE NumSecuencial=" & adoRegistro("NumSecuencial") & " AND CodFondo='" & strCodFondo & "' AND " & _
                            "CodAdministradora='" & gstrCodAdministradora & "' AND NumOrdenCobroPago=''"
                        Set adoAuxiliar = .Execute
                        
                        Do While Not adoAuxiliar.EOF
                            
                            Set adoTemporal = New ADODB.Recordset
                            
                            .CommandText = "SELECT SUM(MontoNetoSolicitud) MontoSuscripcion,SUM(CantCuotas) CuotasSuscripcion " & _
                                "FROM ParticipeSolicitud " & _
                                "WHERE CodParticipe='" & adoAuxiliar("CodParticipe") & "' AND NumSolicitud='" & adoAuxiliar("NumSolicitud") & "' AND " & _
                                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                            Set adoTemporal = .Execute
                            
                            If Not adoTemporal.EOF Then
                                curMontoSuscripcion = adoTemporal("MontoSuscripcion")
                                dblCuotasSuscripcion = adoTemporal("CuotasSuscripcion")
                            End If
                            adoTemporal.Close: Set adoTemporal = Nothing
                        
                            If CDbl(adoRegistro("PorcenPago")) > 0 Then
                                dblCuotasPagadas = Round(dblCuotasSuscripcion * CDbl(adoRegistro("PorcenPago")) * 0.01, 5)
        
                                .CommandText = "UPDATE ParticipePagoSuscripcionTmp SET " & _
                                    "MontoLiquidacion=" & curMontoSuscripcion * adoRegistro("PorcenPago") * 0.01 & "," & _
                                    "CantCuotasPagadas=" & dblCuotasPagadas & " " & _
                                    "WHERE CodParticipe='" & adoAuxiliar("CodParticipe") & "' AND NumSolicitud='" & adoAuxiliar("NumSolicitud") & "' AND " & _
                                    "NumSecuencial=" & adoRegistro("NumSecuencial") & " AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                            Else
                                dblCuotasPagadas = Round((CDbl(adoRegistro("MontoPago")) * CDbl(curMontoSuscripcion)) / dblCuotasSuscripcion, 5)
        
                                .CommandText = "UPDATE ParticipePagoSuscripcionTmp SET " & _
                                    "MontoLiquidacion=" & CDec(adoRegistro("MontoPago")) & "," & _
                                    "CantCuotasPagadas=" & dblCuotasPagadas & " " & _
                                    "WHERE CodParticipe='" & adoAuxiliar("CodParticipe") & "' AND NumSolicitud='" & adoAuxiliar("NumSolicitud") & "' AND " & _
                                    "NumSecuencial=" & adoRegistro("NumSecuencial") & " AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                            End If
                            adoConn.Execute .CommandText
                            
                            adoAuxiliar.MoveNext
                        Loop
                    End If
                    adoAuxiliar.Close: Set adoAuxiliar = Nothing
                                        
                    adoRegistro.MoveNext
                Loop
                adoRegistro.Close: Set adoRegistro = Nothing
                
            End With
        
            Me.MousePointer = vbDefault
                            
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabPagos
                .TabEnabled(0) = True
                .TabEnabled(1) = False
                .Tab = 0
            End With
            Call Buscar
        End If
    End If
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

Private Function TodoOK() As Boolean
        
    Dim adoRegistro     As ADODB.Recordset
    Dim strFechaBase    As String
    
    TodoOK = False
            
    If CDbl(txtPorcenPago.Text) = 0 And CCur(txtMontoPago.Text) = 0 Then
        MsgBox "El Porcentaje o Monto de Pago no puede ser cero!.", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
        
    strFechaBase = Convertyyyymmdd(dtpFechaDesde.Value)
    
'    Set adoRegistro = New ADODB.Recordset
'    With adoComm
'        .CommandText = "SELECT NumSecuencial,FechaDesde,FechaHasta FROM FondoPagoSuscripcion " & _
'            "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'            "(FechaDesde>='" & strFechaBase & "' AND FechaHasta<='" & strFechaBase & "')"
'        Set adoRegistro = .Execute
'
'        If Not adoRegistro.EOF Then
'            MsgBox "Existe el rango de fecha " & CStr(adoRegistro("FechaDesde")) & "-" & CStr(adoRegistro("FechaHasta")), vbCritical, Me.Caption
'            Exit Function
'        End If
'        adoRegistro.Close: Set adoRegistro = Nothing
'    End With
    
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function
Public Sub Imprimir()

End Sub

Public Sub Buscar()
        
    Dim strSQL As String
    Set adoConsulta = New ADODB.Recordset
    
    strSQL = "SELECT NumSecuencial,FechaDesde,FechaHasta,PorcenPago,MontoPago " & _
        "FROM FondoPagoSuscripcion  " & _
        "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' " & _
        "ORDER BY NumSecuencial"
                        
    strEstado = Reg_Defecto
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
        
    tdgConsulta.DataSource = adoConsulta
    
    If adoConsulta.RecordCount > 0 Then strEstado = Reg_Consulta
    
End Sub

Public Sub Modificar()

    If strCodFondo = Valor_Caracter Then
        MsgBox "No existen fondos definidos...", vbCritical, Me.Caption
        Exit Sub
    End If
    
    frmMainMdi.stbMdi.Panels(3).Text = "Modificar..."
    
    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabPagos
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .Tab = 1
        End With
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro         As ADODB.Recordset
    Dim intNumSecuencial    As Integer
    
    Select Case strModo
        Case Reg_Adicion
            intNumSecuencial = adoConsulta.RecordCount + 1
            lblDescripFondo.Caption = Trim(cboFondo.Text)
            lblNumSecuencial.Caption = CStr(intNumSecuencial)
            dtpFechaDesde.Value = gdatFechaActual
            dtpFechaHasta.Value = gdatFechaActual
            txtPorcenPago.Text = "0"
            txtMontoPago.Text = "0"
            
            dtpFechaDesde.SetFocus
        
        Case Reg_Edicion
            Set adoRegistro = New ADODB.Recordset

            intNumSecuencial = CInt(tdgConsulta.Columns(0))

            adoComm.CommandText = "SELECT * FROM FondoPagoSuscripcion " & _
                "WHERE NumSecuencial=" & intNumSecuencial & " AND CodFondo='" & _
                strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            Set adoRegistro = adoComm.Execute

            If Not adoRegistro.EOF Then
                lblDescripFondo.Caption = Trim(cboFondo.Text)
                lblNumSecuencial.Caption = CStr(intNumSecuencial)
                
                dtpFechaDesde.Value = adoRegistro("FechaDesde")
                dtpFechaHasta.Value = adoRegistro("FechaHasta")
                                
                txtPorcenPago.Text = CStr(adoRegistro("PorcenPago"))
                txtMontoPago.Text = CStr(adoRegistro("MontoPago"))
                                                
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
    
    End Select
    
End Sub

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String

    If tabPagos.Tab = 1 Then Exit Sub

    Select Case Index
        Case 1
            gstrNameRepo = "FondoPagoSuscripcion"

            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(1)
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

            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora

    End Select

    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub
Public Sub Adicionar()
    
    If strCodFondo = Valor_Caracter Then
        MsgBox "No existen fondos definidos...", vbCritical, Me.Caption
        Exit Sub
    End If
    
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar..."
                
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabPagos
        .TabEnabled(0) = False
        .Tab = 1
    End With
                
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
    
    Call ValidarPermisoUsoControl(Trim(gstrLogin), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
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

Private Sub CargarListas()
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
            
End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabPagos.Tab = 0
    tabPagos.TabEnabled(1) = False
    
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 14
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 14
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 14
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 20
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 26
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub Form_Terminate()

    Dim adoRegistro     As ADODB.Recordset
    Dim dblPorcenTotal  As Double, curMontoTotal    As Currency

    'If Not adoConsulta.Enabled = False Then Exit Sub
    If adoConsulta.RecordCount = 0 Then Exit Sub
    
    adoConsulta.MoveFirst
    
    dblPorcenTotal = 0: curMontoTotal = 0

    Do While Not adoConsulta.EOF

        dblPorcenTotal = dblPorcenTotal + CDbl(adoConsulta.Fields("PorcenPago"))
        curMontoTotal = curMontoTotal + CDbl(adoConsulta.Fields("MontoPago"))

        adoConsulta.MoveNext
    Loop

    If dblPorcenTotal > 0 Then
        If dblPorcenTotal < 100 Then
            MsgBox "El Porcentaje Total es menor al 100%, verifique!", vbCritical, Me.Caption
            Me.Show
            Exit Sub
        End If
    End If
    
    If curMontoTotal > 0 Then
        Set adoRegistro = New ADODB.Recordset
        
        adoComm.CommandText = "SELECT MontoEmitido FROM Fondo WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
        Set adoRegistro = adoComm.Execute
        
        If Not adoRegistro.EOF Then
            If curMontoTotal < CCur(adoRegistro("MontoEmitido")) Then
                MsgBox "El Monto Total es menor al Monto Emitido, verifique!", vbCritical, Me.Caption
                Me.Show
                Exit Sub
            End If
        End If
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    Call OcultarReportes
    Set frmFondoPagoSuscripcion = Nothing
    
End Sub



Private Sub tabPagos_Click(PreviousTab As Integer)

    Select Case tabPagos.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabPagos.Tab = 0
        
    End Select
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
        
    If ColIndex = 3 Then
        Call DarFormatoValor(Value, Decimales_Tasa2)
    End If
    
    If ColIndex = 4 Then
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

Private Sub txtMontoPago_Change()

    Call FormatoCajaTexto(txtMontoPago, Decimales_Monto)
    
End Sub

Private Sub txtMontoPago_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtMontoPago, Decimales_Monto)
    
End Sub


Private Sub txtPorcenPago_Change()

    Call FormatoCajaTexto(txtPorcenPago, Decimales_Tasa2)
    
End Sub


Private Sub txtPorcenPago_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtPorcenPago, Decimales_Tasa2)
    
End Sub


