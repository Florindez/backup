VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{442CAE95-1D41-47B1-BE83-6995DA3CE254}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmTasas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tasas"
   ClientHeight    =   5970
   ClientLeft      =   1380
   ClientTop       =   900
   ClientWidth     =   6240
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
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5970
   ScaleWidth      =   6240
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   4440
      TabIndex        =   5
      Top             =   5160
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
      TabIndex        =   4
      Top             =   5160
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1296
      Buttons         =   2
      Caption0        =   "&Modificar"
      Tag0            =   "3"
      ToolTipText0    =   "Modificar"
      Caption1        =   "&Buscar"
      Tag1            =   "5"
      ToolTipText1    =   "Buscar"
      UserControlWidth=   2700
   End
   Begin TabDlg.SSTab tabTasas 
      Height          =   4815
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmTasas.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTipoCambio"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdgConsulta"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmTasas.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraDetalle"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -72360
         TabIndex        =   8
         Top             =   3960
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
      Begin VB.Frame fraDetalle 
         Height          =   3375
         Left            =   -74760
         TabIndex        =   14
         Top             =   480
         Width           =   5280
         Begin VB.TextBox txtValorTasa 
            Alignment       =   1  'Right Justify
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
            Left            =   1440
            TabIndex        =   7
            Top             =   2760
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker dtpFechaTipoCambio 
            Height          =   285
            Left            =   1440
            TabIndex        =   6
            Top             =   2160
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
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
            CurrentDate     =   38806
         End
         Begin VB.Label lblClase 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   1440
            TabIndex        =   26
            Top             =   840
            Width           =   3495
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clase"
            Height          =   195
            Index           =   9
            Left            =   360
            TabIndex        =   25
            Top             =   840
            Width           =   480
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mes"
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   22
            Top             =   1335
            Width           =   360
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            X1              =   360
            X2              =   5000
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
            Height          =   195
            Index           =   7
            Left            =   360
            TabIndex        =   21
            Top             =   2760
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            Height          =   195
            Index           =   6
            Left            =   360
            TabIndex        =   20
            Top             =   2160
            Width           =   540
         End
         Begin VB.Label lblPeriodo 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   3600
            TabIndex        =   19
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblMes 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   1440
            TabIndex        =   18
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblTipo 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   1440
            TabIndex        =   17
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año"
            Height          =   195
            Index           =   5
            Left            =   3120
            TabIndex        =   16
            Top             =   1335
            Width           =   345
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   15
            Top             =   375
            Width           =   390
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmTasas.frx":0038
         Height          =   1875
         Left            =   240
         OleObjectBlob   =   "frmTasas.frx":0052
         TabIndex        =   3
         Top             =   2415
         Width           =   5280
      End
      Begin VB.Frame fraTipoCambio 
         Caption         =   "Criterios de Búsqueda"
         Height          =   1815
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   5280
         Begin VB.ComboBox cboClaseTasa 
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
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   720
            Width           =   3735
         End
         Begin VB.ComboBox cboTipoTasa 
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
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   3735
         End
         Begin VB.ComboBox cboMes 
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
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   1125
            Width           =   1455
         End
         Begin VB.ComboBox cboAnio 
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
            Height          =   315
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1125
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clase"
            Height          =   195
            Index           =   8
            Left            =   360
            TabIndex        =   23
            Top             =   720
            Width           =   480
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   13
            Top             =   375
            Width           =   390
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mes"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   12
            Top             =   1140
            Width           =   360
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año"
            Height          =   195
            Index           =   0
            Left            =   2880
            TabIndex        =   11
            Top             =   1140
            Width           =   345
         End
      End
   End
End
Attribute VB_Name = "frmTasas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Mantenimiento de Tasas VAC"
Option Explicit

Dim strTipoTasa         As String, arrTipoTasa()        As String
Dim strAnio             As String, arrAnio()            As String
Dim strMes              As String, arrMes()             As String
Dim strClaseTasa        As String, arrClaseTasa()       As String
Dim strDiaInicial       As String, strDiaFinal          As String
Dim strFechaDesde       As String, strFechaHasta        As String
Dim strEstado           As String
Dim adoConsulta         As ADODB.Recordset
Dim indSortAsc          As Boolean, indSortDesc         As Boolean

Public Sub Adicionar()

End Sub

Public Sub Buscar()

    Dim strSQL As String
    Dim intmes As Integer, intAnio As Integer
    Dim intTemporal As Integer
    Dim datFechaInicioMes As Date, datFechaFinMes As Date
    Dim datFechaTemporal As Date
    
    Set adoConsulta = New ADODB.Recordset
    
    intmes = CInt(strMes)
    intAnio = CInt(strAnio)
    
    If intmes = 1 Then
        intTemporal = UltimoDiaMes(12, intAnio - 1)
        datFechaTemporal = Convertddmmyyyy(Format(intAnio - 1, "0000") & Format(12, "00") & Format(intTemporal, "00"))
    Else
        intTemporal = UltimoDiaMes(intmes - 1, intAnio)
        datFechaTemporal = Convertddmmyyyy(Format(intAnio, "0000") & Format(intmes - 1, "00") & Format(intTemporal, "00"))
    End If
    datFechaInicioMes = DateAdd("d", 1, datFechaTemporal)
    strDiaInicial = CStr(datFechaInicioMes)
    
    intTemporal = UltimoDiaMes(intmes, intAnio)
    datFechaTemporal = Convertddmmyyyy(Format(intAnio, "0000") & Format(intmes, "00") & Format(intTemporal, "00"))
    datFechaFinMes = DateAdd("d", 1, datFechaTemporal)
    strDiaFinal = CStr(datFechaTemporal)
            
    strSQL = "{ call up_GNGenFechasTasas('" & strTipoTasa & "','" & strClaseTasa & "','" & Convertyyyymmdd(datFechaInicioMes) & "','" & Convertyyyymmdd(datFechaFinMes) & "' ) }"
    adoConn.Execute strSQL
    
    strSQL = "SELECT FechaRegistro,ValorTasa " & _
        "FROM InversionTasaTemporal " & _
        "ORDER BY FechaRegistro"
        
    strEstado = Reg_Consulta
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

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabTasas
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .Tab = 0
    End With
    strEstado = Reg_Consulta
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Vista Activa"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Por Rango de Fecha"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    
End Sub



Public Sub Eliminar()

End Sub


Public Sub Grabar()

    Dim adoRegistro As ADODB.Recordset
    Dim intRegistro As Integer, intAccion           As Integer
    Dim intNumError As Integer
    Dim strFecha    As String, strFechaSiguiente    As String
        
    Set adoRegistro = New ADODB.Recordset
    
    On Error GoTo CtrlError
    
    If TodoOK() Then
        With adoComm
            '*** Actualizar valor de cambio en tabla temporal ***
            .CommandText = "UPDATE InversionTasaTemporal SET " & _
                "ValorTasa=" & CDec(txtValorTasa.Text) & " " & _
                "WHERE (FechaRegistro>='" & strFechaDesde & "'AND FechaRegistro<'" & strFechaHasta & "')"
            adoConn.Execute .CommandText
                
            '*** Registro del valor de cambio en tabla definitiva ***
            .CommandText = "SELECT * FROM InversionTasaTemporal"
            Set adoRegistro = .Execute
        
            Do While Not adoRegistro.EOF
                strFecha = Convertyyyymmdd(adoRegistro("FechaRegistro"))
                strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, adoRegistro("FechaRegistro")))
            
                .CommandText = "UPDATE InversionTasa SET " & _
                    "ValorTasa=" & CDbl(adoRegistro("ValorTasa")) & " " & _
                    "WHERE (FechaRegistro>='" & strFecha & "'AND FechaRegistro<'" & strFechaSiguiente & "') AND " & _
                    "CodClaseTasa='" & strClaseTasa & "' AND CodTasa='" & strTipoTasa & "'"
                adoConn.Execute .CommandText, intRegistro
                
                If intRegistro = 0 And CDbl(adoRegistro("ValorTasa")) > 0 Then
                    .CommandText = "INSERT INTO InversionTasa VALUES ('" & strTipoTasa & "','" & strClaseTasa & "','" & _
                        Convertyyyymmdd(adoRegistro("FechaRegistro")) & "'," & _
                        CDbl(adoRegistro("ValorTasa")) & ")"
                    adoConn.Execute .CommandText
                End If
                
                adoRegistro.MoveNext
            Loop
            adoRegistro.Close: Set adoRegistro = Nothing
        End With
        
        MsgBox Mensaje_Edicion_Exitosa, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
                
        cmdOpcion.Visible = True
        With tabTasas
            .TabEnabled(0) = True
            .Tab = 0
        End With
        Call Buscar
    End If
    Exit Sub
    
CtrlError:
    intAccion = ControlErrores
    Select Case intAccion
        Case 0: Resume
        Case 1: Resume Next
        Case Else
            intNumError = err.Number
            err.Raise Number:=intNumError
            err.Clear
    End Select
        
End Sub


Private Function TodoOK() As Boolean
        
    TodoOK = False
    
    If CDbl(txtValorTasa.Text) = 0 Then
        MsgBox "Debe ingresar el valor de la tasa.", vbCritical, Me.Caption
        txtValorTasa.SetFocus
        Exit Function
    End If
        
    '*** Si todo pasó OK ***
    TodoOK = True
  
End Function
Public Sub Imprimir()
    
    Call SubImprimir(1)
    
End Sub

Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabTasas
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .Tab = 1
        End With
        
    End If
    
End Sub


Private Sub LlenarFormulario(strModo As String)

    On Error GoTo CtrlError            '/**/ HMC Habilitamos la rutina de Errores Existente
    

    Dim adoRegistro   As ADODB.Recordset
    Dim intAccion As Integer, lngNumError   As Long
    
    Select Case strModo
        Case Reg_Edicion
            Set adoRegistro = New ADODB.Recordset
            
            adoComm.CommandText = "SELECT MAX(FechaRegistro) FechaRegistro FROM InversionTasaTemporal"
            Set adoRegistro = adoComm.Execute
            
            dtpFechaTipoCambio.MaxDate = "01/01/9999" 'HMC
            dtpFechaTipoCambio.MinDate = "01/01/1000" 'HMC
            
            If Not adoRegistro.EOF Then
                dtpFechaTipoCambio.MaxDate = adoRegistro("FechaRegistro")
            End If
            adoRegistro.Close
            
            adoComm.CommandText = "SELECT MIN(FechaRegistro) FechaRegistro FROM InversionTasaTemporal"
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                dtpFechaTipoCambio.MinDate = adoRegistro("FechaRegistro")
            End If
            adoRegistro.Close
            
            strFechaDesde = Convertyyyymmdd(tdgConsulta.Columns(0))
            strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, tdgConsulta.Columns(0)))
            
            adoComm.CommandText = "SELECT FechaRegistro,ValorTasa FROM InversionTasaTemporal " & _
                "WHERE FechaRegistro >='" & strFechaDesde & "' AND FechaRegistro <'" & strFechaHasta & "'"
                
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                lblTipo.Caption = Trim(cboTipoTasa.Text)
                lblClase.Caption = Trim(cboClaseTasa.Text)
                lblMes.Caption = Trim(cboMes.Text)
                lblPeriodo.Caption = Trim(cboAnio.Text)
                
                dtpFechaTipoCambio.Value = adoRegistro("FechaRegistro")
                txtValorTasa.Text = CStr(adoRegistro("ValorTasa"))
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
    
    End Select
    Exit Sub
    
CtrlError:                                  '/**/
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


Public Sub Salir()

    Unload Me
    
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vModify
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



Public Sub SubImprimir(Index As Integer)

    Dim strSeleccionRegistro    As String
    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()

    Set frmReporte = New frmVisorReporte
    
    ReDim aReportParamS(2)
    ReDim aReportParamFn(5)
    ReDim aReportParamF(5)

    aReportParamFn(0) = "Usuario"
    aReportParamFn(1) = "FechaDel"
    aReportParamFn(2) = "FechaAl"
    aReportParamFn(3) = "Hora"
    aReportParamFn(4) = "DescripTasa"
    aReportParamFn(5) = "NombreEmpresa"

    aReportParamF(0) = gstrLogin
    aReportParamF(3) = Format(Time, "hh:mm:ss")
    aReportParamF(4) = Trim(cboTipoTasa.Text)
    aReportParamF(5) = gstrNombreEmpresa & Space(1)
    
    aReportParamS(0) = strTipoTasa

    Select Case Index
        Case 1
            If cboTipoTasa.ListIndex < 0 Then
                MsgBox "Seleccione Tipo.", vbCritical
                Exit Sub
            End If

            If cboMes.ListIndex < 0 Then
                MsgBox "Seleccione Mes.", vbCritical
                Exit Sub
            End If

            If cboAnio.ListIndex < 0 Then
                MsgBox "Seleccione Año.", vbCritical
                Exit Sub
            End If
            aReportParamF(1) = strDiaInicial
            aReportParamF(2) = strDiaFinal
            
            aReportParamS(1) = Convertyyyymmdd(CVDate(strDiaInicial))
            aReportParamS(2) = Convertyyyymmdd(DateAdd("d", 1, CVDate(strDiaFinal)))

        Case 2
            If cboTipoTasa.ListIndex < 0 Then
                MsgBox "Seleccione Tipo.", vbCritical
                Exit Sub
            End If

            strSeleccionRegistro = "{InversionTasa.FechaRegistro} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal
            If gstrSelFrml <> "0" Then
                Me.Refresh
                aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
                aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                
                aReportParamS(1) = Convertyyyymmdd(CVDate(aReportParamF(1)))
                aReportParamS(2) = Convertyyyymmdd(DateAdd("d", 1, CVDate(aReportParamF(2))))
            Else
                Exit Sub
            End If
    End Select

    gstrNameRepo = "Tasas"
    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub



Private Sub cboAnio_Click()

    strAnio = "0000"
    If cboAnio.ListIndex < 0 Then Exit Sub
    
    strAnio = Trim(arrAnio(cboAnio.ListIndex))
    Me.Refresh
    
    If strMes = "00" Then Exit Sub
    
    Call Buscar
'    Me.MousePointer = vbHourglass
'    LSubLimGrd
'    DoCalendar CInt(strMes), CInt(strAnio)
'    Me.MousePointer = vbDefault
    
End Sub


Private Sub cboClaseTasa_Click()

    strClaseTasa = Valor_Caracter
    If cboClaseTasa.ListIndex < 0 Then Exit Sub
    
    strClaseTasa = Trim(arrClaseTasa(cboClaseTasa.ListIndex))
    
    If strClaseTasa = Valor_Caracter Then strClaseTasa = "00"
    
    cboMes_Click
    
End Sub


Private Sub cboMes_Click()

    strMes = "00"
    If cboMes.ListIndex < 0 Then Exit Sub
        
    strMes = arrMes(cboMes.ListIndex)
    Me.Refresh
    
    If strAnio = "0000" Then Exit Sub
    
    Call Buscar
        
End Sub


Private Sub cboTipoTasa_Click()

    Dim strSQL As String
    
    strTipoTasa = Valor_Caracter
    If cboTipoTasa.ListIndex < 0 Then Exit Sub
    
    strTipoTasa = Trim(arrTipoTasa(cboTipoTasa.ListIndex))
    
    '*** Clase de Tasa ***
    strSQL = "SELECT CodClaseTasa CODIGO,DescripClaseTasa DESCRIP FROM ClaseTasa " & _
        "WHERE CodTasa='" & strTipoTasa & "' ORDER BY DescripClaseTasa"
    CargarControlLista strSQL, cboClaseTasa, arrClaseTasa(), Valor_Caracter
    
    If cboClaseTasa.ListCount > 0 Then
        cboClaseTasa.ListIndex = 0
    Else
        strClaseTasa = "00"
    End If
    
    cboMes_Click
    
End Sub


Private Sub dtpFechaTipoCambio_Change()

    Dim adoRegistro     As ADODB.Recordset
    
    strFechaDesde = Convertyyyymmdd(dtpFechaTipoCambio.Value)
    strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaTipoCambio.Value))
    
    Set adoRegistro = New ADODB.Recordset
    
    adoComm.CommandText = "SELECT ValorTasa FROM InversionTasaTemporal " & _
        "WHERE FechaRegistro >='" & strFechaDesde & "' AND FechaRegistro <'" & strFechaHasta & "'"
        
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        txtValorTasa.Text = CStr(adoRegistro("ValorTasa"))
    End If
    adoRegistro.Close: Set adoRegistro = Nothing
            
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
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    txtValorTasa.Text = "0"
    tabTasas.Tab = 0
    tabTasas.TabEnabled(1) = False
    strMes = "00"
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 30
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 50
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub
Private Sub CargarListas()

    Dim strSQL As String, intRegistro As Integer
        
    '*** Tipo de Tasa ***
    strSQL = "SELECT CodTasa CODIGO,DescripTasa DESCRIP FROM TipoTasa ORDER BY DescripTasa"
    CargarControlLista strSQL, cboTipoTasa, arrTipoTasa(), Valor_Caracter
    
    intRegistro = ObtenerItemLista(arrTipoTasa(), Codigo_Tipo_Ajuste_Vac)
    If intRegistro >= 0 Then cboTipoTasa.ListIndex = intRegistro
            
    '*** Meses ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='DSCMES' ORDER BY CodParametro"
    CargarControlLista strSQL, cboMes, arrMes(), Valor_Caracter
    
    '*** Años ***
    strSQL = "SELECT ValorParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='RNGANI' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboAnio, arrAnio(), Valor_Caracter
                        
    If gstrPeriodoActual = Valor_Caracter Then
        intRegistro = ObtenerItemLista(arrAnio(), Format(Year(Date), "0000"))
        If intRegistro >= 0 Then cboAnio.ListIndex = intRegistro
        
        intRegistro = ObtenerItemLista(arrMes(), Format(Month(Date), "00"))
        If intRegistro >= 0 Then cboMes.ListIndex = intRegistro
    Else
        intRegistro = ObtenerItemLista(arrAnio(), gstrPeriodoActual)
        If intRegistro >= 0 Then cboAnio.ListIndex = intRegistro
        
        intRegistro = ObtenerItemLista(arrMes(), gstrMesActual)
        If intRegistro >= 0 Then cboMes.ListIndex = intRegistro
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmTasas = Nothing
        
End Sub

Private Sub tabTasas_Click(PreviousTab As Integer)

    Select Case tabTasas.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabTasas.Tab = 0
        
    End Select
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 1 Then
        Call DarFormatoValor(Value, Decimales_Tasa)
    End If
    
End Sub

Private Sub txtValorTasa_Change()

    Call FormatoCajaTexto(txtValorTasa, Decimales_Tasa)
    
End Sub

Private Sub txtValorTasa_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtValorTasa, Decimales_Tasa)
    
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
