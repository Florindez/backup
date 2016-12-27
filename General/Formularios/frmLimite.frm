VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{442CAE95-1D41-47B1-BE83-6995DA3CE254}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmLimite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estructura de Limites de Inversión"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   Icon            =   "frmLimite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   10215
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   8160
      TabIndex        =   1
      Top             =   6000
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
      Left            =   840
      TabIndex        =   0
      Top             =   6000
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   1296
      Buttons         =   3
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Eliminar"
      Tag2            =   "4"
      ToolTipText2    =   "Eliminar"
      UserControlWidth=   4200
   End
   Begin TabDlg.SSTab tabLimite 
      Height          =   5655
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
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
      TabPicture(0)   =   "frmLimite.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgDetalle"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdgConsulta"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraCriterios"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmLimite.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraDetalle"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraLimite"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -68640
         TabIndex        =   3
         Top             =   4800
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
      Begin VB.Frame fraCriterios 
         Caption         =   "Criterios de búsqueda"
         Height          =   855
         Left            =   360
         TabIndex        =   16
         Top             =   480
         Width           =   9015
         Begin VB.ComboBox cboTipoReglamento 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   360
            Width           =   5775
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Reglamento"
            Height          =   195
            Index           =   4
            Left            =   480
            TabIndex        =   18
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame fraDetalle 
         Caption         =   "Detalle"
         Height          =   2295
         Left            =   -74640
         TabIndex        =   11
         Top             =   2400
         Width           =   9015
         Begin VB.TextBox txtDescripDetalle 
            Height          =   285
            Left            =   1680
            MaxLength       =   100
            TabIndex        =   15
            Top             =   1320
            Width           =   6975
         End
         Begin VB.Label lblCodLimite 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1680
            TabIndex        =   14
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   13
            Top             =   1320
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   12
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.Frame fraLimite 
         Caption         =   "Limite"
         Height          =   1815
         Left            =   -74640
         TabIndex        =   5
         Top             =   480
         Width           =   9015
         Begin VB.TextBox txtDescripEstructura 
            Height          =   285
            Left            =   1680
            MaxLength       =   70
            TabIndex        =   2
            Top             =   1065
            Width           =   6975
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   9
            Top             =   1080
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   8
            Top             =   420
            Width           =   495
         End
         Begin VB.Label lblCodEstructura 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1680
            TabIndex        =   6
            Top             =   405
            Width           =   2055
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmLimite.frx":0044
         Height          =   1800
         Left            =   360
         OleObjectBlob   =   "frmLimite.frx":005E
         TabIndex        =   7
         ToolTipText     =   "Seleccionar para el mantenimiento de limites"
         Top             =   1440
         Width           =   9015
      End
      Begin TrueOleDBGrid60.TDBGrid tdgDetalle 
         Bindings        =   "frmLimite.frx":310E
         Height          =   1800
         Left            =   360
         OleObjectBlob   =   "frmLimite.frx":3127
         TabIndex        =   10
         ToolTipText     =   "Seleccionar para el mantenimiento del detalle"
         Top             =   3360
         Width           =   9015
      End
   End
End
Attribute VB_Name = "frmLimite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrTipoReglamento()     As String, strCodTipoReglamento As String
Dim strCodEstructura        As String, strCodLimite         As String
Dim strCodEstado            As String, strEstadoDetalle     As String
Dim strEstado               As String, strSQL               As String
Dim adoConsulta             As ADODB.Recordset, adoDetalle  As ADODB.Recordset
Dim indSortAsc              As Boolean, indSortDesc         As Boolean

Private Sub BuscarDetalle(ByVal strpCodigo As String)

    Set adoDetalle = New ADODB.Recordset

    strSQL = "SELECT CodLimite,DescripLimite,IndTitulo,Estado " & _
        "FROM LimiteReglamentoEstructuraDetalle WHERE CodEstructura='" & strpCodigo & "' AND Estado='" & Estado_Activo & "' " & _
        "ORDER BY DescripLimite"
        
    strEstadoDetalle = Reg_Defecto
    With adoDetalle
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    tdgDetalle.DataSource = adoDetalle
    
    If adoDetalle.RecordCount > 0 Then strEstadoDetalle = Reg_Consulta
    
End Sub

Private Sub cboTipoReglamento_Click()

    strCodTipoReglamento = Valor_Caracter
    If cboTipoReglamento.ListIndex < 0 Then Exit Sub
    
    strCodTipoReglamento = Trim(arrTipoReglamento(cboTipoReglamento.ListIndex))
    
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

Public Sub Buscar()

    Set adoConsulta = New ADODB.Recordset
    
    strSQL = "SELECT CodEstructura, DescripEstructura FROM LimiteReglamentoEstructura " & _
        "WHERE CodReglamento='" & strCodTipoReglamento & "' AND Estado='" & Estado_Activo & "' " & _
        "ORDER BY DescripEstructura"
        
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
    
    Call tdgConsulta_Click
            
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Listado"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    
End Sub
Private Sub CargarListas()
    
    '*** Reglamentos ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPREG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoReglamento, arrTipoReglamento(), Valor_Caracter
    
    If cboTipoReglamento.ListCount > 0 Then cboTipoReglamento.ListIndex = 0
            
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

Public Sub Adicionar()
    
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Limites..."
                    
    strEstado = Reg_Adicion
    strEstadoDetalle = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabLimite
        .TabEnabled(0) = False
        .Tab = 1
        .TabEnabled(1) = True
    End With
    
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabLimite
        .TabEnabled(0) = True
        .Tab = 0
        .TabEnabled(1) = False
    End With
    Call Buscar
    
End Sub

Public Sub Eliminar()

    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        If MsgBox(Mensaje_Eliminacion, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            '*** Cambiar de Estado ***
            If fraLimite.Enabled Then
                adoComm.CommandText = "UPDATE LimiteReglamentoEstructura SET Estado='" & Estado_Eliminado & "' " & _
                    "WHERE CodEstructura='" & tdgConsulta.Columns(0).Value & "'"
            Else
                adoComm.CommandText = "UPDATE LimiteReglamentoEstructuraDetalle SET Estado='" & Estado_Eliminado & "' " & _
                    "WHERE CodEstructura='" & tdgConsulta.Columns(0).Value & "' AND CodLimite='" & tdgDetalle.Columns(0).Value & "'"
            End If
            adoConn.Execute adoComm.CommandText
            
            If fraDetalle.Enabled Then Call BuscarDetalle(tdgConsulta.Columns(0).Value)
            
            tabLimite.TabEnabled(0) = True
            tabLimite.Tab = 0
            Call Buscar
            
            Exit Sub
        End If
    End If
    
End Sub

Public Sub Grabar()

    Dim adoRegistro         As ADODB.Recordset, adoRec      As ADODB.Recordset
    Dim intAccion           As Integer, lngNumError         As Long
    Dim dblTipCambio        As Double
    Dim strFechaAnterior    As String, strFechaSiguiente    As String
    Dim strNumDetalleFile   As String
    Dim datFechaFinPeriodo  As Date
    
    If strEstado = Reg_Consulta Then Exit Sub
    If strEstadoDetalle = Reg_Consulta Then Exit Sub
    
    On Error GoTo CtrlError
    
    If strEstado = Reg_Adicion Or strEstadoDetalle = Reg_Adicion Then
        If TodoOK() Then
            Me.MousePointer = vbHourglass
            
            strCodEstructura = lblCodEstructura.Caption
            strCodLimite = lblCodLimite.Caption
            '*** Guardar ***
            With adoComm
                If fraLimite.Enabled Then
                    .CommandText = "INSERT INTO LimiteReglamentoEstructura VALUES('" & strCodEstructura & "','" & _
                        Trim(txtDescripEstructura.Text) & "','" & strCodTipoReglamento & "','" & Estado_Activo & "')"
                    adoConn.Execute .CommandText
                Else
                    .CommandText = "INSERT INTO LimiteReglamentoEstructuraDetalle VALUES('" & strCodLimite & "','" & _
                        Trim(txtDescripDetalle.Text) & "','" & strCodEstructura & "','','" & Estado_Activo & "')"
                    adoConn.Execute .CommandText
                End If
            End With
            
            Me.MousePointer = vbDefault
                        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabLimite
                .TabEnabled(0) = True
                .Tab = 0
            End With
            If fraLimite.Enabled Then
                Call Buscar
            Else
                Call BuscarDetalle(strCodEstructura)
            End If
        End If
    End If
    
    If strEstado = Reg_Edicion Or strEstadoDetalle = Reg_Edicion Then
        If TodoOK() Then
            Me.MousePointer = vbHourglass
            
            '*** Guardar ***
            With adoComm
                If fraLimite.Enabled Then
                    .CommandText = "UPDATE LimiteReglamentoEstructura SET DescripEstructura='" & Trim(txtDescripEstructura.Text) & "' " & _
                        "WHERE CodEstructura='" & strCodEstructura & "'"
                    adoConn.Execute .CommandText
                Else
                    .CommandText = "UPDATE LimiteReglamentoEstructuraDetalle SET DescripLimite='" & Trim(txtDescripDetalle.Text) & "' " & _
                        "WHERE CodEstructura='" & strCodEstructura & "' AND " & _
                        "CodLimite='" & strCodLimite & "'"
                    adoConn.Execute .CommandText
                End If
            End With
            
            Me.MousePointer = vbDefault
                        
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabLimite
                .TabEnabled(0) = True
                .Tab = 0
                .TabEnabled(1) = False
            End With
            If fraLimite.Enabled Then
                Call Buscar
            Else
                Call BuscarDetalle(strCodEstructura)
            End If
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
        
    TodoOK = False
        
    If fraLimite.Enabled Then
        If Trim(txtDescripEstructura.Text) = Valor_Caracter Then
            MsgBox "Debe indicar la descripción", vbCritical
            txtDescripEstructura.SetFocus
            Exit Function
        End If
    Else
        If Trim(txtDescripDetalle.Text) = Valor_Caracter Then
            MsgBox "Debe indicar la descripción", vbCritical
            txtDescripDetalle.SetFocus
            Exit Function
        End If
    End If
    
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function
Public Sub Imprimir()
    
    Call SubImprimir(1)
    
End Sub

Public Sub Modificar()

    If strEstado = Reg_Consulta Or strEstadoDetalle = Reg_Consulta Then
        strEstado = Reg_Edicion
        strEstadoDetalle = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabLimite
            .TabEnabled(0) = False
            .Tab = 1
            .TabEnabled(1) = True
        End With
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro     As ADODB.Recordset
    Dim intRegistro     As Integer
    
    Select Case strModo
        Case Reg_Adicion
            Set adoRegistro = New ADODB.Recordset
            
            If fraLimite.Enabled Then
                adoComm.CommandText = "SELECT COUNT(*) CodEstructura FROM LimiteReglamentoEstructura"
                Set adoRegistro = adoComm.Execute
                
                If Not adoRegistro.EOF Then
                    lblCodEstructura.Caption = Format(adoRegistro("CodEstructura") + 1, "00")
                Else
                    lblCodEstructura.Caption = Format(1, "00")
                End If
                adoRegistro.Close: Set adoRegistro = Nothing
    
                txtDescripEstructura.Text = Valor_Caracter
                
                txtDescripEstructura.SetFocus
            Else
                lblCodEstructura.Caption = Trim(tdgConsulta.Columns(0).Value)
                strCodEstructura = Trim(lblCodEstructura.Caption)
            
                txtDescripEstructura.Text = Trim(tdgConsulta.Columns(1).Value)
            
                adoComm.CommandText = "SELECT COUNT(*) CodLimite FROM LimiteReglamentoEstructuraDetalle"
                Set adoRegistro = adoComm.Execute
                
                If Not adoRegistro.EOF Then
                    lblCodLimite.Caption = Format(adoRegistro("CodLimite") + 1, "00")
                Else
                    lblCodLimite.Caption = Format(1, "00")
                End If
                adoRegistro.Close: Set adoRegistro = Nothing
    
                txtDescripDetalle.Text = Valor_Caracter
                
                txtDescripDetalle.SetFocus
            End If
        
        Case Reg_Edicion
            lblCodEstructura.Caption = Trim(tdgConsulta.Columns(0).Value)
            strCodEstructura = Trim(lblCodEstructura.Caption)
            
            txtDescripEstructura.Text = Trim(tdgConsulta.Columns(1).Value)
            
            lblCodLimite.Caption = Trim(tdgDetalle.Columns(0).Value)
            strCodLimite = Trim(lblCodLimite.Caption)
            txtDescripDetalle.Text = Trim(tdgDetalle.Columns(1).Value)
            
            If fraLimite.Enabled Then txtDescripEstructura.SetFocus
            If fraDetalle.Enabled Then txtDescripDetalle.SetFocus
            
    End Select
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String

    If tabLimite.Tab = 1 Then Exit Sub

    Select Case Index
        Case 1
            gstrNameRepo = "LimiteInversionEstructura"

            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(0)
            ReDim aReportParamFn(2)
            ReDim aReportParamF(2)

            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"

            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)

'            If strCodConcepto = Valor_Caracter Then
'                aReportParamS(0) = strCodConcepto
'                aReportParamS(1) = Codigo_Listar_Todos
'            Else
'                aReportParamS(0) = strCodConcepto
'                aReportParamS(1) = Codigo_Listar_Individual
'            End If

    End Select

    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    strEstadoDetalle = Reg_Defecto
    tabLimite.Tab = 0
    tabLimite.TabEnabled(1) = False
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 12
    tdgDetalle.Columns(2).Width = tdgDetalle.Width * 0.01 * 12
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmMoneda = Nothing
    
End Sub

Private Sub tabLimite_Click(PreviousTab As Integer)

    Select Case tabLimite.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabLimite.Tab = 0
        
    End Select
    
End Sub

Private Sub tdgConsulta_Click()

    tdgConsulta.HeadBackColor = &HFFC0C0
    tdgConsulta_GotFocus
    tdgDetalle.HeadBackColor = &H8000000F
    
End Sub

Private Sub tdgConsulta_GotFocus()

    fraLimite.Enabled = True
    fraDetalle.Enabled = False
    
End Sub

Private Sub tdgConsulta_SelChange(Cancel As Integer)

    Dim strCodBusqueda      As String

    If adoConsulta.RecordCount > 0 Then
        '*** Obtener el código del gasto ***
        strCodBusqueda = tdgConsulta.Columns(0).Value

        Call BuscarDetalle(strCodBusqueda)
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

Private Sub tdgDetalle_Click()

    tdgDetalle.HeadBackColor = &HFFC0C0
    tdgDetalle_GotFocus
    tdgConsulta.HeadBackColor = &H8000000F

End Sub

Private Sub tdgDetalle_GotFocus()

    fraLimite.Enabled = False
    fraDetalle.Enabled = True
    
End Sub

Private Sub tdgDetalle_HeadClick(ByVal ColIndex As Integer)
    
    Dim strColNameTDB  As String
    Static numColindex As Integer
    Static strPrevColumTDB As String
    '** agregar para que no se raye la seleccion de registro con ordenamiento
    strColNameTDB = tdgDetalle.Columns(ColIndex).DataField
    
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

    tdgDetalle.Splits(0).Columns(numColindex).HeadingStyle.ForegroundPicture = Null

    Call OrdenarDBGrid(ColIndex, adoDetalle, tdgDetalle)
    
    numColindex = ColIndex
    
    '****
    strPrevColumTDB = strColNameTDB
    '***
    
End Sub
