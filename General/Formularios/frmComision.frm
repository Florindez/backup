VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{442CAE95-1D41-47B1-BE83-6995DA3CE254}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmComision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comisiones"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   7950
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   6120
      TabIndex        =   1
      Top             =   4080
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
      Left            =   360
      TabIndex        =   0
      Top             =   4080
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
   Begin TabDlg.SSTab tabComision 
      Height          =   3735
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6588
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
      TabPicture(0)   =   "frmComision.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmComision.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDetalle"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAccion"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -70680
         TabIndex        =   3
         Top             =   2880
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
         Height          =   2295
         Left            =   -74640
         TabIndex        =   5
         Top             =   480
         Width           =   6735
         Begin VB.ComboBox cboEstado 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1680
            Width           =   2060
         End
         Begin VB.ComboBox cboTipoComision 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   810
            Width           =   2060
         End
         Begin VB.TextBox txtDescripComision 
            Height          =   285
            Left            =   2040
            MaxLength       =   20
            TabIndex        =   2
            Top             =   1305
            Width           =   4335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estado"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   11
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   10
            Top             =   1260
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   9
            Top             =   840
            Width           =   315
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
         Begin VB.Label lblCodComision 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2040
            TabIndex        =   6
            Top             =   400
            Width           =   2055
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmComision.frx":0038
         Height          =   2475
         Left            =   360
         OleObjectBlob   =   "frmComision.frx":0052
         TabIndex        =   7
         Top             =   600
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmComision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrTipoComision()       As String, strCodTipoComision   As String
Dim arrEstado()             As String, strCodEstado         As String
Dim strCodComision          As String
Dim strEstado               As String, strSQL               As String
Dim adoConsulta             As ADODB.Recordset
Dim indSortAsc              As Boolean, indSortDesc         As Boolean

Private Sub cboEstado_Click()

    strCodEstado = Valor_Caracter
    
    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strCodEstado = arrEstado(cboEstado.ListIndex)
    
End Sub

Private Sub cboTipoComision_Click()

    strCodTipoComision = Valor_Caracter
    
    If cboTipoComision.ListIndex < 0 Then Exit Sub
    
    strCodTipoComision = arrTipoComision(cboTipoComision.ListIndex)
    
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
    
    strSQL = "SELECT CodComision,CodTipoComision,DescripComision,COMEMP.Estado,DescripParametro DescripTipoComision " & _
        "FROM ComisionEmpresa COMEMP JOIN AuxiliarParametro AP ON(AP.CodParametro=COMEMP.CodTipoComision AND AP.CodTipoParametro='COMFON') " & _
        "ORDER BY CodTipoComision"
        
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

Private Sub CargarReportes()

'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Listado"
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    
End Sub
Private Sub CargarListas()
        
    '*** Tipo Comisión Administradora ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro " & _
        "WHERE CodTipoParametro='COMFON' AND Estado='" & Estado_Activo & "' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoComision, arrTipoComision(), Sel_Defecto
    
    If cboTipoComision.ListCount > 0 Then cboTipoComision.ListIndex = 0
    
    '*** Estado Registro ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro " & _
        "WHERE CodTipoParametro='ESTREG' AND CodParametro<>'" & Estado_Eliminado & "'  ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Valor_Caracter
    
    If cboEstado.ListCount > 0 Then cboEstado.ListIndex = 0
            
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
    
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Comisiones..."
                    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabComision
        .TabEnabled(0) = False
        .Tab = 1
        .TabEnabled(1) = True
    End With
    
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabComision
        .TabEnabled(0) = True
        .Tab = 0
        .TabEnabled(1) = False
    End With
    Call Buscar
    
End Sub

Public Sub Eliminar()

End Sub

Public Sub Grabar()

    Dim adoRegistro         As ADODB.Recordset
    Dim intAccion           As Integer, lngNumError         As Long
    Dim strNumDetalleFile   As String
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    On Error GoTo CtrlError
    
    If strEstado = Reg_Adicion Then
        If TodoOK() Then
            Me.MousePointer = vbHourglass
            
            '*** Guardar ***
            With adoComm
                Set adoRegistro = New ADODB.Recordset
                .CommandText = "SELECT COUNT(*) NumDetalle FROM InversionDetalleFile WHERE CodFile='098'"
                Set adoRegistro = .Execute
                
                If Not adoRegistro.EOF Then
                    If IsNull(adoRegistro("NumDetalle")) Then
                        strNumDetalleFile = "001"
                    Else
                        strNumDetalleFile = Format(CInt(adoRegistro("NumDetalle")) + 1, "000")
                    End If
                Else
                    strNumDetalleFile = "001"
                End If
                adoRegistro.Close: Set adoRegistro = Nothing
                
                .CommandText = "{ call up_GNManComisionEmpresa('" & _
                    strCodComision & "','" & strCodTipoComision & "','" & _
                    Trim(txtDescripComision.Text) & "','" & strNumDetalleFile & "','" & strCodEstado & "','I') }"
                adoConn.Execute .CommandText
                
                .CommandText = "INSERT INTO InversionDetalleFile VALUES('" & strNumDetalleFile & "','098','" & _
                        Trim(txtDescripComision.Text) & "','','X','')"
                    adoConn.Execute .CommandText
            End With
            
            Me.MousePointer = vbDefault
                        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabComision
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
    End If
    
    If strEstado = Reg_Edicion Then
        If TodoOK() Then
            Me.MousePointer = vbHourglass
            
            '*** Guardar ***
            With adoComm
                .CommandText = "{ call up_GNManComisionEmpresa('" & _
                    strCodComision & "','" & strCodTipoComision & "','" & _
                    Trim(txtDescripComision.Text) & "','" & strNumDetalleFile & "','" & strCodEstado & "','U') }"
                adoConn.Execute .CommandText
                
                .CommandText = "UPDATE InversionDetalleFile SET DescripDetalleFile='" & Trim(txtDescripComision.Text) & "' " & _
                        "WHERE CodFile='098' AND CodDetalleFile='0" & strCodComision & "'"
                    adoConn.Execute .CommandText
            End With
            
            Me.MousePointer = vbDefault
                        
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabComision
                .TabEnabled(0) = True
                .Tab = 0
                .TabEnabled(1) = False
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
        
    TodoOK = False
        
    If Trim(txtDescripComision.Text) = Valor_Caracter Then
        MsgBox "Debe indicar la descripción", vbCritical
        txtDescripComision.SetFocus
        Exit Function
    End If
    
    '*** Si todo paso OK ***
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
        With tabComision
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
            
            adoComm.CommandText = "SELECT COUNT(*) SecComision FROM ComisionEmpresa"
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                lblCodComision.Caption = Format(adoRegistro("SecComision") + 1, "00")
            Else
                lblCodComision.Caption = Format(1, "00")
            End If
            adoRegistro.Close: Set adoRegistro = Nothing

            strCodComision = Trim(lblCodComision.Caption)
            txtDescripComision.Text = Valor_Caracter
            
            intRegistro = ObtenerItemLista(arrEstado(), Estado_Activo)
            If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
            cboEstado.Enabled = False
            
            cboTipoComision.SetFocus
        
        Case Reg_Edicion
            lblCodComision.Caption = Trim(tdgConsulta.Columns(1).Value)
            strCodComision = Trim(lblCodComision.Caption)
            
            Call CargarListas
            
            txtDescripComision.Text = Trim(tdgConsulta.Columns(3).Value)
            
            intRegistro = ObtenerItemLista(arrTipoComision(), tdgConsulta.Columns(2).Value)
            If intRegistro >= 0 Then cboTipoComision.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrEstado(), tdgConsulta.Columns(4).Value)
            If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
            cboEstado.Enabled = True
            
            cboTipoComision.SetFocus
    
    End Select
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String

    If tabComision.Tab = 1 Then Exit Sub

    Select Case Index
        Case 1
            gstrNameRepo = "ComisionEmpresa"

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
    tabComision.Tab = 0
    tabComision.TabEnabled(1) = False
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 24
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 12
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmMoneda = Nothing
    
End Sub

Private Sub tabComision_Click(PreviousTab As Integer)

    Select Case tabComision.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabComision.Tab = 0
        
    End Select
    
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
