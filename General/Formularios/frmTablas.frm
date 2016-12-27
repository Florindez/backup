VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{442CAE95-1D41-47B1-BE83-6995DA3CE254}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmTablas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tablas Generales"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   7950
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   6240
      TabIndex        =   2
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
      Left            =   600
      TabIndex        =   1
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
      Caption2        =   "&Buscar"
      Tag2            =   "5"
      ToolTipText2    =   "Buscar"
      UserControlWidth=   4200
   End
   Begin TabDlg.SSTab tabCatalogo 
      Height          =   5655
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
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
      TabPicture(0)   =   "frmTablas.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTipoCambio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmTablas.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraDetalle"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAccion"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   4320
         TabIndex        =   4
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
      Begin VB.Frame fraTipoCambio 
         Caption         =   "Criterios de Búsqueda"
         Height          =   975
         Left            =   -74640
         TabIndex        =   9
         Top             =   480
         Width           =   6735
         Begin VB.ComboBox cboConcepto 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   4815
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tabla"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   10
            Top             =   375
            Width           =   405
         End
      End
      Begin VB.Frame fraDetalle 
         Height          =   4215
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   6705
         Begin VB.TextBox txtValores 
            Height          =   285
            Index           =   0
            Left            =   2280
            TabIndex        =   3
            Top             =   680
            Width           =   4095
         End
         Begin VB.Label lblDescripCampo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   12
            Top             =   700
            Width           =   495
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tabla"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   11
            Top             =   375
            Width           =   405
         End
         Begin VB.Label lblConcepto 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2280
            TabIndex        =   7
            Top             =   360
            Width           =   4095
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmTablas.frx":0038
         Height          =   3435
         Left            =   -74640
         OleObjectBlob   =   "frmTablas.frx":0052
         TabIndex        =   8
         Top             =   1590
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmTablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrConcepto()           As String
Dim strCodConcepto          As String, strEstado               As String
Dim adoConsulta             As ADODB.Recordset
Dim indSortAsc              As Boolean, indSortDesc             As Boolean

Private Sub cboConcepto_Click()

    strCodConcepto = Valor_Caracter
    If cboConcepto.ListIndex < 0 Then Exit Sub
    
    strCodConcepto = Trim(arrConcepto(cboConcepto.ListIndex))
    
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

    Dim strSQL As String
    
    Set adoConsulta = New ADODB.Recordset
        
    If strCodConcepto = Valor_Caracter Then Exit Sub
    
    strSQL = "SELECT * FROM " & strCodConcepto
        
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

    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Vista Activa"
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    
    
End Sub
Private Sub CargarListas()

    Dim strSQL As String, intRegistro As Integer
        
    '*** Tablas del Sistema ***
    strSQL = "SELECT CodTabla CODIGO,DescripTabla DESCRIP FROM TablaSistema ORDER BY DescripTabla"
    CargarControlLista strSQL, cboConcepto, arrConcepto(), Sel_Defecto
    
    If cboConcepto.ListCount > 0 Then cboConcepto.ListIndex = 0
            
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

    If cboConcepto.ListIndex = 0 Then Exit Sub
    
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Datos de Catálogo..."
                    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabCatalogo
        .TabEnabled(0) = False
        .Tab = 1
    End With
    
End Sub

Public Sub Cancelar()

    Dim intContador As Integer
    
    cmdOpcion.Visible = True
    With tabCatalogo
        .TabEnabled(0) = True
        .Tab = 0
    End With
    Call Buscar
    Call LimpiarControles
    
End Sub
Public Sub LimpiarControles()
    
    Dim intContador As Integer
    
    For intContador = 1 To (tdgConsulta.Columns.Count - 1)
        Unload txtValores(intContador)
        Unload lblDescripCampo(intContador)
    Next
    txtValores(0).Enabled = True


End Sub


Public Sub Eliminar()

End Sub

Public Sub Grabar()

    Dim adoresult As ADODB.Recordset, adoRec As ADODB.Recordset
    Dim strSQL As String, dblTipCambio As Double
    Dim strFechaAnterior As String, strFechaSiguiente As String
    Dim datFechaFinPeriodo As Date
    Dim strSeparadorCampo As String
    Dim strValores As String
    Dim strCampos As String
    Dim strCamposValores As String
    Dim intSecuencial As Long
    Dim intContador As Long
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    strCamposValores = ""
    
    If strEstado = Reg_Adicion Then
        If TodoOK() Then
            Me.MousePointer = vbHourglass
            
            Set adoresult = New ADODB.Recordset
            
            '*** Obtener el número secuencial ***
            With adoComm
                .CommandText = "SELECT MAX(" & Trim(lblDescripCampo(0).Caption) & ") AS NumSecuencial FROM " & strCodConcepto
                Set adoresult = .Execute
            End With
            
            If Not adoresult.EOF Then
                If IsNull(adoresult("NumSecuencial")) Then
                    intSecuencial = 1
                Else
                    intSecuencial = CInt(adoresult("NumSecuencial")) + 1
                End If
            Else
                intSecuencial = 1
            End If
            adoresult.Close: Set adoresult = Nothing
            
            txtValores(0).Text = Format(CStr(intSecuencial), "000")
            
            strSeparadorCampo = ","
            
            For intContador = 0 To (tdgConsulta.Columns.Count - 1)
                If intContador = (tdgConsulta.Columns.Count - 1) Then strSeparadorCampo = ""
                strValores = strValores & FormateaCampo(CStr(txtValores(intContador).Text), CInt(txtValores(intContador).Tag)) & strSeparadorCampo
                strCampos = strCampos & lblDescripCampo(intContador).Caption & strSeparadorCampo
            Next
           
            '*** Guardar ***
            With adoComm
                .CommandText = "{ call up_GNManCatalogoTabla('" & _
                    strCodConcepto & "','" & Trim(strCampos) & "','" & _
                    Trim(strValores) & "','" & _
                    Trim(strCamposValores) & "','" & _
                    "I'" & ") }"
                                
                adoConn.Execute .CommandText
                            
            End With
                                                                                                                        
            Me.MousePointer = vbDefault
                        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabCatalogo
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
            Call LimpiarControles
        End If
    End If
    
    If strEstado = Reg_Edicion Then
        If TodoOK() Then
            Me.MousePointer = vbHourglass
            
            
            strSeparadorCampo = ","
            
            strValores = FormateaCampo(CStr(txtValores(0).Text), CInt(txtValores(0).Tag))
            strCampos = lblDescripCampo(0).Caption
            
            For intContador = 1 To (tdgConsulta.Columns.Count - 1)
                If intContador = (tdgConsulta.Columns.Count - 1) Then strSeparadorCampo = ""
                strCamposValores = strCamposValores & lblDescripCampo(intContador).Caption & " = " & FormateaCampo(CStr(txtValores(intContador).Text), CInt(txtValores(intContador).Tag)) & strSeparadorCampo
            Next
            
            '*** Guardar ***
            With adoComm
                .CommandText = "{ call up_GNManCatalogoTabla('" & _
                    strCodConcepto & "','" & Trim(strCampos) & "','" & _
                    Trim(strValores) & "','" & _
                    Trim(strCamposValores) & "','" & _
                    "U'" & ") }"
                                                    
                adoConn.Execute .CommandText
                                            
            End With

            Me.MousePointer = vbDefault
                        
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabCatalogo
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
            Call LimpiarControles
        End If
    End If
    
    
End Sub

Private Function TodoOK() As Boolean
        
    Dim intContador As Integer
    
    TodoOK = False
        
    For intContador = 0 To txtValores.UBound
        If Trim(txtValores(intContador).Text) = Valor_Caracter Then
            MsgBox "Debe indicar el Código", vbCritical
            txtValores(intContador).SetFocus
            Exit Function
        End If
    Next
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
        With tabCatalogo
            .TabEnabled(0) = False
            .Tab = 1
        End With
        
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro     As ADODB.Recordset
    Dim intContador     As Integer
    
    Select Case strModo
        Case Reg_Adicion
            lblConcepto.Caption = Trim(cboConcepto.Text)
            
'            txtCodigo.Text = Valor_Caracter
'            txtCodigo.Enabled = True
'            txtAuxiliar.Text = Valor_Caracter
'            txtDescripParametro.Text = Valor_Caracter
'            txtNumValidacion.Text = "0"
'
'            txtCodigo.SetFocus
            
            For intContador = 0 To (tdgConsulta.Columns.Count - 1)
                
                If intContador > 0 Then Load txtValores(intContador)
                txtValores(intContador).Text = Valor_Caracter
                If intContador > 0 Then txtValores(intContador).Top = txtValores(intContador - 1).Top + 300
                txtValores(intContador).Visible = True
                txtValores(intContador).Tag = adoConsulta.Fields(intContador).Type
                
                If intContador > 0 Then Load lblDescripCampo(intContador)
                lblDescripCampo(intContador).Caption = tdgConsulta.Columns(intContador).Name
                If intContador > 0 Then lblDescripCampo(intContador).Top = lblDescripCampo(intContador - 1).Top + 300
                lblDescripCampo(intContador).Visible = True
                
            Next
            txtValores(0).Text = "[AUTOGENERADO]"
            txtValores(0).Enabled = False
        
        
        Case Reg_Edicion
            lblConcepto.Caption = Trim(cboConcepto.Text)
            
            For intContador = 0 To (tdgConsulta.Columns.Count - 1)
                
                If intContador > 0 Then Load txtValores(intContador)
                txtValores(intContador).Text = tdgConsulta.Columns(intContador).Value
                If intContador > 0 Then txtValores(intContador).Top = txtValores(intContador - 1).Top + 300
                txtValores(intContador).Visible = True
                txtValores(intContador).Tag = adoConsulta.Fields(intContador).Type
                                
                If intContador > 0 Then Load lblDescripCampo(intContador)
                lblDescripCampo(intContador).Caption = tdgConsulta.Columns(intContador).Name
                If intContador > 0 Then lblDescripCampo(intContador).Top = lblDescripCampo(intContador - 1).Top + 300
                lblDescripCampo(intContador).Visible = True
                
            Next
            txtValores(0).Enabled = False
    
    End Select
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub SubImprimir(Index As Integer)

    Dim strSeleccionRegistro    As String
    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()

    Set frmReporte = New frmVisorReporte
    
    ReDim aReportParamS(1)
    ReDim aReportParamFn(2)
    ReDim aReportParamF(2)

    aReportParamFn(0) = "Usuario"
    aReportParamFn(1) = "Hora"
    aReportParamFn(2) = "NombreEmpresa"

    aReportParamF(0) = gstrLogin
    aReportParamF(1) = Format(Time, "hh:mm:ss")
    aReportParamF(2) = gstrNombreEmpresa & Space(1)
    
    aReportParamS(1) = strCodConcepto

    Select Case Index
        Case 1
            If cboConcepto.ListIndex <= 0 Then
                MsgBox "Seleccione Tabla.", vbCritical
                Exit Sub
            End If
    End Select

    gstrNameRepo = "TablaSistema"
    gstrSelFrml = ""
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
        
End Sub
Private Sub InicializarValores()

    strEstado = Reg_Defecto
    tabCatalogo.Tab = 0
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmCatalogo = Nothing
    
End Sub

Private Sub fraDetalle_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub tabCatalogo_Click(PreviousTab As Integer)

    Select Case tabCatalogo.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabCatalogo.Tab = 0
        
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
