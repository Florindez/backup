VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{442CAE95-1D41-47B1-BE83-6995DA3CE254}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmCatalogo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogos"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   7950
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   6360
      TabIndex        =   2
      Top             =   4800
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
      TabIndex        =   1
      Top             =   4800
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   1296
      Buttons         =   4
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Eliminar"
      Tag2            =   "4"
      ToolTipText2    =   "Eliminar"
      Caption3        =   "&Buscar"
      Tag3            =   "5"
      ToolTipText3    =   "Buscar"
      UserControlWidth=   5700
   End
   Begin TabDlg.SSTab tabCatalogo 
      Height          =   4455
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7858
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
      TabPicture(0)   =   "frmCatalogo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTipoCambio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmCatalogo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraDetalle"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -70680
         TabIndex        =   6
         Top             =   3600
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
         Left            =   360
         TabIndex        =   12
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
            Caption         =   "Concepto"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   13
            Top             =   375
            Width           =   690
         End
      End
      Begin VB.Frame fraDetalle 
         Height          =   3135
         Left            =   -74640
         TabIndex        =   8
         Top             =   450
         Width           =   6735
         Begin VB.TextBox txtNumValidacion 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   20
            Top             =   2640
            Width           =   1335
         End
         Begin VB.TextBox txtAuxiliar 
            Height          =   285
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   4
            Top             =   1710
            Width           =   1335
         End
         Begin VB.TextBox txtDescripParametro 
            Height          =   285
            Left            =   1800
            MaxLength       =   60
            TabIndex        =   5
            Top             =   2160
            Width           =   4575
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   3
            Top             =   1260
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Num. Validación"
            Height          =   195
            Index           =   6
            Left            =   360
            TabIndex        =   19
            Top             =   2640
            Width           =   1155
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   18
            Top             =   2175
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Auxiliar"
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   17
            Top             =   1725
            Width           =   495
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   16
            Top             =   1275
            Width           =   495
         End
         Begin VB.Label lblDescrip 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   15
            Top             =   825
            Width           =   375
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Concepto"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   14
            Top             =   375
            Width           =   690
         End
         Begin VB.Label lblConcepto 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   10
            Top             =   360
            Width           =   4575
         End
         Begin VB.Label lblTipo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   9
            Top             =   810
            Width           =   1335
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmCatalogo.frx":0038
         Height          =   2595
         Left            =   360
         OleObjectBlob   =   "frmCatalogo.frx":0052
         TabIndex        =   11
         Top             =   1560
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmCatalogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrConcepto()           As String
Dim strCodConcepto          As String, strEstado               As String
Dim strSQL                  As String
Dim adoConsulta             As ADODB.Recordset
Dim indSortAsc              As Boolean, indSortDesc            As Boolean

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

    Set adoConsulta = New ADODB.Recordset
        
    strSQL = "SELECT * FROM AuxiliarParametro WHERE CodTipoParametro='" & strCodConcepto & "' AND Estado='" & Estado_Activo & "'"
        
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

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Vista Activa"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    
End Sub
Private Sub CargarListas()
    
    '*** Tipo de Conceptos ***
    strSQL = "SELECT CodTipoParametro CODIGO,DescripCatalogo DESCRIP FROM ConceptoCatalogo ORDER BY DescripCatalogo"
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
        .TabEnabled(1) = True
    End With
    
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabCatalogo
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
            adoComm.CommandText = "UPDATE AuxiliarParametro SET Estado='" & Estado_Eliminado & "' " & _
                "WHERE CodTipoParametro='" & strCodConcepto & "' AND CodParametro='" & tdgConsulta.Columns(0).Value & "'"
            adoConn.Execute adoComm.CommandText
            
            tabCatalogo.TabEnabled(0) = True
            tabCatalogo.Tab = 0
            Call Buscar
            
            Exit Sub
        End If
    End If
    
End Sub

Public Sub Grabar()

    Dim adoresult           As ADODB.Recordset, adoRec      As ADODB.Recordset
    Dim intAccion           As Integer, lngNumError         As Long
    Dim dblTipCambio        As Double
    Dim strFechaAnterior    As String, strFechaSiguiente    As String
    Dim datFechaFinPeriodo  As Date
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    On Error GoTo CtrlError
    
    If strEstado = Reg_Adicion Then
        If TodoOK() Then
            Me.MousePointer = vbHourglass
            
            '*** Guardar ***
            With adoComm
                .CommandText = "INSERT INTO AuxiliarParametro VALUES('" & _
                    strCodConcepto & "','" & Trim(txtCodigo.Text) & "','" & _
                    Trim(txtDescripParametro.Text) & "','" & _
                    Trim(txtAuxiliar.Text) & "'," & CInt(txtNumValidacion.Text) & ",'" & Estado_Activo & "')"
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
        End If
    End If
    
    If strEstado = Reg_Edicion Then
        If TodoOK() Then
            Me.MousePointer = vbHourglass
            
            '*** Guardar ***
            With adoComm
                .CommandText = "UPDATE AuxiliarParametro SET " & _
                    "DescripParametro='" & Trim(txtDescripParametro.Text) & "'," & _
                    "ValorParametro='" & Trim(txtAuxiliar.Text) & "'," & _
                    "NumValidacion=" & CInt(txtNumValidacion.Text) & " " & _
                    "WHERE CodTipoParametro='" & strCodConcepto & "' AND CodParametro='" & Trim(txtCodigo.Text) & "'"
                adoConn.Execute .CommandText
            End With

            Me.MousePointer = vbDefault
                        
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabCatalogo
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
        
    If Trim(txtCodigo.Text) = Valor_Caracter Then
        MsgBox "Debe indicar el Código", vbCritical
        txtCodigo.SetFocus
        Exit Function
    End If
    
    If Trim(txtDescripParametro.Text) = Valor_Caracter Then
        MsgBox "Debe indicar la descripción", vbCritical
        txtDescripParametro.SetFocus
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
        With tabCatalogo
            .TabEnabled(0) = False
            .Tab = 1
            .TabEnabled(1) = True
        End With
        
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Select Case strModo
        Case Reg_Adicion
            lblConcepto.Caption = Trim(cboConcepto.Text)
            lblTipo.Caption = strCodConcepto
            
            txtCodigo.Text = Valor_Caracter
            txtCodigo.Enabled = True
            txtAuxiliar.Text = Valor_Caracter
            txtDescripParametro.Text = Valor_Caracter
            txtNumValidacion.Text = "0"
            
            txtCodigo.SetFocus
        
        Case Reg_Edicion
            lblConcepto.Caption = Trim(cboConcepto.Text)
            lblTipo.Caption = Trim(tdgConsulta.Columns(3).Value)
            
            txtCodigo.Text = Trim(tdgConsulta.Columns(0).Value)
            txtCodigo.Enabled = False
            txtAuxiliar.Text = Trim(tdgConsulta.Columns(2).Value)
            txtDescripParametro.Text = Trim(tdgConsulta.Columns(1).Value)
            txtNumValidacion.Text = CInt(tdgConsulta.Columns(4).Value)
            txtAuxiliar.SetFocus
    
    End Select
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String

    If tabCatalogo.Tab = 1 Then Exit Sub

    Select Case Index
        Case 1
            gstrNameRepo = "Catalogo"

            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(1)
            ReDim aReportParamFn(2)
            ReDim aReportParamF(2)

            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"

            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)

            If strCodConcepto = Valor_Caracter Then
                aReportParamS(0) = strCodConcepto
                aReportParamS(1) = Codigo_Listar_Todos
            Else
                aReportParamS(0) = strCodConcepto
                aReportParamS(1) = Codigo_Listar_Individual
            End If

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
    tabCatalogo.Tab = 0
    tabCatalogo.TabEnabled(1) = False
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 14
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 60
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 10
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmCatalogo = Nothing
    
End Sub

Private Sub tabCatalogo_Click(PreviousTab As Integer)

    Select Case tabCatalogo.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabCatalogo.Tab = 0
        
    End Select
    
End Sub

Private Sub txtNumValidacion_Change()

    Call FormatoCajaTexto(txtNumValidacion, 0)
    
End Sub

Private Sub txtNumValidacion_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N")
    
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
