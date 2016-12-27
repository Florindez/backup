VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmPoliticaInversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Limites de Inversión"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   10665
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   8880
      TabIndex        =   4
      Top             =   5400
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
      Left            =   720
      TabIndex        =   3
      Top             =   5400
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
   Begin TabDlg.SSTab tabPoliticaInversion 
      Height          =   5055
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8916
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
      TabPicture(0)   =   "frmPoliticaInversion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCriterios"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdgConsulta"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmPoliticaInversion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(1)=   "fraDetalle"
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -67920
         TabIndex        =   7
         Top             =   4200
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
         Height          =   3615
         Left            =   -74760
         TabIndex        =   14
         Top             =   480
         Width           =   9735
         Begin VB.ComboBox cboCriterio 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   2140
            Width           =   6975
         End
         Begin VB.TextBox txtPorcenMaximo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            MaxLength       =   40
            TabIndex        =   6
            Top             =   3000
            Width           =   1460
         End
         Begin VB.TextBox txtPorcenMinimo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            MaxLength       =   40
            TabIndex        =   5
            Top             =   2585
            Width           =   1460
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Criterio Comparación"
            Height          =   195
            Index           =   9
            Left            =   360
            TabIndex        =   25
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label lblConcepto 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2400
            TabIndex        =   24
            Top             =   1725
            Width           =   6975
         End
         Begin VB.Label lblLimite 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2400
            TabIndex        =   23
            Top             =   1305
            Width           =   6975
         End
         Begin VB.Label lblTipoReglamento 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2400
            TabIndex        =   22
            Top             =   480
            Width           =   6975
         End
         Begin VB.Label lblDescripFondo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2400
            TabIndex        =   21
            Top             =   900
            Width           =   6975
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "% Máximo"
            Height          =   195
            Index           =   8
            Left            =   360
            TabIndex        =   20
            Top             =   3020
            Width           =   705
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "% Mínimo"
            Height          =   195
            Index           =   7
            Left            =   360
            TabIndex        =   19
            Top             =   2605
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
            Height          =   195
            Index           =   6
            Left            =   360
            TabIndex        =   18
            Top             =   1745
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Limite"
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   17
            Top             =   1325
            Width           =   405
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   16
            Top             =   920
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Reglamento"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   15
            Top             =   500
            Width           =   1215
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmPoliticaInversion.frx":0038
         Height          =   2175
         Left            =   240
         OleObjectBlob   =   "frmPoliticaInversion.frx":0052
         TabIndex        =   10
         Top             =   2280
         Width           =   9735
      End
      Begin VB.Frame fraCriterios 
         Caption         =   "Criterios de búsqueda"
         Height          =   1695
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   9735
         Begin VB.ComboBox cboLimite 
            Height          =   315
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1080
            Width           =   5775
         End
         Begin VB.ComboBox cboTipoReglamento 
            Height          =   315
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   720
            Width           =   5775
         End
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   5775
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Limite"
            Height          =   195
            Index           =   2
            Left            =   840
            TabIndex        =   13
            Top             =   1080
            Width           =   405
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Reglamento"
            Height          =   195
            Index           =   1
            Left            =   840
            TabIndex        =   12
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            Height          =   195
            Index           =   0
            Left            =   840
            TabIndex        =   11
            Top             =   360
            Width           =   450
         End
      End
   End
End
Attribute VB_Name = "frmPoliticaInversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrTipoReglamento()     As String, arrFondo()           As String
Dim arrEstructura()         As String, arrCriterio()        As String

Dim strCodTipoReglamento    As String, strCodFondo          As String
Dim strCodEstructura        As String, strCodLimite         As String
Dim strCodCriterio          As String
Dim strEstado               As String, strSQL               As String
Dim adoConsulta             As ADODB.Recordset
Dim indSortAsc              As Boolean, indSortDesc         As Boolean

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String

    If tabPoliticaInversion.Tab = 1 Then Exit Sub

    Select Case Index
        Case 1
            gstrNameRepo = "LimiteInversion"

            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(2)
            ReDim aReportParamFn(3)
            ReDim aReportParamF(3)

            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            aReportParamFn(3) = "Fondo"
           

            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
            aReportParamF(3) = cboFondo.List(cboFondo.ListIndex)
            
            
            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora
            aReportParamS(2) = strCodTipoReglamento
            

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

Private Sub cboCriterio_Click()

    strCodCriterio = Valor_Caracter
    If cboCriterio.ListIndex < 0 Then Exit Sub
    
    strCodCriterio = Trim(arrCriterio(cboCriterio.ListIndex))
    
End Sub


Private Sub cboFondo_Click()

    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    If strCodFondo = Valor_Caracter Then strCodFondo = "000"
    
End Sub


Private Sub cboLimite_Click()

    strCodEstructura = Valor_Caracter
    If cboLimite.ListIndex < 0 Then Exit Sub
    
    strCodEstructura = Trim(arrEstructura(cboLimite.ListIndex))
    
    Call Buscar
    
End Sub


Private Sub cboTipoReglamento_Click()

    strCodTipoReglamento = Valor_Caracter
    If cboTipoReglamento.ListIndex < 0 Then Exit Sub
    
    strCodTipoReglamento = Trim(arrTipoReglamento(cboTipoReglamento.ListIndex))
    
    strSQL = "SELECT CodEstructura CODIGO, DescripEstructura DESCRIP FROM LimiteReglamentoEstructura " & _
        "WHERE CodReglamento='" & strCodTipoReglamento & "' AND Estado='" & Estado_Activo & "' ORDER BY DescripEstructura"
    CargarControlLista strSQL, cboLimite, arrEstructura(), Valor_Caracter
    
    If cboLimite.ListCount > 0 Then cboLimite.ListIndex = 0
    
    '*** Criterios Comparación ***
    strSQL = "SELECT CodCriterio CODIGO,DescripCriterio DESCRIP FROM CriterioBaseReglamento WHERE CodReglamento='" & strCodTipoReglamento & "' ORDER BY DescripCriterio"
    CargarControlLista strSQL, cboCriterio, arrCriterio(), Sel_Defecto
    
End Sub


Private Sub Form_Activate()

    Call CargarReportes
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Límites de Inversión"
    
End Sub
Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
                
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
    With tabPoliticaInversion
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .Tab = 0
    End With
    strEstado = Reg_Consulta
    
End Sub
Public Sub Grabar()

    Dim intContador     As Integer, intRegistro     As Integer
    Dim intAccion       As Integer, lngNumError     As Long
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    On Error GoTo CtrlError
            
    If strEstado = Reg_Edicion Then
        
        frmMainMdi.stbMdi.Panels(3).Text = "Actualizar la Política de Inversión..."
        
        Me.MousePointer = vbHourglass
        
        With adoComm
            .CommandText = "UPDATE LimiteReglamento SET " & _
                "CodCriterio='" & strCodCriterio & "'," & _
                "PorcenMinimo=" & CDec(txtPorcenMinimo.Text) & "," & _
                "PorcenMaximo=" & CDec(txtPorcenMaximo.Text) & " " & _
                "WHERE CodLimite='" & strCodLimite & "' AND CodEstructura='" & strCodEstructura & "' AND " & _
                "CodReglamento='" & strCodTipoReglamento & "' AND CodFondo='" & strCodFondo & "' AND " & _
                "CodAdministradora='" & gstrCodAdministradora & "' AND CodCriterio='" & tdgConsulta.Columns(2).Value & "'"
            adoConn.Execute .CommandText, intRegistro
        
            If intRegistro = 0 Then
                .CommandText = "INSERT INTO LimiteReglamento VALUES ('" & _
                    gstrCodAdministradora & "','" & strCodFondo & "','" & _
                    strCodTipoReglamento & "','" & strCodEstructura & "','" & _
                    strCodLimite & "','" & strCodCriterio & "'," & CDec(txtPorcenMinimo.Text) & "," & _
                    CDec(txtPorcenMaximo.Text) & ")"
                adoConn.Execute .CommandText
            End If
        End With
    
        Me.MousePointer = vbDefault
                        
        MsgBox Mensaje_Edicion_Exitosa, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        cmdOpcion.Visible = True
        With tabPoliticaInversion
            .TabEnabled(0) = True
            .TabEnabled(1) = False
            .Tab = 0
        End With
        Call Buscar

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
Public Sub Imprimir()

    Call SubImprimir(1)
    
    
End Sub

Public Sub Buscar()

    Set adoConsulta = New ADODB.Recordset
                
    strSQL = "SELECT LRED.CodLimite,DescripLimite,CodCriterio,isnull(PorcenMinimo,0) PorcenMinimo,isnull(PorcenMaximo,0) PorcenMaximo " & _
        "FROM LimiteReglamento LR RIGHT JOIN LimiteReglamentoEstructuraDetalle LRED ON(LRED.CodLimite=LR.CodLimite AND LRED.CodEstructura=LR.CodEstructura AND " & _
        "CodReglamento='" & strCodTipoReglamento & "' AND " & _
        "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "') " & _
        "WHERE LRED.CodEstructura='" & strCodEstructura & "' AND IndTitulo<>'X' AND LRED.Estado='" & _
        Estado_Activo & "' ORDER BY DescripLimite"
                        
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

    If strEstado = Reg_Consulta Or strEstado = Reg_Defecto Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabPoliticaInversion
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .Tab = 1
        End With
    End If
        
End Sub
Private Sub LlenarFormulario(strModo As String)

    Dim intRegistro As Integer
    
    Select Case strModo
        Case Reg_Edicion
            lblTipoReglamento.Caption = Trim(cboTipoReglamento.Text)
            lblDescripFondo.Caption = Trim(cboFondo.Text)
            lblLimite.Caption = Trim(cboLimite.Text)
            strCodLimite = Trim(tdgConsulta.Columns(0).Value)
            lblConcepto.Caption = Trim(tdgConsulta.Columns(1).Value)
            lblConcepto.ToolTipText = Trim(tdgConsulta.Columns(1).Value)
            
            If cboCriterio.ListCount > 0 Then cboCriterio.ListIndex = 0
            
            intRegistro = ObtenerItemLista(arrCriterio(), tdgConsulta.Columns(2).Value)
            If intRegistro >= 0 Then cboCriterio.ListIndex = intRegistro
            
            txtPorcenMinimo.Text = CStr(tdgConsulta.Columns(3).Value)
            txtPorcenMaximo.Text = CStr(tdgConsulta.Columns(4).Value)

    End Select
    
End Sub
Private Sub Form_Deactivate()

    Call OcultarReportes
    
End Sub


Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call DarFormato
    Call Buscar
    
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
    
    '*** Reglamentos ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPREG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoReglamento, arrTipoReglamento(), Valor_Caracter
    
    If cboTipoReglamento.ListCount > 0 Then cboTipoReglamento.ListIndex = 0
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Sel_Todos
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
End Sub
Private Sub InicializarValores()
                        
    '*** Valores Iniciales ***
    tabPoliticaInversion.Tab = 0
    tabPoliticaInversion.TabEnabled(1) = False
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 70
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 12
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 12
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub tabPoliticaInversion_Click(PreviousTab As Integer)

    Select Case tabPoliticaInversion.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabPoliticaInversion.Tab = 0
        
    End Select
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 3 Then
        Call DarFormatoValor(Value, Decimales_Tasa2)
    End If
    
    If ColIndex = 4 Then
        Call DarFormatoValor(Value, Decimales_Tasa2)
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

Private Sub txtPorcenMaximo_Change()

    Call FormatoCajaTexto(txtPorcenMaximo, Decimales_Tasa2)
    
End Sub

Private Sub txtPorcenMaximo_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtPorcenMaximo, Decimales_Tasa2)
    
End Sub


Private Sub txtPorcenMinimo_Change()

    Call FormatoCajaTexto(txtPorcenMinimo, Decimales_Tasa2)
    
End Sub


Private Sub txtPorcenMinimo_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtPorcenMinimo, Decimales_Tasa2)
    
End Sub


