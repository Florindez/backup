VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmFondoParametro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Par�metros Generales por Fondo"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   12900
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   10320
      TabIndex        =   14
      Top             =   7200
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      Visible0        =   0   'False
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   1080
      TabIndex        =   13
      Top             =   7200
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   1296
      Buttons         =   3
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      Visible0        =   0   'False
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      Visible1        =   0   'False
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Buscar"
      Tag2            =   "5"
      Visible2        =   0   'False
      ToolTipText2    =   "Buscar"
      UserControlWidth=   4200
   End
   Begin TabDlg.SSTab tabParametro 
      Height          =   7035
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   12409
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
      TabPicture(0)   =   "frmFondoParametro.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraLista"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmFondoParametro.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(1)=   "fraParametros"
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -66240
         TabIndex        =   2
         Top             =   5640
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   1296
         Buttons         =   2
         Caption0        =   "&Guardar"
         Tag0            =   "2"
         Visible0        =   0   'False
         ToolTipText0    =   "Guardar"
         Caption1        =   "&Cancelar"
         Tag1            =   "8"
         Visible1        =   0   'False
         ToolTipText1    =   "Cancelar"
         UserControlWidth=   2700
      End
      Begin VB.Frame fraParametros 
         Height          =   4695
         Left            =   -74640
         TabIndex        =   3
         Top             =   660
         Width           =   11955
         Begin VB.ComboBox cboEstado 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   2520
            Width           =   2745
         End
         Begin VB.TextBox txtCodigo 
            Height          =   315
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   7
            Top             =   480
            Width           =   2080
         End
         Begin VB.TextBox txtDescripParametro 
            Height          =   315
            Left            =   1800
            MaxLength       =   200
            TabIndex        =   6
            Top             =   982
            Width           =   9645
         End
         Begin VB.TextBox txtValorParametro 
            Height          =   285
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   5
            Top             =   2016
            Width           =   2775
         End
         Begin VB.ComboBox cboTipoDato 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1484
            Width           =   2775
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   15
            Top             =   2540
            Width           =   495
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   11
            Top             =   2040
            Width           =   360
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Dato"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   10
            Top             =   1504
            Width           =   705
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Descripci�n"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   9
            Top             =   1002
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "C�digo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   8
            Top             =   500
            Width           =   495
         End
      End
      Begin VB.Frame fraLista 
         Height          =   5895
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   12015
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   450
            Width           =   7425
         End
         Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
            Height          =   4065
            Left            =   360
            OleObjectBlob   =   "frmFondoParametro.frx":0038
            TabIndex        =   12
            Top             =   1170
            Width           =   11205
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fondo"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   18
            Top             =   495
            Width           =   1035
         End
      End
   End
End
Attribute VB_Name = "frmFondoParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Ficha de Mantenimiento de Costos de Negociaci�n"
Option Explicit

Dim arrTipoDato()       As String, arrEstado()          As String
Dim arrFondo()          As String

Dim strCodTipoDato      As String, strCodEstado         As String
Dim strEstado           As String, strSQL               As String
Dim strCodFondo         As String
Dim adoConsulta         As ADODB.Recordset
Dim indSortAsc          As Boolean, indSortDesc         As Boolean

Public Sub Adicionar()

    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Costo..."
                
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabParametro
        .TabEnabled(0) = False
        .Tab = 1
        .TabEnabled(1) = True
    End With
    
End Sub

Private Sub LlenarFormulario(strModo As String)
    
    Dim adoRegistro         As ADODB.Recordset
    Dim intRegistro         As Integer
    
    Select Case strModo
        Case Reg_Adicion
            
            txtCodigo.Text = "GENERADO"
            
            txtDescripParametro.Text = Valor_Caracter

            cboTipoDato.ListIndex = -1
            If cboTipoDato.ListCount > 0 Then cboTipoDato.ListIndex = 0
            
            cboEstado.ListIndex = -1
            intRegistro = ObtenerItemLista(arrEstado(), Estado_Activo)
            If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
            
            txtValorParametro.Text = Valor_Caracter
            txtCodigo.Enabled = False
            txtDescripParametro.SetFocus
                        
        Case Reg_Edicion
            txtCodigo.Text = tdgConsulta.Columns(0).Value
            txtDescripParametro.Text = tdgConsulta.Columns(1).Value

            If cboTipoDato.ListCount > 0 Then cboTipoDato.ListIndex = 0
            intRegistro = ObtenerItemLista(arrTipoDato(), tdgConsulta.Columns(3).Value)
            If intRegistro >= 0 Then cboTipoDato.ListIndex = intRegistro
            
            txtValorParametro = tdgConsulta.Columns(2).Value
            
            If cboEstado.ListCount > 0 Then cboEstado.ListIndex = 0
            intRegistro = ObtenerItemLista(arrEstado(), tdgConsulta.Columns(4).Value)
            If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
            
            txtCodigo.Enabled = False
                
    End Select
    
End Sub
Public Sub Buscar()

    Set adoConsulta = New ADODB.Recordset

    strSQL = "SELECT AP.DescripParametro DescripTipoDato," & _
        "FPG.CodParametro,TipoValor,FPG.ValorParametro,FPG.DescripParametro,FPG.Estado " & _
        "FROM FondoParametroGeneral FPG JOIN AuxiliarParametro AP ON(AP.CodParametro=FPG.TipoValor AND AP.CodTipoParametro='TIPDAT') " & _
        "WHERE FPG.Estado='" & Estado_Activo & "' AND FPG.CodFondo = '" & strCodFondo & "' AND " & _
        "FPG.CodAdministradora = '" & gstrCodAdministradora & "'"
        
    strEstado = Reg_Defecto
    
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    tdgConsulta.DataSource = adoConsulta
       
    If adoConsulta.RecordCount > 0 Then
        strEstado = Reg_Consulta
    Else
        If MsgBox("No se encentran parametros definidos para este Fondo!" & vbNewLine & vbNewLine & _
            "Desea Importar los Par�metros de la Definici�n General de Par�metros?", vbQuestion + vbYesNo) = vbYes Then
            
            With adoComm
                .CommandText = "{ call up_GNProcImportarFondoParametroGeneral('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "') }"
                adoConn.Execute .CommandText
            End With
            
            Call Buscar
        
        End If
    End If

    
    
    
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabParametro
        .TabEnabled(0) = True
        .Tab = 0
        .TabEnabled(1) = False
    End With
    Call Buscar
    
End Sub

Private Sub CargarListas()


    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Sel_Todos
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0

    '*** Tipo de Dato ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPDAT' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoDato, arrTipoDato(), Valor_Caracter
    
    '*** Estado ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='ESTREG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Sel_Defecto
    
End Sub

Public Sub Eliminar()

End Sub

Public Sub Grabar()
            
    Dim intAccion   As Integer, lngNumError     As Long
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    On Error GoTo CtrlError
    
    'If strEstado = Reg_Adicion Then
        If TodoOK() Then
                
            With adoComm
                .CommandText = "{ call up_GNManFondoParametroGeneral ('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    Trim(txtCodigo.Text) & "','" & Trim(txtDescripParametro.Text) & "','" & _
                    strCodTipoDato & "','" & Trim(txtValorParametro.Text) & "','" & _
                    strCodEstado & "','" & IIf(strEstado = Reg_Adicion, "I", IIf(strEstado = Reg_Edicion, "U", "D")) & "')}"
                adoConn.Execute .CommandText
            End With
            
            Me.MousePointer = vbDefault
        
            If strEstado = Reg_Adicion Then
                MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            End If
            
            If strEstado = Reg_Edicion Then
                MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            End If
            
            If strEstado = Reg_Eliminacion Then
                MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation
            End If
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acci�n"
            
            cmdOpcion.Visible = True
            With tabParametro
                .TabEnabled(0) = True
                .Tab = 0
                .TabEnabled(1) = False
            End With
            
            Call Buscar
        End If
    'End If
    
    '*** Set de Par�metros Globales ***
    If Not CargarParametrosGlobales(strCodFondo) Then Exit Sub
    
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

End Sub

Public Sub Modificar()

    If strEstado = Reg_Defecto Then Exit Sub
    
    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabParametro
            .TabEnabled(0) = False
            .Tab = 1
            .TabEnabled(1) = True
        End With
        'Call Habilita
    End If
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String

    If tabParametro.Tab = 1 Then Exit Sub
    
    Select Case Index
        Case 1
            gstrNameRepo = "Parametro"
                        
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
            
            aReportParamS(0) = "01"
            aReportParamS(1) = " "
            
    End Select

    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub

Private Sub cboEstado_Click()

    strCodEstado = Valor_Caracter
    
    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strCodEstado = Trim(arrEstado(cboEstado.ListIndex))
    
End Sub


Private Sub cboFondo_Click()
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    If strCodFondo = "" Then strCodFondo = "000"
    
    fraParametros.Caption = cboFondo.List(cboFondo.ListIndex)
    
    Call Buscar
    
End Sub

Private Sub cboTipoDato_Click()

    strCodTipoDato = Valor_Caracter
    
    If cboTipoDato.ListIndex < 0 Then Exit Sub
    
    strCodTipoDato = Trim(arrTipoDato(cboTipoDato.ListIndex))
            
End Sub


Private Sub Form_Activate()

    frmMainMdi.stbMdi.Panels(3).Text = Me.Caption
   
End Sub

Private Sub Form_Deactivate()

    Call OcultarReportes
    
End Sub


Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call CargarReportes
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

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Vista Activa"
    
End Sub
Private Sub InicializarValores()
    
    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabParametro.Tab = 0
    tabParametro.TabEnabled(1) = False
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 8
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 75
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 8
    
    Set adoConsulta = New ADODB.Recordset
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
                 
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmFondoParametro = Nothing
    frmMainMdi.stbMdi.Panels(3).Text = "Acci�n"
    
End Sub

Private Function TodoOK() As Boolean

    TodoOK = False
    
    If strCodTipoDato = Codigo_TipoDato_Numerico Then
        If Not IsNumeric(txtValorParametro) Then
            MsgBox "El valor no es un dato num�rico...", vbCritical
            Exit Function
        End If
    ElseIf strCodTipoDato = Codigo_TipoDato_AlfaNumerico Then
        If IsDate(txtValorParametro) Then
            MsgBox "El valor no es un dato alfanum�rico...", vbCritical
            Exit Function
        End If
    Else
        If Not IsDate(txtValorParametro) Then
            MsgBox "El valor no es un dato fecha...", vbCritical
            Exit Function
        End If
    End If
    
    '*** Si todo paso OK ***
    TodoOK = True

End Function

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

Private Sub tabParametro_Click(PreviousTab As Integer)

    Select Case tabParametro.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabParametro.Tab = 0
    End Select
    
End Sub

Private Sub txtValorParametro_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
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
