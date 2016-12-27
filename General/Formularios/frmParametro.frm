VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmParametro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros Generales"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   11355
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   9480
      TabIndex        =   14
      Top             =   6360
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
      TabIndex        =   13
      Top             =   6360
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
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   6180
      Top             =   6360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabParametro 
      Height          =   6195
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10927
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
      TabPicture(0)   =   "frmParametro.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraLista"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmParametro.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraParametros"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAccion"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -67080
         TabIndex        =   2
         Top             =   4440
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
      Begin VB.Frame fraParametros 
         Height          =   3255
         Left            =   -74640
         TabIndex        =   3
         Top             =   540
         Width           =   10545
         Begin VB.ComboBox cboEstado 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   2520
            Width           =   2080
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
            Width           =   8325
         End
         Begin VB.TextBox txtValorParametro 
            Height          =   315
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   5
            Top             =   2016
            Width           =   2080
         End
         Begin VB.ComboBox cboTipoDato 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1484
            Width           =   2080
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
            Caption         =   "Descripción"
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
            Caption         =   "Código"
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
         Height          =   4515
         Left            =   360
         TabIndex        =   1
         Top             =   750
         Width           =   10545
         Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
            Bindings        =   "frmParametro.frx":0038
            Height          =   3615
            Left            =   360
            OleObjectBlob   =   "frmParametro.frx":0052
            TabIndex        =   12
            Top             =   450
            Width           =   9945
         End
      End
   End
End
Attribute VB_Name = "frmParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Ficha de Mantenimiento de Costos de Negociación"
Option Explicit

Dim arrTipoDato()       As String, arrEstado()          As String

Dim strCodTipoDato      As String, strCodEstado         As String
Dim strEstado           As String, strSQL               As String


Public Sub Adicionar()

    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Costo..."
                
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabParametro
        .TabEnabled(0) = False
        .Tab = 1
    End With
    
End Sub

Private Sub LlenarFormulario(strModo As String)
    
    Dim adoRegistro         As ADODB.Recordset
    Dim intRegistro         As Integer
    
    Select Case strModo
        Case Reg_Adicion
            Set adoRegistro = New ADODB.Recordset
            
            adoComm.CommandText = "SELECT COUNT(*) SecuencialParametro FROM ParametroGeneral"
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                txtCodigo.Text = Format(adoRegistro("SecuencialParametro") + 1, "00")
            Else
                txtCodigo.Text = "01"
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
            
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

    strSQL = "SELECT AP.DescripParametro DescripTipoDato," & _
        "PG.CodParametro,TipoValor,PG.ValorParametro,PG.DescripParametro,PG.Estado " & _
        "FROM ParametroGeneral PG JOIN AuxiliarParametro AP ON(AP.CodParametro=PG.TipoValor AND AP.CodTipoParametro='TIPDAT') " & _
        "WHERE PG.Estado='" & Estado_Activo & "'"
        
    strEstado = Reg_Defecto
    With adoConsulta
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With

    tdgConsulta.Refresh

    If adoConsulta.Recordset.RecordCount > 0 Then strEstado = Reg_Consulta
    
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabParametro
        .TabEnabled(0) = True
        .Tab = 0
    End With
    Call Buscar
    
End Sub

Private Sub CargarListas()

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
    
    If strEstado = Reg_Adicion Then
        If TodoOK() Then
            With adoComm
                .CommandText = "INSERT INTO ParametroGeneral VALUES ('" & _
                    Trim(txtCodigo.Text) & "','" & Trim(txtDescripParametro.Text) & "','" & _
                    strCodTipoDato & "','" & Trim(txtValorParametro.Text) & "','" & _
                    strCodEstado & "')"
                adoConn.Execute .CommandText
            End With
            
            Me.MousePointer = vbDefault
        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabParametro
                .TabEnabled(0) = True
                .Tab = 0
            End With
            
            Call Buscar
        End If
    End If
    
    If strEstado = Reg_Edicion Then
        If TodoOK() Then
            With adoComm
                .CommandText = "UPDATE ParametroGeneral SET " & _
                    "DescripParametro='" & Trim(txtDescripParametro.Text) & "'," & _
                    "TipoValor='" & strCodTipoDato & "'," & _
                    "ValorParametro='" & Trim(txtValorParametro.Text) & "'," & _
                    "Estado='" & strCodEstado & "' " & _
                    "WHERE CodParametro='" & Trim(txtCodigo.Text) & "'"
                adoConn.Execute .CommandText
            End With
            
            Me.MousePointer = vbDefault
        
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabParametro
                .TabEnabled(0) = True
                .Tab = 0
            End With
            
            Call Buscar
        End If
    End If
    
    '*** Set de Parámetros Globales ***
    If Not CargarParametrosGlobales() Then Exit Sub
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

            ReDim aReportParamS(0)
            ReDim aReportParamFn(2)
            ReDim aReportParamF(2)
                        
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
                        
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
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
    Call Buscar
    Call DarFormato
    
    CentrarForm Me
        
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
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
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 50
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 20
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
                 
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set frmParametro = Nothing
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub

Private Function TodoOK() As Boolean

    TodoOK = False
    
    If strCodTipoDato = Codigo_TipoDato_Numerico Then
        If Not IsNumeric(txtValorParametro) Then
            MsgBox "El valor no es un dato numérico...", vbCritical
            Exit Function
        End If
    ElseIf strCodTipoDato = Codigo_TipoDato_AlfaNumerico Then
        If IsDate(txtValorParametro) Then
            MsgBox "El valor no es un dato alfanumérico...", vbCritical
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


