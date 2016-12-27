VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Begin VB.Form frmVistaProceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vista Proceso"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   10770
   Begin TabDlg.SSTab tabVistaProceso 
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   14843
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
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
      TabCaption(0)   =   "Vista"
      TabPicture(0)   =   "frmVistaProceso.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdOpcion"
      Tab(0).Control(1)=   "cmdSalir"
      Tab(0).Control(2)=   "tdgConsulta"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Datos"
      TabPicture(1)   =   "frmVistaProceso.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frVistaProceso"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame frVistaProceso 
         Height          =   7275
         Left            =   120
         TabIndex        =   4
         Top             =   420
         Width           =   10155
         Begin VB.Frame frVistaProcesoDetalleParametro 
            Caption         =   "Detalle Parametro"
            Height          =   2445
            Left            =   150
            TabIndex        =   21
            Top             =   4470
            Width           =   9825
            Begin TrueOleDBGrid60.TDBGrid tdgVistaProcesoDetalleParametro 
               Bindings        =   "frmVistaProceso.frx":0038
               Height          =   2055
               Left            =   120
               OleObjectBlob   =   "frmVistaProceso.frx":0052
               TabIndex        =   22
               Top             =   270
               Width           =   9525
            End
         End
         Begin VB.Frame frDatosVista 
            Height          =   2235
            Left            =   300
            TabIndex        =   14
            Top             =   1830
            Width           =   3585
            Begin VB.ComboBox cboVistaUsuario 
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   540
               Width           =   3225
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vista Usuario"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   20
               Top             =   150
               Width           =   930
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Codigo"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   19
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label lblCodVistaUsuario 
               Height          =   255
               Left            =   960
               TabIndex        =   18
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo"
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   17
               Top             =   1830
               Width           =   315
            End
            Begin VB.Label lblTipoVistaUsuario 
               Height          =   255
               Left            =   960
               TabIndex        =   16
               Top             =   1830
               Width           =   2415
            End
         End
         Begin VB.Frame frVistaProcesoDetalle 
            Caption         =   "(*) Detalle"
            Height          =   2775
            Left            =   150
            TabIndex        =   8
            Top             =   1500
            Width           =   9825
            Begin VB.CommandButton cmdQuitar 
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   3900
               Style           =   1  'Graphical
               TabIndex        =   12
               ToolTipText     =   "Quitar"
               Top             =   1320
               Width           =   375
            End
            Begin VB.CommandButton cmdAgregar 
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   3900
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Agregar"
               Top             =   900
               Width           =   375
            End
            Begin TrueOleDBGrid60.TDBGrid tdgVistaProcesoDetalle 
               Bindings        =   "frmVistaProceso.frx":45BA
               Height          =   2355
               Left            =   4380
               OleObjectBlob   =   "frmVistaProceso.frx":45D4
               TabIndex        =   9
               Top             =   240
               Width           =   5265
            End
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   315
            Left            =   2400
            TabIndex        =   5
            Top             =   810
            Width           =   5010
         End
         Begin VB.Label lblCodVistaProceso 
            Caption         =   "Label1"
            Height          =   255
            Left            =   2400
            TabIndex        =   13
            Top             =   420
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo"
            Height          =   195
            Index           =   1
            Left            =   420
            TabIndex        =   7
            Top             =   420
            Width           =   495
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion"
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   6
            Top             =   810
            Width           =   840
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmVistaProceso.frx":7F0F
         Height          =   6945
         Left            =   -74790
         OleObjectBlob   =   "frmVistaProceso.frx":7F29
         TabIndex        =   1
         Top             =   510
         Width           =   9975
      End
      Begin TAMControls.ucBotonEdicion cmdSalir 
         Height          =   390
         Left            =   -66030
         TabIndex        =   2
         Top             =   7740
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   688
         Caption0        =   "&Salir"
         Tag0            =   "9"
         Visible0        =   0   'False
         ToolTipText0    =   "Salir"
         UserControlHeight=   390
         UserControlWidth=   1200
      End
      Begin TAMControls.ucBotonEdicion cmdOpcion 
         Height          =   390
         Left            =   -74820
         TabIndex        =   3
         Top             =   7740
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   688
         Buttons         =   3
         Caption0        =   "&Nuevo"
         Tag0            =   "0"
         Visible0        =   0   'False
         ToolTipText0    =   "Nuevo"
         Caption1        =   "&Modificar"
         Tag1            =   "1"
         Visible1        =   0   'False
         ToolTipText1    =   "Modificar"
         Caption2        =   "&Eliminar"
         Tag2            =   "4"
         Visible2        =   0   'False
         ToolTipText2    =   "Eliminar"
         UserControlHeight=   390
         UserControlWidth=   4200
      End
      Begin TAMControls.ucBotonEdicion cmdAccion 
         Height          =   390
         Left            =   7260
         TabIndex        =   10
         Top             =   7830
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   688
         Buttons         =   2
         Caption0        =   "&Guardar"
         Tag0            =   "2"
         Visible0        =   0   'False
         ToolTipText0    =   "Guardar"
         Caption1        =   "&Cancelar"
         Tag1            =   "8"
         Visible1        =   0   'False
         ToolTipText1    =   "Cancelar"
         UserControlHeight=   390
         UserControlWidth=   2700
      End
   End
End
Attribute VB_Name = "frmVistaProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arrCodigoVista() As String, adoConsulta As ADODB.Recordset
Dim strEstado        As String, strCodVistaProceso As String
Dim adoVistaProcesoDetalle As ADODB.Recordset, adoVistaProcesoDetalleParametro As ADODB.Recordset

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

Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabVistaProceso.Tab = 0
    tabVistaProceso.TabEnabled(1) = False

    
    lblCodVistaProceso.FontBold = True
    lblCodVistaUsuario.FontBold = True
    lblTipoVistaUsuario.FontBold = True
    
    frVistaProcesoDetalle.FontBold = True
    frVistaProcesoDetalle.ForeColor = &H800000
    
    frVistaProcesoDetalleParametro.FontBold = True
    frVistaProcesoDetalleParametro.ForeColor = &H800000
    frVistaProcesoDetalleParametro.Visible = False
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me

End Sub

Private Sub CargarListas()

  Dim strSQL As String
  
  strSQL = "SELECT CodVistaUsuario AS CODIGO,DescripVista AS DESCRIP FROM VistaUsuario"
            
  CargarControlLista strSQL, cboVistaUsuario, arrCodigoVista(), Sel_Defecto

End Sub

Private Sub CargarReportes()

'   frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Diario General"
    
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Diario General (ME)"
    
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Mayor General"
    
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Text = "Mayor General (ME)"
    
    
End Sub

Public Sub Buscar()

    Dim strSQL As String
    
    Set adoConsulta = New ADODB.Recordset
           
    strSQL = "SELECT CodVistaProceso,DescripVistaProceso FROM VistaProceso " & _
                "WHERE IndVigente='X'"
                        
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

Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
            
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vQuery
            Call Modificar
        Case vDelete
            Call Eliminar
'        Case vSearch
'            Call Buscar
'        Case vReport
'            Call Imprimir
        Case vSave
            Call Grabar
        Case vCancel
            Call Cancelar
        Case vExit
            Call Salir
        
    End Select
    
End Sub

Public Sub Adicionar()
    If strEstado = Reg_Consulta Then
        strEstado = Reg_Adicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabVistaProceso
            .TabEnabled(0) = False
            .Tab = 1
            .TabEnabled(1) = True
        End With
    End If
    
End Sub

Public Sub Modificar()
    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabVistaProceso
            .TabEnabled(0) = False
            .Tab = 1
            .TabEnabled(1) = True
        End With
    End If
    
End Sub

Public Sub Grabar()

    Dim objVistaProcesoDetalleXML  As DOMDocument60
    Dim objVistaProcesoDetalleParametroXML  As DOMDocument60
    Dim strMsgError  As String
    Dim strVistaProcesoDetalleXML As String
    Dim strVistaProcesoDetalleParametroXML As String
    

    If strEstado = Reg_Consulta Then Exit Sub

    If strEstado = Reg_Adicion Then

    If MsgBox(Mensaje_Adicion, vbQuestion + vbYesNo + vbDefaultButton2, gstrNombreEmpresa) = vbNo Then Exit Sub

        If TodoOK() Then

            Me.MousePointer = vbHourglass

            On Error GoTo Ctrl_Error

            With adoComm
                
                Call NumerarVistaProcesoDetalle
                
                Call XMLADORecordset(objVistaProcesoDetalleXML, "VistaProcesoDetalle", "Estructura", adoVistaProcesoDetalle, strMsgError)
    
                strVistaProcesoDetalleXML = objVistaProcesoDetalleXML.xml
                
                Call XMLADORecordset(objVistaProcesoDetalleParametroXML, "VistaProcesoDetalleParametro", "Estructura", adoVistaProcesoDetalleParametro, strMsgError)
    
                strVistaProcesoDetalleParametroXML = objVistaProcesoDetalleParametroXML.xml
    
                .CommandText = "{ call up_ACManVistaProcesoXML('" & _
                lblCodVistaProceso.Caption & "','" & txtDescripcion.Text & "','X','" & _
                strVistaProcesoDetalleXML & "','" & strVistaProcesoDetalleParametroXML & "','I') }"
    
                adoConn.Execute .CommandText
            
            End With

            Me.MousePointer = vbDefault

            MsgBox Mensaje_Adicion_Exitosa, vbExclamation

            frmMainMdi.stbMdi.Panels(3).Text = "Acción"

            cmdOpcion.Visible = True
            With tabVistaProceso
                .TabEnabled(0) = True
                .Tab = 0
                .TabEnabled(1) = False
            End With

            Call Limpiar
            Call Buscar

            tdgVistaProcesoDetalle.DataSource = Nothing '
            tdgVistaProcesoDetalleParametro.DataSource = Nothing '

        End If
    End If

    If strEstado = Reg_Edicion Then

    If MsgBox(Mensaje_Edicion, vbQuestion + vbYesNo + vbDefaultButton2, gstrNombreEmpresa) = vbNo Then Exit Sub

        If TodoOK() Then

            Me.MousePointer = vbHourglass


            On Error GoTo Ctrl_Error
            
            With adoComm
                
                Call NumerarVistaProcesoDetalle
                
                Call XMLADORecordset(objVistaProcesoDetalleXML, "VistaProcesoDetalle", "Estructura", adoVistaProcesoDetalle, strMsgError)
    
                strVistaProcesoDetalleXML = objVistaProcesoDetalleXML.xml
                
                Call XMLADORecordset(objVistaProcesoDetalleParametroXML, "VistaProcesoDetalleParametro", "Estructura", adoVistaProcesoDetalleParametro, strMsgError)
    
                strVistaProcesoDetalleParametroXML = objVistaProcesoDetalleParametroXML.xml
                        
                .CommandText = "{ call up_ACManVistaProcesoXML('" & _
                    lblCodVistaProceso.Caption & "','" & txtDescripcion.Text & "','X','" & _
                    strVistaProcesoDetalleXML & "','" & strVistaProcesoDetalleParametroXML & "','U') }"
    
                adoConn.Execute .CommandText
            
            End With
            
            Me.MousePointer = vbDefault

            MsgBox Mensaje_Edicion_Exitosa, vbExclamation

            frmMainMdi.stbMdi.Panels(3).Text = "Acción"

            cmdOpcion.Visible = True
            With tabVistaProceso
                .TabEnabled(0) = True
                .Tab = 0
                .TabEnabled(1) = False
            End With

            Call Limpiar
            Call Buscar

            tdgVistaProcesoDetalle.DataSource = Nothing '
            tdgVistaProcesoDetalleParametro.DataSource = Nothing '
            
        End If

    End If


Exit Sub

Ctrl_Error:

'        adoComm.CommandText = "ROLLBACK TRAN ProcAsiento"
'        adoConn.Execute adoComm.CommandText

        MsgBox err.Description & vbCrLf & Mensaje_Proceso_NoExitoso, vbCritical
        Me.MousePointer = vbDefault


End Sub

Public Sub Eliminar()

    If MsgBox(Mensaje_Eliminacion, vbQuestion + vbYesNo + vbDefaultButton2, gstrNombreEmpresa) = vbNo Then Exit Sub

    If strEstado = Reg_Consulta Then

            Me.MousePointer = vbHourglass

                '*** Guardar ***
            With adoComm
                .CommandText = "{ call up_ACManVistaProcesoXML('" & _
                tdgConsulta.Columns(0) & "','','','','','D') }"

                adoConn.Execute .CommandText

            End With

            Me.MousePointer = vbDefault

            MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation

            frmMainMdi.stbMdi.Panels(3).Text = "Acción"

            Call Buscar

    End If

End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    cmdAccion.Button(0).Enabled = True
    
    With tabVistaProceso
        .TabEnabled(0) = True
        .Tab = 0
        .TabEnabled(1) = False
    End With
    
    Call Limpiar
    'Call LimpiarFiltro
    Call Buscar
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub


Public Sub LlenarFormulario(ByVal strModo As String)

    'Dim strCodVistaUsuario As String
    Dim strSQL As String
    
'    Dim adoRegistro As ADODB.Recordset
    
    'Dim intCont As Integer, intRegistro As Integer
    
'    txtValorFormulaVista.SelStart = 0
''    txtValorFormulaVista.SelLength = intTamaño
'    txtValorFormulaVista.SelColor = &H808000
    'txtValorFormulaVista.Font.Bold = False
    
    Select Case strModo
    
    Case Reg_Adicion
        
        cboVistaUsuario.ListIndex = 0
        
        frVistaProceso.Caption = "Nueva Vista"
        frVistaProceso.ForeColor = &H800000
        frVistaProceso.FontBold = True
        frVistaProceso.Font = "Arial"
        
        frVistaProcesoDetalleParametro.Visible = False
    
        lblCodVistaProceso.Caption = NuevoCodigo()
        
'        cmdEditarQuery.Caption = "Ok"
'        cmdEditarQuery.Enabled = False
'        cboIdVariableFiltro.ListIndex = 0
'        chkFiltro.Value = Unchecked
        
        txtDescripcion.SetFocus
        
        Call ConfiguraRecordsetProcesoDetalle
        
        tdgVistaProcesoDetalle.DataSource = adoVistaProcesoDetalle
        tdgVistaProcesoDetalleParametro.DataSource = adoVistaProcesoDetalleParametro
        
    Case Reg_Edicion
    
        cboVistaUsuario.ListIndex = 0
            
        strCodVistaProceso = tdgConsulta.Columns(0)
        
        lblCodVistaProceso.Caption = strCodVistaProceso
        
        txtDescripcion.Text = Trim(tdgConsulta.Columns(1))
        
        frVistaProceso.Caption = "Vista: " + tdgConsulta.Columns(1)
        frVistaProceso.ForeColor = &H800000
        frVistaProceso.FontBold = True
        frVistaProceso.Font = "Arial"
        
        'cmdEditarQuery.Caption = "Editar Query"
                
        Call ConfiguraRecordsetProcesoDetalle
        
        strSQL = "SELECT CodVistaProceso,VU.CodVistaUsuario AS CodVistaUsuario," & _
                "DescripVista As DescripVistaUsuario, DescripParametro As TipoVista " & _
                "FROM VistaProcesoDetalle VPD JOIN VistaUsuario VU " & _
                "ON (VPD.CodVistaUsuario=VU.CodVistaUsuario) " & _
                "JOIN AuxiliarParametro AP ON (AP.CodParametro=VU.TipoVista AND AP.CodTipoParametro='TIPVIS') " & _
                "WHERE CodVistaProceso='" & tdgConsulta.Columns(0).Value & "' ORDER BY Secuencial"
            
        Call CargarGrillaCampos(strSQL)
        
        strSQL = "SELECT CodVistaProceso,VPDP.CodVistaUsuario AS CodVistaUsuario,DescripVista AS DescripVistaUsuario," & _
                "CodParametroVistaProceso , SecVistaProceso, SecVistaUsuario " & _
                "FROM VistaProcesoDetalleParametro VPDP JOIN VistaUsuario VU " & _
                "ON (VPDP.CodVistaUsuario=VU.CodVistaUsuario) " & _
                "WHERE CodVistaProceso='" & strCodVistaProceso & "'"
        
        Call CargarGrillaCamposParametro(strSQL)
        
        If adoVistaProcesoDetalleParametro.RecordCount > 0 Then
        
            frVistaProcesoDetalleParametro.Visible = True
            
        Else
        
            frVistaProcesoDetalleParametro.Visible = False
        
        End If
        
        txtDescripcion.SetFocus
        
'        Call ConfiguraRecordsetCampos
'
'        strSql = "SELECT CodVistaUsuario,SecCampo,NombreCampo,DescripParametro AS DescripTipo," & _
'                     "TipoCampo , IdVariable FROM VistaUsuarioCampo VUC JOIN AuxiliarParametro AP " & _
'                        "ON(VUC.TipoCampo=AP.CodParametro AND AP.CodTipoParametro='SUBREP') " & _
'                        "WHERE CodVistaUsuario='" & strCodVistaUsuario & "' AND TipoCampo<>'03'"
'
'        Call CargarGrillaCampos(strSql, strCodVistaUsuario, True)
'
'        Call CargarTipoDatos
'
'        strSql = "SELECT CodVistaUsuario,SecCampo,NombreCampo,DescripParametro AS DescripTipo," & _
'                     "TipoCampo , IdVariable FROM VistaUsuarioCampo VUC JOIN AuxiliarParametro AP " & _
'                        "ON(VUC.TipoCampo=AP.CodParametro AND AP.CodTipoParametro='SUBREP') " & _
'                        "WHERE CodVistaUsuario='" & strCodVistaUsuario & "' AND TipoCampo='03'"
'
'        Call CargarGrillaCamposFiltro(strSql)
'
'        Call BloquearEdicionGrilla
'
'        If adoRegistroCampos.RecordCount > 0 Then
'
'            tabCampos.Visible = True
'            tabCampos.Tab = 0
'            tdgCamposVista.Caption = "Campos - " & Trim(txtDescripcion.Text)
'        Else
'            tabCampos.Visible = False
'
'        End If
'
'        If adoRegistroCamposFiltro.RecordCount > 0 Then
'
'            tdgCamposFiltro.Caption = "Campos - " & Trim(txtDescripcion.Text)
'            chkFiltro.Value = Checked
'        Else
'            chkFiltro.Value = Unchecked
'
'        End If
               
    End Select
    
End Sub

Private Sub tdgConsulta_DblClick()

    Call Modificar
    
End Sub

Private Sub cboVistaUsuario_Click()
    
    Dim strSQL As String
    Dim adoRegistro As ADODB.Recordset
    
    lblCodVistaUsuario.Caption = Valor_Caracter
    lblTipoVistaUsuario.Caption = Valor_Caracter
    
    Set adoRegistro = New ADODB.Recordset
    
    lblCodVistaUsuario.Caption = arrCodigoVista(cboVistaUsuario.ListIndex)
    
    strSQL = "SELECT DescripParametro AS DescripTipoVista " & _
            "FROM VistaUsuario VU JOIN AuxiliarParametro AP " & _
            "ON (VU.TipoVista=AP.CodParametro AND AP.CodTipoParametro='TIPVIS') " & _
            "WHERE CodVistaUsuario='" & lblCodVistaUsuario.Caption & "'"
    
    With adoComm
    
        .CommandText = strSQL
        
        Set adoRegistro = .Execute
    
    End With
    
    Do Until adoRegistro.EOF
    
        lblTipoVistaUsuario.Caption = adoRegistro.Fields("DescripTipoVista")
        
        adoRegistro.MoveNext
    
    Loop
    
End Sub

Private Sub cmdAgregar_Click()
    
    If TodoOkDetalle() Then
        
        adoVistaProcesoDetalle.AddNew
        adoVistaProcesoDetalle.Fields("CodVistaProceso") = lblCodVistaProceso.Caption
        adoVistaProcesoDetalle.Fields("CodVistaUsuario") = lblCodVistaUsuario.Caption
        adoVistaProcesoDetalle.Fields("DescripVistaUsuario") = Trim(cboVistaUsuario.Text)
        adoVistaProcesoDetalle.Fields("TipoVista") = lblTipoVistaUsuario.Caption
        
        adoVistaProcesoDetalle.Update
        
        Call AdicionarParametrosGrilla
        
        
        If adoVistaProcesoDetalle.RecordCount = 0 Then
        
            cmdQuitar.Enabled = False
        Else
        
            cmdQuitar.Enabled = True
        End If
        
        If Not adoVistaProcesoDetalleParametro.EOF Then
        
            frVistaProcesoDetalleParametro.Visible = True
        Else
        
            frVistaProcesoDetalleParametro.Visible = False
        End If
        
        Call LimpiarDetalle
    
    End If
    
End Sub

Private Sub cmdQuitar_Click()

    Dim dblBook As String
    Dim intIndice As Integer
    
    If adoVistaProcesoDetalle.RecordCount > 0 Then
              
        dblBook = adoVistaProcesoDetalle.AbsolutePosition

        adoVistaProcesoDetalle.Delete

        '**********SUBIR O BAJAR CURSOR SEGUN ELIMINE********
        If adoVistaProcesoDetalle.RecordCount >= 1 And dblBook = 1 Then
        
            adoVistaProcesoDetalle.MoveFirst
            tdgVistaProcesoDetalle.MoveFirst
            
            cmdQuitar.Enabled = True

        ElseIf adoVistaProcesoDetalle.RecordCount = 0 Then

            cmdQuitar.Enabled = False

        Else

            intIndice = CInt(dblBook) - 1
            dblBook = intIndice
            adoVistaProcesoDetalle.AbsolutePosition = CDbl(dblBook)
            
            cmdQuitar.Enabled = True

        End If
        
        Call AdicionarParametrosGrilla
        
        If Not adoVistaProcesoDetalleParametro.EOF Then
        
            frVistaProcesoDetalleParametro.Visible = True
        Else
        
            frVistaProcesoDetalleParametro.Visible = False
        End If
        
        Call LimpiarDetalle
    
    End If

End Sub

Private Sub ConfiguraRecordsetProcesoDetalle()

    Set adoVistaProcesoDetalle = New ADODB.Recordset

    With adoVistaProcesoDetalle
    
       .CursorLocation = adUseClient
       .Fields.Append "CodVistaProceso", adChar, 3
       .Fields.Append "CodVistaUsuario", adChar, 3
       .Fields.Append "DescripVistaUsuario", adVarChar, 200
       .Fields.Append "TipoVista", adVarChar, 60
       .Fields.Append "Secuencial", adInteger
'       .CursorType = adOpenStatic

       .LockType = adLockBatchOptimistic
    End With
    
    adoVistaProcesoDetalle.Open
    
    Set adoVistaProcesoDetalleParametro = New ADODB.Recordset

    With adoVistaProcesoDetalleParametro
    
       .CursorLocation = adUseClient
       .Fields.Append "CodVistaProceso", adChar, 3
       .Fields.Append "CodVistaUsuario", adChar, 3
       .Fields.Append "DescripVistaUsuario", adVarChar, 200
       .Fields.Append "CodParametroVistaProceso", adVarChar, 200
       .Fields.Append "SecVistaProceso", adInteger
       .Fields.Append "SecVistaUsuario", adInteger
'       .CursorType = adOpenStatic

       .LockType = adLockBatchOptimistic
    End With
    
    adoVistaProcesoDetalleParametro.Open
    
End Sub

Private Sub CargarGrillaCampos(ByVal strSQL As String)

    Dim adoRegistro As ADODB.Recordset
    
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
    
        .CommandText = strSQL
        
        Set adoRegistro = .Execute
        
        Do Until adoRegistro.EOF
            
            adoVistaProcesoDetalle.AddNew
            adoVistaProcesoDetalle.Fields("CodVistaProceso") = adoRegistro.Fields("CodVistaProceso")
            adoVistaProcesoDetalle.Fields("CodVistaUsuario") = adoRegistro.Fields("CodVistaUsuario")
            adoVistaProcesoDetalle.Fields("DescripVistaUsuario") = adoRegistro.Fields("DescripVistaUsuario")
            adoVistaProcesoDetalle.Fields("TipoVista") = adoRegistro.Fields("TipoVista")
            
            adoVistaProcesoDetalle.Update

            adoRegistro.MoveNext
            
        Loop

        End With

        tdgVistaProcesoDetalle.DataSource = adoVistaProcesoDetalle

End Sub

Private Sub CargarGrillaCamposParametro(ByVal strSQL As String)

    Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
    
        .CommandText = strSQL
        
        Set adoRegistro = .Execute
        
        Do Until adoRegistro.EOF
            
            adoVistaProcesoDetalleParametro.AddNew
            adoVistaProcesoDetalleParametro.Fields("CodVistaProceso") = adoRegistro.Fields("CodVistaProceso")
            adoVistaProcesoDetalleParametro.Fields("CodVistaUsuario") = adoRegistro.Fields("CodVistaUsuario")
            adoVistaProcesoDetalleParametro.Fields("DescripVistaUsuario") = adoRegistro.Fields("DescripVistaUsuario")
            adoVistaProcesoDetalleParametro.Fields("CodParametroVistaProceso") = adoRegistro.Fields("CodParametroVistaProceso")
            adoVistaProcesoDetalleParametro.Fields("SecVistaProceso") = adoRegistro.Fields("SecVistaProceso")
            adoVistaProcesoDetalleParametro.Fields("SecVistaUsuario") = adoRegistro.Fields("SecVistaUsuario")
            
            adoVistaProcesoDetalleParametro.Update

            adoRegistro.MoveNext
            
        Loop

        End With

        tdgVistaProcesoDetalleParametro.DataSource = adoVistaProcesoDetalleParametro

End Sub

Private Sub AdicionarParametrosGrilla()

    Dim arrVistasEnGrilla() As String
    Dim adoRegistroAux As ADODB.Recordset
    Dim intIndice As Integer, i As Integer, j As Integer
    Dim arrayInicializado As Boolean
    Dim intSecuencia As Integer
    Dim arrNombreParametrosCargados() As String
    Dim arrSecuenciaParametrosCargados() As String
    
    intSecuencia = 1
    intIndice = 0
    arrayInicializado = False
    
    If Not adoVistaProcesoDetalle.EOF Then
    
        'BORRO REGISTROS DE LA GRILLA DE PARAMETROS
        
        If Not adoVistaProcesoDetalleParametro.EOF Then
        
            adoVistaProcesoDetalleParametro.MoveFirst
            
            Do Until adoVistaProcesoDetalleParametro.EOF
            
                adoVistaProcesoDetalleParametro.Delete
                
                adoVistaProcesoDetalleParametro.MoveNext
                
            Loop
        
        End If
        '***RECUPERO TODAS LAS VISTAS CARGADAS EN LA GRILLA ACTUALMENTE
        Set adoRegistroAux = adoVistaProcesoDetalle.Clone
        
        If Not adoRegistroAux.EOF Then
        
            adoRegistroAux.MoveFirst
            
            Do Until adoRegistroAux.EOF
                
                ReDim Preserve arrVistasEnGrilla(intIndice)
                
                arrVistasEnGrilla(intIndice) = adoRegistroAux.Fields("CodVistaUsuario")
                
                intIndice = intIndice + 1
                
                adoRegistroAux.MoveNext
                
            Loop
            
            
            adoRegistroAux.Close: Set adoRegistroAux = Nothing
        
        
            intIndice = 0
            '***RECUPERO LOS PARAMETROS DE LAS VISTAS OBTENIDAS
            
            For i = 0 To UBound(arrVistasEnGrilla)
            
            Set adoRegistroAux = New ADODB.Recordset
            
                With adoComm
                
                    .CommandText = "SELECT VUC.CodVistaUsuario AS CodVistaUsuario,DescripVista AS DescripVistaUsuario," & _
                                    "NombreCampo,SecCampo FROM VistaUsuarioCampo VUC " & _
                                    "JOIN VistaUsuario VU ON (VUC.CodVistaUsuario=VU.CodVistaUsuario) " & _
                                    "WHERE TipoCampo='01' AND VUC.CodVistaUsuario='" & arrVistasEnGrilla(i) & _
                                    "' ORDER BY SecCampo"
                     
                    Set adoRegistroAux = .Execute
                    
                    adoRegistroAux.MoveFirst
                    
                    Do Until adoRegistroAux.EOF
                    
                        With adoVistaProcesoDetalleParametro
                        
                            adoVistaProcesoDetalleParametro.AddNew
                            
                            adoVistaProcesoDetalleParametro.Fields("CodVistaProceso") = lblCodVistaProceso.Caption
                            adoVistaProcesoDetalleParametro.Fields("CodVistaUsuario") = adoRegistroAux.Fields("CodVistaUsuario")
                            adoVistaProcesoDetalleParametro.Fields("DescripVistaUsuario") = adoRegistroAux.Fields("DescripVistaUsuario")
                            
                            
                            '**COMPROBAR SI EXISTEN YA EN LA GRILLA,SI ES ASI SE LE ASIGNA EL MISMO
                            'NUMERO DE SECUENCIA PROCESO SINO SE INCREMENTEA NORMALMENTE
                            
                            If arrayInicializado = False Then
                            
                                arrayInicializado = True
                            
                                ReDim Preserve arrNombreParametrosCargados(intIndice)
                                ReDim Preserve arrSecuenciaParametrosCargados(intIndice)
                                
                                adoVistaProcesoDetalleParametro.Fields("CodParametroVistaProceso") = adoRegistroAux.Fields("NombreCampo")
                                
                                arrNombreParametrosCargados(intIndice) = adoRegistroAux.Fields("NombreCampo")
                                
                                adoVistaProcesoDetalleParametro.Fields("SecVistaProceso") = intSecuencia
                                
                                arrSecuenciaParametrosCargados(intIndice) = intSecuencia
                                
                                intSecuencia = intSecuencia + 1
                            
                            Else
                                
                                adoVistaProcesoDetalleParametro.Fields("CodParametroVistaProceso") = adoRegistroAux.Fields("NombreCampo")
                                
                                'primero se le asigna la secuencia normal y se avanza la secuencia
                                adoVistaProcesoDetalleParametro.Fields("SecVistaProceso") = intSecuencia
                                intSecuencia = intSecuencia + 1
                                'si lo encuentra en el arreglo entonces reemplaza la secuencia con la encontrada y retroce la secuencia
                                For j = 0 To UBound(arrNombreParametrosCargados)
                                
                                    If adoRegistroAux.Fields("NombreCampo") = arrNombreParametrosCargados(j) Then
                                    
                                        adoVistaProcesoDetalleParametro.Fields("SecVistaProceso") = arrSecuenciaParametrosCargados(j)
                                        intSecuencia = intSecuencia - 1
                                    End If
                                
                                Next
                                
                                ReDim Preserve arrNombreParametrosCargados(intIndice)
                                ReDim Preserve arrSecuenciaParametrosCargados(intIndice)
                                
                                arrNombreParametrosCargados(intIndice) = adoVistaProcesoDetalleParametro.Fields("CodParametroVistaProceso")
                                arrSecuenciaParametrosCargados(intIndice) = adoVistaProcesoDetalleParametro.Fields("SecVistaProceso")
                                
                                
                            End If
                            
                            adoVistaProcesoDetalleParametro.Fields("SecVistaUsuario") = adoRegistroAux.Fields("SecCampo")
                            
                            adoVistaProcesoDetalleParametro.Update
                            
                            intIndice = intIndice + 1
                            
                            adoRegistroAux.MoveNext
                        
                        End With
                    
                    Loop
                
                End With
                
            adoRegistroAux.Close: Set adoRegistroAux = Nothing
                 
            Next
        
        End If
    
    End If

End Sub

Public Function NuevoCodigo() As String

    Dim adoRegistro As ADODB.Recordset
    Dim strSQL As String
    Dim strCODIGO As String, intCodigo As Integer
  
    NuevoCodigo = Valor_Caracter

    Set adoRegistro = New ADODB.Recordset

    With adoComm

        strSQL = "SELECT MAX(CodVistaProceso) AS CodVistaProceso FROM VistaProceso"

        .CommandText = strSQL

        Set adoRegistro = .Execute

        Do Until adoRegistro.EOF

            strCODIGO = adoRegistro("CodVistaProceso")
            adoRegistro.MoveNext

        Loop

    End With

        intCodigo = CInt(strCODIGO) + 1

        strCODIGO = CStr(intCodigo)

        Select Case Len(strCODIGO)

            Case 1

            strCODIGO = "00" + strCODIGO

            Case 2

            strCODIGO = "0" + strCODIGO

        End Select

    adoRegistro.Close: Set adoRegistro = Nothing

    NuevoCodigo = strCODIGO

End Function

Private Sub Limpiar()

    lblCodVistaProceso.Caption = Valor_Caracter
    txtDescripcion.Text = Valor_Caracter

End Sub

Private Sub LimpiarDetalle()

    cboVistaUsuario.ListIndex = 0
    lblCodVistaUsuario.Caption = Valor_Caracter
    lblTipoVistaUsuario.Caption = Valor_Caracter
    
End Sub

Private Sub NumerarVistaProcesoDetalle()
    
    Dim intSec As Integer
    
    intSec = 0
    
    If Not adoVistaProcesoDetalle.EOF Or adoVistaProcesoDetalle.RecordCount > 0 Then
        adoVistaProcesoDetalle.MoveFirst
        While Not adoVistaProcesoDetalle.EOF
            intSec = intSec + 1
            adoVistaProcesoDetalle("Secuencial").Value = intSec
            adoVistaProcesoDetalle.MoveNext
        Wend
    End If
    
End Sub

Private Function TodoOK() As Boolean

    TodoOK = False
    
    If Trim(txtDescripcion.Text) = Valor_Caracter Then
        MsgBox "La descripcion no puede estar vacia", vbCritical
        txtDescripcion.SetFocus
        Exit Function
    End If
    
    TodoOK = True

End Function

Private Function TodoOkDetalle() As Boolean
    
    Dim adoRegistroAux As ADODB.Recordset
    
    Set adoRegistroAux = New ADODB.Recordset
    
    TodoOkDetalle = False
    
    If cboVistaUsuario.ListIndex <= 0 Then
        MsgBox "Debe seleccionar una opcion", vbCritical
        cboVistaUsuario.SetFocus
        Exit Function
    End If
    
    If Not adoVistaProcesoDetalle.EOF Then
    
       Set adoRegistroAux = adoVistaProcesoDetalle.Clone
       
       If Not adoRegistroAux.EOF Then
       
        adoRegistroAux.MoveFirst
        
        Do Until adoRegistroAux.EOF
         
             If (Trim(lblCodVistaUsuario.Caption) = adoRegistroAux("CodVistaUsuario").Value) Then
                     
                  GoTo Ctrl_Identidad
                     
             End If
             
             adoRegistroAux.MoveNext
         
         Loop
        
        End If
    
    End If
    
    TodoOkDetalle = True
    
    Exit Function
    
Ctrl_Identidad:
    
    MsgBox "Esta vista ya existe en la grilla", vbCritical
    cboVistaUsuario.SetFocus

End Function
