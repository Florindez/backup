VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios"
   ClientHeight    =   4665
   ClientLeft      =   1185
   ClientTop       =   1365
   ClientWidth     =   8040
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
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4665
   ScaleWidth      =   8040
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   6360
      TabIndex        =   2
      Top             =   3840
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
      Left            =   480
      TabIndex        =   1
      Top             =   3840
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
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   4830
      Top             =   3720
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabUsuario 
      Height          =   3645
      Left            =   0
      TabIndex        =   9
      Top             =   60
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6429
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
      TabPicture(0)   =   "frmUsuarios.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmUsuarios.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdActualizar"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraDatos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdAccion"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -70200
         TabIndex        =   8
         Top             =   2760
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
      Begin VB.Frame fraDatos 
         Caption         =   "Datos"
         Height          =   2175
         Left            =   -74760
         TabIndex        =   11
         Top             =   480
         Width           =   7575
         Begin VB.ComboBox cboPromotor 
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
            Left            =   2430
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   3825
         End
         Begin VB.TextBox txtDescripUsuario 
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
            Left            =   2430
            MaxLength       =   40
            TabIndex        =   6
            Top             =   1155
            Width           =   3780
         End
         Begin VB.TextBox txtIdUsuario 
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
            Left            =   2430
            MaxLength       =   12
            TabIndex        =   5
            Text            =   " "
            Top             =   840
            Width           =   1860
         End
         Begin VB.CheckBox chkPromotor 
            Caption         =   "Promotor"
            Height          =   195
            Left            =   1080
            TabIndex        =   3
            Top             =   375
            Width           =   1095
         End
         Begin VB.ComboBox cboPerfil 
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
            Left            =   2430
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1560
            Width           =   3825
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Nombres"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   14
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            Caption         =   "ID Usuario"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   13
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Perfil"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   12
            Top             =   1575
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "&Permisos"
         Height          =   735
         Left            =   -71880
         Picture         =   "frmUsuarios.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2760
         Visible         =   0   'False
         Width           =   1200
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmUsuarios.frx":05E5
         Height          =   2655
         Left            =   240
         OleObjectBlob   =   "frmUsuarios.frx":05FF
         TabIndex        =   0
         Top             =   600
         Width           =   7575
      End
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrPromotor()               As String, arrPerfil()                  As String
Dim strCodPromotor              As String, strCodPerfil                 As String
Dim strEstado                   As String

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



Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabUsuario
        .TabEnabled(0) = True
        .Tab = 0
    End With
    Call Buscar
    
End Sub

Private Sub CargarListas()

    Dim strSQL As String
                  
    '*** Promotor ***
    strSQL = "SELECT CodPersona CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Relacionado & "' AND " & _
        "CodSucursal<>'999' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboPromotor, arrPromotor(), Sel_Defecto
            
    '*** Perfil ***
    strSQL = "SELECT CodPerfil CODIGO,DescripPerfil DESCRIP FROM Funcion ORDER BY DescripPerfil"
    CargarControlLista strSQL, cboPerfil, arrPerfil(), Sel_Defecto
        
End Sub
Private Sub Deshabilita()

    
End Sub

Public Sub Eliminar()

    If gstrLogin <> "sa" Then
        MsgBox "Solo el Administrador de la Base de Datos puede eliminar usuarios", vbCritical, gstrNombreEmpresa
        Exit Sub
    End If
    
    If strEstado = Reg_Edicion Or strEstado = Reg_Consulta Then
        frmMainMdi.stbMdi.Panels(3).Text = "Eliminar usuario..."

        If MsgBox("Eliminar el Usuario (" & Trim(tdgConsulta.Columns(0)) & ") ?", vbQuestion + vbYesNo + vbDefaultButton2, gstrNombreEmpresa) = vbYes Then
            If Trim(tdgConsulta.Columns(0)) = "sa" Then
                MsgBox "No puede eliminar al administrador de la Base de Datos", vbCritical, gstrNombreEmpresa
                frmMainMdi.stbMdi.Panels(3).Text = "Acción"
                Exit Sub
            End If

            frmMainMdi.stbMdi.Panels(3).Text = "Eliminando usuario..."
            With adoComm
                '*** Eliminando de la BD ***
                .CommandText = "EXEC sp_dropuser " & Trim(tdgConsulta.Columns(0))
                adoConn.Execute .CommandText

                '*** Eliminando de la BD ***
                .CommandText = "EXEC sp_droplogin " & Trim(tdgConsulta.Columns(0))
                adoConn.Execute .CommandText

                .CommandText = "DELETE UsuarioSistema WHERE IdUsuario='" & Trim(tdgConsulta.Columns(0)) & "'"
                adoConn.Execute .CommandText
            End With

            frmMainMdi.stbMdi.Panels(3).Text = "Usuario eliminado..."
            
            MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation, "Observación"

            tabUsuario.Tab = 0
            Call Buscar
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        Else
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        End If
    End If
            
End Sub


Public Sub Grabar()

    Dim intAccion   As Integer, lngNumError     As Long
    Dim sqlins      As String
    Dim nac1        As String
            
    If strEstado = Reg_Consulta Then Exit Sub
    
    On Error GoTo CtrlError
    
    If strEstado = Reg_Adicion Or strEstado = Reg_Edicion Then
        If TodoOK() Then
            frmMainMdi.stbMdi.Panels(3).Text = "Grabar Usuario..."
            
            Me.MousePointer = vbHourglass
            With adoComm
                '*** Agrega nuevo login en la BD ***
                '*** Por default el Id del Usuario es igual al login_id         ***
                '*** Por default el grupo asignado al ID del Usuario es fondos  ***
                '*** Por default el password asignado es igual al login_id      ***
                .CommandText = "{ call up_ACProcUsuario('" & _
                    Trim(txtIdUsuario.Text) & "','" & Trim(txtIdUsuario.Text) & "','" & Trim(txtDescripUsuario.Text) & "','" & _
                    strCodPerfil & "','" & strCodPromotor & "','" & _
                    Estado_Activo & "','" & _
                    gstrDataBase & "','" & IIf(strEstado = Reg_Adicion, "I", "U") & "') }"
                adoConn.Execute .CommandText
                
            End With
        
            Me.MousePointer = vbDefault
                            
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabUsuario
                .TabEnabled(0) = True
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
        
    TodoOK = False
    
    If Trim(txtIdUsuario) = Valor_Caracter Then
        MsgBox "El Campo ID Usuario no es Válido!.", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
    
    If Trim(txtDescripUsuario) = Valor_Caracter Then
        MsgBox "El Campo Nombres no es Valido!.", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
    
    If chkPromotor.Value Then
        If strCodPromotor = Valor_Caracter Then
            MsgBox "Debe Seleccionar el Promotor!.", vbCritical, gstrNombreEmpresa
            Exit Function
        End If
    End If
                
    If strCodPerfil = Valor_Caracter Then
        MsgBox "Debe Seleccionar el perfil de acceso del usuario.", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
        
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function
Private Sub Habilita()

    
    
End Sub


Public Sub Imprimir()

End Sub

Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabUsuario
            .TabEnabled(0) = False
            .Tab = 1
        End With
        Call Habilita
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro     As ADODB.Recordset
    Dim strCodUsuario   As String
    Dim intRegistro     As Integer
    
    Select Case strModo
        Case Reg_Adicion
            chkPromotor.Value = vbChecked: chkPromotor.Enabled = False
            cboPromotor.ListIndex = -1
            If cboPromotor.ListCount > 0 Then cboPromotor.ListIndex = 0
            
            txtIdUsuario.Text = Valor_Caracter
            txtIdUsuario.Enabled = True
            txtDescripUsuario.Text = Valor_Caracter
                        
            cboPerfil.ListIndex = -1
            If cboPerfil.ListCount > 0 Then cboPerfil.ListIndex = 0
            
            txtIdUsuario.SetFocus
                        
        Case Reg_Edicion
            Set adoRegistro = New ADODB.Recordset

            strCodUsuario = Trim(tdgConsulta.Columns(0))

            adoComm.CommandText = "SELECT * FROM UsuarioSistema WHERE IdUsuario='" & strCodUsuario & "'"
            Set adoRegistro = adoComm.Execute

            If Not adoRegistro.EOF Then
                chkPromotor.Value = vbChecked: chkPromotor.Enabled = False
                intRegistro = ObtenerItemLista(arrPromotor(), adoRegistro("CodPromotor"))
                If intRegistro >= 0 Then cboPromotor.ListIndex = intRegistro
                
                txtIdUsuario.Text = strCodUsuario
                txtIdUsuario.Enabled = False
                txtDescripUsuario.Text = Trim(adoRegistro("DescripUsuario"))

                intRegistro = ObtenerItemLista(arrPerfil(), adoRegistro("CodPerfil"))
                If intRegistro >= 0 Then cboPerfil.ListIndex = intRegistro
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
    
    End Select
    
End Sub
Public Sub Salir()

    Unload Me
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Listado de Usuarios"
    
End Sub


Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
            
End Sub



Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String

    If tabUsuario.Tab = 1 Then Exit Sub
    
    Select Case Index
        Case 1
            gstrNameRepo = "UsuarioSistema"
                        
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
                        
            aReportParamS(0) = "001"
            aReportParamS(1) = gstrCodAdministradora
            
    End Select

    gstrSelFrml = ""
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub




Private Sub cboPerfil_Click()

    strCodPerfil = Valor_Caracter
    If cboPerfil.ListIndex < 0 Then Exit Sub
  
    strCodPerfil = Trim(arrPerfil(cboPerfil.ListIndex))
    
End Sub


Private Sub cboPromotor_Click()

    strCodPromotor = Valor_Caracter
    If cboPromotor.ListIndex < 0 Then Exit Sub
  
    strCodPromotor = Trim(arrPromotor(cboPromotor.ListIndex))
    txtDescripUsuario.Text = Trim(cboPromotor.Text)
    
End Sub


Private Sub chkPromotor_Click()

    If chkPromotor.Value Then
        cboPromotor.Visible = True
    Else
        cboPromotor.Visible = False
    End If
    
End Sub

Private Sub cmdActualizar_Click()

    If Len(Trim(txtIdUsuario.Text)) = 0 Then
        MsgBox "Seleccione un Usuario ", vbCritical, "Observación"
        Exit Sub
    Else
        If MsgBox("Actualizar los Permisos Asignados a " & Trim(txtIdUsuario.Text) & "?.", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            LPermUsua Trim(txtIdUsuario.Text)
        End If
    End If
            
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
    
    CentrarForm Me

End Sub

Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabUsuario.Tab = 0
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 18
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 45
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 25
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmUsuarios = Nothing
        
End Sub





Public Sub Buscar()
        
    Dim strSQL As String
    
    strSQL = "SELECT IdUsuario,DescripUsuario,U.CodPerfil,CodPromotor,DescripPerfil " & _
        "FROM UsuarioSistema U JOIN Funcion F ON(F.CodPerfil=U.CodPerfil) " & _
        "WHERE EstadoUsuario='" & Estado_Activo & "' AND IdUsuario<>'sa'" & _
        "ORDER BY DescripUsuario"
                        
    strEstado = Reg_Defecto
    With adoConsulta
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With
        
    tdgConsulta.Refresh
    
    If adoConsulta.Recordset.RecordCount > 0 Then strEstado = Reg_Consulta
    
End Sub

Private Sub LPermUsua(s_Usuario As String)

    Dim adoresult As ADODB.Recordset
    Dim strSQL As String
    
    Set adoresult = New ADODB.Recordset
    
    With adoComm
        strSQL = "SELECT name  FROM sysobjects WHERE type='U' or type='P' ORDER BY type,name"
        .CommandText = strSQL
        Set adoresult = .Execute
        Do Until adoresult.EOF
            .CommandText = "GRANT ALL ON " & Trim(adoresult!Name) & " TO " & s_Usuario
            .Execute
            
            adoresult.MoveNext
        Loop
        adoresult.Close: Set adoresult = Nothing
    End With
    
End Sub




Public Sub Adicionar()

    If gstrLogin <> "sa" Then
        MsgBox "Solo el Administrador de la Base de Datos puede crear usuarios", vbCritical, gstrNombreEmpresa
        Exit Sub
    End If
    
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Usuario..."
                
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabUsuario
        .TabEnabled(0) = False
        .Tab = 1
    End With
    Call Habilita
                
End Sub

Private Sub tabUsuario_Click(PreviousTab As Integer)

    Select Case tabUsuario.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabUsuario.Tab = 0
        
    End Select
    
End Sub

Private Sub txtIdUsuario_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(LCase(Chr(KeyAscii)))
    
End Sub


