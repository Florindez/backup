VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmPerfil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perfiles"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   6885
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   5160
      TabIndex        =   9
      Top             =   4560
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
      Left            =   480
      TabIndex        =   8
      Top             =   4560
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
      Caption2        =   "&Eliminar"
      Tag2            =   "4"
      Visible2        =   0   'False
      ToolTipText2    =   "Eliminar"
      UserControlWidth=   4200
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   3960
      Top             =   4200
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
   Begin TabDlg.SSTab tabPerfil 
      Height          =   4245
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7488
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
      TabPicture(0)   =   "frmPerfil.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmPerfil.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDatos"
      Tab(1).Control(1)=   "adoModulo"
      Tab(1).Control(2)=   "cmdAccion"
      Tab(1).ControlCount=   3
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -71640
         TabIndex        =   10
         Top             =   3360
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
      Begin MSAdodcLib.Adodc adoModulo 
         Height          =   330
         Left            =   -74160
         Top             =   3360
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
      Begin VB.Frame fraDatos 
         Caption         =   "Datos"
         Height          =   2775
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   5895
         Begin TrueOleDBGrid60.TDBGrid tdgModulo 
            Bindings        =   "frmPerfil.frx":0038
            Height          =   1455
            Left            =   1800
            OleObjectBlob   =   "frmPerfil.frx":0050
            TabIndex        =   11
            Top             =   1080
            Width           =   3735
         End
         Begin VB.CheckBox chkAdministrador 
            Caption         =   "Administrador"
            Height          =   195
            Left            =   240
            TabIndex        =   4
            ToolTipText     =   "Marcar para seleccionar todos los módulos"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtCodPerfil 
            Height          =   285
            Left            =   1800
            MaxLength       =   12
            TabIndex        =   3
            Text            =   " "
            Top             =   360
            Width           =   1020
         End
         Begin VB.TextBox txtDescripPerfil 
            Height          =   285
            Left            =   1800
            MaxLength       =   40
            TabIndex        =   2
            Top             =   675
            Width           =   3740
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   5
            Top             =   720
            Width           =   840
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmPerfil.frx":28A2
         Height          =   3015
         Left            =   360
         OleObjectBlob   =   "frmPerfil.frx":28BC
         TabIndex        =   7
         Top             =   600
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmPerfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strEstado           As String

Private Sub BuscarModulos()

    Dim strSQL As String
    
    strSQL = "SELECT CodModulo,DescripModulo " & _
        "FROM ModuloSistema " & _
        "ORDER BY DescripModulo"
                            
    With adoModulo
        .ConnectionString = gstrConnectConsulta
        .CommandTimeout = 0
        .ConnectionTimeout = 0
        .RecordSource = strSQL
        .Refresh
    End With
        
    tdgModulo.Refresh
            
End Sub

Private Sub chkAdministrador_Click()

    If chkAdministrador.Value Then
    
        adoModulo.Recordset.MoveFirst
        Do While Not adoModulo.Recordset.EOF
            tdgModulo.SelBookmarks.Add adoModulo.Recordset.Bookmark
        
            adoModulo.Recordset.MoveNext
        Loop
    Else
        Dim intContador     As Integer, intRegistro     As Integer
        
        If tdgModulo.SelBookmarks.Count = 0 Then Exit Sub
        
        intContador = tdgModulo.SelBookmarks.Count - 1

        For intRegistro = 0 To intContador
            tdgModulo.SelBookmarks.Remove (intContador - intRegistro)
            tdgModulo.Refresh
        Next
                
    End If
    
End Sub

Private Sub Form_Activate()

    Call CargarReportes
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Listado de Perfiles"
    
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

Public Sub Salir()

    Unload Me
    
End Sub
Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabPerfil
        .TabEnabled(0) = True
        .Tab = 0
    End With
    strEstado = Reg_Consulta
    
End Sub
Public Sub Grabar()
                
    Dim intContador     As Integer, intRegistro     As Integer
    Dim intAccion       As Integer, lngNumError     As Long
    Dim strPerfilAcceso As String
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    On Error GoTo CtrlError
    
    If strEstado = Reg_Adicion Then
        If TodoOK() Then
            frmMainMdi.stbMdi.Panels(3).Text = "Grabar Perfil..."
            
            Me.MousePointer = vbHourglass
            
            intContador = tdgModulo.SelBookmarks.Count - 1

            For intRegistro = 0 To intContador
                tdgModulo.Row = tdgModulo.SelBookmarks(intRegistro) - 1
                tdgModulo.Refresh

                strPerfilAcceso = strPerfilAcceso & Trim(tdgModulo.Columns(0))

            Next
            
            With adoComm
                .CommandText = "{ call up_GNManFuncion('" & _
                    Trim(txtCodPerfil.Text) & "','" & Trim(txtDescripPerfil.Text) & "','" & _
                    strPerfilAcceso & "','" & _
                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','I') }"
                adoConn.Execute .CommandText
            End With
        
            Me.MousePointer = vbDefault
                            
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabPerfil
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
    End If
    
    If strEstado = Reg_Edicion Then
        If TodoOK() Then
            frmMainMdi.stbMdi.Panels(3).Text = "Grabar Perfil..."
            
            Me.MousePointer = vbHourglass
            
            intContador = tdgModulo.SelBookmarks.Count - 1

            For intRegistro = 0 To intContador
                tdgModulo.Row = tdgModulo.SelBookmarks(intRegistro) - 1
                tdgModulo.Refresh

                strPerfilAcceso = strPerfilAcceso & Trim(tdgModulo.Columns(0))

            Next
            
            With adoComm
                .CommandText = "{ call up_GNManFuncion('" & _
                    Trim(txtCodPerfil.Text) & "','" & Trim(txtDescripPerfil.Text) & "','" & _
                    strPerfilAcceso & "','" & _
                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','U') }"
                adoConn.Execute .CommandText
            End With
        
            Me.MousePointer = vbDefault
                            
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabPerfil
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
    
'    If Trim(txtIdUsuario) = Valor_Caracter Then
'        MsgBox "El Campo ID Usuario no es Válido!.", vbCritical, gstrNombreEmpresa
'        Exit Function
'    End If
'
'    If Trim(txtDescripUsuario) = Valor_Caracter Then
'        MsgBox "El Campo Nombres no es Valido!.", vbCritical, gstrNombreEmpresa
'        Exit Function
'    End If
'
'    If chkPromotor.Value Then
'        If strCodPromotor = Valor_Caracter Then
'            MsgBox "Debe Seleccionar el Promotor!.", vbCritical, gstrNombreEmpresa
'            Exit Function
'        End If
'    End If
'
'    If strCodPerfil = Valor_Caracter Then
'        MsgBox "Debe Seleccionar el perfil de acceso del usuario.", vbCritical, gstrNombreEmpresa
'        Exit Function
'    End If
        
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function
Public Sub Imprimir()

End Sub

Public Sub Buscar()
        
    Dim strSQL As String
    
    strSQL = "SELECT CodPerfil,DescripPerfil,PerfilAcceso " & _
        "FROM Funcion " & _
        "ORDER BY DescripPerfil"
                        
    strEstado = Reg_Defecto
    With adoConsulta
        .ConnectionString = gstrConnectConsulta
        .CommandTimeout = 0
        .ConnectionTimeout = 0
        .RecordSource = strSQL
        .Refresh
    End With
        
    tdgConsulta.Refresh
    
    If adoConsulta.Recordset.RecordCount > 0 Then strEstado = Reg_Consulta
    
End Sub

Public Sub Eliminar()
        
    If strEstado = Reg_Edicion Or strEstado = Reg_Consulta Then
        frmMainMdi.stbMdi.Panels(3).Text = "Eliminar perfil..."

        If MsgBox("Eliminar el Perfil (" & Trim(tdgConsulta.Columns(1)) & ") ?", vbQuestion + vbYesNo + vbDefaultButton2, gstrNombreEmpresa) = vbYes Then

            frmMainMdi.stbMdi.Panels(3).Text = "Eliminando perfil..."
            With adoComm
                .CommandText = "DELETE Funcion WHERE CodPerfil='" & Trim(tdgConsulta.Columns(0)) & "'"
                adoConn.Execute .CommandText
            End With

            frmMainMdi.stbMdi.Panels(3).Text = "Perfil eliminado..."
            
            MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation, Me.Caption

            tabPerfil.Tab = 0
            Call Buscar
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        Else
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        End If
    End If
            
End Sub
Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabPerfil
            .TabEnabled(0) = False
            .Tab = 1
        End With
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro     As ADODB.Recordset
    Dim strCodPerfil   As String
    Dim intRegistro     As Integer
    
    Select Case strModo
        Case Reg_Adicion
            chkAdministrador.Value = vbUnchecked
            
            txtCodPerfil.Text = Valor_Caracter
            txtCodPerfil.Enabled = True
            txtDescripPerfil.Text = Valor_Caracter
                                                
            Call BuscarModulos
            
            txtCodPerfil.SetFocus
                        
        Case Reg_Edicion
            Set adoRegistro = New ADODB.Recordset

            strCodPerfil = Trim(tdgConsulta.Columns(0))

            adoComm.CommandText = "SELECT * FROM Funcion WHERE CodPerfil='" & strCodPerfil & "'"
            Set adoRegistro = adoComm.Execute

            If Not adoRegistro.EOF Then
                chkAdministrador.Value = vbUnchecked
                txtCodPerfil.Text = strCodPerfil
                txtCodPerfil.Enabled = False
                txtDescripPerfil.Text = Trim(adoRegistro("DescripPerfil"))

                Call BuscarModulos
                
                If InStr(adoRegistro("PerfilAcceso"), "A") = 1 Then
                    chkAdministrador.Value = vbChecked
                Else
                    adoModulo.Recordset.MoveFirst
                    Do While Not adoModulo.Recordset.EOF
                        If InStr(adoRegistro("PerfilAcceso"), adoModulo.Recordset.Fields("CodModulo").Value) > 0 Then
                            tdgModulo.SelBookmarks.Add adoModulo.Recordset.Bookmark
                        End If
                                            
                        adoModulo.Recordset.MoveNext
                    Loop
                End If
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
    
    End Select
    
End Sub

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String

    If tabPerfil.Tab = 1 Then Exit Sub
    
    Select Case Index
        Case 1
            gstrNameRepo = "Funcion"
                        
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
Public Sub Adicionar()
    
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Perfil..."
                
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabPerfil
        .TabEnabled(0) = False
        .Tab = 1
    End With
                
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
Private Sub CargarListas()
    
        
End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabPerfil.Tab = 0
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 16
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 60
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmPerfil = Nothing
    
End Sub


Private Sub tabPerfil_Click(PreviousTab As Integer)

    Select Case tabPerfil.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabPerfil.Tab = 0
        
    End Select
    
End Sub

