VERSION 5.00
Begin VB.Form frmAcceso 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acceso"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   ControlBox      =   0   'False
   Icon            =   "frmAcceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   860.317
   ScaleMode       =   0  'User
   ScaleWidth      =   1355.114
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3960
      Left            =   30
      TabIndex        =   7
      Top             =   47
      Width           =   4665
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3105
         Picture         =   "frmAcceso.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3030
         Width           =   1395
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "Con&traseña"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   140
         Picture         =   "frmAcceso.frx":227C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3030
         Width           =   1395
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "Co&nectar"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   1620
         Picture         =   "frmAcceso.frx":2845
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3030
         Width           =   1395
      End
      Begin VB.TextBox txtUsuario 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1710
         MaxLength       =   12
         TabIndex        =   8
         Top             =   2190
         Width           =   2535
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   2580
         Width           =   2535
      End
      Begin VB.Image imgLogo 
         BorderStyle     =   1  'Fixed Single
         Height          =   1860
         Left            =   225
         Picture         =   "frmAcceso.frx":2DDC
         Stretch         =   -1  'True
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   510
         TabIndex        =   11
         Top             =   2250
         Width           =   660
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   510
         TabIndex        =   10
         Top             =   2640
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4680
      TabIndex        =   5
      Top             =   8880
      Visible         =   0   'False
      Width           =   1400
   End
   Begin VB.ComboBox cboAdministradora 
      Height          =   315
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   8040
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.ComboBox cboTipoAdministradora 
      Height          =   315
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   7440
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6480
      TabIndex        =   2
      Top             =   8880
      Visible         =   0   'False
      Width           =   1400
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   4320
      Top             =   4470
      Width           =   2520
   End
   Begin VB.Label lblNombreSistema 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spectrum Fondos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   555
      Left            =   1920
      TabIndex        =   6
      Top             =   4320
      Width           =   4065
   End
   Begin VB.Label lblDescrip 
      Caption         =   "Administradora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   1
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDescrip 
      Caption         =   "Tipo Administradora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image imgTAM 
      Height          =   4665
      Left            =   840
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   405
   End
End
Attribute VB_Name = "frmAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Ficha de Acceso General"
Option Explicit

Dim arrAdministradora()     As String, arrTipoAdministradora() As String, Ind As Integer
Dim strPasar                 As String

Private Sub CargarTipoAdministradora()

    Dim strSQL As String
                        
    '*** Tipo de Administradora ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPADM' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoAdministradora, arrTipoAdministradora(), Sel_Defecto
    
    If cboTipoAdministradora.ListCount > 0 Then cboTipoAdministradora.ListIndex = 0
    
End Sub

Private Function ContinuarAcceso() As Boolean

    Dim adoFondo            As ADODB.Recordset, adoUsuario      As ADODB.Recordset
    Dim adoCuotas           As ADODB.Recordset, adoRegistroTmp  As ADODB.Recordset
    Dim adoRegistro         As ADODB.Recordset, adoRegistroAux  As ADODB.Recordset
    Dim adoExisteUsuario    As ADODB.Recordset
    Dim strPerfil           As String, Msg                      As String
    Dim res                 As Integer, intDias                 As Integer
    Dim strDescripUsuario   As String ', strPasar                 As String
    Dim vntFchUser          As Variant
            
    ContinuarAcceso = False
    
    strPasar = "NO"
    
    
    With adoCommSeguridad
        Set adoUsuario = New ADODB.Recordset
        Set adoExisteUsuario = New ADODB.Recordset
        
        '---/// Validacion de Usuario Seguridad
        
        .CommandText = "SELECT IdUsuario,dbo.uf_SELeerClaveEnch(Passw),DescripUsuario FROM UsuarioSistema WHERE IdUsuario='" & gstrLoginUS & "'"
        Set adoExisteUsuario = .Execute

        If adoExisteUsuario.EOF Then
            adoExisteUsuario.Close: Set adoExisteUsuario = Nothing
            MsgBox "El Usuario no Existe en la Base de Datos", vbCritical, Me.Caption
            Exit Function
        End If
        
        strDescripUsuario = Trim(adoExisteUsuario("DescripUsuario"))
                        
        frmMainMdi.stbMdi.Panels(3).Text = "Verificando accesos..."
        
        '*** Verificación de Accesos de Usuario ***
        strPerfil = frmMainMdi.Tag
        
        '---/// Integración de seguridad
        If Not ValidarPermisoAccesoObjeto(Trim(gstrLoginUS), Trim(App.Title), Codigo_Tipo_Objeto_Modulo) Then
            Exit Function
        End If
        
        '*** Verificacion de la Version del Módulo ***
        frmMainMdi.stbMdi.Panels(3).Text = "Verificando versión..."
        Call VerifVersion
        
        adoExisteUsuario.Close: Set adoExisteUsuario = Nothing
                
        frmMainMdi.txtUsuarioSistema.Text = Trim(strDescripUsuario)
        
        With frmMainMdi.stbMdi
            .Panels(1).Text = gstrNombreEmpresa & Space(1)
            .Panels(2).Text = gstrNombreSistema & Space(1) & gstrVersionSistema & Space(1)
            .Panels(3).Text = "Acción"
        End With
                
        'Unload frmAcceso
                                    
    End With
    
    ContinuarAcceso = True
    
End Function

Private Sub HabilitarControles(blnValor As Boolean)

    cmdAccion(0).Enabled = Not blnValor
    cmdAccion(1).Enabled = Not blnValor
    txtUsuario.Enabled = Not blnValor
    txtPassword.Enabled = Not blnValor
    'cboTipoAdministradora.Enabled = blnValor
    'cboAdministradora.Enabled = blnValor
    'cmdAceptar.Enabled = blnValor
    
End Sub

Private Function ValidarIngreso() As Boolean

    Dim blnValida As Boolean
    
    blnValida = True
    
    If Trim(cboTipoAdministradora.Text) = Sel_Defecto Then
        MsgBox "Seleccione el Tipo de Administradora", vbExclamation, gstrNombreEmpresa
        ValidarIngreso = Not blnValida
        Exit Function
    End If

    If Trim(cboAdministradora.Text) = Sel_Defecto Then
        MsgBox "Seleccione la Administradora", vbExclamation, gstrNombreEmpresa
        ValidarIngreso = Not blnValida
        Exit Function
    End If
    
    ValidarIngreso = blnValida
    
End Function

Private Sub VerifVersion()

    Dim strVerArch   As String
    Dim adoRecord    As ADODB.Recordset
    Dim strCodModulo    As String
    
    strCodModulo = "05"
    
    Set adoRecord = New ADODB.Recordset
    
    adoCommSeguridad.CommandText = "SELECT VersionModulo FROM ModuloSistema WHERE CodModulo='" & strCodModulo & "'"   '& frmMainMdi.Tag &
    Set adoRecord = adoCommSeguridad.Execute
    
    If Not adoRecord.EOF Then
        strVerArch = Trim(adoRecord("VersionModulo"))
                
        If ((App.Major & "." & App.Minor & "." & App.Revision) <> strVerArch) Then
            MsgBox "Advertencia, la versión que tiene registrada está desactualizada." & Chr(10) & Chr(13) & "Comunicarse con el Area de Sistemas.", vbCritical, gstrNombreEmpresa
            adoRecord.Close: Set adoRecord = Nothing
            Unload frmAcceso
            End
        End If
    End If
    adoRecord.Close: Set adoRecord = Nothing
        
End Sub

Private Sub cboAdministradora_Click()

    gstrCodAdministradora = ""
    If cboAdministradora.ListIndex < 0 Then Exit Sub
    
    gstrCodAdministradora = arrAdministradora(cboAdministradora.ListIndex)
    
End Sub

Private Sub cboTipoAdministradora_Click()

    Dim strSQL As String
    
    gstrTipoAdministradora = ""
    If cboTipoAdministradora.ListIndex < 0 Then Exit Sub
    
    gstrTipoAdministradora = arrTipoAdministradora(cboTipoAdministradora.ListIndex)
        
    strSQL = "SELECT CodAdministradora CODIGO,DescripAdministradora DESCRIP FROM Administradora WHERE CodTipoAdministradora='" & gstrTipoAdministradora & "'"
    CargarControlLista strSQL, cboAdministradora, arrAdministradora(), Sel_Defecto
    
    If cboAdministradora.ListCount > 0 Then cboAdministradora.ListIndex = 0
    
End Sub



Private Sub cmdAccion_Click(Index As Integer)
            
    Dim adoAdministradora As ADODB.Recordset '*** NUEVO AFP INTEGRA
    Dim adoCantidadFondo As ADODB.Recordset '*** NUEVO AFP INTEGRA
    Dim adoFondo As ADODB.Recordset '*** NUEVO AFP INTEGRA
    Dim adoUsuario    As ADODB.Recordset
    Dim strCodModulo As String
            
    Dim lastError As Long
    Dim msgbuf As String
    
    strCodModulo = "01"
        
    On Error GoTo cmdConectar_Error

    Select Case Index
    
        Case 0: '*** Aceptar ***
        
            MousePointer = vbHourglass
            Ind = 1
            
            '*** Cargar los parámetros globales de conexión ***

            '*** Conexión a la Base de Datos ***
            Set adoConn = New ADODB.Connection
            Set adoConnSeguridad = New ADODB.Connection
            
                       
            gstrLoginUS = Trim(txtUsuario.Text)
            'gstrRptConnectODBC = "DSN=" & gstrODBCName32 & ";UID=" & gstrLogin & ";PWD=" & gstrPassword & ";App=" & App.Title & ";"
            
            If adoConnSeguridad.State = 1 Then
                adoConnSeguridad.Close:  Set adoConnSeguridad = Nothing
            End If
            
'--------------------------------------------- CONEXION SEGURIDAD
            If adoConn.State = 1 Then
                adoConn.Close:  Set adoConn = Nothing
            End If
                        
            'Conexion a Seguridad
                                         
            '*** SQLOLEDB - Base de Datos ***
            gstrConnectSeguridad = "User ID=" & gstrLoginSeguridad & ";Password=" & gstrPasswSeguridad & ";" & _
                                "Data Source=" & gstrServerSeguridad & ";" & _
                                "Initial Catalog=" & gstrDataBaseSeguridad & ";" & _
                                "Application Name=" & App.Title & ";" & _
                                "Auto Translate=False"
   
            
            frmMainMdi.stbMdi.Panels(3).Text = "Conectando a la Base de Datos..."
            With adoConnSeguridad
                .Provider = "SQLOLEDB"
                .ConnectionString = gstrConnectSeguridad
                .CommandTimeout = 0
                .ConnectionTimeout = 0
                .Open
            End With
            
            frmMainMdi.stbMdi.Panels(3).Text = "Conexión a Seguridad establecida..."
            
            Set adoCommSeguridad = New ADODB.Command
            adoCommSeguridad.CommandTimeout = 0
            Set adoCommSeguridad.ActiveConnection = adoConnSeguridad
            
            
             With adoCommSeguridad
       
                Set adoUsuario = New ADODB.Recordset
        
                '---/// Validacion de Usuario Seguridad
        
                .CommandText = "SELECT SuperAdmin,  dbo.uf_SELeerClaveEnch(SAPassw) SAPassw FROM ModuloConexion WHERE CodModulo='" & strCodModulo & "'"
                Set adoUsuario = .Execute

                gstrLogin = adoUsuario("SuperAdmin"): gstrPassword = adoUsuario("SAPassw")
                
            End With
 '----------------------------------------------------------------------------------------------------------
        
            
            
    '       Comentado y enviado al modulo de Seleccion de Fondos de acuerdo a la mejora pedida por Andres..!!
            gstrLogin = Trim$(txtUsuario.Text)
            gstrPassword = txtPassword.Text
                
            If Not ContinuarAcceso Then
                MousePointer = vbDefault
                txtUsuario.Text = Valor_Caracter: txtPassword.Text = Valor_Caracter
                txtUsuario.SetFocus
                Exit Sub
            End If
                
    '---CONEXION FONDOS----
            gstrConnectConsulta = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & gstrLogin & ";Password=" & gstrPassword & ";" & _
                                  "Data Source=" & gstrServer & ";" & _
                                  "Initial Catalog=" & gstrDataBase & ";"

            '*** SQLOLEDB - Base de Datos ***
            gstrConnect = "User ID=" & gstrLogin & ";Password=" & gstrPassword & ";" & _
                                "Data Source=" & gstrServer & ";" & _
                                "Initial Catalog=" & gstrDataBase & ";" & _
                                "Application Name=" & App.Title & ";" & _
                                "Auto Translate=False"


            gstrConnectNET = "User ID=" & gstrLogin & ";Password=" & gstrPassword & ";Data Source=" & gstrServer & ";Initial Catalog=" & gstrDataBase & ""

            frmMainMdi.stbMdi.Panels(3).Text = "Conectando a la Base de Datos..."
            With adoConn
                .Provider = "SQLOLEDB"
                .ConnectionString = gstrConnect
                .CommandTimeout = 0
                .ConnectionTimeout = 0
                .Open
            End With

'            adoConn.Open gstrRptConnectODBC

            frmMainMdi.stbMdi.Panels(3).Text = "Conexión establecida..."

            Set adoComm = New ADODB.Command
            adoComm.CommandTimeout = 0
            Set adoComm.ActiveConnection = adoConn

                
            '------------------------ Administradora
            With adoComm
                                
                .CommandText = "SELECT CodAdministradora,FechaInicio, DescripAdministradora " & _
                                " FROM Administradora WHERE Estado='01'"
                Set adoAdministradora = .Execute
                
            End With
            
            If adoAdministradora.EOF = False Then
                gstrCodAdministradora = Trim(CStr(adoAdministradora("CodAdministradora")))
                gstrNombreAdministradora = Trim(CStr(adoAdministradora("DescripAdministradora")))
                gstrFechaActual = Convertyyyymmdd(adoAdministradora("FechaInicio"))
                gdatFechaActual = Trim(CDate(adoAdministradora("FechaInicio")))
            Else
                MsgBox "Debe Ingresar la Administradora, de lo contrario el Sistema presentara Errores", vbCritical, Me.Caption
                Exit Sub
            End If
    '----------------------------------------------------------------------------------
    
    '--------------- Verificacion de Existencia de Fondos
        
    
            With adoComm
                        
                .CommandText = "SELECT COUNT(CodFondo) CantidadFondo " & _
                                 " FROM Fondo WHERE Estado='01' AND CodAdministradora = '" & gstrCodAdministradora & "'"
                Set adoCantidadFondo = .Execute
                        
                Dim cant As Integer
                                
                cant = CInt(adoCantidadFondo("CantidadFondo").Value)
                
                If cant <= 0 Then
                    
                    If adoCantidadFondo.EOF = False Then
                        'If App.Title = "General" Then
                        '    MsgBox "Bienvenido a " & gstrNombreSistema & Space(1) & gstrVersionSistema, vbExclamation, gstrNombreEmpresa
                        '    gdatFechaActual = Trim(CDate(adoAdministradora("FechaInicio")))
                        'Else
                            MsgBox "No existen Fondos definidos en el Sistema.", vbExclamation, gstrNombreEmpresa
                            adoCantidadFondo.Close: Set adoCantidadFondo = Nothing
                        'End If
                        Call LDiasNUtil
                        Call CargarParametrosGlobales
                    
                    End If
            
                Else
                    gboolMostrarSelectAdministradora = True
                End If
                
            End With

            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            frmMainMdi.txtEmpresa = gstrNombreAdministradora

            '*****************************
            
        Case 1
            frmCambioClave.Show vbModal
            txtUsuario.Text = Valor_Caracter
            txtPassword.Text = Valor_Caracter
            txtUsuario.SetFocus
            Exit Sub
    End Select
  
cmdConectar_Fin:
    MousePointer = vbDefault
    Unload Me
    Exit Sub
    
cmdConectar_Error:
    MousePointer = vbDefault
    With err
        
        frmMainMdi.stbMdi.Panels(3).Text = "Error de conexión..."
        MsgBox "Error! " & .Description, vbCritical, Me.Caption
        txtPassword.Text = Valor_Caracter
        

    End With
    MousePointer = vbDefault
    txtPassword.SetFocus
    'Resume cmdConectar_Fin
    
End Sub

Private Sub cmdAceptar_Click()

    Dim adoFon As ADODB.Recordset, adoUsr As ADODB.Recordset
    Dim adoCuo As ADODB.Recordset, adoresultTmp As ADODB.Recordset
    Dim adoresult As ADODB.Recordset, adoResultAux As ADODB.Recordset
    Dim strPerfil As String, Msg As String, res As Integer, intDias As Integer
    Dim strDscUser As String, strPasar As String
    Dim vntFchUser As Variant
            
    strPasar = "NO"
    Ind = 1
    If Not ValidarIngreso() Then Exit Sub
    
    With adoComm
        Set adoUsr = New ADODB.Recordset
        
        .CommandText = "SELECT DescripUsuario,CodPromotor,PerfilAcceso FROM UsuarioSistema WHERE IdUsuario='" & gstrLogin & "'"
        Set adoUsr = .Execute
        
        If adoUsr.EOF Then
            adoUsr.Close: Set adoUsr = Nothing
            Exit Sub
        End If
        
        strDscUser = adoUsr("DescripUsuario")
        gstrCodPromotor = Trim(adoUsr("CodPromotor"))
                        
        frmMainMdi.stbMdi.Panels(3).Text = "Verificando accesos..."
        '*** Verificación de Accesos de Usuario y Seteo de Variables Globales ***
        strPerfil = frmMainMdi.Tag
    
        '*** Verificación de autorización de acceso al Módulo        ***
        '*** Los supervisores están autorizados a todos los  Módulos ***
        If InStr(adoUsr("PerfilAcceso"), strPerfil) = 0 And InStr(adoUsr("PerfilAcceso"), "S") = 0 Then
            MsgBox "El acceso a este Módulo no está permitido para este Usuario", vbCritical, gstrNombreEmpresa
            txtUsuario.SetFocus
            txtUsuario.Text = "": txtPassword.Text = ""
            adoUsr.Close: Set adoUsr = Nothing
            Call HabilitarControles(False)
            cboAdministradora.Clear: cboTipoAdministradora.Clear
            Exit Sub
        End If
       
        '*** Verificacion de la Version del Modulo ***
        frmMainMdi.stbMdi.Panels(3).Text = "Verificando versión..."
        Call VerifVersion
        
        '*** Verificación de los fondos existentes ***
        Set adoFon = New ADODB.Recordset
        
        .CommandText = "SELECT CodFondo FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "'"
        Set adoFon = .Execute
        
        If adoFon.EOF Then
            If App.Title = "General" Then
                MsgBox "Bienvenido a " & gstrNombreSistema, vbExclamation, gstrNombreEmpresa
                strPasar = "SI"
            Else
                MsgBox "No existen Fondos definidos en el Sistema", vbExclamation, gstrNombreEmpresa
                adoFon.Close: Set adoFon = Nothing
                Exit Sub
            End If
        End If
        adoFon.Close: Set adoFon = Nothing
        
        '***  Adicionar Opciones ***
'                Call VerOpciones(strPerfil)
    
        '*** Conexión a la Fecha de Procesamiento del Fondo ***
        If strPasar = "NO" Then
            Set adoCuo = New ADODB.Recordset
            
            .CommandText = "SELECT FechaCuota,ValorTipoCambio FROM FondoValorCuota WHERE IndAbierto='X'"
            Set adoCuo = .Execute
            
            If adoCuo.EOF Then
                MsgBox "Fecha de Procesamiento no registrada", vbCritical, gstrNombreEmpresa
                adoCuo.Close: Set adoCuo = Nothing
                Exit Sub
            Else
                gdatFechaActual = adoCuo("FechaCuota"): gdblTipoCambio = CDbl(adoCuo("ValorTipoCambio"))
            End If
            adoCuo.Close: Set adoCuo = Nothing
                                            
            gstrDiaActual = Format(Day(gdatFechaActual), "00")
            gstrMesActual = Format(Month(gdatFechaActual), "00")
            gstrPeriodoActual = Format(Year(gdatFechaActual), "0000")
        Else
            Set adoCuo = New ADODB.Recordset
            
            .CommandText = "{ call up_ACSelDatos(0) }"
            Set adoCuo = .Execute
            
            If Not adoCuo.EOF Then
                gstrFechaActual = Convertyyyymmdd(adoCuo("FechaServidor"))
                gdatFechaActual = CVDate(adoCuo("FechaServidor"))
            End If
            adoCuo.Close: Set adoCuo = Nothing
        End If
        
        frmMainMdi.stbMdi.Panels(3).Text = "Seteo de parámetros..."
        '*** Set de Parámetros Globales ***
        Set adoresult = New ADODB.Recordset
        
        .CommandText = "SELECT CodParametro,ValorParametro FROM ParametroGeneral"
        Set adoresult = .Execute
        
        If adoresult.EOF Then
            MsgBox "Verifique la Definición de Parámetros Globales del Sistema.", vbOKOnly + vbCritical, "Acceso"
            adoresult.Close: Set adoresult = Nothing
            Exit Sub
        End If
        
        Do While Not adoresult.EOF
            Select Case adoresult("CodParametro")
                Case "01" 'Tasa IGV
                    gdblTasaIgv = CDbl(adoresult("ValorParametro")) / 100
                Case "02" 'Inicial Título Manual
                    gstrInicialTitulo = adoresult("ValorParametro")
                Case "03" 'Dias de Pago de Rescates T+n
                    gintDiasPagoRescate = CInt(adoresult("ValorParametro"))
                Case "04" 'Dias de Pago de Com.Administración Cartera
                    gstrFchPagAdm = Trim(adoresult("ValorParametro"))
                Case "07"
                    gstrClaseTipoCambioFondo = Trim(adoresult("ValorParametro"))
                Case "08"
                    gstrValorTipoCambioCierre = Trim(adoresult("ValorParametro"))
            End Select
            adoresult.MoveNext
        Loop
        adoresult.Close
                    
        adoComm.CommandText = "SELECT CodSucursal,CodAgencia FROM InstitucionPersona " & _
        "WHERE TipoPersona='01' AND CodPersona='" & gstrCodPromotor & "'"
        Set adoresult = adoComm.Execute
    
        If Not adoresult.EOF Then
            gstrCodSucursal = Trim(adoresult("CodSucursal"))
            gstrCodAgencia = Trim(adoresult("CodAgencia"))
        End If
        adoresult.Close: Set adoresult = Nothing
                        
        '*** Verificación Caducidad Contraseña ***
    '                If (CVDate(convertddmmyyyy(gstrFechaAct)) - CVDate(convertddmmyyyy(adoUsr("FCH_MODI")))) > gintDiasDuracionClave Then
    '                    '*** Cambiar Contraseña ***
    '                    Set adoConn = Nothing: MousePointer = vbDefault
    '
    '                    gstrLogin = Trim(txtUsuario.Text): gstrPassword = Trim(txtPassword.Text)
    '                    adoConn.Open "DSN=" & gstrODBCName32 & ";UID=" & gstrLogin & ";PWD=" & gstrPassword & ";"
    '                    Set adoComm.ActiveConnection = adoConn
    '                    frmSYSchgpwd.Show vbModal
    '                    frmSYSchgpwd.lbl_UsrNam.Caption = Trim(txtUsuario.Text)
    '                    Exit Sub
    '                Else
    '                    If (CVDate(convertddmmyyyy(gstrFechaAct)) - CVDate(convertddmmyyyy(adoUsr("FCH_MODI")))) > (gintDiasDuracionClave - gintDiasVctoClave) Then
    '                        If MsgBox("Faltan " & CStr(gintDiasDuracionClave - (CVDate(convertddmmyyyy(gstrFechaAct)) - CVDate(convertddmmyyyy(adoUsr("FCH_MODI"))))) & " días para vencer su Contraseña." & Chr(10) & Chr(13) & "Desea Cambiarla ahora ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    '                            Set adoConn = Nothing: MousePointer = vbDefault
    '
    '                            gstrLogin = Trim(txtUsuario.Text): gstrPassword = Trim(txtPassword.Text)
    '                            adoConn.Open "DSN=" & gstrODBCName32 & ";UID=" & gstrLogin & ";PWD=" & gstrPassword & ";"
    '                            Set adoComm.ActiveConnection = adoConn
    '                            frmSYSchgpwd.Show vbModal
    '                            frmSYSchgpwd.lbl_UsrNam.Caption = Trim(txtUsuario.Text)
    '                            Exit Sub
    '                        End If
    '                    End If
    '                End If
    '                Me.Refresh
        
        '*** Carga Array con Dias No Utiles ***
        Call LDiasNUtil
        
        frmMainMdi.stbMdi.Panels(3).Text = "Verificando..."
        
        adoUsr.Close: Set adoUsr = Nothing
        
        frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        frmMainMdi.txtUsuarioSistema.Text = Trim(strDscUser)
        gstrNombreEmpresa = Trim(cboAdministradora.Text)
        
        With frmMainMdi.stbMdi
            .Panels(1).Text = Trim(cboTipoAdministradora.Text) & " - " & gstrNombreEmpresa & Space(1)
            .Panels(2).Text = gstrNombreSistema & Space(1)
            .Panels(3).Text = "Acción"
        End With
                
        Unload frmAcceso
                                    
    End With
                              
End Sub
Private Sub cmdCancelar_Click()

    Set adoConn = Nothing
    End
            
End Sub

Private Sub cmdSalir_Click()

Dim Valor As Integer
    'Set frmAcceso = Nothing
    'Valor = MsgBox("¿Está seguro de salir del sistema SPECTRUM?", vbYesNo + vbQuestion, "SPECTRUM")
    'If Valor = 6 Then
    Ind = 0
    Unload Me
'        If Trim(frmMainMdi.txtUsuarioSistema.Text) = Valor_Caracter Then
'            Set frmAcceso = Nothing ' HMC
'            Set frmMainMdi = Nothing
'            End
'        End If
    'Else
        'txtUsuario.SetFocus
    'End If
End Sub

Private Sub Form_Load()
    
    Dim strRutaLogo As String

    strRutaLogo = gstrImagePath & "LogoCliente.jpg" 'App.Path & "\Logo\Logo.jpg"
    
    If Len(dir(strRutaLogo, vbArchive)) <> 0 Then
        imgLogo.Picture = LoadPicture(strRutaLogo)
    End If
    
    txtUsuario.Text = Valor_Caracter: txtPassword.Text = Valor_Caracter
'    txtUsuario.SetFocus
    Call HabilitarControles(False)
                                  
End Sub
Private Sub Form_Resize()

   ' imgTAM.Left = 0: imgTAM.Top = 0
   ' imgTAM.Width = frmAcceso.Width: imgTAM.Height = frmAcceso.Height
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim Valor As Integer
    'Set frmAcceso = Nothing                     ' HMC
    
    If Ind <> 1 Then
       ' If strPasar = "" Or strPasar = "NO" Then ' HMC
                Valor = MsgBox("¿Está seguro de salir del sistema SPECTRUM?", vbQuestion + vbYesNo, "SPECTRUM")
            If Valor = 6 Then
                'If Trim(frmMainMdi.txtUsuarioSistema.Text) = Valor_Caracter Then
        '            Set frmAcceso = Nothing ' HMC
                    Set frmMainMdi = Nothing
                    End
                'End If
        '    Else                ' HMC
        '        Cancel = 1      ' HMC
            Else
                Cancel = 1
            End If
    End If
        
    'Else                        ' HMC
    '    If strPasar = "" Or strPasar = "NO" Then    'HMC
            
'            Valor = MsgBox("Esta seguro de salir del sistema SPECTRUM", vbYesNo, "SPECTRUM") 'HMC
     '       If Valor = 6 Then   ' HMC
      '          If Trim(frmMainMdi.txtUsuarioSistema.Text) = Valor_Caracter Then           'HMC
      '              Set frmAcceso = Nothing         'HMC
      '              Set frmMainMdi = Nothing        'HMC
      '              End         ' HMC
      '          End If          ' HMC
      '      Else                ' HMC
      '          Cancel = 1      ' HMC
      '      End If              ' HMC
                
      '  End If                  ' HMC
   ' End If                      ' HMC
    
    Ind = 0
    Exit Sub
End Sub



Private Sub txtPassword_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(LCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
       cmdAccion_Click (0)
    End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(LCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then   'HMC 05:37 p.m. 29/08/2008
    txtPassword.SetFocus    'HMC 05:37 p.m. 29/08/2008
    End If                  'HMC 05:37 p.m. 29/08/2008
    
End Sub



