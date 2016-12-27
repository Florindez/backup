VERSION 5.00
Begin VB.Form frmSeleccionaAdministradora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entidad Administradora"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4800
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
      Left            =   2550
      Picture         =   "frmSeleccionaAdministradora.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3150
      Width           =   1275
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Continuar"
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
      Left            =   1005
      Picture         =   "frmSeleccionaAdministradora.frx":0582
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3150
      Width           =   1215
   End
   Begin VB.ComboBox cboFondo 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2580
      Width           =   4245
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione un Fondo y pulse Continuar"
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
      Index           =   33
      Left            =   660
      TabIndex        =   1
      Top             =   2220
      Width           =   3345
   End
   Begin VB.Image imgLogo 
      BorderStyle     =   1  'Fixed Single
      Height          =   1860
      Left            =   270
      Stretch         =   -1  'True
      Top             =   180
      Width           =   4215
   End
End
Attribute VB_Name = "frmSeleccionaAdministradora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSQL As String, strCodFondo As String
Dim arrAdministradora() As String

Private Sub cboFondo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAccion_Click (0)
    End If
End Sub

Private Sub cmdSalir_Click()
    Dim Valor As Integer
    'Set frmAcceso = Nothing
    Valor = MsgBox("¿Está seguro de salir del sistema SPECTRUM?", vbYesNo + vbQuestion, "SPECTRUM")
    If Valor = 6 Then
        Set frmMainMdi = Nothing
        End
    End If
    
End Sub

Private Sub Form_Load()

    Call CargarListas
    
    CentrarForm Me
    
End Sub

Private Sub CargarListas()

    strSQL = "SELECT CodFondo CODIGO,DescripFondo DESCRIP FROM Fondo WHERE Estado='01'"
    CargarControlLista strSQL, cboFondo, arrAdministradora, Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
End Sub

Private Sub cboFondo_Click()
    
    Dim strRutaLogo As String
    Dim adoFondoFrecuencia As ADODB.Recordset
    
    strCodFondo = Valor_Caracter

    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrAdministradora(cboFondo.ListIndex))
    
    strRutaLogo = gstrImagePath & strCodFondo & ".jpg" 'App.Path & "\Logo\Logo.jpg"
    
    If Len(dir(strRutaLogo, vbArchive)) <> 0 Then
        imgLogo.Picture = LoadPicture(strRutaLogo)
    End If
    
    Set adoFondoFrecuencia = New ADODB.Recordset
    
        With adoComm
        
            .CommandText = "SELECT FrecuenciaValorizacion FROM Fondo WHERE CodFondo='" & Trim(arrAdministradora(cboFondo.ListIndex)) & "' " & _
                                " AND CodAdministradora='" & gstrCodAdministradora & "'"
            
            Set adoFondoFrecuencia = .Execute
        End With
        
        If App.Title = "General" Then
            If Trim(adoFondoFrecuencia("FrecuenciaValorizacion")) = "05" Then
                
                frmMainMdi.mnuProcesos(1).Caption = "Cierre Mensual"
            Else
                frmMainMdi.mnuProcesos(1).Caption = "Cierre Diario"
            End If
        End If
    
End Sub

Private Sub cmdAccion_Click(Index As Integer)
        
    Dim adoAdministradora As ADODB.Recordset
    Dim adoFondo As ADODB.Recordset
        
'    With adoComm
'
'        .CommandText = "SELECT DescripAdministradora FROM Administradora WHERE Estado='01' AND CodAdministradora='" & Trim(strCodFondo) & "'"
'        Set adoAdministradora = .Execute
'
'    End With
'
'    If Not adoAdministradora.EOF Then
'
'        gstrNombreAdministradora = Trim(adoAdministradora("DescripAdministradora").Value)
'
'    End If
'
'    gstrCodAdministradora = strCodFondo
    

    '*** Verificación de los fondos existentes ***
    
    gstrCodFondoContable = strCodFondo
    
    With adoComm
    
        Set adoFondo = New ADODB.Recordset

        .CommandText = "SELECT CodFondo,DescripFondo FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "'"
        Set adoFondo = .Execute
    
        If adoFondo.EOF Then
            If App.Title = "General" Then
                MsgBox "Bienvenido a " & gstrNombreSistema & Space(1) & gstrVersionSistema, vbExclamation, gstrNombreEmpresa
            Else
                MsgBox "No existen Fondos definidos en el Sistema.", vbExclamation, gstrNombreEmpresa
                adoFondo.Close: Set adoFondo = Nothing
                GoTo exit_form
            End If
        End If
        
        '*******************************************
        
    End With
    
    
    If Not ContinuarAcceso Then Exit Sub
    
    
    
'    If Not CargarParametrosGlobales() Then GoTo exit_form
    
    frmMainMdi.txtEmpresa = cboFondo.Text
    
    Unload Me
    
    Exit Sub
    
exit_form:

    Exit Sub
       
End Sub


Private Function ContinuarAcceso() As Boolean

    Dim adoFondo            As ADODB.Recordset, adoUsuario      As ADODB.Recordset
    Dim adoCuotas           As ADODB.Recordset, adoRegistroTmp  As ADODB.Recordset
    Dim adoRegistro         As ADODB.Recordset, adoRegistroAux  As ADODB.Recordset
    Dim strPerfil           As String, Msg                      As String
    Dim res                 As Integer, intDias                 As Integer
    Dim strDescripUsuario   As String
    Dim vntFchUser          As Variant
            
    ContinuarAcceso = False
    
    strPasar = "NO"
    
    
    With adoComm
'        Set adoUsuario = New ADODB.Recordset
'
'        .CommandText = "SELECT DescripUsuario FROM UsuarioSistema WHERE IdUsuario='" & gstrLogin & "'"
'        Set adoUsuario = .Execute
'
'        If adoUsuario.EOF Then
'            adoUsuario.Close: Set adoUsuario = Nothing
'            Exit Function
'        End If
'
'        strDescripUsuario = Trim(adoUsuario("DescripUsuario"))
''        gstrCodPromotor = Trim(adoUsuario("CodPromotor"))
'
'        frmMainMdi.stbMdi.Panels(3).Text = "Verificando accesos..."
'        '*** Verificación de Accesos de Usuario ***
'        strPerfil = frmMainMdi.Tag
        
        '---/// Integración de seguridad
'        If Not ValidarPermisoAccesoObjeto(Trim(gstrLogin), Trim(App.Title), Codigo_Tipo_Objeto_Modulo) Then
'            Exit Function
'        End If
'
        '*** Verificacion de la Version del Módulo ***
'        frmMainMdi.stbMdi.Panels(3).Text = "Verificando versión..."
'        Call VerifVersion
        
        '*** Es el Módulo de Administración ? ***
'        If App.Title = "Administracion" Then
'            MsgBox "Bienvenido a " & gstrNombreSistema & Space(1) & gstrVersionSistema, vbExclamation, gstrNombreEmpresa
'            strPasar = "SI"
'        End If
'
'        '*** Conexión a la Fecha de Procesamiento del Fondo ***
           If strPasar = "NO" Then
'            Set adoCuotas = New ADODB.Recordset
'            .CommandText = "SELECT ValorParametro FROM ParametroGeneral WHERE CodParametro = '22'"
'            Set adoCuotas = .Execute
'            If Not adoCuotas.EOF Then
'
'                gstrCodFondoContable = Trim(adoCuotas("ValorParametro"))
'                adoCuotas.Close
'            End If
            
            
            
            
               .CommandText = "SELECT FechaCuota,ValorTipoCambio FROM FondoValorCuota WHERE IndAbierto='X' " & IIf(gstrCodFondoContable = Valor_Caracter, "", " AND CodFondo = '" & gstrCodFondoContable & "'")
            Set adoCuotas = .Execute
            
            'If Not adoCuotas.EOF Then
            
             If adoCuotas.EOF Then
                
                 adoCuotas.Close
    
                
                .CommandText = "SELECT MAX(FechaFinal) FechaFinal FROM PeriodoContable WHERE IndCierre='X' and MesContable = '99'"
                Set adoCuotas = .Execute
            
                If Not adoCuotas.EOF Then
                    gdatFechaActual = adoCuotas("FechaFinal"): gdblTipoCambio = 1
                    gstrFechaActual = Convertyyyymmdd(adoCuotas("FechaFinal"))
                Else
                    MsgBox "Fecha de Procesamiento no registrada", vbCritical, gstrNombreEmpresa
                    adoCuotas.Close: Set adoCuotas = Nothing
                    Exit Function
                End If
                
            Else
                gdatFechaActual = adoCuotas("FechaCuota"): gdblTipoCambio = CDbl(adoCuotas("ValorTipoCambio"))
                gstrFechaActual = Convertyyyymmdd(adoCuotas("FechaCuota"))
            End If
            adoCuotas.Close: Set adoCuotas = Nothing
            
            gstrDiaActual = Format(Day(gdatFechaActual), "00")
            gstrMesActual = Format(Month(gdatFechaActual), "00")
            gstrPeriodoActual = Format(Year(gdatFechaActual), "0000")
        Else
            Set adoCuotas = New ADODB.Recordset
            
            .CommandText = "{ call up_ACSelDatos(0) }"
            Set adoCuotas = .Execute
            
            If Not adoCuotas.EOF Then
                gstrFechaActual = Convertyyyymmdd(adoCuotas("FechaServidor"))
                gdatFechaActual = CVDate(adoCuotas("FechaServidor"))
            End If
            adoCuotas.Close
            
            gstrDiaActual = Format(Day(gdatFechaActual), "00")
            gstrMesActual = Format(Month(gdatFechaActual), "00")
            gstrPeriodoActual = Format(Year(gdatFechaActual), "0000")
            
            If App.Title = "Administracion" Then
                .CommandText = "SELECT CodAdministradora,DescripAdministradora FROM Administradora " & _
                    "WHERE CodTipoAdministradora='" & Codigo_Tipo_Fondo_Administradora & "' AND IndDefecto='X'"
                Set adoCuotas = .Execute
                
                If Not adoCuotas.EOF Then
                    frmMainMdi.txtEmpresa.Text = Trim(adoCuotas("DescripAdministradora"))
                    gstrCodAdministradoraContable = adoCuotas("CodAdministradora")
                Else
                    frmMainMdi.txtEmpresa.Text = "Entidad No Definida"
                    gstrCodAdministradoraContable = Valor_Caracter
                End If
                adoCuotas.Close
            
                .CommandText = "SELECT FechaContable,ValorTipoCambio FROM AdministradoraCalendario WHERE IndAbierto='X'"
                Set adoCuotas = .Execute
                
                If adoCuotas.EOF Then
                    gstrFechaActual = Valor_Caracter
                Else
                    gdatFechaActual = adoCuotas("FechaContable"): gdblTipoCambio = CDbl(adoCuotas("ValorTipoCambio"))
                    gstrFechaActual = Convertyyyymmdd(adoCuotas("FechaContable"))
                    
                    gstrDiaActual = Format(Day(gdatFechaActual), "00")
                    gstrMesActual = Format(Month(gdatFechaActual), "00")
                    gstrPeriodoActual = Format(Year(gdatFechaActual), "0000")
                End If
                adoCuotas.Close
            End If
            Set adoCuotas = Nothing
        End If
        
        frmMainMdi.stbMdi.Panels(3).Text = "Seteo de parámetros..."
        
        '*** SE COMENTO PARA AFP INTEGRA ***
        If Not CargarParametrosGlobales() Then Exit Function
        
        Set adoRegistro = New ADODB.Recordset
                    
        adoComm.CommandText = "SELECT CodSucursal,CodAgencia FROM InstitucionPersona " & _
        "WHERE TipoPersona='01' AND CodPersona='" & gstrCodPromotor & "'"
        Set adoRegistro = adoComm.Execute
    
        If Not adoRegistro.EOF Then
            gstrCodSucursal = Trim(adoRegistro("CodSucursal"))
            gstrCodAgencia = Trim(adoRegistro("CodAgencia"))
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
        
        '*** Carga Array con Dias No Utiles ***
        Call LDiasNUtil 'ACR 17/09/09: Este metodo debe cambiarse por uno mas optimo
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
'        adoUsuario.Close: Set adoUsuario = Nothing
'        adoRegistroAux.Close: Set adoRegistroAux = Nothing
                
        frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
'        frmMainMdi.txtUsuarioSistema.Text = Trim(strDescripUsuario)
'
'        With frmMainMdi.stbMdi
'            .Panels(1).Text = gstrNombreEmpresa & Space(1)
'            .Panels(2).Text = gstrNombreSistema & Space(1) & gstrVersionSistema & Space(1)
'            .Panels(3).Text = "Acción"
'        End With
'
'        Unload frmAcceso
                                    
    End With
    
    ContinuarAcceso = True
    
End Function
