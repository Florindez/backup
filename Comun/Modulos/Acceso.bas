Attribute VB_Name = "Acceso"
Option Explicit

Public gstrConnect                                      As String
Public gstrConnectSeguridad                             As String
Public gstrServerSeguridad                              As String
Public gstrDataBaseSeguridad                            As String
Public gstrLoginSeguridad                               As String
Public gstrServer                                       As String
Public gstrPasswSeguridad                               As String
Public gstrDataBase                                     As String
Public gstrLogin                                        As String
Public gstrLoginUS                                      As String
Public gstrTipoAdministradora                           As String
Public gstrTipoAdministradoraContable                   As String
Public gstrCodAdministradora                            As String
Public gstrCodAdministradoraContable                    As String
Public gstrCodFondoContable                             As String
Public gstrPassword                                     As String
Public gstrFechaActual                                  As String
Public gdatFechaActual                                  As Date
Public gstrDiaActual                                    As String
Public gstrMesActual                                    As String
Public gstrPeriodoActual                                As String
Public gdblTipoCambio                                   As Double
Public gstrGrupo                                        As String
Public gstrConnectConsulta                              As String
Public gstrNombreEmpresa                                As String
Public gstrNombreSistema                                As String
Public gstrVersionSistema                               As String
Public gstrCodMoneda                                    As String
Public gdblTasaIgv                                      As Double
Public gdblTasaRetencion                                As Double
Public gdblTasaDetraccion                               As Double
Public gintDiasInversionRV                              As Integer
Public gintDiasInversionRF                              As Integer
Public gcurMontoMaximoRetencion                         As Currency
Public gstrClaseTipoCambioFondo                         As String
Public gstrCodClaseTipoCambioFondo                      As String
Public gstrValorTipoCambioCierre                        As String
Public gstrClaseTipoCambioOperacionFondo                As String
Public gstrCodClaseTipoCambioOperacionFondo             As String
Public gstrValorTipoCambioOperacion                     As String
Public gstrClaseTipoCambioLiquidacionRC                 As String
Public gstrCodClaseTipoCambioLiquidacionRC              As String
Public gstrValorTipoCambioLiquidacionRC                 As String
Public gstrInicialTitulo                                As String
Public gstrCuentaIgv                                    As String
Public gstrCuentaImptoRenta                             As String
Public gstrValoracionBonosLocal                         As String
Public gstrValoracionBonosExterior                      As String
Public gintDiasDuracionClave                            As Integer
Public gintDiasVctoClave                                As Integer
Public gstrTratamientoContableComisionValorLocal        As String
Public gstrTratamientoContableComisionValorExtranjero   As String
Public gstrTratamientoContableIGVValorLocal             As String
Public gstrTratamientoContableIGVValorExtranjero        As String


Public gstrRptPath                          As String
Public gstrImagePath                        As String
Public gstrTempPath                         As String
Public gstrBackupPath                       As String
Public gstrCargaOperacionesPath             As String


'Public gstrBancoDefecto                     As String
Public gintDiasPagoRescate                  As Integer
Public gstrFchPagAdm                        As String
Public gstrpHorSusc                         As String

Public gstrFormatoFechaCliente              As String
Public gstrFormatoFechaInterno              As String


Public Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Const HKEY_CURRENT_USER = &H80000001
Public Const KEY_QUERY_VALUE = &H1
Public Const REG_SZ As Long = 1
Public Const REG_DWORD As Long = 4
Public Const ERROR_NONE = 0
Global Const REG_OPTION_NON_VOLATILE = 0
Global Const KEY_ALL_ACCESS = &H3F

Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
        "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
        As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
        As Long, phkResult As Long, lpdwDisposition As Long) As Long

Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
        "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
        String, ByVal cbData As Long) As Long

Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
        "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
        ByVal cbData As Long) As Long


Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Public Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Function CargarParametrosGlobales(Optional strCodFondo As String = "000") As Boolean

    Dim adoRegistro     As ADODB.Recordset
    
    CargarParametrosGlobales = False
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm

        .CommandText = "{call up_ACCargarParametrosGenerales ('" & _
                    gstrCodFondoContable & "','" & gstrCodAdministradora & "')}"
        Set adoRegistro = .Execute
            
        If adoRegistro.EOF Then
            MsgBox "Verifique la Definición de Parámetros Globales del Sistema.", vbOKOnly + vbCritical, "Acceso"
            adoRegistro.Close: Set adoRegistro = Nothing
            Exit Function
        Else
            gdblTasaIgv = adoRegistro("TasaIgv")
            '*** Inicial Título Valor Manual ***
            gstrInicialTitulo = adoRegistro("InicialTitulo")
            '*** Dias de Pago de Inversiones RF ***
            gintDiasInversionRF = adoRegistro("DiasInversionRF")
            '*** Dias de Pago de Inversiones RV ***
            gintDiasInversionRV = adoRegistro("DiasInversionRV")
            gstrValoracionBonosLocal = adoRegistro("ValoracionBonosLocal")
            gstrValoracionBonosExterior = adoRegistro("ValoracionBonosExterior")
            gstrClaseTipoCambioFondo = adoRegistro("ClaseTipoCambioFondo")
            gstrCodClaseTipoCambioFondo = adoRegistro("CodClaseTipoCambioFondo")
            gstrValorTipoCambioCierre = adoRegistro("ValorTipoCambioCierre")
            gdblTasaDetraccion = adoRegistro("TasaDetraccion")
            gdblTasaRetencion = adoRegistro("TasaRetencion")
            gcurMontoMaximoRetencion = adoRegistro("MontoMaximoRetencion")
            '*** Días Duración Contraseña ***
            gintDiasDuracionClave = adoRegistro("DiasDuracionClave")
            '*** Días Vcto. Contraseña ***
            gintDiasVctoClave = adoRegistro("DiasVctoClave")
            gstrClaseTipoCambioOperacionFondo = Trim(adoRegistro("ClaseTipoCambioOperacionFondo"))
            gstrCodClaseTipoCambioOperacionFondo = Trim(adoRegistro("CodClaseTipoCambioOperacionFondo"))
            gstrValorTipoCambioOperacion = Trim(adoRegistro("ValorTipoCambioOperacion"))
            gstrClaseTipoCambioLiquidacionRC = Trim(adoRegistro("ClaseTipoCambioLiquidacionRC"))
            gstrCodClaseTipoCambioLiquidacionRC = Trim(adoRegistro("CodClaseTipoCambioLiquidacionRC"))
            gstrValorTipoCambioLiquidacionRC = Trim(adoRegistro("ValorTipoCambioLiquidacionRC"))
            
            gstrTratamientoContableComisionValorLocal = Trim(adoRegistro("TratamientoContableComisionValorLocal"))
            gstrTratamientoContableComisionValorExtranjero = Trim(adoRegistro("TratamientoContableComisionValorExtranjero"))
            gstrTratamientoContableIGVValorLocal = Trim(adoRegistro("TratamientoContableIGVValorLocal"))
            gstrTratamientoContableIGVValorExtranjero = Trim(adoRegistro("TratamientoContableIGVValorExtranjero"))
       
        
        End If
            
        adoRegistro.Close: Set adoRegistro = Nothing
    
    End With
            
    CargarParametrosGlobales = True
   
    
End Function

Public Function QueryValueEx(ByVal lngValorClave As Long, ByVal strNombreValor As String, vntValorClave As Variant) As Long

    Dim lngLongitud As Long, lngRetorno As Long
    Dim lngTipo As Long, lngValor As Long
    Dim strValor As String

    On Error GoTo QueryValueExError

    '***  Determinar la longitud y el tipo de dato a leer ***
    lngRetorno = RegQueryValueExNULL(lngValorClave, strNombreValor, 0&, lngTipo, 0&, lngLongitud)
    If lngRetorno <> ERROR_NONE Then Error 5

    Select Case lngTipo
        '*** Para strings ***
        Case REG_SZ:
            strValor = String(lngLongitud, 0)

            lngRetorno = RegQueryValueExString(lngValorClave, strNombreValor, 0&, lngTipo, strValor, lngLongitud)
            If lngRetorno = ERROR_NONE Then
                vntValorClave = Left$(strValor, lngLongitud - 1)
            Else
                vntValorClave = Empty
            End If
            
        '*** Para DWORDS ***
        Case REG_DWORD:
            lngRetorno = RegQueryValueExLong(lngValorClave, strNombreValor, 0&, lngTipo, lngValor, lngLongitud)
            If lngRetorno = ERROR_NONE Then vntValorClave = lngValor
            
        Case Else
            'all other data types not supported
            lngRetorno = -1
            
    End Select

QueryValueExExit:
    QueryValueEx = lngRetorno
    Exit Function

QueryValueExError:
    Resume QueryValueExExit
    
End Function
 Private Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
   Dim hNewKey As Long 'handle a la nueva clave
   Dim lRetVal As Long 'resultado de la funcion RegCreateKeyEx

   lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
             vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
             0&, hNewKey, lRetVal)
   RegCloseKey (hNewKey)
 End Sub
  Private Sub SetValue(sKeyName As String, sValueName As String, _
             vValueSetting As Variant, lValueType As Long)
   Dim lRetVal As Long 'resultado de la funcion SetValueEx
   Dim hKey As Long 'handle de la clave abierta

  'abrir la clave especificada
   lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, _
                          KEY_ALL_ACCESS, hKey)
   lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
   RegCloseKey (hKey)
End Sub
Public Function SetValueEx(ByVal hKey As Long, sValueName As String, _
       lType As Long, vValue As Variant) As Long

  Dim lValue As Long
  Dim sValue As String

  Select Case lType
     Case REG_SZ
          sValue = vValue & Chr$(0)
          SetValueEx = RegSetValueExString(hKey, sValueName, 0&, _
                                           lType, sValue, Len(sValue))
     Case REG_DWORD
          lValue = vValue
          SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, _
                                         lType, lValue, 4)
  End Select
End Function

Public Sub GenerarParametrosGlobales()

    
    CreateNewKey Clave_Registro_Sistema, HKEY_CURRENT_USER
    
    SetValue Clave_Registro_Sistema, Clave_Registro_TipoAdministradora, "02", REG_SZ
    SetValue Clave_Registro_Sistema, Clave_Registro_Administradora, "002", REG_SZ
    SetValue Clave_Registro_Sistema, Clave_Registro_TipoAdministradoraContable, "02", REG_SZ
    SetValue Clave_Registro_Sistema, Clave_Registro_NombreSistema, "Spectrum Fondos", REG_SZ
    SetValue Clave_Registro_Sistema, Clave_Registro_Empresa, "Cartisa SAF", REG_SZ
    SetValue Clave_Registro_Sistema, Clave_Registro_Servidor, "TAM_LT01\NCF", REG_SZ
    SetValue Clave_Registro_Sistema, Clave_Registro_BaseDatos, "fondos", REG_SZ
    SetValue Clave_Registro_Sistema, Clave_Registro_FormatoFechaCliente, "dd/mm/yyyy", REG_SZ
    SetValue Clave_Registro_Sistema, Clave_Registro_FormatoFechaInterno, "yyyymmdd", REG_SZ
    SetValue Clave_Registro_Sistema, Clave_Registro_Version, "1.0", REG_SZ
    SetValue Clave_Registro_Sistema, Clave_Registro_Cliente, "Cartisa SAF", REG_SZ
    
    #If RELEASEMODE = 1 Then     'en modo release
        SetValue Clave_Registro_Sistema, Clave_Registro_RutaReportes, App.Path & "\Reportes\", REG_SZ
        SetValue Clave_Registro_Sistema, Clave_Registro_RutaLogoTAM, App.Path & "\Imagenes\", REG_SZ
    #ElseIf TESTINGMODE = 1 Then 'en modo testing
        SetValue Clave_Registro_Sistema, Clave_Registro_RutaReportes, "D:\Data\NCF SAFI\Spectrum\Fuentes\Reportes\", REG_SZ
        SetValue Clave_Registro_Sistema, Clave_Registro_RutaLogoTAM, "D:\Data\NCF SAFI\Spectrum\Fuentes\Imagenes\", REG_SZ
    #End If
    
    
End Sub

Public Sub ObtenerParametrosGlobales()

    Dim lngValorRetorno As Long         'result of the API functions
    Dim lngCodigoLlave  As Long         'handle of opened key
    Dim vntValor        As Variant      'setting of queried value
       
    lngValorRetorno = RegOpenKeyEx(HKEY_CURRENT_USER, Clave_Registro_Sistema, 0, KEY_QUERY_VALUE, lngCodigoLlave)
    
    If lngValorRetorno <> ERROR_NONE Then
        Call GenerarParametrosGlobales
        lngValorRetorno = RegOpenKeyEx(HKEY_CURRENT_USER, Clave_Registro_Sistema, 0, KEY_QUERY_VALUE, lngCodigoLlave)
    End If
    
    '*** Código de Tipo de Administradora ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, Clave_Registro_TipoAdministradora, vntValor)
    gstrTipoAdministradora = vntValor
    '*** Código de Tipo de Administradora Contable ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, Clave_Registro_TipoAdministradoraContable, vntValor)
    gstrTipoAdministradoraContable = vntValor
    '*** Código de Administradora ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, Clave_Registro_Administradora, vntValor)
    gstrCodAdministradora = vntValor
    '*** Nombre del Sistema ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, Clave_Registro_NombreSistema, vntValor)
    gstrNombreSistema = vntValor
    '*** Versión del Sistema ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, Clave_Registro_Version, vntValor)
    gstrVersionSistema = vntValor
    '*** Nombre del Servidor ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, Clave_Registro_Servidor, vntValor)
    gstrServer = vntValor
    '*** Nombre de la Base de Datos ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, Clave_Registro_BaseDatos, vntValor)
    gstrDataBase = vntValor
    '*** Ruta de los Reportes ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, Clave_Registro_RutaReportes, vntValor)
    gstrRptPath = vntValor
    '*** Formato de Fecha Cliente ***
'    lngValorRetorno = QueryValueEx(lngCodigoLlave, Clave_Registro_RutaLogoTAM, vntValor)
'    gstrLogoPath = vntValor
    
    lngValorRetorno = QueryValueEx(lngCodigoLlave, Clave_Registro_RutaLogoTAM, vntValor)
    gstrImagePath = vntValor
    
    lngValorRetorno = QueryValueEx(lngCodigoLlave, Clave_Registro_RutaTemporal, vntValor)
    gstrTempPath = vntValor
    
    lngValorRetorno = QueryValueEx(lngCodigoLlave, Clave_Registro_RutaBackups, vntValor)
    gstrBackupPath = vntValor
    
    lngValorRetorno = QueryValueEx(lngCodigoLlave, Clave_Registro_RutaCargaOperaciones, vntValor)
    gstrCargaOperacionesPath = vntValor
    
    '*** Formato de Fecha Cliente ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, Clave_Registro_FormatoFechaCliente, vntValor)
    gstrFormatoFechaCliente = vntValor
    '*** Formato de Fecha Interno ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, Clave_Registro_FormatoFechaInterno, vntValor)
    gstrFormatoFechaInterno = vntValor
    '*** Nombre de Cliente ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, Clave_Registro_Cliente, vntValor)
    gstrNombreEmpresa = vntValor
    
    RegCloseKey (lngCodigoLlave)
                
End Sub

Public Sub ObtenerParametrosGlobalesSeguridad()

    Dim lngValorRetorno As Long         'result of the API functions
    Dim lngCodigoLlave  As Long         'handle of opened key
    Dim vntValor        As Variant      'setting of queried value
       
    lngValorRetorno = RegOpenKeyEx(HKEY_CURRENT_USER, Clave_Registro_Sistema_Seguridad, 0, KEY_QUERY_VALUE, lngCodigoLlave)
    
'    If lngValorRetorno <> ERROR_NONE Then
'        Call GenerarParametrosGlobales
'        lngValorRetorno = RegOpenKeyEx(HKEY_CURRENT_USER, Clave_Registro_Sistema, 0, KEY_QUERY_VALUE, lngCodigoLlave)
'    End If
    
   
    '*** Nombre del Servidor ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, "Servidor", vntValor)
    gstrServerSeguridad = vntValor
    '*** Nombre de la Base de Datos ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, "Base de Datos", vntValor)
    gstrDataBaseSeguridad = vntValor
    '*** Nombre de Usuario ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, "Usuario Seguridad", vntValor)
    gstrLoginSeguridad = vntValor
    '*** Password ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, "Password U. Seguridad", vntValor)
    gstrPasswSeguridad = vntValor
    
    
    RegCloseKey (lngCodigoLlave)
                
End Sub

Sub VerPrinc(strPerfil As String)
    
    Dim adoresult As ADODB.Recordset

    With adoComm
        Set adoresult = New ADODB.Recordset
        
        .CommandText = "SELECT FLG_ACT,NUM_IND FROM SYOPCION WHERE COD_MODU='" & strPerfil & "' AND RIGHT(NOM_OPC,1)='1' ORDER BY COD_OPC"
        Set adoresult = .Execute
        
        If Not adoresult.EOF Then
            Do While Not adoresult.EOF
'                If Trim(adoresult("FLG_ACT")) = "X" Then
'                    frmMainMdi.mnuopc1(CInt(adoresult("NUM_IND"))).Enabled = True
'                Else
'                    frmMainMdi.mnuopc1(CInt(adoresult("NUM_IND"))).Enabled = False
'                End If
'                If Trim$(adoresult!FLG_VIS) = "X" Then
'                    frmSYSMainMdi.mnuopc1(CInt(adoresult("NUM_IND"))).Visible = True
'                Else
'                    frmSYSMainMdi.mnuopc1(CInt(adoresult("NUM_IND"))).Visible = False
'                End If
                adoresult.MoveNext
            Loop
        End If
        adoresult.Close: Set adoresult = Nothing
    End With
    
End Sub
Sub VerPrincusu(strPerfil As String)
    
    Dim adoresult As ADODB.Recordset

    With adoComm
        Set adoresult = New ADODB.Recordset
        
        .CommandText = "SELECT FLG_ACT,NUM_IND FROM SYOPCION WHERE COD_MODU='" & strPerfil & "' AND RIGHT(NOM_OPC,1)='1' ORDER BY COD_OPC"
        Set adoresult = .Execute
        
        If Not adoresult.EOF Then
            Do While Not adoresult.EOF
'                If Trim(adoresult("FLG_ACT")) = "X" Then
'                    frmMainMdi.mnuopc1(CInt(adoresult("NUM_IND"))).Enabled = True
'                Else
'                    frmMainMdi.mnuopc1(CInt(adoresult("NUM_IND"))).Enabled = False
'                End If
'
'                If Trim(adoresult!flg_vis) = "X" Then
'                    frmMainMdi.mnuopc1(CInt(adoresult("NUM_IND"))).Visible = True
'                Else
'                    frmMainMdi.mnuopc1(CInt(adoresult("NUM_IND"))).Visible = False
'                End If
                adoresult.MoveNext
            Loop
        End If
        adoresult.Close: Set adoresult = Nothing
    End With
    
End Sub


Public Sub ControlOpciones(strpCodPerfil As String)

    Dim adoOpciones     As ADODB.Recordset
    Dim ctrlOpciones    As Control
    Dim intNumIndice    As Integer

    On Error Resume Next
   
    Set adoOpciones = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT NombreMenu,IndActivo,IndiceMenu,IndVisible " & _
            "FROM MenuSistema WHERE CodModulo='" & frmMainMdi.Tag & "' ORDER BY CodMenu"
        Set adoOpciones = .Execute
        
        If Not adoOpciones.EOF Then
            Do While Not adoOpciones.EOF
                intNumIndice = CInt(adoOpciones("IndiceMenu"))
            
                For Each ctrlOpciones In frmMainMdi.Controls
                    If Trim(adoOpciones("NombreMenu")) = ctrlOpciones.Name Then
                        If intNumIndice = ctrlOpciones.Index Then
                            If adoOpciones("IndActivo") = Valor_Indicador Then
                                ctrlOpciones.Enabled = True
                            Else
                                ctrlOpciones.Enabled = False
                            End If
                            
                            If adoOpciones("IndVisible") = Valor_Indicador Then
                                ctrlOpciones.Visible = True
                            Else
                                ctrlOpciones.Visible = False
                            End If
                            
                            Exit For
                        End If
                    End If
                Next
                adoOpciones.MoveNext
            Loop
        End If
        adoOpciones.Close: Set adoOpciones = Nothing
    End With
    
End Sub

Sub Main()

    '*** Cargar formulario principal ***
'    frmMainMdi.Show
    Call ObtenerParametrosGlobales
    Call ObtenerParametrosGlobalesSeguridad
    
    frmSplash.Show
    Sleep 0&
    'frmAcceso.Show
    'Sleep 0&
    
End Sub


