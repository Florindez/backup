Attribute VB_Name = "Comun"
Option Explicit

Public gs_FormName As String


Public adoConn                          As ADODB.Connection
Public adoConnSeguridad                 As ADODB.Connection
Public adoCommAux                       As ADODB.Command
Public adoComm                          As ADODB.Command
Public adoCommSeguridad                 As ADODB.Command
Public gstrRptConnectODBC               As String

Public gstrCodPromotor                  As String '*** variable para seleccionar operador por defecto = usuario ***
Public gstrCodSucursal                  As String
Public gstrCodAgencia                   As String
Public gstrCodParticipe                 As String
Public gstrCodParticipeTransferente     As String
Public gstrCodParticipeTransferido      As String
Public gstrFormulario                   As String
Public garrTipoDocumento()              As String
Public garrTipoDocumentoTransferente()  As String
Public garrTipoDocumentoTransferido()   As String
Public garrFondo()                      As String
Public garrParticipe()                  As String
Public gboolMostrarSelectAdministradora As Boolean
Public gstrNombreAdministradora         As String


Public gobjReport                       As Object
Public gstrCodCuenta                    As String
Public gstrNumAsiento                   As String
Public gstrCodMonedaReporte             As String
Public gindMonedaContable               As String
Public gstrFchDel                       As String
Public gstrFchAl                        As String
Public gstrSelFrml                      As String        '*** Selection formula de reportes ***
Public gstraFormRepo()                  As String
Public gstrstoparRep()                  As String
Public gstrNameRepo                     As String
Public gstrTipoCart                     As String
Public gblnRollBack                     As Boolean
Public gblnErr                          As Boolean, gstrMenErr As String
Public gstrfchlim                       As String
Public gstrglbfile                      As String
Public gstrglbanal                      As String
Public gstrglbclsb                      As String
Public gstrglbfond                      As String
Public gintglbflg                       As Integer
Public gdblRestCIIUFond                 As Double
Public gdblRestGrupFond                 As Double
Public gdblRestEmprFond                 As Double
Public gdblRestEmprAcci                 As Double
Public gdblRestEmprObli                 As Double

'***** BMM CAMBIOS ******
Public gstrCodVistaProceso As String
Public gstrCodVistaUsuario As String
'************************

Public gstrTextoAdministradorFormula            As String
Public gstrConnectNET                           As String
Public gstrConnectNETSeguridad                  As String


Public strCtaInversion                  As String
'ACR
Public strCtaInversionCostoSAB              As String
Public strCtaInversionCostoBVL              As String
Public strCtaInversionCostoCavali           As String
Public strCtaInversionCostoFondoGarantia    As String
Public strCtaInversionCostoConasev          As String
Public strCtaInversionCostoIGV              As String
Public strCtaInversionCostoCompromiso       As String
Public strCtaInversionCostoResponsabilidad  As String
Public strCtaInversionCostoFondoLiquidacion As String
Public strCtaInversionCostoComisionEspecial As String
Public strCtaInversionCostoGastosBancarios  As String
'ACR
Public strCtaProvInteres                As String
Public strCtaInteres                    As String
Public strCtaInteresCastigado           As String
Public strCtaCosto                      As String
Public strCtaIngresoOperacional         As String
Public strCtaInteresVencido             As String
Public strCtaVacCorrido                 As String
Public strCtaXPagar                     As String
Public strCtaXCobrar                    As String
Public strCtaInteresCorrido             As String
Public strCtaProvReajusteK              As String
Public strCtaReajusteK                  As String
Public strCtaProvFlucMercado            As String
Public strCtaFlucMercado                As String
'ACR 07/05/2012
Public strCtaProvFlucMercadoPerdida     As String
Public strCtaFlucMercadoPerdida         As String
'JJCC 20/05/2012
Public strCtaXPagarEmitida              As String
Public strCtaGananciaRedondeo           As String
Public strCtaPerdidaRedondeo            As String
'ACR 07/05/2012
Public strCtaProvInteresVac             As String
Public strCtaInteresVac                 As String
Public strCtaIntCorridoK                As String
Public strCtaProvFlucK                  As String
Public strCtaFlucK                      As String
Public strCtaInversionTransito          As String
Public strCtaProvGasto                  As String
Public strCtaImpuesto                   As String
Public strCtaCostoBVL                   As String
Public strCtaCostoSAB                   As String
Public strCtaCostoCavali                As String
Public strCtaCostoConasev               As String
Public strCtaCostoFondoGarantia         As String
Public strCtaCostoFondoLiquidacion      As String
Public strCtaGastoBancario              As String
Public strCtaComisionEspecial           As String
Public strCtaCompromiso                 As String
Public strCtaResponsabilidad            As String
Public strCtaME                         As String
Public strCtaMN                         As String
Public strCtaComision                   As String
Public strCtaDetraccion                 As String
Public strCtaRetencion                  As String


Public curCtaInversion                  As Currency
'ACR
Public curCtaInversionCostoSAB              As Currency
Public curCtaInversionCostoBVL              As Currency
Public curCtaInversionCostoCavali           As Currency
Public curCtaInversionCostoFondoGarantia    As Currency
Public curCtaInversionCostoConasev          As Currency
Public curCtaInversionCostoIGV              As Currency
Public curCtaInversionCostoCompromiso       As Currency
Public curCtaInversionCostoResponsabilidad  As Currency
Public curCtaInversionCostoFondoLiquidacion As Currency
Public curCtaInversionCostoComisionEspecial As Currency
Public curCtaInversionCostoGastosBancarios  As Currency
'ACR
Public curCtaProvInteres                As Currency
Public curCtaInteres                    As Currency
Public curCtaInteresCastigado           As Currency
Public curCtaCosto                      As Currency
Public curCtaIngresoOperacional         As Currency
Public curCtaInteresVencido             As Currency
Public curCtaVacCorrido                 As Currency
Public curCtaXPagar                     As Currency
Public curCtaXCobrar                    As Currency
Public curCtaInteresCorrido             As Currency
Public curCtaProvReajusteK              As Currency
Public curCtaReajusteK                  As Currency
Public curCtaProvFlucMercado            As Currency
Public curCtaFlucMercado                As Currency

Public curCtaProvFlucMercadoPerdida     As Currency
Public curCtaFlucMercadoPerdida         As Currency

Public curCtaProvInteresVac             As Currency
Public curCtaInteresVac                 As Currency
Public curCtaIntCorridoK                As Currency
Public curCtaProvFlucK                  As Currency
Public curCtaFlucK                      As Currency
Public curCtaInversionTransito          As Currency
Public curCtaProvGasto                  As Currency
Public curCtaImpuesto                   As Currency
Public curCtaImpuestoCredito            As Currency
Public curCtaCostoBVL                   As Currency
Public curCtaCostoSAB                   As Currency
Public curCtaCostoCavali                As Currency
Public curCtaCostoConasev               As Currency
Public curCtaCostoFondoLiquidacion      As Currency
Public curCtaCostoFondoGarantia         As Currency
Public curCtaGastoBancario              As Currency
Public curCtaComisionEspecial           As Currency
Public curCtaCompromiso                 As Currency
Public curCtaResponsabilidad            As Currency
Public curCtaME                         As Currency
Public curCtaMN                         As Currency
Public curDifReaCap                     As Currency

'*** Para Manejo de Días No Utiles
Global gvntDiasNUtil()                  As Variant

'Global gstrTipCns                       As String         '** Tipo de consulta Est.Cuentas/Operac.Partícipes
'Global gstrCodFon                       As String
'Global gstrFldCnsCrt                    As String      '** campo para consulta de certificados
'Global gstrNomSoli                      As String
'Global gstrCodUnicBco                   As String
'Global gstrFchCierreAux                 As String
'Global gstrflgvcon                      As String
'Global gstrtipval                       As String
'Global gstrCodUnico                     As String
'Global gstrVarMant                      As String
'Global gstrNomOpc                       As String
'Global gstrNumInd                       As String
'Global gstrVarMantPartic                As String
'***************************** Calculate_accrued_interest *********************************
Public Const GERMAN = 1, SPEC_GERMAN = 2, ENGLISH = 3, FRENCH = 4, US = 5, ISMA_YEAR = 6, ISMA_99N = 7, ISMA_99U = 8 ' day count methods
Public Const ERROR_BAD_DCM = -1, ERROR_BAD_DATES = -2 ' error returns
'*************************************************************************************************
Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

'**************************************************** Variables  Api************************************************************************'
'API used for timings
Declare Function GetTickCount Lib "Kernel32" () As Long
'APT used to get value out of INI file
Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'API used for Cut, Copy , Paste and Undo
Declare Function GetKeyState Lib "User" (ByVal nVirtKey As Integer) As Integer

Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'*** Tipo de busqueda de tipo de cambio ***
Public Const Tipo_Busqueda_Tipo_Cambio_Directo = "1"
Public Const Tipo_Busqueda_Tipo_Cambio_Inverso = "2"
Public Const Tipo_Busqueda_Tipo_Cambio_Iterativo_Directo = "3"
Public Const Tipo_Busqueda_Tipo_Cambio_Iterativo_Inverso = "4"
Public Const Tipo_Busqueda_Tipo_Cambio_Multiple = "5"

'*** Constantes de Error ***
Public Const Codigo_Error_RegistroDuplicado = -2147217873
Public Const Codigo_Error_ArchivoNoExiste = -2147206461

'*** Constantes de Control de Errores  ***   HMC 04:22 p.m. 19/09/2008
Public Const TituloError = "Ocurrió un Problema - Spectrum Fondos"  'HMC
Public Const DescripcionError = "Por Favor pongase en Contacto con el Provedor de Software."
Public Const DescripcionTecnica = "Descripción Técnica : "          'HMC

Public Const LR_LOADFROMFILE = &H10
Public Const IMAGE_BITMAP = 0
Public Const IMAGE_ICON = 1
Public Const IMAGE_CURSOR = 2
Public Const IMAGE_ENHMETAFILE = 3
Public Const CF_BITMAP = 2

Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long

Public Enum HH_COMMAND 'HMC
HH_DISPLAY_TOPIC = &H0
HH_HELP_FINDER = &H0
HH_DISPLAY_TOC = &H1
HH_DISPLAY_INDEX = &H2
HH_DISPLAY_SEARCH = &H3
HH_SET_WIN_TYPE = &H4
HH_GET_WIN_TYPE = &H5
HH_GET_WIN_HANDLE = &H6
HH_GET_INFO_TYPES = &H7
HH_SET_INFO_TYPES = &H8
HH_SYNC = &H9
HH_ADD_NAV_UI = &HA
HH_ADD_BUTTON = &HB
HH_GETBROWSER_APP = &HC
HH_KEYWORD_LOOKUP = &HD
HH_DISPLAY_TEXT_POPUP = &HE
HH_HELP_CONTEXT = &HF
HH_TP_HELP_CONTEXTMENU
HH_TP_HELP_WM_HELP = &H11
HH_CLOSE_ALL = &H12
HH_ALINK_LOOKUP = &H13
End Enum

Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As HH_COMMAND, ByVal dwData As Long) As Long

Public Function FormateaCampo(ByRef strValorCampo As String, intTipoDato As Integer) As String

    FormateaCampo = ""
    
    If intTipoDato = adChar Or intTipoDato = adVarChar Or intTipoDato = adDate Then 'String o Date
        strValorCampo = "''" & strValorCampo & "''"
    End If
    
    FormateaCampo = strValorCampo

End Function

Public Function Dias360(ByVal datFechaInicial As Date, ByVal datFechaFinal As Date, ByVal blnMetodo As Boolean) As Integer

    Dim intDias360      As Integer
    Dim lngDiasAnual    As Long, lngDiasMensual     As Long
    Dim lngDiasDiario   As Long
    
    intDias360 = 0
    
    datFechaFinal = DateAdd("d", datFechaFinal, 1)
    lngDiasAnual = (CLng(Year(datFechaFinal)) - CLng(Year(datFechaInicial))) * 360
    lngDiasMensual = (CLng(Month(datFechaFinal)) - CLng(Month(datFechaInicial))) * 30
    lngDiasDiario = (CLng(Day(datFechaFinal)) - CLng(Day(datFechaInicial)))
    
    If blnMetodo = True Then
        intDias360 = CInt(lngDiasAnual + lngDiasMensual + lngDiasDiario)
    Else
        intDias360 = CInt(lngDiasAnual + lngDiasMensual + lngDiasDiario)
    End If
    
    Dias360 = intDias360
                
End Function

Public Function ObtenerParametroDesplazamientoFechaTipoCambio()
 
    Dim adoRegistro As ADODB.Recordset
    Dim intDiasDesplazamiento As Integer
   
    ObtenerParametroDesplazamientoFechaTipoCambio = 0
   
'    'Obteniendo el factor de desplazamiento para el calculo de la fecha de T/C
'    Set adoRegistro = New ADODB.Recordset
'
'    adoComm.CommandText = "SELECT CONVERT(int,ValorParametro) AS DiasDesplaza " & _
'                          "FROM ParametroGeneral WHERE CodParametro = '19'"
'    Set adoRegistro = adoComm.Execute
'
'    If Not (adoRegistro.EOF) Then
'        intDiasDesplazamiento = adoRegistro("DiasDesplaza")
'    End If
'
'    If intDiasDesplazamiento = Null Then
'        intDiasDesplazamiento = 0
'    End If
'
'    adoRegistro.Close
'
'    Set adoRegistro = Nothing
'
'    ObtenerParametroDesplazamientoFechaTipoCambio = intDiasDesplazamiento
 
 
End Function

Public Function FindRecordset(ByRef adoRs As ADODB.Recordset, strCriteria As String) As Boolean

    Dim clone_rs As ADODB.Recordset
    Set clone_rs = adoRs.Clone

    FindRecordset = False
    
    clone_rs.Filter = strCriteria

    If clone_rs.EOF Or clone_rs.BOF Then
        clone_rs.Close
        Set clone_rs = Nothing
        Exit Function
    Else
        adoRs.Bookmark = clone_rs.Bookmark
    End If

    clone_rs.Close
    Set clone_rs = Nothing
    
    FindRecordset = True
      
End Function

Public Function ExisteDinamica(ByVal strpCodFile As String, ByVal strpCodDetalleFile As String, ByVal strpCodAdministradora As String, ByVal strpCodDinamica As String, ByVal strpCodMoneda As String) As Boolean
    Dim adoRegistro     As ADODB.Recordset
    Dim adoAuxiliar     As ADODB.Recordset
    
    ExisteDinamica = True
    
    '*** Verificar Si Existe Dinamica Contable ***
    With adoComm
        Set adoRegistro = New ADODB.Recordset
        
        .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
            "WHERE TipoOperacion='" & strpCodDinamica & "' AND CodFile='" & strpCodFile & "' AND (CodDetalleFile='" & _
            strpCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & strpCodAdministradora & _
            "' AND CodMoneda='" & IIf(strpCodMoneda = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"
        Set adoRegistro = .Execute
    
        If Not adoRegistro.EOF Then
            If CInt(adoRegistro("NumRegistros")) <= 0 Then
                Set adoAuxiliar = New ADODB.Recordset
                
                .CommandText = "SELECT DescripParametro FROM AuxiliarParametro " & _
                    "WHERE CodTipoParametro='OPECAJ' AND CodParametro='" & strpCodDinamica & "'"
                Set adoAuxiliar = .Execute
                
                If Not adoAuxiliar.EOF Then
                    MsgBox "NO EXISTE Dinámica Contable: " & Trim(adoAuxiliar("DescripParametro")), vbCritical
                End If
                adoAuxiliar.Close: Set adoAuxiliar = Nothing
                
                adoRegistro.Close: Set adoRegistro = Nothing
                ExisteDinamica = False: Exit Function
            End If
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    ExisteDinamica = True
    
End Function

Public Sub GenerarLibroMayor(strpCodFondo As String, strpCodMonedaReporte As String, datpFechaInicial As Date, datpFechaFinal As Date)

    Dim adoRegistro         As ADODB.Recordset
    Dim strFechaInicial     As String, strFechaFinal        As String
    Dim strFechaInicialS    As String, strFechaFinalS       As String
    Dim strFechaAnterior    As String
    
    strFechaInicial = Convertyyyymmdd(datpFechaInicial)
    strFechaFinal = Convertyyyymmdd(datpFechaFinal)
    strFechaInicialS = Convertyyyymmdd(DateAdd("d", 1, datpFechaInicial))
    strFechaFinalS = Convertyyyymmdd(DateAdd("d", 1, datpFechaFinal))
    strFechaAnterior = Convertyyyymmdd(DateAdd("d", -1, datpFechaInicial))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        adoConn.CommandTimeout = 10000
'        .CommandText = "SELECT CodCuenta FROM PartidaContableMayor " & _
'            "WHERE (FechaInicial>='" & strFechaInicial & "' AND FechaInicial<'" & strFechaInicialS & "') AND " & _
'            "(FechaFinal>='" & strFechaFinal & "' AND FechaFinal<'" & strFechaFinalS & "') AND " & _
'            "CodFondo='" & strpCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'        Set adoRegistro = .Execute
'
'        If Not adoRegistro.EOF Then
'            If MsgBox("La información para el rango solicitado YA EXISTE !" & vbNewLine & vbNewLine & _
'                "Desea procesarla nuevamente ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                
'                .CommandText = "DELETE PartidaContableMayor " & _
'                    "WHERE (FechaInicial>='" & strFechaInicial & "' AND FechaInicial<'" & strFechaInicialS & "') AND " & _
'                    "(FechaFinal>='" & strFechaFinal & "' AND FechaFinal<'" & strFechaFinalS & "') AND " & _
'                    "CodFondo='" & strpCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'                .CommandText = "DELETE PartidaContableMayorSaldos " & _
'                    "WHERE CodFondo='" & strpCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'                adoConn.Execute .CommandText
                
'                .CommandText = "DELETE PartidaContableAuxiliar " & _
'                    "WHERE (FechaMovimiento>='" & strFechaInicial & "' AND FechaMovimiento<'" & strFechaFinalS & "') AND " & _
'                    "CodFondo='" & strpCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                .CommandText = "DELETE PartidaContableAuxiliar " & _
                    "WHERE CodFondo='" & strpCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                adoConn.Execute .CommandText
                
                '*** Generar Libro Mayor ***
                .CommandText = "{ call up_CNProcLibroMayor2('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strFechaInicial & "','" & strFechaFinal & "','" & strpCodMonedaReporte & "','%') }"
                adoConn.Execute .CommandText
                
'            End If
'        Else
'            .CommandText = "DELETE PartidaContableMayor " & _
'                    "WHERE (FechaInicial>='" & strFechaInicial & "' AND FechaInicial<='" & strFechaInicialS & "') AND " & _
'                    "(FechaFinal>='" & strFechaFinal & "' AND FechaFinal<='" & strFechaFinalS & "') AND " & _
'                    "CodFondo='" & strpCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'            adoConn.Execute .CommandText
'
'            .CommandText = "DELETE PartidaContableAuxiliar " & _
'                "WHERE (FechaMovimiento>='" & strFechaInicial & "' AND FechaMovimiento<='" & strFechaFinalS & "') AND " & _
'                "CodFondo='" & strpCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'            adoConn.Execute .CommandText
'
'            '*** Generar Libro Mayor ***
'            .CommandText = "{ call up_CNProcLibroMayor('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & _
'                strFechaInicial & "','" & strFechaFinal & "','" & strFechaFinalS & "','" & strFechaAnterior & "','" & _
'                strFechaInicial & "') }"
'            adoConn.Execute .CommandText
'        End If
'        adoRegistro.Close: Set adoRegistro = Nothing
        
    End With

End Sub
Public Sub InicializarParametrosFondo()

'modifica ACR: 27/01/2009

    ReDim astrParametrosFondo(PARAM_INICIA To PARAM_FINAL)

    '*** Valores Parámetros Fondo ***
    astrParametrosFondo(PARAM_NUMOPE) = Valor_NumOperacion
    astrParametrosFondo(PARAM_NUMOCR) = Valor_NumOpeCertificado
    astrParametrosFondo(PARAM_NUMSOL) = Valor_NumSolicitud
    astrParametrosFondo(PARAM_NUMCOM) = Valor_NumComprobante
    astrParametrosFondo(PARAM_NUMCER) = Valor_NumCertificado
    astrParametrosFondo(PARAM_NUMCAJ) = Valor_NumOrdenCaja
    astrParametrosFondo(PARAM_NUMINT) = Valor_NumInt
    astrParametrosFondo(PARAM_NUMORD) = Valor_NumOrdenInversion
    astrParametrosFondo(PARAM_NUMKAR) = Valor_NumKardex
    astrParametrosFondo(PARAM_NUMENT) = Valor_NumEntregaEvento
    astrParametrosFondo(PARAM_NUMCOB) = Valor_NumCobertura
    astrParametrosFondo(PARAM_NUMOPC) = Valor_NumOpeCajaBancos
    astrParametrosFondo(PARAM_NUMREC) = Valor_NumRegistroCompra

End Sub
Public Sub GenerarParametrosFondo(ByVal strpCodAdministradora, ByVal strpCodFondo As String)

   Dim sensql()  As String, intContador As Integer
   
   Call InicializarParametrosFondo
   
   ReDim sensql(PARAM_FINAL)
   
   
   For intContador = PARAM_INICIA To PARAM_FINAL
       sensql(intContador) = "INSERT INTO ParametroFondo (CodFondo,CodAdministradora,CodParametro,UltNumParametro) VALUES ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & astrParametrosFondo(intContador) & "',0)"
   Next intContador
   
   
'   SenSql(1) = "INSERT INTO ParametroFondo (CodFondo,CodAdministradora,CodParametro,UltNumParametro) VALUES ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Valor_NumOperacion & "',0)"
'   SenSql(2) = "INSERT INTO ParametroFondo (CodFondo,CodAdministradora,CodParametro,UltNumParametro) VALUES ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Valor_NumOpeCertificado & "',0)"
'   SenSql(3) = "INSERT INTO ParametroFondo (CodFondo,CodAdministradora,CodParametro,UltNumParametro) VALUES ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Valor_NumSolicitud & "',0)"
'   SenSql(4) = "INSERT INTO ParametroFondo (CodFondo,CodAdministradora,CodParametro,UltNumParametro) VALUES ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Valor_NumComprobante & "',0)"
'   SenSql(5) = "INSERT INTO ParametroFondo (CodFondo,CodAdministradora,CodParametro,UltNumParametro) VALUES ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Valor_NumCertificado & "',0)"
'   SenSql(6) = "INSERT INTO ParametroFondo (CodFondo,CodAdministradora,CodParametro,UltNumParametro) VALUES ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Valor_NumOrdenCaja & "',0)"
'   SenSql(7) = "INSERT INTO ParametroFondo (CodFondo,CodAdministradora,CodParametro,UltNumParametro) VALUES ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Valor_NumInt & "',0)"
'   SenSql(8) = "INSERT INTO ParametroFondo (CodFondo,CodAdministradora,CodParametro,UltNumParametro) VALUES ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Valor_NumOrdenInversion & "',0)"
'   SenSql(9) = "INSERT INTO ParametroFondo (CodFondo,CodAdministradora,CodParametro,UltNumParametro) VALUES ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Valor_NumKardex & "',0)"
'   SenSql(10) = "INSERT INTO ParametroFondo (CodFondo,CodAdministradora,CodParametro,UltNumParametro) VALUES ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Valor_NumEntregaEvento & "',0)"
'   SenSql(11) = "INSERT INTO ParametroFondo (CodFondo,CodAdministradora,CodParametro,UltNumParametro) VALUES ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Valor_NumCobertura & "',0)"
  
   
   For intContador = PARAM_INICIA To PARAM_FINAL
      adoComm.CommandText = sensql(intContador)
      adoConn.Execute adoComm.CommandText
   Next
   
     
End Sub

Public Function CalculoVacCorrido(ByVal strpCodTitulo As String, ByVal dblpCantidad As Double, ByVal datpFechaEmision As Date, ByVal datpFechaLiquidacion As Date, ByVal strpCodIndiceFinal As String, ByVal strpTipoAjuste As String, ByVal strpTipoTasa As String, ByVal strpPeriodoPago As String, ByVal strpCodIndiceInicial As String, ByVal intpBase As Integer) As Double

    '*** Inicio de Cálculo Automático de Intereses Corridos ***
    CalculoVacCorrido = 0
    '*** Cálculo de los Intereses Corridos a la Fecha de la Liquidación                       ***
    '*** Si NumCupon='001' AND FechaLiquidacion = FechaInicio Cupón Vigente Días Corridos = 0 ***
    '*** SiNo  Días Corridos = (FechaLiquidacion) - (FechaInicio Cupón Vigente) + 1           ***
    '*** Los Intereses Corridos en Bonos se calculan sobre el Valor Nominal                   ***
    '*** y el las Letras Hipotecarias sobre el Saldo por Amortizar.                           ***
    '*** Para el caso de Bonos VAC se calcula sobre el Valor Nominal Ajustado                 ***
        
    Dim adoRegistroBono     As ADODB.Recordset, adoRegistroTmp  As ADODB.Recordset
    Dim strFechaLiquidacion As String
    Dim intDiaTranscurridos As Integer, intDiasCUP              As Integer
    Dim intDiasCorridos     As Integer, intRes                  As Integer
    Dim intDiasPeriodo      As Integer
    Dim dblIntCorr          As Double, dblCapitalRea            As Double
    Dim dblTasDia           As Double, dblVacCorrido            As Double
    Dim strFechaInicioCupon As String, strFechaFinCupon         As String
    Dim datFechaInicioCupon As Date, datFechaFinCupon           As Date
        
    strFechaLiquidacion = Convertyyyymmdd(datpFechaLiquidacion)
    
    With adoComm
        Set adoRegistroBono = New ADODB.Recordset
        .CommandType = adCmdText
        
        '*** Obtener datos del cupón vigente ***
        .CommandText = "SELECT FactorDiario,FechaInicio,FechaVencimiento,NumCupon,CantDiasPeriodo,FactorDiario1,FechaInicioIndice,FechaFinIndice " & _
            "FROM InstrumentoInversionCalendario WHERE CodTitulo='" & strpCodTitulo & "' AND FechaVencimiento>='" & strFechaLiquidacion & "' ORDER BY FechaVencimiento"
        Set adoRegistroBono = .Execute
        
        If Not adoRegistroBono.EOF Then
            '*** Fecha de inicio del cupón ***
            strFechaInicioCupon = Convertyyyymmdd(adoRegistroBono("FechaInicio"))
            datFechaInicioCupon = adoRegistroBono("FechaInicio")
            datFechaFinCupon = adoRegistroBono("FechaVencimiento")
            intDiasPeriodo = adoRegistroBono("CantDiasPeriodo")
                   
            '*** Días corridos entre el inicio del cupón y la fecha de liquidación ***
            If (strFechaInicioCupon = strFechaLiquidacion) And (adoRegistroBono("NumCupon") = "001") Then
                intDiasCorridos = 0
                intDiaTranscurridos = 0
            Else
                If adoRegistroBono("NumCupon") = "001" Then
                   intDiasCorridos = DateDiff("d", datFechaInicioCupon, datpFechaLiquidacion)
                   intDiaTranscurridos = DateDiff("d", datFechaInicioCupon, datpFechaLiquidacion)
                Else
                   intDiasCorridos = DateDiff("d", datFechaInicioCupon, datpFechaLiquidacion) + 1
                   intDiaTranscurridos = DateDiff("d", datFechaInicioCupon, datpFechaLiquidacion) + 1
                End If
            End If
        
            '*** Obtención de parametros para Bonos VAC ***
            If strpTipoAjuste = Codigo_Tipo_Ajuste_Vac Then
'                If strpCodIndiceFinal = Codigo_Vac_Liquidacion Then  '*** Bonos VAC Periodicos: Cálculo a partir del cupón anterior ***
'
'                    Set adoRegistroTmp = New ADODB.Recordset
'
'                    If CInt(adoRegistroBono("NumCupon")) = 1 Then
'                        '*** Primer cupón construir las fechas y días del cupón ***
'                        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodParametro='" & strpPeriodoPago & "' AND CodTipoParametro='TIPFRE'"
'                        Set adoRegistroTmp = .Execute
'
'                        If Not adoRegistroTmp.EOF Then
'                            intDiasPeriodo = CInt(adoRegistroTmp("ValorParametro"))
'                        End If
'                        adoRegistroTmp.Close: Set adoRegistroTmp = Nothing
'
'                        datFechaFinCupon = DateAdd("d", -1, Convertddmmyyyy(datpFechaEmision))
'                        datFechaInicioCupon = DateAdd("m", Int(intDiasPeriodo / 30) * -1, datFechaFinCupon)
'                        intDiasCUP = DateDiff("d", datFechaInicioCupon, datFechaFinCupon) + 1
'                    Else
'                        '*** Cualquier otro cupón: extraer los datos del cupón anterior ***
'                        .CommandText = "SELECT FechaInicio,FechaVencimiento,CantDiasPeriodo " & _
'                            "FROM InstrumentoInversionCalendario WHERE CodTitulo='" & strpCodTitulo & "' AND NumCupon='" & Format(CInt(adoRegistroBono("NumCupon")) - 1, "000") & "'"
'                        Set adoRegistroTmp = .Execute
'
'                        If Not adoRegistroTmp.EOF Then
'                            datFechaInicioCupon = DateAdd("d", -1, adoRegistroTmp("FechaInicio"))
'                            datFechaFinCupon = adoRegistroTmp("FechaVencimiento")
'                            intDiasCUP = CInt(adoRegistroTmp("CantDiasPeriodo"))
'                        End If
'                        adoRegistroTmp.Close: Set adoRegistroTmp = Nothing
'                    End If
'                Else
'                    strFechaFinCupon = Convertyyyymmdd(adoRegistroBono("FechaVencimiento"))
'                End If
'
'                strFechaInicioCupon = Convertyyyymmdd(datFechaInicioCupon)
'                strFechaFinCupon = Convertyyyymmdd(datFechaFinCupon)
                strFechaInicioCupon = Convertyyyymmdd(adoRegistroBono("FechaInicioIndice"))
                strFechaFinCupon = Convertyyyymmdd(adoRegistroBono("FechaFinIndice"))
            End If
        End If
        adoRegistroBono.Close
                        
        .CommandText = "SELECT CodFile,NumCupon,FechaInicio,FactorDiario,ValorAmortizacion,SaldoAmortizacion,TasaInteres,FactorInteres1,CantDiasPeriodo,FechaVencimiento,FactorDiario1 " & _
            "FROM InstrumentoInversionCalendario WHERE CodTitulo='" & strpCodTitulo & "' AND FechaVencimiento>='" & strFechaLiquidacion & "' ORDER BY FechaVencimiento"
        Set adoRegistroBono = .Execute
        
        If Not adoRegistroBono.EOF Then
            If adoRegistroBono("CodFile") = "005" Then  '*** Bonos ***
                Dim curIntCapRea    As Currency, curDifReaCap   As Currency
                
                '*** Cálculo automático de Intereses Corridos ***
                If strpTipoAjuste = Codigo_Tipo_Ajuste_Vac Then '*** Bonos VAC ***
'                    If strpCodIndiceInicial = Codigo_Vac_Emision Then '*** Factor diario Bonos VAC Periodicos ***
'                        If (adoRegistroBono("FactorInteres1") = 0) Or (adoRegistroBono("FactorInteres1") = Null) Then
'                            MsgBox "Cupón Vigente no tiene factor del periodo sin VAC, VERIFIQUE.", vbCritical, "Aviso"
'                            dblTasDia = 0: dblIntCorr = 0
'                            Exit Function
'                        End If
'                        dblTasDia = ((1 + adoRegistroBono("FactorInteres1")) ^ (1 / adoRegistroBono("CantDiasPeriodo"))) - 1
'                    Else '*** Factor Diario Bonos VAC Al Vcto. ***
'                        dblTasDia = adoRegistroBono("FactorDiario")
'                    End If
                    dblTasDia = adoRegistroBono("FactorInteres1")
                    
                    '*** Cálculo del Capital Nominal Reajustado para todos los Bonos ***
                    dblCapitalRea = CalculaCapitalVAC(strFechaLiquidacion, Convertyyyymmdd(datpFechaEmision), dblpCantidad, strFechaInicioCupon, strFechaFinCupon, strpCodIndiceInicial, strpCodIndiceFinal, intDiasPeriodo, intDiasCorridos, adoRegistroBono("TasaInteres") * 0.01, intpBase)
                                     
                    '*** Cálculo de la diferencia del Capital Reajustado ***
                    If dblCapitalRea > 0 Then
                        curDifReaCap = dblCapitalRea - dblpCantidad
                    End If
                    dblVacCorrido = curDifReaCap
'                    If strpCodIndiceInicial = Codigo_Vac_Emision Then
'                        '*** VAC Corrido Adelantado ***
'                        dblVacCorrido = 0
'                        '*** Interés Corrido del Capital Reajustado para todos los Bonos ***
'                        curIntCapRea = dblCapitalRea * ((1 + dblTasDia) ^ intDiaTranscurridos - 1)
'                        dblIntCorr = dblpCantidad * ((1 + dblTasDia) ^ intDiaTranscurridos - 1)
'                        curIntCapRea = curIntCapRea - dblIntCorr
'                        '*** Vac Corrido ***
''                        dblIntCorr = Round(dblIntCorr + curIntCapRea + curDifReaCap, 2)
'                        dblVacCorrido = curDifReaCap - dblIntCorr
'                    ElseIf strpCodIndiceInicial = Codigo_Vac_Liquidacion Then
'                        '*** VAC Corrido Adelantado ***
'                        If IsNumeric(dblpCantidad) Then
'                           dblVacCorrido = curDifReaCap
'                        Else
'                           dblVacCorrido = 0
'                        End If
'                        '*** Interés Corrido del Capital Reajustado ***
'                        curIntCapRea = dblCapitalRea * ((1 + dblTasDia) ^ intDiaTranscurridos - 1)
'                        dblIntCorr = dblpCantidad * ((1 + dblTasDia) ^ intDiaTranscurridos - 1)
'                        curIntCapRea = curIntCapRea - dblIntCorr
'                        '*** Interés Corrido ***
'                        dblIntCorr = Round(dblIntCorr + curIntCapRea, 2)
'                    End If
                Else '*** Bonos No VAC ***
                    dblTasDia = adoRegistroBono("FactorDiario")
                    '*** VAC Corrido Adelantado e Int. Corrido de Capital Reajustado ***
                    dblVacCorrido = 0: curIntCapRea = 0
                    '*** Interés Corrido ***
                    If strpTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                        dblIntCorr = Format(dblpCantidad * ((1 + dblTasDia) ^ intDiaTranscurridos - 1), "0.00")
                    Else
                        dblIntCorr = Format(dblpCantidad * dblTasDia * intDiaTranscurridos, "0.00")
                    End If
                End If
           
            Else '*** Letras Hipotecarias ***
                Dim curSaldoXAmor   As Currency
                '*** VAC Corrido Adelantado ***
                dblVacCorrido = 0
                '*** Interés Corrido ***
                dblIntCorr = Format((adoRegistroBono("ValorAmortizacion") + adoRegistroBono("SaldoAmortizacion")) * ((1 + adoRegistroBono("FactorDiario")) ^ intDiaTranscurridos - 1), "0.000000")
                curSaldoXAmor = adoRegistroBono("ValorAmortizacion") + adoRegistroBono("SaldoAmortizacion")
                
            End If
        Else
            dblIntCorr = 0: dblVacCorrido = 0
        End If
        adoRegistroBono.Close: Set adoRegistroBono = Nothing
    End With
                    
    CalculoVacCorrido = dblVacCorrido

End Function
Public Function CalculoInteresCorrido(ByVal strpCodTitulo As String, ByVal dblpCantidad As Double, ByVal datpFechaEmision As Date, ByVal datpFechaLiquidacion As Date, ByVal strpCodIndiceFinal As String, ByVal strpTipoAjuste As String, ByVal strpTipoTasa As String, ByVal strpPeriodoPago As String, ByVal strpCodIndiceInicial As String, ByVal strpBaseCalculo As String, ByVal intpBase As Integer, Optional ByVal dblpFactorDiario As Double) As Double

    '*** Inicio de Cálculo Automático de Intereses Corridos ***
    CalculoInteresCorrido = 0
    
    '*** Cálculo de los Intereses Corridos a la Fecha de la Liquidación                       ***
    '*** Si NumCupon='001' AND FechaLiquidacion = FechaInicio Cupón Vigente Días Corridos = 0 ***
    '*** SiNo  Días Corridos = (FechaLiquidacion) - (FechaInicio Cupón Vigente) + 1           ***
    '*** Los Intereses Corridos en Bonos se calculan sobre el Valor Nominal                   ***
    '*** y el las Letras Hipotecarias sobre el Saldo por Amortizar.                           ***
    '*** Para el caso de Bonos VAC se calcula sobre el Valor Nominal Ajustado                 ***
        
    Dim adoRegistroBono     As ADODB.Recordset, adoRegistroTmp  As ADODB.Recordset
    Dim intDiaTranscurridos As Integer, intDiasCUP              As Integer
    Dim intDiasCorridos     As Integer, intRes                  As Integer
    Dim dblIntCorridos      As Double, dblCapitalRea            As Double
    Dim dblTasDia           As Double
    Dim strFechaInicioCupon As String, strFechaFinCupon         As String
    Dim strFechaLiquidacion As String
    Dim datFechaInicioCupon As Date, datFechaFinCupon           As Date
        
    strFechaLiquidacion = Convertyyyymmdd(datpFechaLiquidacion)
    
    With adoComm
        Set adoRegistroBono = New ADODB.Recordset
        .CommandType = adCmdText
        
        '*** Obtener datos del cupón vigente ***
        .CommandText = "SELECT CodFile,FactorDiario,FechaInicio,FechaVencimiento,NumCupon,CantDiasPeriodo,FactorDiario1,FechaInicioIndice,FechaFinIndice " & _
            "FROM InstrumentoInversionCalendario WHERE CodTitulo='" & strpCodTitulo & "' AND FechaVencimiento>='" & strFechaLiquidacion & "' ORDER BY FechaVencimiento"
        Set adoRegistroBono = .Execute
        
        If Not adoRegistroBono.EOF Then
            '*** Fecha de inicio del cupón ***
            strFechaInicioCupon = Convertyyyymmdd(adoRegistroBono("FechaInicio"))
            datFechaInicioCupon = adoRegistroBono("FechaInicio")
            datFechaFinCupon = adoRegistroBono("FechaVencimiento")
                   
            '*** Días corridos entre el inicio del cupón y la fecha de liquidación ***
            If strFechaInicioCupon = strFechaLiquidacion And adoRegistroBono("NumCupon") = "001" And adoRegistroBono("CodFile") = "005" Then
                intDiasCorridos = 0
                intDiaTranscurridos = 0
            Else
                If adoRegistroBono("NumCupon") = "001" And adoRegistroBono("CodFile") = "005" Then
                   intDiasCorridos = DateDiff("d", datpFechaEmision, datpFechaLiquidacion)
                   intDiaTranscurridos = DateDiff("d", datpFechaEmision, datpFechaLiquidacion)
                Else
                   intDiasCorridos = DateDiff("d", datFechaInicioCupon, datpFechaLiquidacion) '+ 1
                   intDiaTranscurridos = DateDiff("d", datFechaInicioCupon, datpFechaLiquidacion) '+ 1
                   intDiasCorridos = Dias360(datFechaInicioCupon, datpFechaLiquidacion, True) '+ 1
                End If
            End If
        
            '*** Obtención de parametros para Bonos VAC ***
            If strpTipoAjuste = Codigo_Tipo_Ajuste_Vac Then
                strFechaInicioCupon = Convertyyyymmdd(adoRegistroBono("FechaInicioIndice"))
                strFechaFinCupon = Convertyyyymmdd(adoRegistroBono("FechaFinIndice"))
            End If
        Else
            '*** Días corridos entre el inicio del cupón y la fecha de liquidación ***
            If datpFechaLiquidacion = gdatFechaActual Then
               intDiasCorridos = DateDiff("d", datpFechaEmision, datpFechaLiquidacion)
               intDiaTranscurridos = DateDiff("d", datpFechaEmision, datpFechaLiquidacion)
            Else
               intDiasCorridos = DateDiff("d", datpFechaEmision, datpFechaLiquidacion) + 1
               intDiaTranscurridos = DateDiff("d", datpFechaEmision, datpFechaLiquidacion) + 1
            End If
        End If
        adoRegistroBono.Close
                        
        .CommandText = "SELECT CodFile,NumCupon,FechaInicio,InstrumentoInversionCalendario.FechaVencimiento,FactorInteres,FactorDiario,ValorAmortizacion,SaldoAmortizacion,TasaInteres,FactorInteres1,CantDiasPeriodo,FactorDiario1 " & _
            "FROM InstrumentoInversionCalendario WHERE CodTitulo='" & strpCodTitulo & "' AND FechaVencimiento>='" & strFechaLiquidacion & "' ORDER BY FechaVencimiento"
        Set adoRegistroBono = .Execute
        
        If Not adoRegistroBono.EOF Then
            If adoRegistroBono("CodFile") = "005" Then  '*** Bonos ***
                '*** Cálculo automático de Intereses Corridos ***
                If strpTipoAjuste = Codigo_Tipo_Ajuste_Vac Then '*** Bonos VAC ***
                    dblTasDia = adoRegistroBono("FactorInteres1")
                    
                    '*** Cálculo del Capital Nominal Reajustado para todos los Bonos ***
                    dblCapitalRea = CalculaCapitalVAC(strFechaLiquidacion, Convertyyyymmdd(datpFechaEmision), dblpCantidad, strFechaInicioCupon, strFechaFinCupon, strpCodIndiceInicial, strpCodIndiceFinal, intDiasCUP, intDiasCorridos, adoRegistroBono("TasaInteres") * 0.01, intpBase)
                 
                    '*** Cálculo de la diferencia del Capital Reajustado ***
                    If dblCapitalRea > 0 Then
                        curDifReaCap = dblCapitalRea - dblpCantidad
                    End If
                 
                    dblIntCorridos = dblCapitalRea * dblTasDia * (intDiaTranscurridos / adoRegistroBono("CantDiasPeriodo"))
                ElseIf strpTipoAjuste = Valor_Caracter Then '*** Bonos No VAC ***
                    dblTasDia = adoRegistroBono("FactorDiario1")
                    
                    '*** Interés Corrido ***
                    If strpTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                        dblIntCorridos = dblpCantidad * ((1 + dblTasDia) ^ intDiaTranscurridos - 1)
                    Else
                        dblIntCorridos = dblpCantidad * dblTasDia * intDiaTranscurridos
                    End If
                Else '*** Bonos con otras tasas de ajuste ***
                    dblTasDia = adoRegistroBono("FactorInteres1")
                    intDiasCUP = CInt(adoRegistroBono("CantDiasPeriodo"))
                    
                    If strpBaseCalculo = Codigo_Base_30_360 Or strpBaseCalculo = Codigo_Base_30_365 Then
                        intDiasCUP = Dias360(adoRegistroBono("FechaInicio"), adoRegistroBono("FechaVencimiento"), True) + 1
                    End If
                    
                    '*** Cálculo del Capital Nominal Reajustado para todos los Bonos ***
                    dblCapitalRea = CalculaCapital(strpTipoAjuste, strFechaLiquidacion, Convertyyyymmdd(datpFechaEmision), dblpCantidad, strFechaInicioCupon, strpCodIndiceInicial, intDiasCUP, intDiasCorridos, adoRegistroBono("TasaInteres") * 0.01, strpBaseCalculo, intpBase)
                 
                    intDiasCUP = CInt(adoRegistroBono("CantDiasPeriodo"))
                    
                    dblIntCorridos = dblCapitalRea * (intDiaTranscurridos / intDiasCUP)
                End If
           
            ElseIf adoRegistroBono("CodFile") = "007" Then  '*** Letras Hipotecarias ***
                Dim curSaldoXAmor   As Currency
                
                '*** Interés Corrido ***
                dblIntCorridos = (adoRegistroBono("ValorAmortizacion") + adoRegistroBono("SaldoAmortizacion")) * ((1 + adoRegistroBono("FactorDiario")) ^ intDiaTranscurridos - 1)
                curSaldoXAmor = adoRegistroBono("ValorAmortizacion") + adoRegistroBono("SaldoAmortizacion")
            
            Else 'para otros casos usa el factor calculado que viene como parametro del presente procedimiento
                        
                dblIntCorridos = 0
                dblTasDia = dblpFactorDiario
                
                '*** Interés Corrido ***
                If strpTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                    dblIntCorridos = dblpCantidad * ((1 + dblTasDia) ^ intDiaTranscurridos - 1)
                Else
                    dblIntCorridos = dblpCantidad * dblTasDia * intDiaTranscurridos
                End If
            
            End If
        Else
            dblIntCorridos = 0
            dblTasDia = dblpFactorDiario
                    
            '*** Interés Corrido ***
            If strpTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                dblIntCorridos = dblpCantidad * ((1 + dblTasDia) ^ intDiaTranscurridos - 1)
            Else
                dblIntCorridos = dblpCantidad * dblTasDia * intDiaTranscurridos
            End If
        End If
        adoRegistroBono.Close: Set adoRegistroBono = Nothing
    End With

                
    CalculoInteresCorrido = dblIntCorridos

End Function
Public Function CalculoInteres(numPorcenTasa As Double, strCodTipoTasa As String, strCodPeriodoTasa As String, strCodBaseCalculo As String, numMontoBaseCalculo As Double, datFechaInicial As Date, datFechaFinal As Date) As Double

        Dim intNumPeriodoAnualTasa As Integer
        Dim intDiasProvision       As Integer
        Dim intDiasBaseAnual       As Integer
        Dim numPorcenTasaAnual     As Double
        Dim numMontoCalculoInteres As Double
        Dim adoConsulta            As ADODB.Recordset
        
        
        With adoComm
            Set adoConsulta = New ADODB.Recordset
    
            '*** Obtener el número de días del periodo de tasa ***
            .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' AND CodParametro='" & strCodPeriodoTasa & "'"
            Set adoConsulta = .Execute
    
            If Not adoConsulta.EOF Then
                intNumPeriodoAnualTasa = CInt(360 / adoConsulta("ValorParametro"))     '*** Numero del periodos por año de la tasa ***
            End If
            adoConsulta.Close: Set adoConsulta = Nothing
        End With
   
        Select Case strCodBaseCalculo
            Case Codigo_Base_30_360:
                intDiasBaseAnual = 360
                intDiasProvision = Dias360(datFechaInicial, datFechaFinal, True)
            Case Codigo_Base_Actual_365:
                intDiasBaseAnual = 365
                intDiasProvision = DateDiff("d", datFechaInicial, datFechaFinal) + 1
            Case Codigo_Base_Actual_360:
                intDiasBaseAnual = 360
                intDiasProvision = DateDiff("d", datFechaInicial, datFechaFinal) + 1
            Case Codigo_Base_30_365:
                intDiasBaseAnual = 365
                intDiasProvision = Dias360(datFechaInicial, datFechaFinal, True)
        End Select

        Select Case strCodTipoTasa
            Case Codigo_Tipo_Tasa_Efectiva:
                numPorcenTasaAnual = (1 + (numPorcenTasa / 100)) ^ (intNumPeriodoAnualTasa) - 1
                numMontoCalculoInteres = Round(numMontoBaseCalculo * ((((1 + numPorcenTasaAnual)) ^ (intDiasProvision / intDiasBaseAnual)) - 1), 2) 'adoRegistro("MontoDevengo") + curMontoRenta
            Case Codigo_Tipo_Tasa_Nominal:
                numPorcenTasaAnual = (numPorcenTasa / 100) * intNumPeriodoAnualTasa
                numMontoCalculoInteres = Round(numMontoBaseCalculo * ((numPorcenTasaAnual * (intDiasProvision / intDiasBaseAnual))), 2)
            Case Codigo_Tipo_Tasa_Flat:
                numPorcenTasaAnual = numPorcenTasa / 100
                numMontoCalculoInteres = Round(numMontoBaseCalculo * (numPorcenTasaAnual), 2)
        End Select

        CalculoInteres = numMontoCalculoInteres


End Function
Public Function CalculaDias(datpFechaInicial As Date, datpFechaFinal As Date, strpCodBaseCalculo As String, Optional strpTipoCalculoDias As String = Tipo_Calculo_Dias_Diferencia)

    Dim intDiasBaseAnual  As Long
    Dim intDiasDiferencia As Long

    Select Case strpCodBaseCalculo
        Case Codigo_Base_30_360:
            intDiasBaseAnual = 360
            intDiasDiferencia = Dias360(datpFechaInicial, datpFechaFinal, True)
        Case Codigo_Base_Actual_365:
            intDiasBaseAnual = 365
            intDiasDiferencia = DateDiff("d", datpFechaInicial, datpFechaFinal) + 1
        Case Codigo_Base_Actual_360:
            intDiasBaseAnual = 360
            intDiasDiferencia = DateDiff("d", datpFechaInicial, datpFechaFinal) + 1
        Case Codigo_Base_30_365:
            intDiasBaseAnual = 365
            intDiasDiferencia = Dias360(datpFechaInicial, datpFechaFinal, True)
    End Select

    If strpTipoCalculoDias = Tipo_Calculo_Dias_Base_Anual Then
        CalculaDias = intDiasBaseAnual
    ElseIf strpTipoCalculoDias = Tipo_Calculo_Dias_Diferencia Then
        CalculaDias = intDiasDiferencia
    Else
        CalculaDias = intDiasDiferencia
    End If

End Function
Public Function CalculaCapitalVAC(ByVal strpFechaLiquidacion As String, ByVal strpFechaEmision As String, ByVal dblpValorNominal As Double, ByVal strpFechaIniCupon As String, ByVal strpFechaFinCupon As String, ByVal strpCodIndiceInicial As String, ByVal strpCodIndiceFinal As String, ByVal intpDiasCupon As Integer, ByVal intpDiasCorridos As Integer, ByVal dblpTasaCupon As Double, ByVal intpBaseAnual As Integer) As Double

    '*** Cálculo del Valor Nominal Reajustado para el Cálculo del ***
    '*** Interés Corrido y el VAC Corrido para Bonos VAC          ***
    Dim adoRecTasa              As ADODB.Recordset
    Dim dblVacEmision           As Double, dblVacLiquidacion        As Double
    Dim dblVacIniCupon          As Double, dblVacFinCupon           As Double
    Dim dblCapitalVac           As Double, dblFactor                As Double
    Dim strFechaEmisionMas1     As String, strFechaLiquidacionMas1  As String
    Dim strFechaInicuponMas1    As String, strFechaFinCuponMas1     As String
    Dim strMensaje              As String

    CalculaCapitalVAC = 0
    
    dblVacEmision = 0: dblVacLiquidacion = 0: dblVacIniCupon = 0: dblVacFinCupon = 0
    dblCapitalVac = 0
    
    strFechaEmisionMas1 = Convertyyyymmdd(DateAdd("d", 1, Convertddmmyyyy(strpFechaEmision)))
    strFechaLiquidacionMas1 = Convertyyyymmdd(DateAdd("d", 1, Convertddmmyyyy(strpFechaLiquidacion)))
    strFechaInicuponMas1 = Convertyyyymmdd(DateAdd("d", 1, Convertddmmyyyy(strpFechaIniCupon)))
    strFechaFinCuponMas1 = Convertyyyymmdd(DateAdd("d", 1, Convertddmmyyyy(strpFechaFinCupon)))
    
    '*** Obtener las Tasas VAC: Emisión, Liquidación, Cupón Inicial y Cupón Final ***
    dblVacEmision = ObtenerTasaAjuste(Codigo_Tipo_Ajuste_Vac, "00", strpFechaEmision, strFechaEmisionMas1)
    dblVacLiquidacion = ObtenerTasaAjuste(Codigo_Tipo_Ajuste_Vac, "00", strpFechaLiquidacion, strFechaLiquidacionMas1)
    dblVacIniCupon = ObtenerTasaAjuste(Codigo_Tipo_Ajuste_Vac, "00", strpFechaIniCupon, strFechaInicuponMas1)
    dblVacFinCupon = ObtenerTasaAjuste(Codigo_Tipo_Ajuste_Vac, "00", strpFechaFinCupon, strFechaFinCuponMas1)
    
'    With adoComm
'        Set adoRecTasa = New ADODB.Recordset

        '*** Obtener las Tasas VAC: Emisión, Liquidación, Cupón Inicial y Cupón Final ***
'        .CommandText = "SELECT FechaRegistro,ValorTasa FROM InversionTasa " & _
'            "WHERE CodTasa='" & Codigo_Tipo_Ajuste_Vac & "' AND " & _
'            "((FechaRegistro>='" & strpFechaEmision & "' AND FechaRegistro<'" & strFechaEmisionMas1 & "') OR " & _
'            "(FechaRegistro>='" & strpFechaLiquidacion & "' AND FechaRegistro<'" & strFechaLiquidacionMas1 & "') OR " & _
'            "(FechaRegistro>='" & strpFechaIniCupon & "' AND FechaRegistro<'" & strFechaInicuponMas1 & "') OR " & _
'            "(FechaRegistro>='" & strpFechaFinCupon & "'AND FechaRegistro<'" & strFechaFinCuponMas1 & "'))"
'        Set adoRecTasa = .Execute
'
'        Do While Not adoRecTasa.EOF
'            Select Case Convertyyyymmdd(adoRecTasa("FechaRegistro"))
'                Case strpFechaEmision                  '*** A la fecha de emisión          ***
'                    dblVacEmision = adoRecTasa("ValorTasa")
'                    If strpFechaEmision = strpFechaLiquidacion Then
'                        dblVacLiquidacion = adoRecTasa("ValorTasa")
'                    End If
'                Case strpFechaLiquidacion                    '*** A la fecha de liquidación      ***
'                    dblVacLiquidacion = adoRecTasa("ValorTasa")
'                    If strpFechaIniCupon = strpFechaLiquidacion Then
'                        dblVacIniCupon = adoRecTasa("ValorTasa")
'                    End If
'                    If strpFechaFinCupon = strpFechaLiquidacion Then
'                        dblVacFinCupon = adoRecTasa("ValorTasa")
'                    End If
'                Case strpFechaIniCupon                   '*** A la fecha de inicio del cupón ***
'                    dblVacIniCupon = adoRecTasa("ValorTasa")
'                Case strpFechaFinCupon                   '*** A la fecha de corte del cupón  ***
'                    dblVacFinCupon = adoRecTasa("ValorTasa")
'            End Select
'            adoRecTasa.MoveNext
'        Loop
'        adoRecTasa.Close: Set adoRecTasa = Nothing
'    End With

    '*** Si es VAC Periodico proyectar Tasa VAC a la fecha de liquidación ***
'    If strpCodIndiceInicial = Codigo_Vac_Emision Then
'        If strpCodIndiceFinal = Codigo_Vac_Liquidacion Then '*** A partir del cupón anterior ***
'            If dblVacIniCupon > 0 And intpDiasCupon > 0 Then
'                dblFactor = (((dblVacFinCupon / dblVacIniCupon) ^ (intpBaseAnual / intpDiasCupon)) * (1 + dblpTasaCupon)) - 1
'                dblVacLiquidacion = (1 + dblFactor) ^ (intpDiasCorridos / intpBaseAnual)
''                dblVacLiquidacion = (dblVacFinCupon / dblVacIniCupon) ^ (intpDiasCorridos / intpDiasCupon)
'            Else
'                MsgBox "La Tasa VAC a la Fecha de Liquidación No Existe, la Operación no se puede realizar.", vbCritical, "Aviso"
'                dblCapitalVac = 0: Exit Function
'            End If
'        Else                                    '*** A partir del cupón vigente ***
'            If dblVacLiquidacion = 0 Then
'                MsgBox "La Tasa VAC a la Fecha de Liquidación No Existe, la Operación no se puede realizar.", vbCritical, "Aviso"
'                dblCapitalVac = 0: Exit Function
'            Else
'                If dblVacEmision > 0 Then
'                    dblVacLiquidacion = (dblVacFinCupon / dblVacEmision)
'                Else
'                    MsgBox "La Tasa VAC a la Fecha de EMISION No Existe, la Operación no se puede realizar.", vbCritical, "Aviso"
'                    dblCapitalVac = 0: Exit Function
'                End If
'            End If
'        End If
'    End If

    '*** Validar Indices VAC ***
    strMensaje = "Falta Registrar el Indice VAC del:" & vbNewLine & vbNewLine
    
    Select Case strpCodIndiceInicial
        Case Codigo_Vac_Emision
            If dblVacEmision = 0 Then strMensaje = strMensaje & CStr(Convertddmmyyyy(strpFechaEmision)) & vbNewLine
        Case Codigo_Vac_InicioPrimerCupon, Codigo_Vac_InicioCuponVigente, Codigo_Vac_InicioCuponAnterior
            If dblVacIniCupon = 0 Then strMensaje = strMensaje & CStr(Convertddmmyyyy(strpFechaIniCupon)) & vbNewLine
    End Select
    
    Select Case strpCodIndiceFinal
        Case Codigo_Vac_Liquidacion
            If dblVacLiquidacion = 0 Then strMensaje = strMensaje & CStr(Convertddmmyyyy(strpFechaLiquidacion)) & vbNewLine
        Case Codigo_Vac_FinPrimerCupon, Codigo_Vac_FinCuponVigente, Codigo_Vac_FinCuponAnterior
            If dblVacIniCupon = 0 Then strMensaje = strMensaje & CStr(Convertddmmyyyy(strpFechaFinCupon)) & vbNewLine
    End Select
    
    '*** Cálculo del Valor Nominal Reajustado ***
    If strpCodIndiceInicial = Codigo_Vac_Emision Then dblVacIniCupon = dblVacEmision
    If strpCodIndiceFinal = Codigo_Vac_Liquidacion Then dblVacFinCupon = dblVacLiquidacion
    
    If dblVacIniCupon > 0 And dblVacFinCupon > 0 Then
    'If dblVacEmision > 0 And dblVacLiquidacion > 0 Then
'        If strpCodIndiceInicial = Codigo_Vac_Emision Then          '*** VAC Periodico      ***
'            dblCapitalVac = Round(dblpValorNominal * dblVacLiquidacion, 2)
'        Else                           '*** VAC Al Vencimiento ***
'            dblCapitalVac = Round(dblpValorNominal * (dblVacLiquidacion / dblVacEmision), 2)
'        End If
        dblCapitalVac = Round(dblpValorNominal * (dblVacFinCupon / dblVacIniCupon), 2)
    Else
        dblCapitalVac = 0
        MsgBox strMensaje, vbCritical, frmMainMdi.Caption
    End If

    CalculaCapitalVAC = dblCapitalVac
    
End Function

Public Function ObtenerGrupoCuentaContable(ByVal strpCodAdministradora As String, ByVal strpCodCuenta As String, ByVal intNumVersion As Integer) As String

    '*** Obtener Cuenta Contable Tipo ***
    Dim adoRegistro     As ADODB.Recordset
   
    ObtenerGrupoCuentaContable = Valor_Caracter
       
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT dbo.uf_CNObtenerGrupoCuenta('" & _
            strpCodAdministradora & "','" & strpCodCuenta & "'," & intNumVersion & ") AS CodGrupoCuenta"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            ObtenerGrupoCuentaContable = adoRegistro("CodGrupoCuenta")
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Function
Public Function CalculaCapital(ByVal strpCodTipoAjuste, ByVal strpFechaLiquidacion As String, ByVal strpFechaEmision As String, ByVal dblpValorNominal As Double, ByVal strpFechaIniCupon As String, ByVal strpCodIndiceInicial As String, ByVal intpDiasCupon As Integer, ByVal intpDiasCorridos As Integer, ByVal dblpTasaCupon As Double, ByVal strpBaseAnual As String, ByVal intpBaseAnual As Integer) As Double

    '*** Cálculo del Valor Nominal Reajustado para el Cálculo del Interés Corrido ***
    Dim adoRegistro             As ADODB.Recordset
    Dim dblTasaEmision          As Double, dblTasaLiquidacion       As Double
    Dim dblTasaIniCupon         As Double
    Dim dblCapital              As Double, dblFactor                As Double
    Dim strFechaEmisionMas1     As String, strFechaLiquidacionMas1  As String
    Dim strFechaInicuponMas1    As String
    Dim strMensaje              As String

    CalculaCapital = 0
    
    dblTasaEmision = 0: dblTasaLiquidacion = 0: dblTasaIniCupon = 0
    dblCapital = 0
    
    strFechaEmisionMas1 = Convertyyyymmdd(DateAdd("d", 1, Convertddmmyyyy(strpFechaEmision)))
    strFechaLiquidacionMas1 = Convertyyyymmdd(DateAdd("d", 1, Convertddmmyyyy(strpFechaLiquidacion)))
    strFechaInicuponMas1 = Convertyyyymmdd(DateAdd("d", 1, Convertddmmyyyy(strpFechaIniCupon)))
    
    '*** Obtener las Tasas: Emisión, Liquidación y Cupón Inicial ***
    dblTasaEmision = ObtenerTasaAjuste(strpCodTipoAjuste, strpCodIndiceInicial, strpFechaEmision, strFechaEmisionMas1)
    dblTasaLiquidacion = ObtenerTasaAjuste(strpCodTipoAjuste, strpCodIndiceInicial, strpFechaLiquidacion, strFechaLiquidacionMas1)
    dblTasaIniCupon = ObtenerTasaAjuste(strpCodTipoAjuste, strpCodIndiceInicial, strpFechaIniCupon, strFechaInicuponMas1)
    
    '*** Validar Indices ***
    strMensaje = "Falta Registrar el Indice del:" & vbNewLine & vbNewLine
    
    If dblTasaLiquidacion = 0 Then strMensaje = strMensaje & CStr(Convertddmmyyyy(strpFechaLiquidacion)) & vbNewLine
    
    '*** Cálculo del Valor Nominal Reajustado ***
    dblTasaIniCupon = dblTasaLiquidacion
    
    If dblTasaIniCupon > 0 Then
        If strpBaseAnual = Codigo_Base_Actual_Actual Then
            Set adoRegistro = New ADODB.Recordset
            
            adoComm.CommandText = "SELECT dbo.uf_ACValidaEsBisiesto(" & CInt(Left(strpFechaLiquidacion, 4)) & ") AS 'EsBisiesto'"
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                If adoRegistro("EsBisiesto") = 0 Then intpBaseAnual = 366
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
        End If
        
        dblCapital = Round(dblpValorNominal * ((dblTasaIniCupon * 0.01) + dblpTasaCupon) * (intpDiasCupon / intpBaseAnual), 2)
    Else
        dblCapital = 0
        MsgBox strMensaje, vbCritical, frmMainMdi.Caption
    End If

    CalculaCapital = dblCapital
    
End Function


Public Function FactorAnual(ByVal dblpFactor As Double, ByVal intpCantDiasPeriodo As Integer, ByVal intpBaseCalculo As Integer, ByVal strpCodTipoTasa As String, ByVal strpIndCapitalizable As String, ByVal strpCodFormaCalculo As String, ByVal intpIndCorrecto As Integer, ByVal intpCantDiasCupon As Integer, ByVal intpNumPeriodosAnual As Integer) As Double

    Dim dblFactorAnual      As Double
    
    FactorAnual = 0
    
    If strpCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then '*** Efectiva ***
       dblFactorAnual = ((1 + (0.01 * dblpFactor)) ^ (intpCantDiasPeriodo / intpBaseCalculo)) - 1
    ElseIf strpIndCapitalizable = Valor_Indicador Then '*** Nominal Capitalizable ***
        If strpCodFormaCalculo = Codigo_Calculo_Prorrateo Then
            dblFactorAnual = (0.01 * dblpFactor) * intpCantDiasPeriodo / intpBaseCalculo
        Else
            dblFactorAnual = (0.01 * dblpFactor) / (intpBaseCalculo / intpCantDiasPeriodo)
        End If
        
       If intpIndCorrecto = 1 Then
            If strpCodFormaCalculo = Codigo_Calculo_Normal Then
                dblFactorAnual = dblFactorAnual / (intpCantDiasCupon / intpCantDiasPeriodo)
            End If
       End If
    ElseIf strpIndCapitalizable = Valor_Caracter Then '*** Nominal No Capitalizable ***
       dblFactorAnual = (0.01 * dblpFactor) / intpNumPeriodosAnual
       If intpIndCorrecto = 1 Then
            dblFactorAnual = dblFactorAnual / (intpCantDiasCupon / intpCantDiasPeriodo)
       End If
    End If
    
    FactorAnual = dblFactorAnual
            
End Function

Public Function FactorDiario(ByVal dblpFactor As Double, ByVal intpCantDiasPeriodo As Integer, ByVal strpCodTipoTasa As String, ByVal strpIndCapitalizable As String, ByVal intpCantDiasCupon As Integer) As Double

    Dim dblFactorDiario      As Double
    
    FactorDiario = 0
    
    If strpCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then '*** Efectiva ***
       dblFactorDiario = ((1 + dblpFactor) ^ (1 / intpCantDiasPeriodo)) - 1
    ElseIf strpIndCapitalizable = Valor_Indicador Then '*** Nominal Capitalizable ***
       dblFactorDiario = dblpFactor / intpCantDiasPeriodo
    ElseIf strpIndCapitalizable = Valor_Caracter Then '*** Nominal No Capitalizable ***
       dblFactorDiario = dblpFactor / intpCantDiasPeriodo
    End If
    
    FactorDiario = dblFactorDiario
            
End Function
Public Function FactorDiarioImplicito(ByVal dblpMontoMFL1 As Double, ByVal dblpMontoMFL2 As Double, ByVal intpCantDiasPeriodo As Integer) As Double

    Dim dblFactorDiario      As Double
    
    FactorDiarioImplicito = 0
    
    dblFactorDiario = ((dblpMontoMFL2 / dblpMontoMFL1) ^ (1 / intpCantDiasPeriodo)) - 1
    
    FactorDiarioImplicito = dblFactorDiario
            
End Function

Public Function FactorDiarioNormal(ByVal dblpFactor As Double, ByVal intpCantDiasPeriodo As Integer, ByVal strpCodTipoTasa As String, ByVal strpIndCapitalizable As String, ByVal intpCantDiasCupon As Integer) As Double

    Dim dblFactorDiario      As Double
    
    FactorDiarioNormal = 0
    
    If strpCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then '*** Efectiva ***
       dblFactorDiario = ((1 + dblpFactor) ^ (1 / intpCantDiasPeriodo)) - 1
    ElseIf strpIndCapitalizable = Valor_Indicador Then '*** Nominal Capitalizable ***
       dblFactorDiario = dblpFactor / intpCantDiasCupon
    ElseIf strpIndCapitalizable = Valor_Caracter Then '*** Nominal No Capitalizable ***
       dblFactorDiario = dblpFactor / intpCantDiasPeriodo
    End If
    
    FactorDiarioNormal = dblFactorDiario
            
End Function
Public Function FactorAnualNormal(ByVal dblpFactor As Double, ByVal intpCantDiasPeriodo As Integer, ByVal intpBaseCalculo As Integer, ByVal strpCodTipoTasa As String, ByVal strpIndCapitalizable As String, ByVal strpCodFormaCalculo As String, ByVal intpIndCorrecto As Integer, ByVal intpCantDiasCupon As Integer, ByVal intpNumPeriodosAnual As Integer) As Double

    Dim dblFactorAnual      As Double
    
    FactorAnualNormal = 0
    
    If strpCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then '*** Efectiva ***
       dblFactorAnual = ((1 + (0.01 * dblpFactor)) ^ (intpCantDiasPeriodo / intpBaseCalculo)) - 1
    ElseIf strpIndCapitalizable = Valor_Indicador Then '*** Nominal Capitalizable ***
        If strpCodFormaCalculo = Codigo_Calculo_Prorrateo Then
            dblFactorAnual = (0.01 * dblpFactor) * intpCantDiasPeriodo / intpBaseCalculo
        Else
            dblFactorAnual = (0.01 * dblpFactor) / (intpBaseCalculo / intpCantDiasPeriodo)
        End If
        
       If intpIndCorrecto = 1 Then
            If strpCodFormaCalculo = Codigo_Calculo_Normal Then
                dblFactorAnual = dblFactorAnual / (intpCantDiasCupon / intpCantDiasPeriodo)
            End If
       End If
    ElseIf strpIndCapitalizable = Valor_Caracter Then '*** Nominal No Capitalizable ***
       dblFactorAnual = (0.01 * dblpFactor) / intpNumPeriodosAnual
       If intpIndCorrecto = 1 Then
            dblFactorAnual = dblFactorAnual / (intpCantDiasCupon / intpCantDiasPeriodo)
       End If
    End If
    
    FactorAnualNormal = dblFactorAnual
            
End Function

Public Function ObtenerTasaAjuste(ByVal strpCodTipoAjuste As String, ByVal strpCodClaseAjuste, ByVal strpFechaInicial As String, ByVal strpFechaFinal As String) As Double

    '*** Función que permite obtener la tasa de ajuste a una fecha determinada  ***
    '*** Si no existe en el rango indicado obtiene la última tasa registrada    ***
    '*** Parámetros :                                                           ***
    '*** Código Tipo Tasa de Ajuste                                             ***
    '*** Código Clase de Tasa de Ajuste                                         ***
    '*** Fecha Inicial (yyyymmdd)                                               ***
    '*** Fecha Final (yyyymmdd)                                                 ***
    
    Dim adoRegistroTasa     As ADODB.Recordset
    Dim dblValorTasa        As Double
    
    ObtenerTasaAjuste = 0
    
    With adoComm
        Set adoRegistroTasa = New ADODB.Recordset

        '*** Obtener la tasa correspondiente a la fecha determinada ***
        .CommandText = "SELECT FechaRegistro,ValorTasa FROM InversionTasa " & _
            "WHERE CodClaseTasa='" & strpCodClaseAjuste & "' AND CodTasa='" & strpCodTipoAjuste & "' AND " & _
            "(FechaRegistro>='" & strpFechaInicial & "' AND FechaRegistro<'" & strpFechaFinal & "')"
        Set adoRegistroTasa = .Execute
        
        If Not adoRegistroTasa.EOF Then
            If IsNull(adoRegistroTasa("ValorTasa")) Then
                dblValorTasa = 0
            Else
                dblValorTasa = adoRegistroTasa("ValorTasa")
            End If
        End If
        adoRegistroTasa.Close
        
        '*** No se encontró tasa registrada ***
        If dblValorTasa = 0 Then
            '*** Es la fecha de liquidación ? ***
            If strpFechaInicial >= gstrFechaActual Then
                Dim strFechaMaxInicial      As String, strFechaMaxFinal     As String
                
                '*** Obtener la fecha de la última tasa registrada ***
                .CommandText = "SELECT MAX(FechaRegistro) FechaRegistro FROM InversionTasa " & _
                    "WHERE CodClaseTasa='" & strpCodClaseAjuste & "' AND CodTasa='" & strpCodTipoAjuste & "'"
                Set adoRegistroTasa = .Execute
                
                If Not adoRegistroTasa.EOF Then
                    If IsNull(adoRegistroTasa("FechaRegistro")) Then
                        strFechaMaxInicial = Valor_Caracter
                        strFechaMaxFinal = Valor_Caracter
                    Else
                        strFechaMaxInicial = Convertyyyymmdd(adoRegistroTasa("FechaRegistro"))
                        strFechaMaxFinal = Convertyyyymmdd(DateAdd("d", 1, adoRegistroTasa("FechaRegistro")))
                    End If
                End If
                adoRegistroTasa.Close
                
                If strFechaMaxInicial <> Valor_Caracter Then
                    '*** Obtener la tasa correspondiente a la fecha determinada ***
                    .CommandText = "SELECT FechaRegistro,ValorTasa FROM InversionTasa " & _
                        "WHERE CodClaseTasa='" & strpCodClaseAjuste & "' AND CodTasa='" & strpCodTipoAjuste & "' AND " & _
                        "(FechaRegistro>='" & strFechaMaxInicial & "' AND FechaRegistro<'" & strFechaMaxFinal & "')"
                    Set adoRegistroTasa = .Execute
                    
                    If Not adoRegistroTasa.EOF Then
                        If IsNull(adoRegistroTasa("ValorTasa")) Then
                            dblValorTasa = 0
                        Else
                            dblValorTasa = adoRegistroTasa("ValorTasa")
                        End If
                    End If
                    adoRegistroTasa.Close
                End If
                
            End If
        End If
        Set adoRegistroTasa = Nothing
    End With
    
    ObtenerTasaAjuste = dblValorTasa
    
End Function
Public Function ControlErrores() As Integer

    Dim intTipoMensaje      As Integer, strMensaje      As String
    Dim intRespuesta        As Integer
    
    '*** Valor de Devolución        Significado         ***
    '*** 0                          Resume              ***
    '*** 1                          Resume Next         ***
    '*** 2                          Error desconocido   ***
    intTipoMensaje = vbExclamation
    Select Case err.Number
        Case Codigo_Error_RegistroDuplicado
            strMensaje = Mensaje_Registro_Duplicado
            intTipoMensaje = vbExclamation + vbOKOnly
        Case Codigo_Error_ArchivoNoExiste
            strMensaje = "Archivo NO EXISTE"
            intTipoMensaje = vbExclamation + vbOKOnly
        Case Else
            strMensaje = "Número de Error" & Space(1) & ":" & Space(1) & _
                CStr(err.Number) & vbNewLine & "Descripción" & Space(1) & ":" & Space(1) & err.Description & _
                vbNewLine & vbNewLine & Mensaje_Proceso_NoExitoso
            intTipoMensaje = vbExclamation + vbRetryCancel
'            ControlErrores = 3
'            Exit Function
    End Select
    
    intRespuesta = MsgBox(strMensaje, intTipoMensaje, "Control de Errores")
    Select Case intRespuesta
        Case 1 '*** Aceptar ***
            ControlErrores = 2
        Case 1, 4   '*** Reintentar ***
            ControlErrores = 0
        Case 5      '*** Ignorar ***
            ControlErrores = 1
        Case 2, 3   '*** Cancelar,Finalizar ***
            ControlErrores = 2
        Case Else
            ControlErrores = 3
    End Select

End Function

Public Sub GuardarBitacoraSistema(ByVal strpCodFondo As String, ByVal strpCodAdministradora As String, ByVal strpFechaLog As String, ByVal strpDescripConcepto As String, ByVal strpDescripCampo As String, ByVal strpValorAnterior As String, ByVal strpValorActual As String, ByVal strpIdUsuario As String)

    adoComm.CommandText = "{ call up_ACAdicBitacoraSistema('" & strpCodFondo & "','" & _
        strpCodAdministradora & "','" & strpFechaLog & "','" & strpDescripConcepto & "','" & _
        strpDescripCampo & "','" & strpValorAnterior & "','" & strpValorActual & "','" & _
        strpIdUsuario & "') }"
    adoConn.Execute adoComm.CommandText
    
End Sub

Public Function ObtenerSecuencialInversionOperacion(ByVal strpCodFondo As String, ByVal strpCodParametro As String) As String

    Dim strNumParametro As String
    
    ObtenerSecuencialInversionOperacion = Valor_Caracter
    
    With adoComm
        '*** Obtener Secuencial ***
        .CommandType = adCmdStoredProc
        
        .CommandText = "up_ACObtenerUltNumero"
        .Parameters.Append .CreateParameter("CodFondo", adChar, adParamInput, 3, strpCodFondo)
        .Parameters.Append .CreateParameter("CodAdministradora", adChar, adParamInput, 3, gstrCodAdministradora)
        .Parameters.Append .CreateParameter("CodParametro", adChar, adParamInput, 6, strpCodParametro)
        .Parameters.Append .CreateParameter("NuevoNumero", adChar, adParamOutput, 10, "")
        .Execute
        
        If Not .Parameters("NuevoNumero") Then
            strNumParametro = .Parameters("NuevoNumero").Value
            .Parameters.Delete ("CodFondo")
            .Parameters.Delete ("CodAdministradora")
            .Parameters.Delete ("CodParametro")
            .Parameters.Delete ("NuevoNumero")
        End If
                        
        .CommandType = adCmdText
        .Parameters.Refresh  'colocar esto sino se cae cuando es llamado mas de una vez....
    
    End With
    
    ObtenerSecuencialInversionOperacion = strNumParametro
    
End Function

'Public Function ObtenerTipoCambio(ByVal strpCodTipoCambio As String, ByVal strpCodValorCambio As String, ByVal datpFechaCierre As Date, ByVal strpCodMoneda As String) As Double
'
'    Dim dblValorParametro As Double
'
'    ObtenerTipoCambio = 0
'
'    With adoComm
'        '*** Obtener Secuencial ***
'        .CommandType = adCmdStoredProc
'
'        .CommandText = "up_ACObtenerTipoCambio"
'        .Parameters.Append .CreateParameter("@CodTipoCambio", adChar, adParamInput, 2, strpCodTipoCambio)
'        .Parameters.Append .CreateParameter("@CodValorCambio", adChar, adParamInput, 10, strpCodValorCambio)
'        .Parameters.Append .CreateParameter("@FechaCierre", adDate, adParamInput, 10, datpFechaCierre)
'        .Parameters.Append .CreateParameter("@CodMoneda", adChar, adParamInput, 2, strpCodMoneda)
'        .Parameters.Append .CreateParameter("@ValorTipoCambio", adDouble, adParamOutput, 11, 0)
'        .Execute
'
'        If Not .Parameters("@ValorTipoCambio") Then
'            dblValorParametro = .Parameters("@ValorTipoCambio").Value
'        End If
'
'        .Parameters.Delete ("@CodTipoCambio")
'        .Parameters.Delete ("@CodValorCambio")
'        .Parameters.Delete ("@FechaCierre")
'        .Parameters.Delete ("@CodMoneda")
'        .Parameters.Delete ("@ValorTipoCambio")
'
'        '.Parameters.Refresh
'
'        .CommandType = adCmdText
'    End With
'
'    ObtenerTipoCambio = dblValorParametro
'
'End Function
'        up_ACObtenerTipoCambioXML]
'(
' @CodMoneda                     Codigo
',@CodMonedaContable             Codigo
',@FechaTipoCambio               datetime
',@TipoCambioReemplazoXML        XML             = '<TipoCambioReemplazo />'
',@ValorTipoCambio               decimal(14,8)   OUTPUT
',@CodTipoCambio                 Codigo          = '04' --CONASEV
',@CodClaseTipoCambio            Codigo          = '01' --COMPRA
',@ModalidadCambio               smallint        = 5    --POR DEFECTO BUSQUEDA MULTIPLE DEL TIPO DE CAMBIO!
')
Public Function ObtenerTipoCambioMonedaXML(ByVal strpCodMonedaOrigen As String, ByVal strpCodMonedaDestino As String, ByVal strpFechaTipoCambio As String, Optional ByVal strpTipoCambioReemplazoXML As String, Optional ByVal strpCodTipoCambio As String = "04", Optional ByVal strpCodValorCambio As String = "01", Optional ByVal numpCodModalidadCalculo As Integer = Tipo_Busqueda_Tipo_Cambio_Multiple) As Double

    Dim dblValorParametro As Double
    
    ObtenerTipoCambioMonedaXML = 0
    
    With adoComm
        '*** Obtener Secuencial ***
        .CommandType = adCmdText
        .CommandType = adCmdStoredProc
        
        
' @CodMoneda                     Codigo
',@CodMonedaContable             Codigo
',@FechaTipoCambio               datetime
',@TipoCambioReemplazoXML        XML             = '<TipoCambioReemplazo />'
',@ValorTipoCambio               decimal(14,8)   OUTPUT
',@CodTipoCambio                 Codigo          = '04' --CONASEV
',@CodClaseTipoCambio            Codigo          = '01' --COMPRA
',@ModalidadCambio               smallint        = 5    --POR DEFECTO BUSQUEDA MULTIPLE DEL TIPO DE CAMBIO!
        
        .CommandText = "up_ACObtenerTipoCambioXML"
        .Parameters.Append .CreateParameter("@CodMoneda", adChar, adParamInput, 2, strpCodMonedaOrigen)
        .Parameters.Append .CreateParameter("@CodMonedaContable", adChar, adParamInput, 10, strpCodMonedaDestino)
        .Parameters.Append .CreateParameter("@FechaTipoCambio", adChar, adParamInput, 8, strpFechaTipoCambio)
        .Parameters.Append .CreateParameter("@TipoCambioReemplazoXML", adLongVarChar, adParamInput, 4000, strpTipoCambioReemplazoXML)
        .Parameters.Append .CreateParameter("@ValorTipoCambio", adDouble, adParamOutput, 11, 0)
        .Parameters.Append .CreateParameter("@CodTipoCambio", adChar, adParamInput, 2, strpCodTipoCambio)
        .Parameters.Append .CreateParameter("@CodClaseTipoCambio", adChar, adParamInput, 2, strpCodValorCambio)
        .Parameters.Append .CreateParameter("@ModalidadCambio", adSmallInt, adParamInput, 3, numpCodModalidadCalculo)
        .Execute
        
        If Not .Parameters("@ValorTipoCambio") Then
            dblValorParametro = .Parameters("@ValorTipoCambio").Value
        End If
        
        .Parameters.Delete ("@CodTipoCambio")
        .Parameters.Delete ("@CodClaseTipoCambio")
        .Parameters.Delete ("@FechaTipoCambio")
        .Parameters.Delete ("@CodMoneda")
        .Parameters.Delete ("@CodMonedaContable")
        .Parameters.Delete ("@ModalidadCambio")
        .Parameters.Delete ("@ValorTipoCambio")
        .Parameters.Delete ("@TipoCambioReemplazoXML")
                        
        '.Parameters.Refresh
        
        .CommandType = adCmdText
    End With
    
    ObtenerTipoCambioMonedaXML = dblValorParametro
    
End Function


Public Function ObtenerTipoCambioMoneda(ByVal strpCodTipoCambio As String, ByVal strpCodValorCambio As String, ByVal datpFechaCierre As Date, ByVal strpCodMonedaOrigen As String, ByVal strpCodMonedaDestino As String, Optional ByVal strpCodModalidadCalculo = Tipo_Busqueda_Tipo_Cambio_Multiple) As Double

    Dim dblValorParametro As Double
    
    ObtenerTipoCambioMoneda = 0
    
    With adoComm
        '*** Obtener Secuencial ***
        .CommandType = adCmdText
        .CommandType = adCmdStoredProc
        
        .CommandText = "up_ACObtenerTipoCambioMoneda1"
        .Parameters.Append .CreateParameter("@CodTipoCambio", adChar, adParamInput, 2, strpCodTipoCambio)
        .Parameters.Append .CreateParameter("@CodValorCambio", adChar, adParamInput, 10, strpCodValorCambio)
        .Parameters.Append .CreateParameter("@FechaTipoCambio", adDate, adParamInput, 10, datpFechaCierre)
        .Parameters.Append .CreateParameter("@CodMoneda", adChar, adParamInput, 2, strpCodMonedaOrigen)
        .Parameters.Append .CreateParameter("@CodMonedaCambio", adChar, adParamInput, 2, strpCodMonedaDestino)
        .Parameters.Append .CreateParameter("@ModalidadCambio", adSmallInt, adParamInput, 2, CInt(strpCodModalidadCalculo))
        .Parameters.Append .CreateParameter("@ValorTipoCambio", adDouble, adParamOutput, 11, 0)
        .Execute
        
        If Not .Parameters("@ValorTipoCambio") Then
            dblValorParametro = .Parameters("@ValorTipoCambio").Value
        End If
        
        .Parameters.Delete ("@CodTipoCambio")
        .Parameters.Delete ("@CodValorCambio")
        .Parameters.Delete ("@FechaTipoCambio")
        .Parameters.Delete ("@CodMoneda")
        .Parameters.Delete ("@CodMonedaCambio")
        .Parameters.Delete ("@ModalidadCambio")
        .Parameters.Delete ("@ValorTipoCambio")
                        
        '.Parameters.Refresh
        
        .CommandType = adCmdText
    End With
    
    ObtenerTipoCambioMoneda = dblValorParametro
    
End Function


Public Function ObtenerTipoCambioMoneda2(ByVal strpCodTipoCambio As String, ByVal strpCodValorCambio As String, ByVal datpFechaCierre As Date, ByVal strpCodMonedaOrigen As String, ByVal strpCodMonedaDestino As String, Optional ByVal strpCodModalidadCalculo = Tipo_Busqueda_Tipo_Cambio_Multiple) As Double

    Dim dblValorParametro As Double
    
    ObtenerTipoCambioMoneda2 = 0
    
    With adoComm
        '*** Obtener Secuencial ***
        .CommandType = adCmdText
        .CommandType = adCmdStoredProc
        
        .CommandText = "up_ACObtenerTipoCambioMoneda2"
        .Parameters.Append .CreateParameter("@CodTipoCambio", adChar, adParamInput, 2, strpCodTipoCambio)
        .Parameters.Append .CreateParameter("@CodValorCambio", adChar, adParamInput, 10, strpCodValorCambio)
        .Parameters.Append .CreateParameter("@FechaTipoCambio", adDate, adParamInput, 10, datpFechaCierre)
        .Parameters.Append .CreateParameter("@CodMoneda", adChar, adParamInput, 2, strpCodMonedaOrigen)
        .Parameters.Append .CreateParameter("@CodMonedaCambio", adChar, adParamInput, 2, strpCodMonedaDestino)
        .Parameters.Append .CreateParameter("@ModalidadCambio", adSmallInt, adParamInput, 2, CInt(strpCodModalidadCalculo))
        .Parameters.Append .CreateParameter("@ValorTipoCambio", adDouble, adParamOutput, 11, 0)
        .Execute
        
        If Not .Parameters("@ValorTipoCambio") Then
            dblValorParametro = .Parameters("@ValorTipoCambio").Value
        End If
        
        .Parameters.Delete ("@CodTipoCambio")
        .Parameters.Delete ("@CodValorCambio")
        .Parameters.Delete ("@FechaTipoCambio")
        .Parameters.Delete ("@CodMoneda")
        .Parameters.Delete ("@CodMonedaCambio")
        .Parameters.Delete ("@ModalidadCambio")
        .Parameters.Delete ("@ValorTipoCambio")
                        
        '.Parameters.Refresh
        
        .CommandType = adCmdText
    End With
    
    ObtenerTipoCambioMoneda2 = dblValorParametro
    
End Function

Public Function ObtenerHoraServidor() As String

    Dim adoRegistro     As ADODB.Recordset
    Dim strHoraServidor As String
    
    ObtenerHoraServidor = "00:00"
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        .CommandText = "{ call up_ACSelDatos(0) }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strHoraServidor = Format(adoRegistro("HoraServidor"), "hh:mm")
        Else
            adoRegistro.Close: Set adoRegistro = Nothing
            Exit Function
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    ObtenerHoraServidor = strHoraServidor
    
End Function
Public Function ObtenerComisionParticipacion(ByVal strpCodComision As String, ByVal strpCodFondo As String, ByVal strpCodAdministradora As String) As Double

    Dim adoRegistro As ADODB.Recordset
    Dim dblComision As Double
    
    ObtenerComisionParticipacion = 0
    
    With adoComm
        Set adoRegistro = New ADODB.Recordset
        .CommandText = "SELECT PorcenComision,MontoComision,IndRango FROM FondoComision " & _
            "WHERE CodComision='" & strpCodComision & "' AND CodFondo='" & strpCodFondo & "' AND " & _
            "CodAdministradora='" & strpCodAdministradora & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            dblComision = CDbl(adoRegistro("PorcenComision")) / 100
        Else
            adoRegistro.Close: Set adoRegistro = Nothing
            Exit Function
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    ObtenerComisionParticipacion = dblComision
    
End Function
Public Function ObtenerComisionParticipacionRango(ByVal strpCodComision As String, ByVal strpCodFondo As String, ByVal strpCodAdministradora As String, ByVal strpFechaSuscripcion As String, ByVal strpFechaRescate As String) As Double

    Dim adoRegistro As ADODB.Recordset
    Dim dblComision As Double
    
    ObtenerComisionParticipacionRango = 0
    
    With adoComm
        Set adoRegistro = New ADODB.Recordset
        .CommandText = "SELECT FCD.PorcenComision FROM FondoComision FC " & _
            "JOIN FondoComisionDetalle FCD ON (FCD.CodFondo = FC.CodFondo AND FCD.CodAdministradora = FC.CodAdministradora AND " & _
            "FCD.CodComision = FC.CodComision) " & _
            "WHERE FC.CodComision='" & strpCodComision & "' AND FC.CodFondo='" & strpCodFondo & "' AND " & _
            "FC.CodAdministradora='" & strpCodAdministradora & "' AND FC.IndRango = 'X' AND FC.IndVigente = 'X' AND '" & _
            strpFechaRescate & "' BETWEEN '" & strpFechaSuscripcion & "' AND DATEADD(day,FCD.NumDias,'" & strpFechaSuscripcion & "')"

        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            dblComision = CDbl(adoRegistro("PorcenComision")) / 100
        Else
            adoRegistro.Close: Set adoRegistro = Nothing
            Exit Function
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    ObtenerComisionParticipacionRango = dblComision
    
End Function

Public Function ObtenerCuotasParticipe(ByVal strpCodParticipe As String, ByVal strNumCertificado As String, ByVal strpCodFondo As String, ByVal strpCodAdministradora As String, ByVal strpIndGarantia As String, ByVal strpIndBloqueo As String) As Double

    Dim adoRegistro     As ADODB.Recordset
    Dim dblCantCuotas   As Double
    
    ObtenerCuotasParticipe = 0
    
    With adoComm
        Set adoRegistro = New ADODB.Recordset
        .CommandText = "SELECT SUM(PCD.CantCuotas) TotalCuotas FROM ParticipeCertificado PC " & _
            "JOIN ParticipeCertificadoDetalle PCD ON (PC.CodFondo = PCD.CodFondo AND PC.CodAdministradora = PCD.CodAdministradora AND " & _
            "PC.NumCertificado = PCD.NumCertificado) " & _
            "WHERE PC.CodParticipe = '" & strpCodParticipe & "' AND " & _
            "PC.CodFondo = '" & strpCodFondo & "' AND " & _
            "PC.CodAdministradora = '" & strpCodAdministradora & "' AND " & _
            "PC.IndGarantia = '" & strpIndGarantia & "' AND " & _
            "PC.IndBloqueo = '" & strpIndBloqueo & "' AND PC.IndVigente = 'X'"
            
        If strNumCertificado <> "" Then
            .CommandText = .CommandText & " AND PC.NumCertificado = '" & strNumCertificado & "'"
        End If
            
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            If IsNull(adoRegistro("TotalCuotas")) Then
                adoRegistro.Close: Set adoRegistro = Nothing
                Exit Function
            Else
                dblCantCuotas = CDbl(adoRegistro("TotalCuotas"))
            End If
        Else
            adoRegistro.Close: Set adoRegistro = Nothing
            Exit Function
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    ObtenerCuotasParticipe = dblCantCuotas
    
End Function

'Public Function ObtenerCuotasParticipe(ByVal strpCodParticipe As String, ByVal strpCodFondo As String, ByVal strpCodAdministradora As String, ByVal strpIndGarantia As String, ByVal strpIndBloqueo As String) As Double
'
'    Dim adoRegistro     As ADODB.Recordset
'    Dim dblCantCuotas   As Double
'
'    ObtenerCuotasParticipe = 0
'
'    With adoComm
'        Set adoRegistro = New ADODB.Recordset
'        .CommandText = "SELECT SUM(CantCuotas) TotalCuotas FROM ParticipeCertificado " & _
'            "WHERE CodParticipe='" & strpCodParticipe & "' AND CodFondo='" & strpCodFondo & "' AND " & _
'            "CodAdministradora='" & strpCodAdministradora & "' AND IndGarantia='" & strpIndGarantia & "' AND " & _
'            "IndBloqueo='" & strpIndBloqueo & "' AND IndVigente='X'"
'        Set adoRegistro = .Execute
'
'        If Not adoRegistro.EOF Then
'            If IsNull(adoRegistro("TotalCuotas")) Then
'                adoRegistro.Close: Set adoRegistro = Nothing
'                Exit Function
'            Else
'                dblCantCuotas = CDbl(adoRegistro("TotalCuotas"))
'            End If
'        Else
'            adoRegistro.Close: Set adoRegistro = Nothing
'            Exit Function
'        End If
'        adoRegistro.Close: Set adoRegistro = Nothing
'    End With
'
'    ObtenerCuotasParticipe = dblCantCuotas
'
'End Function
Public Sub ObtenerCuentasInversion(ByVal strCodFile As String, ByVal strCodDetalleFile As String, ByVal strCodMoneda As String, Optional ByVal strCodSubDetalleFile As String = "000")

    Dim adoRegistro     As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        '*** Obtener Cuentas de Inversión ***
        .CommandText = "SELECT TipoCuentaInversion,CodCuenta FROM DinamicaContable " & _
            "WHERE CodFile='" & strCodFile & "' AND (CodDetalleFile='" & strCodDetalleFile & "' OR CodDetalleFile = '000') " & _
            " AND (CodSubDetalleFile = '" & strCodSubDetalleFile & "' OR CodSubDetalleFile = '000') " & _
            " AND CodMoneda = '" & IIf(strCodMoneda = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' " & _
            "GROUP BY TipoCuentaInversion,CodCuenta " & _
            "ORDER BY TipoCuentaInversion,CodCuenta"
        Set adoRegistro = .Execute
        
        Do While Not adoRegistro.EOF
            Select Case adoRegistro("TipoCuentaInversion")
                Case Codigo_CtaInversion: strCtaInversion = Trim(adoRegistro("CodCuenta"))
                'ACR:
                Case Codigo_CtaInversionCostoSAB: strCtaInversionCostoSAB = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaInversionCostoBVL: strCtaInversionCostoBVL = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaInversionCostoCavali: strCtaInversionCostoCavali = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaInversionCostoFondoGarantia: strCtaInversionCostoFondoGarantia = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaInversionCostoConasev: strCtaInversionCostoConasev = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaInversionCostoIGV: strCtaInversionCostoIGV = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaInversionCostoCompromiso: strCtaInversionCostoCompromiso = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaInversionCostoResponsabilidad: strCtaInversionCostoResponsabilidad = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaInversionCostoFondoLiquidacion: strCtaInversionCostoFondoLiquidacion = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaInversionCostoComisionEspecial: strCtaInversionCostoComisionEspecial = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaInversionCostoGastosBancarios: strCtaInversionCostoGastosBancarios = Trim(adoRegistro("CodCuenta"))
                'ACR:
                Case Codigo_CtaProvInteres: strCtaProvInteres = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaInteres: strCtaInteres = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaInteresCastigado: strCtaInteresCastigado = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaCosto: strCtaCosto = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaIngresoOperacional: strCtaIngresoOperacional = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaInteresVencido: strCtaInteresVencido = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaVacCorrido: strCtaVacCorrido = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaXPagar: strCtaXPagar = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaXCobrar: strCtaXCobrar = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaInteresCorrido: strCtaInteresCorrido = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaProvReajusteK: strCtaProvReajusteK = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaReajusteK: strCtaReajusteK = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaProvFlucMercado: strCtaProvFlucMercado = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaFlucMercado: strCtaFlucMercado = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaProvInteresVac: strCtaProvInteresVac = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaInteresVac: strCtaInteresVac = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaIntCorridoK: strCtaIntCorridoK = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaProvFlucK: strCtaProvFlucK = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaFlucK: strCtaFlucK = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaInversionTransito: strCtaInversionTransito = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaProvGasto: strCtaProvGasto = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaCostoSAB: strCtaCostoSAB = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaCostoBVL: strCtaCostoBVL = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaCostoCavali: strCtaCostoCavali = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaCostoConasev: strCtaCostoConasev = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaCostoFondoGarantia: strCtaCostoFondoGarantia = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaCostoFondoLiquidacion: strCtaCostoFondoLiquidacion = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaCostoGastosBancarios: strCtaGastoBancario = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaCostoComisionEspecial: strCtaComisionEspecial = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaImpuesto: strCtaImpuesto = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaCompromiso: strCtaCompromiso = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaResponsabilidad: strCtaResponsabilidad = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaME: strCtaME = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaMN: strCtaMN = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaComision: strCtaComision = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaDetraccion: strCtaDetraccion = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaRetencion: strCtaRetencion = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaIngresoOperacional_AjusteRedondeo: strCtaGananciaRedondeo = Trim(adoRegistro("CodCuenta"))
                Case Codigo_Perdida_AjusteRedondeo: strCtaPerdidaRedondeo = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaProvFlucMercado_Perdida: strCtaProvFlucMercadoPerdida = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaFlucMercado_Perdida: strCtaFlucMercadoPerdida = Trim(adoRegistro("CodCuenta"))
                Case Codigo_CtaXPagarEmitida: strCtaXPagarEmitida = Trim(adoRegistro("CodCuenta"))
'                Case Codigo_CtaIngresoRendimientoPrestamo: strCtaIngresoRendimientoPrestamo = Trim(adoRegistro("CodCuenta"))
'                Case Codigo_CtaGastoRendimientoPrestamo: strCtaGastoRendimientoPrestamo = Trim(adoRegistro("CodCuenta"))
               
            End Select
            
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close
    End With
    
End Sub

Public Function ObtenerCuentaAdministracion(ByVal strpCodCuenta As String, ByVal strpTipoCuenta As String, Optional strpCodMoneda As String = Codigo_Moneda_Local) As String

    '*** Obtener Cuenta de Administración ***
    Dim strNumParametro As String
    
    ObtenerCuentaAdministracion = Valor_Caracter
    
    With adoComm
        .CommandType = adCmdStoredProc
        
        .CommandText = "up_GNObtenerCuentaAdministracion"
        .Parameters.Append .CreateParameter("CodCuentaAdministracion", adChar, adParamInput, 3, strpCodCuenta)
        .Parameters.Append .CreateParameter("TipoCuenta", adChar, adParamInput, 2, strpTipoCuenta)
        .Parameters.Append .CreateParameter("CodMoneda", adChar, adParamInput, 2, strpCodMoneda)
        .Parameters.Append .CreateParameter("CodCuenta", adChar, adParamOutput, 10, Valor_Caracter)
        .Execute
        
        If Not .Parameters("CodCuenta") Then
            strNumParametro = .Parameters("CodCuenta").Value
            .Parameters.Delete ("CodCuentaAdministracion")
            .Parameters.Delete ("TipoCuenta")
            .Parameters.Delete ("CodMoneda")
            .Parameters.Delete ("CodCuenta")
        End If
                        
        .CommandType = adCmdText
    End With
    
    ObtenerCuentaAdministracion = strNumParametro
    
End Function
Public Function ObtenerCuentaContableTipo(ByVal strpCodAdministradora As String, ByVal strpTipoCuenta As String) As String

    '*** Obtener Cuenta Contable Tipo ***
    Dim strNumParametro As String
    
    ObtenerCuentaContableTipo = Valor_Caracter
    
    With adoComm
        .CommandType = adCmdStoredProc
        
        .CommandText = "up_GNObtenerCuentaContableTipo"
        .Parameters.Append .CreateParameter("CodAdministradora", adChar, adParamInput, 3, strpCodAdministradora)
        .Parameters.Append .CreateParameter("TipoCuenta", adChar, adParamInput, 6, strpTipoCuenta)
        .Parameters.Append .CreateParameter("CodCuenta", adChar, adParamOutput, 10, Valor_Caracter)
        .Execute
        
        If Not .Parameters("CodCuenta") Then
            strNumParametro = .Parameters("CodCuenta").Value
            .Parameters.Delete ("CodAdministradora")
            .Parameters.Delete ("TipoCuenta")
            .Parameters.Delete ("CodCuenta")
        End If
                        
        .CommandType = adCmdText
    End With
    
    ObtenerCuentaContableTipo = strNumParametro
    
End Function
Public Sub GenerarFondoEstructura(ByVal strpTipoAdministradora, ByVal strpCodAdministradora, ByVal strpCodFondo As String, ByVal strCodMoneda As String, ByVal datFechaInicial As Date, ByVal datFechaFinal As Date)

    With adoComm
        .CommandText = "{ call up_GNGenFondoEstructura('" & _
            strpTipoAdministradora & "','" & strpCodAdministradora & "','" & _
            strpCodFondo & "','" & strCodMoneda & "','" & _
            Convertyyyymmdd(datFechaInicial) & "','" & Convertyyyymmdd(datFechaFinal) & "') }"
        adoConn.Execute .CommandText
    End With

End Sub

Public Sub GenerarPeriodoContable(ByVal strpTipoAdministradora, ByVal strpCodAdministradora, ByVal strpCodFondo As String, ByVal strpCodMoneda As String, ByVal datpFechaInicial As Date, ByVal datpFechaFinal As Date)

    With adoComm
        .CommandText = "{ call up_CNGenPeriodoFondo('" & _
            strpTipoAdministradora & "','" & gstrCodAdministradora & "','" & _
            strpCodFondo & "','" & strpCodMoneda & "','" & _
            Convertyyyymmdd(datpFechaInicial) & "','" & Convertyyyymmdd(datpFechaFinal) & "') }"
        adoConn.Execute .CommandText
    End With

End Sub


'Public Sub GenerarPeriodoContable(ByVal strpTipoAdministradora, ByVal strpCodAdministradora, ByVal strpCodFondo As String, ByVal strCodMoneda As String, ByVal datFechaInicial As Date, ByVal datFechaFinal As Date, ByVal ctlControl As Control)
'
'    Dim intPeriodos         As Integer, blnPrimero  As Boolean
'    Dim datFechaActualizada As Date, datFechaPago   As Date
'    Dim arrMes()            As String
'
'    ReDim arrMes(12)
'    arrMes(1) = "Enero": arrMes(2) = "Febrero": arrMes(3) = "Marzo"
'    arrMes(4) = "Abril": arrMes(5) = "Mayo": arrMes(6) = "Junio"
'    arrMes(7) = "Julio": arrMes(8) = "Agosto": arrMes(9) = "Setiembre"
'    arrMes(10) = "Octubre": arrMes(11) = "Noviembre": arrMes(12) = "Diciembre"
'
'    '*** Determinar cantidad de periodos ***
'    intPeriodos = CInt(DateDiff("m", datFechaInicial, datFechaFinal) + 2)
'
'    blnPrimero = True
'    datFechaActualizada = datFechaInicial
'    datFechaPago = DateAdd("d", 1, datFechaActualizada)
'    If Not EsDiaUtil(datFechaPago) Then
'        datFechaPago = ProximoDiaUtil(datFechaPago)
'    End If
'    With adoComm
'
'        '*** Adicionar Periodo Contable ***
'        .CommandText = "INSERT INTO PeriodoContable (CodFondo,CodAdministradora,PeriodoContable,MesContable,DescripPeriodo,FechaInicio,FechaFinal,IndVigente) VALUES ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Format(Year(datFechaActualizada), "0000") & "','00','Apertura " & Format(Year(datFechaActualizada), "0000") & "','" & Convertyyyymmdd(datFechaInicial) & "','" & Convertyyyymmdd(datFechaInicial) & "','' )"
'        adoConn.Execute .CommandText
'
'        Do While DateDiff("d", datFechaActualizada, datFechaFinal) >= 0
'            If strpTipoAdministradora = Codigo_Tipo_Fondo_Administradora Then
'                .CommandText = "INSERT INTO AdministradoraCalendario (CodFondo,CodAdministradora,FechaContable,ValorTipoCambio,CodMoneda) " & _
'                    "VALUES ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Convertyyyymmdd(datFechaActualizada) & "',0,'" & strCodMoneda & "')"
'            Else
'                .CommandText = "INSERT INTO FondoValorCuota (CodFondo,CodAdministradora,FechaCuota,ValorCuotaInicial,ValorCuotaFinal,ValorCuotaInicialReal,ValorCuotaFinalReal,ValorTipoCambio,CodMoneda,CantCuotaInicio,CantCuotaSuscripcionConocida,CantCuotaRedencionConocida,CantCuotaFinal,CantCuotaSuscripcionDesconocida,CantCuotaRedencionDesconocida,CantParticipe,MontoPatrimonio,TasaAdministracion,MontoActivo) " & _
'                    "VALUES ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Convertyyyymmdd(datFechaActualizada) & "',0,0,0,0,0,'" & strCodMoneda & "',0,0,0,0,0,0,0,0,0,0 )"
'            End If
'            adoComm.Execute .CommandText
'
'            If Month(datFechaActualizada) <> Month(DateAdd("d", 1, datFechaActualizada)) Then
'                 If blnPrimero Then
'                    .CommandText = "INSERT INTO PeriodoContable (CodFondo,CodAdministradora,PeriodoContable,MesContable,DescripPeriodo,FechaInicio,FechaFinal,IndVigente,IndApertura) values ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Format(Year(datFechaActualizada), "0000") & "','" & Format(Month(datFechaActualizada), "00") & "','" & arrMes(Month(datFechaActualizada)) & " " & Format(Year(datFechaActualizada), "0000") & "','" & Convertyyyymmdd(datFechaInicial) & "','" & Convertyyyymmdd(datFechaActualizada) & "','','X')"
'                    If strpTipoAdministradora = Codigo_Tipo_Fondo_Administradora Then blnPrimero = False
'                 Else
'                    .CommandText = "INSERT INTO PeriodoContable (CodFondo,CodAdministradora,PeriodoContable,MesContable,DescripPeriodo,FechaInicio,FechaFinal,IndVigente) values ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Format(Year(datFechaActualizada), "0000") & "','" & Format(Month(datFechaActualizada), "00") & "','" & arrMes(Month(datFechaActualizada)) & " " & Format(Year(datFechaActualizada), "0000") & "','" & Convertyyyymmdd(DateAdd("m", -1, DateAdd("d", 1, datFechaActualizada))) & "','" & Convertyyyymmdd(datFechaActualizada) & "','')"
'                 End If
'                 adoConn.Execute adoComm.CommandText
'
''                 If strpTipoAdministradora <> Codigo_Tipo_Fondo_Administradora Then
''                    '*** Pagos a la Administradora ***
''                    If blnPrimero Then
''                       .CommandText = "INSERT INTO FondoPagoAdministradora (CodFondo,CodAdministradora,FechaCorte,FechaPago) values ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Convertyyyymmdd(datFechaActualizada) & "','" & Convertyyyymmdd(datFechaPago) & "')"
''                       blnPrimero = False
''                    Else
''                       .CommandText = "INSERT INTO FondoPagoAdministradora (CodFondo,CodAdministradora,FechaCorte,FechaPago) values ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Convertyyyymmdd(datFechaActualizada) & "','" & Convertyyyymmdd(datFechaPago) & "')"
''                    End If
''                    adoConn.Execute adoComm.CommandText
''                 End If
'
'            End If
'            datFechaActualizada = DateAdd("d", 1, datFechaActualizada)
'            datFechaPago = DateAdd("d", 1, datFechaActualizada)
'            If Not EsDiaUtil(datFechaPago) Then
'                datFechaPago = ProximoDiaUtil(datFechaPago)
'            End If
'            ctlControl.Panels(3).Text = "Creando Registros al -> " & CStr(datFechaActualizada) & "..."
'
'        Loop
'
'        datFechaActualizada = DateAdd("d", -1, datFechaActualizada)
'        .CommandText = "INSERT INTO PeriodoContable (CodFondo,CodAdministradora,PeriodoContable,MesContable,DescripPeriodo,FechaInicio,FechaFinal,IndVigente,IndApertura) " & _
'            "VALUES ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Format(Year(datFechaActualizada), "0000") & "','99','Cierre " & Format(Year(datFechaActualizada), "0000") & "','" & Convertyyyymmdd(datFechaInicial) & "','" & Convertyyyymmdd(datFechaFinal) & "','','X' )"
'        adoConn.Execute .CommandText
'    End With
'
'End Sub

Public Sub GenerarPeriodoComision(ByVal strpTipoAdministradora, ByVal strpCodAdministradora, ByVal strpCodFondo As String, ByVal datFechaInicial As Date, ByVal datFechaFinal As Date, ByVal strpCodComision As String, ByVal strpCodAnalitica, ByVal ctlControl As Control)

    Dim intPeriodos         As Integer, blnPrimero  As Boolean
    Dim datFechaActualizada As Date, datFechaPago   As Date
    Dim datFechaAnterior    As Date
    
         
    '*** Determinar cantidad de periodos ***
    intPeriodos = CInt(DateDiff("m", datFechaInicial, datFechaFinal))
       
    blnPrimero = True
    datFechaActualizada = datFechaInicial
    datFechaPago = DateAdd("d", 1, datFechaActualizada)
    If Not EsDiaUtil(datFechaPago) Then
        datFechaPago = ProximoDiaUtil(datFechaPago)
    End If
    With adoComm
        '*** Adicionar Periodo Comisión ***
        Do While DateDiff("d", datFechaActualizada, datFechaFinal) >= 0
            If Month(datFechaActualizada) <> Month(DateAdd("d", 1, datFechaActualizada)) Then
                 If strpTipoAdministradora <> Codigo_Tipo_Fondo_Administradora Then
                    '*** Pagos a la Administradora ***
                    If blnPrimero Then
                       .CommandText = "INSERT INTO FondoPagoAdministradora (CodFondo,CodAdministradora,CodComision,CodAnalitica,FechaInicio,FechaCorte,FechaPago,FechaLiquidacion) values ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & strpCodComision & "','" & strpCodAnalitica & "','" & Convertyyyymmdd(datFechaInicial) & "','" & Convertyyyymmdd(datFechaActualizada) & "','" & Convertyyyymmdd(datFechaPago) & "','" & Convertyyyymmdd(datFechaPago) & "')"
                       blnPrimero = False
                    Else
                       .CommandText = "INSERT INTO FondoPagoAdministradora (CodFondo,CodAdministradora,CodComision,CodAnalitica,FechaInicio,FechaCorte,FechaPago,FechaLiquidacion) values ('" & strpCodFondo & "','" & strpCodAdministradora & "','" & strpCodComision & "','" & strpCodAnalitica & "','" & Convertyyyymmdd(datFechaAnterior) & "','" & Convertyyyymmdd(datFechaActualizada) & "','" & Convertyyyymmdd(datFechaPago) & "','" & Convertyyyymmdd(datFechaPago) & "')"
                    End If
                    adoConn.Execute adoComm.CommandText
                 End If
                 datFechaAnterior = DateAdd("d", 1, datFechaActualizada)
            End If
            datFechaActualizada = DateAdd("d", 1, datFechaActualizada)
            datFechaPago = DateAdd("d", 1, datFechaActualizada)
            If Not EsDiaUtil(datFechaPago) Then
                datFechaPago = ProximoDiaUtil(datFechaPago)
            End If
            ctlControl.Panels(3).Text = "Creando Cortes al -> " & CStr(datFechaActualizada) & "..."
             
        Loop
        
        datFechaActualizada = DateAdd("d", -1, datFechaActualizada)
    End With
   
End Sub

Public Function Convertyyyymmdd(ByVal datFechaOrigen As Date) As String
    
    Convertyyyymmdd = Format(datFechaOrigen, gstrFormatoFechaInterno)
                
End Function

Public Function ObtenerSaldoFinalCuenta(ByVal strpCodFondo As String, ByVal strpCodAdministradora As String, ByVal strpCodFile As String, ByVal strpCodAnalitica As String, ByVal strpFechaConsulta As String, ByVal strpFechaSiguiente As String, ByVal strpCodCuenta As String, ByVal strpCodMoneda As String) As Currency

    Dim adoRegistro     As ADODB.Recordset
    
    ObtenerSaldoFinalCuenta = 0
        
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "{ call up_ACObtenerSaldoFinalCuenta('" & _
            strpCodFondo & "','" & strpCodAdministradora & "','" & strpCodFile & "','" & _
            strpCodAnalitica & "','" & strpFechaConsulta & "','" & strpFechaSiguiente & "','" & _
            strpCodCuenta & "','" & strpCodMoneda & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            ObtenerSaldoFinalCuenta = CCur(adoRegistro("SaldoCuenta"))
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With

End Function
Public Function ObtenerSaldoFinalContableCuenta(ByVal strpCodFondo As String, ByVal strpCodAdministradora As String, ByVal strpCodFile As String, ByVal strpCodAnalitica As String, ByVal strpFechaConsulta As String, ByVal strpFechaSiguiente As String, ByVal strpCodCuenta As String, ByVal strpCodMoneda As String) As Currency

    Dim adoRegistro     As ADODB.Recordset
    
    Dim ObtenerSaldoFinalCuenta  As Variant
    ObtenerSaldoFinalCuenta = 0
        
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "{ call up_ACObtenerSaldoFinalContableCuenta('" & _
            strpCodFondo & "','" & strpCodAdministradora & "','" & strpCodFile & "','" & _
            strpCodAnalitica & "','" & strpFechaConsulta & "','" & strpFechaSiguiente & "','" & _
            strpCodCuenta & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            ObtenerSaldoFinalCuenta = CCur(adoRegistro("SaldoCuenta"))
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With

End Function

Public Function ObtenerSignoMoneda(ByVal strpMoneda As String) As String

    Dim adoRegistro As ADODB.Recordset
    Dim strSigno    As String
    
    ObtenerSignoMoneda = Valor_Caracter
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT Signo FROM Moneda WHERE CodMoneda='" & strpMoneda & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strSigno = Trim(adoRegistro("Signo"))
        Else
            adoRegistro.Close: Set adoRegistro = Nothing
            Exit Function
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    ObtenerSignoMoneda = strSigno
    
End Function

Public Function ObtenerCodSignoMoneda(ByVal strpMoneda As String) As String

    Dim adoRegistro As ADODB.Recordset
    Dim strSigno    As String
    
    ObtenerCodSignoMoneda = Valor_Caracter
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT CodSigno FROM Moneda WHERE CodMoneda='" & strpMoneda & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strSigno = Trim(adoRegistro("CodSigno"))
        Else
            adoRegistro.Close: Set adoRegistro = Nothing
            Exit Function
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    ObtenerCodSignoMoneda = strSigno
    
End Function

Public Function ObtenerDescripcionMoneda(ByVal strpMoneda As String) As String

    Dim adoRegistro     As ADODB.Recordset
    Dim strDescripcion  As String
    
    ObtenerDescripcionMoneda = Valor_Caracter
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT DescripMoneda FROM Moneda WHERE CodMoneda='" & strpMoneda & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strDescripcion = Trim(adoRegistro("DescripMoneda"))
        Else
            adoRegistro.Close: Set adoRegistro = Nothing
            Exit Function
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    ObtenerDescripcionMoneda = strDescripcion
    
End Function

'/**/
'Public Function ObtenerDescripcionDeMonedaOrigenAMonedaPago(ByVal strpCodAnalitica As String) As String
'
'    Dim adoRegistro     As ADODB.Recordset
'    Dim strDescripcion  As String
'
'    ObtenerDescripcionDeMonedaOrigenAMonedaPago = Valor_Caracter
'
'    Set adoRegistro = New ADODB.Recordset
'    With adoComm
'        '.CommandText = "SELECT DescripMoneda FROM Moneda WHERE CodMoneda='" & strpMoneda & "'"
'        .CommandText = "SELECT i.CodAnalitica, i.CodMoneda,MO.CodSigno, i.CodMoneda1 ,M.CodSigno, DescripcionDeMonedaAMoneda = MO.CodSigno + '/' + M.CodSigno + ' '+ M.CodSigno + '/' + MO.CodSigno FROM InstrumentoInversion i inner join Moneda M on i.CodMoneda1=M.CodMoneda inner join Moneda MO on i.CodMoneda=MO.CodMoneda  WHERE i.CodAnalitica='" & strpCodAnalitica & "'"
'
'        Set adoRegistro = .Execute
'
'        If Not adoRegistro.EOF Then
'            strDescripcion = Trim(adoRegistro("DescripcionDeMonedaAMoneda"))
'        Else
'            adoRegistro.Close: Set adoRegistro = Nothing
'            Exit Function
'        End If
'        adoRegistro.Close: Set adoRegistro = Nothing
'    End With
'
'    ObtenerDescripcionDeMonedaOrigenAMonedaPago = strDescripcion
'
'End Function
'/**/


Public Function ObtenerCodigoParametro(ByVal strpCodTipoParametro As String, ByVal strpValorParametro As String) As String

    Dim adoRegistro     As ADODB.Recordset
    Dim strCODIGO  As String
    
    ObtenerCodigoParametro = Valor_Caracter
   
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandType = adCmdStoredProc
        .CommandText = "up_ACSelAuxiliarParametro"
        '.CommandText = "{call up_ACSelAuxiliarParametro ('" & strpCodTipoParametro & "','" & strpValorParametro & "','I') }"
        'Set adoRegistro = .Execute
            
        .Parameters.Append .CreateParameter("CodTipoParametro", adChar, adParamInput, 6, strpCodTipoParametro)
        .Parameters.Append .CreateParameter("ValorParametro", adChar, adParamInput, 8, strpValorParametro)
        .Parameters.Append .CreateParameter("TipoConsulta", adChar, adParamOutput, 1, "I")
        Set adoRegistro = .Execute
                                
        If Not adoRegistro.EOF Then
            strCODIGO = Trim(adoRegistro("CODIGO"))
        Else
            adoRegistro.Close: Set adoRegistro = Nothing
            Exit Function
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
                        
        .Parameters.Delete ("CodTipoParametro"): .Parameters.Delete ("ValorParametro")
        .Parameters.Delete ("TipoConsulta")
                        
        .CommandType = adCmdText
        
    End With
    
    ObtenerCodigoParametro = strCODIGO
    
End Function

Public Function ObtenerDescripcionParametro(ByVal strpCodTipoParametro As String, ByVal strpValorParametro As String) As String

    Dim adoRegistro     As ADODB.Recordset
    Dim strDescripcion  As String
    
    ObtenerDescripcionParametro = Valor_Caracter
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT DescripParametro FROM AuxiliarParametro " & _
            "WHERE CodTipoParametro='" & strpCodTipoParametro & "' AND CodParametro='" & strpValorParametro & "'"
            Set adoRegistro = .Execute
            
            If Not adoRegistro.EOF Then
                strDescripcion = Trim(adoRegistro("DescripParametro"))
            Else
                adoRegistro.Close: Set adoRegistro = Nothing
                Exit Function
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    ObtenerDescripcionParametro = strDescripcion
    
End Function
Public Function ObtenerNumMaximoDocumentoIdentidad(ByVal strpCodTipoDocumento As String) As Integer

    Dim adoRegistro         As ADODB.Recordset
    Dim intNumValidacion    As Integer
    
    ObtenerNumMaximoDocumentoIdentidad = 0
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT NumValidacion FROM AuxiliarParametro WHERE CodTipoParametro='TIPIDE' AND CodParametro='" & strpCodTipoDocumento & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            intNumValidacion = CInt(adoRegistro("NumValidacion"))
        Else
            adoRegistro.Close: Set adoRegistro = Nothing
            Exit Function
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    ObtenerNumMaximoDocumentoIdentidad = intNumValidacion
    
End Function
Public Function ValidarAnalitica(ByVal strFile As String, ByVal strAnalitica As String, ByVal strFondo As String) As Boolean

    Dim adoRegistro As ADODB.Recordset
    Dim adoConsulta As ADODB.Recordset
    
    ValidarAnalitica = False
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT CodFile FROM InversionFile WHERE CodFile='" & strFile & "' AND IndVigente='X' AND CodFile<>'000'"
        
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            Set adoConsulta = New ADODB.Recordset
            
            Select Case strFile
                Case "001", "002", "003"
                    .CommandText = "SELECT CodAnalitica FROM BancoCuenta " & _
                        "WHERE CodAnalitica='" & strAnalitica & "' AND CodFile='" & strFile & "' AND " & _
                        "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strFondo & "' AND " & _
                        "IndVigente='X'"
                                
                Case "098", "099"
                    .CommandText = "SELECT CodDetalleFile FROM InversionDetalleFile " & _
                        "WHERE CodFile='" & strFile & "' AND IndVigente='X'"
                
                Case Else
                    .CommandText = "SELECT CodAnalitica FROM InstrumentoInversion " & _
                        "WHERE CodAnalitica='" & strAnalitica & "' AND CodFile='" & strFile & "' AND " & _
                        "IndVigente='X'"
            
            End Select
                                    
            Set adoConsulta = .Execute
            
            If adoConsulta.EOF Then
                Exit Function
            End If
            adoConsulta.Close: Set adoConsulta = Nothing
            
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    ValidarAnalitica = True
    
End Function

Public Function Convertddmmyyyy(ByVal strFechaOrigen As String) As Date

    If Trim(strFechaOrigen) = Valor_Caracter Or strFechaOrigen = Null Then
        Convertddmmyyyy = Valor_Caracter
    Else
        strFechaOrigen = Mid(strFechaOrigen, 1, 4) & "/" & Mid(strFechaOrigen, 5, 2) & "/" & Mid(strFechaOrigen, 7, 2)
        Convertddmmyyyy = Format(strFechaOrigen, gstrFormatoFechaCliente)
    End If
                
End Function
'Sub NEWTIRPROM_A(codfond As String, Codfile As String, CodAnal As String, TipVac As String, FECHA As String)
'
'    Dim SenSql As String, adoKar As New Recordset, adoAux As New Recordset
'    Dim n_ValCosto As Currency, n_ValProv As Currency, n_ValInt As Currency, n_ValGP As Currency, n_ValVacP As Currency, n_ValReca As Currency, n_ValIntC As Currency, n_ValMerc As Currency
'    Dim CtaInve As String, CtaFluc As String, CtaIade As String, CtaFlca As String, CtaVacPeriod As String, CtaReca As String, CtaFlme As String
'    Dim CtaVacc As String, CtaPRCVacP As String, CtaIadeC As String, Tip_bono As String
'    Dim nTIRProm, n_Nominal As String * 15, n_ACtual As String * 15, n_TirProm As String * 12
'    Dim d_Fecha
'
'    '*** CD's /Dep. Plazo en Cartera ***
'    SenSql = "SELECT COD_FOND,TIP_MOVI,SLD_FINA,COD_MONE,FCH_MOVI,NRO_KARD,TIR_OPER,SLD_AMORT FROM FMKARDEX "
'    SenSql = SenSql + "WHERE COD_FILE='" + Codfile + "'"
'    SenSql = SenSql + " AND COD_ANAL='" + CodAnal + "'"
'    SenSql = SenSql + " AND COD_FOND='" + codfond + "'"
'    SenSql = SenSql + " AND SLD_FINA > 0"
'    SenSql = SenSql + " AND FLG_ULTI='X'"
'    SenSql = SenSql + " ORDER BY COD_ANAL"
'    adoComm.CommandText = SenSql
'    Set adoKar = adoComm.Execute
'    If Not adoKar.EOF Then
'        If adoKar!TIP_MOVI = "E" Or adoKar!TIP_MOVI = "S" Then
'            adoComm.CommandText = "SELECT TIP_CTA,COD_CTAN,COD_CTAX FROM FMTITCTA WHERE CLS_TITU='" + IIf(Codfile = "03", "I", "R") + "'"
'            Set adoAux = adoComm.Execute
'            Do While Not adoAux.EOF
'                Select Case adoAux!TIP_CTA
'                    Case "A" '*** Cta. de Inversiones ***
'                        CtaInve = adoAux!cod_ctan
'                    Case "B" '*** Prov. Intereses ***
'                        CtaFluc = adoAux!cod_ctan
'                    Case "G" '*** Vac Corrido Periódico ***
'                        CtaVacPeriod = adoAux!cod_ctan
'                    Case "J" '*** Interés Corrido Compra ***
'                        CtaIade = adoAux!cod_ctan
'                    Case "K" '*** Prov. Reajuste Capital Vac Vcto. ***
'                        CtaReca = adoAux!cod_ctan
'                    Case "M" '*** Prov. Fluctuación Mercado ***
'                        CtaFlme = adoAux!cod_ctan
'                    Case "O" '*** Vac Corrido Vcto. ***
'                        CtaVacc = adoAux!cod_ctan
'                    Case "R" '*** Prov. Reajuste Capital Vac Period. ***
'                        CtaPRCVacP = adoAux!cod_ctan
'                    Case "T" '*** Interés Corrido Cap. Reajustado ***
'                        CtaIadeC = adoAux!cod_ctan
'                    Case "X" '*** Prov. Fluctuación Capital V.N. ***
'                        CtaFlca = adoAux!cod_ctan
'                End Select
'                adoAux.MoveNext
'            Loop
'            adoAux.Close: Set adoAux = Nothing
'
'        Do While Not adoKar.EOF
'            n_ValCosto = 0: n_ValProv = 0: n_ValInt = 0: n_ValGP = 0: n_ValVacP = 0: n_ValReca = 0: n_ValIntC = 0: n_ValMerc = 0
'
'            If adoKar!cod_mone = "D" Then
'                SenSql = "SELECT (SIN_MONX + MDE_MONX + MHA_MONX) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "'"
'            Else
'                SenSql = "SELECT (SIN_MONN + MDE_MONN + MHA_MONN) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "' "
'            End If
'            SenSql = SenSql + " AND COD_CTA='" + CtaInve + "'"
'            SenSql = SenSql + " AND COD_FILE='" + Codfile + "' AND COD_ANAL='" + CodAnal + "'"
'            SenSql = SenSql + " AND FCH_SALD='" + FECHA + "'"
'            adoComm.CommandText = SenSql
'            Set adoAux = adoComm.Execute
'            If adoAux.EOF Then
'                n_ValCosto = 0
'            Else
'                n_ValCosto = adoAux!Monto
'            End If
'            adoAux.Close: Set adoAux = Nothing
'
'            If adoKar!cod_mone = "D" Then
'                SenSql = "SELECT (SIN_MONX + MDE_MONX + MHA_MONX) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "'"
'            Else
'                SenSql = "SELECT (SIN_MONN + MDE_MONN + MHA_MONN) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "' "
'            End If
'            SenSql = SenSql + " AND COD_CTA='" + CtaFluc + "'"
'            SenSql = SenSql + " AND COD_FILE='" + Codfile + "' AND COD_ANAL='" + CodAnal + "'"
'            SenSql = SenSql + " AND FCH_SALD='" + FECHA + "'"
'            adoComm.CommandText = SenSql
'            Set adoAux = adoComm.Execute
'            If adoAux.EOF Then
'                n_ValProv = 0
'            Else
'                n_ValProv = adoAux!Monto
'            End If
'            adoAux.Close: Set adoAux = Nothing
'
'            If adoKar!cod_mone = "D" Then
'                SenSql = "SELECT (SIN_MONX + MDE_MONX + MHA_MONX) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "'"
'            Else
'                SenSql = "SELECT (SIN_MONN + MDE_MONN + MHA_MONN) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "' "
'            End If
'            SenSql = SenSql + " AND COD_CTA='" + CtaIade + "'"
'            SenSql = SenSql + " AND COD_FILE='" + Codfile + "' AND COD_ANAL='" + CodAnal + "'"
'            SenSql = SenSql + " AND FCH_SALD='" + FECHA + "'"
'            adoComm.CommandText = SenSql
'            Set adoAux = adoComm.Execute
'            If adoAux.EOF Then
'                n_ValInt = 0
'            Else
'                n_ValInt = adoAux!Monto
'            End If
'            adoAux.Close: Set adoAux = Nothing
'
'            If adoKar!cod_mone = "D" Then
'                SenSql = "SELECT (SIN_MONX + MDE_MONX + MHA_MONX) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "'"
'            Else
'                SenSql = "SELECT (SIN_MONN + MDE_MONN + MHA_MONN) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "' "
'            End If
'            SenSql = SenSql + " AND COD_CTA='" + CtaFlca + "'"
'            SenSql = SenSql + " AND COD_FILE='" + Codfile + "' AND COD_ANAL='" + CodAnal + "'"
'            SenSql = SenSql + " AND FCH_SALD='" + FECHA + "'"
'            adoComm.CommandText = SenSql
'            Set adoAux = adoComm.Execute
'            If adoAux.EOF Then
'                n_ValGP = 0
'            Else
'                n_ValGP = adoAux!Monto
'            End If
'            adoAux.Close: Set adoAux = Nothing
'
'            If adoKar!cod_mone = "D" Then
'                SenSql = "SELECT (SIN_MONX + MDE_MONX + MHA_MONX) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "'"
'            Else
'                SenSql = "SELECT (SIN_MONN + MDE_MONN + MHA_MONN) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "' "
'            End If
'            SenSql = SenSql + " AND COD_CTA='" + CtaFlme + "'"
'            SenSql = SenSql + " AND COD_FILE='" + Codfile + "' AND COD_ANAL='" + CodAnal + "'"
'            SenSql = SenSql + " AND FCH_SALD='" + FECHA + "'"
'            adoComm.CommandText = SenSql
'            Set adoAux = adoComm.Execute
'            If adoAux.EOF Then
'                n_ValMerc = 0
'            Else
'                n_ValMerc = adoAux!Monto
'            End If
'            adoAux.Close: Set adoAux = Nothing
'
'            '*** Hallar TIR Promedio ***
'            d_Fecha = CVDate(Right$(FECHA, 2) + "/" + Mid$(FECHA, 5, 2) + "/" + Left$(FECHA, 4))
'            nTIRProm = TirNoPerCrtDepC(Trim$(adoKar!COD_FOND), Trim$(Codfile), Trim$(CodAnal), d_Fecha, DateAdd("d", 1, d_Fecha), n_ValCosto + n_ValProv + n_ValInt + n_ValGP + n_ValMerc, 0, CDbl(adoKar!SLD_FINA), CDbl(adoKar!SLD_AMORT), adoKar!TIR_OPER / 100, Trim$(TipVac))
'
'            SenSql = "UPDATE FMKARDEX SET TIR_PROM=" & Format(nTIRProm, "0.000000")
'            SenSql = SenSql + " WHERE COD_FILE='" + Codfile + "' AND COD_ANAL='" + CodAnal + "' AND COD_FOND = '" + codfond + "' "
'            SenSql = SenSql + " AND NRO_KARD=" + adoKar!NRO_KARD
'            adoConn.Execute SenSql
'            adoKar.MoveNext
'        Loop
'        End If
'    End If
'    adoKar.Close: Set adoKar = Nothing
'
'End Sub

Sub NEWTIRPROM(Codfile As String, CodAnal As String, TipBono As String, TipVac As String, FECHA As String)

'    Dim SenSql As String, adoKar As New Recordset, adoAux As New Recordset
'    Dim n_ValCosto As Currency, n_ValProv As Currency, n_ValInt As Currency, n_ValGP As Currency, n_ValVacP As Currency, n_ValReca As Currency, n_ValIntC As Currency, n_ValMerc As Currency
'    Dim CtaInve As String, CtaFluc As String, CtaIade As String, CtaFlca As String, CtaVacPeriod As String, CtaReca As String, CtaFlme As String
'    Dim CtaVacc As String, CtaPRCVacP As String, CtaIadeC As String, Tip_bono As String
'    Dim nTIRProm, n_Nominal As String * 15, n_ACtual As String * 15, n_TirProm As String * 12
'    Dim d_Fecha
'
'    '*** Bonos en Cartera ***
'    SenSql = "SELECT COD_FOND,TIP_MOVI,SLD_FINA,COD_MONE,FCH_MOVI,NRO_KARD,TIR_OPER,SLD_AMORT FROM FMKARDEX "
'    SenSql = SenSql + "WHERE COD_FILE='" + Codfile + "'"
'    SenSql = SenSql + " AND COD_ANAL='" + CodAnal + "'"
'    SenSql = SenSql + " AND SLD_FINA > 0"
'    SenSql = SenSql + " AND FLG_ULTI='X'"
'    SenSql = SenSql + " ORDER BY COD_ANAL"
'    adoComm.CommandText = SenSql
'    Set adoKar = adoComm.Execute
'    If Not adoKar.EOF Then
'        If adoKar!TIP_MOVI = "E" Or adoKar!TIP_MOVI = "S" Then
'            adoComm.CommandText = "SELECT TIP_CTA,COD_CTAN,COD_CTAX FROM FMTITCTA WHERE CLS_TITU='" + TipBono + "'"
'            Set adoAux = adoComm.Execute
'            Do While Not adoAux.EOF
'                Select Case adoAux!TIP_CTA
'                    Case "A" '*** Cta. de Inversiones ***
'                        CtaInve = adoAux!cod_ctan
'                    Case "B" '*** Prov. Intereses ***
'                        CtaFluc = adoAux!cod_ctan
'                    Case "G" '*** Vac Corrido Periódico ***
'                        CtaVacPeriod = adoAux!cod_ctan
'                    Case "J" '*** Interés Corrido Compra ***
'                        CtaIade = adoAux!cod_ctan
'                    Case "K" '*** Prov. Reajuste Capital Vac Vcto. ***
'                        CtaReca = adoAux!cod_ctan
'                    Case "M" '*** Prov. Fluctuación Mercado ***
'                        CtaFlme = adoAux!cod_ctan
'                    Case "O" '*** Vac Corrido Vcto. ***
'                        CtaVacc = adoAux!cod_ctan
'                    Case "R" '*** Prov. Reajuste Capital Vac Period. ***
'                        CtaPRCVacP = adoAux!cod_ctan
'                    Case "T" '*** Interés Corrido Cap. Reajustado ***
'                        CtaIadeC = adoAux!cod_ctan
'                    Case "X" '*** Prov. Fluctuación Capital V.N. ***
'                        CtaFlca = adoAux!cod_ctan
'                End Select
'                adoAux.MoveNext
'            Loop
'            adoAux.Close: Set adoAux = Nothing
'
'        Do While Not adoKar.EOF
'            n_ValCosto = 0: n_ValProv = 0: n_ValInt = 0: n_ValGP = 0: n_ValVacP = 0: n_ValReca = 0: n_ValIntC = 0: n_ValMerc = 0
'
'            If adoKar!COD_MONE = "D" Then
'                SenSql = "SELECT (SIN_MONX + MDE_MONX + MHA_MONX) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "'"
'            Else
'                SenSql = "SELECT (SIN_MONN + MDE_MONN + MHA_MONN) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "' "
'            End If
'            SenSql = SenSql + " AND COD_CTA='" + CtaInve + "'"
'            SenSql = SenSql + " AND COD_FILE='" + Codfile + "' AND COD_ANAL='" + CodAnal + "'"
'            SenSql = SenSql + " AND FCH_SALD='" + FECHA + "'"
'            adoComm.CommandText = SenSql
'            Set adoAux = adoComm.Execute
'            If adoAux.EOF Then
'                n_ValCosto = 0
'            Else
'                n_ValCosto = adoAux!Monto
'            End If
'            adoAux.Close: Set adoAux = Nothing
'
'            If adoKar!COD_MONE = "D" Then
'                SenSql = "SELECT (SIN_MONX + MDE_MONX + MHA_MONX) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "'"
'            Else
'                SenSql = "SELECT (SIN_MONN + MDE_MONN + MHA_MONN) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "' "
'            End If
'            SenSql = SenSql + " AND COD_CTA='" + CtaFluc + "'"
'            SenSql = SenSql + " AND COD_FILE='" + Codfile + "' AND COD_ANAL='" + CodAnal + "'"
'            SenSql = SenSql + " AND FCH_SALD='" + FECHA + "'"
'            adoComm.CommandText = SenSql
'            Set adoAux = adoComm.Execute
'            If adoAux.EOF Then
'                n_ValProv = 0
'            Else
'                n_ValProv = adoAux!Monto
'            End If
'            adoAux.Close: Set adoAux = Nothing
'
'            If adoKar!COD_MONE = "D" Then
'                SenSql = "SELECT (SIN_MONX + MDE_MONX + MHA_MONX) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "'"
'            Else
'                SenSql = "SELECT (SIN_MONN + MDE_MONN + MHA_MONN) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "' "
'            End If
'            SenSql = SenSql + " AND COD_CTA='" + CtaIade + "'"
'            SenSql = SenSql + " AND COD_FILE='" + Codfile + "' AND COD_ANAL='" + CodAnal + "'"
'            SenSql = SenSql + " AND FCH_SALD='" + FECHA + "'"
'            adoComm.CommandText = SenSql
'            Set adoAux = adoComm.Execute
'            If adoAux.EOF Then
'                n_ValInt = 0
'            Else
'                n_ValInt = adoAux!Monto
'            End If
'            adoAux.Close: Set adoAux = Nothing
'
'            If adoKar!COD_MONE = "D" Then
'                SenSql = "SELECT (SIN_MONX + MDE_MONX + MHA_MONX) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "'"
'            Else
'                SenSql = "SELECT (SIN_MONN + MDE_MONN + MHA_MONN) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "' "
'            End If
'            SenSql = SenSql + " AND COD_CTA='" + CtaFlca + "'"
'            SenSql = SenSql + " AND COD_FILE='" + Codfile + "' AND COD_ANAL='" + CodAnal + "'"
'            SenSql = SenSql + " AND FCH_SALD='" + FECHA + "'"
'            adoComm.CommandText = SenSql
'            Set adoAux = adoComm.Execute
'            If adoAux.EOF Then
'                n_ValGP = 0
'            Else
'                n_ValGP = adoAux!Monto
'            End If
'            adoAux.Close: Set adoAux = Nothing
'
'            If adoKar!COD_MONE = "D" Then
'                SenSql = "SELECT (SIN_MONX + MDE_MONX + MHA_MONX) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "'"
'            Else
'                SenSql = "SELECT (SIN_MONN + MDE_MONN + MHA_MONN) MONTO FROM FMSALDOS WHERE COD_FOND='" + adoKar!COD_FOND + "' "
'            End If
'            SenSql = SenSql + " AND COD_CTA='" + CtaFlme + "'"
'            SenSql = SenSql + " AND COD_FILE='" + Codfile + "' AND COD_ANAL='" + CodAnal + "'"
'            SenSql = SenSql + " AND FCH_SALD='" + FECHA + "'"
'            adoComm.CommandText = SenSql
'            Set adoAux = adoComm.Execute
'            If adoAux.EOF Then
'                n_ValMerc = 0
'            Else
'                n_ValMerc = adoAux!Monto
'            End If
'            adoAux.Close: Set adoAux = Nothing
'
'            '*** Hallar TIR Promedio ***
'            d_Fecha = CVDate(Right$(FECHA, 2) + "/" + Mid$(FECHA, 5, 2) + "/" + Left$(FECHA, 4))
'            If Codfile = "06" Or Codfile = "10" Then
'                nTIRProm = TirNoPerLetPag(Trim$(Codfile), Trim$(CodAnal), Trim$(adoKar!COD_FOND), d_Fecha, DateAdd("d", 1, d_Fecha), n_ValCosto + n_ValProv + n_ValInt + n_ValGP + n_ValMerc, 0, CDbl(adoKar!SLD_FINA), CDbl(adoKar!SLD_AMORT), adoKar!TIR_OPER / 100, Trim$(TipVac))
'            Else
'                nTIRProm = TirNoPer(Trim$(Codfile), Trim$(CodAnal), d_Fecha, DateAdd("d", 1, d_Fecha), n_ValCosto + n_ValProv + n_ValInt + n_ValGP + n_ValMerc, 0, CDbl(adoKar!SLD_FINA), CDbl(adoKar!SLD_AMORT), adoKar!TIR_OPER / 100, Trim$(TipVac))
'            End If
'
'            SenSql = "UPDATE FMKARDEX SET TIR_PROM=" & Format(nTIRProm, "0.000000")
'            SenSql = SenSql + " WHERE COD_FILE='" + Codfile + "' AND COD_ANAL='" + CodAnal + "'"
'            SenSql = SenSql + " AND NRO_KARD=" + Str(adoKar!NRO_KARD)
'            adoConn.Execute SenSql
'            adoKar.MoveNext
'        Loop
'        End If
'    End If
'    adoKar.Close: Set adoKar = Nothing
    
End Sub

Public Sub OcultarReportes()

    With frmMainMdi.tlbMdi.Buttons("Reportes")
        .ButtonMenus("Repo1").Visible = False
        .ButtonMenus("Repo2").Visible = False
        .ButtonMenus("Repo3").Visible = False
        .ButtonMenus("Repo4").Visible = False
        .ButtonMenus("Repo5").Visible = False
        .ButtonMenus("Repo6").Visible = False
        .ButtonMenus("Repo7").Visible = False
        .ButtonMenus("Repo8").Visible = False
        .ButtonMenus("Repo9").Visible = False
        .ButtonMenus("Repo10").Visible = False
    End With
    
End Sub

Function ValCobPag_Tc(ByVal Tc$, ByVal fech$) As Integer

    Dim adoTcmb As New Recordset
    Dim NewTcMas
    Dim NewTcMen

    adoComm.CommandText = "SELECT VAL_TCMB FROM FMCUOTAS WHERE FCH_CUOT = '" + fech + "' AND "
    adoComm.CommandText = adoComm.CommandText + " COD_FOND <> 'AD'"
    Set adoTcmb = adoComm.Execute
    If adoTcmb.EOF Then
        Exit Function
    End If
    NewTcMas = adoTcmb!VAL_TCMB + (adoTcmb!VAL_TCMB * 0.03)
    NewTcMen = adoTcmb!VAL_TCMB - (adoTcmb!VAL_TCMB * 0.03)
    If (Val(Tc$) >= NewTcMen) And (Val(Tc$) <= NewTcMas) Then
        ValCobPag_Tc = 0  ' T.C. Correcto
    Else
        ValCobPag_Tc = 1  ' T.C. Incorrecto
    End If
    adoTcmb.Close: Set adoTcmb = Nothing
    
End Function

Function JoinLogical(a As String, B As String, lOper As String)

    Dim cTmp As String
    
    cTmp = a
    If a <> "" Then
       If B <> "" Then cTmp = a & lOper & B
    Else
       If B <> "" Then cTmp = B
    End If
    JoinLogical = cTmp
    
End Function

Public Function ValidarCuentaContable(ByVal strpCuenta As String, ByVal strpCodAdministradora) As Boolean

    Dim adoRegistro As ADODB.Recordset
    
    ValidarCuentaContable = False
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT CodCuenta FROM PlanContable WHERE CodCuenta='" & strpCuenta & "' AND CodAdministradora='" & strpCodAdministradora & "' AND IndMovimiento='X'"
        
        Set adoRegistro = .Execute
        
        If adoRegistro.EOF Then
            adoRegistro.Close: Set adoRegistro = Nothing
            Exit Function
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    ValidarCuentaContable = True
    
End Function

Public Function ValidarFechaInicial(ByVal datpFechaConsulta As Date, ByVal strpCodFondo As String, ByVal strpCodAdministradora As String) As Boolean

    '*** Valida la fecha de consulta para que no sea menor a la fecha de inicio del fondo ***
    Dim adoRegistro As ADODB.Recordset
    
    ValidarFechaInicial = False
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT FechaInicioEtapaPreOperativa FROM Fondo " & _
            "WHERE CodAdministradora='" & strpCodAdministradora & "' AND " & _
            "CodFondo='" & strpCodFondo & "'"
        
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            If datpFechaConsulta < adoRegistro("FechaInicioEtapaPreOperativa") Then
                adoRegistro.Close: Set adoRegistro = Nothing
                Exit Function
            End If
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    ValidarFechaInicial = True
    
End Function
Public Function ValidarFile(ByVal strFile As String) As Boolean

    Dim adoRegistro As ADODB.Recordset
    
    ValidarFile = False
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT CodFile FROM InversionFile WHERE CodFile='" & strFile & "' AND IndVigente='X'"
        
        Set adoRegistro = .Execute
        
        If adoRegistro.EOF Then
            adoRegistro.Close: Set adoRegistro = Nothing
            Exit Function
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    ValidarFile = True

End Function

Function WriteINIString(ByVal szItem As String, ByVal szGrup As String, ByVal szDefault As String, ByVal FileName As String)

    Dim tmp As String
    Dim X As Integer

    tmp = String(2048, 32)
    X = WritePrivateProfileString(szGrup, szItem, szDefault, FileName)
          
    WriteINIString = Mid(tmp, 1, X)

End Function

Sub DefCtas(xClsPers As String, xFlgExtr As String, xCodMone As String, s_CtaApo As String, s_CtaSpa As String, s_CtaBpa As String)
    '--------------------------------------------------
    'Parámetros:
    '-xClsPers$  : Clase de Personeria del Partícipe
    '-xFlgExtr$  : Flag de Extrangería del Partícipe
    'Permite Obtener las siguientes cuentas:
    '-s_CtaApo    : Cuenta Aporte
    '-s_CtaBpa    : Cuenta Bajo  la Par
    '-s_CtaSpa    : Cuenta Sobre la Par
    '--------------------------------------------------
    
    Dim adoDefCtas As New ADODB.Recordset
    Dim aCapital() As String
    
    ReDim aCapital(3)
    If xClsPers = "N" Then 'Natural
       If xFlgExtr <> "X" Then  'Nacional
          aCapital(1) = "TIP_DEFI='C' AND COD_DEFI='006'"
          aCapital(2) = "TIP_DEFI='C' AND COD_DEFI='010'"
          aCapital(3) = "TIP_DEFI='C' AND COD_DEFI='014'"
       Else                     'Extranjera
          aCapital(1) = "TIP_DEFI='C' AND COD_DEFI='007'"
          aCapital(2) = "TIP_DEFI='C' AND COD_DEFI='011'"
          aCapital(3) = "TIP_DEFI='C' AND COD_DEFI='015'"
       End If
    
    Else  ' Juridicas
       If xFlgExtr <> "X" Then   'Nacional
          aCapital(1) = "TIP_DEFI='C' AND COD_DEFI='008'"
          aCapital(2) = "TIP_DEFI='C' AND COD_DEFI='012'"
          aCapital(3) = "TIP_DEFI='C' AND COD_DEFI='016'"
       Else                       'Extranjera
          aCapital(1) = "TIP_DEFI='C' AND COD_DEFI='009'"
          aCapital(2) = "TIP_DEFI='C' AND COD_DEFI='013'"
          aCapital(3) = "TIP_DEFI='C' AND COD_DEFI='017'"
       End If
    End If
    
    'Cuenta de Aporte MN/ME
    adoComm.CommandText = "select COD_CTAN, COD_CTAX from FMCTADEF where " & aCapital(1)
    adoDefCtas.Open adoComm.CommandText, adoConn, adOpenStatic
    'Set adoDefCtas = adoComm.Execute
    If xCodMone = "S" Then
       s_CtaApo = IIf(Not IsNull(adoDefCtas!cod_ctan), adoDefCtas!cod_ctan, "")
    Else
       s_CtaApo = IIf(Not IsNull(adoDefCtas!cod_ctax), adoDefCtas!cod_ctax, "")
    End If
    adoDefCtas.Close: Set adoDefCtas = Nothing
    
    'Cuenta Bajo la Par MN/ME
    adoComm.CommandText = "select COD_CTAN, COD_CTAX from FMCTADEF where " & aCapital(2)
    adoDefCtas.Open adoComm.CommandText, adoConn, adOpenStatic
    'Set adoDefCtas = adoComm.Execute
    If xCodMone = "S" Then
       s_CtaBpa = IIf(Not IsNull(adoDefCtas!cod_ctan), adoDefCtas!cod_ctan, "")
    Else
       s_CtaBpa = IIf(Not IsNull(adoDefCtas!cod_ctax), adoDefCtas!cod_ctax, "")
    End If
    adoDefCtas.Close: Set adoDefCtas = Nothing
    
    'Cuenta Sobre la Par MN/ME
    adoComm.CommandText = "select COD_CTAN, COD_CTAX from FMCTADEF where " & aCapital(3)
    adoDefCtas.Open adoComm.CommandText, adoConn, adOpenStatic
    'Set adoDefCtas = adoComm.Execute
    If xCodMone = "S" Then
       s_CtaSpa = IIf(Not IsNull(adoDefCtas!cod_ctan), adoDefCtas!cod_ctan, "")
    Else
       s_CtaSpa = IIf(Not IsNull(adoDefCtas!cod_ctax), adoDefCtas!cod_ctax, "")
    End If
    adoDefCtas.Close: Set adoDefCtas = Nothing
    
End Sub

Public Function EsDiaUtil(ByVal datFecha As Date) As Boolean

    Dim blnValorRetorno As Boolean
    Dim intContador     As Integer
    
    blnValorRetorno = True
    If Weekday(datFecha) = 1 Or Weekday(datFecha) = 7 Then
        blnValorRetorno = False
    Else
        For intContador = 1 To UBound(gvntDiasNUtil)
            If DateDiff("d", datFecha, gvntDiasNUtil(intContador)) = 0 Then
                blnValorRetorno = False
                Exit For
            End If
        Next
    End If
    EsDiaUtil = True 'blnValorRetorno
    
End Function
Public Function DesplazamientoDiaUtil(ByVal datFecha As Date, strCodTipoDesplazamiento As String) As Date

    Dim strFecha    As String
    Dim adoConsulta As ADODB.Recordset
    
    With adoComm
        Set adoConsulta = New ADODB.Recordset
        
        strFecha = Convertyyyymmdd(datFecha)

        '*** Obtener el número de días del peridodo de pago ***
        .CommandText = "SELECT dbo.uf_ACObtenerFechaUtil('" & strFecha & "','" & strCodTipoDesplazamiento & "') as FechaUtil"

        Set adoConsulta = .Execute

        If Not adoConsulta.EOF Then
            DesplazamientoDiaUtil = adoConsulta("FechaUtil") '*** Días del periodo  ***
        End If
        adoConsulta.Close: Set adoConsulta = Nothing
    
    End With
    
   
End Function

Sub LDiasNUtil()
    
    Dim intCont As Integer, n_TotRec As Integer, res As Integer
    Dim adoresultTmp As ADODB.Recordset
    
    Set adoresultTmp = New ADODB.Recordset
    
    adoresultTmp.Open "SELECT FechaFeriado FROM CalendarioNoLaborable", adoConn, adOpenStatic
    
    If adoresultTmp.EOF Then
        ReDim gvntDiasNUtil(0)
        adoresultTmp.Close: Set adoresultTmp = Nothing
        Exit Sub
    End If
    adoresultTmp.MoveLast
    
    n_TotRec = adoresultTmp.RecordCount
    ReDim gvntDiasNUtil(adoresultTmp.RecordCount)
    adoresultTmp.MoveFirst
    
    For intCont = 1 To n_TotRec
        gvntDiasNUtil(intCont) = adoresultTmp("FechaFeriado")
        adoresultTmp.MoveNext
    Next
    adoresultTmp.Close: Set adoresultTmp = Nothing
   
End Sub
Public Sub LCmbLoad(strSQL As String, CtlNam As Control, AmapCmb(), vMenIni As String)

    Dim nCon As Long, objRS As New ADODB.Recordset
    
'    objRS.Open strSQL, adoComm, adOpenKeyset
    
    adoComm.CommandText = strSQL
    Set objRS = adoComm.Execute

    CtlNam.Clear
    nCon = 0
    ReDim AmapCmb(nCon)
    If vMenIni = "" Then
'        CtlNam.AddItem "{Todos}"
    Else
        CtlNam.AddItem vMenIni
        nCon = 0
        'ReDim Preserve AmapCmb(nCon)
        'AmapCmb(nCon) = ""
        'nCon = 1
    End If
     
    Do Until objRS.EOF
        CtlNam.AddItem objRS!DESCRIP
        'CtlNam.AddItem adoCombo.Fields(0)
        ReDim Preserve AmapCmb(nCon)
        AmapCmb(nCon) = objRS!codigo
        'AmapCmb(nCon) = adoCombo.Fields(1)
        objRS.MoveNext
        nCon = nCon + 1
    Loop
   
    objRS.Close: Set objRS = Nothing

End Sub

Public Function ProximoDiaUtil(ByVal datFecha As Date) As Date

    Dim datFechaSiguiente   As Date
    Dim blnIndicador        As Boolean, intContador As Integer
   
    datFechaSiguiente = DateAdd("d", 1, datFecha)
    
    '*** Verificar si es Sábado o Domingo ***
    If Weekday(datFechaSiguiente) = 1 Or Weekday(datFechaSiguiente) = 7 Then
        datFechaSiguiente = DateAdd("d", IIf(Weekday(datFechaSiguiente) = 1, 1, 2), datFechaSiguiente)
    End If
    blnIndicador = True
    Do While blnIndicador
        blnIndicador = False
        For intContador = 1 To UBound(gvntDiasNUtil)
            If DateDiff("d", datFechaSiguiente, gvntDiasNUtil(intContador)) = 0 Then
                blnIndicador = True
                datFechaSiguiente = DateAdd("d", 1, datFechaSiguiente)
                '*** Verificar si es Sábado o Domingo ***
                If Weekday(datFechaSiguiente) = 1 Or Weekday(datFechaSiguiente) = 7 Then
                    datFechaSiguiente = DateAdd("d", IIf(Weekday(datFechaSiguiente) = 1, 1, 2), datFechaSiguiente)
                End If
            End If
        Next
    Loop
    ProximoDiaUtil = datFechaSiguiente
   
End Function

Function AnteriorDiaUtil(ByVal datFecha As Date) As Date

    Dim datFechaAnterior    As Date
    Dim blnIndicador        As Boolean, intContador As Integer
   
    datFechaAnterior = DateAdd("d", -1, datFecha)
    
    '*** Verificar si es Sábado o Domingo ***
    If Weekday(datFechaAnterior) = 1 Or Weekday(datFechaAnterior) = 7 Then
        datFechaAnterior = DateAdd("d", IIf(Weekday(datFechaAnterior) = 1, -2, -1), datFechaAnterior)
    End If
    blnIndicador = True
    Do While blnIndicador
        blnIndicador = False
        For intContador = 1 To UBound(gvntDiasNUtil)
            If DateDiff("d", datFechaAnterior, gvntDiasNUtil(intContador)) = 0 Then
                blnIndicador = True
                datFechaAnterior = DateAdd("d", -1, datFechaAnterior)
                '*** Verificar si es Sábado o Domingo ***
                If Weekday(datFechaAnterior) = 1 Or Weekday(datFechaAnterior) = 7 Then
                    datFechaAnterior = DateAdd("d", IIf(Weekday(datFechaAnterior) = 1, -2, -1), datFechaAnterior)
                End If
            End If
        Next
    Loop
    AnteriorDiaUtil = datFechaAnterior
   
End Function
Public Function TirNoPer(ByVal strpCodTitulo As String, ByVal datpFechaLiquidacion As Date, ByVal datpFechaCupon As Date, ByVal curpSubTotal As Currency, ByVal curpInteresCorrido As Currency, ByVal curpCantidadNominal As Currency, ByVal curpCantidadTitulos As Currency, ByVal dblpTirOPeracion As Double, ByVal strpTipoTitulo As String, ByVal strpCodIndiceInicial As String, ByVal strpCodIndiceFinal As String) As Double

    Dim adoRegistro                 As ADODB.Recordset
    Dim intContador                 As Integer, intDiasPeriodo          As Integer
    Dim intNumCupon                 As Integer, intNroCupoFin           As Integer
    Dim intCntDias                  As Integer, intNumRegistros         As Integer
    Dim curMonto                    As Currency, curNewNominal          As Currency
    Dim dblnTasa                    As Double, dblTasDia                As Double
    Dim dblTasa                     As Double, intAcum                  As Double
    Dim dblAjusteInicial            As Double, dblAjusteFinal           As Double
    Dim strIndAmortizacion          As String, strFecha                 As String
    Dim strTipoTasa                 As String, strIndTasaReal           As String
    Dim strPeriodoPago              As String, strClaseTasa             As String
    Dim strFechaInicialIndice       As String, strFechaFinalIndice      As String
    Dim strFechaInicialIndiceMas1   As String, strFechaFinalIndiceMas1  As String
    Dim datFechaInicio              As Date, datFechaFinal              As Date
    Dim datFechaEmision             As Date, dblValorNominalOriginal    As Double
        
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT ValorNominal, IndAmortizacion,CodTipoTasa,CodTipoVac,IndReal,PeriodoPago,FechaEmision " & _
            "FROM InstrumentoInversion WHERE CodTitulo='" & strpCodTitulo & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strIndAmortizacion = Trim(adoRegistro("IndAmortizacion")): strTipoTasa = Trim(adoRegistro("CodTipoTasa"))
            strIndTasaReal = Trim(adoRegistro("IndReal")): strPeriodoPago = Trim(adoRegistro("PeriodoPago"))
            strClaseTasa = Trim(adoRegistro("CodTipoVac"))
            datFechaEmision = adoRegistro("FechaEmision")
            dblValorNominalOriginal = adoRegistro("ValorNominal")
        End If
        adoRegistro.Close

        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodParametro='" & strPeriodoPago & "' AND CodTipoParametro='TIPFRE'"
        Set adoRegistro = .Execute
        If Not adoRegistro.EOF Then
            intDiasPeriodo = CInt(adoRegistro("ValorParametro"))
        End If
        adoRegistro.Close
        
        .CommandText = "SELECT CONVERT(INT,NumCupon) NumCupon FROM InstrumentoInversionCalendario " & _
            "WHERE CodTitulo='" & strpCodTitulo & "' AND " & _
            "FechaInicio<='" & Convertyyyymmdd(datpFechaCupon) & "' AND " & _
            "FechaVencimiento>='" & Convertyyyymmdd(datpFechaCupon) & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            intNumCupon = adoRegistro("NumCupon")
        End If
        adoRegistro.Close
        
        .CommandText = "SELECT FechaVencimiento,FactorInteres,FactorInteres1,ValorAmortizacion,CantDiasPeriodo,NumCupon,FechaInicioIndice,FechaFinIndice,PorcenAmortizacion " & _
            "FROM InstrumentoInversionCalendario " & _
            "WHERE CodTitulo='" & strpCodTitulo & "' AND " & _
            "CONVERT(INT,NumCupon) >= " & intNumCupon & " ORDER BY NumCupon"
        adoRegistro.Open .CommandText, adoConn, adOpenStatic
        adoRegistro.MoveLast
        intNumRegistros = adoRegistro.RecordCount
        
        'ReDim Array_Monto(intNumRegistros + 1): ReDim Array_Dias(intNumRegistros + 1)
        ReDim Array_Monto(intNumRegistros): ReDim Array_Dias(intNumRegistros)
        
        intAcum = 1: dblTasa = dblpTirOPeracion: intContador = 0
        intNroCupoFin = CInt(adoRegistro("NumCupon"))
        
        adoRegistro.MoveFirst
        
        curMonto = (curpSubTotal + curpInteresCorrido) * -1
        Array_Monto(intContador) = curMonto
        strFecha = Convertyyyymmdd(datpFechaLiquidacion)
        datFechaInicio = datpFechaLiquidacion: Array_Dias(intContador) = datFechaInicio
        dblnTasa = 0
'        If strpTipoTitulo = Codigo_Vac_Periodico Then
'            dblnTasa = adoRegistro("FactorInteres1")
'            If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'                dblTasDia = ((1 + adoRegistro("FactorInteres1")) ^ (1 / adoRegistro("CantDiasPeriodo")))
'            Else
'                If strIndTasaReal = Valor_Indicador Then
'                    dblTasDia = adoRegistro("FactorInteres1") / intDiasPeriodo
'                Else
'                    dblTasDia = adoRegistro("FactorInteres1") / adoRegistro("CantDiasPeriodo")
'                End If
'            End If
'            If strIndAmortizacion = Valor_Indicador Then
'                dblnTasa = adoRegistro("FactorInteres1") + adoRegistro("ValorAmortizacion")
'            End If
'        Else
'            dblnTasa = adoRegistro("FactorInteres1")
'            If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'                dblTasDia = ((1 + adoRegistro("FactorInteres")) ^ (1 / adoRegistro("CantDiasPeriodo")))
'            Else
'                If strIndTasaReal = "X" Then
'                    dblTasDia = adoRegistro("FactorInteres") / intDiasPeriodo
'                Else
'                    dblTasDia = adoRegistro("FactorInteres1") / adoRegistro("CantDiasPeriodo")
'                End If
'            End If
'            If strIndAmortizacion = Valor_Indicador Then
'                dblnTasa = adoRegistro("FactorInteres") + adoRegistro("ValorAmortizacion")
'            End If
'        End If

        'If strIndAmortizacion = Valor_Indicador Then
            'curNewNominal = curpCantidadTitulos
        'Else
            curNewNominal = curpCantidadNominal
        'End If

        Do While Not adoRegistro.EOF
            'strFecha = Convertddmmyyyy(adoRegistro("FechaVencimiento"))
            strFechaInicialIndice = Convertyyyymmdd(adoRegistro("FechaInicioIndice"))
            strFechaInicialIndiceMas1 = Convertyyyymmdd(DateAdd("d", 1, adoRegistro("FechaInicioIndice")))
            If strpCodIndiceInicial = Codigo_Vac_Emision Then
                strFechaInicialIndice = Convertyyyymmdd(datFechaEmision)
                strFechaInicialIndiceMas1 = Convertyyyymmdd(DateAdd("d", 1, datFechaEmision))
            End If
            
            strFechaFinalIndice = Convertyyyymmdd(adoRegistro("FechaFinIndice"))
            strFechaFinalIndiceMas1 = Convertyyyymmdd(DateAdd("d", 1, adoRegistro("FechaFinIndice")))
            If strpCodIndiceFinal = Codigo_Vac_Liquidacion Then
                strFechaFinalIndice = Convertyyyymmdd(datpFechaLiquidacion)
                strFechaFinalIndiceMas1 = Convertyyyymmdd(DateAdd("d", 1, datpFechaLiquidacion))
            End If
            
            '*** Es tasa ajustada ? ***
            If strpTipoTitulo <> Valor_Caracter Then
                If strpTipoTitulo = Codigo_Tipo_Ajuste_Vac Then
                    dblAjusteInicial = ObtenerTasaAjuste(Codigo_Tipo_Ajuste_Vac, "00", strFechaInicialIndice, strFechaInicialIndiceMas1)
                    dblAjusteFinal = ObtenerTasaAjuste(Codigo_Tipo_Ajuste_Vac, "00", strFechaFinalIndice, strFechaFinalIndiceMas1)
                Else
                    dblAjusteInicial = ObtenerTasaAjuste(Codigo_Tipo_Ajuste_Vac, strClaseTasa, strFechaInicialIndice, strFechaInicialIndiceMas1)
                    dblAjusteFinal = 0
                End If
            End If
            
            datFechaFinal = adoRegistro("FechaVencimiento")
            intCntDias = DateDiff("d", datFechaInicio, datFechaFinal)
            If strpTipoTitulo = Codigo_Tipo_Ajuste_Vac Then
                If adoRegistro("FactorInteres1") = 0 Then
                    If strIndAmortizacion = Valor_Indicador Then
                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1) + adoRegistro("ValorAmortizacion")
                        Else
                            If strIndTasaReal = Valor_Indicador Then
                                dblnTasa = (dblTasDia * intDiasPeriodo)
                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
                                    dblnTasa = ((dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo) + adoRegistro("ValorAmortizacion")
                                Else
                                    dblnTasa = dblnTasa + adoRegistro("ValorAmortizacion")
                                End If
                            Else
                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo")) + adoRegistro("ValorAmortizacion")
                            End If
                        End If
                    Else
                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1)
                        Else
                            If strIndTasaReal = Valor_Indicador Then
                                dblnTasa = (dblTasDia * intDiasPeriodo)
                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
                                    dblnTasa = (dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo
                                End If
                            Else
                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo"))
                            End If
                        End If
                    End If
                    curMonto = curpCantidadNominal * dblnTasa
                    If strIndAmortizacion = Valor_Indicador Then
                        curMonto = (curNewNominal * dblnTasa) + (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                        curNewNominal = curNewNominal - (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                    End If
                Else
                    curMonto = curpCantidadNominal * adoRegistro("FactorInteres1") * (dblAjusteFinal / dblAjusteInicial)
                    If strIndAmortizacion = Valor_Indicador Then
                        curMonto = (curNewNominal * adoRegistro("FactorInteres1")) + (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                    End If
                    dblnTasa = adoRegistro("FactorInteres1")
                    If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                        dblTasDia = ((1 + adoRegistro("FactorInteres1")) ^ (1 / adoRegistro("CantDiasPeriodo")))
                    Else
                        If strIndTasaReal = Valor_Indicador Then
                            dblTasDia = adoRegistro("FactorInteres1") / intDiasPeriodo
                        Else
                            dblTasDia = adoRegistro("FactorInteres1") / adoRegistro("CantDiasPeriodo")
                        End If
                    End If
                    If strIndAmortizacion = Valor_Indicador Then
                        dblnTasa = adoRegistro("FactorInteres") + adoRegistro("ValorAmortizacion")
                        curNewNominal = curNewNominal - (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                    End If
                End If
            Else
                If adoRegistro("FactorInteres1") = 0 Then
                    If strIndAmortizacion = Valor_Indicador Then
                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1) + adoRegistro("ValorAmortizacion")
                        Else
                            If strIndTasaReal = Valor_Indicador Then
                                dblnTasa = (dblTasDia * intDiasPeriodo)
                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
                                    dblnTasa = ((dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo) + adoRegistro("ValorAmortizacion")
                                Else
                                    dblnTasa = dblnTasa + adoRegistro("ValorAmortizacion")
                                End If
                            Else
                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo")) + adoRegistro("ValorAmortizacion")
                            End If
                        End If
                    Else
                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1)
                        Else
                            If strIndTasaReal = Valor_Indicador Then
                                dblnTasa = (dblTasDia * intDiasPeriodo)
                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
                                    dblnTasa = (dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo
                                End If
                            Else
                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo"))
                            End If
                        End If
                    End If
                    curMonto = Format(curpCantidadNominal * dblnTasa, "0.00")
                    If strIndAmortizacion = Valor_Indicador Then
                        curMonto = (curNewNominal * dblnTasa) + (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                        curNewNominal = curNewNominal - (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                    End If
                Else
                    curMonto = curpCantidadNominal * adoRegistro("FactorInteres1")
                    If strIndAmortizacion = Valor_Indicador Then
                        curMonto = (curNewNominal * adoRegistro("FactorInteres")) + (curpCantidadTitulos * dblValorNominalOriginal * adoRegistro("PorcenAmortizacion") * 0.01)
                    End If
                    dblnTasa = adoRegistro("FactorInteres1")
                    If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                        dblTasDia = ((1 + adoRegistro("FactorInteres")) ^ (1 / adoRegistro("CantDiasPeriodo")))
                    Else
                        If strIndTasaReal = Valor_Indicador Then
                            dblTasDia = adoRegistro("FactorInteres") / intDiasPeriodo
                        Else
                            dblTasDia = adoRegistro("FactorInteres1") / adoRegistro("CantDiasPeriodo")
                        End If
                    End If
                    If strIndAmortizacion = Valor_Indicador Then
                        dblnTasa = adoRegistro("FactorInteres") + adoRegistro("ValorAmortizacion")
                        curNewNominal = curNewNominal - (curpCantidadTitulos * dblValorNominalOriginal * adoRegistro("PorcenAmortizacion") * 0.01)
                    End If
                End If
            End If
            If intNroCupoFin = adoRegistro("NumCupon") Then
                If strIndAmortizacion = Valor_Indicador Then
                    curMonto = Round(curMonto, 2)
                Else
                    curMonto = Round(curMonto + curpCantidadNominal, 2)
                    If strpTipoTitulo = Codigo_Tipo_Ajuste_Vac Then
                        curMonto = Round((curpCantidadNominal * adoRegistro("FactorInteres") + curpCantidadNominal) * (dblAjusteFinal / dblAjusteInicial), 2)
                    End If
                End If
            End If
            intContador = intContador + 1
            Array_Monto(intContador) = curMonto: Array_Dias(intContador) = datFechaFinal
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing

        TirNoPer = TIR(Array_Monto(), Array_Dias(), dblTasa) * 100
    End With
    
End Function
'Public Function TirNoPer1(ByVal strpCodTitulo As String, ByVal datpFechaLiquidacion As Date, ByVal datpFechaCupon As Date, ByVal curpSubTotal As Currency, ByVal curpInteresCorrido As Currency, ByVal curpCantidadNominal As Currency, ByVal curpCantidadTitulos As Currency, ByVal dblpTirOPeracion As Double, ByVal strpTipoTitulo As String, ByVal strpCodIndiceInicial As String, ByVal strpCodIndiceFinal As String) As Double
'
'    Dim adoRegistro                 As ADODB.Recordset
'    Dim intContador                 As Integer
'    Dim intNumCupon                 As Integer
'    Dim intNumRegistros             As Integer
'    Dim curMonto                    As Currency, curNewNominal          As Currency
'    Dim strFecha                    As String, datFechaInicio           As Date
'    Dim datFechaFinal               As Date
'    Dim dblValorNominalOriginal     As Double
'
'    Set adoRegistro = New ADODB.Recordset
'    With adoComm
'
'        'OBTENER EL CUPON VIGENTE
'        .CommandText = "SELECT CONVERT(INT,NumCupon) NumCupon FROM InstrumentoInversionCalendario " & _
'            "WHERE CodTitulo='" & strpCodTitulo & "' AND " & _
'            "FechaInicio<='" & Convertyyyymmdd(datpFechaCupon) & "' AND " & _
'            "FechaVencimiento>='" & Convertyyyymmdd(datpFechaCupon) & "'"
'        Set adoRegistro = .Execute
'
'        If Not adoRegistro.EOF Then
'            intNumCupon = adoRegistro("NumCupon")
'        End If
'        adoRegistro.Close
'
'        'CALCULAR CURSOR CON CUPONES POR VENCER
'        .CommandText = "SELECT FechaVencimiento,ValorCupon FROM InstrumentoInversionCalendario " & _
'            "WHERE CodTitulo='" & strpCodTitulo & "' AND " & _
'            "CONVERT(INT,NumCupon) >= " & intNumCupon & " ORDER BY NumCupon"
'        adoRegistro.Open .CommandText, adoConn, adOpenStatic
'
'        adoRegistro.MoveLast
'
'        intNumRegistros = adoRegistro.RecordCount
'
'        ReDim Array_Monto(intNumRegistros): ReDim Array_Dias(intNumRegistros)
'
'        intContador = 0
'
'        adoRegistro.MoveFirst
'
'        curMonto = (curpSubTotal + curpInteresCorrido) * -1
'        Array_Monto(intContador) = curMonto
'        strFecha = Convertyyyymmdd(datpFechaLiquidacion)
'        datFechaInicio = datpFechaLiquidacion: Array_Dias(intContador) = datFechaInicio
'        curNewNominal = curpCantidadNominal
'
'        Do While Not adoRegistro.EOF
'
'            datFechaFinal = adoRegistro("FechaVencimiento")
'            curMonto = adoRegistro("ValorCupon") * curpCantidadTitulos
'            intContador = intContador + 1
'            Array_Monto(intContador) = curMonto: Array_Dias(intContador) = datFechaFinal
'            adoRegistro.MoveNext
'
'        Loop
'        adoRegistro.Close: Set adoRegistro = Nothing
'
'        TirNoPer = TIR(Array_Monto(), Array_Dias(), dblTasa) * 100
'    End With
'
'End Function
Function TirNoPerPlazo(strpCodTitulo As String, datpFechaOperacion As Date, datpFechaCupon As Date, datpFechaVenta As Date, curpSubTotalPlazo As Currency, curpSubTotal As Currency, curpInteresCorrido As Currency, curpCantidadNominal As Currency, curpCantidadTitulos As Currency, dblpTirOPeracion As Double, strpTipoTitulo As String, intpDiasPlazo As Integer) As Double

    Dim adoRegistro         As ADODB.Recordset
    Dim intContador         As Integer, intDiasPeriodo  As Integer
    Dim intNumCupon         As Integer, intNroCupoFin   As Integer
    Dim intCntDias          As Integer, intNumRegistros As Integer
    Dim curMonto            As Currency, curNewNominal  As Currency
    Dim dblnTasa            As Double, dblTasDia        As Double
    Dim dblTasa             As Double, intAcum          As Double
    Dim strIndAmortizacion  As String, strFecha         As String
    Dim strTipoTasa         As String, strIndTasaReal   As String
    Dim strPeriodoPago      As String
    Dim datFechaInicio      As Date, datFechaFinal      As Date
        
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT IndAmortizacion,CodTipoTasa,IndReal,PeriodoPago " & _
            "FROM InstrumentoInversion WHERE CodTitulo='" & strpCodTitulo & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strIndAmortizacion = Trim(adoRegistro("IndAmortizacion")): strTipoTasa = Trim(adoRegistro("CodTipoTasa"))
            strIndTasaReal = Trim(adoRegistro("IndReal")): strPeriodoPago = Trim(adoRegistro("PeriodoPago"))
        End If
        adoRegistro.Close

        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodParametro='" & strPeriodoPago & "' AND CodTipoParametro='TIPFRE'"
        Set adoRegistro = .Execute
        If Not adoRegistro.EOF Then
            intDiasPeriodo = CInt(adoRegistro("ValorParametro"))
        End If
        adoRegistro.Close
        
        .CommandText = "SELECT CONVERT(INT,NumCupon) NumCupon FROM InstrumentoInversionCalendario " & _
            "WHERE CodTitulo='" & strpCodTitulo & "' AND " & _
            "FechaInicio<='" & Convertyyyymmdd(datpFechaCupon) & "' AND " & _
            "FechaVencimiento>='" & Convertyyyymmdd(datpFechaCupon) & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            intNumCupon = adoRegistro("NumCupon")
        End If
        adoRegistro.Close
        
        .CommandText = "SELECT FechaVencimiento,FactorInteres,FactorInteres1,ValorAmortizacion,CantDiasPeriodo,NumCupon " & _
            "FROM InstrumentoInversionCalendario " & _
            "WHERE CodTitulo='" & strpCodTitulo & "' AND " & _
            "FechaVencimiento>='" & Convertyyyymmdd(datpFechaOperacion) & "'"
            '"CONVERT(INT,NumCupon) >= " & intNumCupon & " ORDER BY NumCupon"
        adoRegistro.Open .CommandText, adoConn, adOpenStatic
        
        adoRegistro.MoveLast
        intNumRegistros = adoRegistro.RecordCount
                
        ReDim Array_Monto(intNumRegistros + 1): ReDim Array_Dias(intNumRegistros + 1)
        
        intAcum = 1: dblTasa = dblpTirOPeracion: intContador = 0
        intNroCupoFin = CInt(adoRegistro("NumCupon"))
        
        adoRegistro.MoveFirst
        
        curMonto = (curpSubTotal + curpInteresCorrido) * -1
        Array_Monto(intContador) = curMonto
        strFecha = Convertyyyymmdd(datpFechaOperacion)
        datFechaInicio = datpFechaOperacion: Array_Dias(intContador) = datFechaInicio
        dblnTasa = 0
'        If strpTipoTitulo = Codigo_Vac_Periodico Then
'            dblnTasa = adoRegistro("FactorInteres1")
'            If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'                dblTasDia = ((1 + adoRegistro("FactorInteres1")) ^ (1 / adoRegistro("CantDiasPeriodo")))
'            Else
'                If strIndTasaReal = Valor_Indicador Then
'                    dblTasDia = adoRegistro("FactorInteres1") / intDiasPeriodo
'                Else
'                    dblTasDia = adoRegistro("FactorInteres1") / adoRegistro("CantDiasPeriodo")
'                End If
'            End If
'            If strIndAmortizacion = Valor_Indicador Then
'                dblnTasa = adoRegistro("FactorInteres1") + adoRegistro("ValorAmortizacion")
'            End If
'        Else
            dblnTasa = adoRegistro("FactorInteres")
            If strTipoTasa = "X" Then
                dblTasDia = ((1 + adoRegistro("FactorInteres")) ^ (1 / adoRegistro("CantDiasPeriodo")))
            Else
                If strIndTasaReal = "X" Then
                    dblTasDia = adoRegistro("FactorInteres") / intDiasPeriodo
                Else
                    dblTasDia = adoRegistro("FactorInteres") / adoRegistro("CantDiasPeriodo")
                End If
            End If
            If strIndAmortizacion = Valor_Indicador Then
                dblnTasa = adoRegistro("FactorInteres") + adoRegistro("ValorAmortizacion")
            End If
'        End If

        curNewNominal = curpCantidadNominal

        Do While Not adoRegistro.EOF
            'strFecha = Convertddmmyyyy(adoRegistro("FechaVencimiento"))
            datFechaFinal = adoRegistro("FechaVencimiento")
            intCntDias = DateDiff("d", datFechaInicio, datFechaFinal)
'            If strpTipoTitulo = Codigo_Vac_Periodico Then
'                If adoRegistro("FactorInteres1") = 0 Then
'                    If strIndAmortizacion = Valor_Indicador Then
'                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1) + adoRegistro("ValorAmortizacion")
'                        Else
'                            If strIndTasaReal = Valor_Indicador Then
'                                dblnTasa = (dblTasDia * intDiasPeriodo)
'                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
'                                    dblnTasa = ((dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo) + adoRegistro("ValorAmortizacion")
'                                Else
'                                    dblnTasa = dblnTasa + adoRegistro("ValorAmortizacion")
'                                End If
'                            Else
'                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo")) + adoRegistro("ValorAmortizacion")
'                            End If
'                        End If
'                    Else
'                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1)
'                        Else
'                            If strIndTasaReal = Valor_Indicador Then
'                                dblnTasa = (dblTasDia * intDiasPeriodo)
'                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
'                                    dblnTasa = (dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo
'                                End If
'                            Else
'                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo"))
'                            End If
'                        End If
'                    End If
'                    curMonto = curpCantidadNominal * dblnTasa
'                    If strIndAmortizacion = Valor_Indicador Then
'                        curMonto = (curNewNominal * dblnTasa) + (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
'                        curNewNominal = curNewNominal - (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
'                    End If
'                Else
'                    curMonto = curpCantidadNominal * adoRegistro("FactorInteres1")
'                    If strIndAmortizacion = Valor_Indicador Then
'                        curMonto = (curNewNominal * adoRegistro("FactorInteres1")) + (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
'                    End If
'                    dblnTasa = adoRegistro("FactorInteres1")
'                    If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'                        dblTasDia = ((1 + adoRegistro("FactorInteres1")) ^ (1 / adoRegistro("CantDiasPeriodo")))
'                    Else
'                        If strIndTasaReal = Valor_Indicador Then
'                            dblTasDia = adoRegistro("FactorInteres1") / intDiasPeriodo
'                        Else
'                            dblTasDia = adoRegistro("FactorInteres1") / adoRegistro("CantDiasPeriodo")
'                        End If
'                    End If
'                    If strIndAmortizacion = Valor_Indicador Then
'                        dblnTasa = adoRegistro("FactorInteres") + adoRegistro("ValorAmortizacion")
'                        curNewNominal = curNewNominal - (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
'                    End If
'                End If
'            Else
                If adoRegistro("FactorInteres") = 0 Then
                    If strIndAmortizacion = Valor_Indicador Then
                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1) + adoRegistro("ValorAmortizacion")
                        Else
                            If strIndTasaReal = Valor_Indicador Then
                                dblnTasa = (dblTasDia * intDiasPeriodo)
                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
                                    dblnTasa = ((dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo) + adoRegistro("ValorAmortizacion")
                                Else
                                    dblnTasa = dblnTasa + adoRegistro("ValorAmortizacion")
                                End If
                            Else
                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo")) + adoRegistro("ValorAmortizacion")
                            End If
                        End If
                    Else
                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1)
                        Else
                            If strIndTasaReal = Valor_Indicador Then
                                dblnTasa = (dblTasDia * intDiasPeriodo)
                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
                                    dblnTasa = (dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo
                                End If
                            Else
                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo"))
                            End If
                        End If
                    End If
                    curMonto = Format(curpCantidadNominal * dblnTasa, "0.00")
                    If strIndAmortizacion = Valor_Indicador Then
                        curMonto = (curNewNominal * dblnTasa) + (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                        curNewNominal = curNewNominal - (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                    End If
                Else
                    curMonto = curpCantidadNominal * adoRegistro("FactorInteres")
                    If strIndAmortizacion = Valor_Indicador Then
                        curMonto = (curNewNominal * adoRegistro("FactorInteres")) + (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                    End If
                    dblnTasa = adoRegistro("FactorInteres")
                    If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                        dblTasDia = ((1 + adoRegistro("FactorInteres")) ^ (1 / adoRegistro("CantDiasPeriodo")))
                    Else
                        If strIndTasaReal = Valor_Indicador Then
                            dblTasDia = adoRegistro("FactorInteres") / intDiasPeriodo
                        Else
                            dblTasDia = adoRegistro("FactorInteres") / adoRegistro("CantDiasPeriodo")
                        End If
                    End If
                    If strIndAmortizacion = Valor_Indicador Then
                        dblnTasa = adoRegistro("FactorInteres") + adoRegistro("ValorAmortizacion")
                        curNewNominal = curNewNominal - (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                    End If
                End If
'            End If
            If intNroCupoFin = CInt(adoRegistro("NumCupon")) Then
                If strIndAmortizacion = Valor_Indicador Then
                    curMonto = Round(curMonto, 2)
                Else
                    curMonto = Round(curMonto + curpCantidadNominal, 2)
                End If
            End If
            intContador = intContador + 1
            If Abs(DateDiff("d", adoRegistro("FechaVencimiento"), datpFechaVenta)) <= intpDiasPlazo Then
                Array_Monto(intContador) = curMonto
            Else
                Array_Monto(intContador) = 0
            End If
            Array_Dias(intContador) = datFechaFinal
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing

        intContador = intContador + 1
        curMonto = curpSubTotalPlazo
        datFechaFinal = datpFechaVenta
        Array_Monto(intContador) = curMonto: Array_Dias(intContador) = datFechaFinal
        
        TirNoPerPlazo = TIR(Array_Monto(), Array_Dias(), dblTasa) * 100
    End With
    
End Function
'Function Duration(strCodFile As String, strCodAnal As String, vntFecOpe, vntFecCup, curSubTot As Currency, curIntCorr As Currency, curMontNomi As Currency, curMontTitu As Currency, dblTirOpe As Double, strTipBono As String) As Double
'
'    Dim adoresult As New Recordset
'    Dim intCont As Integer, intDiasPer As Integer
'    Dim curMonto As Currency, intNroCupo As Integer, strFecha As String * 10
'    Dim dblnTasa As Double, strAmort As String, curNewNominal As Currency, dblTasDia As Double
'    Dim intAcum As Double, dblTasa As Double, vntFchIni, vntFchFin, intCntDias As Integer, intNroCupoFin As Integer
'    Dim intNroReg As Integer, strTipTasa As String, strFlgReal As String, strPerPago As String
'    Dim dblDuration As Double
'
'    With adoComm
'        .CommandText = "SELECT FLG_AMORT,FLG_EFEC,TAS_REAL,PER_PAGO FROM FMBONOS WHERE COD_FILE='" & Trim$(strCodFile) & "' AND COD_ANAL='" & Trim$(strCodAnal) & "'"
'        Set adoresult = .Execute
'        strAmort = Trim(adoresult!FLG_AMORT): strTipTasa = Trim(adoresult!FLG_EFEC)
'        strFlgReal = Trim(adoresult!TAS_REAL): strPerPago = Trim(adoresult!PER_PAGO)
'        adoresult.Close: Set adoresult = Nothing
'
'        .CommandText = "SELECT CNT_DIAS FROM FMTBFREC WHERE PRD_FREC='" & strPerPago & "'"
'        Set adoresult = .Execute
'        If Not adoresult.EOF Then
'            intDiasPer = adoresult!CNT_DIAS
'        End If
'        adoresult.Close: Set adoresult = Nothing
'
'        .CommandText = "SELECT CONVERT(INT,NRO_CUPO) NRO_CUPO FROM FMCUPON"
'        .CommandText = .CommandText & " WHERE COD_FILE='" & Trim$(strCodFile) & "' AND COD_ANAL='" & Trim$(strCodAnal) & "'"
'        '.CommandText = .CommandText & " AND FCH_INIC<='" & Format(vntFecCup, "yyyymmdd") & "'"
'        .CommandText = .CommandText & " AND FCH_INIC<='" & Convertyyyymmdd(vntFecCup) & "'"
'        '.CommandText = .CommandText & " AND FCH_VCTO>='" & Format(vntFecCup, "yyyymmdd") & "'"
'        .CommandText = .CommandText & " AND FCH_VCTO>='" & Convertyyyymmdd(vntFecCup) & "'"
'        Set adoresult = .Execute
'        intNroCupo = adoresult!NRO_CUPO
'        adoresult.Close: Set adoresult = Nothing
'
'        .CommandText = "SELECT FCH_VCTO,TAS_INTE,TAS_INTE2,VAL_AMOR,CNT_DIAS,NRO_CUPO FROM FMCUPON"
'        .CommandText = .CommandText & " WHERE COD_FILE='" & Trim$(strCodFile) & "' AND COD_ANAL='" & Trim$(strCodAnal) & "'"
'        .CommandText = .CommandText & " AND CONVERT(INT,NRO_CUPO) >= " & intNroCupo & " ORDER BY NRO_CUPO"
'        adoresult.Open .CommandText, adoConn, adOpenStatic
'        adoresult.MoveLast
'        intNroReg = adoresult.RecordCount
'
'        'ReDim Array_Monto(intNroReg + 1): ReDim Array_Dias(intNroReg + 1)
'        ReDim Array_Monto(intNroReg): ReDim Array_Dias(intNroReg)
'        intAcum = 1: dblTasa = dblTirOpe: intCont = 0
'        intNroCupoFin = adoresult!NRO_CUPO
'        adoresult.MoveFirst
'        curMonto = Format((curSubTot + curIntCorr) * -1, "0.00")
'        Array_Monto(intCont) = curMonto
'        strFecha = Convertyyyymmdd(vntFecOpe)
'        strFecha = CStr(Convertddmmyyyy(strFecha))
'        vntFchIni = CVDate(strFecha): Array_Dias(intCont) = vntFchIni
'        dblnTasa = 0
'        If strTipBono = "P" Then
'            dblnTasa = adoresult!TAS_INTE2
'            If strTipTasa = "X" Then
'                dblTasDia = Format(((1 + adoresult!TAS_INTE2) ^ (1 / adoresult!CNT_DIAS)), "0.0000000000000000")
'            Else
'                If strFlgReal = "X" Then
'                    dblTasDia = Format(adoresult!TAS_INTE2 / intDiasPer, "0.0000000000000000")
'                Else
'                    dblTasDia = Format(adoresult!TAS_INTE2 / adoresult!CNT_DIAS, "0.0000000000000000")
'                End If
'            End If
'            If (strAmort = "F" Or strAmort = "V") Then
'                dblnTasa = adoresult!TAS_INTE2 + adoresult!VAL_AMOR
'            End If
'        Else
'            dblnTasa = adoresult!TAS_INTE
'            If strTipTasa = "X" Then
'                dblTasDia = Format(((1 + adoresult!TAS_INTE) ^ (1 / adoresult!CNT_DIAS)), "0.0000000000000000")
'            Else
'                If strFlgReal = "X" Then
'                    dblTasDia = Format(adoresult!TAS_INTE / intDiasPer, "0.0000000000000000")
'                Else
'                    dblTasDia = Format(adoresult!TAS_INTE / adoresult!CNT_DIAS, "0.0000000000000000")
'                End If
'            End If
'            If (strAmort = "F" Or strAmort = "V") Then
'                dblnTasa = adoresult!TAS_INTE + adoresult!VAL_AMOR
'            End If
'        End If
'
'        curNewNominal = curMontNomi
'
'        Do While Not adoresult.EOF
'            'strfecha = Right$(adoresult!FCH_VCTO, 2) & "/" & Mid$(adoresult!FCH_VCTO, 5, 2) & "/" & Left$(adoresult!FCH_VCTO, 4)
'            strFecha = CStr(Convertddmmyyyy(adoresult!fch_vcto))
'            vntFchFin = CVDate(strFecha)
'            intCntDias = DateDiff("d", vntFchIni, vntFchFin)
'            If strTipBono = "P" Then
'                If adoresult!TAS_INTE2 = 0 Then
'                    If (strAmort = "F" Or strAmort = "V") Then
'                        If strTipTasa = "X" Then
'                            dblnTasa = ((dblTasDia ^ adoresult!CNT_DIAS) - 1) + adoresult!VAL_AMOR
'                        Else
'                            If strFlgReal = "X" Then
'                                dblnTasa = (dblTasDia * intDiasPer)
'                                If adoresult!CNT_DIAS < intDiasPer Then
'                                    dblnTasa = ((dblnTasa * adoresult!CNT_DIAS) / intDiasPer) + adoresult!VAL_AMOR
'                                Else
'                                    dblnTasa = dblnTasa + adoresult!VAL_AMOR
'                                End If
'                            Else
'                                dblnTasa = (dblTasDia * adoresult!CNT_DIAS) + adoresult!VAL_AMOR
'                            End If
'                        End If
'                    Else
'                        If strTipTasa = "X" Then
'                            dblnTasa = ((dblTasDia ^ adoresult!CNT_DIAS) - 1)
'                        Else
'                            If strFlgReal = "X" Then
'                                dblnTasa = (dblTasDia * intDiasPer)
'                                If adoresult!CNT_DIAS < intDiasPer Then
'                                    dblnTasa = (dblnTasa * adoresult!CNT_DIAS) / intDiasPer
'                                End If
'                            Else
'                                dblnTasa = (dblTasDia * adoresult!CNT_DIAS)
'                            End If
'                        End If
'                    End If
'                    curMonto = Format(curMontNomi * dblnTasa, "0.00")
'                    If (strAmort = "F" Or strAmort = "V") Then
'                        curMonto = Format((curNewNominal * dblnTasa) + (curMontTitu * adoresult!VAL_AMOR), "0.00")
'                        curNewNominal = curNewNominal - (curMontTitu * adoresult!VAL_AMOR)
'                    End If
'                Else
'                    curMonto = Format(curMontNomi * adoresult!TAS_INTE2, "0.00")
'                    If (strAmort = "F" Or strAmort = "V") Then
'                        curMonto = Format((curNewNominal * adoresult!TAS_INTE2) + (curMontTitu * adoresult!VAL_AMOR), "0.00")
'                    End If
'                    dblnTasa = adoresult!TAS_INTE2
'                    If strTipTasa = "X" Then
'                        dblTasDia = Format(((1 + adoresult!TAS_INTE2) ^ (1 / adoresult!CNT_DIAS)), "0.0000000000000000")
'                    Else
'                        If strFlgReal = "X" Then
'                            dblTasDia = Format(adoresult!TAS_INTE2 / intDiasPer, "0.0000000000000000")
'                        Else
'                            dblTasDia = Format(adoresult!TAS_INTE2 / adoresult!CNT_DIAS, "0.0000000000000000")
'                        End If
'                    End If
'                    If (strAmort = "F" Or strAmort = "V") Then
'                        dblnTasa = adoresult!TAS_INTE + adoresult!VAL_AMOR
'                        curNewNominal = curNewNominal - (curMontTitu * adoresult!VAL_AMOR)
'                    End If
'                End If
'            Else
'                If adoresult!TAS_INTE = 0 Then
'                    If (strAmort = "F" Or strAmort = "V") Then
'                        If strTipTasa = "X" Then
'                            dblnTasa = ((dblTasDia ^ adoresult!CNT_DIAS) - 1) + adoresult!VAL_AMOR
'                        Else
'                            If strFlgReal = "X" Then
'                                dblnTasa = (dblTasDia * intDiasPer)
'                                If adoresult!CNT_DIAS < intDiasPer Then
'                                    dblnTasa = ((dblnTasa * adoresult!CNT_DIAS) / intDiasPer) + adoresult!VAL_AMOR
'                                Else
'                                    dblnTasa = dblnTasa + adoresult!VAL_AMOR
'                                End If
'                            Else
'                                dblnTasa = (dblTasDia * adoresult!CNT_DIAS) + adoresult!VAL_AMOR
'                            End If
'                        End If
'                    Else
'                        If strTipTasa = "X" Then
'                            dblnTasa = ((dblTasDia ^ adoresult!CNT_DIAS) - 1)
'                        Else
'                            If strFlgReal = "X" Then
'                                dblnTasa = (dblTasDia * intDiasPer)
'                                If adoresult!CNT_DIAS < intDiasPer Then
'                                    dblnTasa = (dblnTasa * adoresult!CNT_DIAS) / intDiasPer
'                                End If
'                            Else
'                                dblnTasa = (dblTasDia * adoresult!CNT_DIAS)
'                            End If
'                        End If
'                    End If
'                    curMonto = Format(curMontNomi * dblnTasa, "0.00")
'                    If (strAmort = "F" Or strAmort = "V") Then
'                        curMonto = Format((curNewNominal * dblnTasa) + (curMontTitu * adoresult!VAL_AMOR), "0.00")
'                        curNewNominal = curNewNominal - (curMontTitu * adoresult!VAL_AMOR)
'                    End If
'                Else
'                    curMonto = Format(curMontNomi * adoresult!TAS_INTE, "0.00")
'                    If (strAmort = "F" Or strAmort = "V") Then
'                        curMonto = Format((curNewNominal * adoresult!TAS_INTE) + (curMontTitu * adoresult!VAL_AMOR), "0.00")
'                    End If
'                    dblnTasa = adoresult!TAS_INTE
'                    If strTipTasa = "X" Then
'                        dblTasDia = Format(((1 + adoresult!TAS_INTE) ^ (1 / adoresult!CNT_DIAS)), "0.0000000000000000")
'                    Else
'                        If strFlgReal = "X" Then
'                            dblTasDia = Format(adoresult!TAS_INTE / intDiasPer, "0.0000000000000000")
'                        Else
'                            dblTasDia = Format(adoresult!TAS_INTE / adoresult!CNT_DIAS, "0.0000000000000000")
'                        End If
'                    End If
'                    If (strAmort = "F" Or strAmort = "V") Then
'                        dblnTasa = adoresult!TAS_INTE + adoresult!VAL_AMOR
'                        curNewNominal = curNewNominal - (curMontTitu * adoresult!VAL_AMOR)
'                    End If
'                End If
'            End If
'            If intNroCupoFin = adoresult!NRO_CUPO Then
'                If (strAmort = "F" Or strAmort = "V") Then
'                    curMonto = Format(curMonto, "0.00")
'                Else
'                    curMonto = Format(curMonto + curMontNomi, "0.00")
'                End If
'            End If
'            intCont = intCont + 1
'            Array_Monto(intCont) = curMonto: Array_Dias(intCont) = vntFchFin
'            dblDuration = dblDuration + ((curMonto * intCntDias) / (((1 + dblTasa) ^ (intCntDias / 365)) * 365 * (curSubTot + curIntCorr)))
'            adoresult.MoveNext
'        Loop
'        adoresult.Close: Set adoresult = Nothing
'
'        'TirNoPer = TIR(Array_Monto(), Array_Dias(), dblTasa) * 100
'        Duration = dblDuration
'    End With
'
'End Function
'Function TirNoPerCrtDepC(strCodFon As String, Codfile As String, CodAnal As String, FecOpe, FecCup, SubTot As Double, IntCorr As Double, MontNomi As Double, MontTitu As Double, TirOpe As Double, TipBono As String) As Double
'
'    Dim adoresult As New Recordset
'    Dim n_Cont As Integer
'    Dim n_Monto As Double, n_NroCupo As Integer, s_Fecha As String * 10
'    Dim n_Tasa As Double, s_Amort As String, n_NewNominal As Double, n_TasDia As Double
'    Dim n_Acum As Double, Tasa As Double, d_FchIni, d_FchFin, n_CntDias As Integer, n_NroCupoFin As Integer
'    Dim n_Reg As Integer, i As Integer
'
'    adoComm.CommandText = "SELECT FLG_AMORT FROM FMDEPBAN WHERE COD_FILE='" & Trim$(Codfile) & "' AND COD_ANAL='" & Trim$(CodAnal) & "' and COD_FOND='" & strCodFon & "'"
'    Set adoresult = adoComm.Execute
'    s_Amort = IIf(IsNull(adoresult!FLG_AMORT), "", Trim(adoresult!FLG_AMORT))
'    adoresult.Close: Set adoresult = Nothing
'
'    With adoComm
'        .CommandText = "SELECT CONVERT(INT,NRO_CUPO) NRO_CUPO FROM FMCUPONES"
'        .CommandText = .CommandText & " WHERE COD_FILE='" & Trim$(Codfile) & "' AND COD_ANAL='" & Trim$(CodAnal) & "' AND COD_FOND='" & strCodFon & "'"
'        '.CommandText = .CommandText & " AND FCH_INIC<='" & Format(FecCup, "yyyymmdd") & "'"
'        .CommandText = .CommandText & " AND FCH_INIC<='" & Convertyyyymmdd(FecCup) & "'"
'        '.CommandText = .CommandText & " AND FCH_VCTO>='" & Format(FecCup, "yyyymmdd") & "'"
'        .CommandText = .CommandText & " AND FCH_VCTO>='" & Convertyyyymmdd(FecCup) & "'"
'        Set adoresult = .Execute
'        If Not adoresult.EOF Then
'            n_NroCupo = IIf(IsNull(adoresult!NRO_CUPO), "", adoresult!NRO_CUPO)
'        Else
'            Exit Function
'        End If
'        adoresult.Close: Set adoresult = Nothing
'
'        .CommandText = "SELECT FCH_VCTO,TAS_INTE,TAS_INTE2,VAL_AMOR,CNT_DIAS,NRO_CUPO FROM FMCUPONES"
'        .CommandText = .CommandText & " WHERE COD_FILE='" & Trim(Codfile) & "' AND COD_ANAL='" & Trim(CodAnal) & "'"
'        .CommandText = .CommandText & " AND CONVERT(INT,NRO_CUPO) >= " & n_NroCupo & " ORDER BY NRO_CUPO"
'        adoresult.Open .CommandText, adoConn, adOpenStatic
'
'        adoresult.MoveLast
'        n_Reg = adoresult.RecordCount
'
'        ReDim Array_Monto(n_Reg + 1): ReDim Array_Dias(n_Reg + 1)
'        ReDim Array_Monto(n_Reg): ReDim Array_Dias(n_Reg)
'        n_Acum = 1: Tasa = TirOpe: n_Cont = 0
'        n_NroCupoFin = adoresult!NRO_CUPO
'        adoresult.MoveFirst
'
'        n_Monto = Format((SubTot + IntCorr) * -1, "0.00")
'        Array_Monto(n_Cont) = n_Monto
'        's_fecha = Trim$(Format(FecOpe, "yyyymmdd"))
'        s_Fecha = Convertyyyymmdd(FecOpe)
'        s_Fecha = CStr(Convertddmmyyyy(s_Fecha))
'        d_FchIni = CVDate(s_Fecha): Array_Dias(n_Cont) = d_FchIni
'        n_Tasa = 0
'        If TipBono = "P" Then
'            n_Tasa = adoresult!TAS_INTE2
'            n_TasDia = Format(((1 + adoresult!TAS_INTE2) ^ (1 / adoresult!CNT_DIAS)), "0.0000000000000000")
'            If (s_Amort = "F" Or s_Amort = "V") Then
'                n_Tasa = adoresult!TAS_INTE2 + adoresult!VAL_AMOR
'            End If
'        Else
'            n_Tasa = adoresult!TAS_INTE
'            n_TasDia = Format(((1 + adoresult!TAS_INTE) ^ (1 / adoresult!CNT_DIAS)), "0.0000000000000000")
'            If (s_Amort = "F" Or s_Amort = "V") Then
'                n_Tasa = adoresult!TAS_INTE + adoresult!VAL_AMOR
'            End If
'        End If
'
'        n_NewNominal = MontNomi
'
'        Do While Not adoresult.EOF
'            's_fecha = Right$(adoresult!FCH_VCTO, 2) & "/" & Mid$(adoresult!FCH_VCTO, 5, 2) & "/" & Left$(adoresult!FCH_VCTO, 4)
'            s_Fecha = CStr(Convertddmmyyyy(adoresult!fch_vcto))
'            d_FchFin = CVDate(s_Fecha)
'            n_CntDias = DateDiff("d", d_FchIni, d_FchFin)
'            If TipBono = "P" Then
'                If adoresult!TAS_INTE2 = 0 Then
'                    If (s_Amort = "F" Or s_Amort = "V") Then
'                        n_Tasa = ((n_TasDia ^ adoresult!CNT_DIAS) - 1) + adoresult!VAL_AMOR
'                    Else
'                        n_Tasa = ((n_TasDia ^ adoresult!CNT_DIAS) - 1)
'                    End If
'                    n_Monto = Format(MontNomi * n_Tasa, "0.00")
'                    If (s_Amort = "F" Or s_Amort = "V") Then
'                        n_Monto = Format((n_NewNominal * n_Tasa) + (MontTitu * adoresult!VAL_AMOR), "0.00")
'                        n_NewNominal = n_NewNominal - (MontTitu * adoresult!VAL_AMOR)
'                    End If
'                Else
'                    n_Monto = Format(MontNomi * adoresult!TAS_INTE2, "0.00")
'                    If (s_Amort = "F" Or s_Amort = "V") Then
'                        n_Monto = Format((n_NewNominal * adoresult!TAS_INTE2) + (MontTitu * adoresult!VAL_AMOR), "0.00")
'                    End If
'                    n_Tasa = adoresult!TAS_INTE2
'                    n_TasDia = Format(((1 + adoresult!TAS_INTE2) ^ (1 / adoresult!CNT_DIAS)), "0.0000000000000000")
'                    If (s_Amort = "F" Or s_Amort = "V") Then
'                        n_Tasa = adoresult!TAS_INTE + adoresult!VAL_AMOR
'                        n_NewNominal = n_NewNominal - (MontTitu * adoresult!VAL_AMOR)
'                    End If
'                End If
'            Else
'                If adoresult!TAS_INTE = 0 Then
'                    If (s_Amort = "F" Or s_Amort = "V") Then
'                        n_Tasa = ((n_TasDia ^ adoresult!CNT_DIAS) - 1) + adoresult!VAL_AMOR
'                    Else
'                        n_Tasa = ((n_TasDia ^ adoresult!CNT_DIAS) - 1)
'                    End If
'                    n_Monto = Format(MontNomi * n_Tasa, "0.00")
'                    If (s_Amort = "F" Or s_Amort = "V") Then
'                        n_Monto = Format((n_NewNominal * n_Tasa) + (MontTitu * adoresult!VAL_AMOR), "0.00")
'                        n_NewNominal = n_NewNominal - (MontTitu * adoresult!VAL_AMOR)
'                    End If
'                Else
'                    n_Monto = Format(MontNomi * adoresult!TAS_INTE, "0.00")
'                    If (s_Amort = "F" Or s_Amort = "V") Then
'                        n_Monto = Format((n_NewNominal * adoresult!TAS_INTE) + (MontTitu * adoresult!VAL_AMOR), "0.00")
'                    End If
'                    n_Tasa = adoresult!TAS_INTE
'                    n_TasDia = Format(((1 + adoresult!TAS_INTE) ^ (1 / adoresult!CNT_DIAS)), "0.0000000000000000")
'                    If (s_Amort = "F" Or s_Amort = "V") Then
'                        n_Tasa = adoresult!TAS_INTE + adoresult!VAL_AMOR
'                        n_NewNominal = n_NewNominal - (MontTitu * adoresult!VAL_AMOR)
'                    End If
'                End If
'            End If
'            If n_NroCupoFin = adoresult!NRO_CUPO Then
'                If (s_Amort = "F" Or s_Amort = "V") Then
'                    n_Monto = Format(n_Monto, "0.00")
'                Else
'                    n_Monto = Format(n_Monto + MontNomi, "0.00")
'                End If
'            End If
'            n_Cont = n_Cont + 1
'            Array_Monto(n_Cont) = n_Monto: Array_Dias(n_Cont) = d_FchFin
'            adoresult.MoveNext
'        Loop
'        adoresult.Close: Set adoresult = Nothing
'    End With
'
'    TirNoPerCrtDepC = TIRDepCrt(Array_Monto(), Array_Dias(), Tasa) * 100
'
'End Function
Public Function Duration(ByVal strpCodTitulo As String, ByVal datpFechaLiquidacion As Date, ByVal datpFechaCupon As Date, ByVal curpSubTotal As Currency, ByVal curpInteresCorrido As Currency, ByVal curpCantidadNominal As Currency, ByVal curpCantidadTitulos As Currency, ByVal dblpTirOPeracion As Double, ByVal strpTipoTitulo As String, ByVal strpCodIndiceInicial As String, ByVal strpCodIndiceFinal As String) As Double

    Dim adoRegistro                 As ADODB.Recordset
    Dim intContador                 As Integer, intDiasPeriodo          As Integer
    Dim intNumCupon                 As Integer, intNroCupoFin           As Integer
    Dim intCntDias                  As Integer, intNumRegistros         As Integer
    Dim curMonto                    As Currency, curNewNominal          As Currency
    Dim dblnTasa                    As Double, dblTasDia                As Double
    Dim dblTasa                     As Double, intAcum                  As Double
    Dim dblAjusteInicial            As Double, dblAjusteFinal           As Double
    Dim dblDuration                 As Double
    Dim strIndAmortizacion          As String, strFecha                 As String
    Dim strTipoTasa                 As String, strIndTasaReal           As String
    Dim strPeriodoPago              As String, strClaseTasa             As String
    Dim strFechaInicialIndice       As String, strFechaFinalIndice      As String
    Dim strFechaInicialIndiceMas1   As String, strFechaFinalIndiceMas1  As String
    Dim datFechaInicio              As Date, datFechaFinal              As Date
    Dim datFechaEmision             As Date
        
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT IndAmortizacion,CodTipoTasa,CodTipoVac,IndReal,PeriodoPago,FechaEmision " & _
            "FROM InstrumentoInversion WHERE CodTitulo='" & strpCodTitulo & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strIndAmortizacion = Trim(adoRegistro("IndAmortizacion")): strTipoTasa = Trim(adoRegistro("CodTipoTasa"))
            strIndTasaReal = Trim(adoRegistro("IndReal")): strPeriodoPago = Trim(adoRegistro("PeriodoPago"))
            strClaseTasa = Trim(adoRegistro("CodTipoVac"))
            datFechaEmision = adoRegistro("FechaEmision")
        End If
        adoRegistro.Close

        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodParametro='" & strPeriodoPago & "' AND CodTipoParametro='TIPFRE'"
        Set adoRegistro = .Execute
        If Not adoRegistro.EOF Then
            intDiasPeriodo = CInt(adoRegistro("ValorParametro"))
        End If
        adoRegistro.Close
        
        .CommandText = "SELECT CONVERT(INT,NumCupon) NumCupon FROM InstrumentoInversionCalendario " & _
            "WHERE CodTitulo='" & strpCodTitulo & "' AND " & _
            "FechaInicio<='" & Convertyyyymmdd(datpFechaCupon) & "' AND " & _
            "FechaVencimiento>='" & Convertyyyymmdd(datpFechaCupon) & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            intNumCupon = adoRegistro("NumCupon")
        End If
        adoRegistro.Close
        
        .CommandText = "SELECT FechaVencimiento,FactorInteres,FactorInteres1,ValorAmortizacion,CantDiasPeriodo,NumCupon,FechaInicioIndice,FechaFinIndice,PorcenAmortizacion " & _
            "FROM InstrumentoInversionCalendario " & _
            "WHERE CodTitulo='" & strpCodTitulo & "' AND " & _
            "CONVERT(INT,NumCupon) >= " & intNumCupon & " ORDER BY NumCupon"
        adoRegistro.Open .CommandText, adoConn, adOpenStatic
        adoRegistro.MoveLast
        intNumRegistros = adoRegistro.RecordCount
        
        'ReDim Array_Monto(intNumRegistros + 1): ReDim Array_Dias(intNumRegistros + 1)
        ReDim Array_Monto(intNumRegistros): ReDim Array_Dias(intNumRegistros)
        
        intAcum = 1: dblTasa = dblpTirOPeracion: intContador = 0
        intNroCupoFin = CInt(adoRegistro("NumCupon"))
        
        adoRegistro.MoveFirst
        
        curMonto = (curpSubTotal + curpInteresCorrido) * -1
        Array_Monto(intContador) = curMonto
        strFecha = Convertyyyymmdd(datpFechaLiquidacion)
        datFechaInicio = datpFechaLiquidacion: Array_Dias(intContador) = datFechaInicio
        dblnTasa = 0
'        If strpTipoTitulo = Codigo_Vac_Periodico Then
'            dblnTasa = adoRegistro("FactorInteres1")
'            If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'                dblTasDia = ((1 + adoRegistro("FactorInteres1")) ^ (1 / adoRegistro("CantDiasPeriodo")))
'            Else
'                If strIndTasaReal = Valor_Indicador Then
'                    dblTasDia = adoRegistro("FactorInteres1") / intDiasPeriodo
'                Else
'                    dblTasDia = adoRegistro("FactorInteres1") / adoRegistro("CantDiasPeriodo")
'                End If
'            End If
'            If strIndAmortizacion = Valor_Indicador Then
'                dblnTasa = adoRegistro("FactorInteres1") + adoRegistro("ValorAmortizacion")
'            End If
'        Else
'            dblnTasa = adoRegistro("FactorInteres1")
'            If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'                dblTasDia = ((1 + adoRegistro("FactorInteres")) ^ (1 / adoRegistro("CantDiasPeriodo")))
'            Else
'                If strIndTasaReal = "X" Then
'                    dblTasDia = adoRegistro("FactorInteres") / intDiasPeriodo
'                Else
'                    dblTasDia = adoRegistro("FactorInteres1") / adoRegistro("CantDiasPeriodo")
'                End If
'            End If
'            If strIndAmortizacion = Valor_Indicador Then
'                dblnTasa = adoRegistro("FactorInteres") + adoRegistro("ValorAmortizacion")
'            End If
'        End If

        If strIndAmortizacion = Valor_Indicador Then
            curNewNominal = curpCantidadTitulos
        Else
            curNewNominal = curpCantidadNominal
        End If

        Do While Not adoRegistro.EOF
            'strFecha = Convertddmmyyyy(adoRegistro("FechaVencimiento"))
            strFechaInicialIndice = Convertyyyymmdd(adoRegistro("FechaInicioIndice"))
            strFechaInicialIndiceMas1 = Convertyyyymmdd(DateAdd("d", 1, adoRegistro("FechaInicioIndice")))
            If strpCodIndiceInicial = Codigo_Vac_Emision Then
                strFechaInicialIndice = Convertyyyymmdd(datFechaEmision)
                strFechaInicialIndiceMas1 = Convertyyyymmdd(DateAdd("d", 1, datFechaEmision))
            End If
            
            strFechaFinalIndice = Convertyyyymmdd(adoRegistro("FechaFinIndice"))
            strFechaFinalIndiceMas1 = Convertyyyymmdd(DateAdd("d", 1, adoRegistro("FechaFinIndice")))
            If strpCodIndiceFinal = Codigo_Vac_Liquidacion Then
                strFechaFinalIndice = Convertyyyymmdd(datpFechaLiquidacion)
                strFechaFinalIndiceMas1 = Convertyyyymmdd(DateAdd("d", 1, datpFechaLiquidacion))
            End If
            
            '*** Es tasa ajustada ? ***
            If strpTipoTitulo <> Valor_Caracter Then
                If strpTipoTitulo = Codigo_Tipo_Ajuste_Vac Then
                    dblAjusteInicial = ObtenerTasaAjuste(Codigo_Tipo_Ajuste_Vac, "00", strFechaInicialIndice, strFechaInicialIndiceMas1)
                    dblAjusteFinal = ObtenerTasaAjuste(Codigo_Tipo_Ajuste_Vac, "00", strFechaFinalIndice, strFechaFinalIndiceMas1)
                Else
                    dblAjusteInicial = ObtenerTasaAjuste(Codigo_Tipo_Ajuste_Vac, strClaseTasa, strFechaInicialIndice, strFechaInicialIndiceMas1)
                    dblAjusteFinal = 0
                End If
            End If
            
            datFechaFinal = adoRegistro("FechaVencimiento")
            intCntDias = DateDiff("d", datFechaInicio, datFechaFinal)
            If strpTipoTitulo = Codigo_Tipo_Ajuste_Vac Then
                If adoRegistro("FactorInteres1") = 0 Then
                    If strIndAmortizacion = Valor_Indicador Then
                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1) + adoRegistro("ValorAmortizacion")
                        Else
                            If strIndTasaReal = Valor_Indicador Then
                                dblnTasa = (dblTasDia * intDiasPeriodo)
                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
                                    dblnTasa = ((dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo) + adoRegistro("ValorAmortizacion")
                                Else
                                    dblnTasa = dblnTasa + adoRegistro("ValorAmortizacion")
                                End If
                            Else
                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo")) + adoRegistro("ValorAmortizacion")
                            End If
                        End If
                    Else
                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1)
                        Else
                            If strIndTasaReal = Valor_Indicador Then
                                dblnTasa = (dblTasDia * intDiasPeriodo)
                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
                                    dblnTasa = (dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo
                                End If
                            Else
                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo"))
                            End If
                        End If
                    End If
                    curMonto = curpCantidadNominal * dblnTasa
                    If strIndAmortizacion = Valor_Indicador Then
                        curMonto = (curNewNominal * dblnTasa) + (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                        curNewNominal = curNewNominal - (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                    End If
                Else
                    curMonto = curpCantidadNominal * adoRegistro("FactorInteres1") * (dblAjusteFinal / dblAjusteInicial)
                    If strIndAmortizacion = Valor_Indicador Then
                        curMonto = (curNewNominal * adoRegistro("FactorInteres1")) + (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                    End If
                    dblnTasa = adoRegistro("FactorInteres1")
                    If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                        dblTasDia = ((1 + adoRegistro("FactorInteres1")) ^ (1 / adoRegistro("CantDiasPeriodo")))
                    Else
                        If strIndTasaReal = Valor_Indicador Then
                            dblTasDia = adoRegistro("FactorInteres1") / intDiasPeriodo
                        Else
                            dblTasDia = adoRegistro("FactorInteres1") / adoRegistro("CantDiasPeriodo")
                        End If
                    End If
                    If strIndAmortizacion = Valor_Indicador Then
                        dblnTasa = adoRegistro("FactorInteres") + adoRegistro("ValorAmortizacion")
                        curNewNominal = curNewNominal - (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                    End If
                End If
            Else
                If adoRegistro("FactorInteres1") = 0 Then
                    If strIndAmortizacion = Valor_Indicador Then
                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1) + adoRegistro("ValorAmortizacion")
                        Else
                            If strIndTasaReal = Valor_Indicador Then
                                dblnTasa = (dblTasDia * intDiasPeriodo)
                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
                                    dblnTasa = ((dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo) + adoRegistro("ValorAmortizacion")
                                Else
                                    dblnTasa = dblnTasa + adoRegistro("ValorAmortizacion")
                                End If
                            Else
                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo")) + adoRegistro("ValorAmortizacion")
                            End If
                        End If
                    Else
                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1)
                        Else
                            If strIndTasaReal = Valor_Indicador Then
                                dblnTasa = (dblTasDia * intDiasPeriodo)
                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
                                    dblnTasa = (dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo
                                End If
                            Else
                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo"))
                            End If
                        End If
                    End If
                    curMonto = Format(curpCantidadNominal * dblnTasa, "0.00")
                    If strIndAmortizacion = Valor_Indicador Then
                        curMonto = (curNewNominal * dblnTasa) + (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                        curNewNominal = curNewNominal - (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                    End If
                Else
                    curMonto = curpCantidadNominal * adoRegistro("FactorInteres1")
                    If strIndAmortizacion = Valor_Indicador Then
                        curMonto = (curNewNominal * adoRegistro("FactorInteres")) + (curpCantidadNominal * adoRegistro("PorcenAmortizacion") * 0.01)
                    End If
                    dblnTasa = adoRegistro("FactorInteres1")
                    If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                        dblTasDia = ((1 + adoRegistro("FactorInteres")) ^ (1 / adoRegistro("CantDiasPeriodo")))
                    Else
                        If strIndTasaReal = Valor_Indicador Then
                            dblTasDia = adoRegistro("FactorInteres") / intDiasPeriodo
                        Else
                            dblTasDia = adoRegistro("FactorInteres1") / adoRegistro("CantDiasPeriodo")
                        End If
                    End If
                    If strIndAmortizacion = Valor_Indicador Then
                        dblnTasa = adoRegistro("FactorInteres") + adoRegistro("ValorAmortizacion")
                        curNewNominal = curNewNominal - (curpCantidadNominal * adoRegistro("PorcenAmortizacion") * 0.01)
                    End If
                End If
            End If
            If intNroCupoFin = adoRegistro("NumCupon") Then
                If strIndAmortizacion = Valor_Indicador Then
                    curMonto = Round(curMonto, 2)
                Else
                    curMonto = Round(curMonto + curpCantidadNominal, 2)
                    If strpTipoTitulo = Codigo_Tipo_Ajuste_Vac Then
                        curMonto = Round((curpCantidadNominal * adoRegistro("FactorInteres") + curpCantidadNominal) * (dblAjusteFinal / dblAjusteInicial), 2)
                    End If
                End If
            End If
            intContador = intContador + 1
'            Array_Monto(intContador) = curMonto: Array_Dias(intContador) = datFechaFinal
            dblDuration = dblDuration + ((curMonto * intCntDias) / (((1 + dblTasa) ^ (intCntDias / 365)) * 365 * (curpSubTotal + curpInteresCorrido)))
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing

'        TirNoPer = TIR(Array_Monto(), Array_Dias(), dblTasa) * 100
        Duration = dblDuration
    End With
End Function
Public Function TIR(ByRef vntValFlujo(), ByRef vntFchFlujo(), ByVal dblpValorTir As Double) As Double

    '*** Calcular la TIR tomando en cuenta los flujos futuros ***
    Dim intContador     As Integer, intBaseAnual As Integer
    Dim lngAcumulado    As Long
    Dim vntFchFlujo1()

    'On Error GoTo ErrorHandler1

    lngAcumulado = 0

    ReDim vntFchFlujo1(UBound(vntFchFlujo))

    '*** Base Anual ***
    intBaseAnual = 365

    '*** Cambiar Fechas ***
    For intContador = 0 To UBound(vntValFlujo)
        vntFchFlujo1(intContador) = DateAdd("d", 1, vntFchFlujo(intContador))
    Next

    '*** Calculando ***
    Do Until Abs(VAN(vntValFlujo(), vntFchFlujo(), dblpValorTir, 0)) <= 0.00000001 Or lngAcumulado = 100
        dblpValorTir = dblpValorTir + intBaseAnual * (VAN(vntValFlujo(), vntFchFlujo(), dblpValorTir, 0) / VAN(vntValFlujo(), vntFchFlujo1(), dblpValorTir, 1))
        lngAcumulado = lngAcumulado + 1
    Loop

    TIR = dblpValorTir

ExitFunction1:
    Exit Function

ErrorHandler1:

    MsgBox Error(err), 48
    TIR = 0
    Resume ExitFunction1

End Function


Function VNANoPerLetPag(Codfile As String, CodAnal As String, CODFON As String, FecOpe, FecFlujo, MontNomi As Double, MontTitu As Double, TirOpe As Double, TipBono As String)

    Dim adoresultTmp As New Recordset
    Dim n_Cont As Integer
    Dim n_Monto As Double, n_NroCupo As Integer, s_fecha As String * 10
    Dim n_Tasa As Double, s_Amort As String, n_NewNominal As Double, n_TasDia As Double
    Dim n_Acum As Double, Tasa As Double, d_FchIni, d_FchFin, n_CntDias As Integer, n_NroCupoFin As Integer
    Dim n_Reg As Integer, i As Integer
    
    adoComm.CommandText = "SELECT FLG_AMORT FROM FMLETRAS WHERE COD_FILE='" + Trim$(Codfile) + "' AND COD_ANAL='" + Trim$(CodAnal) + "' AND COD_FOND='" + Trim$(CODFON) + "'"
    Set adoresultTmp = adoComm.Execute
    s_Amort = adoresultTmp!FLG_AMORT
    adoresultTmp.Close: Set adoresultTmp = Nothing

    With adoComm
        .CommandText = "SELECT CONVERT(INT,NRO_CUPO) NRO_CUPO FROM FMCUPONES"
        .CommandText = .CommandText + " WHERE COD_FILE='" + Trim$(Codfile) + "' AND COD_ANAL='" + Trim$(CodAnal) + "' AND COD_FOND='" + Trim$(CODFON) + "'"
        '.CommandText = .CommandText + " AND FCH_INIC<='" + Format(FecFlujo, "yyyymmdd") + "'"
        .CommandText = .CommandText + " AND FCH_INIC<='" + Convertyyyymmdd(FecFlujo) + "'"
        '.CommandText = .CommandText + " AND FCH_VCTO>='" + Format(FecFlujo, "yyyymmdd") + "'"
        .CommandText = .CommandText + " AND FCH_VCTO>='" + Convertyyyymmdd(FecFlujo) + "'"
        Set adoresultTmp = .Execute
    End With
    n_NroCupo = adoresultTmp!NRO_CUPO
    adoresultTmp.Close: Set adoresultTmp = Nothing

    With adoComm
        .CommandText = "SELECT FCH_VCTO,TAS_INTE,TAS_INTE2,VAL_AMOR,CNT_DIAS,NRO_CUPO FROM FMCUPONES"
        .CommandText = .CommandText + " WHERE COD_FILE='" + Trim$(Codfile) + "' AND COD_ANAL='" + Trim$(CodAnal) + "' AND COD_FOND='" + Trim$(CODFON) + "'"
        .CommandText = .CommandText + " AND CONVERT(INT,NRO_CUPO) >= " & n_NroCupo & " ORDER BY NRO_CUPO"
        adoresultTmp.Open .CommandText, adoConn, adOpenStatic
    End With
    
    adoresultTmp.MoveLast
    n_Reg = adoresultTmp.RecordCount

    ReDim Array_Monto(n_Reg + 1): ReDim Array_Dias(n_Reg + 1)

    n_Acum = 1: Tasa = TirOpe: n_Cont = 1
    n_NroCupoFin = adoresultTmp!NRO_CUPO
    adoresultTmp.MoveFirst

    n_Monto = Format(0, "0.00")
    Array_Monto(n_Cont) = 0: Array_Dias(n_Cont) = 0
    's_fecha = Trim$(Format(FecFlujo, "yyyymmdd"))
    s_fecha = Convertyyyymmdd(FecFlujo)
    s_fecha = CStr(Convertddmmyyyy(s_fecha))
    d_FchIni = CVDate(s_fecha)
    n_Tasa = 0
    If TipBono = "P" Then
        n_Tasa = adoresultTmp!TAS_INTE2
        n_TasDia = Format(((1 + adoresultTmp!TAS_INTE2) ^ (1 / adoresultTmp!CNT_DIAS)), "0.0000000000000000")
        If (s_Amort = "F" Or s_Amort = "V") Then
            n_Tasa = adoresultTmp!TAS_INTE2 + adoresultTmp!VAL_AMOR
        End If
    Else
        n_Tasa = adoresultTmp!TAS_INTE
        n_TasDia = Format(((1 + adoresultTmp!TAS_INTE) ^ (1 / adoresultTmp!CNT_DIAS)), "0.0000000000000000")
        If (s_Amort = "F" Or s_Amort = "V") Then
            n_Tasa = adoresultTmp!TAS_INTE + adoresultTmp!VAL_AMOR
        End If
    End If

    n_NewNominal = MontNomi

    Do While Not adoresultTmp.EOF
        's_fecha = Right$(adoresultTmp!FCH_VCTO, 2) + "/" + Mid$(adoresultTmp!FCH_VCTO, 5, 2) + "/" + Left$(adoresultTmp!FCH_VCTO, 4)
        s_fecha = CStr(Convertddmmyyyy(adoresultTmp!fch_vcto))
        d_FchFin = CVDate(s_fecha)
        n_CntDias = DateDiff("d", d_FchIni, d_FchFin)
        If TipBono = "P" Then
            If adoresultTmp!TAS_INTE2 = 0 Then
                If (s_Amort = "F" Or s_Amort = "V") Then
                    n_Tasa = ((n_TasDia ^ adoresultTmp!CNT_DIAS) - 1) + adoresultTmp!VAL_AMOR
                Else
                    n_Tasa = ((n_TasDia ^ adoresultTmp!CNT_DIAS) - 1)
                End If
                n_Monto = Format(MontNomi * n_Tasa, "0.00")
                If (s_Amort = "F" Or s_Amort = "V") Then
                    n_Monto = Format((n_NewNominal * n_Tasa) + (MontTitu * adoresultTmp!VAL_AMOR), "0.00")
                    n_NewNominal = n_NewNominal - (MontTitu * adoresultTmp!VAL_AMOR)
                End If
            Else
                n_Monto = Format(MontNomi * adoresultTmp!TAS_INTE2, "0.00")
                If (s_Amort = "F" Or s_Amort = "V") Then
                    n_Monto = Format((n_NewNominal * adoresultTmp!TAS_INTE2) + (MontTitu * adoresultTmp!VAL_AMOR), "0.00")
                End If
                n_Tasa = adoresultTmp!TAS_INTE2
                n_TasDia = Format(((1 + adoresultTmp!TAS_INTE2) ^ (1 / adoresultTmp!CNT_DIAS)), "0.0000000000000000")
                If (s_Amort = "F" Or s_Amort = "V") Then
                    n_Tasa = adoresultTmp!TAS_INTE + adoresultTmp!VAL_AMOR
                    n_NewNominal = n_NewNominal - (MontTitu * adoresultTmp!VAL_AMOR)
                End If
            End If
        Else
            If adoresultTmp!TAS_INTE = 0 Then
                If (s_Amort = "F" Or s_Amort = "V") Then
                    n_Tasa = ((n_TasDia ^ adoresultTmp!CNT_DIAS) - 1) + adoresultTmp!VAL_AMOR
                Else
                    n_Tasa = ((n_TasDia ^ adoresultTmp!CNT_DIAS) - 1)
                End If
                n_Monto = Format(MontNomi * n_Tasa, "0.00")
                If (s_Amort = "F" Or s_Amort = "V") Then
                    n_Monto = Format((n_NewNominal * n_Tasa) + (MontTitu * adoresultTmp!VAL_AMOR), "0.00")
                    n_NewNominal = n_NewNominal - (MontTitu * adoresultTmp!VAL_AMOR)
                End If
            Else
                n_Monto = Format(MontNomi * adoresultTmp!TAS_INTE, "0.00")
                If (s_Amort = "F" Or s_Amort = "V") Then
                    n_Monto = Format((n_NewNominal * adoresultTmp!TAS_INTE) + (MontTitu * adoresultTmp!VAL_AMOR), "0.00")
                End If
                n_Tasa = adoresultTmp!TAS_INTE
                n_TasDia = Format(((1 + adoresultTmp!TAS_INTE) ^ (1 / adoresultTmp!CNT_DIAS)), "0.0000000000000000")
                If (s_Amort = "F" Or s_Amort = "V") Then
                    n_Tasa = adoresultTmp!TAS_INTE + adoresultTmp!VAL_AMOR
                    n_NewNominal = n_NewNominal - (MontTitu * adoresultTmp!VAL_AMOR)
                End If
            End If
        End If
        If n_NroCupoFin = adoresultTmp!NRO_CUPO Then
            If (s_Amort = "F" Or s_Amort = "V") Then
                n_Monto = Format(n_Monto, "0.00")
            Else
                n_Monto = Format(n_Monto + MontNomi, "0.00")
            End If
        End If
        n_Cont = n_Cont + 1
        Array_Monto(n_Cont) = n_Monto: Array_Dias(n_Cont) = n_CntDias
        adoresultTmp.MoveNext
    Loop
        
    n_Acum = 0
    For i = 1 To n_Cont
        n_Acum = n_Acum + (Array_Monto(i) / ((1 + Tasa) ^ (Array_Dias(i) / 365)))
    Next

    adoresultTmp.Close: Set adoresultTmp = Nothing
    
    VNANoPerLetPag = Format(n_Acum, "0.00")
    
End Function

Public Function VNANoPer(ByVal strpCodTitulo As String, ByVal datpFechaOperacion As Date, ByVal datpFechaCupon As Date, ByVal curpCantidadNominal As Double, ByVal curpCantidadTitulos As Currency, ByVal dblpTirOPeracion As Double, ByVal strpTipoTitulo As String, ByVal strpCodIndiceInicial As String, ByVal strpCodIndiceFinal As String) As Double
'ByVal curpSubTotal As Currency, ByVal curpInteresCorrido As Currency,ByVal strpCodIndiceInicial As String, ByVal strpCodIndiceFinal As String
    Dim adoRegistro             As ADODB.Recordset
    Dim intContador             As Integer, intDiasPeriodo              As Integer
    Dim intNumCupon             As Integer, intNroCupoFin               As Integer
    Dim intCntDias              As Integer, intNumRegistros             As Integer
    Dim curMonto                As Double, curNewNominal              As Double
    Dim dblnTasa                As Double, dblTasDia                    As Double
    Dim dblTasa                 As Double, intAcum                      As Double
    Dim dblAjusteInicial        As Double, dblAjusteFinal               As Double
    Dim strIndAmortizacion      As String, strFecha                     As String
    Dim strTipoTasa             As String, strIndTasaReal               As String
    Dim strPeriodoPago          As String, strClaseTasa                 As String
    Dim strFechaInicialIndice   As String, strFechaInicialIndiceMas1    As String
    Dim strFechaFinalIndice     As String, strFechaFinalIndiceMas1      As String
    Dim datFechaInicio          As Date, datFechaFinal                  As Date
    Dim datFechaEmision         As Date, dblValorNominalOriginal        As Double
        
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT ValorNominal, IndAmortizacion,CodTipoTasa,CodTipoVac,IndReal,PeriodoPago,FechaEmision " & _
            "FROM InstrumentoInversion WHERE CodTitulo='" & strpCodTitulo & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strIndAmortizacion = Trim(adoRegistro("IndAmortizacion")): strTipoTasa = Trim(adoRegistro("CodTipoTasa"))
            strIndTasaReal = Trim(adoRegistro("IndReal")): strPeriodoPago = Trim(adoRegistro("PeriodoPago"))
            strClaseTasa = Trim(adoRegistro("CodTipoVac"))
            datFechaEmision = adoRegistro("FechaEmision")
            dblValorNominalOriginal = adoRegistro("ValorNominal")
        End If
        adoRegistro.Close

        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodParametro='" & strPeriodoPago & "' AND CodTipoParametro='TIPFRE'"
        Set adoRegistro = .Execute
        If Not adoRegistro.EOF Then
            intDiasPeriodo = CInt(adoRegistro("ValorParametro"))
        End If
        adoRegistro.Close
        
        .CommandText = "SELECT CONVERT(INT,NumCupon) NumCupon FROM InstrumentoInversionCalendario " & _
            "WHERE CodTitulo='" & strpCodTitulo & "' AND " & _
            "FechaInicio<='" & Convertyyyymmdd(datpFechaCupon) & "' AND " & _
            "FechaVencimiento>='" & Convertyyyymmdd(datpFechaCupon) & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            intNumCupon = adoRegistro("NumCupon")
        End If
        adoRegistro.Close
        
        .CommandText = "SELECT FechaVencimiento,FactorInteres,FactorInteres1,ValorAmortizacion,CantDiasPeriodo,NumCupon,FechaInicioIndice,FechaFinIndice,PorcenAmortizacion " & _
            "FROM InstrumentoInversionCalendario " & _
            "WHERE CodTitulo='" & strpCodTitulo & "' AND " & _
            "CONVERT(INT,NumCupon) >= " & intNumCupon & " ORDER BY NumCupon"
        adoRegistro.Open .CommandText, adoConn, adOpenStatic
        adoRegistro.MoveLast
        intNumRegistros = adoRegistro.RecordCount
        
        'ReDim Array_Monto(intNumRegistros + 1): ReDim Array_Dias(intNumRegistros + 1)
        ReDim Array_Monto(intNumRegistros): ReDim Array_Dias(intNumRegistros)
        
        intAcum = 1: dblTasa = dblpTirOPeracion: intContador = 0
        intNroCupoFin = CInt(adoRegistro("NumCupon"))
        
        adoRegistro.MoveFirst
        
        curMonto = 0
        Array_Monto(intContador) = curMonto
        strFecha = Convertyyyymmdd(datpFechaOperacion)
        datFechaInicio = datpFechaOperacion: Array_Dias(intContador) = datFechaInicio
        dblnTasa = 0
'        If strpTipoTitulo = Codigo_Vac_Periodico Then
'            dblnTasa = adoRegistro("FactorInteres1")
'            If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'                dblTasDia = ((1 + adoRegistro("FactorInteres1")) ^ (1 / adoRegistro("CantDiasPeriodo")))
'            Else
'                If strIndTasaReal = Valor_Indicador Then
'                    dblTasDia = adoRegistro("FactorInteres1") / intDiasPeriodo
'                Else
'                    dblTasDia = adoRegistro("FactorInteres1") / adoRegistro("CantDiasPeriodo")
'                End If
'            End If
'            If strIndAmortizacion = Valor_Indicador Then
'                dblnTasa = adoRegistro("FactorInteres1") + adoRegistro("ValorAmortizacion")
'            End If
'        Else
'            dblnTasa = adoRegistro("FactorInteres")
'            If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'                dblTasDia = ((1 + adoRegistro("FactorInteres")) ^ (1 / adoRegistro("CantDiasPeriodo")))
'            Else
'                If strIndTasaReal = Valor_Indicador Then
'                    dblTasDia = adoRegistro("FactorInteres") / intDiasPeriodo
'                Else
'                    dblTasDia = adoRegistro("FactorInteres") / adoRegistro("CantDiasPeriodo")
'                End If
'            End If
'            If strIndAmortizacion = Valor_Indicador Then
'                dblnTasa = adoRegistro("FactorInteres") + adoRegistro("ValorAmortizacion")
'            End If
'        End If

'        If strIndAmortizacion = Valor_Indicador Then
'            curNewNominal = curpCantidadTitulos
'        Else
            curNewNominal = curpCantidadNominal
       ' End If

        Do While Not adoRegistro.EOF
            'strFecha = Convertddmmyyyy(adoRegistro("FechaVencimiento"))
            strFechaInicialIndice = Convertyyyymmdd(adoRegistro("FechaInicioIndice"))
            strFechaInicialIndiceMas1 = Convertyyyymmdd(DateAdd("d", 1, adoRegistro("FechaInicioIndice")))
            If strpCodIndiceInicial = Codigo_Vac_Emision Then
                strFechaInicialIndice = Convertyyyymmdd(datFechaEmision)
                strFechaInicialIndiceMas1 = Convertyyyymmdd(DateAdd("d", 1, datFechaEmision))
            End If
            
            strFechaFinalIndice = Convertyyyymmdd(adoRegistro("FechaFinIndice"))
            strFechaFinalIndiceMas1 = Convertyyyymmdd(DateAdd("d", 1, adoRegistro("FechaFinIndice")))
            If strpCodIndiceFinal = Codigo_Vac_Liquidacion Then
                strFechaFinalIndice = Convertyyyymmdd(datpFechaOperacion)
                strFechaFinalIndiceMas1 = Convertyyyymmdd(DateAdd("d", 1, datpFechaOperacion))
            End If
            
            '*** Es tasa ajustada ? ***
            If strpTipoTitulo <> Valor_Caracter Then
                If strpTipoTitulo = Codigo_Tipo_Ajuste_Vac Then
                    dblAjusteInicial = ObtenerTasaAjuste(Codigo_Tipo_Ajuste_Vac, "00", strFechaInicialIndice, strFechaInicialIndiceMas1)
                    dblAjusteFinal = ObtenerTasaAjuste(Codigo_Tipo_Ajuste_Vac, "00", strFechaFinalIndice, strFechaFinalIndiceMas1)
                Else
                    dblAjusteInicial = ObtenerTasaAjuste(Codigo_Tipo_Ajuste_Vac, strClaseTasa, strFechaInicialIndice, strFechaInicialIndiceMas1)
                    dblAjusteFinal = 0
                End If
            End If
            
            datFechaFinal = adoRegistro("FechaVencimiento")
            intCntDias = DateDiff("d", datFechaInicio, datFechaFinal)
            If strpTipoTitulo = Codigo_Tipo_Ajuste_Vac Then
                If adoRegistro("FactorInteres1") = 0 Then
                    If strIndAmortizacion = Valor_Indicador Then
                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1) + adoRegistro("ValorAmortizacion")
                        Else
                            If strIndTasaReal = Valor_Indicador Then
                                dblnTasa = (dblTasDia * intDiasPeriodo)
                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
                                    dblnTasa = ((dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo) + adoRegistro("ValorAmortizacion")
                                Else
                                    dblnTasa = dblnTasa + adoRegistro("ValorAmortizacion")
                                End If
                            Else
                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo")) + adoRegistro("ValorAmortizacion")
                            End If
                        End If
                    Else
                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1)
                        Else
                            If strIndTasaReal = Valor_Indicador Then
                                dblnTasa = (dblTasDia * intDiasPeriodo)
                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
                                    dblnTasa = (dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo
                                End If
                            Else
                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo"))
                            End If
                        End If
                    End If
                    curMonto = curpCantidadNominal * dblnTasa
                    If strIndAmortizacion = Valor_Indicador Then
                        curMonto = (curNewNominal * dblnTasa) + (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                        curNewNominal = curNewNominal - (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                    End If
                Else
                    curMonto = curpCantidadNominal * adoRegistro("FactorInteres1") * (dblAjusteFinal / dblAjusteInicial)
                    If strIndAmortizacion = Valor_Indicador Then
                        curMonto = (curNewNominal * adoRegistro("FactorInteres1")) + (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                    End If
                    dblnTasa = adoRegistro("FactorInteres1")
                    If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                        dblTasDia = ((1 + adoRegistro("FactorInteres1")) ^ (1 / adoRegistro("CantDiasPeriodo")))
                    Else
                        If strIndTasaReal = Valor_Indicador Then
                            dblTasDia = adoRegistro("FactorInteres1") / intDiasPeriodo
                        Else
                            dblTasDia = adoRegistro("FactorInteres1") / adoRegistro("CantDiasPeriodo")
                        End If
                    End If
                    If strIndAmortizacion = Valor_Indicador Then
                        dblnTasa = adoRegistro("FactorInteres") + adoRegistro("ValorAmortizacion")
                        curNewNominal = curNewNominal - (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                    End If
                End If
            Else
                If adoRegistro("FactorInteres") = 0 Then
                    If strIndAmortizacion = Valor_Indicador Then
                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1) + adoRegistro("ValorAmortizacion")
                        Else
                            If strIndTasaReal = Valor_Indicador Then
                                dblnTasa = (dblTasDia * intDiasPeriodo)
                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
                                    dblnTasa = ((dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo) + adoRegistro("ValorAmortizacion")
                                Else
                                    dblnTasa = dblnTasa + adoRegistro("ValorAmortizacion")
                                End If
                            Else
                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo")) + adoRegistro("ValorAmortizacion")
                            End If
                        End If
                    Else
                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1)
                        Else
                            If strIndTasaReal = Valor_Indicador Then
                                dblnTasa = (dblTasDia * intDiasPeriodo)
                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
                                    dblnTasa = (dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo
                                End If
                            Else
                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo"))
                            End If
                        End If
                    End If
                    curMonto = Format(curpCantidadNominal * dblnTasa, "0.00")
                    If strIndAmortizacion = Valor_Indicador Then
                        curMonto = (curNewNominal * dblnTasa) + (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                        curNewNominal = curNewNominal - (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                    End If
                Else
                    curMonto = curpCantidadNominal * adoRegistro("FactorInteres1")
                    If strIndAmortizacion = Valor_Indicador Then
'                        curMonto = (curNewNominal * adoRegistro("FactorInteres")) + (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                        curMonto = (curNewNominal * adoRegistro("FactorInteres")) + (curpCantidadTitulos * dblValorNominalOriginal * adoRegistro("PorcenAmortizacion") * 0.01)
                    End If
                    dblnTasa = adoRegistro("FactorInteres1")
                    If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                        dblTasDia = ((1 + adoRegistro("FactorInteres")) ^ (1 / adoRegistro("CantDiasPeriodo")))
                    Else
                        If strIndTasaReal = Valor_Indicador Then
                            dblTasDia = adoRegistro("FactorInteres") / intDiasPeriodo
                        Else
                            dblTasDia = adoRegistro("FactorInteres1") / adoRegistro("CantDiasPeriodo")
                        End If
                    End If
                    If strIndAmortizacion = Valor_Indicador Then
                        dblnTasa = adoRegistro("FactorInteres") + adoRegistro("ValorAmortizacion")
'                        curNewNominal = curNewNominal - (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                        curNewNominal = curNewNominal - (curpCantidadTitulos * dblValorNominalOriginal * adoRegistro("PorcenAmortizacion") * 0.01)
                    End If
                End If
            End If
            If intNroCupoFin = adoRegistro("NumCupon") Then
                If strIndAmortizacion = Valor_Indicador Then
                    curMonto = curMonto
                Else
                    curMonto = curMonto + curpCantidadNominal
                    If strpTipoTitulo = Codigo_Tipo_Ajuste_Vac Then
                        curMonto = (curpCantidadNominal * adoRegistro("FactorInteres") + curpCantidadNominal) * (dblAjusteFinal / dblAjusteInicial)
                    End If
                End If
            End If
            intContador = intContador + 1
            Array_Monto(intContador) = curMonto: Array_Dias(intContador) = datFechaFinal
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing

        VNANoPer = VAN(Array_Monto(), Array_Dias(), dblTasa * 0.01, 0)
    End With

'    n_Acum = 0
'    For i = 1 To n_Cont
'        n_Acum = n_Acum + (Array_Monto(i) / ((1 + Tasa) ^ (Array_Dias(i) / 365)))
'    Next
'
'    adoResultAux.Close: Set adoResultAux = Nothing
    'VNANoPer = Format(n_Acum, "0.00")
    
End Function

'Public Function VNANoPer1(ByVal strpCodTitulo As String, ByVal datpFechaOperacion As Date, ByVal datpFechaCupon As Date, ByVal curpCantidadNominal As Double, ByVal curpCantidadTitulos As Currency, ByVal dblpTirOPeracion As Double, ByVal strpTipoTitulo As String, ByVal strpCodIndiceInicial As String, ByVal strpCodIndiceFinal As String) As Double
'
'    Dim adoRegistro                 As ADODB.Recordset
'    Dim intContador                 As Integer
'    Dim intNumCupon                 As Integer
'    Dim intNumRegistros             As Integer
'    Dim curMonto                    As Currency, curNewNominal          As Currency
'    Dim strFecha                    As String, datFechaInicio           As Date
'    Dim datFechaFinal               As Date
'    Dim dblValorNominalOriginal     As Double
'    Set adoRegistro = New ADODB.Recordset
'    With adoComm
'
'        'OBTENER EL CUPON VIGENTE
'        .CommandText = "SELECT CONVERT(INT,NumCupon) NumCupon FROM InstrumentoInversionCalendario " & _
'            "WHERE CodTitulo='" & strpCodTitulo & "' AND " & _
'            "FechaInicio<='" & Convertyyyymmdd(datpFechaCupon) & "' AND " & _
'            "FechaVencimiento>='" & Convertyyyymmdd(datpFechaCupon) & "'"
'        Set adoRegistro = .Execute
'
'        If Not adoRegistro.EOF Then
'            intNumCupon = adoRegistro("NumCupon")
'        End If
'        adoRegistro.Close
'
'        'CALCULAR CURSOR CON CUPONES POR VENCER
'        .CommandText = "SELECT FechaVencimiento,ValorCupon FROM InstrumentoInversionCalendario " & _
'            "WHERE CodTitulo='" & strpCodTitulo & "' AND " & _
'            "CONVERT(INT,NumCupon) >= " & intNumCupon & " ORDER BY NumCupon"
'        adoRegistro.Open .CommandText, adoConn, adOpenStatic
'
'        adoRegistro.MoveLast
'
'        intNumRegistros = adoRegistro.RecordCount
'
'        ReDim Array_Monto(intNumRegistros): ReDim Array_Dias(intNumRegistros)
'
'        intContador = 0
'
'        adoRegistro.MoveFirst
'
'        curMonto = (curpSubTotal + curpInteresCorrido) * -1
'        Array_Monto(intContador) = curMonto
'        strFecha = Convertyyyymmdd(datpFechaLiquidacion)
'        datFechaInicio = datpFechaLiquidacion: Array_Dias(intContador) = datFechaInicio
'        curNewNominal = curpCantidadNominal
'
'        Do While Not adoRegistro.EOF
'
'            datFechaFinal = adoRegistro("FechaVencimiento")
'            curMonto = adoRegistro("ValorCupon") * curpCantidadTitulos
'            intContador = intContador + 1
'            Array_Monto(intContador) = curMonto: Array_Dias(intContador) = datFechaFinal
'            adoRegistro.MoveNext
'
'        Loop
'        adoRegistro.Close: Set adoRegistro = Nothing
'
'        VNANoPer = VAN(Array_Monto(), Array_Dias(), dblTasa * 0.01, 0)
'    End With
'
'End Function
Function VNANoPerPlazo(strpCodTitulo As String, datpFechaOperacion As Date, datpFechaCupon As Date, datpFechaVenta As Date, curpCantidadNominal As Currency, curpCantidadTitulos As Currency, dblpTirOPeracion As Double, strpTipoTitulo As String, curpTotalPlazo As Currency) As Double

    Dim adoRegistro         As ADODB.Recordset
    Dim intContador         As Integer, intDiasPeriodo  As Integer
    Dim intNumCupon         As Integer, intNroCupoFin   As Integer
    Dim intCntDias          As Integer, intNumRegistros As Integer
    Dim curMonto            As Currency, curNewNominal  As Currency
    Dim dblnTasa            As Double, dblTasDia        As Double
    Dim dblTasa             As Double, intAcum          As Double
    Dim strIndAmortizacion  As String, strFecha         As String
    Dim strTipoTasa         As String, strIndTasaReal   As String
    Dim strPeriodoPago      As String
    Dim datFechaInicio      As Date, datFechaFinal      As Date
        
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT IndAmortizacion,CodTipoTasa,IndReal,PeriodoPago " & _
            "FROM InstrumentoInversion WHERE CodTitulo='" & strpCodTitulo & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strIndAmortizacion = Trim(adoRegistro("IndAmortizacion")): strTipoTasa = Trim(adoRegistro("CodTipoTasa"))
            strIndTasaReal = Trim(adoRegistro("IndReal")): strPeriodoPago = Trim(adoRegistro("PeriodoPago"))
        End If
        adoRegistro.Close

        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodParametro='" & strPeriodoPago & "' AND CodTipoParametro='TIPFRE'"
        Set adoRegistro = .Execute
        If Not adoRegistro.EOF Then
            intDiasPeriodo = CInt(adoRegistro("ValorParametro"))
        End If
        adoRegistro.Close
        
        .CommandText = "SELECT CONVERT(INT,NumCupon) NumCupon FROM InstrumentoInversionCalendario " & _
            "WHERE CodTitulo='" & strpCodTitulo & "' AND " & _
            "FechaInicio<='" & Convertyyyymmdd(datpFechaCupon) & "' AND " & _
            "FechaVencimiento>='" & Convertyyyymmdd(datpFechaCupon) & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            intNumCupon = adoRegistro("NumCupon")
        End If
        adoRegistro.Close
        
        .CommandText = "SELECT FechaVencimiento,FactorInteres,FactorInteres1,ValorAmortizacion,CantDiasPeriodo,NumCupon " & _
            "FROM InstrumentoInversionCalendario " & _
            "WHERE CodTitulo='" & strpCodTitulo & "' AND " & _
            "CONVERT(INT,NumCupon) >= " & intNumCupon & " ORDER BY NumCupon"
        adoRegistro.Open .CommandText, adoConn, adOpenStatic
        adoRegistro.MoveLast
        intNumRegistros = adoRegistro.RecordCount
        
        'ReDim Array_Monto(intNumRegistros + 1): ReDim Array_Dias(intNumRegistros + 1)
        ReDim Array_Monto(intNumRegistros): ReDim Array_Dias(intNumRegistros)
        
        intAcum = 1: dblTasa = dblpTirOPeracion: intContador = 0
        intNroCupoFin = CInt(adoRegistro("NumCupon"))
        
        adoRegistro.MoveFirst
        
        curMonto = 0
        Array_Monto(intContador) = curMonto
        strFecha = Convertyyyymmdd(datpFechaOperacion)
        datFechaInicio = datpFechaOperacion: Array_Dias(intContador) = datFechaInicio
        dblnTasa = 0
'        If strpTipoTitulo = Codigo_Vac_Periodico Then
'            dblnTasa = adoRegistro("FactorInteres1")
'            If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'                dblTasDia = ((1 + adoRegistro("FactorInteres1")) ^ (1 / adoRegistro("CantDiasPeriodo")))
'            Else
'                If strIndTasaReal = Valor_Indicador Then
'                    dblTasDia = adoRegistro("FactorInteres1") / intDiasPeriodo
'                Else
'                    dblTasDia = adoRegistro("FactorInteres1") / adoRegistro("CantDiasPeriodo")
'                End If
'            End If
'            If strIndAmortizacion = Valor_Indicador Then
'                dblnTasa = adoRegistro("FactorInteres1") + adoRegistro("ValorAmortizacion")
'            End If
'        Else
            dblnTasa = adoRegistro("FactorInteres")
            If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                dblTasDia = ((1 + adoRegistro("FactorInteres")) ^ (1 / adoRegistro("CantDiasPeriodo")))
            Else
                If strIndTasaReal = Valor_Indicador Then
                    dblTasDia = adoRegistro("FactorInteres") / intDiasPeriodo
                Else
                    dblTasDia = adoRegistro("FactorInteres") / adoRegistro("CantDiasPeriodo")
                End If
            End If
            If strIndAmortizacion = Valor_Indicador Then
                dblnTasa = adoRegistro("FactorInteres") + adoRegistro("ValorAmortizacion")
            End If
'        End If

        curNewNominal = curpCantidadNominal

        Do While Not adoRegistro.EOF
            'strFecha = Convertddmmyyyy(adoRegistro("FechaVencimiento"))
            datFechaFinal = adoRegistro("FechaVencimiento")
            intCntDias = DateDiff("d", datFechaInicio, datFechaFinal)
'            If strpTipoTitulo = Codigo_Vac_Periodico Then
'                If adoRegistro("FactorInteres1") = 0 Then
'                    If strIndAmortizacion = Valor_Indicador Then
'                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1) + adoRegistro("ValorAmortizacion")
'                        Else
'                            If strIndTasaReal = Valor_Indicador Then
'                                dblnTasa = (dblTasDia * intDiasPeriodo)
'                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
'                                    dblnTasa = ((dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo) + adoRegistro("ValorAmortizacion")
'                                Else
'                                    dblnTasa = dblnTasa + adoRegistro("ValorAmortizacion")
'                                End If
'                            Else
'                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo")) + adoRegistro("ValorAmortizacion")
'                            End If
'                        End If
'                    Else
'                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1)
'                        Else
'                            If strIndTasaReal = Valor_Indicador Then
'                                dblnTasa = (dblTasDia * intDiasPeriodo)
'                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
'                                    dblnTasa = (dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo
'                                End If
'                            Else
'                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo"))
'                            End If
'                        End If
'                    End If
'                    curMonto = curpCantidadNominal * dblnTasa
'                    If strIndAmortizacion = Valor_Indicador Then
'                        curMonto = (curNewNominal * dblnTasa) + (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
'                        curNewNominal = curNewNominal - (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
'                    End If
'                Else
'                    curMonto = curpCantidadNominal * adoRegistro("FactorInteres1")
'                    If strIndAmortizacion = Valor_Indicador Then
'                        curMonto = (curNewNominal * adoRegistro("FactorInteres1")) + (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
'                    End If
'                    dblnTasa = adoRegistro("FactorInteres1")
'                    If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'                        dblTasDia = ((1 + adoRegistro("FactorInteres1")) ^ (1 / adoRegistro("CantDiasPeriodo")))
'                    Else
'                        If strIndTasaReal = Valor_Indicador Then
'                            dblTasDia = adoRegistro("FactorInteres1") / intDiasPeriodo
'                        Else
'                            dblTasDia = adoRegistro("FactorInteres1") / adoRegistro("CantDiasPeriodo")
'                        End If
'                    End If
'                    If strIndAmortizacion = Valor_Indicador Then
'                        dblnTasa = adoRegistro("FactorInteres") + adoRegistro("ValorAmortizacion")
'                        curNewNominal = curNewNominal - (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
'                    End If
'                End If
'            Else
                If adoRegistro("FactorInteres") = 0 Then
                    If strIndAmortizacion = Valor_Indicador Then
                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1) + adoRegistro("ValorAmortizacion")
                        Else
                            If strIndTasaReal = Valor_Indicador Then
                                dblnTasa = (dblTasDia * intDiasPeriodo)
                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
                                    dblnTasa = ((dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo) + adoRegistro("ValorAmortizacion")
                                Else
                                    dblnTasa = dblnTasa + adoRegistro("ValorAmortizacion")
                                End If
                            Else
                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo")) + adoRegistro("ValorAmortizacion")
                            End If
                        End If
                    Else
                        If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                            dblnTasa = ((dblTasDia ^ adoRegistro("CantDiasPeriodo")) - 1)
                        Else
                            If strIndTasaReal = Valor_Indicador Then
                                dblnTasa = (dblTasDia * intDiasPeriodo)
                                If adoRegistro("CantDiasPeriodo") < intDiasPeriodo Then
                                    dblnTasa = (dblnTasa * adoRegistro("CantDiasPeriodo")) / intDiasPeriodo
                                End If
                            Else
                                dblnTasa = (dblTasDia * adoRegistro("CantDiasPeriodo"))
                            End If
                        End If
                    End If
                    curMonto = Format(curpCantidadNominal * dblnTasa, "0.00")
                    If strIndAmortizacion = Valor_Indicador Then
                        curMonto = (curNewNominal * dblnTasa) + (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                        curNewNominal = curNewNominal - (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                    End If
                Else
                    curMonto = curpCantidadNominal * adoRegistro("FactorInteres")
                    If strIndAmortizacion = Valor_Indicador Then
                        curMonto = (curNewNominal * adoRegistro("FactorInteres")) + (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                    End If
                    dblnTasa = adoRegistro("FactorInteres")
                    If strTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                        dblTasDia = ((1 + adoRegistro("FactorInteres")) ^ (1 / adoRegistro("CantDiasPeriodo")))
                    Else
                        If strIndTasaReal = Valor_Indicador Then
                            dblTasDia = adoRegistro("FactorInteres") / intDiasPeriodo
                        Else
                            dblTasDia = adoRegistro("FactorInteres") / adoRegistro("CantDiasPeriodo")
                        End If
                    End If
                    If strIndAmortizacion = Valor_Indicador Then
                        dblnTasa = adoRegistro("FactorInteres") + adoRegistro("ValorAmortizacion")
                        curNewNominal = curNewNominal - (curpCantidadTitulos * adoRegistro("ValorAmortizacion"))
                    End If
                End If
'            End If
            If intNroCupoFin = adoRegistro("NumCupon") Then
                If strIndAmortizacion = Valor_Indicador Then
                    curMonto = Round(curMonto, 2)
                Else
                    curMonto = Round(curMonto + curpCantidadNominal, 2)
                End If
            End If
            intContador = intContador + 1
            Array_Monto(intContador) = curMonto: Array_Dias(intContador) = datFechaFinal
            
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
        
        intContador = intContador + 1
        curMonto = curpTotalPlazo
        datFechaFinal = datpFechaVenta
        Array_Monto(intContador) = curMonto: Array_Dias(intContador) = datFechaFinal

        VNANoPerPlazo = VAN(Array_Monto(), Array_Dias(), dblTasa * 0.01, 0)
    End With
    
End Function
Public Function VAN(ByRef vntValFlujo(), ByRef vntFchFlujo(), ByVal dblpValorTir As Double, ByVal intpTipoFuncion As Integer) As Double

    '*** Calcular el VAN tomando en cuenta los flujos futuros ***
    Dim intNumDias      As Integer, intContador As Integer
    Dim intBaseAnual    As Integer
    Dim dblValorRetorno As Double

    On Error GoTo Errorhandler

    dblValorRetorno = 0

    If IsNull(intpTipoFuncion) Then
        intpTipoFuncion = 0
    End If

    '*** Base Anual ***
    intBaseAnual = 360

    '*** Función de Valor Actual ***
    For intContador = 0 To UBound(vntValFlujo)
        intNumDias = DateDiff("d", vntFchFlujo(0), vntFchFlujo(intContador))
        If intpTipoFuncion = 0 Then
            dblValorRetorno = dblValorRetorno + (vntValFlujo(intContador) / ((1 + dblpValorTir) ^ (intNumDias / intBaseAnual)))
        Else
            dblValorRetorno = dblValorRetorno + (intNumDias * vntValFlujo(intContador) / ((1 + dblpValorTir) ^ (intNumDias / intBaseAnual)))
        End If
    Next

    VAN = dblValorRetorno

ExitFunction:
    Exit Function

Errorhandler:
    MsgBox Error(err), 48
    dblValorRetorno = 0
    Resume ExitFunction

End Function

Public Function UltimoDiaMes(ByVal intmes As Integer, ByVal intPeriodo As Integer) As Integer

    Dim intValorRetorno As Integer
    
    intValorRetorno = 0
    Select Case intmes
        Case 1, 3, 5, 7, 8, 10, 12
            intValorRetorno = 31
        Case 4, 6, 9, 11
            intValorRetorno = 30
        Case 2
            intValorRetorno = 28
            If intPeriodo Mod 4 = 0 And intPeriodo Mod 100 = 0 Then
               If intPeriodo Mod 400 = 0 Then
                  intValorRetorno = 29
               End If
            ElseIf intPeriodo Mod 4 = 0 Then
               intValorRetorno = 29
            End If
    End Select

    UltimoDiaMes = intValorRetorno
    
End Function

Public Function UltimaFechaMes(ByVal intmes As Integer, ByVal intPeriodo As Integer) As Date

    Dim intDia      As Integer
    Dim datFecha    As Date
    
    intDia = 0
    Select Case intmes
        Case 1, 3, 5, 7, 8, 10, 12
            intDia = 31
        Case 4, 6, 9, 11
            intDia = 30
        Case 2
            intDia = 28
            If intPeriodo Mod 4 = 0 And intPeriodo Mod 100 = 0 Then
               If intPeriodo Mod 400 = 0 Then
                  intDia = 29
               End If
            ElseIf intPeriodo Mod 4 = 0 Then
               intDia = 29
            End If
    End Select
    
    datFecha = Convertddmmyyyy(Format(intPeriodo, "0000") & Format(intmes, "00") & Format(intDia, "00"))

    UltimaFechaMes = datFecha
    
End Function
Public Sub CentrarForm(ByVal Forma As Form)
    
    Forma.Left = (frmMainMdi.Width - Forma.Width) / 2
    Forma.Top = (frmMainMdi.Height - frmMainMdi.tlbMdi.Height - frmMainMdi.stbMdi.Height - 780 - Forma.Height) / 2
    
End Sub

Function strtran(cCad As String, cBus As String, cRep As String) As String
'Reemplaza una cadena en otra, no necesariamente de la misma longitud
'Si no encuentra la cadena buscada retorna cCad
    Dim nPos As Integer, cAcumCad As String, nLen As Integer
    Dim cSubStr1 As String, cSubStr2 As String
    
    nPos = InStr(cCad, cBus)
    If nPos > 0 Then
        cAcumCad = ""
        Do While nPos > 0
            cSubStr1 = Left$(cCad, nPos - 1)
            nLen = Len(cCad) - (nPos + Len(cBus) - 1)
            cSubStr2 = Right$(cCad, nLen)
            'StrTran = cSubStr1$ + cRep + cSubStr2$
            cAcumCad = cAcumCad + cSubStr1 + cRep '+ cSubStr2
            cCad = cSubStr2
            nPos = InStr(cCad, cBus)
        Loop
        strtran = cAcumCad + cSubStr2
    Else
        strtran = cCad
    End If
    
End Function

Public Function ObtenerItemLista(ByRef arrControl() As String, ByVal strItem As String, Optional intPosIni As Integer = 0, Optional intPosFin As Integer = 0) As Integer

    Dim intItems As Integer
    
    On Error GoTo Ctrl_Error
        
    If intPosIni = 0 Then intPosIni = 1
    
    intItems = UBound(arrControl)
    ObtenerItemLista = -1
    For intItems = 0 To UBound(arrControl)

        If Mid(Trim(arrControl(intItems)), intPosIni, IIf(intPosFin = 0, Len(arrControl(intItems)) - intPosIni + 1, intPosFin)) = Trim(strItem) Then
            ObtenerItemLista = intItems
            Exit For
        End If
    Next
    
Exit Function

Ctrl_Error:
    If err.Number = 9 Then
        ObtenerItemLista = -1
    Else
        MsgBox "Error Inesperado...", vbCritical, "Error"
    End If

End Function

Public Function EsBisiesto(ByVal intAnno As Integer) As Integer
   
    '*** Indicar si el año que se indica es bisiesto o no
    Dim vntCurYear As Variant, intRetVal As Integer
   
    intRetVal = False

    vntCurYear = intAnno
    If vntCurYear Mod 4 = 0 Then
        'If vntCurYear Mod 400 = 0 Then
        '    intRetVal = False
        'Else
            intRetVal = True
        'End If
    Else
        intRetVal = False
    End If
    EsBisiesto = intRetVal
   
End Function
Public Function ObtenerMontoArbitraje(dblMonto As Double, dblValorTipoCambio As Double, strCodMonedaParEvaluacion As String, strCodMonedaParPorDefecto As String)

    Dim dblMontoArbitraje As Double

    'Modificado de acuerdo a los últimos cambios (pardefecto)
    If dblValorTipoCambio <> 0 Then
        If strCodMonedaParEvaluacion = strCodMonedaParPorDefecto Then
            dblMontoArbitraje = dblMonto * dblValorTipoCambio
        Else
            dblMontoArbitraje = dblMonto / dblValorTipoCambio
        End If
    Else
        dblMontoArbitraje = 0#
    End If

    ObtenerMontoArbitraje = dblMontoArbitraje

End Function
Public Function ObtenerTipoCambioArbitraje(dblValorTipoCambio As Double, strCodMonedaParEvaluacion As String, strCodMonedaParPorDefecto As String) As Double
 
    Dim dblTipoCambioArbitraje As Double

    'Modificado de acuerdo a los últimos cambios (pardefecto)
    If dblValorTipoCambio <> 0 Then
        If strCodMonedaParEvaluacion = strCodMonedaParPorDefecto Then
            dblTipoCambioArbitraje = dblValorTipoCambio
        Else
            dblTipoCambioArbitraje = 1 / dblValorTipoCambio
        End If
    Else
        dblTipoCambioArbitraje = 0#
    End If

    ObtenerTipoCambioArbitraje = dblTipoCambioArbitraje

End Function
Public Function ObtenerMonedaParPorDefecto(ByVal strpCodTipoCambio As String, ByVal strCodMonedaParEvaluacion As String) As String


    Dim strCodMonedaPar As String
    
    ObtenerMonedaParPorDefecto = strCodMonedaParEvaluacion
    
    With adoComm
        '*** Obtener Secuencial ***
        .CommandType = adCmdStoredProc
        
        .CommandText = "up_ACObtenerMonedaParPorDefecto"
        .Parameters.Append .CreateParameter("@CodTipoCambio", adChar, adParamInput, 2, strpCodTipoCambio)
        .Parameters.Append .CreateParameter("@CodMonedaParEvaluacion", adChar, adParamInput, 4, strCodMonedaParEvaluacion)
        .Parameters.Append .CreateParameter("@CodMonedaParPorDefecto", adChar, adParamOutput, 4, strCodMonedaParEvaluacion)
        .Execute
        
        If Not .Parameters("@CodMonedaParPorDefecto") Then
            strCodMonedaPar = .Parameters("@CodMonedaParPorDefecto").Value
        End If
        
        .Parameters.Delete ("@CodTipoCambio")
        .Parameters.Delete ("@CodMonedaParEvaluacion")
        .Parameters.Delete ("@CodMonedaParPorDefecto")
                        
        '.Parameters.Refresh
        
        .CommandType = adCmdText
    End With
    
    ObtenerMonedaParPorDefecto = strCodMonedaPar

End Function

Public Sub CargarControlLista(ByVal strSentencia As String, ByVal CtrlNombre As Control, ByRef arrControl() As String, ByVal strValor As String, Optional blnActivaItemData As Boolean = False)

    Dim adoBusqueda As ADODB.Recordset
    Dim intCont As Long
    
    Set adoBusqueda = New ADODB.Recordset
    
    adoComm.CommandText = strSentencia
    Set adoBusqueda = adoComm.Execute

    CtrlNombre.Clear
    intCont = 0
    ReDim arrControl(intCont)
    
    If Len(strValor) > 0 Then
        CtrlNombre.AddItem strValor
        intCont = 0
        ReDim Preserve arrControl(intCont)
        arrControl(intCont) = Valor_Caracter
        intCont = 1
    End If
    
    Do Until adoBusqueda.EOF
        CtrlNombre.AddItem adoBusqueda("DESCRIP")
        ReDim Preserve arrControl(intCont)
        arrControl(intCont) = adoBusqueda("CODIGO")
        If blnActivaItemData Then
            CtrlNombre.ItemData(CtrlNombre.NewIndex) = IIf(adoBusqueda("CODIGO") = Valor_Caracter Or Not IsNumeric(adoBusqueda("CODIGO")), 0, adoBusqueda("CODIGO"))
        End If
        adoBusqueda.MoveNext
        intCont = intCont + 1
    Loop
   
    adoBusqueda.Close: Set adoBusqueda = Nothing

End Sub

Public Sub CargarControlListaAuxiliarParametro(ByVal strCodParametro As String, ByVal CtrlNombre As Control, ByRef arrControl() As String, ByVal strValor As String)

    Dim adoBusqueda As ADODB.Recordset
    Dim intCont As Long
    
    Set adoBusqueda = New ADODB.Recordset
    
    adoComm.CommandText = "EXEC up_ACLstAuxiliarParametro '" & strCodParametro & "'"
    Set adoBusqueda = adoComm.Execute

    CtrlNombre.Clear
    intCont = 0
    ReDim arrControl(intCont)
    
    If Len(strValor) > 0 Then
        CtrlNombre.AddItem strValor
        intCont = 0
        ReDim Preserve arrControl(intCont)
        arrControl(intCont) = Valor_Caracter
        intCont = 1
    End If
     
    Do Until adoBusqueda.EOF
        CtrlNombre.AddItem adoBusqueda("DESCRIP")
        ReDim Preserve arrControl(intCont)
        arrControl(intCont) = adoBusqueda("CODIGO")
        adoBusqueda.MoveNext
        intCont = intCont + 1
    Loop
   
    adoBusqueda.Close: Set adoBusqueda = Nothing

End Sub

Public Function GetINIString(ByVal szItem As String, ByVal szDefault As String, ByVal szGrup As String) As String

    Dim tmp As String
    Dim X As Integer
    Dim cFlag As String

    tmp = String$(2048, 32)
    If cFlag = "X" Then
        X = GetPrivateProfileString(szGrup, szItem, szDefault, tmp, Len(tmp), "WIN.INI")
    Else
        X = GetPrivateProfileString(szGrup, szItem, szDefault, tmp, Len(tmp), "fondos.ini")
    End If

    GetINIString = Mid$(tmp, 1, X)
    
End Function

Function TirNoPerLetPag(Codfile As String, CodAnal As String, strCodFon As String, FecOpe, FecCup, SubTot As Double, IntCorr As Double, MontNomi As Double, MontTitu As Double, TirOpe As Double, TipBono As String) As Double
    Dim adoresult As New ADODB.Recordset
    Dim n_Cont As Integer
    Dim n_Monto As Double, n_NroCupo As Integer, s_fecha As String * 10
    Dim n_Tasa As Double, s_Amort As String, n_NewNominal As Double, n_TasDia As Double
    Dim n_Acum As Double, Tasa As Double, d_FchIni, d_FchFin, n_CntDias As Integer, n_NroCupoFin As Integer
    Dim n_Reg As Integer, i As Integer
    
    With adoComm
        .CommandText = "SELECT FLG_AMORT FROM FMLETRAS WHERE COD_FILE='" & Trim$(Codfile) & "' AND COD_ANAL='" & Trim$(CodAnal) & "' AND COD_FOND='" & strCodFon & "'"
        Set adoresult = .Execute
        s_Amort = adoresult!FLG_AMORT
        adoresult.Close: Set adoresult = Nothing
    
        .CommandText = "SELECT CONVERT(INT,NRO_CUPO) NRO_CUPO FROM FMCUPONES"
        .CommandText = .CommandText + " WHERE COD_FILE='" & Trim$(Codfile) & "' AND COD_ANAL='" & Trim$(CodAnal) & "' AND COD_FOND='" & strCodFon & "'"
        '.CommandText = .CommandText + " AND FCH_INIC<='" & Format(FecCup, "yyyymmdd") & "'"
        .CommandText = .CommandText + " AND FCH_INIC<='" & Convertyyyymmdd(FecCup) & "'"
        '.CommandText = .CommandText + " AND FCH_VCTO>='" & Format(FecCup, "yyyymmdd") & "'"
        .CommandText = .CommandText + " AND FCH_VCTO>='" & Convertyyyymmdd(FecCup) & "'"
        Set adoresult = .Execute
        n_NroCupo = adoresult!NRO_CUPO
        adoresult.Close: Set adoresult = Nothing
    
        .CommandText = "SELECT FCH_VCTO,TAS_INTE,TAS_INTE2,VAL_AMOR,CNT_DIAS,NRO_CUPO FROM FMCUPONES"
        .CommandText = .CommandText & " WHERE COD_FILE='" & Trim$(Codfile) & "' AND COD_ANAL='" & Trim$(CodAnal) & "' AND COD_FOND='" & strCodFon & "'"
        .CommandText = .CommandText & " AND CONVERT(INT,NRO_CUPO) >= " & n_NroCupo & " ORDER BY NRO_CUPO"
        adoresult.Open .CommandText, adoConn, adOpenStatic
        'adoresult.MoveLast
        n_Reg = adoresult.RecordCount
    
        ReDim Array_Monto(n_Reg + 1): ReDim Array_Dias(n_Reg + 1)
        ReDim Array_Monto(n_Reg): ReDim Array_Dias(n_Reg)
        n_Acum = 1: Tasa = TirOpe: n_Cont = 0
        n_NroCupoFin = adoresult!NRO_CUPO
        'adoresult.MoveFirst
    
        n_Monto = Format((SubTot + IntCorr) * -1, "0.00")
        Array_Monto(n_Cont) = n_Monto
        's_fecha = Trim$(Format(FecOpe, "yyyymmdd"))
        s_fecha = Convertyyyymmdd(FecOpe)
        s_fecha = CStr(Convertddmmyyyy(s_fecha))
        d_FchIni = CVDate(s_fecha): Array_Dias(n_Cont) = d_FchIni
        n_Tasa = 0
        If TipBono = "P" Then
            n_Tasa = adoresult!TAS_INTE2
            n_TasDia = Format(((1 + adoresult!TAS_INTE2) ^ (1 / adoresult!CNT_DIAS)), "0.0000000000000000")
            If (s_Amort = "F" Or s_Amort = "V") Then
                n_Tasa = adoresult!TAS_INTE2 + adoresult!VAL_AMOR
            End If
        Else
            n_Tasa = adoresult!TAS_INTE
            n_TasDia = Format(((1 + adoresult!TAS_INTE) ^ (1 / adoresult!CNT_DIAS)), "0.0000000000000000")
            If (s_Amort = "F" Or s_Amort = "V") Then
                n_Tasa = adoresult!TAS_INTE + adoresult!VAL_AMOR
            End If
        End If
    
        n_NewNominal = MontNomi
    
        Do While Not adoresult.EOF
            's_fecha = Right$(adoresult!FCH_VCTO, 2) & "/" & Mid$(adoresult!FCH_VCTO, 5, 2) & "/" & Left$(adoresult!FCH_VCTO, 4)
            s_fecha = CStr(Convertddmmyyyy(adoresult!fch_vcto))
            d_FchFin = CVDate(s_fecha)
            n_CntDias = DateDiff("d", d_FchIni, d_FchFin)
            If TipBono = "P" Then
                If adoresult!TAS_INTE2 = 0 Then
                    If (s_Amort = "F" Or s_Amort = "V") Then
                        n_Tasa = ((n_TasDia ^ adoresult!CNT_DIAS) - 1) + adoresult!VAL_AMOR
                    Else
                        n_Tasa = ((n_TasDia ^ adoresult!CNT_DIAS) - 1)
                    End If
                    n_Monto = Format(MontNomi * n_Tasa, "0.00")
                    If (s_Amort = "F" Or s_Amort = "V") Then
                        n_Monto = Format((n_NewNominal * n_Tasa) + (MontTitu * adoresult!VAL_AMOR), "0.00")
                        n_NewNominal = n_NewNominal - (MontTitu * adoresult!VAL_AMOR)
                    End If
                Else
                    n_Monto = Format(MontNomi * adoresult!TAS_INTE2, "0.00")
                    If (s_Amort = "F" Or s_Amort = "V") Then
                        n_Monto = Format((n_NewNominal * adoresult!TAS_INTE2) + (MontTitu * adoresult!VAL_AMOR), "0.00")
                    End If
                    n_Tasa = adoresult!TAS_INTE2
                    n_TasDia = Format(((1 + adoresult!TAS_INTE2) ^ (1 / adoresult!CNT_DIAS)), "0.0000000000000000")
                    If (s_Amort = "F" Or s_Amort = "V") Then
                        n_Tasa = adoresult!TAS_INTE + adoresult!VAL_AMOR
                        n_NewNominal = n_NewNominal - (MontTitu * adoresult!VAL_AMOR)
                    End If
                End If
            Else
                If adoresult!TAS_INTE = 0 Then
                    If (s_Amort = "F" Or s_Amort = "V") Then
                        n_Tasa = ((n_TasDia ^ adoresult!CNT_DIAS) - 1) + adoresult!VAL_AMOR
                    Else
                        n_Tasa = ((n_TasDia ^ adoresult!CNT_DIAS) - 1)
                    End If
                    n_Monto = Format(MontNomi * n_Tasa, "0.00")
                    If (s_Amort = "F" Or s_Amort = "V") Then
                        n_Monto = Format((n_NewNominal * n_Tasa) + (MontTitu * adoresult!VAL_AMOR), "0.00")
                        n_NewNominal = n_NewNominal - (MontTitu * adoresult!VAL_AMOR)
                    End If
                Else
                    n_Monto = Format(MontNomi * adoresult!TAS_INTE, "0.00")
                    If (s_Amort = "F" Or s_Amort = "V") Then
                        n_Monto = Format((n_NewNominal * adoresult!TAS_INTE) + (MontTitu * adoresult!VAL_AMOR), "0.00")
                    End If
                    n_Tasa = adoresult!TAS_INTE
                    n_TasDia = Format(((1 + adoresult!TAS_INTE) ^ (1 / adoresult!CNT_DIAS)), "0.0000000000000000")
                    If (s_Amort = "F" Or s_Amort = "V") Then
                        n_Tasa = adoresult!TAS_INTE + adoresult!VAL_AMOR
                        n_NewNominal = n_NewNominal - (MontTitu * adoresult!VAL_AMOR)
                    End If
                End If
            End If
            If n_NroCupoFin = adoresult!NRO_CUPO Then
                If (s_Amort = "F" Or s_Amort = "V") Then
                    n_Monto = Format(n_Monto, "0.00")
                Else
                    n_Monto = Format(n_Monto + MontNomi, "0.00")
                End If
            End If
            n_Cont = n_Cont + 1
            Array_Monto(n_Cont) = n_Monto: Array_Dias(n_Cont) = d_FchFin
            adoresult.MoveNext
        Loop
        adoresult.Close: Set adoresult = Nothing
    End With
    
    TirNoPerLetPag = TIR(Array_Monto(), Array_Dias(), Tasa) * 100
    
End Function

Function CalculaVANBVAC(File As String, Anal As String, FECHA As String, TIRDiario As Double, TipBono As String) As Double

'    Dim adoResultAux As New ADODB.Recordset, adoresult As New ADODB.Recordset
'    Dim FechaRede As String, ValNomi As Double, nDiasAcum As Integer
'    Dim IntCorrido As Double, AcumFlujo As Double
'    Dim TasDiar As Double, FchIntCor As String, DiasIntCorr As Integer
'    Dim FchCupoMenos1 As Variant, res As Integer
'    Dim n_TasDia As Double, n_Tasa As Double
'
'
'    ' FUNCION QUE CALCULO EL PRECIO DE UN BONO A
'    ' PARTIR DE LA TIR BRUTA DE LA OPERACION.
'
'    With adoComm
'        'Obtener datos del BONO
'        .CommandText = "SELECT * FROM FMBONOS"
'        .CommandText = .CommandText & " WHERE COD_FILE='" & File & "'"
'        .CommandText = .CommandText & " AND COD_ANAL='" & Anal & "'"
'        Set adoresult = .Execute
'
'        'Para Intereses Corridos obtener Cupon Actual donde este incluido FECHA, capturar FCH_INIC, TAS_DIAR
'        'Si el Cupón Vigente es el Primero tomar FCH_INIC, caso contrario tomar FCH_INIC Menos 1 DIA
'        .CommandText = "SELECT NRO_CUPO, FCH_INIC, TAS_DIAR,TAS_INTE2,CNT_DIAS FROM FMCUPON "
'        .CommandText = .CommandText & " WHERE COD_FILE='" & Trim$(File) & "' AND "
'        .CommandText = .CommandText & " COD_ANAL='" & Trim$(Anal) & "' AND FCH_INIC<='" & Fecha & "' AND "
'        .CommandText = .CommandText & " FCH_VCTO>='" & Fecha & "'"
'        Set adoResultAux = .Execute
'        If Not adoResultAux.EOF Then  ' Si tiene cupon vigente
'            If TipBono = "P" Then
'                TasDiar = Format(((1 + adoResultAux!TAS_INTE2) ^ (1 / adoResultAux!CNT_DIAS)) - 1, "0.0000000000000000")
'            Else
'                TasDiar = adoResultAux!TAS_DIAR
'            End If
'
'            If adoResultAux!NRO_CUPO = "001" Then
'                FchIntCor = adoResultAux!FCH_INIC
'            Else
'                FchIntCor = FmtFec(DateAdd("d", -1, DateSerial(Left(adoResultAux!FCH_INIC, 4), Mid(adoResultAux!FCH_INIC, 5, 2), Right(adoResultAux!FCH_INIC, 2))), "WIN", "YYYYMMDD", res)
'            End If
'        Else
'            TasDiar = 0
'            FchIntCor = Fecha   'Como no existe asumo la fecha de cierre para que de 0 dias
'        End If
'        adoResultAux.Close: Set adoResultAux = Nothing
'
'        nDiasAcum = 0: AcumFlujo = 0
'        n_TasDia = 0: n_Tasa = 0
'
'        .CommandText = "SELECT NRO_CUPO, TAS_INTE,TAS_INTE2,VAL_CUPO, FLG_VENC, FCH_INIC, FCH_VCTO,CNT_DIAS,VAL_AMOR FROM FMCUPON "
'        .CommandText = .CommandText & "WHERE COD_FILE='" & Trim$(File) & "' AND "
'        .CommandText = .CommandText & "COD_ANAL='" & Trim$(Anal) & "' AND FLG_VENC <> 'X'"
'        Set adoResultAux = .Execute
'        Do Until adoResultAux.EOF   ' Si el bono tiene cupones vigentes
'            'nDiasAcum = DateDiff("d", DateSerial(Left(Fecha, 4), Mid(Fecha, 5, 2), Right(Fecha, 2)), DateSerial(Left(adoresultAux!FCH_VCTO, 4), Mid(adoresultAux!FCH_VCTO, 5, 2), Right(adoresultAux!FCH_VCTO, 2)))
'            nDiasAcum = DateDiff("d", Convertddmmyyyy(Fecha), Convertddmmyyyy(adoResultAux!FCH_VCTO))
'            If TipBono = "P" Then
'                If adoResultAux!TAS_INTE2 = 0 Then
'                'If adoresult!FLG_AMORT = "X" Then
'                '    n_Tasa = ((n_TasDia# ^ adoresultAux!CNT_DIAS) - 1) + adoresultAux!VAL_AMOR
'                '    AcumFlujo = AcumFlujo + Format((n_Tasa / ((1 + TIRDiario) ^ nDiasAcum)), "0.0000000000") + adoresultAux!VAL_AMOR
'                'Else
'                    n_Tasa = ((n_TasDia ^ adoResultAux!CNT_DIAS) - 1)
'                    AcumFlujo = AcumFlujo + Format((n_Tasa / ((1 + TIRDiario) ^ nDiasAcum)), "0.0000000000")
'                'End If
'                Else
'                'If adoresult!FLG_AMORT = "X" Then
'                '    AcumFlujo = AcumFlujo + Format(((adoresultAux!TAS_INTE + adoresultAux!VAL_AMOR) / ((1 + TIRDiario) ^ nDiasAcum)), "0.0000000000") + adoresultAux!VAL_AMOR
'                'Else
'                    AcumFlujo = AcumFlujo + Format((adoResultAux!TAS_INTE2 / ((1 + TIRDiario) ^ nDiasAcum)), "0.0000000000")
'                'End If
'                    n_TasDia = Format(((1 + adoResultAux!TAS_INTE2) ^ (1 / adoResultAux!CNT_DIAS)), "0.0000000000000000")
'                End If
'            Else
'                If adoResultAux!TAS_INTE = 0 Then
'                'If adoresult!FLG_AMORT = "X" Then
'                '    n_Tasa = ((n_TasDia# ^ adoresultAux!CNT_DIAS) - 1) + adoresultAux!VAL_AMOR
'                '    AcumFlujo = AcumFlujo + Format((n_Tasa / ((1 + TIRDiario) ^ nDiasAcum)), "0.0000000000") + adoresultAux!VAL_AMOR
'                'Else
'                    n_Tasa = ((n_TasDia ^ adoResultAux!CNT_DIAS) - 1)
'                    AcumFlujo = AcumFlujo + Format((n_Tasa / ((1 + TIRDiario) ^ nDiasAcum)), "0.0000000000")
'                'End If
'                Else
'                'If adoresult!FLG_AMORT = "X" Then
'                '    AcumFlujo = AcumFlujo + Format(((adoresultAux!TAS_INTE + adoresultAux!VAL_AMOR) / ((1 + TIRDiario) ^ nDiasAcum)), "0.0000000000") + adoresultAux!VAL_AMOR
'                'Else
'                    AcumFlujo = AcumFlujo + Format((adoResultAux!TAS_INTE / ((1 + TIRDiario) ^ nDiasAcum)), "0.0000000000")
'                'End If
'                    n_TasDia = Format(((1 + adoResultAux!TAS_INTE) ^ (1 / adoResultAux!CNT_DIAS)), "0.0000000000000000")
'                End If
'            End If
'            adoResultAux.MoveNext
'        Loop
'
'        AcumFlujo = AcumFlujo + Format((adoresult!VAL_NOMI / ((1 + TIRDiario) ^ nDiasAcum)), "0.0000000000")  ' Incluir Valor Nominal
'
'        If FchIntCor <> Fecha$ Then 'Existen Intereses Corridos
'            'DiasIntCorr = DateDiff("d", DateSerial(Left(FchIntCor, 4), Mid(FchIntCor, 5, 2), Right(FchIntCor, 2)), DateSerial(Left(Fecha$, 4), Mid(Fecha$, 5, 2), Right(Fecha$, 2)))
'            DiasIntCorr = DateDiff("d", Convertddmmyyyy(FchIntCor), Convertddmmyyyy(Fecha))
'            IntCorrido = Format((((1 + TasDiar) ^ DiasIntCorr) - 1) * adoresult!VAL_NOMI, "0.0000000000")
'        Else
'            IntCorrido = Format(0, "0.0000000000")
'        End If
'
'        AcumFlujo = Format(((AcumFlujo - IntCorrido) / adoresult!VAL_NOMI), "0.0000000000")
'
'        adoResultAux.Close: Set adoResultAux = Nothing
'        adoresult.Close: Set adoresult = Nothing
'    End With
'
'    CalculaVANBVAC = AcumFlujo
    
End Function

Function CalculaVANLH(File As String, Anal As String, FECHA As String, TIRDiario As Double) As Double

    Dim adoResultAux As New ADODB.Recordset, adoresult As New ADODB.Recordset
    Dim FechaRede As String, ValNomi As Double, nDiasAcum As Integer
    Dim IntCorrido As Double, AcumFlujo As Double
    Dim TasDiar As Double, FchIntCor As String, DiasIntCorr As Integer
    Dim FchCupoMenos1 As Variant, res As Integer
    Dim s_SldAmor As Double
   
    ' FUNCION QUE CALCULO EL PRECIO DE UNA LETRA HIP.A
    ' PARTIR DE LA TIR DE LA OPERACION.
   
    With adoComm
        'Obtener datos del BONO
        .CommandText = "SELECT * FROM FMBONOS"
        .CommandText = .CommandText & " WHERE COD_FILE='" & File & "'"
        .CommandText = .CommandText & " AND COD_ANAL='" & Anal & "'"
        Set adoresult = .Execute
   
        '-----------------------------------------------------------------------------------------------------
        ' Para Intereses Corridos obtener Cupon Actual donde este incluido FECHA, capturar FCH_INIC, TAS_DIAR
        ' Si el Cupón Vigente es el Primero tomar FCH_INIC, caso contrario tomar FCH_INIC Menos 1 DIA
        ' Para el CALCULO DEL PRECIO OBTENER EL SALDO POR AMORTIZAR DE LA LETRA HIPOTEC. PARA EL CUPON VIGENTE
        '-----------------------------------------------------------------------------------------------------
        .CommandText = "SELECT NRO_CUPO, FCH_INIC, TAS_DIAR, VAL_AMOR, SLD_AMOR FROM FMCUPON "
        .CommandText = .CommandText & " WHERE COD_FILE='" & Trim$(File) & "' AND "
        .CommandText = .CommandText & " COD_ANAL='" & Trim$(Anal) & "' AND FCH_INIC<='" & FECHA & "' AND "
        .CommandText = .CommandText & " FCH_VCTO>='" & FECHA & "'"
        Set adoResultAux = .Execute
        If Not adoResultAux.EOF Then  ' Si tiene cupon vigente
           TasDiar = adoResultAux!TAS_DIAR
           'H.R.P.>>10/10/97
           If adoResultAux!NRO_CUPO = "001" Then
                FchIntCor = adoResultAux!FCH_INIC
           Else
                'FchIntCor = FmtFec(DateAdd("d", -1, DateSerial(Left(adoresultAux!FCH_INIC, 4), Mid(adoresultAux!FCH_INIC, 5, 2), Right(adoresultAux!FCH_INIC, 2))), "WIN", "YYYYMMDD", res)
                FchIntCor = Convertyyyymmdd(DateAdd("d", -1, Convertddmmyyyy(adoResultAux!FCH_INIC)))
           End If
           s_SldAmor = adoResultAux!VAL_AMOR + adoResultAux!SLD_AMOR
        Else
           TasDiar = 0
           FchIntCor = FECHA   'Como no existe asumo la fecha de cierre para que de 0 dias
           s_SldAmor = 0
        End If
        adoResultAux.Close: Set adoResultAux = Nothing
   
        nDiasAcum = 0
        AcumFlujo = 0
        '------------------------------------------------------------------------------------------
        ' EL CALCULO DEL VAN DE LA LETRA HIPOTECARIA SE REALIZA SOBRE EL VALOR DEL CUPON (VAL_CUPO)
        ' VALOR CONSTANTE EN TODOS LOS CUPONES ( VAL_AMOR + VAL_INTE )
        ' EN LOS FLUJOS ACUMULADOS YA NO SE INCLUYE EL FLUJO CORRESPONDIENTE AL VALOR NOMINAL, PUES
        ' ESTE SE VA CONSIDERANDO DE MANERA PROPORCIONAL EN CADA CUPON DE LA LETRA HIPOTECARIA
        '------------------------------------------------------------------------------------------
        .CommandText = "SELECT NRO_CUPO, VAL_CUPO, FLG_VENC, FCH_INIC, FCH_VCTO, FLG_VIGE FROM FMCUPON "
        .CommandText = .CommandText & "WHERE COD_FILE='" & Trim$(File) & "' AND "
        .CommandText = .CommandText & "COD_ANAL='" & Trim$(Anal) & "' AND FLG_VENC <> 'X'"
        Set adoResultAux = .Execute
        Do Until adoResultAux.EOF   ' Si el bono tiene cupones vigentes
            'nDiasAcum = DateDiff("d", DateSerial(Left(Fecha, 4), Mid(Fecha, 5, 2), Right(Fecha, 2)), DateSerial(Left(adoresultAux!FCH_VCTO, 4), Mid(adoresultAux!FCH_VCTO, 5, 2), Right(adoresultAux!FCH_VCTO, 2)))
            nDiasAcum = DateDiff("d", Convertddmmyyyy(FECHA), Convertddmmyyyy(adoResultAux!fch_vcto))
            AcumFlujo = AcumFlujo + Format((adoResultAux!VAL_CUPO / ((1 + TIRDiario) ^ nDiasAcum)), "0.0000000000")
            adoResultAux.MoveNext
        Loop
        '------------------------------------------------------
        ' LOS INTERESES CORRIDOS SE CALCULAN SOBRE EL SALDO POR
        ' AMORTIZAR DE LA LETRA Y NO SOBRE EL VALOR NOMINAL.
        '------------------------------------------------------
        If FchIntCor <> FECHA$ Then 'Existen Intereses Corridos
           'DiasIntCorr = DateDiff("d", DateSerial(Left(FchIntCor, 4), Mid(FchIntCor, 5, 2), Right(FchIntCor, 2)), DateSerial(Left(Fecha$, 4), Mid(Fecha$, 5, 2), Right(Fecha$, 2)))
           DiasIntCorr = DateDiff("d", Convertddmmyyyy(FchIntCor), Convertddmmyyyy(FECHA))
           If s_SldAmor > 0 Then
              IntCorrido = Format((((1 + TasDiar) ^ DiasIntCorr) - 1) * s_SldAmor, "0.0000000000")
           Else
              IntCorrido = Format((((1 + TasDiar) ^ DiasIntCorr) - 1) * adoresult!VAL_NOMI, "0.0000000000")
           End If
        Else
           IntCorrido = Format(0, "0.0000000000")
        End If
   
        '-----------------------------------------------------------------
        ' PRECIO_LETRA = (V.A.N(LETRA) - INT.CORRIDOS)/SALDO_POR_AMORTIZAR
        '-----------------------------------------------------------------
        If s_SldAmor > 0 Then
           AcumFlujo = Format(((AcumFlujo - IntCorrido) / s_SldAmor), "0.0000000000")
        Else
           AcumFlujo = Format(AcumFlujo - IntCorrido, "0.0000000000")
        End If
   
        adoResultAux.Close: Set adoResultAux = Nothing
        adoresult.Close: Set adoresult = Nothing
    End With
    CalculaVANLH = AcumFlujo

End Function

'Sub LlenaGrdCer(GrdDat As Grid, adoresultAux1 As ADODB.Recordset, Adirreg() As String, n_CntCuo As Double, s_FchHoy As String)
'
'    Dim adoresultaux2 As New Recordset
'    Dim adoResultAux3 As New Recordset
'    Dim adoresultTmp As New Recordset
'    Dim i As Integer, NroFil As Integer, res As Integer
'    Dim s_strCodFon As String, s_CodPar As String
'    Dim n_DifDia As Long, gstrSQL As String, n_TasCom As Double
'    Dim n_ConCuo As Double 'Contador de Cuotas
'    Dim aGrdCnf() As RGrdCnf
'    Dim v_FchHoy As Variant
'    Dim nSW As Integer
'    Dim n_CuoXCer As Double
'
'    On Error Resume Next
'
'    adoresultAux1.MoveFirst
'    s_strCodFon = adoresultAux1!COD_FOND
'    s_CodPar = adoresultAux1!Cod_part
'    'v_FchHoy = FmtFec(s_FchHoy, "yyyymmdd", "win", res)
'    v_FchHoy = Convertddmmyyyy(s_FchHoy)
'
'    '** Configuración de Grilla
'    'ReDim AGrdCnf(1 To 9)
'    ReDim aGrdCnf(1 To 12)
'    'Columna uno Para Indicador
'    aGrdCnf(1).TitDes = "Selección"
'    aGrdCnf(1).DatNom = ""
'    aGrdCnf(1).DatAnc = 130 * 2
'
'    aGrdCnf(2).TitDes = "Certificado"
'    aGrdCnf(2).DatNom = "NRO_CERT"
'    aGrdCnf(2).DatAnc = 130 * 8
'
'    aGrdCnf(3).TitDes = "Tot.Cuotas"
'    aGrdCnf(3).DatNom = "CNT_CUOT"
'    aGrdCnf(3).DatAnc = 130 * 10
'    aGrdCnf(3).DatFmt = "C"
'    aGrdCnf(3).DatJus = 1
'
'    aGrdCnf(4).TitDes = "Rescatar"
'    aGrdCnf(4).DatNom = ""
'    aGrdCnf(4).DatAnc = 130 * 10
'    aGrdCnf(4).DatFmt = "C"
'    aGrdCnf(4).DatJus = 1
'
'    aGrdCnf(5).TitDes = "Fch.Susc."
'    aGrdCnf(5).DatNom = "FCH_SUSC"
'    aGrdCnf(5).DatAnc = 130 * 8
'    aGrdCnf(5).DatFmt = "F"
'
'    aGrdCnf(6).TitDes = "Comisión"
'    aGrdCnf(6).DatNom = ""
'    aGrdCnf(6).DatAnc = 130 * 5
'    aGrdCnf(6).DatFmt = "D"
'    aGrdCnf(6).DatJus = 1
'
'    aGrdCnf(7).TitDes = "NroOper"
'    aGrdCnf(7).DatNom = "NRO_OPER"
'    aGrdCnf(7).DatAnc = 1
'
'    aGrdCnf(8).TitDes = "Tipo"
'    aGrdCnf(8).DatNom = "TIP_OPER"
'    aGrdCnf(8).DatAnc = 1
'
'    aGrdCnf(9).TitDes = "Ext."
'    aGrdCnf(9).DatNom = "FLG_EXTR"
'    aGrdCnf(9).DatAnc = 1
'
'    '*** creadso por l.e inicio
'    aGrdCnf(10).TitDes = "PreImprI."
'    aGrdCnf(10).DatNom = ""
'    aGrdCnf(10).DatAnc = 130 * 8
'
'    aGrdCnf(11).TitDes = "PreImprO"
'    aGrdCnf(11).DatNom = "NRO_DOCU"
'    aGrdCnf(11).DatAnc = 1
'
'    aGrdCnf(12).TitDes = "Flag"
'    aGrdCnf(12).DatNom = ""
'    aGrdCnf(12).DatAnc = 1
'
'    '*** creado por l.e fin
'
'    'Escondo 1ra. Columna
'    GrdDat.ColWidth(0) = 1
'    'Agregar una columna
'
'    'Tasas de Rescate x Fondo
'    adoComm.CommandText = "SELECT CNT_DIA1,TAS_RED1,CNT_DIA2,TAS_RED2,CNT_DIA3,TAS_RED3,CNT_DIA4,TAS_RED4,TAS_RED5,TIP_RESC from FMFONDOS WHERE COD_FOND='" & s_strCodFon & "'"
'    Set adoresultaux2 = adoComm.Execute
'
'    'Limpiar la grilla
'    Do While GrdDat.Rows - 1 > GrdDat.FixedRows
'        GrdDat.RemoveItem (GrdDat.Rows - 1)
'    Loop
'    GrdDat.Row = GrdDat.Rows - 1
'    For i = 0 To GrdDat.Cols - 1
'        GrdDat.Col = i
'        GrdDat.Text = ""
'    Next
'
'    'Títulos de la Grilla
'    'Configurar la grilla
'    GrdDat.Cols = UBound(aGrdCnf) + 1
'    GrdDat.Row = 0
'    GrdDat.Col = 0: GrdDat.Text = "Nro"
'    For i = 1 To GrdDat.Cols - 1
'        GrdDat.Col = i
'        GrdDat.Text = aGrdCnf(i).TitDes
'        GrdDat.FixedAlignment(i) = aGrdCnf(i).TitJus
'        GrdDat.ColAlignment(i) = aGrdCnf(i).DatJus
'        GrdDat.ColWidth(i) = aGrdCnf(i).DatAnc
'    Next
'
'    If adoresultAux1.BOF And adoresultAux1.EOF Then
'        Exit Sub
'    Else
'        adoresultAux1.MoveFirst
'    End If
'    NroFil = 0
'
'    '** Detalle de la Grilla
'    n_ConCuo = 0
'    nSW = True
'    Do Until adoresultAux1.EOF
'        n_ConCuo = n_ConCuo + CDbl(adoresultAux1!CNT_CUOT)
'        NroFil = NroFil + 1
'        ReDim Preserve Adirreg(1 To NroFil)
'        Adirreg(NroFil) = adoresultAux1.Bookmark
'        If NroFil > 1 Then GrdDat.AddItem ""
'        GrdDat.Row = NroFil
'        GrdDat.Col = 0
'        GrdDat.Text = Format$(NroFil, "0000")
'        For i = 1 To GrdDat.Cols - 1
'            GrdDat.Col = i
'
'            '*** creado por l.e inicio
'            If i = 12 Then 'Flags
'                If adoresultAux1!FLG_GARA = "X" Or adoresultAux1!FLG_CUST = " " Then
'                    GrdDat.Text = "X"
'                End If
'            End If
'            '*** creado por l.e fin
'
'            If i = 1 Then 'Indicador
'                If nSW Then
'                    GrdDat.Text = "X"
'                    If n_ConCuo >= n_CntCuo Then
'                        n_CuoXCer = n_CntCuo - (n_ConCuo - CDbl(adoresultAux1!CNT_CUOT))
'                        nSW = False
'                    Else
'                        n_CuoXCer = CDbl(adoresultAux1!CNT_CUOT)
'                    End If
'                Else
'                    n_CuoXCer = 0
'                    GrdDat.Text = ""
'                End If
'            ElseIf i = 4 Then 'Cuotas a Rescatar
'                    If CDbl(n_CuoXCer) > 0 Then
'                        GrdDat.Text = n_CuoXCer
'                    Else
'                        GrdDat.Text = 0
'                    End If
'                ElseIf i = 6 Then 'Comisión
'                        GrdDat.Col = 5 'Fecha
'                        '** Calculo de tasa de la operación
'                        adoComm.CommandText = "SELECT TAS_OPER FROM FMCOMISPART WHERE COD_FOND='" & s_strCodFon & "' AND COD_PART='" + s_CodPar + "' AND TIP_OPER='R' AND TIP_DURA='T' AND FCH_OPER='" & gstrFchCierreAux & "'"
'                        Set adoResultAux3 = adoComm.Execute
'                        If Not adoResultAux3.EOF Then
'                            n_TasCom = adoResultAux3!TAS_OPER
'                            adoResultAux3.Close: Set adoResultAux3 = Nothing
'                        Else
'                            adoResultAux3.Close: Set adoResultAux3 = Nothing
'                            adoComm.CommandText = "SELECT TAS_OPER FROM FMCOMISPART WHERE COD_FOND='" & s_strCodFon & "' AND COD_PART='" & s_CodPar & "' AND TIP_OPER='R' AND TIP_DURA='P' "
'                            Set adoResultAux3 = adoComm.Execute
'                            If Not adoResultAux3.EOF Then
'                                n_TasCom = adoResultAux3!TAS_OPER
'                                adoResultAux3.Close: Set adoResultAux3 = Nothing
'                            Else
'                                adoResultAux3.Close: Set adoResultAux3 = Nothing
'
'                                n_DifDia = DateDiff("d", GrdDat.Text, v_FchHoy)
'                                Select Case Abs(n_DifDia)
'                                    Case 0 To CLng(adoresultaux2!CNT_DIA1)
'                                        n_TasCom = CDbl(adoresultaux2!TAS_RED1)
'                                    Case CLng(adoresultaux2!CNT_DIA1) + 1 To CLng(adoresultaux2!CNT_DIA2)
'                                        n_TasCom = CDbl(adoresultaux2!TAS_RED2)
'                                    Case CLng(adoresultaux2!CNT_DIA2) + 1 To CLng(adoresultaux2!CNT_DIA3)
'                                        n_TasCom = CDbl(adoresultaux2!TAS_RED3)
'                                    Case CLng(adoresultaux2!CNT_DIA3) + 1 To CLng(adoresultaux2!CNT_DIA4)
'                                        n_TasCom = CDbl(adoresultaux2!TAS_RED4)
'                                    Case Is > CLng(adoresultaux2!CNT_DIA4)
'                                        n_TasCom = CDbl(adoresultaux2!TAS_RED5)
'                                End Select
'                            End If
'                        End If
'
'                        '** (1) para los nuevos fondos por L.E INICIO
'                        If adoresultaux2!TIP_RESC = "R" And n_TasCom > 0 Then
'                            Dim xTipOperacion As String
'                            Dim xFchSuscripcion As String
'                            Dim xFchdelDia As String
'                            Dim xValOperacion As Double
'                            Dim xValCuotaAux As Double
'
'                            'GrdDat.Col = 8 'TIPO DE OPERACION
'                            xTipOperacion = adoresultAux1!TIP_OPER ' GrdDat.Text
'                            GrdDat.Col = 5 'FECHA DE SUSCRIPCION
'                            'xFchSuscripcion = Mid$(grddat.Text, 7, 10) & Mid$(grddat.Text, 4, 2) & Mid$(grddat.Text, 1, 2)
'                            xFchSuscripcion = Convertyyyymmdd(GrdDat.Text)
'
'                            'xFchdelDia = Format$(Now, "yyyymmdd")
'                            xFchdelDia = Convertyyyymmdd(Now)
'
'                            adoComm.CommandText = "select VAL_CUOT,VAL_CALC from fmcuotas where fch_cuot='" & xFchSuscripcion & "' and cod_fond='" & s_strCodFon & "'"
'                            Set adoresultTmp = adoComm.Execute
'                            If Not adoresultTmp.EOF Then
'                               Select Case xTipOperacion
'                                      Case "R1", "R2", "SC", "TP", "TT"
'                                           xValOperacion = adoresultTmp!Val_cuot
'                                      Case "R3", "SD", "R4"
'                                           xValOperacion = adoresultTmp!VAL_CALC
'                               End Select
'                            End If
'                            adoresultTmp.Close: Set adoresultTmp = Nothing
'
'                            adoComm.CommandText = "select VAL_CUOT from fmcuotas where fch_cuot='" & gstrFchCierreAux & "' and cod_fond='" & s_strCodFon & "'"
'                            Set adoresultTmp = adoComm.Execute
'                            If Not adoresultTmp.EOF Then
'                               xValCuotaAux = adoresultTmp!Val_cuot
'                            End If
'                            adoresultTmp.Close: Set adoresultTmp = Nothing
'
'                            If xValOperacion > xValCuotaAux Then
'                                n_TasCom = 0
'                            Else
'                                n_TasCom = ((xValCuotaAux - xValOperacion) * (n_TasCom / 100)) / xValCuotaAux
'                                n_TasCom = Val(Format(n_TasCom, "##0.0000000000")) * 100
'                            End If
'                        End If
'                        adoresultaux2.Close: Set adoresultaux2 = Nothing
'                        '** (1) para los nuevos fondos por L.E FIN
'
'                        GrdDat.Col = 6
'                        GrdDat.Text = n_TasCom
'                        'GrdDat.Col = 5
'                    Else
'                        GrdDat.Text = adoresultAux1.Fields(aGrdCnf(i).DatNom)
'                    End If
'
'            If Len(aGrdCnf(i).DatFmt) > 0 Then
'                GrdDat.Text = FmtDat(aGrdCnf(i).DatFmt, GrdDat.Text)
'            End If
'            GrdDat.Col = 3
'        Next
'        adoresultAux1.MoveNext
'    Loop
'
'End Sub

Function TIRDepCrt(n_ValFlujo(), s_FchFlujo(), n_ValTir As Double) As Double

    Dim n_NroDias As Integer, n As Integer, r As Integer, n_ValBase As Integer, c As Integer
    Dim s_FchFlujo1()

    'On Error GoTo ErrorHandler1

    c = 0

    ReDim s_FchFlujo1(UBound(s_FchFlujo))

    '*** Base Anual ***
    n_ValBase = 365

    '*** Cambiar Fechas ***
    For n = 0 To UBound(n_ValFlujo)
        s_FchFlujo1(n) = DateAdd("d", 1, s_FchFlujo(n))
    Next

    '*** Calculando ***
    Do Until Abs(VANDepCrt(n_ValFlujo(), s_FchFlujo(), n_ValTir, 0)) <= 0.00000001 Or c = 100
        n_ValTir = n_ValTir + n_ValBase * (VANDepCrt(n_ValFlujo(), s_FchFlujo(), n_ValTir, 0) / VANDepCrt(n_ValFlujo(), s_FchFlujo1(), n_ValTir, 1))
        c = c + 1
    Loop

    TIRDepCrt = n_ValTir

ExitFunction1:
    Exit Function

ErrorHandler1:

    MsgBox Error(err), vbCritical
    Resume ExitFunction1

End Function

Function VNANoPerCrtDep(Codfile As String, CodAnal As String, Fondo As String, FecOpe, FecFlujo, MontNomi As Double, MontTitu As Double, TirOpe As Double, TipBono As String)

    Dim sensql As String, adoresultTmp As New Recordset
    Dim n_Cont As Integer
    Dim n_Monto As Double, n_NroCupo As Integer, s_fecha As String * 10
    Dim n_Tasa As Double, s_Amort As String, n_NewNominal As Double, n_TasDia As Double
    Dim n_Acum As Double, Tasa As Double, d_FchIni, d_FchFin, n_CntDias As Integer, n_NroCupoFin As Integer
    Dim n_Reg As Integer, i As Integer
    
    adoComm.CommandText = "SELECT FLG_AMORT FROM FMDEPBAN WHERE COD_FILE='" & Trim(Codfile) & "' AND COD_ANAL='" & Trim(CodAnal) & "' AND COD_FOND='" & Fondo & "'"
    Set adoresultTmp = adoComm.Execute
    s_Amort = IIf(IsNull(adoresultTmp!FLG_AMORT), "", adoresultTmp!FLG_AMORT)
    adoresultTmp.Close: Set adoresultTmp = Nothing

    With adoComm
        .CommandText = "SELECT CONVERT(INT,NRO_CUPO) NRO_CUPO FROM FMCUPONES"
        .CommandText = .CommandText & " WHERE COD_FILE='" & Trim(Codfile) & "' AND COD_ANAL='" & Trim(CodAnal) & "' AND COD_FOND='" & Fondo & "'"
        '.CommandText = .CommandText & " AND FCH_INIC<='" & Format(FecFlujo, "yyyymmdd") & "'"
        .CommandText = .CommandText & " AND FCH_INIC<='" & Convertyyyymmdd(FecFlujo) & "'"
        '.CommandText = .CommandText & " AND FCH_VCTO>='" & Format(FecFlujo, "yyyymmdd") & "'"
        .CommandText = .CommandText & " AND FCH_VCTO>='" & Convertyyyymmdd(FecFlujo) & "'"
        Set adoresultTmp = .Execute
    End With
    n_NroCupo = adoresultTmp!NRO_CUPO
    adoresultTmp.Close: Set adoresultTmp = Nothing

    sensql = "SELECT FCH_VCTO,TAS_INTE,TAS_INTE2,VAL_AMOR,CNT_DIAS,NRO_CUPO FROM FMCUPONES"
    sensql = sensql & " WHERE COD_FILE='" & Trim(Codfile) & "' AND COD_ANAL='" & Trim(CodAnal) & "' AND COD_FOND='" & Fondo & "'"
    sensql = sensql & " AND CONVERT(INT,NRO_CUPO) >= " & n_NroCupo & " ORDER BY NRO_CUPO"
    adoresultTmp.Open sensql, adoConn, adOpenStatic
    
    adoresultTmp.MoveLast
    n_Reg = adoresultTmp.RecordCount

    ReDim Array_Monto(n_Reg + 1): ReDim Array_Dias(n_Reg + 1)

    n_Acum = 1: Tasa = TirOpe: n_Cont = 1
    n_NroCupoFin = adoresultTmp!NRO_CUPO
    adoresultTmp.MoveFirst

    n_Monto = Format(0, "0.00")
    Array_Monto(n_Cont) = 0: Array_Dias(n_Cont) = 0
    s_fecha = Convertyyyymmdd(FecFlujo)
    s_fecha = CStr(Convertddmmyyyy(s_fecha))
    d_FchIni = CVDate(s_fecha)
    n_Tasa = 0
    If TipBono = "P" Then
        n_Tasa = adoresultTmp!TAS_INTE2
        n_TasDia = Format(((1 + adoresultTmp!TAS_INTE2) ^ (1 / adoresultTmp!CNT_DIAS)), "0.0000000000000000")
        If (s_Amort = "F" Or s_Amort = "V") Then
            n_Tasa = adoresultTmp!TAS_INTE2 + adoresultTmp!VAL_AMOR
        End If
    Else
        n_Tasa = adoresultTmp!TAS_INTE
        n_TasDia = Format(((1 + adoresultTmp!TAS_INTE) ^ (1 / adoresultTmp!CNT_DIAS)), "0.0000000000000000")
        If (s_Amort = "F" Or s_Amort = "V") Then
            n_Tasa = adoresultTmp!TAS_INTE + adoresultTmp!VAL_AMOR
        End If
    End If

    n_NewNominal = MontNomi

    Do While Not adoresultTmp.EOF
        's_fecha = Right$(adoresultTmp!FCH_VCTO, 2) + "/" + Mid$(adoresultTmp!FCH_VCTO, 5, 2) + "/" + Left$(adoresultTmp!FCH_VCTO, 4)
        s_fecha = CStr(Convertddmmyyyy(adoresultTmp!fch_vcto))
        d_FchFin = CVDate(s_fecha)
        n_CntDias = DateDiff("d", d_FchIni, d_FchFin)
        If TipBono = "P" Then
            If adoresultTmp!TAS_INTE2 = 0 Then
                If (s_Amort = "F" Or s_Amort = "V") Then
                    n_Tasa = ((n_TasDia ^ adoresultTmp!CNT_DIAS) - 1) + adoresultTmp!VAL_AMOR
                Else
                    n_Tasa = ((n_TasDia ^ adoresultTmp!CNT_DIAS) - 1)
                End If
                n_Monto = Format(MontNomi * n_Tasa, "0.00")
                If (s_Amort = "F" Or s_Amort = "V") Then
                    n_Monto = Format((n_NewNominal * n_Tasa) + (MontTitu * adoresultTmp!VAL_AMOR), "0.00")
                    n_NewNominal = n_NewNominal - (MontTitu * adoresultTmp!VAL_AMOR)
                End If
            Else
                n_Monto = Format(MontNomi * adoresultTmp!TAS_INTE2, "0.00")
                If (s_Amort = "F" Or s_Amort = "V") Then
                    n_Monto = Format((n_NewNominal * adoresultTmp!TAS_INTE2) + (MontTitu * adoresultTmp!VAL_AMOR), "0.00")
                End If
                n_Tasa = adoresultTmp!TAS_INTE2
                n_TasDia = Format(((1 + adoresultTmp!TAS_INTE2) ^ (1 / adoresultTmp!CNT_DIAS)), "0.0000000000000000")
                If (s_Amort = "F" Or s_Amort = "V") Then
                    n_Tasa = adoresultTmp!TAS_INTE + adoresultTmp!VAL_AMOR
                    n_NewNominal = n_NewNominal - (MontTitu * adoresultTmp!VAL_AMOR)
                End If
            End If
        Else
            If adoresultTmp!TAS_INTE = 0 Then
                If (s_Amort = "F" Or s_Amort = "V") Then
                    n_Tasa = ((n_TasDia ^ adoresultTmp!CNT_DIAS) - 1) + adoresultTmp!VAL_AMOR
                Else
                    n_Tasa = ((n_TasDia ^ adoresultTmp!CNT_DIAS) - 1)
                End If
                n_Monto = Format(MontNomi * n_Tasa, "0.00")
                If (s_Amort = "F" Or s_Amort = "V") Then
                    n_Monto = Format((n_NewNominal * n_Tasa) + (MontTitu * adoresultTmp!VAL_AMOR), "0.00")
                    n_NewNominal = n_NewNominal - (MontTitu * adoresultTmp!VAL_AMOR)
                End If
            Else
                n_Monto = Format(MontNomi * adoresultTmp!TAS_INTE, "0.00")
                If (s_Amort = "F" Or s_Amort = "V") Then
                    n_Monto = Format((n_NewNominal * adoresultTmp!TAS_INTE) + (MontTitu * adoresultTmp!VAL_AMOR), "0.00")
                End If
                n_Tasa = adoresultTmp!TAS_INTE
                n_TasDia = Format(((1 + adoresultTmp!TAS_INTE) ^ (1 / adoresultTmp!CNT_DIAS)), "0.0000000000000000")
                If (s_Amort = "F" Or s_Amort = "V") Then
                    n_Tasa = adoresultTmp!TAS_INTE + adoresultTmp!VAL_AMOR
                    n_NewNominal = n_NewNominal - (MontTitu * adoresultTmp!VAL_AMOR)
                End If
            End If
        End If
        If n_NroCupoFin = adoresultTmp!NRO_CUPO Then
            If (s_Amort = "F" Or s_Amort = "V") Then
                n_Monto = Format(n_Monto, "0.00")
            Else
                n_Monto = Format(n_Monto + MontNomi, "0.00")
            End If
        End If
        n_Cont = n_Cont + 1
        Array_Monto(n_Cont) = n_Monto: Array_Dias(n_Cont) = n_CntDias
        adoresultTmp.MoveNext
    Loop
        
    n_Acum = 0
    For i = 1 To n_Cont
        n_Acum = n_Acum + (Array_Monto(i) / ((1 + Tasa) ^ (Array_Dias(i) / 365)))
    Next

    adoresultTmp.Close: Set adoresultTmp = Nothing
    
    VNANoPerCrtDep = Format(n_Acum, "0.00")

End Function

Function AccesoForm(ByVal strNomOpc As String, ByVal strNumInd As String) As String

    Dim adoRecord As New Recordset


'    adoComm.CommandText = "Select * from syusuari where user_id = '" & gstrLogin & "'"
'    Set adoRecord = adoComm.Execute
'
'    If (Trim(adoRecord!PERF_ACCE) = "S") Then
'        AccesoForm = "1"
'        adoRecord.Close: Set adoRecord = Nothing
'    Else
'        adoComm.CommandText = "Select syopcpart.* From syopcion, syopcpart where syopcion.COD_MODU = syopcpart.COD_MODU and syopcion.COD_MODU = '" & frmMainMdi.Tag & "' and "
'        adoComm.CommandText = adoComm.CommandText & "syopcion.nom_opc  = '" & strNomOpc & "' and syopcion.num_ind  = '" & strNumInd & "' and syopcion.COD_OPC  = syopcpart.COD_OPC and syopcpart.USER_ID = '" & gstrLogin & "'"
'        Set adoRecord = adoComm.Execute
'        If adoRecord.EOF Then
'            adoRecord.Close: Set adoRecord = Nothing
'            adoComm.CommandText = "Select syopcion.* From syopcion where syopcion.COD_MODU = '" & frmMainMdi.Tag & "' and syopcion.nom_opc = '" & strNomOpc & "' and syopcion.num_ind  = '" & strNumInd & "'"
'            Set adoRecord = adoComm.Execute
'            If adoRecord!flg_vis = "X" Then
'                AccesoForm = "3"
'            Else
'                AccesoForm = "5"
'            End If
'        Else
'            AccesoForm = adoRecord!NIV_ACC
'            If Trim(adoRecord!NIV_ACC) = "" Then
'                AccesoForm = "5"
'            End If
'        End If
'        adoRecord.Close: Set adoRecord = Nothing
'    End If

End Function

Function VANDepCrt(n_ValFlujo(), s_FchFlujo(), n_ValTir As Double, n_TipFunc As Integer) As Double

    Dim n_NroDias As Integer, n As Integer, r As Integer, n_ValBase As Integer
    Dim res As Double

    'On Error GoTo ErrorHandler

    res = 0

    If IsNull(n_TipFunc) Then
        n_TipFunc = 0
    End If

    '*** Base Anual ***
    n_ValBase = 365

    '*** Función de Valor Actual ***
    For n = 0 To UBound(n_ValFlujo)
        n_NroDias = DateDiff("d", s_FchFlujo(0), s_FchFlujo(n))
        If n_TipFunc = 0 Then
            res = res + (n_ValFlujo(n) / ((1 + n_ValTir) ^ (n_NroDias / n_ValBase)))
        Else
            res = res + (n_NroDias * n_ValFlujo(n) / ((1 + n_ValTir) ^ (n_NroDias / n_ValBase)))
        End If
    Next

    VANDepCrt = res

ExitFunction:
    Exit Function

Errorhandler:
    MsgBox Error(err), vbCritical
    Resume ExitFunction

End Function

Function IsValidPath(strDestPath As String, ByVal strDefaultDrive As String) As Integer

    Dim strTMP As String, strdrive As String, strlegalChar As String, strTemp As String
    Dim intBackPos As Integer, intForePos  As Integer, i As Integer, intperiodPos As Integer, intlength As Integer

' arguments:  DestPath$         a string that is a full path
'             DefaultDrive$     the default drive.  eg.  "C:"
'
'  If DestPath$ does not include a drive specification,
'  IsValidPath uses Default Drive
'
'  When IsValidPath is finished, DestPath$ is reformated
'  to the format "X:\dir\dir\dir\"
'
' Result:  True (-1) if path is valid.
'          False (0) if path is invalid
'-------------------------------------------------------
    '----------------------------
    ' Remove left and right spaces
    '----------------------------
    strDestPath = RTrim(LTrim(strDestPath))
    

    '-----------------------------
    ' Check Default Drive Parameter
    '-----------------------------
    If Right(strDefaultDrive, 1) <> ":" Or Len(strDefaultDrive) <> 2 Then
        MsgBox "Parámetro especificado es Inválido.  Ud. indicó,  """ & strDefaultDrive & """.  Debe indicar el drive y un "":"".  Por Ej. , ""C:"", ""D:""...", 64, "Error en Validación"
        GoTo parseErr
    End If
    

    '-------------------------------------------------------
    ' Insert default drive if path begins with root backslash
    '-------------------------------------------------------
    If Left(strDestPath, 1) = "\" Then
        strDestPath = strDefaultDrive + strDestPath
    End If
    
    '-----------------------------
    ' check for invalid characters
    '-----------------------------
    On Error Resume Next
    strTMP = dir(strDestPath)
    If err <> 0 Then
        GoTo parseErr
    End If
    

    '-----------------------------------------
    ' Check for wildcard characters and spaces
    '-----------------------------------------
    If (InStr(strDestPath, "*") <> 0) Then GoTo parseErr
    If (InStr(strDestPath, "?") <> 0) Then GoTo parseErr
    If (InStr(strDestPath, " ") <> 0) Then GoTo parseErr
         
    
    '------------------------------------------
    ' Make Sure colon is in second char position
    '------------------------------------------
    If Mid(strDestPath, 2, 1) <> Chr$(58) Then GoTo parseErr
    

    '-------------------------------
    ' Insert root backslash if needed
    '-------------------------------
    If Len(strDestPath) > 2 Then
      If Right(Left(strDestPath, 3), 1) <> "\" Then
        strDestPath = Left(strDestPath, 2) + "\" + Right(strDestPath, Len(strDestPath) - 2)
      End If
    End If

    '-------------------------
    ' Check drive to install on
    '-------------------------
    strdrive = Left(strDestPath, 1)
    ChDrive (strdrive)                                                        ' Try to change to the dest drive
    If err <> 0 Then GoTo parseErr
    
    '-----------
    ' Add final \
    '-----------
    If Right(strDestPath, 1) <> "\" Then
        strDestPath = strDestPath + "\"
    End If
    

    '-------------------------------------
    ' Root dir is a valid dir
    '-------------------------------------
    If Len(strDestPath) = 3 Then
        If Right(strDestPath, 2) = ":\" Then
            GoTo ParseOK
        End If
    End If
    

    '------------------------
    ' Check for repeated Slash
    '------------------------
    If InStr(strDestPath, "\\") <> 0 Then GoTo parseErr
        
    '--------------------------------------
    ' Check for illegal directory names
    '--------------------------------------
    strlegalChar = "!#$%&'()-0123456789@ABCDEFGHIJKLMNOPQRSTUVWXYZ^_`{}~.üäöÄÖÜß"
    intBackPos = 3
    intForePos = InStr(4, strDestPath, "\")
    Do
        strTemp = Mid(strDestPath, intBackPos + 1, intForePos - intBackPos - 1)
        
        '----------------------------
        ' Test for illegal characters
        '----------------------------
        For i = 1 To Len(strTemp)
            If InStr(strlegalChar, UCase(Mid(strTemp, i, 1))) = 0 Then GoTo parseErr
        Next

        '-------------------------------------------
        ' Check combinations of periods and intlengths
        '-------------------------------------------
        intperiodPos = InStr(strTemp, ".")
        intlength = Len(strTemp)
        If intperiodPos = 0 Then
            If intlength > 12 Then GoTo parseErr                         ' Base too long
        Else
            If intperiodPos > 13 Then GoTo parseErr                      ' Base too long
            If intlength > intperiodPos + 3 Then GoTo parseErr             ' Extension too long
            If InStr(intperiodPos + 1, strTemp, ".") <> 0 Then GoTo parseErr ' Two periods not allowed
        End If

        intBackPos = intForePos
        intForePos = InStr(intBackPos + 1, strDestPath, "\")
    Loop Until intForePos = 0

ParseOK:
    IsValidPath = True
    Exit Function

parseErr:
    IsValidPath = False

End Function
Public Function ObtenerValorTipoCambio(ByVal strpCodMoneda As String, ByVal strpCodMonedaCambio As String, ByVal strpFechaMovimiento As String, ByVal strpFechaTipoCambio As String, ByVal strpCodTipoCambio As String, ByVal strpCodClaseTipoCambio As String, Optional ByVal intpModalidadCambio As Integer = 5, Optional ByVal strpTipoCambioReemplazoXML As String = "<TipoCambioReemplazo />") As Double

    
    Dim adoRegistro As ADODB.Recordset, dblValorTipoCambio As Double
    
    ObtenerValorTipoCambio = 0
    dblValorTipoCambio = 0
    
'    If strpCodMoneda = "02" Then
'        dblValorTipoCambio = 2.65
'    Else
'        dblValorTipoCambio = 1
'    End If
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        .CommandText = "{ call up_ACObtenerTipoCambioContable1('" & strpCodMoneda & "','" & _
                        strpCodMonedaCambio & "','" & strpFechaMovimiento & "','" & strpFechaMovimiento & "','" & _
                        strpCodTipoCambio & "','" & strpCodClaseTipoCambio & "'," & intpModalidadCambio & ",'" & _
                        strpTipoCambioReemplazoXML & "') }"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            dblValorTipoCambio = adoRegistro.Fields("ValorTipoCambio").Value
        Else
            adoRegistro.Close: Set adoRegistro = Nothing
            Exit Function
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    ObtenerValorTipoCambio = dblValorTipoCambio
    
    
End Function
Public Function ValidarForm(frm As VB.Form)
Dim ctl As Control

ValidarForm = False

For Each ctl In frm.Controls
    
    If TypeOf ctl Is TextBox Then
        If ctl.Visible = True And ctl.Enabled = True Then
            
            Select Case ctl.Tag
                Case "S" 'String
                    If Len(Trim(ctl.Text)) = 0 Then
                        MsgBox "Por favor completar la información!", vbExclamation + vbOKOnly
                        ctl.SetFocus
                        Exit Function
                    End If
                        
                Case "N" 'Numericos
                    If Not IsNumeric(ctl.Text) Then
                        MsgBox "Por favor ingresar correctamente la información!", vbExclamation + vbOKOnly
                        ctl.SetFocus
                        Exit Function
                    End If
                
                'Case "F" 'Fechas
                '    If Not IsDate(ctl.Text) Or Len(ctl.Text) <> 10 Then
                '        MsgBox "Por favor ingresar correctamente la información!", vbExclamation + vbOKOnly
                '        ctl.SetFocus
                '        Exit Function
                '    End If
            
            End Select
        End If
    End If
    
    If TypeOf ctl Is ComboBox Then
        If ctl.Visible = True And ctl.Enabled = True And Len(Trim(ctl.Tag)) > 0 Then
            If ctl.ListIndex = -1 And Trim(ctl.Text) = "" Then
                MsgBox "Por favor completar la información!", vbExclamation + vbOKOnly
                ctl.SetFocus
                Exit Function
            End If
        End If
    End If
    
    'If TypeOf ctl Is OptionButton Then
    '    If ctl.Visible = True And ctl.Enabled = True Then
    '
    '    End If
    'End If

Next

ValidarForm = True

End Function
Public Function ObtenerNuevaAnalitica(strCodTipoInstrumento As String)

    Dim adoRegistro     As ADODB.Recordset
    Dim strCodAnalitica As String

    Set adoRegistro = New ADODB.Recordset
    
    strCodAnalitica = "00000000"
    
    With adoComm
        '*** Obtener el número de la analítica ***
        .CommandText = "{call up_ACSelDatosParametro(21,'" & strCodTipoInstrumento & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodAnalitica = Format(CInt(adoRegistro("NumUltimo")) + 1, "00000000")
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With

    ObtenerNuevaAnalitica = strCodAnalitica

End Function
Public Sub OrdenarDBGrid(ByVal ColIndex As Integer, ByRef rs As ADODB.Recordset, ByRef DBGrid As Object)

    Dim strColName As String
    Dim strImagePath As String
    Static bSortAsc As Boolean
    Static strPrevCol As String

    strColName = DBGrid.Columns(ColIndex).DataField


    If strColName = strPrevCol Then
    
        If bSortAsc Then
            rs.Sort = strColName & " DESC"
            bSortAsc = False
            strImagePath = gstrImagePath & "SortDown.bmp"
        Else
            rs.Sort = strColName
            bSortAsc = True
            strImagePath = gstrImagePath & "SortUp.bmp"
        End If
        
    Else
        rs.Sort = strColName
        bSortAsc = True
        strImagePath = gstrImagePath & "SortUp.bmp"
    End If
      
    strPrevCol = strColName
    
    DBGrid.Splits(0).Columns(ColIndex).HeadingStyle.ForegroundPicture = LoadPicture(strImagePath)
    DBGrid.Splits(0).Columns(ColIndex).HeadingStyle.ForegroundPicturePosition = dbgFPRight
    DBGrid.Splits(0).Columns(ColIndex).HeadingStyle.TransparentForegroundPicture = True
    
    DBGrid.Refresh

  
 End Sub
Public Function ObtenerDescripcionCuenta(ByVal strpCodCuenta As String) As String

    Dim adoRegistro     As ADODB.Recordset
    Dim strDescripcion  As String
    
    ObtenerDescripcionCuenta = Valor_Caracter
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT DescripCuenta FROM PlanContable " & _
            "WHERE CodCuenta='" & strpCodCuenta & "' AND CodAdministradora = '" & gstrCodAdministradora + "'"
            Set adoRegistro = .Execute
            
            If Not adoRegistro.EOF Then
                strDescripcion = Trim(adoRegistro("DescripCuenta"))
            Else
                adoRegistro.Close: Set adoRegistro = Nothing
                Exit Function
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    ObtenerDescripcionCuenta = strDescripcion
    
End Function


'
' Calculate_accrued_interest
'
Public Function AI_Factor(ByVal DCM As Long, ByVal D1M1Y1 As Date, ByVal D2M2Y2 As Date, ByVal D3M3Y3 As Date, ByVal F As Double, ByVal Maturity As Date) As Double
' DCM is day count method code
' D1M1Y1 is D1.M1.Y1 etc.
' F is the coupon frequency
' Maturity is the maturity date
Dim D1, D1x, M1, Y1 As Long ' variables used to hold day, month and year
Dim D2, D2x, M2, Y2 As Long ' values separately
Dim D3, M3, Y3 As Long
Dim Anchor As Date ' "anchor" date: start date for notional...
Dim AD, AM, AY As Long ' ...periods
Dim WM, WY As Long ' "working dates" in notional period loop
Dim Target As Date ' end date for notional period loop
Dim n, Nx As Long ' number of interest-bearing days
Dim Y As Double ' length of a year (ISMA-Year)
Dim L As Long ' regular coupon length in months
Dim c, Cx As Double ' notional period length in days
Dim Fx As Double ' applicable coupon frequency
Dim Periodic, Regular As Boolean ' various flags
Dim Direction As Long
Dim CurrC, NextC, TempD As Date ' used for temporary serial date values
Dim i As Long ' temporary loop variable
'
' Check input dates
'
If (D1M1Y1 = 0) Or (D2M2Y2 = 0) Or (D3M3Y3 = 0) Or (D2M2Y2 < D1M1Y1) Or (D3M3Y3 <= D1M1Y1) Or (Maturity = 0) Then
    AI_Factor = ERROR_BAD_DATES
    Exit Function
End If
'
' Determine Number of Interest-bearing days, N
'
Select Case DCM
Case GERMAN ' RULE 1
    D1 = Day(D1M1Y1): M1 = Month(D1M1Y1): Y1 = Year(D1M1Y1)
    D2 = Day(D2M2Y2): M2 = Month(D2M2Y2): Y2 = Year(D2M2Y2)
    If D1 = 31 Then
        D1x = 30
    ElseIf IsFebUltimo(D1M1Y1) Then ' end of February
        D1x = 30
    Else
        D1x = D1
    End If
    If D2 = 31 Then
        D2x = 30
    ElseIf IsFebUltimo(D2M2Y2) Then ' end of February
        D2x = 30
    Else
        D2x = D2
    End If
    n = (D2x - D1x) + 30 * (M2 - M1) + 360 * (Y2 - Y1)
Case SPEC_GERMAN ' RULE 2
    D1 = Day(D1M1Y1): M1 = Month(D1M1Y1): Y1 = Year(D1M1Y1)
    D2 = Day(D2M2Y2): M2 = Month(D2M2Y2): Y2 = Year(D2M2Y2)
    If D1 = 31 Then
        D1x = 30
    Else
        D1x = D1
    End If
    If D2 = 31 Then
        D2x = 30
    Else
        D2x = D2
    End If
    n = (D2x - D1x) + 30 * (M2 - M1) + 360 * (Y2 - Y1)
Case ENGLISH, FRENCH, ISMA_YEAR, ISMA_99N, ISMA_99U ' RULES 3, 4, 6, 7
    n = D2M2Y2 - D1M1Y1
Case US ' RULE 5
    D1 = Day(D1M1Y1): M1 = Month(D1M1Y1): Y1 = Year(D1M1Y1)
    D2 = Day(D2M2Y2): M2 = Month(D2M2Y2): Y2 = Year(D2M2Y2)
    D1x = D1: D2x = D2
    If IsFebUltimo(D1M1Y1) And IsFebUltimo(D2M2Y2) Then
        D2x = 30
    End If
    If IsFebUltimo(D1M1Y1) Then
        D1x = 30
    End If
    If (D2x = 31) And (D1x >= 30) Then
        D2x = 30
    End If
    If D1x = 31 Then
        D1x = 30
    End If
    n = (D2x - D1x) + 30 * (M2 - M1) + 360 * (Y2 - Y1)
Case Else
    AI_Factor = ERROR_BAD_DCM
    Exit Function
End Select
'
' Determine Basic Accrued Interest Factor
'
Select Case DCM
Case GERMAN, SPEC_GERMAN, FRENCH, US ' RULES 8, 9, 11, 12
    AI_Factor = n / 360# ' force double precision arithmetic!
Case ENGLISH ' RULE 10
    AI_Factor = n / 365# ' force double precision arithmetic!
Case ISMA_YEAR
    D1 = Day(D1M1Y1): M1 = Month(D1M1Y1): Y1 = Year(D1M1Y1)
    D3 = Day(D3M3Y3): M3 = Month(D3M3Y3): Y3 = Year(D3M3Y3)
    If F = 1 Then ' RULE 14
            i = (D3M3Y3 - D1M1Y1)
        If (i = 365) Or (i = 366) Then
            Y = i
        Else
            Y = 365
            For i = Y1 To Y3
                TempD = GetUltimo(i, 2) ' last day in February
                If (Day(TempD) = 29) And (TempD > D1M1Y1) And (TempD <= D3M3Y3) Then
                    Y = 366
                    Exit For
                End If
            Next i
        End If
    Else ' RULE 15
        If ((Y3 Mod 4 = 0) And (Y3 Mod 100 <> 0)) Or (Y3 Mod 400 = 0) Then
            Y = 366
        Else
            Y = 365
        End If
    End If
    AI_Factor = n / Y ' RULE 13
Case ISMA_99N, ISMA_99U
    D1 = Day(D1M1Y1): M1 = Month(D1M1Y1): Y1 = Year(D1M1Y1)
    D3 = Day(D3M3Y3): M3 = Month(D3M3Y3): Y3 = Year(D3M3Y3)
    ' check whether the frequency is periodic or not and look if the period is regular
    ' set up default values (assume aperiodic, irregular unless otherwise)
    Periodic = False ' aperiodic
    L = 12 ' regular period length in months
    Fx = 1 ' applicable coupon frequency
    Regular = False
    If F >= 1 Then ' RULE 21
        If (12 \ F) = (12 / F) Then ' RULES 19, 20
            Periodic = True ' periodic
            L = 12 \ F ' regular period length in months
            Fx = F ' applicable coupon frequency
            Regular = False ' default: not regular
            If ((Y3 - Y1) * 12 + (M3 - M1)) = L Then ' RULES 23, 24
                If DCM = ISMA_99N Then ' ISMA-99 Normal
                    If (D1 = D3) Then
                        Regular = True
                    ElseIf InvalidDate(Y1, M1, D3) And IsUltimo(D1M1Y1) Then
                        Regular = True
                    ElseIf InvalidDate(Y3, M3, D1) And IsUltimo(D3M3Y3) Then
                        Regular = True
                    End If
                Else ' ISMA-99 Ultimo
                    If IsUltimo(D1M1Y1) And IsUltimo(D3M3Y3) Then Regular = True
                End If
            End If
        End If
    End If
    If Regular Then ' RULE 17
        c = (D3M3Y3 - D1M1Y1)
        AI_Factor = (1 / Fx) * (n / c)
    Else ' generate notional periods
            AI_Factor = 0#
            If D3M3Y3 = Maturity Then ' RULE 18
                Direction = 1 ' ... forwards
                Anchor = D1M1Y1
                AY = Y1: AM = M1: AD = D1
                Target = D3M3Y3
            Else
                Direction = -1 ' ... backwards
                Anchor = D3M3Y3
                AY = Y3: AM = M3: AD = D3
                Target = D1M1Y1
            End If
        CurrC = Anchor ' start notional loop
        i = 0
        While Direction * (CurrC - Target) < 0
            i = i + Direction
            WY = GetNewYear(AY, AM, (i * L)) ' next notional year and...
            WM = GetNewMonth(AM, (i * L)) ' ...month (handling year changes)
                If DCM = ISMA_99N Then ' ISMA-99 Normal
                    If InvalidDate(WY, WM, AD) Then ' RULE 23
                        NextC = GetUltimo(WY, WM)
                    Else
                        NextC = DateSerial(WY, WM, AD)
                    End If
                Else ' ISMA-99 Ultimo
                    NextC = GetUltimo(WY, WM) ' RULE 24
                End If
            Nx = Min(D2M2Y2, Max(NextC, CurrC)) - Max(D1M1Y1, Min(CurrC, NextC))
            Cx = Direction * (NextC - CurrC)
                If Nx > 0 Then ' RULE 22
                    AI_Factor = AI_Factor + (Nx / Cx) ' RULE 21
                End If
            CurrC = NextC
        Wend
            AI_Factor = AI_Factor / Fx ' RULE 22
    End If
Case Else
    End Select
End Function
Function Min(ByVal a As Date, ByVal B As Date) As Date ' lesser of two dates
    If a < B Then Min = a Else Min = B
End Function
Function Max(ByVal c As Date, ByVal D As Date) As Date ' greater of two dates
    If c > D Then Max = c Else Max = D
End Function
Public Function GetUltimo(ByVal YY As Long, _
ByVal MM As Long) As Date ' last day in month MM.YY
    ' NB: MM.YY must be valid
    ' if MM = 12 then handles new year boundary
    GetUltimo = DateSerial(YY + (MM \ 12), (MM Mod 12) + 1, 1) - 1
End Function
Public Function IsUltimo(ByVal DS As Date) As Boolean ' is DS last day in month
    IsUltimo = (Day(DS + 1) = 1)
End Function
Public Function IsFebUltimo(ByVal DS As Date) As Boolean ' is DS last day in February
    IsFebUltimo = (Day(DS + 1) = 1) And (Month(DS + 1) = 3)
End Function
Public Function InvalidDate(ByVal YY As Long, ByVal MM As Long, ByVal DD As Long) As Boolean ' check if a valid date
    InvalidDate = (Month(DateSerial(YY, MM, 1) + DD - 1) <> MM)
End Function
Public Function GetNewMonth(ByVal MM As Long, ByVal Num As Long) As Long ' new month MM +/- Num months
Dim NM As Long ' NB: MM must be valid
    NM = MM + Num
    If NM > 0 Then
        GetNewMonth = (NM - 1) Mod 12 + 1
    Else
        GetNewMonth = 12 + (NM Mod 12)
    End If
End Function
Public Function GetNewYear(ByVal YY As Long, ByVal MM As Long, ByVal Num As Long) As Long ' get new year starting from MM.YY
Dim NM As Long ' going +/- Num months (MM.YY valid)
    NM = MM + Num
    If NM > 0 Then
        GetNewYear = YY + ((NM - 1) \ 12)
    Else
        GetNewYear = YY - 1 + (NM \ 12)
    End If
End Function

Public Function ValidarPermisoAccesoObjeto(ByVal strIdUsuario As String, ByVal strCodObjeto As String, ByVal strTipoObjeto As String) As Boolean
    
    Dim adoAtributoPermitido As ADODB.Recordset, adoValorPermiso As ADODB.Recordset
    Dim strCodAtributo      As String, arrAtributosPermitidos() As String, mensaje As String
    Dim sesionValida        As Boolean
    Dim cont As Integer, i As Integer
    
    ValidarPermisoAccesoObjeto = False
    
    Dim strBDSeguridad As String, strServerSeguridad As String
    Dim strBDSeguridadVa As String, strServerSeguridadVa As String
    Dim strSeguridadActivada As String, strSeguridadActivadaVa As String
    
    Dim lngValorRetorno As Long         'result of the API functions
    Dim lngCodigoLlave  As Long         'handle of opened key
    Dim vntValor        As Variant      'setting of queried value
    Dim adoConnSeguridad                          As ADODB.Connection
    Dim adoCommSeguridad                          As ADODB.Command
    Dim strEstadoObjeto As String
    
'    ValidarPermisoAccesoObjeto = True
'    Exit Function
    
    strBDSeguridad = "Base de Datos"
    strServerSeguridad = "Servidor"
    strSeguridadActivada = "Activada"
    
    lngValorRetorno = RegOpenKeyEx(HKEY_CURRENT_USER, Clave_Registro_Sistema_Seguridad, 0, KEY_QUERY_VALUE, lngCodigoLlave)
    '*** Nombre de Servidor ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, strServerSeguridad, vntValor)
    strServerSeguridadVa = vntValor
    '*** Nombre de Base de Datos ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, strBDSeguridad, vntValor)
    strBDSeguridadVa = vntValor
    '*** Indicador de seguridad activada ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, strSeguridadActivada, vntValor)
    strSeguridadActivadaVa = vntValor
            
    If strSeguridadActivadaVa = "0" Then 'Seguridad no activada
        ValidarPermisoAccesoObjeto = True
        Exit Function
    End If
    
    Set adoConnSeguridad = New ADODB.Connection
    Set adoCommSeguridad = New ADODB.Command
     
    '*** SQLOLEDB - Base de Datos ***
    gstrConnect = "User ID=" & gstrLoginSeguridad & ";Password=" & gstrPasswSeguridad & ";" & _
                        "Data Source=" & strServerSeguridadVa & ";" & _
                        "Initial Catalog=" & strBDSeguridadVa & ";" & _
                        "Application Name=" & App.Title & ";" & _
                        "Auto Translate=False"
    
    With adoConnSeguridad
        .Provider = "SQLOLEDB"
        .ConnectionString = gstrConnect
        .CommandTimeout = 0
        .ConnectionTimeout = 0
        .Open
    End With
    
    Set adoCommSeguridad = New ADODB.Command
    Set adoCommSeguridad.ActiveConnection = adoConnSeguridad
    
    With adoCommSeguridad
        
        '** Obtener Atributos permitidos para el tipo de objeto **
        Set adoAtributoPermitido = New ADODB.Recordset
        .CommandText = "{ call up_SEObtenerAtributoPermitidoObjeto('" & strTipoObjeto & "')}"
    
        Set adoAtributoPermitido = .Execute
         
        If Not adoAtributoPermitido.EOF Then
            Do Until adoAtributoPermitido.EOF
                ReDim Preserve arrAtributosPermitidos(cont)
                arrAtributosPermitidos(cont) = adoAtributoPermitido("CodAtributo")
                cont = cont + 1
                adoAtributoPermitido.MoveNext
            Loop
            adoAtributoPermitido.Close: Set adoAtributoPermitido = Nothing
        
        Else
            
            strEstadoObjeto = "6"
            
            .CommandText = "{ call up_SEActualizarObjetoSistemaLog('" & strCodObjeto & "','" & strEstadoObjeto & "','" & strIdUsuario & "')}"
        
            adoCommSeguridad.Execute
            
            MsgBox "Error: (No existe atributos permitidos para el Tipo de Objeto del Sistema)", vbCritical, gstrNombreEmpresa
            adoAtributoPermitido.Close: Set adoAtributoPermitido = Nothing
            Exit Function
        End If
        
        'inserta atributos
        
       
        '*** Mensaje **
        If strTipoObjeto = Codigo_Tipo_Objeto_Modulo Then
            mensaje = "El acceso a este Módulo no está permitido para este Usuario"
        ElseIf strTipoObjeto = Codigo_Tipo_Objeto_Formulario Then
            mensaje = "El acceso a este Formulario no está permitido para este Usuario"
        End If
        
        '** Validacion **
        For i = 0 To UBound(arrAtributosPermitidos)
            strCodAtributo = Trim(arrAtributosPermitidos(i))
            'No evalua cuando el atributo es no definido
            If strCodAtributo <> Codigo_Atributo_NoDefinido Then
            
                Set adoValorPermiso = New ADODB.Recordset
                .CommandText = "{ call up_SEValidarPermisoSistema('" & Trim(strIdUsuario) & "','" & _
                                Trim(strCodObjeto) & "','" & strCodAtributo & "' ) }"
                Set adoValorPermiso = .Execute
                
                If Not adoValorPermiso.EOF Then
                    Do Until adoValorPermiso.EOF
                        If Not IsNull(adoValorPermiso("ValorAtributo")) Then
                            Select Case strCodAtributo
                                Case Codigo_Atributo_Acceso:
                                    sesionValida = CBool(adoValorPermiso("ValorAtributo"))
                                    If Not sesionValida Then
                                        
'                                        strEstadoObjeto = "2"
'
'                                        .CommandText = "{ call up_SEActualizarObjetoSistemaLog('" & strCodObjeto & "','" & strEstadoObjeto & "','" & strIdUsuario & "')}"
'
'                                        adoCommSeguridad.Execute
                                        
                                        MsgBox mensaje, vbCritical, gstrNombreEmpresa
                                        Exit Function
                                    Else
                                    
'                                        strEstadoObjeto = "1"
'
'                                        .CommandText = "{ call up_SEActualizarObjetoSistemaLog('" & strCodObjeto & "','" & strEstadoObjeto & "','" & strIdUsuario & "')}"
'
'                                        adoCommSeguridad.Execute
                                    
                                    End If
                            End Select
                        
                        Else
                            
'                            strEstadoObjeto = "3"
'
'                            .CommandText = "{ call up_SEActualizarObjetoSistemaLog('" & strCodObjeto & "','" & strEstadoObjeto & "','" & strIdUsuario & "')}"
'
'                            adoCommSeguridad.Execute
             
                            'El comportamiento cuando no esta definido el atributo, en un modulo
                            'es mostrar el error, y en un formulario, dejar su estado normal
                            If strTipoObjeto = Codigo_Tipo_Objeto_Modulo Then
                                MsgBox "(No se ah definido ningun acceso para este usuario)", vbCritical, gstrNombreEmpresa
                                adoValorPermiso.Close: Set adoValorPermiso = Nothing
                                Exit Function
                            End If
                        End If
                        adoValorPermiso.MoveNext
                    Loop
                    adoValorPermiso.Close: Set adoValorPermiso = Nothing
                Else
                
'                    strEstadoObjeto = "4"
'
'                    .CommandText = "{ call up_SEActualizarObjetoSistemaLog('" & strCodObjeto & "','" & strEstadoObjeto & "','" & strIdUsuario & "')}"
'
'                    adoCommSeguridad.Execute
                
                    MsgBox "(No se ha definido ningun acceso para este usuario o el acceso no esta definido en el sistema)", vbCritical, gstrNombreEmpresa
                    adoValorPermiso.Close: Set adoValorPermiso = Nothing
                    Exit Function
                End If
   
            Else
            
                'Si solo existe una atributo y este es "NO DEFINIDO" entonces registro en log.
                If UBound(arrAtributosPermitidos) + 1 = 1 Then
'                    strEstadoObjeto = "5"
'
'                    .CommandText = "{ call up_SEActualizarObjetoSistemaLog('" & strCodObjeto & "','" & strEstadoObjeto & "','" & strIdUsuario & "')}"
'
'                    adoCommSeguridad.Execute
                End If
            
            End If
            
        Next
        
    End With
    
    adoConnSeguridad.Close: Set adoConnSeguridad = Nothing
    
    ValidarPermisoAccesoObjeto = True
    
End Function


Public Sub ValidarPermisoUsoMenu(ByVal strIdUsuario As String, ByVal frmPadre As Object, ByVal strCodObjetoPadre As String, ByVal strseparador As String)
    
    Dim adoAtributoPermitido    As ADODB.Recordset, adoValorPermiso As ADODB.Recordset
    Dim strCodAtributo      As String, arrAtributosPermitidos() As String, strNombreObjeto As String
    Dim cont As Integer, i As Integer
    
    Dim objControl As Control
    
    
    Dim strBDSeguridad As String, strServerSeguridad As String
    Dim strBDSeguridadVa As String, strServerSeguridadVa As String
    Dim strSeguridadActivada As String, strSeguridadActivadaVa As String
    Dim lngValorRetorno As Long         'result of the API functions
    Dim lngCodigoLlave  As Long         'handle of opened key
    Dim vntValor        As Variant      'setting of queried value
    Dim adoConnSeguridad                          As ADODB.Connection
    Dim adoCommSeguridad                          As ADODB.Command
    
    
    strBDSeguridad = "Base de Datos"
    strServerSeguridad = "Servidor"
    strSeguridadActivada = "Activada"

    
    lngValorRetorno = RegOpenKeyEx(HKEY_CURRENT_USER, Clave_Registro_Sistema_Seguridad, 0, KEY_QUERY_VALUE, lngCodigoLlave)
    '*** Nombre de Servidor ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, strServerSeguridad, vntValor)
    strServerSeguridadVa = vntValor
    '*** Nombre de Base de Datos ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, strBDSeguridad, vntValor)
    strBDSeguridadVa = vntValor

    '*** Indicador de seguridad activada ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, strSeguridadActivada, vntValor)
    strSeguridadActivadaVa = vntValor
            
    If strSeguridadActivadaVa = "0" Then 'Seguridad no activada
        Exit Sub
    End If


    Set adoConnSeguridad = New ADODB.Connection
    Set adoCommSeguridad = New ADODB.Command
    
    '*** SQLOLEDB - Base de Datos ***
    gstrConnect = "User ID=" & gstrLoginSeguridad & ";Password=" & gstrPasswSeguridad & ";" & _
                        "Data Source=" & strServerSeguridadVa & ";" & _
                        "Initial Catalog=" & strBDSeguridadVa & ";" & _
                        "Application Name=" & App.Title & ";" & _
                        "Auto Translate=False"
    
    With adoConnSeguridad
        .Provider = "SQLOLEDB"
        .ConnectionString = gstrConnect
        .CommandTimeout = 0
        .ConnectionTimeout = 0
        .Open
    End With
    
    Set adoCommSeguridad = New ADODB.Command
    Set adoCommSeguridad.ActiveConnection = adoConnSeguridad
    
    With adoCommSeguridad
    
        '** Obtener Atributos permitidos para el tipo de objeto **
        Set adoAtributoPermitido = New ADODB.Recordset
        .CommandText = "{ call up_SEObtenerAtributoPermitidoObjeto('" & Codigo_Tipo_Objeto_Menu & "')}"
    
        Set adoAtributoPermitido = .Execute
         
        If Not adoAtributoPermitido.EOF Then
            Do Until adoAtributoPermitido.EOF
                ReDim Preserve arrAtributosPermitidos(cont)
                arrAtributosPermitidos(cont) = adoAtributoPermitido("CodAtributo")
                cont = cont + 1
                adoAtributoPermitido.MoveNext
            Loop
            adoAtributoPermitido.Close: Set adoAtributoPermitido = Nothing
        Else
            MsgBox "Error: (No existe atributos permitidos a Objeto del Sistema)", vbCritical, gstrNombreEmpresa
            adoAtributoPermitido.Close: Set adoAtributoPermitido = Nothing
            Exit Sub
        End If
    
        For Each objControl In frmPadre.Controls
            If TypeOf objControl Is Menu Then
                For i = 0 To UBound(arrAtributosPermitidos)
                    strCodAtributo = Trim(arrAtributosPermitidos(i))
                    Set adoValorPermiso = New ADODB.Recordset
                    strNombreObjeto = strCodObjetoPadre & strseparador & objControl.Name & "(" & objControl.Index & ")"
                    .CommandText = "{ call up_SEValidarPermisoSistema('" & Trim(strIdUsuario) & "','" & _
                                    strNombreObjeto & "','" & strCodAtributo & "' ) }"
                    Set adoValorPermiso = .Execute
                    
                    If Not adoValorPermiso.EOF Then
                        Do Until adoValorPermiso.EOF
                            If Not IsNull(adoValorPermiso("ValorAtributo")) Then
                                Select Case strCodAtributo
                                    Case Codigo_Atributo_Enabled:
                                    objControl.Enabled = CBool(adoValorPermiso("ValorAtributo"))
                                    Case Codigo_Atributo_Visible 'Visible predomina a Enabled
                                    objControl.Visible = CBool(adoValorPermiso("ValorAtributo"))
                                    objControl.Enabled = CBool(adoValorPermiso("ValorAtributo"))
                                End Select
                            End If
                            adoValorPermiso.MoveNext
                        Loop
                        adoValorPermiso.Close: Set adoValorPermiso = Nothing
                    End If
                    
                Next
            End If
        Next
               
    End With
    
    adoConnSeguridad.Close: Set adoConnSeguridad = Nothing
    
End Sub

Public Sub ValidarPermisoUsoControl(ByVal strIdUsuario As String, ByVal frmPadre As Object, ByVal strCodObjetoPadre As String, ByVal strseparador As String)
    
    Dim adoAtributoPermitido    As ADODB.Recordset, adoValorPermiso As ADODB.Recordset
    Dim strCodAtributo      As String, arrAtributosPermitidos() As String, strNombreObjeto As String
    Dim cont As Integer, i As Long, j As Long
    
    Dim objControl As Control
    
    Dim strBDSeguridad As String, strServerSeguridad As String
    Dim strBDSeguridadVa As String, strServerSeguridadVa As String
    Dim strSeguridadActivada As String, strSeguridadActivadaVa As String
    Dim lngValorRetorno As Long         'result of the API functions
    Dim lngCodigoLlave  As Long         'handle of opened key
    Dim vntValor        As Variant      'setting of queried value
    Dim adoConnSeguridad                          As ADODB.Connection
    Dim adoCommSeguridad                          As ADODB.Command
    
'    Exit Sub

    strBDSeguridad = "Base de Datos"
    strServerSeguridad = "Servidor"
    strSeguridadActivada = "Activada"
    
    lngValorRetorno = RegOpenKeyEx(HKEY_CURRENT_USER, Clave_Registro_Sistema_Seguridad, 0, KEY_QUERY_VALUE, lngCodigoLlave)
    '*** Nombre de Servidor ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, strServerSeguridad, vntValor)
    strServerSeguridadVa = vntValor
    '*** Nombre de Base de Datos ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, strBDSeguridad, vntValor)
    strBDSeguridadVa = vntValor
    
    '*** Indicador de seguridad activada ***
    lngValorRetorno = QueryValueEx(lngCodigoLlave, strSeguridadActivada, vntValor)
    strSeguridadActivadaVa = vntValor
            
    If strSeguridadActivadaVa = "0" Then 'Seguridad no activada
        Exit Sub
    End If
    
    Set adoConnSeguridad = New ADODB.Connection
    Set adoCommSeguridad = New ADODB.Command
     
    '*** SQLOLEDB - Base de Datos ***
    gstrConnect = "User ID=" & gstrLoginSeguridad & ";Password=" & gstrPasswSeguridad & ";" & _
                        "Data Source=" & strServerSeguridadVa & ";" & _
                        "Initial Catalog=" & strBDSeguridadVa & ";" & _
                        "Application Name=" & App.Title & ";" & _
                        "Auto Translate=False"
    
    With adoConnSeguridad
        .Provider = "SQLOLEDB"
        .ConnectionString = gstrConnect
        .CommandTimeout = 0
        .ConnectionTimeout = 0
        .Open
    End With
    
    Set adoCommSeguridad = New ADODB.Command
    Set adoCommSeguridad.ActiveConnection = adoConnSeguridad
       
    
    
    With adoCommSeguridad
    
        '** Obtener Atributos permitidos para el tipo de objeto **
        Set adoAtributoPermitido = New ADODB.Recordset
        .CommandText = "{ call up_SEObtenerAtributoPermitidoObjeto('" & Codigo_Tipo_Objeto_Control & "')}"
    
        Set adoAtributoPermitido = .Execute
         
        If Not adoAtributoPermitido.EOF Then
            Do Until adoAtributoPermitido.EOF
                ReDim Preserve arrAtributosPermitidos(cont)
                arrAtributosPermitidos(cont) = adoAtributoPermitido("CodAtributo")
                cont = cont + 1
                adoAtributoPermitido.MoveNext
            Loop
            adoAtributoPermitido.Close: Set adoAtributoPermitido = Nothing
        Else
            MsgBox "Error: (No existe atributos permitidos a Objeto del Sistema)", vbCritical, gstrNombreEmpresa
            adoAtributoPermitido.Close: Set adoAtributoPermitido = Nothing
            Exit Sub
        End If
    
        For Each objControl In frmPadre.Controls
            If TypeOf objControl Is ucBotonEdicion2 Then
                For i = 0 To UBound(arrAtributosPermitidos)
                    strCodAtributo = Trim(arrAtributosPermitidos(i))
                    For j = 0 To objControl.Buttons - 1
                        Set adoValorPermiso = New ADODB.Recordset
                        strNombreObjeto = strCodObjetoPadre & strseparador & objControl.Name & _
                                            "(" & objControl.Button(j).Index & ")"
                        .CommandText = "{ call up_SEValidarPermisoSistema('" & Trim(strIdUsuario) & "','" & _
                                    strNombreObjeto & "','" & strCodAtributo & "' ) }"
                        
                        Set adoValorPermiso = .Execute
                    
                        If Not adoValorPermiso.EOF Then
                            Do Until adoValorPermiso.EOF
                                If Not IsNull(adoValorPermiso("ValorAtributo")) Then
                                    Select Case strCodAtributo
                                        Case Codigo_Atributo_Enabled:
                                        objControl.Button(j).Enabled = CBool(adoValorPermiso("ValorAtributo"))
                                        Case Codigo_Atributo_Visible
                                        objControl.Button(j).Visible = CBool(adoValorPermiso("ValorAtributo"))
                                        objControl.Button(j).Enabled = CBool(adoValorPermiso("ValorAtributo"))
                                    End Select
                                End If
                                adoValorPermiso.MoveNext
                            Loop
                            adoValorPermiso.Close: Set adoValorPermiso = Nothing
                        End If
                    Next
                Next
            Else
                For i = 0 To UBound(arrAtributosPermitidos)
                    strCodAtributo = Trim(arrAtributosPermitidos(i))
                    Set adoValorPermiso = New ADODB.Recordset
                    strNombreObjeto = strCodObjetoPadre & strseparador & objControl.Name
                    
                    .CommandText = "{ call up_SEValidarPermisoSistema('" & Trim(strIdUsuario) & "','" & _
                                    strNombreObjeto & "','" & strCodAtributo & "' ) }"
                        
                    Set adoValorPermiso = .Execute
                    
                    If Not adoValorPermiso.EOF Then
                        Do Until adoValorPermiso.EOF
                            If Not IsNull(adoValorPermiso("ValorAtributo")) Then
                                Select Case strCodAtributo
                                    Case Codigo_Atributo_Enabled:
                                    objControl.Enabled = CBool(adoValorPermiso("ValorAtributo"))
                                    Case Codigo_Atributo_Visible
                                    objControl.Visible = CBool(adoValorPermiso("ValorAtributo"))
                                    objControl.Enabled = CBool(adoValorPermiso("ValorAtributo"))
                                End Select
                            End If
                            adoValorPermiso.MoveNext
                        Loop
                        adoValorPermiso.Close: Set adoValorPermiso = Nothing
                    End If
                Next
            End If
        Next
               
    End With
    
    adoConnSeguridad.Close: Set adoConnSeguridad = Nothing
    
End Sub

Public Function CalculaFechaSiguienteCalendario(ByVal pdatFechaInicial, indexBaseCalculo, indexPeriodoCupon, indexUnidadPeriodo, intUnidadesPeriodo) As Date
    Dim result As Date
    If indexBaseCalculo < 2 Then 'caso 360
        Select Case indexPeriodoCupon
            Case 0
                result = DateAdd("d", 360, pdatFechaInicial)
            Case 1
                result = DateAdd("d", 180, pdatFechaInicial)
            Case 2
                result = DateAdd("d", 90, pdatFechaInicial)
            Case 3
                result = DateAdd("d", 60, pdatFechaInicial)
            Case 4
                result = DateAdd("d", 30, pdatFechaInicial)
            Case 5
                result = DateAdd("d", 15, pdatFechaInicial)
            Case 6
                result = DateAdd("d", 1, pdatFechaInicial)
            Case 7
                Select Case indexUnidadPeriodo
                    Case 0
                        result = DateAdd("d", intUnidadesPeriodo * 360, pdatFechaInicial)
                    Case 1
                        result = DateAdd("d", intUnidadesPeriodo * 180, pdatFechaInicial)
                    Case 2
                        result = DateAdd("d", intUnidadesPeriodo * 90, pdatFechaInicial)
                    Case 3
                        result = DateAdd("d", intUnidadesPeriodo * 60, pdatFechaInicial)
                    Case 4
                        result = DateAdd("d", intUnidadesPeriodo * 30, pdatFechaInicial)
                    Case 5
                        result = DateAdd("d", intUnidadesPeriodo * 15, pdatFechaInicial)
                    Case 6
                        result = DateAdd("d", intUnidadesPeriodo, pdatFechaInicial)
                End Select
        End Select
    Else    'caso actual
        Select Case indexPeriodoCupon
            Case 0
                result = DateAdd("yyyy", 1, pdatFechaInicial)
            Case 1
                result = DateAdd("m", 6, pdatFechaInicial)
            Case 2
                result = DateAdd("m", 3, pdatFechaInicial)
            Case 3
                result = DateAdd("m", 2, pdatFechaInicial)
            Case 4
                result = DateAdd("m", 1, pdatFechaInicial)
            Case 5
                result = DateAdd("d", 15, pdatFechaInicial)
            Case 6
                result = DateAdd("d", 1, pdatFechaInicial)
            Case 7
            Select Case indexUnidadPeriodo
                Case 0
                    result = DateAdd("yyyy", intUnidadesPeriodo, pdatFechaInicial)
                Case 1
                    result = DateAdd("m", intUnidadesPeriodo * 6, pdatFechaInicial)
                Case 2
                    result = DateAdd("m", intUnidadesPeriodo * 3, pdatFechaInicial)
                Case 3
                    result = DateAdd("m", intUnidadesPeriodo * 2, pdatFechaInicial)
                Case 4
                    result = DateAdd("m", intUnidadesPeriodo, pdatFechaInicial)
                Case 5
                    result = DateAdd("d", intUnidadesPeriodo * 15, pdatFechaInicial)
                Case 6
                    result = DateAdd("d", intUnidadesPeriodo, pdatFechaInicial)
            End Select
        End Select
    End If
    
    CalculaFechaSiguienteCalendario = result
    
End Function

 Public Function aFind(rstTemp As ADODB.Recordset, strEncontrar As String, Optional strTipoBusqueda As String = "R") As Boolean
  'Utiliza el método Find basado en la entrada del usuario.
  Dim iNum As Integer
    
  If strTipoBusqueda = "R" Then 'Busqueda Rapida
    For iNum = 0 To rstTemp.Fields.Count - 1
      rstTemp.MoveFirst
      If (Not rstTemp.EOF) And (rstTemp.Fields(iNum).Type = adChar Or rstTemp.Fields(iNum).Type = adVarChar) Then
        rstTemp.Find rstTemp.Fields(iNum).Name & " like '" & strEncontrar & "%'"
        If Not rstTemp.EOF Then Exit For
      End If
    Next iNum
  End If
  
'  If strTipoBusqueda = "E" Then 'Busqueda Especial
'    rstTemp.MoveFirst
'    For iNum = 0 To rstTemp.Fields.Count - 1
'        If bRegistro = True Then rstTemp.MoveFirst
'            If bRegistro = False Then
'               rstTemp.AbsolutePosition = vPosicion
'               rstTemp.MoveNext
'            End If
'            If (Not rstTemp.EOF) And (rstTemp.Fields(iNum).Type = adChar Or rstTemp.Fields(iNum).Type = adVarChar) Then
'               rstTemp.Find rstTemp.Fields(iNum).Name & " Like '%" & strEncontrar & "%'"
'               'vPosicion = rstTemp.Bookmark
'               If Not rstTemp.EOF Then
'                  vPosicion = rstTemp.Bookmark
'                  bRegistro = False
'                  Exit For
'               End If
'            Else
'               rstTemp.MoveLast
'               If Not rstTemp.EOF Then rstTemp.MoveNext
'            End If
'        End If
'    Next iNum
' End If
  
 aFind = Not (rstTemp.EOF)


End Function


'Private Sub ProvisionGastosFondo(strTipoCierre As String, strIndNoIncluyeEnPreCierre As String)
Public Sub ProvisionGastosFondo(ByVal strCodFondo As String, ByVal indEjecucion As String, ByVal strFechaCierre As String, ByVal strFechaSiguiente As String, ByVal strCodMoneda As String, ByVal strCodModulo As String, strTipoCierre As String, strIndNoIncluyeEnPreCierre As String, Optional strNumGasto As String = "", Optional strParCodigoDinamicaGasto As String = "")
  
    Dim adoRegistro             As ADODB.Recordset
    Dim adoConsulta             As ADODB.Recordset
    Dim strCodFile              As String, strCodDetalleFile            As String
    Dim strNumAsiento           As String, strDescripAsiento            As String
    Dim strIndDebeHaber         As String, strDescripMovimiento         As String
    Dim strDescripGasto         As String, strFechaGrabar               As String
    Dim intDiasProvision        As Long, intCantRegistros            As Integer
    Dim intContador             As Integer, intDiasCorridos             As Long
    Dim curMontoRenta           As Currency, curSaldoProvision          As Currency
    Dim curMontoMovimientoMN    As Currency, curMontoMovimientoME       As Currency
    Dim curMontoContable        As Currency, curValorAnterior           As Currency
    Dim curValorActual          As Currency
    Dim dblValorTipoCambio      As Double
    Dim dblValorAjusteProv      As Double
    Dim curValorTotal           As Currency
    Dim intNumDiasPeriodo       As Integer
    Dim intDiasProvision1       As Integer
    Dim strIndUltimoMovimiento  As String
    Dim strIndTipoComision      As String
    Dim strCodSubDetalleFile    As String
    'Dim objParser               As New clsParser
    Dim indCumpleCondicion      As Boolean

    frmMainMdi.stbMdi.Panels(3).Text = "Provisionando Gastos del Fondo..."
    
    Set adoRegistro = New ADODB.Recordset
    Set adoConsulta = New ADODB.Recordset
    With adoComm
    
       
        If strTipoCierre = Codigo_Cierre_Definitivo Then
            .CommandText = "SELECT *,'" & strTipoCierre & "' AS TipoCierre,'" & strFechaCierre & "' AS FechaCierre" & " FROM FondoGasto FG " & _
                 "JOIN FondoGastoPeriodo FGP ON (FG.CodFondo = FGP.CodFondo AND FG.CodAdministradora = FGP.CodAdministradora AND FG.NumGasto = FGP.NumGasto) " & _
                 "WHERE FGP.FechaInicio <= '" & strFechaCierre & "' AND FGP.FechaVencimiento >= '" & strFechaCierre & "' AND FG.CodFondo='" & strCodFondo & "' AND " & _
                "FG.CodAdministradora='" & gstrCodAdministradora & "' AND FG.IndVigente='X' AND FG.IndNoIncluyeEnBalancePreCierre = '" & strIndNoIncluyeEnPreCierre & "' AND " & _
                "FG.CodFile = '" & Codigo_File_Gasto & "'"
        Else
            .CommandText = "SELECT *,'" & strTipoCierre & "' AS TipoCierre,'" & strFechaCierre & "' AS FechaCierre" & " FROM FondoGastoTmp FG " & _
                 "JOIN FondoGastoPeriodoTmp FGP ON (FG.CodFondo = FGP.CodFondo AND FG.CodAdministradora = FGP.CodAdministradora AND FG.NumGasto = FGP.NumGasto) " & _
                 "WHERE FGP.FechaInicio <= '" & strFechaCierre & "' AND FGP.FechaVencimiento >= '" & strFechaCierre & "' AND FG.CodFondo='" & strCodFondo & "' AND " & _
                "FG.CodAdministradora='" & gstrCodAdministradora & "' AND FG.IndVigente='X' AND FG.IndNoIncluyeEnBalancePreCierre = '" & strIndNoIncluyeEnPreCierre & "' AND " & _
                "FG.CodFile = '" & Codigo_File_Gasto & "'"
        End If
                        
        Set adoRegistro = .Execute
        
        Do While Not adoRegistro.EOF
        
            If indEjecucion = "C" And adoRegistro("CodModalidadCalculo") = Codigo_Modalidad_Devengo_Inmediata Then GoTo Siguiente
        
            strCodFile = Trim(adoRegistro("CodFile"))
            
            strDescripGasto = adoRegistro("DescripGasto")
            
            If adoRegistro("CodModalidadCalculo") = Codigo_Modalidad_Devengo_Provision Then
                strCodSubDetalleFile = "001"
            End If
            If adoRegistro("CodModalidadCalculo") = Codigo_Modalidad_Devengo_Inmediata Then
                strCodSubDetalleFile = "002"
            End If
            If adoRegistro("CodModalidadCalculo") = Codigo_Modalidad_Devengo_Ganancia_Diferida Then
                strCodSubDetalleFile = "003"
            End If
             
            .CommandText = "SELECT CodDetalleFile FROM InversionDetalleFile " & _
                "WHERE CodFile='" & strCodFile & "' AND DescripDetalleFile='" & adoRegistro("CodCuenta") & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                strCodDetalleFile = adoConsulta("CodDetalleFile")
            End If
            adoConsulta.Close
            Set adoConsulta = New ADODB.Recordset
            
            If Trim(adoRegistro("CodMoneda")) <> strCodMoneda Then  'Codigo_Moneda_Local
                'Por defecto obtener el tipo de cambio SUNAT de la fecha del documento indicada en el registro de compras
                .CommandText = "SELECT FechaComprobante, CodMonedaPago " & _
                    " FROM RegistroCompra RC " & _
                    " WHERE RC.NumGasto = " & CInt(adoRegistro("NumGasto")) & " AND RC.CodFondo = '" & strCodFondo & "' AND " & _
                    " RC.CodAdministradora = '" & gstrCodAdministradora & "' AND RC.FechaPago = '" & strFechaCierre & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    'Es tipo de cambio SUNAT de la fecha de emision del documento si el documento es factura
                    dblValorTipoCambio = ObtenerTipoCambioMoneda(Codigo_TipoCambio_SBS, Codigo_Valor_TipoCambioVenta, adoConsulta("FechaComprobante"), adoRegistro("CodMoneda"), strCodMoneda)
                Else
                    dblValorTipoCambio = ObtenerTipoCambioMoneda(Codigo_TipoCambio_SBS, Codigo_Valor_TipoCambioVenta, Convertddmmyyyy(strFechaCierre), adoRegistro("CodMoneda"), strCodMoneda)
                End If
                adoConsulta.Close
                Set adoConsulta = New ADODB.Recordset
            Else
                dblValorTipoCambio = 1
            End If
                        
'            '*** Verificar Dinamica Contable ***
'            .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
'                "WHERE TipoOperacion='" & Codigo_Dinamica_Gasto & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
'                strCodDetalleFile & "' OR CodDetalleFile='000') AND CodSubDetalleFile = '" & strCodSubDetalleFile & "' AND CodMoneda = '" & IIf(adoRegistro("CodMoneda") <> Codigo_Moneda_Local, Codigo_Moneda_Extranjero, Codigo_Moneda_Local) & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'            Set adoConsulta = .Execute
'
'            If Not adoConsulta.EOF Then
'                If CInt(adoConsulta("NumRegistros")) > 0 Then
'                    intCantRegistros = CInt(adoConsulta("NumRegistros"))
'                Else
'                    MsgBox "NO EXISTE Dinámica Contable para la provisión", vbCritical
'                    adoConsulta.Close: Set adoConsulta = Nothing
'                    GoTo Siguiente
'                End If
'            End If
'            adoConsulta.Close
                        
            '*** Obtener Descripción del Gasto ***
'            .CommandText = "SELECT DescripCuenta FROM PlanContable WHERE CodCuenta='" & adoRegistro("CodCuenta") & "'"
'            Set adoConsulta = .Execute
'
'            If Not adoConsulta.EOF Then
'                strDescripGasto = Trim(adoConsulta("DescripCuenta")) '& " - " & adoRegistro("CodFondoSerie")
'            End If
'            adoConsulta.Close
'
'
'            '*** Obtener Descripción de Serie ***
'            .CommandText = "SELECT DescripFondoSerie FROM FondoSerie WHERE CodFondo='" & strCodFondo & "' AND " & _
'                           "CodAdministradora = '" & gstrCodAdministradora & "' AND CodFondoSerie = '" & adoRegistro("CodFondoSerie") & "'"
'            Set adoConsulta = .Execute
'
'            If Not adoConsulta.EOF Then
'                strDescripGasto = strDescripGasto & " - " & Trim(adoConsulta("DescripFondoSerie"))
'            End If
'            adoConsulta.Close
            
            
            '*** Obtener las cuentas de inversión ***
            'Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile)
'            Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, adoRegistro("CodMoneda"), strCodSubDetalleFile)
            
            curSaldoProvision = 0

            If strIndNoIncluyeEnPreCierre = Valor_Indicador Then
            
                'ESTA LOGICA HAY QUE VER COMO SE CAMBIA!
                If Trim(adoRegistro("CodCuenta")) = "6730110" Then
                    strIndTipoComision = "F"
                ElseIf Trim(adoRegistro("CodCuenta")) = "6730120" Then
                    strIndTipoComision = "V"
                End If
                
                '*** Obtener Saldo de Gasto NO CRISTALIZADO ***
                If strTipoCierre = Codigo_Cierre_Simulacion Then
                    curSaldoProvision = ObtenerComisionNoCristalizadaTmp(strCodFondo, gstrCodAdministradora, adoRegistro("CodFondoSerie"), adoRegistro("CodMoneda"), strFechaCierre, strIndTipoComision)
                Else
                    curSaldoProvision = ObtenerComisionNoCristalizada(strCodFondo, gstrCodAdministradora, adoRegistro("CodFondoSerie"), adoRegistro("CodMoneda"), strFechaCierre, strIndTipoComision)
                End If
            
            Else

                '*** Obtener Saldo de Gasto NO CRISTALIZADO ***
                .CommandText = "{ call up_GNObtieneSaldoGasto ('" & _
                                strCodFondo & "','" & gstrCodAdministradora & "'," & _
                                adoRegistro("NumGasto") & "," & adoRegistro("NumPeriodo") & ",'" & strFechaCierre & "','" & _
                                strTipoCierre & "') }"
                                
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curSaldoProvision = CDbl(adoConsulta("MontoDevengoAcumulado"))
                Else
                    curSaldoProvision = 0
                End If
                adoConsulta.Close
            
            End If
            
'            '*** Obtener Saldo de Inversión ***
'            If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
'                .CommandText = "SELECT SaldoFinalMN Saldo "
'            Else
'                .CommandText = "SELECT SaldoFinalME Saldo "
'            End If
'
'            If strTipoCierre = Codigo_Cierre_Simulacion Then
'                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
'            Else
'                .CommandText = .CommandText & "FROM PartidaContableSaldos "
'            End If
'
'            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
'                "CodCuenta='" & strCtaProvGasto & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
'                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                "CodFondo='" & strCodFondo & "' AND CodMonedaContable = '" & strCodMoneda & "'"
'            Set adoConsulta = .Execute
'
'            If Not adoConsulta.EOF Then
'                curSaldoProvision = CDbl(adoConsulta("Saldo"))
'            Else
'                curSaldoProvision = 0
'            End If
'            adoConsulta.Close
                                
            intDiasProvision = DateDiff("d", adoRegistro("FechaInicio"), adoRegistro("FechaVencimiento")) + 1
            intDiasCorridos = DateDiff("d", adoRegistro("FechaInicio"), gdatFechaActual) + 1
            
            curValorAnterior = curSaldoProvision
            
'            If adoRegistro("CodTipoCalculo") = Codigo_Tipo_Gasto_Periodico Then
'                Set adoConsulta = New ADODB.Recordset
'
''                '*** Obtener el número de días del periodo de devengo ***
''                .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' AND CodParametro='" & adoRegistro("CodPeriodoGasto") & "'"
''                Set adoConsulta = adoComm.Execute
''
''                If Not adoConsulta.EOF Then
''                    intDiasProvision = CInt(adoConsulta("ValorParametro")) '*** Días del periodo  ***
''                Else
''                    intDiasProvision = 0
''                End If
''                adoConsulta.Close
'
'                '*** Obtener el número de días del periodo de devengo ***
''            Else
''                intNumDiasPeriodo = 0
''                intDiasProvision = 0
'            End If
            
            'JAFR: Aqui CodFrecuenciaCalculo se refiere a la FRECUENCIA de devengo.
            .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' AND CodParametro='" & adoRegistro("CodFrecuenciaCalculo") & "'"
            Set adoConsulta = adoComm.Execute
    
            If Not adoConsulta.EOF Then
                intNumDiasPeriodo = CInt(adoConsulta("ValorParametro")) '*** Días del periodo  ***
            Else
                intNumDiasPeriodo = 0
            End If
            adoConsulta.Close: Set adoConsulta = Nothing
            
            
            'JCB
            Dim MyParser As New clsParser
            Dim strIndCondicional As String
            Dim strFormulaMonto As String
            Dim strFormulaCondicion As String
            Dim strCodFormulaDatos As String
            Dim strParametros As String
            Dim strMsgError As String
            Dim strFechaPago As String
            Dim strFechaVencimiento As String
                                   
            indCumpleCondicion = True
            
            curValorTotal = 0

            If adoRegistro("CodTipoGasto") = Codigo_Tipo_Calculo_Fijo Then 'Fijo
'                If adoRegistro("CodTipoValor") = Codigo_Tipo_Costo_Porcentaje Then
'                    curValorTotal = CalculoInteres(adoRegistro("PorcenGasto"), adoRegistro("CodTipoTasa"), adoRegistro("CodPeriodoTasa"), adoRegistro("CodBaseAnual"), adoRegistro("MontoBaseCalculo"), adoRegistro("FechaInicio"), adoRegistro("FechaVencimiento"))
'                Else
                 curValorTotal = adoRegistro("MontoGasto")
'                End If
            Else ' Calculamos ejecutando la formula
                'Traemos los datos de la formula
                .CommandText = "SELECT FormulaMonto, indCondicion, FormulaCondicion, CodFormulaDatos  FROM Formula WHERE CodFormula='" & adoRegistro("CodFormula") & "'"
                Set adoConsulta = adoComm.Execute
        
                If Not adoConsulta.EOF Then
                    strIndCondicional = "" & adoConsulta("indCondicion")
                    strFormulaMonto = "" & adoConsulta("FormulaMonto")
                    strFormulaCondicion = "" & adoConsulta("FormulaCondicion")
                    strCodFormulaDatos = "" & adoConsulta("CodFormulaDatos")
                Else
                    strIndCondicional = ""
                    strFormulaMonto = ""
                    strFormulaCondicion = ""
                    strCodFormulaDatos = ""
                End If
                adoConsulta.Close: Set adoConsulta = Nothing
                
'                strParametros = strCodFondo & "|" & gstrCodAdministradora & "|" & adoRegistro("CodFondoSerie") & "|" & strFechaCierre & "|" & strTipoCierre
                
                'Parametros configurados
'                MsgBox adoRegistro.GetRows(1, intFila)

                If strIndCondicional = Valor_Indicador Then
                    indCumpleCondicion = MyParser.ParseExpression(strFormulaCondicion, strCodFormulaDatos, adoRegistro, strMsgError)
                    If strMsgError <> "" Then
                        MsgBox strMsgError, vbCritical
                        Exit Sub
                    End If
                End If
                If indCumpleCondicion Then
                    curValorTotal = Round(MyParser.ParseExpression(strFormulaMonto, strCodFormulaDatos, adoRegistro, strMsgError), 2)

                    If strMsgError <> "" Then
                        MsgBox strMsgError, vbCritical
                        Exit Sub
                    End If
                End If
            End If
            
            Set MyParser = Nothing
                
            If indCumpleCondicion Then
            
                 '--- AGREGANDO EL IGV
'                If adoRegistro("CodAfectacion") = Codigo_Afecto Then   'new
'                    'Si hay impuesto a provisionar con el gasto (asumimos que se trata del impuesto IGV)
'                    curValorTotal = Round(curValorTotal * (1 + gdblTasaIgv), 2)  'new
'                End If
          
                'UltimoDiaMes
                'Para el calculo prorratea sobre la base de Actual/x --osea sobre el numero real de dias del mes!
                If adoRegistro("CodFormaCalculo") = Codigo_Tipo_Devengo_Alicuota_Lineal Then
                    If intDiasProvision <> 0 And intNumDiasPeriodo <> 0 And curValorTotal <> 0 Then
                        If intDiasProvision Mod intNumDiasPeriodo = 0 Then
                            curMontoRenta = Round(curValorTotal / (intDiasProvision / intNumDiasPeriodo), 2)
                        Else
                            curMontoRenta = 0
                        End If
                    Else
                        curMontoRenta = 0
                    End If
                    
                    curValorActual = curSaldoProvision + curMontoRenta
                    
                ElseIf adoRegistro("CodFormaCalculo") = Codigo_Tipo_Devengo_Alicuota_Incremental Then
                    If intDiasProvision <> 0 And intNumDiasPeriodo <> 0 And curValorTotal <> 0 Then
                        If intDiasProvision Mod intNumDiasPeriodo = 0 Then
                            curMontoRenta = Round(curValorTotal / (intDiasProvision / intNumDiasPeriodo), 2) * intDiasCorridos - curSaldoProvision
                        Else
                            curMontoRenta = 0
                        End If
                    Else
                        curMontoRenta = 0
                    End If
                   
                    curValorActual = curSaldoProvision + curMontoRenta
                ElseIf adoRegistro("CodFormaCalculo") = Codigo_Tipo_Devengo_Valor_Total_Incremental Then
                    curValorActual = curValorTotal
                    curMontoRenta = curValorActual - curSaldoProvision
                Else 'No Porratea, es inmediato : Codigo_Tipo_Devengo_Valor_Total
                    curMontoRenta = curValorTotal
                    
                    curValorActual = curSaldoProvision + curMontoRenta
                    curValorTotal = curValorActual
                End If
                
                '08/07/2010
                '--- ADICIONAR EL IGV AL GASTO SI ESTA DEFINIDO ASI
                ' Se hizo en cada if
                
                '--- FIN : ADICIONAR EL IGV AL GASTO SI ESTA DEFINIDO ASI
                
                'Control de remanentes
                If adoRegistro("FechaVencimiento") = gdatFechaActual Then
                    If (curValorTotal - curValorActual) <> 0 Then
                        dblValorAjusteProv = (curValorTotal - curValorActual)
                        curMontoRenta = curMontoRenta + dblValorAjusteProv
                        curValorActual = curValorActual + dblValorAjusteProv
                    End If
               
                    'Inserta el registro de la provision del gasto del dia, asi sea cero.
                    .CommandText = "{ call up_GNManFondoGastoDevengo('"
                    If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNManFondoGastoDevengoTmp('"  '*** Simulación ***
                    
                    .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "'," & _
                                    adoRegistro("NumGasto") & "," & adoRegistro("NumPeriodo") & ",'" & strFechaCierre & "','" & _
                                    adoRegistro("CodFondoSerie") & "','" & adoRegistro("CodMoneda") & "'," & _
                                    curMontoRenta & ") }"
                                    
                    adoConn.Execute .CommandText
                    
                    'JAFR 09/03/11: la grabacion de la orden de pago estaba aqui, se movió con su propia condición.
                Else
                    'Inserta el registro de la provision del gasto del dia, asi sea cero.
                    .CommandText = "{ call up_GNManFondoGastoDevengo('"
                    If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNManFondoGastoDevengoTmp('"  '*** Simulación ***
                    
                    .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "'," & _
                                    adoRegistro("NumGasto") & "," & adoRegistro("NumPeriodo") & ",'" & strFechaCierre & "','" & _
                                    adoRegistro("CodFondoSerie") & "','" & adoRegistro("CodMoneda") & "'," & _
                                    curMontoRenta & ") }"
                                    
                    adoConn.Execute .CommandText
                    
                    'Inicializa el registro de la provision del gasto para el dia siguiente.
                    .CommandText = "{ call up_GNProcFondoGastoDevengoInicial('"
                    .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "'," & _
                                    adoRegistro("NumGasto") & "," & adoRegistro("NumPeriodo") & ",'" & strFechaCierre & "','" & _
                                    adoRegistro("CodFondoSerie") & "','" & adoRegistro("CodMoneda") & "','" & strTipoCierre & "') }"
                                    
                    adoConn.Execute .CommandText
                End If
                                            
                'CONTABILIZAR LA PROVISIÓN DEL GASTO -- 35: TipoOperacion ProvisionGasto
                .CommandText = "{ call up_ACProcContabilizarOperacion('"
                .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                strFechaCierre & "','" & strCodFile & "','" & Format(CStr(adoRegistro("NumGasto")), "0000000000") & "','" & Codigo_Caja_Provision_Gasto & "') }"
                adoConn.Execute .CommandText
                                            
                'JAFR 09/03/11 la grabacion de la orden de pago (solo si es provisión)
                If (adoRegistro("CodModalidadCalculo") = Codigo_Modalidad_Devengo_Provision) And (adoRegistro("FechaPago") = gdatFechaActual) Then
                    'Genera orden de pago del gasto
                    Dim montoOrdenPago As Double
                    
                    If adoRegistro("CodModalidadPago") = Codigo_Modalidad_Pago_Vencimiento Then
                        'JAFR: Caso de gasto al vencimiento:
                        montoOrdenPago = curValorActual
                    Else
                        'JAFR: caso de gasto adelantado
                        montoOrdenPago = curValorTotal
                    End If
                        
                    
                    .CommandText = "{ call up_GNManOrdenPago('"
                    If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNManOrdenPagoTmp('"  '*** Simulación ***
                   
                    strFechaPago = Convertyyyymmdd(adoRegistro("FechaPago"))
                    strFechaVencimiento = Convertyyyymmdd(adoRegistro("FechaVencimiento"))
                   
                    .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "'," & _
                                    "NULL," & adoRegistro("NumGasto") & "," & adoRegistro("NumPeriodo") & ",'" & _
                                    adoRegistro("CodMoneda") & "','" & montoOrdenPago & "'," & _
                                    montoOrdenPago & ",'" & strFechaVencimiento & "','" & _
                                    strFechaPago & "',NULL,NULL,NULL,'01','I') }"
                                    
                    adoConn.Execute .CommandText
                End If
                'Fin JAFR
                
                '*** Provisión ***
                If curMontoRenta <> 0 Then
                    strDescripAsiento = "Provisión" & Space(1) & strDescripGasto
                    strDescripMovimiento = strDescripGasto
                    If curMontoRenta > 0 Then strDescripMovimiento = strDescripGasto
                                                    
                    .CommandType = adCmdStoredProc
                    '*** Obtener el número del parámetro **
                    .CommandText = "up_ACObtenerUltNumero"
                    If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "up_GNObtenerUltNumeroTmp"  '*** Simulación ***
                    
                    .Parameters.Append .CreateParameter("CodFondo", adChar, adParamInput, 3, strCodFondo)
                    .Parameters.Append .CreateParameter("CodAdministradora", adChar, adParamInput, 3, gstrCodAdministradora)
                    .Parameters.Append .CreateParameter("CodParametro", adChar, adParamInput, 6, Valor_NumComprobante)
                    .Parameters.Append .CreateParameter("NuevoNumero", adChar, adParamOutput, 10, "")
                    .Execute
                    
                    If Not .Parameters("NuevoNumero") Then
                        strNumAsiento = .Parameters("NuevoNumero").Value
                        .Parameters.Delete ("CodFondo")
                        .Parameters.Delete ("CodAdministradora")
                        .Parameters.Delete ("CodParametro")
                        .Parameters.Delete ("NuevoNumero")
                    End If
                    
                    .CommandType = adCmdText
                                                    
                    'On Error GoTo Ctrl_Error
                    
                    '*** Contabilizar ***
                    strFechaGrabar = strFechaCierre & Space(1) & Format(Time, "hh:mm")
                    strIndUltimoMovimiento = ""
                    
                    '*** Cabecera ***
'''                    .CommandText = "{ call up_ACAdicAsientoContable('"
'''                    If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableTmp('"  '*** Simulación ***
'''
'''                    .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
'''                        strFechaGrabar & "','" & _
'''                        gstrPeriodoActual & "','" & gstrMesActual & "','" & Tipo_Asiento_Provision_Gastos_Proveedores & "','" & _
'''                        strDescripAsiento & "','" & Trim(adoRegistro("CodMoneda")) & "','" & _
'''                        Codigo_Moneda_Local & "','',''," & _
'''                        CDec(curMontoRenta) & ",'" & Estado_Activo & "'," & _
'''                        intCantRegistros & ",'" & _
'''                        strFechaCierre & Space(1) & Format(Time, "hh:ss") & "','" & _
'''                        strCodModulo & "',''," & _
'''                        CDec(dblValorTipoCambio) & ",'','','" & _
'''                        strDescripAsiento & "','','X','') }"
'''                    adoConn.Execute .CommandText
'''
'''                    '*** Detalle ***
'''                    .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
'''                        "WHERE TipoOperacion='" & Codigo_Dinamica_Gasto & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
'''                        strCodDetalleFile & "' OR CodDetalleFile='000') AND CodSubDetalleFile = '" & strCodSubDetalleFile & "' AND CodMoneda = '" & IIf(adoRegistro("CodMoneda") <> Codigo_Moneda_Local, Codigo_Moneda_Extranjero, Codigo_Moneda_Local) & "' AND CodAdministradora='" & gstrCodAdministradora & "' " & _
'''                        "ORDER BY NumSecuencial"
'''                    Set adoConsulta = .Execute
'''
'''                    Do While Not adoConsulta.EOF
'''
'''                        Select Case Trim(adoConsulta("TipoCuentaInversion"))
'''                            Case Codigo_CtaInversion
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaProvInteres
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaInteres
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaInteresVencido
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaVacCorrido
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaXPagar
'''                                curMontoMovimientoMN = curMontoRenta
'''                           '     If adoRegistro("CodAfectacion") = Codigo_Afecto Then
''''                                    If adoRegistro("IndGeneraCreditoFiscal") = Valor_Indicador Then
'''                                        curMontoMovimientoMN = curMontoRenta
''''                                    Else
''' '                                       curMontoMovimientoMN = Round(curMontoMovimientoMN * (1 + gdblTasaIgv), 2)
'''                                    'End If
'''                               ' End If
'''
'''                            Case Codigo_CtaXCobrar
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaInteresCorrido
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaProvReajusteK
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaReajusteK
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaProvFlucMercado
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaFlucMercado
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaProvInteresVac
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaInteresVac
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaIntCorridoK
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaProvFlucK
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaFlucK
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaInversionTransito
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaProvGasto
'''                                curMontoMovimientoMN = curMontoRenta
''''                                If adoRegistro("CodAfectacion") = Codigo_Afecto Then   'new
''''                                    'Si hay impuesto a provisionar con el gasto (asumimos que se trata del impuesto IGV)
''''                                    curMontoMovimientoMN = Round(curMontoMovimientoMN / (1 + gdblTasaIgv), 2) 'new
''''                                End If
'''
'''                            Case Codigo_CtaIngresoOperacional
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaCosto
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaCostoSAB
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaCostoBVL
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaCostoCavali
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaCostoFondoLiquidacion
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaCostoGastosBancarios
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaCostoComisionEspecial
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaCostoFondoGarantia
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaCostoConasev
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                            Case Codigo_CtaComision
'''                                curMontoMovimientoMN = curMontoRenta
'''                                '--- QUITANDO EL IGV
''''                                If adoRegistro("CodAfectacion") = Codigo_Afecto Then
''''                                    'Si hay impuesto a provisionar con el gasto (asumimos que se trata del impuesto IGV)
''''                                    curMontoMovimientoMN = Round(curMontoMovimientoMN / (1 + gdblTasaIgv), 2) 'new
''''                                End If
'''
'''                            Case Codigo_CtaImpuesto
'''                                curMontoMovimientoMN = curMontoRenta
'''                                If adoRegistro("CodAfectacion") = Codigo_Afecto Then
'''                                    'Si hay impuesto a provisionar con el gasto (asumimos que se trata del impuesto IGV)
'''                                    If adoRegistro("IndGeneraCreditoFiscal").Value = Valor_Indicador Then 'genera credito fiscal
'''                                        curMontoMovimientoMN = Round(curMontoMovimientoMN * gdblTasaIgv, 2) 'Round((curMontoMovimientoMN / (1 + gdblTasaIgv)) * gdblTasaIgv, 2)   'new
'''                                    Else
'''                                        curMontoMovimientoMN = 0#
'''                                    End If
'''                                Else
'''                                    curMontoMovimientoMN = 0#
'''                                End If
'''
'''                            Case Codigo_CtaGastoImpuesto
'''                                curMontoMovimientoMN = curMontoRenta
'''                                If adoRegistro("CodAfectacion") = Codigo_Afecto Then
'''                                    'Si hay impuesto a provisionar con el gasto (asumimos que se trata del impuesto IGV)
'''                                    If adoRegistro("IndGeneraCreditoFiscal") = Valor_Indicador Then
'''                                        curMontoMovimientoMN = 0#
'''                                    Else
'''                                        curMontoMovimientoMN = Round((curMontoMovimientoMN / (1 + gdblTasaIgv)) * gdblTasaIgv, 2)   'new
'''                                    End If
'''                                Else
'''                                    curMontoMovimientoMN = 0#
'''                                End If
'''
'''                            'JAFR 09/03/11: Gastos carga diferida
'''                            Case Codigo_CtaCargaDiferidaGasto 'equivalente a la cuenta 46 de provision normal!
'''                                curMontoMovimientoMN = curMontoRenta
'''
'''                                If adoRegistro("CodAfectacion") = Codigo_Afecto Then   'new
'''                                    curMontoMovimientoMN = Round(curMontoMovimientoMN / (1 + gdblTasaIgv), 2) 'new
'''                                End If
'''
'''                        End Select
'''
'''                        strIndDebeHaber = Trim(adoConsulta("IndDebeHaber"))
'''                        If strIndDebeHaber = "H" Then
'''                            curMontoMovimientoMN = curMontoMovimientoMN * -1
'''                            If curMontoMovimientoMN > 0 Then strIndDebeHaber = "D"
'''                        ElseIf strIndDebeHaber = "D" Then
'''                            If curMontoMovimientoMN < 0 Then strIndDebeHaber = "H"
'''                        End If
'''
'''                        If strIndDebeHaber = "T" Then
'''                            If curMontoMovimientoMN > 0 Then
'''                                strIndDebeHaber = "D"
'''                            Else
'''                                strIndDebeHaber = "H"
'''                            End If
'''                        End If
'''                        strDescripMovimiento = Trim(adoConsulta("DescripDinamica"))
'''                        curMontoMovimientoME = 0
'''                        curMontoContable = curMontoMovimientoMN
'''
'''                        If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
'''                            curMontoContable = Round(curMontoMovimientoMN * CDec(dblValorTipoCambio), 2)
'''                            curMontoMovimientoME = curMontoMovimientoMN
'''                            curMontoMovimientoMN = 0
'''                        End If
'''
'''                        '*** Movimiento ***
'''                        If curMontoContable <> 0 Then
'''                            .CommandText = "{ call up_ACProcAsientoContableDetalle1('"
'''                            'If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_ACProcAsientoContableDetalleTmp('"
'''
'''    '                        .CommandText = .CommandText & strNumAsiento & "','" & strCodFondo & "','" & _
'''    '                            gstrCodAdministradora & "'," & _
'''    '                            CInt(adoConsulta("NumSecuencial")) & ",'" & _
'''    '                            strFechaGrabar & "','" & _
'''    '                            gstrPeriodoActual & "','" & _
'''    '                            gstrMesActual & "','" & _
'''    '                            strDescripMovimiento & "','" & _
'''    '                            strIndDebeHaber & "','" & _
'''    '                            Trim(adoConsulta("CodCuenta")) & "','" & _
'''    '                            Trim(adoRegistro("CodMoneda")) & "'," & _
'''    '                            CDec(curMontoMovimientoMN) & "," & _
'''    '                            CDec(curMontoMovimientoME) & "," & _
'''    '                            CDec(curMontoContable) & ",'" & _
'''    '                            Trim(adoRegistro("CodFile")) & "','" & _
'''    '                            Trim(adoRegistro("CodAnalitica")) & "') }"
'''    '                        adoConn.Execute .CommandText
'''
'''                            .CommandText = .CommandText & strNumAsiento & "','" & strCodFondo & "','" & _
'''                                gstrCodAdministradora & "'," & _
'''                                CInt(adoConsulta("NumSecuencial")) & ",'" & _
'''                                strFechaGrabar & "','" & _
'''                                gstrPeriodoActual & "','" & _
'''                                gstrMesActual & "','" & _
'''                                strDescripMovimiento & "','" & _
'''                                strIndDebeHaber & "','" & _
'''                                Trim(adoConsulta("CodCuenta")) & "','" & _
'''                                Trim(adoRegistro("CodMoneda")) & "'," & _
'''                                CDec(curMontoMovimientoMN) & "," & _
'''                                CDec(curMontoMovimientoME) & "," & _
'''                                CDec(curMontoContable) & ",'" & _
'''                                Trim(adoRegistro("CodFile")) & "','" & _
'''                                Trim(adoRegistro("CodAnalitica")) & "','" & _
'''                                strIndUltimoMovimiento & "','" & strFechaCierre & "','" & gstrCodClaseTipoCambioOperacionFondo & "','" & _
'''                                gstrValorTipoCambioOperacion & "',0,'" & XML_TipoCambioReemplazo & "','" & _
'''                                XML_MonedaContable & "','" & XML_MontoMovimientoContable & "','" & strTipoCierre & "') }"
'''
'''                            adoConn.Execute .CommandText
'''
'''    '                        '*** Saldos ***
'''    '                        .CommandText = "{ call up_ACGenPartidaContableSaldos('"
'''    '                        If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNGenPartidaContableSaldosTmp('"
'''    '
'''    '                        .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & _
'''    '                            gstrPeriodoActual & "','" & gstrMesActual & "','" & _
'''    '                            Trim(adoConsulta("CodCuenta")) & "','" & _
'''    '                            Trim(adoRegistro("CodFile")) & "','" & _
'''    '                            Trim(adoRegistro("CodAnalitica")) & "','" & _
'''    '                            strFechaCierre & "','" & _
'''    '                            strFechaSiguiente & "'," & _
'''    '                            CDec(curMontoMovimientoMN) & "," & _
'''    '                            CDec(curMontoMovimientoME) & "," & _
'''    '                            CDec(curMontoContable) & ",'" & _
'''    '                            strIndDebeHaber & "','" & _
'''    '                            Trim(adoRegistro("CodMoneda")) & "') }"
'''    '                        adoConn.Execute .CommandText
'''
'''                            '*** Validar valor de cuenta contable ***
'''                            If Trim(adoConsulta("CodCuenta")) = Valor_Caracter Then
'''                                MsgBox "Registro Nro. " & CStr(intContador) & " de Asiento Contable no tiene cuenta asignada", vbCritical, "Valorización Depósitos"
'''                                gblnRollBack = True
'''                                Exit Sub
'''                            End If
'''                        End If
'''                        adoConsulta.MoveNext
'''                    Loop
'''                    adoConsulta.Close: Set adoConsulta = Nothing
'''
'''                    '-- Verifica y ajusta posibles descuadres
'''                    .CommandText = "{ call up_ACProcAsientoContableAjuste('"
'''                    If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_ACProcAsientoContableAjusteTmp('"  '*** Simulación ***
'''                    .CommandText = .CommandText & strCodFondo & "','" & _
'''                            gstrCodAdministradora & "','" & _
'''                            strNumAsiento & "') }"
'''                    adoConn.Execute .CommandText
'''
'''                    '*** Actualizar el número del parámetro **
'''                    .CommandText = "{ call up_ACActUltNumero('"
'''                    If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNActUltNumeroTmp('"
'''
'''                    .CommandText = .CommandText & strCodFondo & "','" & _
'''                        gstrCodAdministradora & "','" & _
'''                        Valor_NumComprobante & "','" & _
'''                        strNumAsiento & "') }"
'''                    adoConn.Execute .CommandText
'''
                
                    
                    If Convertyyyymmdd(adoRegistro("FechaVencimiento")) = strFechaCierre And _
                       Convertyyyymmdd(adoRegistro("FechaFinal")) = strFechaCierre Then
                        If strTipoCierre = Codigo_Cierre_Definitivo Then
                            .CommandText = "UPDATE FondoGasto SET "
                        Else
                            .CommandText = "UPDATE FondoGastoTmp SET "
                        End If
                       
                        .CommandText = .CommandText & "IndVigente='' " & _
                                        "WHERE NumGasto=" & adoRegistro("NumGasto") & " AND CodCuenta='" & Trim(adoRegistro("CodCuenta")) & _
                                        "' AND CodFondo='" & adoRegistro("CodFondo") & "' AND CodAdministradora='" & adoRegistro("CodAdministradora") & "'"
                        adoConn.Execute .CommandText
                    End If
                    
'                    If strIndNoIncluyeEnPreCierre = "X" Then
'                        adoComm.CommandText = "{ call up_GNActComisionAdministradora('" & _
'                            strCodFondo & "','" & gstrCodAdministradora & "','" & _
'                            strFechaCierre & "','" & strFechaSiguiente & "','" & strCodMoneda & "','"
'
'                       If strTipoCierre = Codigo_Cierre_Definitivo Then
'                            adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Definitivo & "') }"
'                       Else
'                            adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Simulacion & "') }"
'                       End If
'
'                        adoConn.Execute .CommandText
'                    End If
                    
                End If
                                            
            End If 'indCumpleCondicion = true
Siguiente:
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    Exit Sub
  
Ctrl_Error:
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
 
End Sub

Public Function ObtenerComisionNoCristalizada(strCodFondo As String, strCodAdministradora As String, strCodFondoSerie As String, strCodMoneda As String, strFechaDevengo As String, strIndTipoComision As String)

    Dim dblMontoDevengoAcumulado As Double

    dblMontoDevengoAcumulado = 0

    With adoComm
        '*** Obtener Secuencial ***
        .CommandType = adCmdStoredProc
    
        .CommandText = "up_ACObtenerComisionNoCristalizada"
        .Parameters.Append .CreateParameter("@CodFondo", adChar, adParamInput, 3, strCodFondo)
        .Parameters.Append .CreateParameter("@CodAdministradora", adChar, adParamInput, 3, strCodAdministradora)
        .Parameters.Append .CreateParameter("@CodFondoSerie", adChar, adParamInput, 3, strCodFondoSerie)
        .Parameters.Append .CreateParameter("@CodMoneda", adChar, adParamInput, 2, strCodMoneda)
        .Parameters.Append .CreateParameter("@FechaDevengo", adChar, adParamInput, 10, strFechaDevengo)
        .Parameters.Append .CreateParameter("@IndTipoComision", adChar, adParamInput, 1, strIndTipoComision)
        .Parameters.Append .CreateParameter("@MontoDevengoAcumulado", adDouble, adParamOutput, 18, dblMontoDevengoAcumulado)
        .Execute
    
        dblMontoDevengoAcumulado = .Parameters("@MontoDevengoAcumulado").Value
            
        .Parameters.Delete ("@CodFondo")
        .Parameters.Delete ("@CodAdministradora")
        .Parameters.Delete ("@FechaDevengo")
        .Parameters.Delete ("@CodFondoSerie")
        .Parameters.Delete ("@CodMoneda")
        .Parameters.Delete ("@IndTipoComision")
        .Parameters.Delete ("@MontoDevengoAcumulado")
    
        .CommandType = adCmdText
        .Parameters.Refresh  'colocar esto sino se cae cuando es llamado mas de una vez....
    
    End With

    ObtenerComisionNoCristalizada = dblMontoDevengoAcumulado


End Function

Public Function ObtenerComisionNoCristalizadaTmp(strCodFondo As String, strCodAdministradora As String, strCodFondoSerie As String, strCodMoneda As String, strFechaDevengo As String, strIndTipoComision As String)

    Dim dblMontoDevengoAcumulado As Double

    dblMontoDevengoAcumulado = 0

    With adoComm
        '*** Obtener Secuencial ***
        .CommandType = adCmdStoredProc
    
        .CommandText = "up_ACObtenerComisionNoCristalizadaTmp"
        .Parameters.Append .CreateParameter("@CodFondo", adChar, adParamInput, 3, strCodFondo)
        .Parameters.Append .CreateParameter("@CodAdministradora", adChar, adParamInput, 3, strCodAdministradora)
        .Parameters.Append .CreateParameter("@CodFondoSerie", adChar, adParamInput, 3, strCodFondoSerie)
        .Parameters.Append .CreateParameter("@CodMoneda", adChar, adParamInput, 2, strCodMoneda)
        .Parameters.Append .CreateParameter("@FechaDevengo", adChar, adParamInput, 10, strFechaDevengo)
        .Parameters.Append .CreateParameter("@IndTipoComision", adChar, adParamInput, 1, strIndTipoComision)
        .Parameters.Append .CreateParameter("@MontoDevengoAcumulado", adDouble, adParamOutput, 18, dblMontoDevengoAcumulado)
        .Execute
    
        dblMontoDevengoAcumulado = .Parameters("@MontoDevengoAcumulado").Value
            
        .Parameters.Delete ("@CodFondo")
        .Parameters.Delete ("@CodAdministradora")
        .Parameters.Delete ("@FechaDevengo")
        .Parameters.Delete ("@CodFondoSerie")
        .Parameters.Delete ("@CodMoneda")
        .Parameters.Delete ("@IndTipoComision")
        .Parameters.Delete ("@MontoDevengoAcumulado")
    
        .CommandType = adCmdText
        .Parameters.Refresh  'colocar esto sino se cae cuando es llamado mas de una vez....
    
    End With

    ObtenerComisionNoCristalizadaTmp = dblMontoDevengoAcumulado

End Function

Public Sub ValidaExisteTipoCambio(strCodTipoCambio As String, strFechaConsulta As String)
    Dim adoConsulta As ADODB.Recordset
    Dim strTipoCambio As String
    
    adoComm.CommandText = "SELECT DescripParametro FROM AuxiliarParametro WHERE CodTipoParametro = 'TIPCAM' AND CodParametro = '01'"
    Set adoConsulta = adoComm.Execute
    If Not adoConsulta.EOF Then
        strTipoCambio = adoConsulta("DescripParametro") & " "
    Else
        strTipoCambio = Valor_Caracter
    End If
    
    ' Por ahora solo se comprueba de dolares a soles
    adoComm.CommandText = "SELECT ValorTipoCambioCompra, ValorTipoCambioVenta FROM TipoCambioFondo " & _
    "WHERE CodTipoCambio = '" & strCodTipoCambio & "' and FechaTipoCambio = '" & strFechaConsulta & "' " & _
    "AND CodMoneda = '02' and CodMonedaCambio = '01' "
    
    Set adoConsulta = adoComm.Execute
    
    If Not adoConsulta.EOF Then
        If adoConsulta("ValorTipoCambioCompra").Value = 0 Then
            MsgBox "El tipo de cambio de Compra " & strTipoCambio & "no está registrado para el día " & Convertddmmyyyy(strFechaConsulta), vbExclamation, "Registro de Tipo de Cambio"
        End If
        If adoConsulta("ValorTipoCambioVenta").Value = 0 Then
            MsgBox "El tipo de cambio de Venta " & strTipoCambio & "no está registrado para el día " & Convertddmmyyyy(strFechaConsulta), vbExclamation, "Registro de Tipo de Cambio"
        End If
    Else
        MsgBox "El tipo de cambio " & strTipoCambio & "no está registrado para el día " & Convertddmmyyyy(strFechaConsulta), vbExclamation, "Registro de Tipo de Cambio"
    End If
End Sub

Public Function VerificaOrdenPendienteFacturacion(strCodFondo As String) As Boolean
    Dim adoConsulta As ADODB.Recordset
    Dim result As Boolean
    
    result = True
    
    adoComm.CommandText = "SELECT COUNT(*) CantOrdenes FROM OrdenPago WHERE CodFondo = '" & strCodFondo & "' AND Estado = '01'"
   
    Set adoConsulta = adoComm.Execute
    
    If Not adoConsulta.EOF Then
        If adoConsulta("CantOrdenes").Value > 0 Then
            If MsgBox("Existen ordenes de pago pendientes de registrar en Compras." & vbNewLine & _
                        "¿Seguro de continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                result = result And False
            End If
        End If
    End If
    
    adoComm.CommandText = "SELECT COUNT(*) CantOrdenes FROM OrdenCobro WHERE CodFondo = '" & strCodFondo & "' AND Estado = '01'"
   
    Set adoConsulta = adoComm.Execute
    
    If Not adoConsulta.EOF And result Then
        If adoConsulta("CantOrdenes").Value > 0 Then
            If MsgBox("Existen ordenes de cobro pendientes de registrar en Ventas." & vbNewLine & _
                        "¿Seguro de continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                result = result And False
            End If
        End If
    End If
    
    VerificaOrdenPendienteFacturacion = result
    
End Function
