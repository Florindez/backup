VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmCierreDiario 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7080
   ClientLeft      =   1440
   ClientTop       =   1695
   ClientWidth     =   8805
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
   Icon            =   "frmCierreDiario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   8805
   Begin VB.Frame fraCierre 
      Caption         =   "Proceso de Cierre"
      Height          =   6915
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   8715
      Begin VB.Frame fraFechaCierre 
         Caption         =   "Fecha de Cierre"
         Height          =   1335
         Left            =   360
         TabIndex        =   21
         Top             =   960
         Width           =   7935
         Begin VB.ComboBox cboPeriodoContable 
            Enabled         =   0   'False
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
            Left            =   1260
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   360
            Width           =   3975
         End
         Begin MSComCtl2.DTPicker dtpFechaCierre 
            Height          =   345
            Left            =   1260
            TabIndex        =   22
            Top             =   760
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   178454529
            CurrentDate     =   38068
         End
         Begin MSComCtl2.DTPicker dtpFechaCierreHasta 
            Height          =   345
            Left            =   3930
            TabIndex        =   24
            Top             =   760
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   178454529
            CurrentDate     =   38068
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Periodo"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   10
            Left            =   330
            TabIndex        =   26
            Top             =   420
            Width           =   885
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Hasta"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   3000
            TabIndex        =   25
            Top             =   820
            Width           =   765
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Desde"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   330
            TabIndex        =   23
            Top             =   820
            Width           =   765
         End
      End
      Begin VB.CheckBox chkSimulacion 
         Caption         =   "Simular el Valor de Cuota"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         ToolTipText     =   "Marcar para proceso de simulación"
         Top             =   2400
         Width           =   2895
      End
      Begin VB.ComboBox cboFondo 
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
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   570
         Width           =   6825
      End
      Begin VB.CommandButton cmdCierre 
         Caption         =   "&Procesar"
         Height          =   735
         Left            =   7140
         Picture         =   "frmCierreDiario.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5010
         Width           =   1200
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   735
         Left            =   7140
         Picture         =   "frmCierreDiario.frx":09AA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5970
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker dtpFechaEntrega 
         Height          =   315
         Left            =   5640
         TabIndex        =   4
         Top             =   7950
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   178454529
         CurrentDate     =   38068
      End
      Begin TrueOleDBGrid60.TDBGrid tdgTipoCambioCierre 
         Bindings        =   "frmCierreDiario.frx":0F2C
         Height          =   1695
         Left            =   360
         OleObjectBlob   =   "frmCierreDiario.frx":0F4E
         TabIndex        =   19
         Top             =   5010
         Width           =   3015
      End
      Begin TrueOleDBGrid60.TDBGrid dbgFondoSeries 
         Height          =   2085
         Left            =   360
         OleObjectBlob   =   "frmCierreDiario.frx":3DD6
         TabIndex        =   28
         Top             =   2740
         Width           =   7995
      End
      Begin VB.Label lblDescrip 
         Alignment       =   2  'Center
         Caption         =   "Tipo de Cambio Oficial"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   8
         Left            =   3570
         TabIndex        =   20
         Top             =   8070
         Width           =   3015
      End
      Begin VB.Label lblRentabilidad 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0000"
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   1
         Left            =   7530
         TabIndex        =   18
         Top             =   9030
         Width           =   1635
      End
      Begin VB.Label lblValorDIR 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00000000"
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   1
         Left            =   5610
         TabIndex        =   17
         Top             =   9030
         Width           =   1800
      End
      Begin VB.Label lblValorAIR 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100.00000000"
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   1
         Left            =   3690
         TabIndex        =   16
         Top             =   9030
         Width           =   1800
      End
      Begin VB.Label lblDescrip 
         Caption         =   "dd/mm/yyyy"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   7
         Left            =   2460
         TabIndex        =   15
         Top             =   9030
         Width           =   1095
      End
      Begin VB.Label lblDescrip 
         Caption         =   "dd/mm/yyyy"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   6
         Left            =   2460
         TabIndex        =   14
         Top             =   8670
         Width           =   1095
      End
      Begin VB.Label lblValorAIR 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100.00000000"
         Height          =   285
         Index           =   0
         Left            =   3690
         TabIndex        =   12
         Top             =   8670
         Width           =   1800
      End
      Begin VB.Label lblValorDIR 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00000000"
         Height          =   285
         Index           =   0
         Left            =   5610
         TabIndex        =   11
         Top             =   8670
         Width           =   1800
      End
      Begin VB.Label lblRentabilidad 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0000"
         Height          =   285
         Index           =   0
         Left            =   7530
         TabIndex        =   10
         Top             =   8670
         Width           =   1605
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fondo"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   435
         TabIndex        =   9
         Top             =   615
         Width           =   615
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fecha de Entrega de Redenciones"
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   2
         Left            =   2505
         TabIndex        =   8
         Top             =   7995
         Width           =   3105
      End
      Begin VB.Label lblDescrip 
         Alignment       =   2  'Center
         Caption         =   "Valor A.I.R."
         ForeColor       =   &H00800000&
         Height          =   165
         Index           =   3
         Left            =   3735
         TabIndex        =   7
         Top             =   8430
         Width           =   1455
      End
      Begin VB.Label lblDescrip 
         Alignment       =   2  'Center
         Caption         =   "Valor D.I.R."
         ForeColor       =   &H00800000&
         Height          =   165
         Index           =   4
         Left            =   5655
         TabIndex        =   6
         Top             =   8430
         Width           =   1455
      End
      Begin VB.Label lblDescrip 
         Alignment       =   2  'Center
         Caption         =   "Rentabilidad"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   5
         Left            =   7575
         TabIndex        =   5
         Top             =   8430
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc adoTipoCambioCierre 
      Height          =   375
      Left            =   0
      Top             =   3810
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
End
Attribute VB_Name = "frmCierreDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()              As String
Dim arrPeriodo()            As String
Dim strCodFondo             As String, strFechaCierre               As String
Dim strFechaAnterior        As String, strFechaSiguiente            As String
Dim strFechaAnteAnterior    As String, strFechaSubSiguiente         As String
Dim strCodMoneda            As String, strCodModulo                 As String
Dim strSQL                  As String, datFechaCierre               As Date

Dim dblValNuevaCuota        As Double, dblValorCuotaNominal         As Double
Dim dblValNuevaCuotaReal    As Double

'*** Variables para los códigos de cuentas contables ***
Dim strCodCuentaValuacion As String, strCodCuentaResultados         As String
Dim intNumReproceso As Integer

'--------------------------------------- Variables C.Mensual

Dim strPeriodoContable          As String, strFechaCierreHasta              As String
Dim strMesContable              As String, strFechaCierreHastaSiguiente     As String
Dim strFechaCierreDesde         As String, strIndCobrar                     As String
Dim strTipoFondoFrecuencia     As String
Dim strFrecuenciaValorizacionMensual    As String
Dim adoFondoTipo                As ADODB.Recordset
Dim dblTCCierre                 As Double
Dim strTipoFondo                As String
'---------------------------------------


Private Sub GenOrdenPagoAdministradora()

    Dim adoRegistro                 As ADODB.Recordset
    Dim adoAuxiliar                 As ADODB.Recordset
    Dim strCodCuentaSuscripcion     As String, strCodCuentaRedencion        As String
    Dim strCodCuentaAdministracion  As String, strFechaPago                 As String
    Dim strNumCaja                  As String, strDescripAsiento            As String
    Dim strIndDebeHaber             As String, strCodFile                   As String
    Dim strCodAnalitica             As String, strCodCreditoFiscal          As String
    Dim curCuentaSuscripcion        As Currency, curCuentaRedencion         As Currency
    Dim curCuentaAdministracion     As Currency, curCuentaTotal             As Currency
    Dim intSecuencial               As Integer
    Dim strFechaInicio              As String
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
    
     '"WHERE (FechaCorte>='" & strFechaCierre & "' AND FechaCorte<'" & strFechaSiguiente & "') AND "
    
        .CommandText = "SELECT FechaInicio, FechaCorte,FechaPago,FPA.CodComision,CodFile,CodDetalleFile,CodAnalitica,CodCreditoFiscal " & _
            "FROM FondoPagoAdministradora FPA JOIN FondoComision FC " & _
            "ON(FC.CodComision=FPA.CodComision AND FC.CodFondo=FPA.CodFondo AND FC.CodAdministradora=FPA.CodAdministradora) " & _
            "WHERE FechaInicio = '" & strFechaCierre & "' AND " & _
            "FPA.CodFondo='" & strCodFondo & "' AND FPA.CodAdministradora='" & gstrCodAdministradora & "' AND FC.IndVigencia='X'"
        Set adoRegistro = .Execute
        
        Do While Not adoRegistro.EOF
            Set adoAuxiliar = New ADODB.Recordset
            
            '*** Obtener el número secuencial ***
            .CommandText = "SELECT MAX(NumGasto) NumSecuencial FROM FondoGasto " & _
                "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            Set adoAuxiliar = .Execute
            
            If Not adoAuxiliar.EOF Then
                If IsNull(adoAuxiliar("NumSecuencial")) Then
                    intSecuencial = 1
                Else
                    intSecuencial = CInt(adoAuxiliar("NumSecuencial")) + 1
                End If
            Else
                intSecuencial = 1
            End If
            adoAuxiliar.Close
            
            '*** Obtener descripción de comisión ***
            .CommandText = "SELECT DescripComision FROM ComisionEmpresa WHERE CodComision='" & adoRegistro("CodComision") & "'"
            Set adoAuxiliar = .Execute
            
            If Not adoAuxiliar.EOF Then
                strDescripAsiento = "Comisión " & Trim(adoAuxiliar("DescripComision")) & Space(1) & CStr(adoRegistro("Fechacorte"))
            End If
            adoAuxiliar.Close: Set adoAuxiliar = Nothing
                
            strFechaPago = Convertyyyymmdd(adoRegistro("FechaCorte"))
            strFechaInicio = Convertyyyymmdd(adoRegistro("FechaInicio"))
            strCodFile = adoRegistro("CodFile")
            strCodAnalitica = adoRegistro("CodAnalitica")
            strCodCreditoFiscal = adoRegistro("CodCreditoFiscal")
            
            '*** Obtener Cuentas ***
            strCodCuentaSuscripcion = ObtenerCuentaAdministracion("001", "R")
            strCodCuentaRedencion = ObtenerCuentaAdministracion("002", "R")
            strCodCuentaAdministracion = ObtenerCuentaAdministracion("003", "R")
            
            Call ObtenerCuentasInversion(strCodFile, adoRegistro("CodDetalleFile"), Trim(adoRegistro("CodMoneda")))
            
            '*** Obtener Saldos ***
            curCuentaAdministracion = Abs(ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaCierre, strFechaSiguiente, strCtaXPagar, strCodMoneda))
            
            '*** Guardar ***
            .CommandText = "{ call up_GNManFondoGasto('" & strCodFondo & "','" & _
                gstrCodAdministradora & "','" & Codigo_Gasto_Provision & "','" & strCtaComision & "'," & intSecuencial & ",'" & _
                strCodFile & "','" & strCodAnalitica & "','" & _
                Codigo_Frecuencia_Diaria & "','" & strCodMoneda & "','" & _
                Codigo_Pago_Vencimiento & "','" & strCodCreditoFiscal & "','" & strDescripAsiento & "','" & _
                strFechaCierre & "','" & Convertyyyymmdd(CVDate(Valor_Fecha)) & "','" & _
                strFechaInicio & "','" & strFechaPago & "'," & _
                gdblTipoCambio & "," & CDec(curCuentaAdministracion) & ",0,'','X','I') }"
            adoConn.Execute .CommandText
            

            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub
Private Sub GenFondoGastoAdmin()

    Dim adoRegistro                 As ADODB.Recordset
    Dim adoAuxiliar                 As ADODB.Recordset
    Dim strCodCuentaSuscripcion     As String, strCodCuentaRedencion        As String
    Dim strCodCuentaAdministracion  As String, strFechaPago                 As String
    Dim strNumCaja                  As String, strDescripAsiento            As String
    Dim strIndDebeHaber             As String, strCodFile                   As String
    Dim strCodAnalitica             As String, strCodCreditoFiscal          As String
    Dim curCuentaSuscripcion        As Currency, curCuentaRedencion         As Currency
    Dim curCuentaAdministracion     As Currency, curCuentaTotal             As Currency
    Dim intSecuencial               As Integer
    Dim strFechaInicio              As String
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
    
     '"WHERE (FechaCorte>='" & strFechaCierre & "' AND FechaCorte<'" & strFechaSiguiente & "') AND "
    
        .CommandText = "SELECT FechaInicio, FechaCorte,FechaPago,FPA.CodComision,CodFile,CodDetalleFile,FPA.CodAnalitica,CodCreditoFiscal " & _
            "FROM FondoPagoAdministradora FPA JOIN FondoComision FC " & _
            "ON(FC.CodAnalitica=FPA.CodAnalitica AND FC.CodComision=FPA.CodComision AND FC.CodFondo=FPA.CodFondo AND FC.CodAdministradora=FPA.CodAdministradora) " & _
            "WHERE FechaInicio = '" & strFechaCierre & "' AND " & _
            "FPA.CodFondo='" & strCodFondo & "' AND FPA.CodAdministradora='" & gstrCodAdministradora & "' AND IndVigente='X'"
        Set adoRegistro = .Execute
        
        Do While Not adoRegistro.EOF
            Set adoAuxiliar = New ADODB.Recordset
            
            '*** Obtener el número secuencial ***
            .CommandText = "SELECT MAX(NumGasto) NumSecuencial FROM FondoGasto " & _
                "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            Set adoAuxiliar = .Execute
            
            If Not adoAuxiliar.EOF Then
                If IsNull(adoAuxiliar("NumSecuencial")) Then
                    intSecuencial = 1
                Else
                    intSecuencial = CInt(adoAuxiliar("NumSecuencial")) + 1
                End If
            Else
                intSecuencial = 1
            End If
            adoAuxiliar.Close
            
            '*** Obtener descripción de comisión ***
            .CommandText = "SELECT DescripComision FROM ComisionEmpresa WHERE CodComision='" & adoRegistro("CodComision") & "'"
            Set adoAuxiliar = .Execute
            
            If Not adoAuxiliar.EOF Then
                strDescripAsiento = "Comisión " & Trim(adoAuxiliar("DescripComision")) & Space(1) & CStr(adoRegistro("Fechacorte"))
            End If
            adoAuxiliar.Close: Set adoAuxiliar = Nothing
                
            strFechaPago = Convertyyyymmdd(adoRegistro("FechaCorte"))
            strFechaInicio = Convertyyyymmdd(adoRegistro("FechaInicio"))
            strCodFile = adoRegistro("CodFile")
            strCodAnalitica = adoRegistro("CodAnalitica")
            strCodCreditoFiscal = adoRegistro("CodCreditoFiscal")
            
            '*** Obtener Cuentas ***
            strCodCuentaSuscripcion = ObtenerCuentaAdministracion("001", "R")
            strCodCuentaRedencion = ObtenerCuentaAdministracion("002", "R")
            strCodCuentaAdministracion = ObtenerCuentaAdministracion("003", "R")
            
            Call ObtenerCuentasInversion(strCodFile, adoRegistro("CodDetalleFile"), Trim(adoRegistro("CodMoneda")))
            
            '*** Obtener Saldos ***
            curCuentaAdministracion = Abs(ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaCierre, strFechaSiguiente, strCtaXPagar, strCodMoneda))
            
            '*** Guardar ***
            .CommandText = "{ call up_GNManFondoGasto('" & strCodFondo & "','" & _
                gstrCodAdministradora & "','" & Codigo_Gasto_Provision & "','" & strCtaComision & "'," & intSecuencial & ",'" & _
                strCodFile & "','" & strCodAnalitica & "','" & _
                Codigo_Frecuencia_Diaria & "','" & strCodMoneda & "','" & _
                Codigo_Pago_Vencimiento & "','" & strCodCreditoFiscal & "','" & strDescripAsiento & "','" & _
                strFechaCierre & "','" & Convertyyyymmdd(CVDate(Valor_Fecha)) & "','" & _
                strFechaInicio & "','" & strFechaPago & "'," & _
                gdblTipoCambio & "," & CDec(curCuentaAdministracion) & ",0,'','X','','I') }"
            adoConn.Execute .CommandText
            

            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub
Private Sub GenFondoGastoAdminTmp()

    Dim adoRegistro                 As ADODB.Recordset
    Dim adoAuxiliar                 As ADODB.Recordset
    Dim strCodCuentaSuscripcion     As String, strCodCuentaRedencion        As String
    Dim strCodCuentaAdministracion  As String, strFechaPago                 As String
    Dim strNumCaja                  As String, strDescripAsiento            As String
    Dim strIndDebeHaber             As String, strCodFile                   As String
    Dim strCodAnalitica             As String, strCodCreditoFiscal          As String
    Dim curCuentaSuscripcion        As Currency, curCuentaRedencion         As Currency
    Dim curCuentaAdministracion     As Currency, curCuentaTotal             As Currency
    Dim intSecuencial               As Integer
    Dim strFechaInicio              As String
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
    
     '"WHERE (FechaCorte>='" & strFechaCierre & "' AND FechaCorte<'" & strFechaSiguiente & "') AND "
    
        .CommandText = "SELECT FechaInicio, FechaCorte,FechaPago,FPA.CodComision,CodFile,CodDetalleFile,FPA.CodAnalitica,CodCreditoFiscal " & _
            "FROM FondoPagoAdministradora FPA JOIN FondoComision FC " & _
            "ON(FC.CodAnalitica=FPA.CodAnalitica AND FC.CodComision=FPA.CodComision AND FC.CodFondo=FPA.CodFondo AND FC.CodAdministradora=FPA.CodAdministradora) " & _
            "WHERE FechaInicio = '" & strFechaCierre & "' AND " & _
            "FPA.CodFondo='" & strCodFondo & "' AND FPA.CodAdministradora='" & gstrCodAdministradora & "' AND IndVigente='X'"
        Set adoRegistro = .Execute
        
        Do While Not adoRegistro.EOF
            Set adoAuxiliar = New ADODB.Recordset
            
            '*** Obtener el número secuencial ***
            .CommandText = "SELECT MAX(NumGasto) NumSecuencial FROM FondoGastoTmp " & _
                "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            Set adoAuxiliar = .Execute
            
            If Not adoAuxiliar.EOF Then
                If IsNull(adoAuxiliar("NumSecuencial")) Then
                    intSecuencial = 1
                Else
                    intSecuencial = CInt(adoAuxiliar("NumSecuencial")) + 1
                End If
            Else
                intSecuencial = 1
            End If
            adoAuxiliar.Close
            
            '*** Obtener descripción de comisión ***
            .CommandText = "SELECT DescripComision FROM ComisionEmpresa WHERE CodComision='" & adoRegistro("CodComision") & "'"
            Set adoAuxiliar = .Execute
            
            If Not adoAuxiliar.EOF Then
                strDescripAsiento = "Comisión " & Trim(adoAuxiliar("DescripComision")) & Space(1) & CStr(adoRegistro("Fechacorte"))
            End If
            adoAuxiliar.Close: Set adoAuxiliar = Nothing
                
            strFechaPago = Convertyyyymmdd(adoRegistro("FechaCorte"))
            strFechaInicio = Convertyyyymmdd(adoRegistro("FechaInicio"))
            strCodFile = adoRegistro("CodFile")
            strCodAnalitica = adoRegistro("CodAnalitica")
            strCodCreditoFiscal = adoRegistro("CodCreditoFiscal")
            
            '*** Obtener Cuentas ***
            strCodCuentaSuscripcion = ObtenerCuentaAdministracion("001", "R")
            strCodCuentaRedencion = ObtenerCuentaAdministracion("002", "R")
            strCodCuentaAdministracion = ObtenerCuentaAdministracion("003", "R")
            
            Call ObtenerCuentasInversion(strCodFile, adoRegistro("CodDetalleFile"), Trim(adoRegistro("CodMoneda")))
            
            '*** Obtener Saldos ***
            curCuentaAdministracion = Abs(ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaCierre, strFechaSiguiente, strCtaXPagar, strCodMoneda))
            
            '*** Guardar ***
            .CommandText = "{ call up_GNManFondoGastoTmp('" & strCodFondo & "','" & _
                gstrCodAdministradora & "','" & Codigo_Gasto_Provision & "','" & strCtaComision & "'," & intSecuencial & ",'" & _
                strCodFile & "','" & strCodAnalitica & "','" & _
                Codigo_Frecuencia_Diaria & "','" & strCodMoneda & "','" & _
                Codigo_Pago_Vencimiento & "','" & strCodCreditoFiscal & "','" & strDescripAsiento & "','" & _
                strFechaCierre & "','" & Convertyyyymmdd(CVDate(Valor_Fecha)) & "','" & _
                strFechaInicio & "','" & strFechaPago & "'," & _
                gdblTipoCambio & "," & CDec(curCuentaAdministracion) & ",0,'','X','','I') }"
            adoConn.Execute .CommandText
            

            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub
Private Sub GenOrdenPagoCuotaSuscripcion(strpTipoValuacion As String)

    Dim adoRegistro             As ADODB.Recordset, adoTemporal As ADODB.Recordset
    Dim strNumCaja              As String, strNumOperacion      As String
    Dim strNumOrdenCobroPago    As String, strCodFile           As String
    Dim strCodAnalitica         As String, strDescripAsiento    As String
    Dim strCodCuenta            As String, strClasePersona      As String
    Dim strTipoPago             As String
    Dim strIndDebeHaber         As String, strNumFolio          As String
    Dim dblCantCuotas           As Double, dblValorCuota        As Double
    Dim dblPorcenParcial        As Double, dblCantCuotasReal    As Double
    Dim curMontoTotal           As Currency, curMontoComision   As Currency
    Dim curMontoIgv             As Currency, curSubTotal        As Currency
    Dim curMontoAporte          As Currency, curMontoSobreLaPar As Currency
    Dim curMontoBajoLaPar       As Currency, curMontoIgvOrden   As Currency
    Dim curMontoComisionOrden   As Currency, curTotalOperacion  As Currency
    Dim curMontoTotalOrden      As Currency, curSubTotalReal    As Currency
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        .CommandText = "SELECT CodParticipe,NumSolicitud,NumSecuencial,MontoLiquidacion FROM ParticipePagoSuscripcion " & _
            "WHERE (FechaLiquidacion>='" & strFechaSiguiente & "' AND FechaLiquidacion<'" & strFechaSubSiguiente & "') AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND FechaPago='01/01/1900'"
        Set adoRegistro = .Execute
        
        Set adoTemporal = New ADODB.Recordset
        
        Do Until adoRegistro.EOF
            '*** Obtener secuencial ***
            strNumCaja = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOrdenCaja)
                                            
            '*** Obtener número de orden de pago anterior ***
            .CommandText = "SELECT NumOrdenCobroPago FROM ParticipePagoSuscripcion " & _
                "WHERE FechaLiquidacion<'" & strFechaSiguiente & "' AND CodParticipe='" & adoRegistro("CodParticipe") & "' AND " & _
                "NumSolicitud='" & adoRegistro("NumSolicitud") & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            Set adoTemporal = .Execute
            
            If Not adoTemporal.EOF Then
                strNumOrdenCobroPago = adoTemporal("NumOrdenCobroPago")
            End If
            adoTemporal.Close
            
            '*** Obtener demás información para generar la orden ***
            .CommandText = "SELECT NumOperacion,CodBanco,CodCuenta,CodFile,CodAnalitica,CodMoneda,DescripOrden,NumFolio,TipoPago " & _
                "FROM MovimientoFondo WHERE NumOrdenCobroPago='" & strNumOrdenCobroPago & "' AND " & _
                "CodParticipe='" & adoRegistro("CodParticipe") & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            Set adoTemporal = .Execute
            
            If Not adoTemporal.EOF Then
                strNumOperacion = adoTemporal("NumOperacion")
                strCtaXCobrar = adoTemporal("CodCuenta")
                strCodFile = adoTemporal("CodFile")
                strCodAnalitica = adoTemporal("CodAnalitica")
                strCodMoneda = adoTemporal("CodMoneda")
                strDescripAsiento = adoTemporal("DescripOrden")
                strNumFolio = adoTemporal("NumFolio")
                strTipoPago = adoTemporal("TipoPago")
            End If
            adoTemporal.Close
            
            '*** Obtener datos de la operación ***
            .CommandText = "SELECT ClasePersona,CantCuotas,MontoTotal, MontoComision, MontoIgv, ValorCuota " & _
                "FROM ParticipeOperacion WHERE NumOperacion='" & strNumOperacion & "' AND " & _
                "CodParticipe='" & adoRegistro("CodParticipe") & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "EstadoOperacion='" & Estado_Activo & "'"
            Set adoTemporal = .Execute
            
            If Not adoTemporal.EOF Then
                strClasePersona = adoTemporal("ClasePersona")
                dblCantCuotas = adoTemporal("CantCuotas")
                curMontoTotal = adoTemporal("MontoTotal")
                curMontoComision = adoTemporal("MontoComision")
                curMontoIgv = adoTemporal("MontoIgv")
                dblValorCuota = adoTemporal("ValorCuota")
            End If
            adoTemporal.Close
                        
            '*** Orden de Cobro/Pago ***
            .CommandText = "{ call up_ACAdicMovimientoFondo('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strNumCaja & "','" & strFechaSiguiente & "','" & Trim(frmMainMdi.Tag) & "','" & strNumOperacion & "','" & strFechaSiguiente & "','" & _
                "','','E','" & strCtaXCobrar & "'," & CDec(adoRegistro("MontoLiquidacion")) & ",'" & _
                strCodFile & "','" & strCodAnalitica & "','" & strCodMoneda & "','" & _
                strDescripAsiento & "','" & Codigo_Caja_Suscripcion & "','','" & Estado_Caja_NoConfirmado & "','" & _
                adoRegistro("CodParticipe") & "','" & strNumFolio & "','" & strTipoPago & "','','','','','','','" & gstrLogin & "') }"
            adoConn.Execute .CommandText
                
            '*** Cálculos ***
            dblPorcenParcial = CDbl(adoRegistro("MontoLiquidacion")) / curMontoTotal
            curSubTotal = (dblCantCuotas * dblPorcenParcial * dblValorCuota)
            curMontoAporte = (dblCantCuotas * dblPorcenParcial * dblValorCuotaNominal)
            
            If dblValorCuota - dblValorCuotaNominal > 0 Then
                curMontoSobreLaPar = (dblValorCuota - dblValorCuotaNominal) * Abs(dblCantCuotas * dblPorcenParcial)
            Else
                curMontoSobreLaPar = 0
            End If
            
            If dblValorCuota - dblValorCuotaNominal < 0 Then
                curMontoBajoLaPar = (dblValorCuota - dblValorCuotaNominal) * Abs(dblCantCuotas * dblPorcenParcial)
            Else
                curMontoBajoLaPar = 0
            End If
                                    
            curMontoComisionOrden = curMontoComision * dblPorcenParcial
            curMontoIgvOrden = curMontoIgv * dblPorcenParcial
            curMontoTotalOrden = CDbl(adoRegistro("MontoLiquidacion"))
            curTotalOperacion = Round(curMontoTotalOrden, 2)
            
            curSubTotalReal = curMontoTotalOrden - curMontoComisionOrden - curMontoIgvOrden
            If strpTipoValuacion = Codigo_Asignacion_TMenos1 Then
                dblCantCuotasReal = curSubTotalReal / dblValNuevaCuotaReal
            Else
                dblCantCuotasReal = 0
            End If
            
            '*** Movimientos del Detalle de la Orden ***
            strCodCuenta = ObtenerCuentaAdministracion("018", "R")
            strDescripAsiento = "Operac.de Suscrip"
            strIndDebeHaber = "D"
            strCodFile = "000"
            strCodAnalitica = "00000000"

            '*** Orden de Cobro/Pago Detalle ***
            .CommandText = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strNumCaja & "','" & strFechaSiguiente & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripAsiento & "','" & _
                strIndDebeHaber & "','" & Trim(strCodCuenta) & "'," & CDec(adoRegistro("MontoLiquidacion")) & ",'" & _
                strCodFile & "','" & strCodAnalitica & "','" & strCodMoneda & "','') }"
            adoConn.Execute .CommandText
            
            If strClasePersona = Codigo_Persona_Natural Then
                strCodCuenta = ObtenerCuentaAdministracion("006", "C")
            Else
                strCodCuenta = ObtenerCuentaAdministracion("009", "C")
            End If
            strDescripAsiento = "Capital Fijo"
            strIndDebeHaber = "H"
            strCodFile = "000"
            strCodAnalitica = "00000000"

            '*** Orden de Cobro/Pago Detalle ***
            .CommandText = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strNumCaja & "','" & strFechaSiguiente & "',2,'" & Trim(frmMainMdi.Tag) & "','" & strDescripAsiento & "','" & _
                strIndDebeHaber & "','" & Trim(strCodCuenta) & "'," & CDec(curMontoAporte) * -1 & ",'" & _
                strCodFile & "','" & strCodAnalitica & "','" & strCodMoneda & "','') }"
            adoConn.Execute .CommandText
            
            If strClasePersona = Codigo_Persona_Natural Then
                If curMontoBajoLaPar > 0 Then
                    strCodCuenta = ObtenerCuentaAdministracion("010", "C")
                Else
                    strCodCuenta = ObtenerCuentaAdministracion("014", "C")
                End If
            Else
                If curMontoBajoLaPar > 0 Then
                    strCodCuenta = ObtenerCuentaAdministracion("013", "C")
                Else
                    strCodCuenta = ObtenerCuentaAdministracion("017", "C")
                End If
            End If
            strDescripAsiento = "Capital Adicional"
            strIndDebeHaber = "H"
            strCodFile = "000"
            strCodAnalitica = "00000000"

            '*** Orden de Cobro/Pago Detalle ***
            .CommandText = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strNumCaja & "','" & strFechaSiguiente & "',3,'" & Trim(frmMainMdi.Tag) & "','" & strDescripAsiento & "','" & _
                strIndDebeHaber & "','" & Trim(strCodCuenta) & "'," & CDec(curMontoBajoLaPar) * -1 & ",'" & _
                strCodFile & "','" & strCodAnalitica & "','" & strCodMoneda & "','') }"
            adoConn.Execute .CommandText
            
            strCodCuenta = ObtenerCuentaAdministracion("001", "R")
            strDescripAsiento = "Comisiones"
            strIndDebeHaber = "D"
            strCodFile = "000"
            strCodAnalitica = "00000000"

            '*** Orden de Cobro/Pago Detalle ***
            .CommandText = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strNumCaja & "','" & strFechaSiguiente & "',4,'" & Trim(frmMainMdi.Tag) & "','" & strDescripAsiento & "','" & _
                strIndDebeHaber & "','" & Trim(strCodCuenta) & "'," & CDec(curMontoComisionOrden) & ",'" & _
                strCodFile & "','" & strCodAnalitica & "','" & strCodMoneda & "','') }"
            adoConn.Execute .CommandText

            strCodCuenta = ObtenerCuentaAdministracion("001", "R")
            strDescripAsiento = "Tributos x Pagar"
            strIndDebeHaber = "D"
            strCodFile = "000"
            strCodAnalitica = "00000000"

            '*** Orden de Cobro/Pago Detalle ***
            .CommandText = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strNumCaja & "','" & strFechaSiguiente & "',5,'" & Trim(frmMainMdi.Tag) & "','" & strDescripAsiento & "','" & _
                strIndDebeHaber & "','" & Trim(strCodCuenta) & "'," & CDec(curMontoIgvOrden) & ",'" & _
                strCodFile & "','" & strCodAnalitica & "','" & strCodMoneda & "','') }"
            adoConn.Execute .CommandText
            
'            strCodCuenta = ObtenerCuentaAdministracion("030", "R")
'            strDescripAsiento = "Suscrip.X Cobrar"
            strCodCuenta = ObtenerCuentaAdministracion("018", "R")
            strDescripAsiento = "Operac.de Suscrip"
            strIndDebeHaber = "H"
            strCodFile = "000"
            strCodAnalitica = "00000000"
                        
            '*** Orden de Cobro/Pago Detalle ***
            .CommandText = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strNumCaja & "','" & strFechaSiguiente & "',6,'" & Trim(frmMainMdi.Tag) & "','" & strDescripAsiento & "','" & _
                strIndDebeHaber & "','" & Trim(strCodCuenta) & "'," & CDec(curMontoAporte) * -1 & ",'" & _
                strCodFile & "','" & strCodAnalitica & "','" & strCodMoneda & "','') }"
            adoConn.Execute .CommandText
                            
            '*** Actualizar Secuencial ***
            .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                Valor_NumOrdenCaja & "','" & strNumCaja & "') }"
            adoConn.Execute .CommandText
            
            '*** Actualizar Número de Orden y Cuotas Reales ***
            .CommandText = "UPDATE ParticipePagoSuscripcion " & _
                "SET NumOrdenCobroPago='" & strNumCaja & "',CantCuotasPagadasReal=" & dblCantCuotasReal & " " & _
                "WHERE NumSecuencial=" & adoRegistro("NumSecuencial") & " AND NumSolicitud='" & adoRegistro("NumSolicitud") & "' AND " & _
                "CodParticipe='" & adoRegistro("CodParticipe") & "' AND CodFondo='" & strCodFondo & "' AND " & _
                "CodAdministradora='" & gstrCodAdministradora & "'"
            adoConn.Execute .CommandText
                    
            adoRegistro.MoveNext
        Loop
        Set adoTemporal = Nothing
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub

Private Sub ValorizacionRentaFijaCortoplazo(strTipoCierre As String)

    Dim adoRegistro                 As ADODB.Recordset, adoConsulta         As ADODB.Recordset
    Dim dblPrecioCierre             As Double, dblPrecioPromedio            As Double
    Dim dblTirCierre                As Double, dblFactorDiario              As Double
    Dim dblTasaInteres              As Double, dblFactorDiarioCupon         As Double
    Dim curSaldoGPCapital           As Currency, curSaldoFluctuacionMercado As Currency
    Dim curSaldoInversion           As Currency, curSaldoInteresCorrido     As Currency
    Dim curSaldoFluctuacion         As Currency, curValorAnterior           As Currency
    Dim curValorActual              As Currency, curMontoRenta              As Currency
    Dim curMontoContable            As Currency, curMontoMovimientoMN       As Currency
    Dim curMontoMovimientoME        As Currency, curSaldoValorizar          As Currency
    Dim curMontoProvisionCapital    As Currency, curMontoFluctuacionMercado As Currency
    Dim intCantRegistros            As Integer, intContador                 As Integer
    Dim intRegistro                 As Integer, intBaseCalculo              As Integer
    Dim intDiasPlazo                As Integer, intDiasDeRenta              As Integer
    Dim strNumAsiento               As String, strDescripAsiento            As String
    Dim strDescripMovimiento        As String, strIndDebeHaber              As String
    Dim strCodCuenta                As String, strFiles                     As String
    Dim strCodFile                  As String, strModalidadInteres          As String
    Dim strCodTasa                  As String, strIndCuponCero              As String
    Dim strCodDetalleFile           As String, strNemonico                  As String
    Dim strCodIndiceInicial         As String, strCodIndiceFinal            As String
    Dim strFechaGrabar              As String, strBaseAnual                 As String
    Dim dblTipoCambioCierre         As Double
    
    Dim dblValorTipoCambio          As Double, strTipoDocumento As String
    Dim strNumDocumento             As String
    
    Dim strTipoPersonaContraparte   As String, strCodPersonaContraparte As String
    
    Dim strIndContracuenta          As String, strCodContracuenta As String
    Dim strCodFileContracuenta      As String, strCodAnaliticaContracuenta As String
    Dim strIndUltimoMovimiento      As String
    
    
    '*** Rentabilidad de Valores de Renta Fija Corto Plazo ***
    frmMainMdi.stbMdi.Panels(3).Text = "Calculando rentabilidad de Valores de Renta Fija Corto Plazo..."
       
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT IK.MontoSaldo,IK.CodTitulo,II.CodMoneda,IK.CodAnalitica,IK.CodFile,SaldoFinal,CodDetalleFile,CodSubDetalleFile,CodTipoVac,CuponCalculo," & _
            "ValorNominal,CodTipoTasa,BaseAnual,CodTasa,TasaInteres,DiasPlazo,IndCuponCero,FechaEmision,TirPromedio,CodTipoAjuste,Nemotecnico,II.CodEmisor " & _
            "FROM InversionKardex IK JOIN InstrumentoInversion II " & _
            "ON(II.CodTitulo=IK.CodTitulo) " & _
            "WHERE IK.CodFile IN ('006','010','012','014','015') AND FechaOperacion < '" & strFechaSiguiente & "' AND SaldoFinal > 0 AND " & _
            "IK.CodAdministradora='" & gstrCodAdministradora & "' AND IK.CodFondo='" & strCodFondo & "' AND IK.NumKardex = dbo.uf_IVObtenerUltimoMovimientoKardexValor(IK.CodFondo,IK.CodAdministradora,IK.CodTitulo,'" & strFechaCierre & "')"
        Set adoRegistro = .Execute
    
        Do Until adoRegistro.EOF
            strCodFile = Trim(adoRegistro("CodFile"))
            strCodDetalleFile = Trim(adoRegistro("CodDetalleFile"))
            strModalidadInteres = Trim(adoRegistro("CodDetalleFile"))
            strCodTasa = Trim(adoRegistro("CodTipoTasa"))
            strBaseAnual = Trim(adoRegistro("BaseAnual"))
            dblTasaInteres = CDbl(adoRegistro("TasaInteres"))
            intDiasPlazo = CInt(adoRegistro("DiasPlazo"))
            strIndCuponCero = Trim(adoRegistro("IndCuponCero"))
            strNemonico = Trim(adoRegistro("Nemotecnico"))
            curSaldoValorizar = CCur(adoRegistro("SaldoFinal")) '("MontoSaldo")
            intDiasDeRenta = DateDiff("d", CVDate(adoRegistro("FechaEmision")), gdatFechaActual) '+ 1
            
            If strBaseAnual = Codigo_Base_30_360 Or strBaseAnual = Codigo_Base_30_365 Then intDiasDeRenta = Dias360(CVDate(adoRegistro("FechaEmision")), gdatFechaActual, True) '+ 1
            
            Set adoConsulta = New ADODB.Recordset
                        
            '*** Verificar Dinamica Contable ***
'            .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
'                "WHERE TipoOperacion='" & Codigo_Dinamica_Provision & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
'                strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
'                IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"
'            Set adoConsulta = .Execute
'
'            If Not adoConsulta.EOF Then
'                If CInt(adoConsulta("NumRegistros")) > 0 Then
'                    intCantRegistros = CInt(adoConsulta("NumRegistros"))
'                Else
'                    MsgBox "NO EXISTE Dinámica Contable para la valorización", vbCritical
'                    adoConsulta.Close: Set adoConsulta = Nothing
'                    Exit Sub
'                End If
'            End If
'            adoConsulta.Close
                        
            '*** Obtener Ultimo Precio de Cierre registrado ***
            .CommandText = "{ call up_IVSelDatoInstrumentoInversion(2,'" & _
                Trim(adoRegistro("CodTitulo")) & "','" & strFechaCierre & "') }"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                dblPrecioCierre = CDbl(adoConsulta("PrecioCierre"))
                dblTirCierre = CDbl(adoConsulta("TirCierre"))
                dblPrecioPromedio = CDbl(adoConsulta("PrecioPromedio"))
            End If
            adoConsulta.Close
            
            '*** Obtener el factor diario del cupón ***
            .CommandText = "SELECT FactorDiario FROM InstrumentoInversionCalendario " & _
                "WHERE CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                dblFactorDiarioCupon = CDbl(adoConsulta("FactorDiario"))
            End If
            adoConsulta.Close
            
            '*** Obtener las cuentas de inversión ***
            Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, Trim(adoRegistro("CodMoneda")))
            
            '*** Obtener tipo de cambio ***
            'dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
            dblTipoCambioCierre = ObtenerValorTipoCambio(adoRegistro("CodMoneda"), Codigo_Moneda_Local, strFechaCierre, strFechaCierre, gstrCodClaseTipoCambioOperacionFondo, Codigo_Valor_TipoCambioCompra)
            
            '*** Obtener Saldo de Inversión ***
            If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaInversion & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoInversion = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Interés Corrido ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaInteresCorrido & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoInteresCorrido = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Provisión ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvInteres & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoFluctuacion = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Provisión G/P Capital ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvFlucK & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoGPCapital = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Fluctuación Mercado ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvFlucMercado & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoFluctuacionMercado = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
                        
            curValorAnterior = curSaldoInteresCorrido + curSaldoFluctuacion
            
            If adoRegistro("BaseAnual") = Codigo_Base_Actual_365 Or adoRegistro("BaseAnual") = Codigo_Base_30_365 Or adoRegistro("BaseAnual") = Codigo_Base_Actual_Actual Then
                intBaseCalculo = 365
            Else
                intBaseCalculo = 360
            End If
            
            If Trim(adoRegistro("CodSubDetalleFile")) <> Valor_Caracter Then strModalidadInteres = Trim(adoRegistro("CodSubDetalleFile"))
            
            'Con tasa de interes
            If ((adoRegistro("CodFile") = "006" Or adoRegistro("CodFile") = "012") And adoRegistro("CodDetalleFile") = "001") Or adoRegistro("CodFile") = "010" Then

                If strIndCuponCero = Valor_Indicador Then
                    dblFactorDiario = dblFactorDiarioCupon
                Else
                    If strCodTasa = Codigo_Tipo_Tasa_Efectiva Then
                        If strBaseAnual = Codigo_Base_30_360 Or strBaseAnual = Codigo_Base_30_365 Then
                            dblFactorDiario = ((1 + dblTasaInteres * 0.01) ^ (intDiasDeRenta / intBaseCalculo)) - 1
                        Else
                            dblFactorDiario = ((1 + dblTasaInteres * 0.01) ^ (intDiasDeRenta / intBaseCalculo)) - 1
                        End If
'                        dblFactorDiario = ((1 + CDbl(((1 + (dblTasaInteres / 100)) ^ (intDiasPlazo / intBaseCalculo)) - 1)) ^ (1 / intDiasPlazo)) - 1
                    Else
                        If strBaseAnual = Codigo_Base_30_360 Or strBaseAnual = Codigo_Base_30_365 Then
                            dblFactorDiario = (((dblTasaInteres * 0.01) / intBaseCalculo) * intDiasDeRenta)
                        Else
                            dblFactorDiario = (((dblTasaInteres * 0.01) / intBaseCalculo) * intDiasDeRenta)
                        End If
'                        dblFactorDiario = (CDbl(((1 + (dblTasaInteres / 100)) / intBaseCalculo)))
                    End If
                End If
                
                curValorActual = Round(curSaldoValorizar * dblFactorDiario, 2)
                
            Else  'al descuento
            
                dblFactorDiario = dblFactorDiarioCupon
             
                '*** Obtener el factor diario del cupón ***
                .CommandText = "SELECT FactorDiario, ValorInteres + ValorAmortizacion AS SaldoValorizar FROM InstrumentoInversionCalendario " & _
                    "WHERE CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curSaldoValorizar = CDbl(adoConsulta("SaldoValorizar")) 'MFL2
                End If
                adoConsulta.Close
                
                curValorAnterior = curSaldoValorizar / ((1 + dblFactorDiario) ^ (intDiasPlazo - intDiasDeRenta + 1))
                
                curValorActual = curSaldoValorizar / ((1 + dblFactorDiario) ^ (intDiasPlazo - intDiasDeRenta))
            End If
            
                
                        
            curMontoRenta = Round(curValorActual - curValorAnterior, 2)
            
'inicio comentarios acr: 23/03/2013
'            '*** Cálculo Provisión G/P Capital ***
'            'If strOrigen = "L" Then
'                '*** VAN AL DIA ANTERIOR AL CIERRE ***
'                curValorAnterior = curSaldoInversion + curSaldoInteresCorrido + curSaldoFluctuacion + curSaldoGPCapital
'
'                '*** CALCULO DEL VAN A LA FECHA DE CIERRE ***
'                If CDbl(adoRegistro("TirPromedio")) <> 0 Then
'                    If strModalidadInteres = Codigo_Interes_Descuento Then curSaldoValorizar = curSaldoInversion
'
'                    Dim datFechaGP  As Date, datFechaSiguienteGP    As Date
'                    Dim dblValorTir As Double
'
'                    datFechaGP = Convertddmmyyyy(strFechaCierre)
'                    datFechaSiguienteGP = Convertddmmyyyy(strFechaSiguiente)
'
'                    dblValorTir = CDbl(adoRegistro("TirPromedio"))
'
'                    curValorActual = VNANoPer(adoRegistro("CodTitulo"), datFechaSiguienteGP, datFechaSiguienteGP, curSaldoValorizar, curSaldoValorizar, dblValorTir, adoRegistro("CodTipoAjuste"), Valor_Caracter, Valor_Caracter)
'
'                    '*** CALCULO DEL MONTO DE GANANCIA/PERDIDA DE curCapital ***
'                    curMontoProvisionCapital = Round(curValorActual - curValorAnterior - curMontoRenta, 2)
'                Else
'                    curMontoProvisionCapital = 0
'                End If
'
'            'End If
'
'            '*** Cálculo Fluctuación Mercado ***
'            If dblTirCierre > 0 Then
'                curValorAnterior = curSaldoInversion + curSaldoInteresCorrido + curSaldoFluctuacion + curSaldoGPCapital + curSaldoFluctuacionMercado
'
'                'curValorActual= CDbl(VNANoPer(Trim$(adoresult!COD_FILE), Trim$(adoresult!COD_ANAL), vntFeccie, vntFeccieMas1Dia, CDbl(curCapital), SldAmort, dblTirHoy, strTipVac))
'                curValorActual = VNANoPer(adoRegistro("CodTitulo"), datFechaSiguienteGP, datFechaSiguienteGP, curSaldoValorizar, curSaldoValorizar, dblTirCierre, adoRegistro("CodTipoAjuste"), Valor_Caracter, Valor_Caracter)
'
'                curMontoFluctuacionMercado = Round(curValorActual - curValorAnterior - curMontoRenta - curMontoProvisionCapital, 2)
'            Else
'                curMontoFluctuacionMercado = 0
'            End If
'fin comentarios acr: 23/03/2013

            '*** Contabilización ***
            If curMontoRenta <> 0 Then 'Or curMontoProvisionCapital <> 0 Or curMontoFluctuacionMercado <> 0 Then
                'strDescripAsiento = "Valorización" & Space(1) & "(" & Trim(adoRegistro("CodFile")) & "-" & Trim(adoRegistro("CodAnalitica")) & ")"
                strDescripAsiento = "Valorización" & Space(1) & strNemonico
                strDescripMovimiento = "Pérdida"
                If curMontoRenta > 0 Then strDescripMovimiento = "Ganancia"
                                                
                .CommandType = adCmdStoredProc
                '*** Obtener el número del parámetro **
                .CommandText = "up_ACObtenerUltNumero"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "up_ACObtenerUltNumeroTmp"  '*** Simulación ***
                
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
                
'                .CommandText = "BEGIN TRAN ProcAsiento"
'                adoConn.Execute .CommandText
                
                On Error GoTo Ctrl_Error
                
                '*** Contabilizar ***
                strFechaGrabar = strFechaCierre & Space(1) & Format(Time, "hh:mm")
                
                '*** Cabecera ***
                .CommandText = "{ call up_ACAdicAsientoContable('"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableTmp('"  '*** Simulación ***
                
                .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
                    strFechaGrabar & "','" & _
                    gstrPeriodoActual & "','" & gstrMesActual & "','" & _
                    "','" & _
                    strDescripAsiento & "','" & Trim(adoRegistro("CodMoneda")) & "','" & _
                    Codigo_Moneda_Local & "','" & _
                    "','" & _
                    "'," & _
                    CDec(curMontoRenta) & ",'" & Estado_Activo & "'," & _
                    intCantRegistros & ",'" & _
                    strFechaCierre & Space(1) & Format(Time, "hh:ss") & "','" & _
                    strCodModulo & "','" & _
                    "'," & _
                    dblTipoCambioCierre & ",'" & _
                    "','" & _
                    "','" & _
                    strDescripAsiento & "','" & _
                    "','" & _
                    "X','') }"
                adoConn.Execute .CommandText
                
                '*** Detalle ***
                .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
                    "WHERE TipoOperacion='" & Codigo_Dinamica_Provision & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                    IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"
                    
                Set adoConsulta = .Execute
        
                Do While Not adoConsulta.EOF
                
                    Select Case Trim(adoConsulta("TipoCuentaInversion"))
                        Case Codigo_CtaInversion
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvInteres
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteres
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaCosto
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaIngresoOperacional
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresVencido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaVacCorrido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaXPagar
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaXCobrar
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresCorrido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvReajusteK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaReajusteK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvFlucMercado
                            curMontoMovimientoMN = curMontoFluctuacionMercado
                            strDescripMovimiento = "Pérdida"
                            If curMontoFluctuacionMercado > 0 Then strDescripMovimiento = "Ganancia"
                            
                        Case Codigo_CtaFlucMercado
                            curMontoMovimientoMN = curMontoFluctuacionMercado
                            
                        Case Codigo_CtaProvInteresVac
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresVac
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaIntCorridoK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvFlucK
                            curMontoMovimientoMN = curMontoProvisionCapital
                            strDescripMovimiento = "Pérdida"
                            If curMontoProvisionCapital > 0 Then strDescripMovimiento = "Ganancia"
                            
                        Case Codigo_CtaFlucK
                            curMontoMovimientoMN = curMontoProvisionCapital
                            
                        Case Codigo_CtaInversionTransito
                            curMontoMovimientoMN = curMontoRenta
                            
                    End Select
                    
                    dblValorTipoCambio = 1
                    
                    strIndDebeHaber = Trim(adoConsulta("IndDebeHaber"))
                    
                    If strIndDebeHaber = "H" Then
                        curMontoMovimientoMN = curMontoMovimientoMN * -1
                        If curMontoMovimientoMN > 0 Then strIndDebeHaber = "D"
                    ElseIf strIndDebeHaber = "D" Then
                        If curMontoMovimientoMN < 0 Then strIndDebeHaber = "H"
                    End If
                    
                    If strIndDebeHaber = "T" Then
                        If curMontoMovimientoMN > 0 Then
                            strIndDebeHaber = "D"
                        Else
                            strIndDebeHaber = "H"
                        End If
                    End If
                    strDescripMovimiento = Trim(adoConsulta("DescripDinamica"))
                    curMontoMovimientoME = 0
                    curMontoContable = curMontoMovimientoMN
       
                    If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
                        curMontoContable = Round(curMontoMovimientoMN * dblTipoCambioCierre, 2)
                        curMontoMovimientoME = curMontoMovimientoMN
                        curMontoMovimientoMN = 0
                        dblValorTipoCambio = dblTipoCambioCierre
                    End If
                                
                    strTipoDocumento = ""
                    strNumDocumento = ""
                    strTipoPersonaContraparte = ""
                    strCodPersonaContraparte = ""
                    strIndContracuenta = ""
                    strCodContracuenta = ""
                    strCodFileContracuenta = ""
                    strCodAnaliticaContracuenta = ""
                    strIndUltimoMovimiento = ""
                    
                    If curMontoContable <> 0 Then
                                
                        '*** Movimiento ***

                        .CommandText = "{ call up_ACAdicAsientoContableDetalle('"
                        If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableDetalleTmp('"
    
                        .CommandText = .CommandText & strNumAsiento & "','" & strCodFondo & "','" & _
                            gstrCodAdministradora & "'," & _
                            CInt(adoConsulta("NumSecuencial")) & ",'" & _
                            strFechaGrabar & "','" & _
                            gstrPeriodoActual & "','" & _
                            gstrMesActual & "','" & _
                            strDescripMovimiento & "','" & _
                            strIndDebeHaber & "','" & _
                            Trim(adoConsulta("CodCuenta")) & "','" & _
                            Trim(adoRegistro("CodMoneda")) & "'," & _
                            CDec(curMontoMovimientoMN) & "," & _
                            CDec(curMontoMovimientoME) & "," & _
                            CDec(curMontoContable) & "," & _
                            dblValorTipoCambio & ",'" & _
                            Trim(adoRegistro("CodFile")) & "','" & _
                            Trim(adoRegistro("CodAnalitica")) & "','" & _
                            strTipoDocumento & "','" & _
                            strNumDocumento & "','" & _
                            strTipoPersonaContraparte & "','" & _
                            strCodPersonaContraparte & "','" & _
                            strIndContracuenta & "','" & _
                            strCodContracuenta & "','" & _
                            strCodFileContracuenta & "','" & _
                            strCodAnaliticaContracuenta & "','" & _
                            strIndUltimoMovimiento & "') }"

                        adoConn.Execute .CommandText
                
                        '*** Validar valor de cuenta contable ***
                        If Trim(adoConsulta("CodCuenta")) = Valor_Caracter Then
                            MsgBox "Registro Nro. " & CStr(intContador) & " de Asiento Contable no tiene cuenta asignada", vbCritical, "Valorización Depósitos"
                            gblnRollBack = True
                            Exit Sub
                        End If
                    
                    End If
                    
                    adoConsulta.MoveNext
                Loop
                adoConsulta.Close: Set adoConsulta = Nothing
                                
                '*** Actualizar el número del parámetro **
                .CommandText = "{ call up_ACActUltNumero('"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNActUltNumeroTmp('"
                
                .CommandText = .CommandText & strCodFondo & "','" & _
                    gstrCodAdministradora & "','" & _
                    Valor_NumComprobante & "','" & _
                    strNumAsiento & "') }"
                    
                adoConn.Execute .CommandText
                                
'                .CommandText = "COMMIT TRAN ProcAsiento"
'                adoConn.Execute .CommandText
        
            End If
                                    
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
                
    End With
    
    Exit Sub
       
Ctrl_Error:
'    adoComm.CommandText = "ROLLBACK TRAN ProcAsiento"
'    adoConn.Execute adoComm.CommandText
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Resume Next
    Me.MousePointer = vbDefault
    
End Sub

Private Sub VencimientoCobertura()

    Dim adoRegistro             As ADODB.Recordset, adoConsulta     As ADODB.Recordset
    Dim dblPrecioCierre         As Double, dblPrecioPromedio        As Double
    Dim dblTirCierre            As Double, dblFactorDiario          As Double
    Dim dblTasaInteres          As Double, dblFactorDiarioCupon     As Double
    Dim dblPrecioUnitario       As Double, dblValorPromedioKardex   As Double
    Dim dblInteresCorridoPromedio As Double, dblTirOperacionKardex  As Double
    Dim dblTirPromedioKardex    As Double, dblTirNetaKardex         As Double
    Dim curSaldoInversion       As Currency, curSaldoInteresCorrido As Currency
    Dim curSaldoFluctuacion     As Currency, curValorAnterior       As Currency
    Dim curValorActual          As Currency, curMontoRenta          As Currency
    Dim curMontoContable        As Currency, curMontoMovimientoMN   As Currency
    Dim curMontoMovimientoME    As Currency, curSaldoValorizar      As Currency
    Dim curCantMovimiento       As Currency, curKarValProm          As Currency
    Dim curValorMovimiento      As Currency, curSaldoInicialKardex  As Currency
    Dim curSaldoFinalKardex     As Currency, curValorSaldoKardex    As Currency
    Dim curValComi              As Currency, curVacCorrido          As Currency
    Dim curSaldoAmortizacion    As Currency
    Dim intCantRegistros        As Integer, intContador             As Integer
    Dim intRegistro             As Integer, intBaseCalculo          As Integer
    Dim intDiasPlazo            As Integer, intDiasDeRenta          As Integer
    Dim strNumAsiento           As String, strDescripAsiento        As String
    Dim strNumOperacion         As String, strNumKardex             As String
    Dim strNumCaja              As String, strFechaPago             As String
    Dim strCodTitulo            As String, strCodEmisor             As String
    Dim strDescripMovimiento    As String, strIndDebeHaber          As String
    Dim strCodCuenta            As String, strFiles                 As String
    Dim strCodFile              As String, strModalidadInteres      As String
    Dim strCodTasa              As String, strIndCuponCero          As String
    Dim strCodDetalleFile       As String, strCodAnalitica          As String
    Dim strCodSubDetalleFile    As String, strFechaGrabar           As String
    Dim strSQLOperacion         As String, strSQLKardex             As String
    Dim strSQLOrdenCaja         As String, strSQLOrdenCajaDetalle   As String
    Dim strIndUltimoMovimiento  As String, strTipoMovimientoKardex  As String
    Dim blnVenceTitulo          As Boolean, blnVenceCupon           As Boolean
    Dim dblTipoCambioCierre     As Double


    '*** Verificación de Vencimiento de Valores de Depósito ***
    frmMainMdi.stbMdi.Panels(3).Text = "Verificando Vencimiento de Valores de Depósito..."
       
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT IK.CodTitulo,II.CodMoneda,IK.CodAnalitica,IK.CodFile,SaldoFinal,CodDetalleFile,CodSubDetalleFile,MontoCobertura," & _
            "ValorNominal,CodTipoTasa,BaseAnual,CodTasa,TasaInteres,DiasPlazo,IndCuponCero,FechaEmision,Nemotecnico,TipoCambioSpot,TipoCambioFuturo,ICO.FechaVencimiento " & _
            "FROM InversionKardex IK JOIN InversionCobertura ICO ON(ICO.CodTitulo=IK.CodTitulo AND ICO.CodFondo=IK.CodFondo AND ICO.CodAdministradora=IK.CodAdministradora) " & _
            "JOIN InstrumentoInversion II ON(II.CodTitulo=IK.CodTitulo) " & _
            "WHERE IK.CodFile IN ('013') AND IK.FechaOperacion <='" & strFechaCierre & "' AND SaldoFinal > 0 AND " & _
            "IK.CodAdministradora='" & gstrCodAdministradora & "' AND IK.CodFondo='" & strCodFondo & "' AND IndUltimoMovimiento='X'"
        Set adoRegistro = .Execute
    
        Do Until adoRegistro.EOF
            blnVenceTitulo = False: blnVenceCupon = False
            
            '*** Obtener Secuenciales ***
            strNumAsiento = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumComprobante)
            strNumOperacion = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOperacion)
            strNumKardex = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumKardex)
            strNumCaja = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOrdenCaja)
            
            '*** Fecha de Vencimiento del Título = Fecha de Cierre ***
            If Convertyyyymmdd(adoRegistro("FechaVencimiento")) = strFechaCierre Then blnVenceTitulo = True
            
            '*** Si vence el título o el cupón ***
            If blnVenceTitulo Or blnVenceCupon Then
                strCodTitulo = Trim(adoRegistro("CodTitulo"))
                strCodFile = Trim(adoRegistro("CodFile"))
                strCodDetalleFile = Trim(adoRegistro("CodDetalleFile"))
                strCodSubDetalleFile = Trim(adoRegistro("CodSubDetalleFile"))
                strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
                strCodEmisor = Trim(adoRegistro("CodEmisor"))
                strModalidadInteres = Trim(adoRegistro("CodDetalleFile"))
                strCodTasa = Trim(adoRegistro("CodTasa"))
                dblTasaInteres = CDbl(adoRegistro("TasaInteres"))
                intDiasPlazo = CInt(adoRegistro("DiasPlazo"))
                strIndCuponCero = Trim(adoRegistro("IndCuponCero"))
                curSaldoValorizar = CCur(adoRegistro("SaldoFinal"))
                curKarValProm = CDbl(adoRegistro("ValorPromedio"))
                intDiasDeRenta = DateDiff("d", CVDate(adoRegistro("FechaEmision")), gdatFechaActual) + 1
            
                Set adoConsulta = New ADODB.Recordset
                        
                '*** Verificar Dinamica Contable ***
                .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
                    "WHERE TipoOperacion='" & Codigo_Dinamica_Vencimiento & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                    IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"

                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    If CInt(adoConsulta("NumRegistros")) > 0 Then
                        intCantRegistros = CInt(adoConsulta("NumRegistros"))
                    Else
                        MsgBox "NO EXISTE Dinámica Contable para la valorización", vbCritical
                        adoConsulta.Close: Set adoConsulta = Nothing
                        Exit Sub
                    End If
                End If
                adoConsulta.Close
                
                '*** Obtener la Fecha de Pago ***
                .CommandText = "SELECT FechaPago FROM InstrumentoInversionCalendario " & _
                    "WHERE CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    strFechaPago = Convertyyyymmdd(adoConsulta("FechaPago"))
                End If
                adoConsulta.Close
            
                '*** Obtener las cuentas de inversión ***
                Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, Trim(adoRegistro("CodMoneda")))
                
                '*** Obtener tipo de cambio ***
                'dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
                dblTipoCambioCierre = ObtenerValorTipoCambio(adoRegistro("CodMoneda"), Codigo_Moneda_Local, strFechaCierre, strFechaCierre, gstrCodClaseTipoCambioOperacionFondo, Codigo_Valor_TipoCambioCompra)

                
                '*** Obtener Saldo de Inversión ***
                If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaInversion & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                    "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curSaldoInversion = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                '*** Obtener Saldo Compromiso ***
                If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaCompromiso & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                    "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curCtaCompromiso = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                '*** Obtener Saldo de Provisión ***
                If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaProvInteres & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                    "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curSaldoFluctuacion = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                If blnVenceTitulo Then
                    '*** Calculos ***
                    curCtaMN = adoRegistro("MontoCobertura")
                    
                    curCtaXCobrar = Round(curSaldoInversion + curSaldoInteresCorrido + curSaldoFluctuacion, 2)
                    curCtaInversion = curSaldoInversion
                    curCtaCosto = curSaldoInversion
                    curCtaInteresCorrido = curSaldoInteresCorrido
                    curCtaProvInteres = curSaldoFluctuacion
                    curCtaIngresoOperacional = curCtaXCobrar - curCtaProvInteres
                    
                    curCantMovimiento = CCur(adoRegistro("SaldoFinal"))
                    dblPrecioUnitario = curKarValProm
                    curValorMovimiento = dblPrecioUnitario * curCantMovimiento * CCur(adoRegistro("ValorNominal")) * -1
                    curSaldoInicialKardex = CCur(adoRegistro("SaldoFinal"))
                    curSaldoFinalKardex = 0
                    curValorSaldoKardex = 0
                    dblValorPromedioKardex = 0
                    dblInteresCorridoPromedio = 0
                    curValComi = 0
                    curVacCorrido = 0
                    dblTirOperacionKardex = 0
                    dblTirPromedioKardex = 0
                    dblTirNetaKardex = 0
                    curSaldoAmortizacion = 0
                End If
                
                '************************
                '*** Armar sentencias ***
                '************************
                strDescripAsiento = "Vencimiento" & Space(1) & "(" & strCodFile & "-" & strCodAnalitica & ")"
                '*** Operación ***
                strSQLOperacion = "{ call up_IVAdicInversionOperacion('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strNumOperacion & "','" & strFechaSiguiente & "','" & strCodTitulo & "','" & Left(strFechaSiguiente, 4) & "','" & _
                    Mid(strFechaSiguiente, 5, 2) & "','','" & Estado_Activo & "','" & strCodAnalitica & "','" & _
                    strCodFile & "','" & strCodAnalitica & "','" & strCodDetalleFile & "','" & strCodSubDetalleFile & "','" & _
                    Codigo_Caja_Vencimiento & "','','','" & strDescripAsiento & "','" & strCodEmisor & "','" & _
                    "','','','" & strFechaSiguiente & "','" & strFechaSiguiente & "','" & _
                    strFechaSiguiente & "','" & strCodMoneda & "'," & CDec(adoRegistro("SaldoFinal")) & "," & CDec(gdblTipoCambio) & "," & _
                    CDec(adoRegistro("ValorNominal")) & "," & CDec(adoRegistro("PrecioUnitario")) & "," & CDec(adoRegistro("MontoMovimiento")) & "," & CDec(adoRegistro("SaldoInteresCorrido")) & "," & _
                    "0,0,0,0,0,0,0," & CDec(curCtaXCobrar) & ",0,0,0,0,0,0,0,0,0," & _
                    "0,0,0,0,'X','" & strNumAsiento & "','','','" & _
                    "','','','',0,'','','','',''," & CDec(dblTasaInteres) & "," & _
                    "0,0,'','','','" & gstrLogin & "') }"
                                                
                strIndUltimoMovimiento = "X"
                strTipoMovimientoKardex = "S"
                '*** Kardex ***
                strSQLKardex = "{ call up_IVAdicInversionKardex('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strCodTitulo & "','" & strNumKardex & "','" & strFechaSiguiente & "','" & Left(strFechaSiguiente, 4) & "','" & _
                    Mid(strFechaSiguiente, 5, 2) & "','','" & strNumOperacion & "','" & strCodEmisor & "','','O','" & _
                    strFechaSiguiente & "','" & strTipoMovimientoKardex & "','O'," & curCantMovimiento & ",'" & strCodMoneda & "'," & _
                    dblPrecioUnitario & "," & curValorMovimiento & "," & curValComi & "," & curSaldoInicialKardex & "," & _
                    curSaldoFinalKardex & "," & curValorSaldoKardex & ",'" & strDescripAsiento & "'," & dblValorPromedioKardex & ",'" & _
                    strIndUltimoMovimiento & "','" & strCodFile & "','" & strCodAnalitica & "'," & dblInteresCorridoPromedio & "," & _
                    curSaldoInteresCorrido & "," & dblTirOperacionKardex & "," & dblTirPromedioKardex & "," & curVacCorrido & "," & _
                    dblTirNetaKardex & "," & curSaldoAmortizacion & ") }"
    
                '*** Orden de Cobro/Pago ***
                strSQLOrdenCaja = "{ call up_ACAdicMovimientoFondo('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strNumCaja & "','" & strFechaSiguiente & "','" & Trim(frmMainMdi.Tag) & "','" & strNumOperacion & "','" & strFechaPago & "','" & _
                    strNumAsiento & "','','E','" & strCtaXCobrar & "'," & CDec(curCtaXCobrar) & ",'" & _
                    strCodFile & "','" & strCodAnalitica & "','" & strCodMoneda & "','" & _
                    strDescripAsiento & "','" & Codigo_Caja_Vencimiento & "','','" & Estado_Caja_NoConfirmado & "','','','','','','','','','','" & gstrLogin & "') }"
                
                '*** Orden de Cobro/Pago Detalle ***
                strSQLOrdenCajaDetalle = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strNumCaja & "','" & strFechaSiguiente & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripAsiento & "','" & _
                    "H','" & strCtaXCobrar & "'," & CDec(curCtaXCobrar) * -1 & ",'" & _
                    strCodFile & "','" & strCodAnalitica & "','" & strCodMoneda & "','') }"
                                
                '*** Monto Orden ***
                If curCtaXCobrar > 0 Then
                                                                                            
                    On Error GoTo Ctrl_Error
                    
'                    .CommandText = "BEGIN TRANSACTION ProcAsiento"
'                    adoConn.Execute .CommandText
                                                            
                    '*** Actualizar indicador de último movimiento en Kardex ***
                    .CommandText = "UPDATE InversionKardex SET IndUltimoMovimiento='' " & _
                        "WHERE CodAnalitica='" & strCodAnalitica & "' AND CodFile='" & strCodFile & "' AND " & _
                        "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                        "IndUltimoMovimiento='X'"
        
                    adoConn.Execute .CommandText
        
                    '*** Inserta movimiento en el kardex ***
                    adoConn.Execute strSQLKardex
                    
                    '*** Operación ***
                    adoConn.Execute strSQLOperacion
                    
                    '*** Contabilizar ***
                    strFechaGrabar = strFechaSiguiente & Space(1) & Format(Time, "hh:mm")
                    
                    '*** Cabecera ***
                    .CommandText = "{ call up_ACAdicAsientoContable('"
                    .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
                        strFechaGrabar & "','" & _
                        Left(strFechaSiguiente, 4) & "','" & Mid(strFechaSiguiente, 5, 2) & "','" & _
                        "','" & _
                        strDescripAsiento & "','" & Trim(adoRegistro("CodMoneda")) & "','" & _
                        Codigo_Moneda_Local & "','" & _
                        "','" & _
                        "'," & _
                        CDec(curCtaXCobrar) & ",'" & Estado_Activo & "'," & _
                        intCantRegistros & ",'" & _
                        strFechaSiguiente & Space(1) & Format(Time, "hh:ss") & "','" & _
                        strCodModulo & "','" & _
                        "'," & _
                        dblTipoCambioCierre & ",'" & _
                        "','" & _
                        "','" & _
                        strDescripAsiento & "','" & _
                        "','" & _
                        "X','') }"
                    adoConn.Execute .CommandText
                    
                    '*** Detalle ***
                    .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
                        "WHERE TipoOperacion='" & Codigo_Dinamica_Vencimiento & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                        strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                        IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"

                    Set adoConsulta = .Execute
            
                    Do While Not adoConsulta.EOF
                    
                        Select Case Trim(adoConsulta("TipoCuentaInversion"))
                            Case Codigo_CtaInversion
                                curMontoMovimientoMN = curCtaInversion
                                
                            Case Codigo_CtaProvInteres
                                curMontoMovimientoMN = curCtaProvInteres
                                
                            Case Codigo_CtaInteres
                                curMontoMovimientoMN = curCtaInteres
                                
                            Case Codigo_CtaCosto
                                curMontoMovimientoMN = curCtaCosto
                                
                            Case Codigo_CtaIngresoOperacional
                                curMontoMovimientoMN = curCtaIngresoOperacional
                                
                            Case Codigo_CtaInteresVencido
                                curMontoMovimientoMN = curCtaInteresVencido
                                
                            Case Codigo_CtaVacCorrido
                                curMontoMovimientoMN = curCtaVacCorrido
                                
                            Case Codigo_CtaXPagar
                                curMontoMovimientoMN = curCtaXPagar
                                
                            Case Codigo_CtaXCobrar
                                curMontoMovimientoMN = curCtaXCobrar
                                
                            Case Codigo_CtaInteresCorrido
                                curMontoMovimientoMN = curCtaInteresCorrido
                                
                            Case Codigo_CtaProvReajusteK
                                curMontoMovimientoMN = curCtaProvReajusteK
                                
                            Case Codigo_CtaReajusteK
                                curMontoMovimientoMN = curCtaReajusteK
                                
                            Case Codigo_CtaProvFlucMercado
                                curMontoMovimientoMN = curCtaProvFlucMercado
                                
                            Case Codigo_CtaFlucMercado
                                curMontoMovimientoMN = curCtaFlucMercado
                                
                            Case Codigo_CtaProvInteresVac
                                curMontoMovimientoMN = curCtaProvInteresVac
                                
                            Case Codigo_CtaInteresVac
                                curMontoMovimientoMN = curCtaInteresVac
                                
                            Case Codigo_CtaIntCorridoK
                                curMontoMovimientoMN = curCtaIntCorridoK
                                
                            Case Codigo_CtaProvFlucK
                                curMontoMovimientoMN = curCtaProvFlucK
                                
                            Case Codigo_CtaFlucK
                                curMontoMovimientoMN = curCtaFlucK
                                
                            Case Codigo_CtaInversionTransito
                                curMontoMovimientoMN = curCtaInversionTransito
                                
                            Case Codigo_CtaME
                                curMontoMovimientoMN = curCtaME
                                
                            Case Codigo_CtaMN
                                curMontoMovimientoMN = curCtaMN
                                
                            Case Codigo_CtaCompromiso
                                curMontoMovimientoMN = curCtaCompromiso
                                
                            Case Codigo_CtaResponsabilidad
                                curMontoMovimientoMN = curCtaResponsabilidad
                                
                        End Select
                        
                        strIndDebeHaber = Trim(adoConsulta("IndDebeHaber"))
                        If strIndDebeHaber = "H" Then
                            curMontoMovimientoMN = curMontoMovimientoMN * -1
                            If curMontoMovimientoMN > 0 Then strIndDebeHaber = "D"
                        ElseIf strIndDebeHaber = "D" Then
                            If curMontoMovimientoMN < 0 Then strIndDebeHaber = "H"
                        End If
                        
                        If strIndDebeHaber = "T" Then
                            If curMontoMovimientoMN > 0 Then
                                strIndDebeHaber = "D"
                            Else
                                strIndDebeHaber = "H"
                            End If
                        End If
                        strDescripMovimiento = Trim(adoConsulta("DescripDinamica"))
                        curMontoMovimientoME = 0
                        curMontoContable = curMontoMovimientoMN
            
                        If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
                            curMontoContable = Round(curMontoMovimientoMN * dblTipoCambioCierre, 2)
                            curMontoMovimientoME = curMontoMovimientoMN
                            curMontoMovimientoMN = 0
                        End If
                                    
                        '*** Movimiento ***
                        .CommandText = "{ call up_ACAdicAsientoContableDetalle('"
                        .CommandText = .CommandText & strNumAsiento & "','" & strCodFondo & "','" & _
                            gstrCodAdministradora & "'," & _
                            CInt(adoConsulta("NumSecuencial")) & ",'" & _
                            strFechaGrabar & "','" & _
                            Left(strFechaSiguiente, 4) & "','" & _
                            Mid(strFechaSiguiente, 5, 2) & "','" & _
                            strDescripMovimiento & "','" & _
                            strIndDebeHaber & "','" & _
                            Trim(adoConsulta("CodCuenta")) & "','" & _
                            Trim(adoRegistro("CodMoneda")) & "'," & _
                            CDec(curMontoMovimientoMN) & "," & _
                            CDec(curMontoMovimientoME) & "," & _
                            CDec(curMontoContable) & ",'" & _
                            Trim(adoRegistro("CodFile")) & "','" & _
                            Trim(adoRegistro("CodAnalitica")) & "','','') }"
                        adoConn.Execute .CommandText
                    
                        '*** Saldos ***
'                        .CommandText = "{ call up_ACGenPartidaContableSaldos('"
'                        .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & _
'                            Left(strFechaSiguiente, 4) & "','" & Mid(strFechaSiguiente, 5, 2) & "','" & _
'                            Trim(adoConsulta("CodCuenta")) & "','" & _
'                            Trim(adoRegistro("CodFile")) & "','" & _
'                            Trim(adoRegistro("CodAnalitica")) & "','" & _
'                            strFechaSiguiente & "','" & _
'                            strFechaSubSiguiente & "'," & _
'                            CDec(curMontoMovimientoMN) & "," & _
'                            CDec(curMontoMovimientoME) & "," & _
'                            CDec(curMontoContable) & ",'" & _
'                            strIndDebeHaber & "','" & _
'                            Trim(adoRegistro("CodMoneda")) & "') }"
'                        adoConn.Execute .CommandText
                                        
                        '*** Validar valor de cuenta contable ***
                        If Trim(adoConsulta("CodCuenta")) = Valor_Caracter Then
                            MsgBox "Registro Nro. " & CStr(intContador) & " de Asiento Contable no tiene cuenta asignada", vbCritical, "Valorización Depósitos"
                            gblnRollBack = True
                            Exit Sub
                        End If
                        
                        adoConsulta.MoveNext
                    Loop
                    adoConsulta.Close: Set adoConsulta = Nothing
                                    
                    '*** Orden de Cobro ***
                    adoConn.Execute strSQLOrdenCaja
                    adoConn.Execute strSQLOrdenCajaDetalle
        
                    '*** Actualizar Secuenciales **
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumComprobante & "','" & strNumAsiento & "') }"
                    adoConn.Execute .CommandText
                    
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumOperacion & "','" & strNumOperacion & "') }"
                    adoConn.Execute .CommandText
                    
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumKardex & "','" & strNumKardex & "') }"
                    adoConn.Execute .CommandText
                    
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumOrdenCaja & "','" & strNumCaja & "') }"
                    adoConn.Execute .CommandText
                                                                            
'                    .CommandText = "COMMIT TRANSACTION ProcAsiento"
'                    adoConn.Execute .CommandText
            
                End If
            
            End If

            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
                
    End With
    
    Exit Sub
       
Ctrl_Error:
'    adoComm.CommandText = "ROLLBACK TRANSACTION ProcAsiento"
'    adoConn.Execute adoComm.CommandText
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
    
End Sub

Private Sub VencimientoGastosFondo()

    Dim adoRegistro             As ADODB.Recordset
    Dim adoConsulta             As ADODB.Recordset
    Dim adoAuxiliar             As ADODB.Recordset
    Dim strNumCaja              As String, strCodFile                   As String
    Dim strCodDetalleFile       As String, strDescripGasto              As String
    Dim strSQLOrdenCaja         As String, strSQLOrdenCajaDetalle       As String
    Dim strSQLOrdenCajaMN       As String, strSQLOrdenCajaDetalleMN     As String
    Dim strSQLOrdenCajaDetalleI As String, strSQLOrdenCajaDetalleMNI    As String
    Dim strIndDetraccion        As String, strIndImpuesto               As String
    Dim strIndRetencion         As String, strCodCreditoFiscal          As String
    Dim strCodAnalitica         As String, strCodMonedaGasto            As String
    Dim strFechaGrabar          As String
    Dim curSaldoProvision       As Currency, curValorImpuesto           As Currency
    Dim blnVenceGasto           As Boolean

    frmMainMdi.stbMdi.Panels(3).Text = "Vencimiento Gastos del Fondo..."
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT * FROM FondoGasto " & _
            "WHERE FechaFinal<'" & strFechaSiguiente & "' AND CodFondo='" & strCodFondo & "' AND " & _
            "CodAdministradora='" & gstrCodAdministradora & "' AND CodFile='099' AND IndVigente='X'"
        Set adoRegistro = .Execute
        
        Do While Not adoRegistro.EOF
            blnVenceGasto = False
            strCodCreditoFiscal = Trim(adoRegistro("CodCreditoFiscal"))
            
            '*** Obtener Secuenciales ***
            strNumCaja = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOrdenCaja)
            
            '*** Fecha de Vencimiento del Gasto = Fecha de Cierre ***
            If Convertyyyymmdd(adoRegistro("FechaFinal")) = strFechaCierre Then blnVenceGasto = True
            
            If adoRegistro("CodTipoGasto") = Codigo_Gasto_MismoDia Then blnVenceGasto = False
            
            '*** Si vence la provisión del Gasto ***
            If blnVenceGasto Then
                strCodFile = Trim(adoRegistro("CodFile"))
                strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
                strCodMonedaGasto = Trim(adoRegistro("CodMoneda"))
                        
                Set adoConsulta = New ADODB.Recordset
                
                .CommandText = "SELECT CodDetalleFile FROM InversionDetalleFile " & _
                    "WHERE CodFile='" & strCodFile & "' AND CodDetalleFile='" & adoRegistro("CodCuenta") & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    strCodDetalleFile = adoConsulta("CodDetalleFile")
                End If
                adoConsulta.Close
            
                '*** Obtener Descripción del Gasto ***
                If strCodFile = "099" Then
                    .CommandText = "SELECT DescripCuenta FROM PlanContable WHERE CodCuenta='" & Trim(adoRegistro("CodCuenta")) & "'"
                Else
                    .CommandText = "SELECT DescripComision DescripCuenta FROM ComisionEmpresa WHERE CodDetalleFile='" & Trim(adoRegistro("CodCuenta")) & "'"
                End If
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    strDescripGasto = Trim(adoConsulta("DescripCuenta"))
                End If
                adoConsulta.Close
            
                '*** Obtener las cuentas de inversión ***
                Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, Trim(adoRegistro("CodMoneda")))
            
                '*** Obtener Saldo de Inversión ***
                If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaXPagar & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curSaldoProvision = Abs(CDbl(adoConsulta("Saldo")))
                End If
                adoConsulta.Close
                                                                                                                    
                .CommandText = "SELECT CodDetraccionSiNo,CodFormaPagoDetraccion,CodMonedaDetraccion,CodFileDetraccion,CodAnaliticaDetraccion,MontoDetraccion,TipoCambioPago,MontoPago,MontoTotal,FechaPago,Importe,ValorImpuesto,ValorTotal,CodTipoComprobante " & _
                    "FROM RegistroCompra WHERE NumGasto=" & CInt(adoRegistro("NumGasto")) & " AND CodFondo='" & strCodFondo & "' AND " & _
                    "CodAdministradora='" & gstrCodAdministradora & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    Set adoAuxiliar = New ADODB.Recordset
                    
                    .CommandText = "SELECT IndImpuesto,IndRetencion " & _
                        "FROM TipoComprobantePago WHERE CodTipoComprobantePago='" & adoConsulta("CodTipoComprobante") & "'"
                    Set adoAuxiliar = .Execute
            
                    If Not adoAuxiliar.EOF Then
                        strIndImpuesto = Trim(adoAuxiliar("IndImpuesto"))
                        strIndRetencion = Trim(adoAuxiliar("IndRetencion"))
                
'                        strCtaImpuesto = ObtenerCuentaAdministracion("025", "R")
'                        If strIndRetencion = Valor_Indicador Then strCtaImpuesto = ObtenerCuentaAdministracion("036", "R")
                    End If
                    adoAuxiliar.Close: Set adoAuxiliar = Nothing
                
                    If adoConsulta("CodDetraccionSiNo") = Codigo_Respuesta_Si Then
                        strIndDetraccion = Valor_Indicador
                        
                        '*** Orden de Cobro/Pago ***
                        strSQLOrdenCaja = "{ call up_ACAdicMovimientoFondo('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                            strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "','" & Trim(frmMainMdi.Tag) & "','','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "','" & _
                            "','','S','" & strCtaXPagar & "'," & CDec(adoConsulta("MontoPago")) * -1 & ",'" & _
                            strCodFile & "','" & strCodAnalitica & "','" & strCodMonedaGasto & "','" & _
                            strDescripGasto & "','" & Codigo_Caja_Gasto & "','','" & Estado_Caja_NoConfirmado & "','','','','','','','','','','" & gstrLogin & "') }"
                                        
                        '*** Orden de Cobro/Pago Detalle Impuesto ***
                        curValorImpuesto = CCur(adoConsulta("ValorImpuesto") * (1 - gdblTasaDetraccion))
                        
                        If strIndRetencion = Valor_Indicador Then
'                            strSQLOrdenCajaDetalleI = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "',2,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                "H','" & strCtaImpuesto & "'," & CDec(curValorImpuesto * -1) & ",'" & _
                                strCodFile & "','" & strCodAnalitica & "','" & strCodMonedaGasto & "','') }"
                                
                            '*** Orden de Cobro/Pago Detalle ***
                            strSQLOrdenCajaDetalle = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                "D','" & strCtaXPagar & "'," & CDec(adoConsulta("MontoPago")) & ",'" & _
                                strCodFile & "','" & strCodAnalitica & "','" & strCodMonedaGasto & "','') }"
                        Else
                            If strCodCreditoFiscal = Codigo_Tipo_Credito_RentaNoGravada Then
                                strSQLOrdenCajaDetalleI = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                    strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "',2,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                    "D','" & strCtaImpuesto & "'," & CDec(curValorImpuesto) & ",'" & _
                                    strCodFile & "','" & strCodAnalitica & "','" & strCodMonedaGasto & "','') }"
                                    
                                '*** Orden de Cobro/Pago Detalle ***
                                strSQLOrdenCajaDetalle = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                    strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                    "D','" & strCtaXPagar & "'," & CDec(adoConsulta("MontoPago") - curValorImpuesto) & ",'" & _
                                    strCodFile & "','" & strCodAnalitica & "','" & strCodMonedaGasto & "','') }"
                            Else
                                '*** Orden de Cobro/Pago Detalle ***
                                strSQLOrdenCajaDetalle = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                    strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                    "D','" & strCtaXPagar & "'," & CDec(adoConsulta("MontoPago")) & ",'" & _
                                    strCodFile & "','" & strCodAnalitica & "','" & strCodMonedaGasto & "','') }"
                            End If
                        End If
                    Else
                        '*** Orden de Cobro/Pago ***
                        strSQLOrdenCaja = "{ call up_ACAdicMovimientoFondo('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                            strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "','" & Trim(frmMainMdi.Tag) & "','','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "','" & _
                            "','','S','" & strCtaXPagar & "'," & CDec(adoConsulta("ValorTotal") * -1) & ",'" & _
                            strCodFile & "','" & strCodAnalitica & "','" & strCodMonedaGasto & "','" & _
                            strDescripGasto & "','" & Codigo_Caja_Vencimiento & "','','" & Estado_Caja_NoConfirmado & "','','','','','','','','','','" & gstrLogin & "') }"
                
                        If strCodCreditoFiscal = Codigo_Tipo_Credito_RentaNoGravada Then
                            '*** Orden de Cobro/Pago Detalle ***
                            strSQLOrdenCajaDetalle = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                "D','" & strCtaXPagar & "'," & CDec(adoConsulta("ValorTotal")) & ",'" & _
                                strCodFile & "','" & strCodAnalitica & "','" & strCodMonedaGasto & "','') }"
                        Else
                            '*** Orden de Cobro/Pago Detalle ***
                            strSQLOrdenCajaDetalle = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                "D','" & strCtaXPagar & "'," & CDec(adoConsulta("Importe")) & ",'" & _
                                strCodFile & "','" & strCodAnalitica & "','" & strCodMonedaGasto & "','') }"
                        End If
                    
                        If strCodCreditoFiscal <> Codigo_Tipo_Credito_RentaNoGravada Then
                            '*** Orden de Cobro/Pago Detalle ***
                            If strIndRetencion = Valor_Indicador Then
                                strSQLOrdenCajaDetalleI = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                    strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "',2,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                    "H','" & strCtaImpuesto & "'," & CDec(adoConsulta("ValorImpuesto") * -1) & ",'" & _
                                    strCodFile & "','" & strCodAnalitica & "','" & strCodMonedaGasto & "','') }"
                            Else
                                strSQLOrdenCajaDetalleI = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                    strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "',2,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                    "D','" & strCtaImpuesto & "'," & CDec(adoConsulta("ValorImpuesto")) & ",'" & _
                                    strCodFile & "','" & strCodAnalitica & "','" & strCodMonedaGasto & "','') }"
                            End If
                        End If
                    End If
                                                                 
                    On Error GoTo Ctrl_Error
                
                    '*** Orden de Cobro ***
                    adoConn.Execute strSQLOrdenCaja
                    adoConn.Execute strSQLOrdenCajaDetalle
                    If strCodCreditoFiscal = Codigo_Tipo_Credito_RentaNoGravada Then
                        adoConn.Execute strSQLOrdenCajaDetalleI
                    End If
                    
                    If strIndDetraccion = Valor_Indicador Then
                        '*** Actualizar el número del parámetro **
                        .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & _
                            gstrCodAdministradora & "','" & Valor_NumOrdenCaja & "','" & strNumCaja & "') }"
                        adoConn.Execute .CommandText
                    
                        '*** Obtener Secuenciales ***
                        strNumCaja = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOrdenCaja)
            
                        '*** Orden de Cobro/Pago Detracción ***
                        strSQLOrdenCajaMN = "{ call up_ACAdicMovimientoFondo('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                            strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "','" & Trim(frmMainMdi.Tag) & "','','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "','" & _
                            "','','S','" & strCtaXPagar & "'," & CDec(adoConsulta("MontoDetraccion") * -1) & ",'" & _
                            strCodFile & "','" & strCodAnalitica & "','" & adoConsulta("CodMonedaDetraccion") & "','" & _
                            strDescripGasto & "','" & Codigo_Caja_Gasto & "','','" & Estado_Caja_NoConfirmado & "','','','','','','','','','','" & gstrLogin & "') }"
                                        
                        '*** Orden de Cobro/Pago Detalle Detracción Impuesto ***
                        curValorImpuesto = CCur(adoConsulta("ValorImpuesto") * gdblTasaDetraccion)
                                                                            
                        If strIndRetencion = Valor_Indicador Then
'                            strSQLOrdenCajaDetalleMNI = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "',2,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                "H','" & strCtaImpuesto & "'," & CDec(curValorImpuesto * -1) & ",'" & _
                                strCodFile & "','" & strCodAnalitica & "','" & strCodMonedaGasto & "','') }"
                                
                            '*** Orden de Cobro/Pago Detalle Detracción ***
                            If strCodMonedaGasto <> Codigo_Moneda_Local Then
                                strSQLOrdenCajaDetalleMN = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                    strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                    "D','" & strCtaXPagar & "'," & CDec(adoConsulta("MontoTotal") - adoConsulta("MontoPago")) & ",'" & _
                                    strCodFile & "','" & strCodAnalitica & "','" & strCodMonedaGasto & "','') }"
                            Else
                                strSQLOrdenCajaDetalleMN = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                    strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                    "D','" & strCtaXPagar & "'," & CDec(adoConsulta("MontoDetraccion")) & ",'" & _
                                    strCodFile & "','" & strCodAnalitica & "','" & Codigo_Moneda_Local & "','') }"
                            End If
                        Else
                            If strCodCreditoFiscal = Codigo_Tipo_Credito_RentaNoGravada Then
                                strSQLOrdenCajaDetalleMNI = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                    strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "',2,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                    "D','" & strCtaImpuesto & "'," & CDec(curValorImpuesto) & ",'" & _
                                    strCodFile & "','" & strCodAnalitica & "','" & strCodMonedaGasto & "','') }"
                            End If
                            
                            If strCodCreditoFiscal = Codigo_Tipo_Credito_RentaNoGravada Then
                                '*** Orden de Cobro/Pago Detalle Detracción ***
                                strSQLOrdenCajaDetalleMN = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                    strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                    "D','" & strCtaXPagar & "'," & CDec(adoConsulta("MontoTotal") - adoConsulta("MontoPago")) & ",'" & _
                                    strCodFile & "','" & strCodAnalitica & "','" & strCodMonedaGasto & "','') }"
                            Else
                                '*** Orden de Cobro/Pago Detalle Detracción ***
                                If strCodMonedaGasto <> Codigo_Moneda_Local Then
                                    strSQLOrdenCajaDetalleMN = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                        strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                        "D','" & strCtaXPagar & "'," & CDec(adoConsulta("MontoTotal") - adoConsulta("MontoPago")) & ",'" & _
                                        strCodFile & "','" & strCodAnalitica & "','" & strCodMonedaGasto & "','') }"
                                Else
                                    strSQLOrdenCajaDetalleMN = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                        strNumCaja & "','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                        "D','" & strCtaXPagar & "'," & CDec(adoConsulta("MontoDetraccion")) & ",'" & _
                                        strCodFile & "','" & strCodAnalitica & "','" & Codigo_Moneda_Local & "','') }"
                                End If
                            End If
                        End If
                        
                        adoConn.Execute strSQLOrdenCajaMN
                        adoConn.Execute strSQLOrdenCajaDetalleMN
                        If strCodCreditoFiscal = Codigo_Tipo_Credito_RentaNoGravada Then
                            adoConn.Execute strSQLOrdenCajaDetalleMNI
                        End If
                        
                        If strIndRetencion <> Valor_Indicador Then
                            .CommandText = "UPDATE MovimientoFondo SET ValorTipoCambio=" & adoConsulta("TipoCambioPago") & " " & _
                                "WHERE NumOrdenCobroPago='" & strNumCaja & "' AND CodFondo='" & strCodFondo & "' AND " & _
                                "CodAdministradora='" & gstrCodAdministradora & "'"
                            adoConn.Execute .CommandText
                        End If
                    End If
                
                    '*** Actualizar el número del parámetro **
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & _
                        gstrCodAdministradora & "','" & Valor_NumOrdenCaja & "','" & strNumCaja & "') }"
                    adoConn.Execute .CommandText
                End If
                
                '*** Actualizar Vigencia del Gasto ***---AQUI-------
                '*** El Gasto Puede Estar Totalmente Provisionado
                '    Pero no ha sido contabilizado asi que sigue vigente *********------
                '.CommandText = "UPDATE FondoGasto SET IndVigente='',FechaConfirma='" & gstrFechaActual & "' " & _
                    '"WHERE NumGasto=" & adoRegistro("NumGasto") & " AND " & _
                    '"CodCuenta='" & Trim(adoRegistro("CodCuenta")) & "' AND CodFondo='" & strCodFondo & "' AND " & _
                    '"CodAdministradora='" & gstrCodAdministradora & "' AND IndVigente='X'"
                'adoConn.Execute .CommandText
                .CommandText = "UPDATE FondoGasto SET FechaConfirma='" & gstrFechaActual & "' " & _
                    "WHERE NumGasto=" & adoRegistro("NumGasto") & " AND " & _
                    "CodCuenta='" & Trim(adoRegistro("CodCuenta")) & "' AND CodFondo='" & strCodFondo & "' AND " & _
                    "CodAdministradora='" & gstrCodAdministradora & "' AND IndVigente='X'"
                adoConn.Execute .CommandText
            End If
'            adoConsulta.Close: Set adoConsulta = Nothing
                                                                                                   
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    Exit Sub
  
Ctrl_Error:
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
    
End Sub

Private Sub VerVenRepBONOS()

'    Dim strCtaAux As String, intCont As Integer, n_TotComi As Double, n_ValVcto As Double
'    Dim opeinv As ROpeInv, Sec_ord As Long
'    Dim Sec_tmp As Long, sDescri As String, Sec_Kar As Long
'    Dim adoTit As New Recordset, adoAux As New Recordset
'
'    On Error GoTo Ctrlerror
'
'    '*** RUTINA DE CONTROL DE VENCIMIENTOS DE INSTR. RENTA FIJA ***
'    '*** DE CORTO PLAZO:  OP. REP. CON BONOS                    ***
'    lblMensaje.Caption = "Verificando Vcto. de Operaciones de Reporte con BONOS..."
'
'    '*** Verificar Vencimientos ***
'    With adoComm
'        .CommandText = "SELECT * FROM FMLETRAS WHERE "
'        .CommandText = .CommandText & "COD_FOND=" & "'" & strCodFondo & "' AND "
'        .CommandText = .CommandText & "FCH_VCTO=" & "'" & strFechaSiguiente & "' AND "
'        .CommandText = .CommandText & "COD_FILE ='08' AND CLS_TITU='S' AND "
'        .CommandText = .CommandText & "STA_ORDE='A'"
'        Set adoTit = .Execute
'    End With
'    Do Until adoTit.EOF
'       LGenSec strCodFondo, "ORD": Sec_ord = Sec_ord + 1
'
'       LIniRCabAsiCon CabAsi 'Cabecera del asiento
'       CabAsi.CNT_MOVI = 9
'       'Son nueve movimientos contables para registrar el vencimiento:
'       CabAsi.DSL_COMP = "Vcto." & adoTit!COD_FILE & "-" & adoTit!COD_ANAL & "-" & Trim(adoTit!DSC_LETR)
'       CabAsi.COD_FOND = strCodFondo
'       CabAsi.COD_MONC = "S"
'       CabAsi.COD_MONE = adoTit!COD_MONE
'       CabAsi.FCH_COMP = strFechaSiguiente
'       CabAsi.FCH_CONT = strFechaSiguiente
'       CabAsi.FLG_AUTO = "X"
'       CabAsi.FLG_CONT = "X"
'       CabAsi.GEN_COMP = "GENERADO"
'       CabAsi.GLO_COMP = Trim$(CabAsi.DSL_COMP)
'       CabAsi.HOR_COMP = strHorCie
'       CabAsi.MES_CONT = Mid$(strFechaSiguiente, 5, 2)
'       'Incrementa el correlativo de número de comprobante si es cierre...
'       LGenSec strCodFondo, "COM": lngSecCom = lngSecCom + 1
'       CabAsi.NRO_COMP = Format$(lngSecCom, "00000000")
'       CabAsi.NRO_DOCU = ""
'       CabAsi.NRO_OPER = Format$(lngSecOpe, "00000000")
'       CabAsi.PER_DIGI = ""
'       CabAsi.PER_REVI = ""
'       CabAsi.prd_cont = Left$(strFechaSiguiente, 4)
'       CabAsi.STA_COMP = ""
'       CabAsi.SUB_SIST = "I"
'       CabAsi.TIP_CAMB = Format(CDbl(mhrTipCam.Text), "0.0000")
'       CabAsi.TIP_COMP = ""
'       CabAsi.TIP_DOCU = ""
'       CabAsi.VAL_COMP = 0
'       'Detalle del asiento
'       ReDim DetAsi(1 To CabAsi.CNT_MOVI)
'       For intCont = 1 To CabAsi.CNT_MOVI
'          LIniRDetAsiCon DetAsi(intCont) 'Inicializar
'          DetAsi(intCont).COD_ANAL = adoTit!COD_ANAL 'asignar partes comunes
'          DetAsi(intCont).COD_FILE = adoTit!COD_FILE
'          DetAsi(intCont).COD_FOND = strCodFondo
'          DetAsi(intCont).COD_MONE = adoTit!COD_MONE
'          DetAsi(intCont).FCH_MOVI = strFechaSiguiente
'          DetAsi(intCont).FLG_PROC = "X"
'          DetAsi(intCont).MES_COMP = Mid$(strFechaSiguiente, 5, 2)
'          DetAsi(intCont).NRO_COMP = CabAsi.NRO_COMP
'          DetAsi(intCont).prd_cont = Left$(strFechaSiguiente, 4)
'          DetAsi(intCont).SEC_MOVI = Format$(intCont, "000")
'          DetAsi(intCont).TIP_GENR = "P"
'       Next
'
'       'Obtener cuentas para los movimientos contables
'       adoComm.CommandText = "SELECT * FROM FMTITCTA"
'       adoComm.CommandText = adoComm.CommandText & " WHERE CLS_TITU=" & "'" & adoTit!CLS_TITU & "'"
'       Set adoAux = adoComm.Execute
'       Do Until adoAux.EOF
'          'If adoTit!COD_MONE = "S" Then
'             strCtaAux = adoAux!cod_ctan
'          'Else
'          '   strCtaAux = adoaux!cod_ctax
'          'End If
'          Select Case adoAux!TIP_CTA
'             Case "D" 'Gastos / Costo Enajenacion    '67133
'                DetAsi(1).COD_CTA = strCtaAux
'             Case "A" 'Inversiones                   '1833
'                DetAsi(2).COD_CTA = strCtaAux
'             Case "C" 'R.A.N.R.                      '773332
'                DetAsi(3).COD_CTA = strCtaAux
'             Case "B" 'PROV. DE INTERESES            '189332
'                DetAsi(4).COD_CTA = strCtaAux
'             Case "I" 'Cuentas por cobrar            '12133
'                DetAsi(5).COD_CTA = strCtaAux
'             Case "E" 'Ingresos Operacionales        '77133
'                DetAsi(7).COD_CTA = strCtaAux
'             Case "J" 'Intereses Corridos Adelantados '189331
'                DetAsi(8).COD_CTA = strCtaAux
'             Case "G" 'INTERESES COBRADOS
'                DetAsi(9).COD_CTA = strCtaAux
'          End Select
'          adoAux.MoveNext
'       Loop
'       adoAux.Close: Set adoAux = Nothing
'
'       'Realiza consulta para obtener el Monto a Cobrar y las Comis. al Vcto. del reporte
'       adoComm.CommandText = "SELECT VAL_AGE2,VAL_BOL2,VAL_CNS2,VAL_IGV2,VAL_TOT2 FROM FMOPEINV WHERE COD_FOND='" & adoTit!COD_FOND & "' AND COD_FILE='" & adoTit!COD_FILE & "' AND COD_ANAL='" & adoTit!COD_ANAL & "'"
'       Set adoAux = adoComm.Execute
'       n_TotComi = Format(adoAux!VAL_AGE2 + adoAux!VAL_BOL2 + adoAux!VAL_CNS2 + adoAux!VAL_IGV2, "0.00")
'       n_ValVcto = Format(adoAux!VAL_TOT2, "0.00")  'Monto a Cobrar
'       adoAux.Close: Set adoAux = Nothing
'
'       'a.- Detalle de las Cuentas de Costo Enajenacion y de Inversiones
'       DetAsi(1).FLG_DEHA = "D"
'       DetAsi(1).DSC_MOVI = "Costo Enajenación >>" & adoTit!COD_FILE & "-" & adoTit!COD_ANAL
'       DetAsi(2).FLG_DEHA = "H"
'       DetAsi(2).DSC_MOVI = "Inversiones >> " & adoTit!COD_FILE & "-" & adoTit!COD_ANAL
'
'       'Traer saldo de cuenta de inversiones
'       With adoComm
'        .CommandText = "SELECT SFI_MONN,SFI_MONX,SFI_CONT FROM FMSALDOS"
'        .CommandText = .CommandText & " WHERE COD_CTA = '" & Trim$(DetAsi(2).COD_CTA) & "'"
'        .CommandText = .CommandText & " AND COD_FILE='" & adoTit!COD_FILE & "'"
'        .CommandText = .CommandText & " AND COD_ANAL='" & adoTit!COD_ANAL & "'"
'        .CommandText = .CommandText & " AND COD_FOND='" & strCodFondo & "'"
'        .CommandText = .CommandText & " AND FCH_SALD='" & strFeccie & "'"
'        Set adoAux = .Execute
'       End With
'       If adoAux.EOF Then
'          DetAsi(1).VAL_MOVN = 0
'          DetAsi(1).VAL_MOVX = 0
'          DetAsi(1).VAL_CONT = 0
'       Else
'          If adoTit!COD_MONE = "S" Then
'             DetAsi(1).VAL_MOVN = Format(adoAux!SFI_MONN, "0.00")
'             DetAsi(1).VAL_CONT = Format(adoAux!SFI_CONT, "0.00")
'             DetAsi(1).VAL_MOVX = 0
'          Else
'             DetAsi(1).VAL_MOVX = Format(adoAux!SFI_MONX, "0.00")
'             DetAsi(1).VAL_CONT = Format(adoAux!SFI_CONT, "0.00")
'             DetAsi(1).VAL_MOVN = 0
'          End If
'       End If
'        adoAux.Close: Set adoAux = Nothing
'
'       '*** Detalle de la Cuenta de Inversiones
'       If adoTit!COD_MONE = "S" Then
'          DetAsi(2).VAL_MOVN = Format((DetAsi(1).VAL_MOVN) * -1, "0.00")
'          DetAsi(2).VAL_CONT = Format(DetAsi(2).VAL_MOVN, "0.00")
'          DetAsi(2).VAL_MOVX = 0
'       Else
'          DetAsi(2).VAL_MOVX = Format((DetAsi(1).VAL_MOVX) * -1, "0.00")
'          DetAsi(2).VAL_CONT = Format((DetAsi(1).VAL_CONT) * -1, "0.00")
'          DetAsi(2).VAL_MOVN = 0
'       End If
'
'       '*** Detalle de las Cuentas de Prov.Intereses y de RANR
'       DetAsi(3).FLG_DEHA = "D"
'       DetAsi(3).DSC_MOVI = "RANR >> " & adoTit!COD_FILE & "-" & adoTit!COD_ANAL
'
'       'Traer saldo de cuenta de PROV. de Intereses
'       With adoComm
'        .CommandText = "SELECT SFI_MONN,SFI_MONX,SFI_CONT FROM FMSALDOS "
'        .CommandText = .CommandText & " WHERE COD_CTA = '" & Trim$(DetAsi(4).COD_CTA) & "'"
'        .CommandText = .CommandText & " AND COD_FILE='" & adoTit!COD_FILE & "'"
'        .CommandText = .CommandText & " AND COD_ANAL='" & adoTit!COD_ANAL & "'"
'        .CommandText = .CommandText & " AND COD_FOND='" & strCodFondo & "'"
'        .CommandText = .CommandText & " AND FCH_SALD='" & strFeccie & "'"
'        Set adoAux = .Execute
'       End With
'       If adoAux.EOF Then
'          DetAsi(3).VAL_MOVN = 0
'          DetAsi(3).VAL_MOVX = 0
'          DetAsi(3).VAL_CONT = 0
'       Else
'          If adoTit!COD_MONE = "S" Then
'             DetAsi(3).VAL_MOVN = Format(adoAux!SFI_MONN, "0.00")
'             DetAsi(3).VAL_MOVX = 0
'             DetAsi(3).VAL_CONT = Format(adoAux!SFI_CONT, "0.00")
'          Else
'             DetAsi(3).VAL_MOVN = 0
'             DetAsi(3).VAL_MOVX = Format(adoAux!SFI_MONX, "0.00")
'             DetAsi(3).VAL_CONT = Format(adoAux!SFI_CONT, "0.00")
'          End If
'       End If
'       adoAux.Close: Set adoAux = Nothing
'
'       '* Detalle de la Cuenta de Provisión 192001
'       DetAsi(4).FLG_DEHA = "H"
'       DetAsi(4).DSC_MOVI = "Prov. de Intereses >> " & adoTit!COD_FILE & "-" & adoTit!COD_ANAL
'       If adoTit!COD_MONE = "S" Then
'          DetAsi(4).VAL_MOVN = Format((DetAsi(3).VAL_MOVN) * -1, "0.00")
'          DetAsi(4).VAL_MOVX = 0
'          DetAsi(4).VAL_CONT = Format(DetAsi(4).VAL_MOVN, "0.00")
'       Else
'          DetAsi(4).VAL_MOVX = Format((DetAsi(3).VAL_MOVX) * -1, "0.00")
'          DetAsi(4).VAL_MOVN = 0
'          DetAsi(4).VAL_CONT = Format((DetAsi(3).VAL_CONT) * -1, "0.00")
'       End If
'
'       '* Detalle de la Cuenta Cuentas por Cobrar 164001
'       DetAsi(5).FLG_DEHA = "D"
'       DetAsi(5).DSC_MOVI = "Ctas. por Cobrar >> " & adoTit!COD_FILE & "-" & adoTit!COD_ANAL
'       If adoTit!COD_MONE = "S" Then
'          DetAsi(5).VAL_MOVN = Format(n_ValVcto, "0.00")
'          DetAsi(5).VAL_MOVX = 0
'          DetAsi(5).VAL_CONT = Format(n_ValVcto, "0.00")
'       Else
'          DetAsi(5).VAL_MOVX = Format(n_ValVcto, "0.00")
'          DetAsi(5).VAL_MOVN = 0
'          DetAsi(5).VAL_CONT = Format(n_ValVcto * mhrTipCam.Text, "0.00")
'       End If
'
'       '* Detalle de la Cuenta 671001 Gastos de Gestion: Comisiones + IGV
'       DetAsi(6).COD_CTA = TraeCta("G", "023", adoTit!COD_MONE)
'       DetAsi(6).FLG_DEHA = "D"
'       DetAsi(6).DSC_MOVI = "Gastos Operacionales >> " & adoTit!COD_FILE & "-" & adoTit!COD_ANAL
'       If adoTit!COD_MONE = "S" Then
'          DetAsi(6).VAL_MOVN = Format(n_TotComi, "0.00")
'          DetAsi(6).VAL_CONT = Format(n_TotComi, "0.00")
'          DetAsi(6).VAL_MOVX = 0
'       Else
'          DetAsi(6).VAL_MOVX = Format(n_TotComi, "0.00")
'          DetAsi(6).VAL_CONT = Format(n_TotComi * mhrTipCam.Text, "0.00")
'          DetAsi(6).VAL_MOVN = 0
'       End If
'
'       '* Detalle de la Cuenta de Ingresos por Enajenación 774001
'       DetAsi(7).FLG_DEHA = "H"
'       DetAsi(7).DSC_MOVI = "Ingresos Operacionales >> " & adoTit!COD_FILE & "-" & adoTit!COD_ANAL
'       If adoTit!COD_MONE = "S" Then
'          DetAsi(7).VAL_MOVN = Format(adoTit!VAL_FINA * -1, "0.00")
'          DetAsi(7).VAL_CONT = Format(adoTit!VAL_FINA * -1, "0.00")
'          DetAsi(7).VAL_MOVX = 0
'       Else
'          DetAsi(7).VAL_MOVX = Format(adoTit!VAL_FINA * -1, "0.00")
'          DetAsi(7).VAL_CONT = Format(adoTit!VAL_FINA * mhrTipCam.Text * -1, "0.00")
'          DetAsi(7).VAL_MOVN = 0
'       End If
'
'       'c.- Detalle de la Cuenta de Inter. Corridos
'       DetAsi(8).FLG_DEHA = "H"
'       DetAsi(8).DSC_MOVI = "Intereses Corridos >> " & adoTit!COD_FILE & "-" & adoTit!COD_ANAL
'
'       '*** Traer saldo de cuenta de PROV. de Intereses ***
'       With adoComm
'        .CommandText = "SELECT SFI_MONN,SFI_MONX,SFI_CONT FROM FMSALDOS "
'        .CommandText = .CommandText & " WHERE COD_CTA='" & Trim$(DetAsi(8).COD_CTA) & "'"
'        .CommandText = .CommandText & " AND COD_FILE='" & adoTit!COD_FILE & "'"
'        .CommandText = .CommandText & " AND COD_ANAL='" & adoTit!COD_ANAL & "'"
'        .CommandText = .CommandText & " AND COD_FOND='" & strCodFondo & "'"
'        .CommandText = .CommandText & " AND FCH_SALD='" & strFeccie & "'"
'        Set adoAux = .Execute
'       End With
'       If adoAux.EOF Then
'          DetAsi(8).VAL_MOVN = 0
'          DetAsi(8).VAL_MOVX = 0
'          DetAsi(8).VAL_CONT = 0
'       Else
'          If adoTit!COD_MONE = "S" Then
'             DetAsi(8).VAL_MOVN = Format(adoAux!SFI_MONN * -1, "0.00")
'             DetAsi(8).VAL_MOVX = 0
'             DetAsi(8).VAL_CONT = Format(adoAux!SFI_CONT * -1, "0.00")
'          Else
'             DetAsi(8).VAL_MOVN = 0
'             DetAsi(8).VAL_MOVX = Format(adoAux!SFI_MONX * -1, "0.00")
'             DetAsi(8).VAL_CONT = Format(adoAux!SFI_CONT * -1, "0.00")
'          End If
'       End If
'       adoAux.Close: Set adoAux = Nothing
'
'       '* Detalle de la Cuenta de Intereses Cobrados 772001
'       DetAsi(9).FLG_DEHA = "H"
'       DetAsi(9).DSC_MOVI = "Intereses Cobrados >> " & adoTit!COD_FILE & "-" & adoTit!COD_ANAL
'       If adoTit!COD_MONE = "S" Then
'          DetAsi(9).VAL_MOVN = Format((DetAsi(5).VAL_MOVN + DetAsi(6).VAL_MOVN + DetAsi(7).VAL_MOVN + DetAsi(8).VAL_MOVN) * -1, "0.00")
'          DetAsi(9).VAL_CONT = Format((DetAsi(5).VAL_CONT + DetAsi(6).VAL_CONT + DetAsi(7).VAL_CONT + DetAsi(8).VAL_CONT) * -1, "0.00")
'          DetAsi(9).VAL_MOVX = 0
'       Else
'          DetAsi(9).VAL_MOVX = Format((DetAsi(5).VAL_MOVX + DetAsi(6).VAL_MOVX + DetAsi(7).VAL_MOVX + DetAsi(8).VAL_MOVX) * -1, "0.00")
'          DetAsi(9).VAL_CONT = Format((DetAsi(5).VAL_CONT + DetAsi(6).VAL_CONT + DetAsi(7).VAL_CONT + DetAsi(8).VAL_CONT) * -1, "0.00")
'          DetAsi(9).VAL_MOVN = 0
'       End If
'
'       '* Crear movimiento para liquidación
'       LIniMovCta MovCta
'       MovCta.COD_ANAL = adoTit!COD_ANAL
'       MovCta.COD_FILE = adoTit!COD_FILE
'       MovCta.COD_FOND = strCodFondo
'       MovCta.COD_MONE = adoTit!COD_MONE
'       MovCta.com_orig = Trim$(CabAsi.DSL_COMP)
'       MovCta.FCH_CREA = strFeccie
'       'MACR 14/04/1999
'       MovCta.FCH_OBLI = strFechaSiguiente
'       MovCta.NRO_OPER = Format$(lngSecOpe, "00000000")
'       LGenSec strCodFondo, "TMP": Sec_tmp = Sec_tmp + 1: MovCta.NRO_MCTA = Format$(Sec_tmp, "00000000")
'       MovCta.SUB_SIST = "A"
'       MovCta.TIP_MOVI = "E"
'       MovCta.TIP_OPER = "28"
'
'       If adoTit!COD_MONE = "S" Then
'          MovCta.VAL_MOVI = Format(DetAsi(5).VAL_MOVN, "0.00")
'       Else
'          MovCta.VAL_MOVI = Format(DetAsi(5).VAL_MOVX, "0.00")
'       End If
'       MovCta.NRO_COMP = Format$(lngSecCom, "00000000")
'       MovCta.COD_CTA = gstrBancDefa
'       MovCta.Cod_part = ""
'       LGraMovCta MovCta
'
'       'Detalle movimiento de temporales
'       LIniMovTmp Movtmp
'       Movtmp.COD_ANAL = adoTit!COD_ANAL
'       Movtmp.COD_CTA = DetAsi(5).COD_CTA ' Cuentas por Cobrar
'       Movtmp.COD_FILE = adoTit!COD_FILE
'       Movtmp.COD_FOND = strCodFondo
'       Movtmp.COD_MONE = adoTit!COD_MONE
'       Movtmp.DSC_MOVI = CabAsi.DSL_COMP
'       Movtmp.FLG_DEHA = "H"
'       Movtmp.NRO_MCTA = MovCta.NRO_MCTA
'       Movtmp.nro_temp = "00000001"
'       Movtmp.SEC_MOVI = "001"
'       Movtmp.SUB_SIST = "A"
'       Movtmp.VAL_MOVI = Format(MovCta.VAL_MOVI * -1, "0.00")
'       LGraMovTmp Movtmp
'
'       LGraAsiCont CabAsi, DetAsi() 'Grabar el asiento
'
'       '*** Actualizar Valores de Renta Fija Vencidos ***
'       With adoComm
'        .CommandText = "UPDATE FMLETRAS SET STA_ORDE='R'"
'        .CommandText = .CommandText & " WHERE COD_FOND='" & strCodFondo & "'"
'        .CommandText = .CommandText & " AND CLS_TITU='" & adoTit!CLS_TITU & "'"
'        .CommandText = .CommandText & " AND COD_FILE='" & adoTit!COD_FILE & "'"
'        .CommandText = .CommandText & " AND COD_ANAL='" & adoTit!COD_ANAL & "'"
'        adoConn.Execute .CommandText
'       End With
'
'       'Actualizar Operaciones vencidas
'       With adoComm
'        .CommandText = "UPDATE FMOPEINV SET STA_INVE='X'"
'        .CommandText = .CommandText & ",FCH_VCTO='" & strFechaSiguiente & "'"
'        .CommandText = .CommandText & " WHERE COD_FOND='" & strCodFondo & "'"
'        .CommandText = .CommandText & " AND COD_FILE='" & adoTit!COD_FILE & "'"
'        .CommandText = .CommandText & " AND COD_ANAL='" & adoTit!COD_ANAL & "'"
'        .CommandText = .CommandText & " AND NRO_OPER=" & CVar(adoTit!NRO_OPE1)
'        adoConn.Execute .CommandText
'       End With
'
'       '*** Actualizar la Operacion de Inversion ***
'       LIniOpeInv opeinv
'       opeinv.COD_FOND = strCodFondo
'       opeinv.prd_cont = Left$(strFechaSiguiente, 4)
'       opeinv.MES_CONT = Mid$(strFechaSiguiente, 5, 2)
'       opeinv.NRO_OPER = lngSecOpe
'       opeinv.STA_INVE = "X"
'       opeinv.NRO_ORDE = Sec_ord
'       opeinv.CLS_TITU = adoTit!CLS_TITU
'       opeinv.COD_TITU = adoTit!COD_FILE & adoTit!COD_ANAL
'       opeinv.COD_LETR = adoTit!COD_LETR
'       opeinv.COD_FILE = adoTit!COD_FILE
'       opeinv.COD_ANAL = adoTit!COD_ANAL
'       opeinv.TIP_OPER = "V" 'VENCIMIENTO:VENTA
'       opeinv.DSL_OPER = "Vcto. " & adoTit!COD_FILE & "-" & adoTit!COD_ANAL & "-" & Trim$(adoTit!DSC_LETR)
'       opeinv.COD_EMPR = adoTit!COD_EMPR
'       opeinv.COD_AGEN = adoTit!COD_SAB
'       opeinv.ACC_INVE = "S"
'       opeinv.TIP_DOCU = "O"
'       opeinv.FCH_OPER = strFechaSiguiente
'       opeinv.FCH_CONF = strFechaSiguiente
'       opeinv.FCH_LIQU = strFechaSiguiente
'       opeinv.FCH_VCTO = strFechaSiguiente
'       opeinv.COD_MONE = adoTit!COD_MONE
'       opeinv.CNT_INVE = 1
'       opeinv.VAL_TCMB = Format(mhrTipCam, "0.0000")
'       opeinv.VAL_UNIT = Format(adoTit!VAL_FINA, "0.00")
'       opeinv.VAL_IADE = Format(adoTit!VAL_INT2, "0.00")
'       opeinv.VAL_AGEN = Format(adoTit!TAS_AGE2, "0.00")
'       opeinv.VAL_AGE2 = Format(0, "0.00")
'       opeinv.VAL_CNSV = Format(adoTit!TAS_CNS2, "0.00")
'       opeinv.VAL_CNS2 = Format(0, "0.00")
'       opeinv.VAL_BOLS = Format(adoTit!TAS_BOL2, "0.00")
'       opeinv.VAL_BOL2 = Format(0, "0.00")
'       opeinv.VAL_IGV = Format(0, "0.00")
'       opeinv.VAL_IGV2 = Format(0, "0.00")
'       opeinv.VAL_TOT2 = Format(0, "0.00")
'       opeinv.VAL_VCTO = Format(0, "0.00")
'       opeinv.FLG_CONT = "X"
'       opeinv.NRO_COMP = Format$(lngSecCom, "00000000")
'       opeinv.NRO_DOCU = ""
'       opeinv.USU_ACTU = Trim$(gstrLogin)
'       opeinv.FCH_ACTU = strFechaSiguiente
'       opeinv.HOR_ACTU = Format(Time$, "hh:mm")
'       If adoTit!COD_MONE = "D" Then
'          opeinv.VAL_TOTA = DetAsi(5).VAL_MOVX
'       Else
'          opeinv.VAL_TOTA = DetAsi(5).VAL_CONT
'       End If
'
'       '*** REGISTRO DE GRUPO ECONOMICO ***
'       adoComm.CommandText = "SELECT COD_GRUP FROM FMPERSON WHERE COD_PERS='" & adoTit!COD_EMPR & "'"
'       Set adoAux = adoComm.Execute
'       opeinv.COD_GRUP = adoAux!COD_GRUP
'       adoAux.Close: Set adoAux = Nothing
'
'       LGraOpeInv opeinv 'Inserta el Registro en FMOPEINV
'
'       '*** Actualizar kardex de Inventario ***
'       With adoComm
'        .CommandText = "SELECT * FROM FMKARDEX"
'        .CommandText = .CommandText & " WHERE COD_FOND='" & strCodFondo & "'"
'        .CommandText = .CommandText & " AND COD_FILE='" & adoTit!COD_FILE & "'"
'        .CommandText = .CommandText & " AND COD_ANAL='" & adoTit!COD_ANAL & "'"
'        .CommandText = .CommandText & " ORDER BY NRO_OPER DESC"
'        Set adoAux = .Execute
'       End With
'       If adoAux.EOF Then
'          adoAux.Close: Set adoAux = Nothing 'No existe kardex (No se muestra mensaje por no parar el proceso)
'       Else
'          If adoAux!SLD_FINA < 1 Then
'             adoAux.Close: Set adoAux = Nothing 'No existe stock (No se muestra mensaje por no parar el proceso)
'          Else
'             'completar valores del kardex
'             LIniKar Kardex 'inicializa
'             Kardex.COD_FOND = strCodFondo
'             Kardex.prd_cont = Left$(strFechaSiguiente, 4)
'             Kardex.MES_CONT = Mid$(strFechaSiguiente, 5, 2)
'             Kardex.NRO_OPER = Format$(lngSecOpe, "00000000")
'             Kardex.COD_FILE = adoTit!COD_FILE
'             Kardex.COD_ANAL = adoTit!COD_ANAL
'             Kardex.COD_TITU = adoTit!COD_FILE & adoTit!COD_ANAL
'             Kardex.FCH_MOVI = strFechaSiguiente
'             Kardex.TIP_MOVI = "S"
'             Kardex.TIP_ORIG = "O"
'             Kardex.CNT_MOVI = 1
'             Kardex.COD_MONE = adoTit!COD_MONE
'             Kardex.FCH_CORT = strFechaSiguiente
'             Kardex.CNT_CORT = 0
'             Kardex.FLG_ULTI = "X"
'             Kardex.FCH_OPER = strFechaSiguiente
'             Kardex.VAL_UNIT = adoAux!VAL_PROM
'             Kardex.VAL_MOVI = adoAux!VAL_PROM * -1
'             Kardex.VAL_COMI = 0
'
'             For intCont = 6 To 6 'Comisiones e igv correspondiente
'                If adoTit!COD_MONE = "S" Then
'                   Kardex.VAL_COMI = Kardex.VAL_COMI + DetAsi(intCont).VAL_MOVN
'                Else
'                   Kardex.VAL_COMI = Kardex.VAL_COMI + DetAsi(intCont).VAL_MOVX
'                End If
'             Next
'
'             Kardex.COM_MOVI = CabAsi.DSL_COMP
'             Kardex.CLS_TITU = adoTit!CLS_TITU
'             LGenSec strCodFondo, "KAR": Sec_Kar = Sec_Kar + 1  'NUEVO
'             Kardex.NRO_KARD = Sec_Kar
'             Kardex.COD_EMPR = adoTit!COD_EMPR
'             Kardex.TIP_OBLI = "O"
'             Kardex.SLD_INIC = adoAux!SLD_FINA
'             Kardex.SLD_FINA = adoAux!SLD_FINA - 1   'se asume salida de 1 reporte
'             Kardex.VAL_SALD = adoAux!VAL_SALD + Kardex.VAL_MOVI + Kardex.VAL_COMI
'
'             If Kardex.SLD_FINA <> 0 Then
'                Kardex.VAL_PROM = Format(Kardex.VAL_SALD / Kardex.SLD_FINA, "0.00")
'             Else
'                Kardex.VAL_PROM = 0
'             End If
'             adoAux.Close: Set adoAux = Nothing
'
'             'Actualiza flag de ultimo movimiento en regitros actuales
'             With adoComm
'                .CommandText = "UPDATE FMKARDEX SET FLG_ULTI='' "
'                .CommandText = .CommandText & " WHERE COD_FOND='" & strCodFondo & "'"
'                .CommandText = .CommandText & " AND COD_FILE='" & adoTit!COD_FILE & "'"
'                .CommandText = .CommandText & " AND COD_ANAL='" & adoTit!COD_ANAL & "'"
'                adoConn.Execute .CommandText
'             End With
'             'inserta movimiento en el kardex
'             LGraKar Kardex
'          End If
'       End If
'       'Siguiente letra
'       adoTit.MoveNext
'    Loop
'    adoTit.Close: Set adoTit = Nothing
'
'    Exit Sub
'
'Ctrlerror:
'    Resume Next
  
End Sub


Private Sub cboFondo_Click()

    Dim adoRegistro     As ADODB.Recordset, adoFondo As ADODB.Recordset
    Dim intRespuesta    As Integer
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
        
            dtpFechaCierre.Value = adoRegistro("FechaCuota")
            dtpFechaCierreHasta.Value = adoRegistro("FechaCuota")
            dtpFechaCierreHasta.Enabled = False
            datFechaCierre = dtpFechaCierre.Value
            strTipoFondo = adoRegistro("TipoFondo")
            
            gdatFechaActual = adoRegistro("FechaCuota")
            
            If Me.Tag = "R" Then
                
                '*** Obtener Fecha Hábil antes del Reproceso ***
                adoComm.CommandText = "SELECT NumReproceso,FechaHabil FROM FondoReproceso WHERE CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND Estado='" & Estado_Activo & "'"
                Set adoFondo = adoComm.Execute
                
                If Not adoFondo.EOF Then
                    intNumReproceso = adoFondo("NumReproceso")
                    dtpFechaCierreHasta.Value = adoFondo("FechaHabil")
                    dtpFechaCierreHasta.Enabled = True
                End If
                adoFondo.Close: Set adoFondo = Nothing
                   
            End If
            
            Call ActualizarFechasCierre(dtpFechaCierre.Value)
    
            Call ValidarFechas
            
            strCodMoneda = adoRegistro("CodMoneda")
            lblValorAIR(0).Caption = CStr(adoRegistro("ValorCuotaInicial"))
            lblValorDIR(0).Caption = "0"
            lblRentabilidad(1).Caption = "0"
            
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            cmdCierre.Enabled = True
        
            'ACTUALIZA PARAMETROS GLOBALES POR FONDO
            If Not CargarParametrosGlobales(strCodFondo) Then Exit Sub

        Else
            cmdCierre.Enabled = False
            MsgBox "Periodo contable no vigente para este fondo! Debe aperturar primero un periodo contable para este fondo!", vbExclamation + vbOKOnly, Me.Caption
            Exit Sub
        End If
        adoRegistro.Close
        
'        .CommandText = " SELECT ValorCuotaInicial,ValorCuotaFinal FROM FondoValorCuota " & _
'            "WHERE (FechaCuota >='" & strFechaAnterior & "' AND FechaCuota <'" & strFechaCierre & "') AND " & _
'            "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "'"
'        Set adoRegistro = .Execute
'
'        If Not adoRegistro.EOF Then
'            lblValorAIR(1).Caption = CStr(adoRegistro("ValorCuotaInicial"))
'            lblValorDIR(1).Caption = CStr(adoRegistro("ValorCuotaFinal"))
'
'            If CDbl(lblValorAIR(0).Caption) > 0 Then
'                lblRentabilidad(0).Caption = CStr((((CDbl(lblValorDIR(0).Caption) / CDbl(lblValorAIR(0).Caption)) ^ 365) - 1) * 100)
'            End If
'        Else
'            lblValorAIR(1).Caption = "0"
'            lblValorDIR(1).Caption = "0"
'        End If
'        adoRegistro.Close: Set adoRegistro = Nothing
'        lblDescrip(7).Caption = Convertddmmyyyy(strFechaAnterior)
                
        Call BuscarTipoCambio
        
        Call BuscarFondoSeries(Codigo_Cierre_Definitivo)
            
        '-----------------------------------
         Set adoFondoTipo = New ADODB.Recordset
         
         With adoComm
    
            .CommandText = " SELECT FrecuenciaValorizacion FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "'"
            Set adoFondoTipo = .Execute
            
            strTipoFondoFrecuencia = adoFondoTipo("FrecuenciaValorizacion")
        
        End With
        '--------------------------------
        
        frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        
        '*** Periodo Actual ***
        strSQL = "{ call up_CNSelPeriodoContableVigente ('" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        CargarControlLista strSQL, cboPeriodoContable, arrPeriodo(), ""
    
        If cboPeriodoContable.ListCount > 0 Then cboPeriodoContable.ListIndex = 0
        
    End With
    
End Sub

Private Sub ValidarFechas()
    
    If EsDiaUtil(dtpFechaEntrega.Value) Then
      dtpFechaEntrega.Value = dtpFechaEntrega
    Else
      dtpFechaEntrega.Value = ProximoDiaUtil(dtpFechaEntrega)
    End If
    
End Sub
Private Function TodoOK() As Boolean
                
    Dim adoConsulta As ADODB.Recordset
    Dim strMensaje  As String
    
    TodoOK = False
                
    If cboFondo.ListCount = 0 Then
        MsgBox "No existen fondos definidos...", vbCritical, Me.Caption
        Exit Function
    End If
    
    '*** Verificar si existen operaciones con retención para liberar ***
    If Not VerificarOperacionRetencion(strCodFondo, strFechaCierre, strFechaSiguiente) Then
        MsgBox "Existen Operaciones con Retención para Liberar...", vbCritical, Me.Caption
        Exit Function
    End If
        
    '*** Verificar si existen ordenes de inversión pendientes de confirmación ***
    If Not VerificarOrdenInversion(strCodFondo, strFechaCierre, strFechaSiguiente) Then
        MsgBox "Existen ordenes de inversión pendientes de ser confirmadas...", vbCritical, Me.Caption
        Exit Function
    End If
        
    '*** Verificar si existe nuevo periodo contable ***
    If Not VerificarPeriodoContable(strCodFondo, strFechaCierre, strFechaSiguiente) Then
        MsgBox "Por favor genere el nuevo periodo contable...", vbCritical, Me.Caption
        Exit Function
    End If
                                
    Set adoConsulta = New ADODB.Recordset
    '*** Se Realizó Cierre anteriormente ? ***
    adoComm.CommandText = "{ call up_GNValidaCierreRealizado('" & _
        strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaCierre & "','" & _
        strFechaSiguiente & "') }"
    Set adoConsulta = adoComm.Execute
    
    If Not adoConsulta.EOF Then
        If adoConsulta("IndCierre") = Valor_Indicador Then
            MsgBox "El Cierre Diario del Día " & CStr(dtpFechaCierre.Value) & " ya fué realizado antes.", vbCritical, Me.Caption
            adoConsulta.Close: Set adoConsulta = Nothing
            Exit Function
        End If
    End If
    adoConsulta.Close

    '*** Cerrar periodo ? ***
    adoComm.CommandText = "SELECT IndCierre FROM PeriodoContable " & _
        "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
        "MesContable='99' AND FechaFinal='" & strFechaAnterior & "'"
    Set adoConsulta = adoComm.Execute
    
    If Not adoConsulta.EOF Then
        If adoConsulta("IndCierre") = Valor_Caracter Then
            MsgBox "Por favor procese el Cierre Anual.", vbCritical, Me.Caption
            adoConsulta.Close: Set adoConsulta = Nothing
            Exit Function
        End If
    End If
    adoConsulta.Close

    '*** Existe un descuadre en la contabilidad ? ***
    adoComm.CommandText = "{ call up_ACValidaCuadreContabilidad('" & _
        strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaCierre & "','" & _
        strFechaSiguiente & "') }"
    Set adoConsulta = adoComm.Execute
    
    If Not adoConsulta.EOF Then
        If Not IsNull(adoConsulta("MontoContable")) Then
            If CCur(adoConsulta("MontoContable")) <> 0 Then
                If MsgBox("Existe un descuadre contable de " & CStr(adoConsulta("MontoContable")) & " para el día " & CStr(dtpFechaCierre.Value) & "." & vbNewLine & vbNewLine & "Desea Continuar ?.", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                    adoConsulta.Close: Set adoConsulta = Nothing
                    Exit Function
                End If
            End If
        End If
    End If
    adoConsulta.Close

    '*** Cierre en fecha aún no abierta para el Fondo ***
    adoComm.CommandText = "{ call up_GNValidaFechaNoAbierta('" & _
        strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaCierre & "','" & _
        strFechaSiguiente & "') }"
    Set adoConsulta = adoComm.Execute
    
    If Not adoConsulta.EOF Then
        If Trim(adoConsulta("IndAbierto")) = Valor_Caracter Then
            MsgBox "El Día " & CStr(dtpFechaCierre.Value) & " aún no ha sido abierto.", vbCritical, Me.Caption
            adoConsulta.Close: Set adoConsulta = Nothing
            Exit Function
        End If
    End If
    adoConsulta.Close
        
    '"LEFT JOIN InstrumentoPrecioTir IPT ON(IPT.CodTitulo=IK.CodTitulo AND IPT.IndUltimoPrecio='X') " & _

        
    '*** Verificar si existen valores sin precio o tir de mercado en cartera ***
    adoComm.CommandText = "SELECT II.Nemotecnico," & _
        "ISNULL(dbo.uf_IVObtenerUltimoPrecioTitulo(IK.CodTitulo, '" & strFechaCierre & "'),0) AS PrecioCierre, " & _
        "ISNULL(dbo.uf_IVObtenerUltimaTirTitulo(IK.CodTitulo, '" & strFechaCierre & "'),0) AS TirCierre " & _
        "FROM InversionKardex IK " & _
        "JOIN InstrumentoInversion II ON(II.CodTitulo=IK.CodTitulo) " & _
        "JOIN InversionFile IVF ON(IVF.CodFile=IK.CodFile AND (IndTir='X' OR IndPrecio='X')) " & _
        "WHERE IK.CodFondo = '" & strCodFondo & "' AND IK.CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
        "IK.SaldoFinal > 0 AND " & _
        "IK.NumKardex = dbo.uf_IVObtenerUltimoMovimientoKardexValor('" & strCodFondo & "','" & gstrCodAdministradora & "', IK.CodTitulo, '" & strFechaCierre & "') " & _
        "ORDER BY II.Nemotecnico"
    Set adoConsulta = adoComm.Execute
    
    strMensaje = Valor_Caracter
    Do While Not adoConsulta.EOF
        If adoConsulta("PrecioCierre") = 0 And adoConsulta("TirCierre") = 0 Then
            strMensaje = strMensaje & adoConsulta("Nemotecnico") & vbNewLine
        End If
        
        adoConsulta.MoveNext
    Loop
    adoConsulta.Close: Set adoConsulta = Nothing
    
    If strMensaje <> Valor_Caracter Then
        strMensaje = "Existen los siguientes valores sin Precio o Tir de Mercado :" & vbNewLine & vbNewLine & strMensaje
        If MsgBox(strMensaje & vbNewLine & vbNewLine & "Desea Continuar ?.", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            Exit Function
        End If
    End If
    

    '*** Si todo paso OK ***
    TodoOK = True
  
End Function


Private Sub cboPeriodoContable_Click()
        
      If strTipoFondoFrecuencia = strFrecuenciaValorizacionMensual Then
      
            Dim adoRegistro As ADODB.Recordset
            
            strPeriodoContable = Valor_Caracter
            If cboPeriodoContable.ListIndex < 0 Then Exit Sub
        
            strPeriodoContable = Mid(arrPeriodo(cboPeriodoContable.ListIndex), 1, 4)
            
            strMesContable = Mid(arrPeriodo(cboPeriodoContable.ListIndex), 5, 2)
            
            Set adoRegistro = New ADODB.Recordset
            
            '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
            adoComm.CommandText = "{ call up_ACSelDatosParametro(73,'" & strCodFondo & "','" & gstrCodAdministradora & "','" & strPeriodoContable & "') }"
            Set adoRegistro = adoComm.Execute
                
            If Not adoRegistro.EOF Then
                dtpFechaCierre.Value = adoRegistro("FechaInicio")
                dtpFechaCierreHasta.Value = adoRegistro("FechaFinal")
                strFechaCierreDesde = Convertyyyymmdd(dtpFechaCierre.Value)
                strFechaCierreHasta = Convertyyyymmdd(dtpFechaCierreHasta.Value)
                strFechaCierreHastaSiguiente = Convertyyyymmdd(DateAdd("d", 1, dtpFechaCierreHasta.Value))
                datFechaCierre = dtpFechaCierre.Value
            End If
            
            adoRegistro.Close: Set adoRegistro = Nothing
        
            Call ActualizarFechasCierre(dtpFechaCierre.Value)
            
            Call BuscarTipoCambio
            
    End If
    
End Sub

Private Sub cmdCierre_Click()
      Dim fech As Date
      Dim fechs As Date
      Dim F As String
      Dim fs As String
      
      Dim i As Integer
        
        If strTipoFondoFrecuencia = strFrecuenciaValorizacionMensual Then

            If TodoOKMensual = True Then
                Call CierreMensual
            End If

        Else

            If TodoOK = True Then
                Call CierreDiario
            End If
            
        
        End If
    
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub


Private Sub dtpFechaCierre_Change()

    strFechaCierre = Convertyyyymmdd(dtpFechaCierre.Value)
    gstrPeriodoActual = Format(Year(dtpFechaCierre.Value), "0000")
    gstrMesActual = Format(Month(dtpFechaCierre.Value), "00")
    gstrDiaActual = Format(Day(dtpFechaCierre.Value), "00")
    strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, dtpFechaCierre.Value))
    strFechaAnterior = Convertyyyymmdd(DateAdd("d", -1, dtpFechaCierre.Value))
    strFechaAnteAnterior = Convertyyyymmdd(DateAdd("d", -2, dtpFechaCierre.Value))
    strFechaSubSiguiente = Convertyyyymmdd(DateAdd("d", 2, dtpFechaCierre.Value))
            
End Sub


Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Load()
  
    Call InicializarValores
    Call CargarListas
    Call DarFormato
    Call Ocultar_Recursos
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
        
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
        
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
            
End Sub
Private Sub CargarListas()
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
            
End Sub
Private Sub InicializarValores()

    dtpFechaCierre.Value = gdatFechaActual
    dtpFechaEntrega.Value = DateAdd("d", gintDiasPagoRescate, dtpFechaCierre.Value)
    
    strFrecuenciaValorizacionMensual = "05"
    
    Call ValidarFechas
    'txtTipoCambio.Text = "0"
   
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgTipoCambioCierre.Columns(1).Width = tdgTipoCambioCierre.Width * 0.01 * 35
    
    
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set frmCierreDiario = Nothing
    
End Sub

Private Sub GenRedVDesc()
   
End Sub

Private Sub GenOperaciones_SuscripcionValorDesconocido()
   
    Dim adoRegistro             As ADODB.Recordset, adoTemporal     As ADODB.Recordset
    Dim adoAuxiliar             As ADODB.Recordset
    Dim curMonto                As Currency, curMontoMN             As Currency
    Dim curMontoME              As Currency, curMontoContable       As Currency
    Dim curMontoSobreLaPar      As Currency, curMontoBajoLaPar      As Currency
    Dim dblCantCuota            As Double, curValorNominal          As Currency
    Dim dblCantCuotasInicial    As Double, dblCantCuotaReal         As Currency
    Dim strNumOperacion         As String, strCodParticipe          As String
    Dim strCodMonedaOperacion   As String, strCodClasePersoneria    As String
    Dim strIndExtranjero        As String, strCuentaAporte          As String
    Dim strCuentaSobreLaPar     As String, strCuentaBajoLaPar       As String
    Dim strCuentaSuscripcion    As String, strCuentaComision        As String
    Dim strNumAsiento           As String, strDescripAsiento        As String
    Dim strDescripMovimiento    As String, strIndDebeHaber          As String
    Dim strCodCuenta            As String, strCodFile               As String
    Dim strCodAnalitica         As String, strFechaGrabar           As String
    Dim intFor                  As Integer
    Dim dblTipoCambioCierre     As Double
    
    frmMainMdi.stbMdi.Panels(3).Text = "Verificando Suscripciones a Valor Desconocido..."
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        '** Procesar operaciones de suscripción a valor desconocido  y conocido T ***
        .CommandText = "SELECT MF.*, PO.* " & _
            "FROM  MovimientoFondo MF JOIN ParticipeOperacion PO ON(PO.CodFondo=MF.CodFondo AND PO.CodAdministradora=MF.CodAdministradora AND PO.NumOperacion=MF.NumOperacion) " & _
            "WHERE (MF.FechaObligacion >='" & strFechaCierre & "' AND MF.FechaObligacion <'" & strFechaSiguiente & "') AND " & _
            "MF.CodFondo='" & strCodFondo & "' AND MF.CodAdministradora='" & gstrCodAdministradora & "' AND " & _
            "MF.ModuloOrigen='P' AND MF.TipoOperacion='" & Codigo_Caja_Suscripcion & "' AND " & _
            "MF.IndContable = '' AND EstadoOrden='" & Estado_Caja_Confirmado & "' AND PO.IndRetencion <> 'C' AND PO.IndRetencion <> 'X'"
        Set adoRegistro = .Execute
    
        Do While Not adoRegistro.EOF
            strNumOperacion = adoRegistro("NumOperacion")
            strCodParticipe = adoRegistro("CodParticipe")
            strCodMonedaOperacion = adoRegistro("CodMoneda")
            strCodClasePersoneria = adoRegistro("ClasePersona")
            strIndExtranjero = adoRegistro("IndExtranjero")
            strCodFile = "000"
            strCodAnalitica = "00000000"
        
            '*** Obtiene el número de cuotas a valor desconocido ***
            curMonto = CCur(Abs(adoRegistro("MontoOrdenCobroPago"))) - (CCur(adoRegistro("MontoComision")) + CCur(adoRegistro("MontoIgv")))
            dblCantCuota = Round(curMonto / dblValNuevaCuota, Decimales_CantCuota)
            dblCantCuotaReal = Round(curMonto / dblValNuevaCuotaReal, Decimales_CantCuota)

            '*** Obtiene los montos del aporte y Capital adicional ***
            curValorNominal = dblCantCuota * dblValorCuotaNominal
            If (dblValNuevaCuota - dblValorCuotaNominal) > 0 Then
                curMontoSobreLaPar = Round((dblValNuevaCuota - dblValorCuotaNominal) * dblCantCuota, Decimales_Monto)
                curMontoBajoLaPar = 0
            Else
                curMontoBajoLaPar = Round((dblValNuevaCuota - dblValorCuotaNominal) * dblCantCuota, Decimales_Monto)
                curMontoSobreLaPar = 0
            End If
            
            '*** Obtener tipo de cambio ***
            'dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
            dblTipoCambioCierre = ObtenerValorTipoCambio(adoRegistro("CodMoneda"), Codigo_Moneda_Local, strFechaCierre, strFechaCierre, gstrCodClaseTipoCambioOperacionFondo, Codigo_Valor_TipoCambioCompra)
                        
                        
            '*** Actualizar el valor cuota y la cantidad de cuotas en la operación ***
            .CommandText = "UPDATE ParticipeOperacion SET CantCuotas=" & dblCantCuota & ",ValorCuota=" & dblValNuevaCuota & "," & _
                "MontoTotal=" & CDec(adoRegistro("MontoOrdenCobroPago")) & ",EstadoOperacion='" & Estado_Activo & "' " & _
                "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND NumOperacion='" & strNumOperacion & "' AND " & _
                "(FechaConfirmacion >= '" & strFechaCierre & "' AND FechaConfirmacion < '" & strFechaSiguiente & "')"
            adoConn.Execute .CommandText
            
            '*** Obtener el Saldo Inicial de Cuotas ***
            .CommandText = "SELECT SUM(CantCuotas) TotalCuotas FROM ParticipeCertificado " & _
                "WHERE FechaSuscripcion <'" & strFechaSiguiente & "' AND IndVigente='X' AND CodParticipe='" & strCodParticipe & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            Set adoTemporal = .Execute
            
            If Not adoTemporal.EOF Then
                If Not IsNull(adoTemporal("TotalCuotas")) Then
                   dblCantCuotasInicial = adoTemporal("TotalCuotas")
                Else
                   dblCantCuotasInicial = 0
                End If
            End If
            adoTemporal.Close: Set adoTemporal = Nothing
                        
            '*** Actualiza el detalle de la operación ***
            .CommandText = "UPDATE ParticipeOperacionDetalle SET " & _
                "CantCuotas=" & dblCantCuota & ",ValorCuota=" & dblValNuevaCuota & ",MontoTotal=" & CDec(Abs(adoRegistro("MontoOrdenCobroPago"))) & "," & _
                "MontoAporte=" & curValorNominal & ",MontoSobreLaPar=" & curMontoSobreLaPar & "," & _
                "MontoBajoLaPar=" & curMontoBajoLaPar & ",SaldoInicialCuotas=" & dblCantCuotasInicial & " " & _
                "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "NumOperacion='" & strNumOperacion & "'"
            adoConn.Execute .CommandText
            
            '*** Actualiza los certificados ***
            .CommandText = "UPDATE ParticipeCertificado SET " & _
                "CantCuotas=" & dblCantCuota & ",ValorCuota=" & dblValNuevaCuota & ",IndContable='X',IndVigente='X'," & _
                "CantCuotasPagadas=" & dblCantCuota & ",FechaSuscripcion='" & strFechaCierre & "' " & _
                "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "NumOperacion='" & strNumOperacion & "'"
            adoConn.Execute .CommandText
                                    
            '*** Actualiza Pagos Suscripción ***
            .CommandText = "UPDATE ParticipePagoSuscripcion SET " & _
                "CantCuotasPagadas=" & dblCantCuota & ",CantCuotasPagadasReal=" & dblCantCuotaReal & " " & _
                "WHERE CodParticipe='" & strCodParticipe & "' AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "(FechaPago >='" & strFechaCierre & "' AND FechaPago <'" & strFechaSiguiente & "')"
            adoConn.Execute .CommandText
            
            '*** Actualiza el número de cuotas en el Kardex ***
            .CommandText = "UPDATE FondoValorCuota SET " & _
                "CantCuotaSuscripcionDesconocida = CantCuotaSuscripcionDesconocida + " & dblCantCuota & "," & _
                "CantCuotaFinal = CantCuotaFinal + " & dblCantCuota & "," & _
                "CantCuotaSuscripcionPagada = CantCuotaSuscripcionPagada + " & dblCantCuota & "," & _
                "CantCuotaFinalPagada = CantCuotaFinalPagada + " & dblCantCuota & "," & _
                "CantCuotaSuscripcionPagadaReal = CantCuotaSuscripcionPagadaReal + " & dblCantCuotaReal & "," & _
                "CantCuotaFinalPagadaReal = CantCuotaFinalPagadaReal + " & dblCantCuotaReal & " " & _
                "WHERE (FechaCuota >='" & strFechaCierre & "' AND FechaCuota <'" & strFechaSiguiente & "') AND " & _
                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            adoConn.Execute .CommandText
                        
            '*** Obtención de cuentas
            .CommandType = adCmdStoredProc
                
            .CommandText = "up_GNObtenerCuentaAporte"
            .Parameters.Append .CreateParameter("ClasePersoneria", adChar, adParamInput, 2, strCodClasePersoneria)
            .Parameters.Append .CreateParameter("IndExtranjero", adChar, adParamInput, 1, strIndExtranjero)
            .Parameters.Append .CreateParameter("CodMoneda", adChar, adParamInput, 2, strCodMonedaOperacion)
            .Parameters.Append .CreateParameter("CuentaAporte", adChar, adParamOutput, 10, strCuentaAporte)
            .Parameters.Append .CreateParameter("CuentaSobreLaPar", adChar, adParamOutput, 10, strCuentaSobreLaPar)
            .Parameters.Append .CreateParameter("CuentaBajoLaPar", adChar, adParamOutput, 10, strCuentaBajoLaPar)
            .Execute
            
            If Not .Parameters("CuentaAporte") Then
                strCuentaAporte = .Parameters("CuentaAporte").Value
                strCuentaSobreLaPar = .Parameters("CuentaSobreLaPar").Value
                strCuentaBajoLaPar = .Parameters("CuentaBajoLaPar").Value
                .Parameters.Delete ("ClasePersoneria"): .Parameters.Delete ("IndExtranjero")
                .Parameters.Delete ("CodMoneda"): .Parameters.Delete ("CuentaAporte")
                .Parameters.Delete ("CuentaSobreLaPar"): .Parameters.Delete ("CuentaBajoLaPar")
            End If
            
            .CommandText = "up_GNObtenerCuentaAdministracion"
            .Parameters.Append .CreateParameter("CodCuentaAdministracion", adChar, adParamInput, 3, "018")
            .Parameters.Append .CreateParameter("TipoCuenta", adChar, adParamInput, 1, "R")
            .Parameters.Append .CreateParameter("CodCuentaMN", adChar, adParamOutput, 10, strCuentaSuscripcion)
            .Execute
            
            If Not .Parameters("CodCuentaMN") Then
                strCuentaSuscripcion = .Parameters("CodCuentaMN").Value
                .Parameters("CodCuentaAdministracion").Value = "001"
                .Parameters("CodCuentaMN").Value = strCuentaComision
            End If
            
            .Execute
            
            If Not .Parameters("CodCuentaMN") Then
                strCuentaComision = .Parameters("CodCuentaMN").Value
                .Parameters.Delete ("CodCuentaAdministracion"): .Parameters.Delete ("TipoCuenta")
                .Parameters.Delete ("CodCuentaMN")
            End If
                                                
            '*** Obtener el número del parámetro **
            .CommandText = "up_ACObtenerUltNumero"
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
            
            '*** Empieza la parte contable ***
            strDescripAsiento = "Suscripción a Valor Desconocido por " & Format$(Abs(adoRegistro("MontoOrdenCobroPago")), "###,###,###,##0.00")
            
            '*** Contabilizar ***
            strFechaGrabar = strFechaCierre & Space(1) & Format(Time, "hh:mm")
                
            '*** Cabecera Asiento Contable***
            .CommandText = "{ call up_ACAdicAsientoContable('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strNumAsiento & "','" & strFechaGrabar & "','" & _
                gstrPeriodoActual & "','" & gstrMesActual & "','','" & _
                strDescripAsiento & "','" & strCodMonedaOperacion & "','" & _
                Codigo_Moneda_Local & "','',''," & _
                CDec(Abs(adoRegistro("MontoOrdenCobroPago"))) & ",'" & Estado_Activo & "',5,'" & _
                strFechaCierre & Space(1) & Format(Time, "hh:ss") & "','" & _
                frmMainMdi.Tag & "',''," & _
                dblTipoCambioCierre & ",'','','" & _
                strDescripAsiento & "','','X','') }"
            adoConn.Execute .CommandText

            For intFor = 1 To 5
                Select Case intFor
                    Case 1
                        strDescripMovimiento = "Suscripción a valor desconocido"
                        strIndDebeHaber = "D"
                        curMontoMN = CCur(Abs(adoRegistro("MontoOrdenCobroPago")))
                        strCodCuenta = strCuentaSuscripcion
                        
                    Case 2
                        strDescripMovimiento = "Comisiones"
                        strIndDebeHaber = "H"
                        curMontoMN = CCur(Abs(adoRegistro("MontoComision"))) * -1
                        strCodCuenta = strCuentaComision
                        
                    Case 3
                        strDescripMovimiento = "IGV"
                        strIndDebeHaber = "H"
                        curMontoMN = CCur(Abs(adoRegistro("MontoIgv"))) * -1
                        strCodCuenta = strCuentaComision
                        
                    Case 4
                        strDescripMovimiento = "Aporte Fijo"
                        strIndDebeHaber = "H"
                        curMontoMN = curValorNominal * -1
                        strCodCuenta = strCuentaAporte
                        
                    Case 5
                        strDescripMovimiento = "Aporte Adicional"
                        strIndDebeHaber = "H"
                        strCodCuenta = strCuentaSobreLaPar
                        If curMontoSobreLaPar = 0 And curMontoBajoLaPar = 0 Then
                            curMontoMN = 0
                        ElseIf curMontoSobreLaPar > 0 Then
                            strIndDebeHaber = "H"
                            curMontoMN = curMontoSobreLaPar * -1
                            strCodCuenta = strCuentaSobreLaPar
                        Else
                            strIndDebeHaber = "D"
                            curMontoMN = curMontoBajoLaPar * -1
                            strCodCuenta = strCuentaBajoLaPar
                        End If
                        
                End Select
                curMontoME = 0
                curMontoContable = curMontoMN
                
                If strCodMonedaOperacion <> Codigo_Moneda_Local Then
                    curMontoContable = Round(curMontoMN * dblTipoCambioCierre, 2)
                    curMontoME = curMontoMN
                    curMontoMN = 0
                End If
                
                '*** Movimiento ***
                .CommandText = "{ call up_ACAdicAsientoContableDetalle('" & _
                    strNumAsiento & "','" & strCodFondo & "','" & _
                    gstrCodAdministradora & "'," & _
                    intFor & ",'" & _
                    strFechaGrabar & "','" & _
                    gstrPeriodoActual & "','" & gstrMesActual & "','" & _
                    strDescripMovimiento & "','" & strIndDebeHaber & "','" & _
                    strCodCuenta & "','" & strCodMonedaOperacion & "'," & _
                    CDec(curMontoMN) & "," & _
                    CDec(curMontoME) & "," & _
                    CDec(curMontoContable) & ",'" & _
                    strCodFile & "','" & strCodAnalitica & "','','') }"
                adoConn.Execute .CommandText
                                                    
            Next
            
            '*** Actualizar Secuenciales ***
            .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                Valor_NumComprobante & "','" & strNumAsiento & "') }"
            adoConn.Execute .CommandText
            
            '*** Actualizar Estado de la orden de caja ***
            .CommandText = "UPDATE MovimientoFondo SET EstadoOrden='" & Estado_Caja_Confirmado & "' " & _
                "WHERE (FechaObligacion >='" & strFechaCierre & "' AND FechaObligacion <'" & strFechaSiguiente & "') AND " & _
                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "ModuloOrigen='P' AND TipoOperacion='" & Codigo_Caja_Suscripcion & "' AND " & _
                "IndContable = 'X' AND EstadoOrden='" & Estado_Caja_NoConfirmado & "'"
            adoConn.Execute .CommandText
            
            adoRegistro.MoveNext
        
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
        
End Sub



Private Function LFchPriMov(s_ParCodFon As String, s_ParCodFil, s_ParCodAna, s_ParTipMov As String) As String

    Dim sensql As String, adoresultTmp As New Recordset
    
    sensql = "SELECT MIN(FCH_MOVI) FCH_MOVI FROM FMKARDEX WHERE "
    sensql = sensql + "COD_FOND='" + s_ParCodFon + "'"
    sensql = sensql + " AND COD_FILE='" + s_ParCodFil + "'"
    sensql = sensql + " AND COD_ANAL='" + s_ParCodAna + "'"
    sensql = sensql + " AND TIP_MOVI='" + s_ParTipMov + "'"
    sensql = sensql + " AND FLG_NCNF<>'X'"
    adoComm.CommandText = sensql
    Set adoresultTmp = adoComm.Execute
    If Not adoresultTmp.EOF Then
        If Not IsNull(adoresultTmp!FCH_MOVI) Then
            LFchPriMov = adoresultTmp!FCH_MOVI
        Else
            LFchPriMov = ""
            MsgBox "Error fecha kardex " & s_ParCodFil & "-" & s_ParCodAna
        End If
    Else
        LFchPriMov = ""
        MsgBox "Error fecha kardex " & s_ParCodFil & "-" & s_ParCodAna
    End If
    adoresultTmp.Close: Set adoresultTmp = Nothing
    
End Function

Private Function LFchUltMov(s_ParCodFon$, s_ParCodFil, s_ParCodAna, s_ParTipMov$) As String

    Dim sensql As String, adoresultTmp As New Recordset
    
    sensql = "SELECT MAX(FCH_MOVI) FCH_MOVI FROM FMKARDEX WHERE "
    sensql = sensql + "COD_FOND='" + s_ParCodFon$ + "'"
    sensql = sensql + " AND COD_FILE='" + s_ParCodFil + "'"
    sensql = sensql + " AND COD_ANAL='" + s_ParCodAna + "'"
    sensql = sensql + " AND TIP_MOVI='" + s_ParTipMov$ + "'"
    sensql = sensql + " AND FLG_NCNF<>'X'"
    adoComm.CommandText = sensql
    Set adoresultTmp = adoComm.Execute
    If Not adoresultTmp.EOF Then
        If Not IsNull(adoresultTmp!FCH_MOVI) Then
            LFchUltMov = adoresultTmp!FCH_MOVI
        Else
            LFchUltMov = ""
            MsgBox "Error fecha kardex " & s_ParCodFil & "-" & s_ParCodAna
        End If
    Else
        LFchUltMov = ""
        MsgBox "Error fecha kardex " & s_ParCodFil & "-" & s_ParCodAna
    End If
    adoresultTmp.Close: Set adoresultTmp = Nothing
    
End Function

'Private Sub ProvisionGastosFondo(strTipoCierre As String, strIndNoIncluyeEnPreCierre As String)
'
'    Dim adoRegistro             As ADODB.Recordset
'    Dim adoConsulta             As ADODB.Recordset
'    Dim strCodFile              As String, strCodDetalleFile            As String
'    Dim strNumAsiento           As String, strDescripAsiento            As String
'    Dim strIndDebeHaber         As String, strDescripMovimiento         As String
'    Dim strDescripGasto         As String, strFechaGrabar               As String
'    Dim intDiasProvision        As Integer, intCantRegistros            As Integer
'    Dim intContador             As Integer, intDiasCorridos             As Integer
'    Dim curMontoRenta           As Currency, curSaldoProvision          As Currency
'    Dim curMontoMovimientoMN    As Currency, curMontoMovimientoME       As Currency
'    Dim curMontoContable        As Currency, curValorAnterior           As Currency
'    Dim curValorActual          As Currency
'    Dim dblValorAjusteProv      As Double
'    Dim curValorTotal           As Currency
'    Dim intNumDiasPeriodo       As Integer
'    Dim intDiasProvision1       As Integer
'    Dim strTipoAuxiliar         As String
'    Dim strCodAuxiliar          As String
'    Dim strCodAnalitica         As String
'    Dim dblTipoCambioCierre     As Double
'    Dim indCumpleCondicion      As Boolean
'    Dim intNumRegistro          As Integer
'    Dim strSecOrdenPago         As String
'    Dim strNumCaja              As String
'    Dim intSecuencial           As Integer
'    Dim strIndVigente           As String
'
'    Dim dblValorTipoCambio          As Double, strTipoDocumento             As String
'    Dim strNumDocumento             As String, strTipoPersonaContraparte    As String
'    Dim strCodPersonaContraparte    As String
'    Dim strIndContracuenta          As String, strCodContracuenta           As String
'    Dim strCodFileContracuenta      As String, strCodAnaliticaContracuenta  As String
'    Dim strIndUltimoMovimiento      As String
'    Dim strIndTipoCambioContable    As String
'
'
'    frmMainMdi.stbMdi.Panels(3).Text = "Provisionando Gastos del Fondo..."
'
'    Set adoRegistro = New ADODB.Recordset
'    Set adoConsulta = New ADODB.Recordset
'
'    With adoComm
'
'
'        If strTipoCierre = Codigo_Cierre_Definitivo Then
'            .CommandText = "SELECT FG.*,FGP.NumPeriodo,FGP.FechaInicio,FGP.FechaVencimiento,'" & strTipoCierre & "' AS TipoCierre,'" & strFechaCierre & "' AS FechaCierre" & " FROM FondoGasto FG " & _
'                 "JOIN FondoGastoPeriodo FGP ON (FG.CodFondo = FGP.CodFondo AND FG.CodAdministradora = FGP.CodAdministradora AND FG.NumGasto = FGP.NumGasto) " & _
'                 "WHERE FGP.FechaInicio <= '" & strFechaCierre & "' AND FGP.FechaVencimiento >= '" & strFechaCierre & "' AND FG.CodFondo='" & strCodFondo & "' AND " & _
'                "FG.CodAdministradora='" & gstrCodAdministradora & "' AND FG.IndVigente='X' AND FG.IndNoIncluyeEnBalancePreCierre = '" & strIndNoIncluyeEnPreCierre & "'" & _
'                " AND FG.CodAplicacionDevengo = '" & Codigo_Aplica_Devengo_Periodica & "'"
'        Else
'            .CommandText = "SELECT FG.*,FGP.NumPeriodo,FGP.FechaInicio,FGP.FechaVencimiento,'" & strTipoCierre & "' AS TipoCierre,'" & strFechaCierre & "' AS FechaCierre" & " FROM FondoGastoTmp FG " & _
'                 "JOIN FondoGastoPeriodoTmp FGP ON (FG.CodFondo = FGP.CodFondo AND FG.CodAdministradora = FGP.CodAdministradora AND FG.NumGasto = FGP.NumGasto) " & _
'                 "WHERE FGP.FechaInicio <= '" & strFechaCierre & "' AND FGP.FechaVencimiento >= '" & strFechaCierre & "' AND FG.CodFondo='" & strCodFondo & "' AND " & _
'                "FG.CodAdministradora='" & gstrCodAdministradora & "' AND FG.IndVigente='X' AND FG.IndNoIncluyeEnBalancePreCierre = '" & strIndNoIncluyeEnPreCierre & "'" & _
'                " AND FG.CodAplicacionDevengo = '" & Codigo_Aplica_Devengo_Periodica & "'"
'        End If
'
'        Set adoRegistro = .Execute
'
'        Do While Not adoRegistro.EOF
'
'            strCodFile = Trim(adoRegistro("CodFile"))
'            strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
'
'            strTipoAuxiliar = "02"
'            strCodAuxiliar = adoRegistro("TipoProveedor") & adoRegistro("CodProveedor")
'
'            strIndVigente = adoRegistro("IndVigente")
'
'            .CommandText = "SELECT CodDetalleFile FROM InversionDetalleFile " & _
'                "WHERE CodFile='" & strCodFile & "' AND DescripDetalleFile='" & adoRegistro("CodCuenta") & "'"
'            Set adoConsulta = .Execute
'
'            If Not adoConsulta.EOF Then
'                strCodDetalleFile = adoConsulta("CodDetalleFile")
'            End If
'            adoConsulta.Close
'            Set adoConsulta = New ADODB.Recordset
'
'            '*** Obtener tipo de cambio ***
'            'dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
'            dblTipoCambioCierre = ObtenerValorTipoCambio(adoRegistro("CodMoneda"), Codigo_Moneda_Local, strFechaCierre, strFechaCierre, gstrCodClaseTipoCambioOperacionFondo, Codigo_Valor_TipoCambioCompra)
'
'
'            '*** Verificar Dinamica Contable ***
'            .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
'                "WHERE TipoOperacion='" & Codigo_Dinamica_Gasto & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
'                strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & _
'                "' AND CodMoneda= '" & IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"
'            Set adoConsulta = .Execute
'
'            If Not adoConsulta.EOF Then
'                If CInt(adoConsulta("NumRegistros")) > 0 Then
'                    intCantRegistros = CInt(adoConsulta("NumRegistros"))
'                Else
'                    MsgBox "NO EXISTE Dinámica Contable para la provisión", vbCritical
'                    adoConsulta.Close: Set adoConsulta = Nothing
'                    Exit Sub
'                End If
'            End If
'            adoConsulta.Close
'
'            '*** Obtener Descripción del Gasto ***
'            .CommandText = "SELECT DescripCuenta FROM PlanContable WHERE CodCuenta='" & adoRegistro("CodCuenta") & "'"
'            Set adoConsulta = .Execute
'
'            If Not adoConsulta.EOF Then
'                strDescripGasto = Trim(adoConsulta("DescripCuenta"))
'            End If
'            adoConsulta.Close
'
'            '*** Obtener las cuentas de inversión ***
'            Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, Trim(adoRegistro("CodMoneda")))
'
'            '*** Obtener Saldo de Provision ***
'            If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
'                .CommandText = "SELECT SaldoFinalContable Saldo "
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
'            'If Trim(adoRegistro("CodCuenta")) = Codigo_Cuenta_Comision_Fija Or Trim(adoRegistro("CodCuenta")) = Codigo_Cuenta_Comision_Variable Then
'            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
'                "CodCuenta='" & strCtaXPagar & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
'                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                "CodFondo='" & strCodFondo & "' AND CodMoneda ='" & adoRegistro("CodMoneda") & "' AND CodMonedaContable = '" & Codigo_Moneda_Local & "'"
'
'            Set adoConsulta = .Execute
'
'            If Not adoConsulta.EOF Then
'                curSaldoProvision = CDbl(adoConsulta("Saldo")) * -1
'            Else
'                curSaldoProvision = 0
'            End If
'            adoConsulta.Close
'
'            intDiasProvision = DateDiff("d", adoRegistro("FechaInicio"), adoRegistro("FechaVencimiento")) + 1
'            intDiasCorridos = DateDiff("d", adoRegistro("FechaInicio"), gdatFechaActual) + 1
'
'            curValorAnterior = curSaldoProvision
'
'            If adoRegistro("CodAplicacionDevengo") = Codigo_Aplica_Devengo_Periodica Then
'                Set adoConsulta = New ADODB.Recordset
'
'                '*** Obtener el número de días del peridodo de pago ***
'                .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' AND CodParametro='" & adoRegistro("CodFrecuenciaDevengo") & "'"
'                Set adoConsulta = adoComm.Execute
'
'                If Not adoConsulta.EOF Then
'                    intNumDiasPeriodo = CInt(adoConsulta("ValorParametro")) '*** Días del periodo  ***
'                Else
'                    intNumDiasPeriodo = 0
'                End If
'                adoConsulta.Close: Set adoConsulta = Nothing
'            Else
'                intNumDiasPeriodo = 0
'            End If
'
'            If adoRegistro("CodTipoValor") = Codigo_Tipo_Costo_Porcentaje Then
'                curValorTotal = CalculoInteres(adoRegistro("PorcenGasto"), adoRegistro("CodTipoTasa"), adoRegistro("CodPeriodoTasa"), adoRegistro("CodBaseAnual"), adoRegistro("MontoBaseCalculo"), adoRegistro("FechaInicio"), adoRegistro("FechaVencimiento"))
'            Else
'                curValorTotal = adoRegistro("MontoGasto")
'            End If
'
'            'JJCC 16/04/2012
'            Dim MyParser As New clsParser
'            Dim strIndCondicional As String
'            Dim strFormulaMonto As String
'            Dim strFormulaCondicion As String
'            Dim strCodFormulaDatos As String
'            Dim strParametros As String
'            Dim strMsgError As String
'            Dim strFechaPago As String
'            Dim strFechaVencimiento As String
'
'            indCumpleCondicion = True
'
'            curValorTotal = 0
'
'            If adoRegistro("CodFijaVariable") = Codigo_Tipo_Calculo_Fijo Then 'Fijo
'                If adoRegistro("CodTipoValor") = Codigo_Tipo_Costo_Porcentaje Then
'                    curValorTotal = CalculoInteres(adoRegistro("PorcenGasto"), adoRegistro("CodTipoTasa"), adoRegistro("CodPeriodoTasa"), adoRegistro("CodBaseAnual"), adoRegistro("MontoBaseCalculo"), adoRegistro("FechaInicio"), adoRegistro("FechaVencimiento"))
'                Else
'                    curValorTotal = adoRegistro("MontoGasto")
'                End If
'            Else ' Calculamos ejecutando la formula
'                'Traemos los datos de la formula
'                .CommandText = "SELECT FormulaMonto, indCondicion, FormulaCondicion, CodFormulaDatos  FROM Formula WHERE CodFormula='" & adoRegistro("CodFormula") & "'"
'                Set adoConsulta = adoComm.Execute
'
'                If Not adoConsulta.EOF Then
'                    strIndCondicional = adoConsulta("indCondicion")
'                    strFormulaMonto = adoConsulta("FormulaMonto")
'                    strFormulaCondicion = adoConsulta("FormulaCondicion")
'                    strCodFormulaDatos = adoConsulta("CodFormulaDatos")
'                Else
'                    strIndCondicional = ""
'                    strFormulaMonto = ""
'                    strFormulaCondicion = ""
'                    strCodFormulaDatos = ""
'                End If
'                adoConsulta.Close: Set adoConsulta = Nothing
'
''                strParametros = strCodFondo & "|" & gstrCodAdministradora & "|" & adoRegistro("CodFondoSerie") & "|" & strFechaCierre & "|" & strTipoCierre
'
'                'Parametros configurados
''                MsgBox adoRegistro.GetRows(1, intFila)
'
'                If strIndCondicional = Valor_Indicador Then
'                    indCumpleCondicion = MyParser.ParseExpression(strFormulaCondicion, strCodFormulaDatos, adoRegistro, strMsgError)
'                    If strMsgError <> "" Then
'                        MsgBox strMsgError, vbCritical
'                        Exit Sub
'                    End If
'                End If
'                If indCumpleCondicion Then
'                    curValorTotal = Round(MyParser.ParseExpression(strFormulaMonto, strCodFormulaDatos, adoRegistro, strMsgError), 2)
'
'                    If strMsgError <> "" Then
'                        MsgBox strMsgError, vbCritical
'                        Exit Sub
'                    End If
'                End If
'            End If
'
'            Set MyParser = Nothing
'            'FIN JJCC 16/04/2012
'
'
'            'Para el calculo prorratea sobre la base de Actual/x --osea sobre el numero real de dias del mes!
'            If adoRegistro("CodAplicacionDevengo") = Codigo_Aplica_Devengo_Periodica And adoRegistro("CodFijaVariable") = Codigo_Tipo_Calculo_Fijo Then
'                If intNumDiasPeriodo <> 0 Then
'                    If intDiasCorridos Mod intNumDiasPeriodo = 0 Then
'                        curMontoRenta = Round(curValorTotal / (intDiasProvision / intNumDiasPeriodo), 2)
'                    Else
'                        curMontoRenta = 0
'                    End If
'                Else
'                    curMontoRenta = 0
'                End If
'                curValorActual = curMontoRenta + curValorAnterior 'adoRegistro("MontoDevengo") + curMontoRenta
'            Else 'No Porratea, es inmediato
'                curMontoRenta = curValorTotal
'                curValorActual = curMontoRenta + curValorAnterior
'            End If
'
'            'UltimoDiaMes
'            'Control de remanentes
'            If adoRegistro("FechaVencimiento") = gdatFechaActual And adoRegistro("CodFijaVariable") = Codigo_Tipo_Calculo_Fijo Then
'                If (curValorTotal - curValorActual) <> 0 Then
'                    dblValorAjusteProv = (curValorTotal - curValorActual)
'                    curMontoRenta = curMontoRenta + dblValorAjusteProv
'                    curValorActual = curValorActual + dblValorAjusteProv
'                End If
'            End If
'
'            '*** Provisión ***
'            If curMontoRenta <> 0 Then
'                strDescripAsiento = "Provisión" & Space(1) & strDescripGasto
'                strDescripMovimiento = strDescripGasto
'                If curMontoRenta > 0 Then strDescripMovimiento = strDescripGasto
'
'                .CommandType = adCmdStoredProc
'                '*** Obtener el número del parámetro **
'                .CommandText = "up_ACObtenerUltNumero"
'                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "up_ACObtenerUltNumeroTmp"  '*** Simulación ***
'
'                .Parameters.Append .CreateParameter("CodFondo", adChar, adParamInput, 3, strCodFondo)
'                .Parameters.Append .CreateParameter("CodAdministradora", adChar, adParamInput, 3, gstrCodAdministradora)
'                .Parameters.Append .CreateParameter("CodParametro", adChar, adParamInput, 6, Valor_NumComprobante)
'                .Parameters.Append .CreateParameter("NuevoNumero", adChar, adParamOutput, 10, "")
'                .Execute
'
'                If Not .Parameters("NuevoNumero") Then
'                    strNumAsiento = .Parameters("NuevoNumero").Value
'                    .Parameters.Delete ("CodFondo")
'                    .Parameters.Delete ("CodAdministradora")
'                    .Parameters.Delete ("CodParametro")
'                    .Parameters.Delete ("NuevoNumero")
'                End If
'
'                .CommandType = adCmdText
'
'                'On Error GoTo Ctrl_Error
'
'                '*** Contabilizar ***
'                strFechaGrabar = strFechaCierre & Space(1) & Format(Time, "hh:mm")
'
'                '*** Cabecera ***
'                .CommandText = "{ call up_ACAdicAsientoContable('"
'                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableTmp('"  '*** Simulación ***
'
'                .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
'                    strFechaGrabar & "','" & _
'                    gstrPeriodoActual & "','" & gstrMesActual & "','" & Codigo_Tipo_Asiento_Provision_Gastos & "','" & _
'                    strDescripAsiento & "','" & Trim(adoRegistro("CodMoneda")) & "','" & _
'                    Codigo_Moneda_Local & "','',''," & _
'                    CDec(curMontoRenta) & ",'" & Estado_Activo & "'," & _
'                    intCantRegistros & ",'" & _
'                    strFechaCierre & Space(1) & Format(Time, "hh:ss") & "','" & _
'                    strCodModulo & "',''," & _
'                    CDec(dblTipoCambioCierre) & ",'','','" & _
'                    strDescripAsiento & "','','X','') }"
'                adoConn.Execute .CommandText
'
'                '*** Detalle ***
'                .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
'                    "WHERE TipoOperacion='" & Codigo_Dinamica_Gasto & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
'                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
'                    IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"
'                Set adoConsulta = .Execute
'
'                Do While Not adoConsulta.EOF
'
'                    curMontoMovimientoMN = 0
'
'                    Select Case Trim(adoConsulta("TipoCuentaInversion"))
'                        Case Codigo_CtaInversion
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaProvInteres
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaInteres
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaInteresVencido
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaVacCorrido
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaXPagar
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaXCobrar
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaInteresCorrido
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaProvReajusteK
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaReajusteK
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaProvFlucMercado
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaFlucMercado
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaProvInteresVac
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaInteresVac
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaIntCorridoK
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaProvFlucK
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaFlucK
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaInversionTransito
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaProvGasto
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaIngresoOperacional
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaCosto
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaCostoSAB
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaCostoBVL
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaCostoCavali
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaCostoFondoLiquidacion
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaCostoGastosBancarios
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaCostoComisionEspecial
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaCostoFondoGarantia
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaCostoConasev
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaComision
'                            curMontoMovimientoMN = curMontoRenta
'
'                    End Select
'
'                    strIndDebeHaber = Trim(adoConsulta("IndDebeHaber"))
'
'                    If strIndDebeHaber = "H" Then
'                        curMontoMovimientoMN = curMontoMovimientoMN * -1
'                        If curMontoMovimientoMN > 0 Then strIndDebeHaber = "D"
'                    ElseIf strIndDebeHaber = "D" Then
'                        If curMontoMovimientoMN < 0 Then strIndDebeHaber = "H"
'                    End If
'
'                    If strIndDebeHaber = "T" Then
'                        If curMontoMovimientoMN > 0 Then
'                            strIndDebeHaber = "D"
'                        Else
'                            strIndDebeHaber = "H"
'                        End If
'                    End If
'
'                    strDescripMovimiento = Trim(adoConsulta("DescripDinamica"))
'                    curMontoMovimientoME = 0
'                    curMontoContable = curMontoMovimientoMN
'
'                    dblValorTipoCambio = 1
'
'                    If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
'                        curMontoContable = Round(curMontoMovimientoMN * CDec(dblTipoCambioCierre), 2)
'                        curMontoMovimientoME = curMontoMovimientoMN
'                        curMontoMovimientoMN = 0
'                        dblValorTipoCambio = dblTipoCambioCierre
'                    End If
'
'                    strTipoDocumento = ""
'                    strNumDocumento = ""
'                    strTipoPersonaContraparte = ""
'                    strCodPersonaContraparte = ""
'                    strIndContracuenta = ""
'                    strCodContracuenta = ""
'                    strCodFileContracuenta = ""
'                    strCodAnaliticaContracuenta = ""
'                    strIndUltimoMovimiento = ""
'
'                    strIndTipoCambioContable = Valor_Caracter
'
'                    If Mid(adoConsulta("CodCuenta"), 1, 1) <> "6" And Mid(adoConsulta("CodCuenta"), 1, 1) <> "7" And Mid(adoConsulta("CodCuenta"), 1, 1) <> "5" Then
'                        strIndTipoCambioContable = Valor_Indicador
'                    End If
'
'                    If curMontoContable <> 0 Then
'                        '*** Movimiento ***
'                        .CommandText = "{ call up_ACAdicAsientoContableDetalle('"
'                        If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableDetalleTmp('"
'
'                        .CommandText = .CommandText & strNumAsiento & "','" & strCodFondo & "','" & _
'                            gstrCodAdministradora & "'," & _
'                            CInt(adoConsulta("NumSecuencial")) & ",'" & _
'                            strFechaGrabar & "','" & _
'                            gstrPeriodoActual & "','" & _
'                            gstrMesActual & "','" & _
'                            strDescripMovimiento & "','" & _
'                            strIndDebeHaber & "','" & _
'                            Trim(adoConsulta("CodCuenta")) & "','" & _
'                            Trim(adoRegistro("CodMoneda")) & "'," & _
'                            CDec(curMontoMovimientoMN) & "," & _
'                            CDec(curMontoMovimientoME) & "," & _
'                            CDec(curMontoContable) & "," & _
'                            dblValorTipoCambio & ",'" & _
'                            Trim(adoRegistro("CodFile")) & "','" & _
'                            Trim(adoRegistro("CodAnalitica")) & "','" & _
'                            strTipoDocumento & "','" & _
'                            strNumDocumento & "','" & _
'                            strTipoPersonaContraparte & "','" & _
'                            strCodPersonaContraparte & "','" & _
'                            strIndContracuenta & "','" & _
'                            strCodContracuenta & "','" & _
'                            strCodFileContracuenta & "','" & _
'                            strCodAnaliticaContracuenta & "','" & _
'                            strIndUltimoMovimiento & "','','','<TipoCambioReemplazo/>','I','" & gstrCodClaseTipoCambioOperacionFondo & "','" & Codigo_Valor_TipoCambioCompra & "','" & strIndTipoCambioContable & "') }"
'
'                        adoConn.Execute .CommandText
'
'                        '*** Validar valor de cuenta contable ***
'                        If Trim(adoConsulta("CodCuenta")) = Valor_Caracter Then
'                            MsgBox "Registro Nro. " & CStr(intContador) & " de Asiento Contable no tiene cuenta asignada", vbCritical, "Valorización Depósitos"
'                            gblnRollBack = True
'                            Exit Sub
'                        End If
'                    End If
'
'                    adoConsulta.MoveNext
'                Loop
'                adoConsulta.Close: Set adoConsulta = Nothing
'
'                '-- Verifica y ajusta posibles descuadres
'                .CommandText = "{ call up_ACProcAsientoContableAjusteTC('"
'                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_ACProcAsientoContableAjusteTCTmp('"
'
'                .CommandText = .CommandText & strCodFondo & "','" & _
'                        gstrCodAdministradora & "','" & _
'                        strNumAsiento & "') }"
'                adoConn.Execute .CommandText
'
'                '*** Actualizar el número del parámetro **
'                .CommandText = "{ call up_ACActUltNumero('"
'                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNActUltNumeroTmp('"
'
'                .CommandText = .CommandText & strCodFondo & "','" & _
'                    gstrCodAdministradora & "','" & _
'                    Valor_NumComprobante & "','" & _
'                    strNumAsiento & "') }"
'                adoConn.Execute .CommandText
'
'                If strTipoCierre = Codigo_Cierre_Definitivo Then
'                    .CommandText = "UPDATE FondoGasto SET "
'                Else
'                    .CommandText = "UPDATE FondoGastoTmp SET "
'                End If
'
'                If Convertyyyymmdd(adoRegistro("FechaFinal")) = strFechaCierre Then
'                    strIndVigente = Valor_Caracter
'                    .CommandText = .CommandText & "IndVigente='" & strIndVigente & "',"
'                End If
'
'                .CommandText = .CommandText & "MontoDevengo=" & curValorActual & _
'                " WHERE NumGasto=" & adoRegistro("NumGasto") & " AND " & _
'                       "CodCuenta='" & Trim(adoRegistro("CodCuenta")) & "' AND CodFondo='" & strCodFondo & "' AND " & _
'                       "CodAdministradora='" & gstrCodAdministradora & "'"
'
'                adoConn.Execute .CommandText
'
'                ''REVISAR CUANDO HAYA COMISION VARIABLE
'                If strIndNoIncluyeEnPreCierre = Valor_Indicador Then
'                    If strTipoCierre = Codigo_Cierre_Definitivo Then
'                        .CommandText = "UPDATE FondoValorCuota SET "
'                    Else
'                        .CommandText = "UPDATE FondoValorCuotaTmp SET "
'                    End If
'
'                    .CommandText = .CommandText & "MontoComisionAdminFija = " & curMontoRenta & _
'                    " WHERE FechaCuota = '" & strFechaCierre & "' AND " & _
'                           "CodFondo = '" & strCodFondo & "' AND " & _
'                           "CodAdministradora = '" & gstrCodAdministradora & "'"
'
'                    adoConn.Execute .CommandText
'                End If
'
'                curSaldoProvision = curSaldoProvision + curMontoRenta
'
'                'JJCC:FIN DE MES PARA LA REMUNERACIÓN (FACTURACIÓN)
'                Dim datPrimerDia As Date, datUltimoDia As Date
'
'                datPrimerDia = DateSerial(Year(gdatFechaActual), Month(gdatFechaActual), 1)
'                If datPrimerDia < adoRegistro("FechaInicio") Then
'                    datPrimerDia = adoRegistro("FechaInicio")
'                End If
'
'                datUltimoDia = DateAdd("d", -1, DateSerial(Year(DateAdd("m", 1, gdatFechaActual)), Month(DateAdd("m", 1, gdatFechaActual)), 1))
'                If datUltimoDia > adoRegistro("FechaVencimiento") Then
'                    datUltimoDia = adoRegistro("FechaVencimiento")
'                End If
'
'                If gdatFechaActual = adoRegistro("FechaVencimiento") Then
'
'                    .CommandText = "{ call up_GNManFondoGasto('" & strCodFondo & "','" & _
'                        gstrCodAdministradora & "'," & adoRegistro("NumGasto") & ",'" & Convertyyyymmdd(adoRegistro("FechaDefinicion")) & "','" & adoRegistro("CodCuenta") & "','" & _
'                        adoRegistro("CodFile") & "','" & adoRegistro("CodAnalitica") & "','" & adoRegistro("TipoProveedor") & "','" & adoRegistro("CodProveedor") & "','" & adoRegistro("DescripGasto") & "','" & _
'                        Convertyyyymmdd(adoRegistro("FechaConfirma")) & "','" & Convertyyyymmdd(adoRegistro("FechaInicial")) & "','" & Convertyyyymmdd(adoRegistro("FechaFinal")) & "','" & _
'                        adoRegistro("CodTipoGasto") & "','" & Valor_Caracter & "','" & strIndVigente & "'," & dblTipoCambioCierre & ",'" & adoRegistro("CodMoneda") & "','" & adoRegistro("CodTipoValor") & "'," & _
'                        CCur(curValorActual) & "," & adoRegistro("PorcenGasto") & ",'" & adoRegistro("CodTipoTasa") & "','" & adoRegistro("CodPeriodoTasa") & "'," & adoRegistro("MontoBaseCalculo") & ",'" & adoRegistro("CodBaseAnual") & "',0,'" & _
'                        adoRegistro("CodModalidadPago") & "','" & adoRegistro("CodTipoPago") & "'," & adoRegistro("NumPeriodoPago") & ",'" & adoRegistro("CodPeriodoPago") & "','" & _
'                        Convertyyyymmdd(adoRegistro("FechaVencimiento")) & "','" & adoRegistro("CodTipoDesplazamiento") & "','" & adoRegistro("IndFinMes") & "','" & adoRegistro("CodTipoDevengo") & "','" & _
'                        adoRegistro("CodAplicacionDevengo") & "','" & adoRegistro("CodFrecuenciaDevengo") & "','" & adoRegistro("CodCreditoFiscal") & "','" & adoRegistro("IndNoIncluyeEnBalancePreCierre") & "','U','" & adoRegistro("CodFijaVariable") & "','" & adoRegistro("CodFormula") & "') }"
'
'                    adoConn.Execute .CommandText
'
'                End If
'
'            End If
'
'            adoRegistro.MoveNext
'        Loop
'        adoRegistro.Close: Set adoRegistro = Nothing
'    End With
'
'    Exit Sub
'
'Ctrl_Error:
'    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
'    Me.MousePointer = vbDefault
'
'End Sub

Private Sub ValorizacionDepositos(strTipoCierre As String)

    Dim adoRegistro             As ADODB.Recordset, adoConsulta     As ADODB.Recordset
    Dim dblPrecioCierre         As Double, dblPrecioPromedio        As Double
    Dim dblTirCierre            As Double, dblFactorDiario          As Double
    Dim dblTasaInteres          As Double, dblFactorDiarioCupon     As Double
    Dim curSaldoInversion       As Currency, curSaldoInteresCorrido As Currency
    Dim curSaldoFluctuacion     As Currency, curValorAnterior       As Currency
    Dim curValorActual          As Currency, curMontoRenta          As Currency
    Dim curMontoContable        As Currency, curMontoMovimientoMN   As Currency
    Dim curMontoMovimientoME    As Currency, curSaldoValorizar      As Currency
    Dim intCantRegistros        As Integer, intContador             As Integer
    Dim intRegistro             As Integer, intBaseCalculo          As Integer
    Dim intDiasPlazo            As Integer, intDiasDeRenta          As Integer
    Dim strNumAsiento           As String, strDescripAsiento        As String
    Dim strDescripMovimiento    As String, strIndDebeHaber          As String
    Dim strCodCuenta            As String, strFiles                 As String
    Dim strCodFile              As String, strModalidadInteres      As String
    Dim strCodTasa              As String, strIndCuponCero          As String
    Dim strCodDetalleFile       As String, strNemonico              As String
    Dim strFechaGrabar          As String, strBaseAnual             As String
    Dim dblTipoCambioCierre     As Double, strNumOperacion          As String
    Dim curMontoRentaContable       As Currency
    
    Dim dblValorTipoCambio          As Double, strTipoDocumento             As String
    Dim strNumDocumento             As String, strTipoPersonaContraparte    As String
    Dim strCodPersonaContraparte    As String
    Dim strIndContracuenta          As String, strCodContracuenta           As String
    Dim strCodFileContracuenta      As String, strCodAnaliticaContracuenta  As String
    Dim strIndUltimoMovimiento      As String, strCodAnalitica              As String
    Dim strCodTitulo                As String, strNumCuota                  As String
        
    '*****
    Dim dblValorNominal         As Double
    
    '*** Rentabilidad de Valores de Depósito ***
    frmMainMdi.stbMdi.Panels(3).Text = "Calculando rentabilidad de Valores de Depósito..."
       
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT IO.NumOperacion, CC.CodFile, CC.CodAnalitica, CC.CodTitulo, CC.NumCuota, CC.NumSecuencial, IO.FechaVencimiento  " & _
            "FROM InversionKardex IK " & _
            "JOIN InstrumentoInversion II ON (II.CodFondo = IK.CodFondo and II.CodAdministradora = IK.CodAdministradora " & _
            "AND II.CodFile=IK.CodFile AND II.CodAnalitica=IK.CodAnalitica and IK.CodTitulo = II.CodTitulo) " & _
            "JOIN InversionOperacion IO ON (IO.CodFondo = IK.CodFondo AND IO.CodAdministradora = IK.CodAdministradora AND IO.NumOperacion = IK.NumOperacion) " & _
            "JOIN InversionOperacionCalendarioCuota CC ON (CC.CodFondo = IK.CodFondo AND CC.CodAdministradora = IK.CodAdministradora AND " & _
            "CC.NumOperacionOrig = dbo.uf_IVObtenerNumOperacionOrigen(IK.CodFondo,IK.CodAdministradora,IK.CodFile,IK.CodAnalitica) " & _
            "AND CC.CodFile = IK.CodFile AND CC.CodAnalitica = IK.CodAnalitica) " & _
            "WHERE " & _
            "IK.CodAdministradora = '" & gstrCodAdministradora & "' AND IK.CodFondo = '" & strCodFondo & "' AND " & _
            "IK.CodFile IN ('003','011') AND IK.SaldoFinal > 0 AND " & _
            "IK.NumKardex = dbo.uf_IVObtenerUltimoMovimientoKardexValor(IK.CodFondo,IK.CodAdministradora,IK.CodTitulo,'" & strFechaCierre & "') AND " & _
            "CC.NumSecuencial = dbo.uf_IVObtenerUltimoCalendarioCuotaVigente(CC.CodFondo,CC.CodAdministradora,CC.CodFile,CC.CodAnalitica,CC.NumCuota,'" & strFechaCierre & "') ORDER BY IK.CodFile, IK.CodAnalitica"
        Set adoRegistro = .Execute
    
        Do Until adoRegistro.EOF
            strNumOperacion = Trim(adoRegistro("NumOperacion"))
            strCodFile = Trim(adoRegistro("CodFile"))
            strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
            strCodTitulo = Trim(adoRegistro("CodTitulo"))
            strNumCuota = Trim$(adoRegistro("NumCuota"))
        
            '*** Grabamos Devengado ***
            .CommandText = "{ call up_IVAdicInversionDevengado('" & strCodFondo & "','" & gstrCodAdministradora & _
                            "','" & strCodFile & "','" & strCodAnalitica & "','" & strCodTitulo & _
                            "','" & strNumCuota & "','" & strFechaCierre & "','" & strTipoCierre & "') }"
            .Execute
            
            '*** Contabilizamos Rendimiento Devengado ***
            .CommandText = "{ call up_ACProcContabilizarOperacion('" & strCodFondo & "','" & gstrCodAdministradora & _
                "','" & strFechaCierre & "','" & strCodFile & "','" & strNumOperacion & "', '" & Codigo_Caja_Provision & "') }"
            .Execute
                
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
        
        
'        Do Until adoRegistro.EOF
'            strCodFile = Trim(adoRegistro("CodFile"))
'            strCodDetalleFile = Trim(adoRegistro("CodDetalleFile"))
'            strModalidadInteres = Trim(adoRegistro("CodDetalleFile"))
'            strCodTasa = Trim(adoRegistro("CodTipoTasa"))
'            strBaseAnual = Trim(adoRegistro("BaseAnual"))
'            dblTasaInteres = CDbl(adoRegistro("TasaInteres"))
'            intDiasPlazo = CInt(adoRegistro("DiasPlazo"))
'            strIndCuponCero = Trim(adoRegistro("IndCuponCero"))
'            strNemonico = Trim(adoRegistro("Nemotecnico"))
'            curSaldoValorizar = CCur(adoRegistro("SaldoFinal"))
'
'            curSaldoFluctuacion = 0
'            curSaldoInversion = 0
'            curSaldoInteresCorrido = 0
'            dblPrecioCierre = 0
'            dblTirCierre = 0
'            dblPrecioPromedio = 0
'
'            '********
'            dblValorNominal = CDbl(adoRegistro("ValorNominal"))
'
'            intDiasDeRenta = DateDiff("d", CVDate(adoRegistro("FechaEmision")), gdatFechaActual) + 1
'
'            If strBaseAnual = Codigo_Base_30_360 Or strBaseAnual = Codigo_Base_30_365 Then intDiasDeRenta = Dias360(CVDate(adoRegistro("FechaEmision")), gdatFechaActual, True) + 1
'
'            Set adoConsulta = New ADODB.Recordset
'
'            '*** Verificar Dinamica Contable ***
'            .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
'                "WHERE TipoOperacion='" & Codigo_Dinamica_Provision & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
'                strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
'                IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"
'
'            Set adoConsulta = .Execute
'
'            If Not adoConsulta.EOF Then
'                If CInt(adoConsulta("NumRegistros")) > 0 Then
'                    intCantRegistros = CInt(adoConsulta("NumRegistros"))
'                Else
'                    MsgBox "NO EXISTE Dinámica Contable para la valorización", vbCritical
'                    adoConsulta.Close: Set adoConsulta = Nothing
'                    Exit Sub
'                End If
'            End If
'            adoConsulta.Close
'
'            '*** Obtener Ultimo Precio de Cierre registrado ***
'            .CommandText = "{ call up_IVSelDatoInstrumentoInversion(2,'" & _
'                Trim(adoRegistro("CodTitulo")) & "','19000101') }"
'            Set adoConsulta = .Execute
'
'            If Not adoConsulta.EOF Then
'                dblPrecioCierre = CDbl(adoConsulta("PrecioCierre"))
'                dblTirCierre = CDbl(adoConsulta("TirCierre"))
'                dblPrecioPromedio = CDbl(adoConsulta("PrecioPromedio"))
'            End If
'            adoConsulta.Close
'
'
'            '*** Obtener el factor diario del cupón ***
'            .CommandText = "SELECT FactorDiario FROM InstrumentoInversionCalendario " & _
'                "WHERE CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "'"
'            Set adoConsulta = .Execute
'
'            If Not adoConsulta.EOF Then
'                dblFactorDiarioCupon = CDbl(adoConsulta("FactorDiario"))
'            End If
'            adoConsulta.Close
'
'            '*** Obtener las cuentas de inversión ***
'            Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, Trim(adoRegistro("CodMoneda")))
'
'            '*** Obtener tipo de cambio ***
'            'dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
'            dblTipoCambioCierre = ObtenerValorTipoCambio(adoRegistro("CodMoneda"), Codigo_Moneda_Local, strFechaCierre, strFechaCierre, gstrCodClaseTipoCambioOperacionFondo, Codigo_Valor_TipoCambioCompra)
'
'
'            '*** Obtener Saldo de Inversión ***
'            If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
'                .CommandText = "SELECT SaldoFinalContable Saldo "
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
'                "CodCuenta='" & strCtaInversion & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
'                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
'                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
'
'            Set adoConsulta = .Execute
'
'            If Not adoConsulta.EOF Then
'                curSaldoInversion = CDbl(adoConsulta("Saldo"))
'            End If
'            adoConsulta.Close
'
'            '*** Obtener Saldo de Interés Corrido ***
'            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
'                .CommandText = "SELECT SaldoFinalContable Saldo "
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
'                "CodCuenta='" & strCtaInteresCorrido & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
'                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
'                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
'
'            Set adoConsulta = .Execute
'
'            If Not adoConsulta.EOF Then
'                curSaldoInteresCorrido = CDbl(adoConsulta("Saldo"))
'            End If
'            adoConsulta.Close
'
'            '*** Obtener Saldo de Provisión ***
'            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
'                .CommandText = "SELECT SaldoFinalContable Saldo "
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
'                "CodCuenta='" & strCtaProvInteres & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
'                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
'                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
'
'            Set adoConsulta = .Execute
'
'            If Not adoConsulta.EOF Then
'                curSaldoFluctuacion = CDbl(adoConsulta("Saldo"))
'            End If
'            adoConsulta.Close
'
'            curValorAnterior = curSaldoInteresCorrido + curSaldoFluctuacion
'
'            If strBaseAnual = Codigo_Base_Actual_365 Or strBaseAnual = Codigo_Base_30_365 Or strBaseAnual = Codigo_Base_Actual_Actual Then
'                intBaseCalculo = 365
'            Else
'                intBaseCalculo = 360
'            End If
'
'            If Trim(adoRegistro("CodSubDetalleFile")) <> Valor_Caracter Then strModalidadInteres = Trim(adoRegistro("CodSubDetalleFile"))
'
'                    If strCodTasa = Codigo_Tipo_Tasa_Efectiva Then
'                        If strBaseAnual = Codigo_Base_30_360 Or strBaseAnual = Codigo_Base_30_365 Then
'                            dblFactorDiario = ((1 + dblTasaInteres * 0.01) ^ (intDiasDeRenta / intBaseCalculo)) - 1
'                        Else
'                            dblFactorDiario = ((1 + dblTasaInteres * 0.01) ^ (intDiasDeRenta / intBaseCalculo)) - 1
'                        End If
''                        dblFactorDiario = ((1 + CDbl(((1 + (dblTasaInteres / 100)) ^ (intDiasPlazo / intBaseCalculo)) - 1)) ^ (1 / intDiasPlazo)) - 1
'                    Else
'                        If strBaseAnual = Codigo_Base_30_360 Or strBaseAnual = Codigo_Base_30_365 Then
'                            dblFactorDiario = (((dblTasaInteres * 0.01) / intBaseCalculo) * intDiasDeRenta)
'                        Else
'                            dblFactorDiario = (((dblTasaInteres * 0.01) / intBaseCalculo) * intDiasDeRenta)
'                        End If
''                        dblFactorDiario = (CDbl(((1 + (dblTasaInteres / 100)) / intBaseCalculo)))
'                    End If
''                End If
'
'                curValorActual = Round(curSaldoValorizar * dblValorNominal * dblFactorDiario, 2)
'
'            curMontoRenta = Round(curValorActual - curValorAnterior, 2)
'
'            '*** Ganancia/Pérdida ***
'            If curMontoRenta <> 0 Then
'                'strDescripAsiento = "Valorización" & Space(1) & "(" & Trim(adoRegistro("CodFile")) & "-" & Trim(adoRegistro("CodAnalitica")) & ")"
'                strDescripAsiento = "Valorización" & Space(1) & strNemonico
'                strDescripMovimiento = "Pérdida"
'                If curMontoRenta > 0 Then strDescripMovimiento = "Ganancia"
'
'                If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
'                    curMontoRentaContable = Round(curMontoRenta * dblTipoCambioCierre, 2)
'                Else
'                    curMontoRentaContable = curMontoRenta
'                End If
'
'                .CommandType = adCmdStoredProc
'                '*** Obtener el número del parámetro **
'                .CommandText = "up_ACObtenerUltNumero"
'                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "up_ACObtenerUltNumeroTmp"  '*** Simulación ***
'
'                .Parameters.Append .CreateParameter("CodFondo", adChar, adParamInput, 3, strCodFondo)
'                .Parameters.Append .CreateParameter("CodAdministradora", adChar, adParamInput, 3, gstrCodAdministradora)
'                .Parameters.Append .CreateParameter("CodParametro", adChar, adParamInput, 6, Valor_NumComprobante)
'                .Parameters.Append .CreateParameter("NuevoNumero", adChar, adParamOutput, 10, "")
'                .Execute
'
'                If Not .Parameters("NuevoNumero") Then
'                    strNumAsiento = .Parameters("NuevoNumero").Value
'                    .Parameters.Delete ("CodFondo")
'                    .Parameters.Delete ("CodAdministradora")
'                    .Parameters.Delete ("CodParametro")
'                    .Parameters.Delete ("NuevoNumero")
'                End If
'
'                .CommandType = adCmdText
'
'                '*** Contabilizar ***
'                strFechaGrabar = strFechaCierre & Space(1) & Format(Time, "hh:mm")
'                '*** Cabecera ***
'                .CommandText = "{ call up_ACAdicAsientoContable('"
'                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableTmp('"  '*** Simulación ***
'
'                .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
'                    strFechaGrabar & "','" & gstrPeriodoActual & "','" & gstrMesActual & "','" & Valor_Caracter & "','" & _
'                    strDescripAsiento & "','" & Trim(adoRegistro("CodMoneda")) & "','" & _
'                    Codigo_Moneda_Local & "','" & "','" & "'," & _
'                    CDec(curMontoRenta) & ",'" & Estado_Activo & "'," & _
'                    intCantRegistros & ",'" & strFechaCierre & Space(1) & Format(Time, "hh:ss") & "','" & _
'                    strCodModulo & "','" & "'," & dblTipoCambioCierre & ",'" & _
'                    "','" & "','" & strDescripAsiento & "','" & "','" & "X','') }"
'                adoConn.Execute .CommandText
'
'                '*** Detalle ***
'                .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
'                    "WHERE TipoOperacion='" & Codigo_Dinamica_Provision & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
'                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
'                    IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"
'                Set adoConsulta = .Execute
'
'                Do While Not adoConsulta.EOF
'
'                    Select Case Trim(adoConsulta("TipoCuentaInversion"))
'                        Case Codigo_CtaInversion
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaProvInteres
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaInteres
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaCosto
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaIngresoOperacional
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaInteresVencido
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaVacCorrido
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaXPagar
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaXCobrar
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaInteresCorrido
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaProvReajusteK
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaReajusteK
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaProvFlucMercado
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaFlucMercado
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaProvInteresVac
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaInteresVac
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaIntCorridoK
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaProvFlucK
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaFlucK
'                            curMontoMovimientoMN = curMontoRenta
'
'                        Case Codigo_CtaInversionTransito
'                            curMontoMovimientoMN = curMontoRenta
'
'                    End Select
'
'                    strIndDebeHaber = Trim(adoConsulta("IndDebeHaber"))
'                    If strIndDebeHaber = "H" Then
'                        curMontoMovimientoMN = curMontoMovimientoMN * -1
'                        If curMontoMovimientoMN > 0 Then strIndDebeHaber = "D"
'                    ElseIf strIndDebeHaber = "D" Then
'                        If curMontoMovimientoMN < 0 Then strIndDebeHaber = "H"
'                    End If
'
'                    If strIndDebeHaber = "T" Then
'                        If curMontoMovimientoMN > 0 Then
'                            strIndDebeHaber = "D"
'                        Else
'                            strIndDebeHaber = "H"
'                        End If
'                    End If
'                    strDescripMovimiento = Trim(adoConsulta("DescripDinamica"))
'                    curMontoMovimientoME = 0
'                    curMontoContable = curMontoMovimientoMN
'
'                    dblValorTipoCambio = 1
'
'                    If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
'                        curMontoContable = Round(curMontoMovimientoMN * dblTipoCambioCierre, 2)
'                        'curMontoContable = Round(curMontoMovimientoMN * 2.697, 2)
'                        curMontoMovimientoME = curMontoMovimientoMN
'                        dblValorTipoCambio = dblTipoCambioCierre
'                    End If
'
'                    strTipoDocumento = ""
'                    strNumDocumento = ""
'                    strTipoPersonaContraparte = ""
'                    strCodPersonaContraparte = ""
'                    strIndContracuenta = ""
'                    strCodContracuenta = ""
'                    strCodFileContracuenta = ""
'                    strCodAnaliticaContracuenta = ""
'                    strIndUltimoMovimiento = ""
'
'                    '*** Movimiento ***
'                    .CommandText = "{ call up_ACAdicAsientoContableDetalle('"
'                    If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableDetalleTmp('"
'
'                    .CommandText = .CommandText & strNumAsiento & "','" & strCodFondo & "','" & _
'                        gstrCodAdministradora & "'," & _
'                        CInt(adoConsulta("NumSecuencial")) & ",'" & _
'                        strFechaGrabar & "','" & _
'                        gstrPeriodoActual & "','" & _
'                        gstrMesActual & "','" & _
'                        strDescripMovimiento & "','" & _
'                        strIndDebeHaber & "','" & _
'                        Trim(adoConsulta("CodCuenta")) & "','" & _
'                        Trim(adoRegistro("CodMoneda")) & "'," & _
'                        CDec(curMontoMovimientoMN) & "," & _
'                        CDec(curMontoMovimientoME) & "," & _
'                        CDec(curMontoContable) & "," & _
'                        dblValorTipoCambio & ",'" & _
'                        Trim(adoRegistro("CodFile")) & "','" & _
'                        Trim(adoRegistro("CodAnalitica")) & "','" & _
'                        strTipoDocumento & "','" & _
'                        strNumDocumento & "','" & _
'                        strTipoPersonaContraparte & "','" & _
'                        strCodPersonaContraparte & "','" & _
'                        strIndContracuenta & "','" & _
'                        strCodContracuenta & "','" & _
'                        strCodFileContracuenta & "','" & _
'                        strCodAnaliticaContracuenta & "','" & _
'                        strIndUltimoMovimiento & "') }"
'
'                    adoConn.Execute .CommandText
'
'
'                    '*** Validar valor de cuenta contable ***
'                    If Trim(adoConsulta("CodCuenta")) = Valor_Caracter Then
'                        MsgBox "Registro Nro. " & CStr(intContador) & " de Asiento Contable no tiene cuenta asignada", vbCritical, "Valorización Depósitos"
'                        gblnRollBack = True
'                        Exit Sub
'                    End If
'
'                    '*** Insertar en up_GNManInversionValorizacionDiaria **
'                    .CommandText = "{ call up_GNManInversionValorizacionDiaria('"
'                    If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNManInversionValorizacionDiariaTmp('"
'
'                    .CommandText = .CommandText & strCodFondo & "','" & _
'                        gstrCodAdministradora & "','" & _
'                        strFechaCierre & "','" & _
'                        Trim(adoRegistro("CodTitulo")) & "','" & _
'                        strCodFile & "','" & _
'                        Trim(adoRegistro("CodAnalitica")) & "','" & _
'                        Trim(adoRegistro("Nemotecnico")) & "','" & _
'                        Trim(adoRegistro("CodMoneda")) & "','" & _
'                        Codigo_Moneda_Local & "'," & _
'                        dblTasaInteres & "," & _
'                        0 & "," & _
'                        dblFactorDiario & "," & _
'                        adoRegistro("SaldoFinal") & "," & _
'                        intDiasDeRenta & "," & _
'                        adoRegistro("ValorNominal") & "," & _
'                        curValorAnterior & "," & _
'                        curValorActual & "," & _
'                        curMontoRenta & "," & _
'                        0 & "," & _
'                        0 & "," & _
'                        0 & "," & _
'                        curMontoRentaContable & "," & _
'                        0 & "," & _
'                        0 & "," & _
'                        0 & ",'" & gstrCodClaseTipoCambioOperacionFondo & "'," & dblTipoCambioCierre & ",'" & adoRegistro("CodEmisor") & "') }"
'                    adoConn.Execute .CommandText
'
'                    adoConsulta.MoveNext
'                Loop
'                adoConsulta.Close: Set adoConsulta = Nothing
'
'                '*** Actualizar el número del parámetro **
'                .CommandText = "{ call up_ACActUltNumero('"
'                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNActUltNumeroTmp('"
'
'                .CommandText = .CommandText & strCodFondo & "','" & _
'                    gstrCodAdministradora & "','" & _
'                    Valor_NumComprobante & "','" & _
'                    strNumAsiento & "') }"
'
'                adoConn.Execute .CommandText
'
'                '.CommandText = "COMMIT TRAN ProcAsiento"
'                'adoConn.Execute .CommandText
'
'            End If
'
'            adoRegistro.MoveNext
'        Loop
'        adoRegistro.Close: Set adoRegistro = Nothing
                
    End With
    
    Exit Sub
       
Ctrl_Error:
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
    
End Sub

Private Sub ValorizacionCobertura(strTipoCierre As String)

    Dim adoRegistro             As ADODB.Recordset, adoConsulta     As ADODB.Recordset
    Dim dblPrecioCierre         As Double, dblPrecioPromedio        As Double
    Dim dblTirCierre            As Double, dblFactorDiario          As Double
    Dim dblTasaInteres          As Double, dblFactorDiarioCupon     As Double
    Dim dblTipoCambioSpot       As Double, dblTipoCambioFuturo      As Double
    Dim curSaldoInversion       As Currency, curSaldoInteresCorrido As Currency
    Dim curSaldoFluctuacion     As Currency, curValorAnterior       As Currency
    Dim curValorActual          As Currency, curMontoRenta          As Currency
    Dim curMontoContable        As Currency, curMontoMovimientoMN   As Currency
    Dim curMontoMovimientoME    As Currency, curSaldoValorizar      As Currency
    Dim intCantRegistros        As Integer, intContador             As Integer
    Dim intRegistro             As Integer, intBaseCalculo          As Integer
    Dim intDiasPlazo            As Integer, intDiasDeRenta          As Integer
    Dim strNumAsiento           As String, strDescripAsiento        As String
    Dim strDescripMovimiento    As String, strIndDebeHaber          As String
    Dim strCodCuenta            As String, strFiles                 As String
    Dim strCodFile              As String, strModalidadInteres      As String
    Dim strCodTasa              As String, strIndCuponCero          As String
    Dim strCodDetalleFile       As String, strNemonico              As String
    Dim strFechaGrabar          As String
    Dim dblTipoCambioCierre     As Double
    
    '*** Rentabilidad de Valores de Depósito ***
    frmMainMdi.stbMdi.Panels(3).Text = "Calculando rentabilidad de Valores Coberturados..."
       
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT IK.CodTitulo,II.CodMoneda,IK.CodAnalitica,IK.CodFile,SaldoFinal,CodDetalleFile,CodSubDetalleFile,MontoCobertura," & _
            "ValorNominal,CodTipoTasa,BaseAnual,CodTasa,TasaInteres,DiasPlazo,IndCuponCero,FechaEmision,Nemotecnico,TipoCambioSpot,TipoCambioFuturo,ICO.FechaVencimiento " & _
            "FROM InversionKardex IK JOIN InversionCobertura ICO ON(ICO.CodTitulo=IK.CodTitulo AND ICO.CodFondo=IK.CodFondo AND ICO.CodAdministradora=IK.CodAdministradora) " & _
            "JOIN InstrumentoInversion II ON(II.CodTitulo=IK.CodTitulo) " & _
            "WHERE IK.CodFile IN ('013') AND IK.FechaOperacion <='" & strFechaCierre & "' AND SaldoFinal > 0 AND " & _
            "IK.CodAdministradora='" & gstrCodAdministradora & "' AND IK.CodFondo='" & strCodFondo & "' AND IndUltimoMovimiento='X'"
        Set adoRegistro = .Execute
    
        Do Until adoRegistro.EOF
            strCodFile = Trim(adoRegistro("CodFile"))
            strCodDetalleFile = Trim(adoRegistro("CodDetalleFile"))
            strModalidadInteres = Trim(adoRegistro("CodDetalleFile"))
            strCodTasa = Trim(adoRegistro("CodTasa"))
            dblTasaInteres = CDbl(adoRegistro("TasaInteres"))
            intDiasPlazo = CInt(adoRegistro("DiasPlazo"))
            strIndCuponCero = Trim(adoRegistro("IndCuponCero"))
            strNemonico = Trim(adoRegistro("Nemotecnico"))
            curSaldoValorizar = CCur(adoRegistro("ValorNominal")) * CCur(adoRegistro("SaldoFinal"))
            dblTipoCambioSpot = CDbl(adoRegistro("TipoCambioSpot"))
            dblTipoCambioFuturo = CDbl(adoRegistro("TipoCambioFuturo"))
            intDiasDeRenta = DateDiff("d", gdatFechaActual, CVDate(adoRegistro("FechaVencimiento")))
            
            Set adoConsulta = New ADODB.Recordset
                        
            '*** Verificar Dinamica Contable ***
            .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
                "WHERE TipoOperacion='" & Codigo_Dinamica_Provision & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"

            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                If CInt(adoConsulta("NumRegistros")) > 0 Then
                    intCantRegistros = CInt(adoConsulta("NumRegistros"))
                Else
                    MsgBox "NO EXISTE Dinámica Contable para la valorización", vbCritical
                    adoConsulta.Close: Set adoConsulta = Nothing
                    Exit Sub
                End If
            End If
            adoConsulta.Close
            
            '*** Obtener las cuentas de inversión ***
            Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, Trim(adoRegistro("CodMoneda")))
            
            '*** Obtener tipo de cambio ***
            'dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
            dblTipoCambioCierre = ObtenerValorTipoCambio(adoRegistro("CodMoneda"), Codigo_Moneda_Local, strFechaCierre, strFechaCierre, gstrCodClaseTipoCambioOperacionFondo, Codigo_Valor_TipoCambioCompra)
            
            
            '*** Obtener Saldo de Provisión ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvInteres & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoFluctuacion = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
                        
            curSaldoInversion = adoRegistro("MontoCobertura")
            
            dblFactorDiario = ((dblTipoCambioFuturo / dblTipoCambioSpot) ^ (1 / 365)) - 1
            curValorAnterior = curSaldoInversion / ((1 + dblFactorDiario) ^ intDiasDeRenta) + curSaldoFluctuacion
            
            curValorActual = curSaldoValorizar * dblTipoCambioCierre
                        
            curMontoRenta = Round(curValorActual - curValorAnterior, 2)
            
            '*** Ganancia/Pérdida ***
            If curMontoRenta <> 0 Then
                'strDescripAsiento = "Valorización" & Space(1) & "(" & Trim(adoRegistro("CodFile")) & "-" & Trim(adoRegistro("CodAnalitica")) & ")"
                strDescripAsiento = "Valorización" & Space(1) & strNemonico
                strDescripMovimiento = "Pérdida"
                If curMontoRenta > 0 Then strDescripMovimiento = "Ganancia"
                                                
                .CommandType = adCmdStoredProc
                '*** Obtener el número del parámetro **
                .CommandText = "up_ACObtenerUltNumero"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "up_ACObtenerUltNumeroTmp"  '*** Simulación ***
                
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
                
                On Error GoTo Ctrl_Error
                
                '*** Contabilizar ***
                strFechaGrabar = strFechaCierre & Space(1) & Format(Time, "hh:mm")
                '*** Cabecera ***
                .CommandText = "{ call up_ACAdicAsientoContable('"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableTmp('"  '*** Simulación ***
                
                .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
                    strFechaGrabar & "','" & gstrPeriodoActual & "','" & gstrMesActual & "','" & _
                    "','" & strDescripAsiento & "','" & Trim(adoRegistro("CodMoneda")) & "','" & _
                    Codigo_Moneda_Local & "','" & "','" & "'," & _
                    CDec(curMontoRenta) & ",'" & Estado_Activo & "'," & _
                    intCantRegistros & ",'" & strFechaCierre & Space(1) & Format(Time, "hh:ss") & "','" & _
                    strCodModulo & "','" & "'," & dblTipoCambioCierre & ",'" & _
                    "','" & "','" & strDescripAsiento & "','" & "','" & "X','') }"
                adoConn.Execute .CommandText
                
                '*** Detalle ***
                .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
                    "WHERE TipoOperacion='" & Codigo_Dinamica_Provision & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                    IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"

                Set adoConsulta = .Execute
        
                Do While Not adoConsulta.EOF
                
                    Select Case Trim(adoConsulta("TipoCuentaInversion"))
                        Case Codigo_CtaInversion
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvInteres
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteres
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaCosto
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaIngresoOperacional
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresVencido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaVacCorrido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaXPagar
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaXCobrar
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresCorrido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvReajusteK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaReajusteK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvFlucMercado
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaFlucMercado
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvInteresVac
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresVac
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaIntCorridoK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvFlucK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaFlucK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInversionTransito
                            curMontoMovimientoMN = curMontoRenta
                            
                    End Select
                    
                    strIndDebeHaber = Trim(adoConsulta("IndDebeHaber"))
                    If strIndDebeHaber = "H" Then
                        curMontoMovimientoMN = curMontoMovimientoMN * -1
                        If curMontoMovimientoMN > 0 Then strIndDebeHaber = "D"
                    ElseIf strIndDebeHaber = "D" Then
                        If curMontoMovimientoMN < 0 Then strIndDebeHaber = "H"
                    End If
                    
                    If strIndDebeHaber = "T" Then
                        If curMontoMovimientoMN > 0 Then
                            strIndDebeHaber = "D"
                        Else
                            strIndDebeHaber = "H"
                        End If
                    End If
                    strDescripMovimiento = Trim(adoConsulta("DescripDinamica"))
                    curMontoMovimientoME = 0
                    curMontoContable = curMontoMovimientoMN
        
                    If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
                        curMontoContable = Round(curMontoMovimientoMN * dblTipoCambioCierre, 2)
                        curMontoMovimientoME = curMontoMovimientoMN
                        curMontoMovimientoMN = 0
                    End If
                                
                    '*** Movimiento ***
                    .CommandText = "{ call up_ACAdicAsientoContableDetalle('"
                    If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableDetalleTmp('"

                    .CommandText = .CommandText & strNumAsiento & "','" & strCodFondo & "','" & _
                        gstrCodAdministradora & "'," & _
                        CInt(adoConsulta("NumSecuencial")) & ",'" & _
                        strFechaGrabar & "','" & _
                        gstrPeriodoActual & "','" & _
                        gstrMesActual & "','" & _
                        strDescripMovimiento & "','" & _
                        strIndDebeHaber & "','" & _
                        Trim(adoConsulta("CodCuenta")) & "','" & _
                        Trim(adoRegistro("CodMoneda")) & "'," & _
                        CDec(curMontoMovimientoMN) & "," & _
                        CDec(curMontoMovimientoME) & "," & _
                        CDec(curMontoContable) & "," & _
                        Trim(adoRegistro("CodFile")) & "','" & _
                        Trim(adoRegistro("CodAnalitica")) & "','','') }"

                    adoConn.Execute .CommandText
                                    
                    '*** Validar valor de cuenta contable ***
                    If Trim(adoConsulta("CodCuenta")) = Valor_Caracter Then
                        MsgBox "Registro Nro. " & CStr(intContador) & " de Asiento Contable no tiene cuenta asignada", vbCritical, "Valorización Depósitos"
                        gblnRollBack = True
                        Exit Sub
                    End If
                    
                    adoConsulta.MoveNext
                Loop
                adoConsulta.Close: Set adoConsulta = Nothing
                                
                '*** Actualizar el número del parámetro **
                .CommandText = "{ call up_ACActUltNumero('"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNActUltNumeroTmp('"
                
                .CommandText = .CommandText & strCodFondo & "','" & _
                    gstrCodAdministradora & "','" & _
                    Valor_NumComprobante & "','" & _
                    strNumAsiento & "') }"
                    
                adoConn.Execute .CommandText
        
            End If
                                    
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
                
    End With
    
    Exit Sub
       
Ctrl_Error:
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
    
End Sub
Private Sub ValorizacionRentaFijaLargoPlazo(strTipoCierre As String)

    Dim adoRegistro             As ADODB.Recordset, adoConsulta             As ADODB.Recordset
    Dim dblPrecioCierre         As Double, dblPrecioPromedio                As Double
    Dim dblTirCierre            As Double, dblFactorDiario                  As Double
    Dim dblTasaInteres          As Double, dblFactorDiarioCupon             As Double
    Dim curSaldoInversion       As Currency, curSaldoInteresCorrido         As Currency
    Dim curSaldoFluctuacion     As Currency, curValorAnterior               As Currency
    Dim curMontoProvision       As Currency, curMontoFluctuacionMercado     As Currency
    Dim curSaldoVacCorrido      As Currency, curSaldoFluctuacionMercado     As Currency
    Dim curSaldoFluctuacionVac  As Currency, curSaldoFluctuacionReajuste    As Currency
    Dim curMontoProvisionCapital As Currency, curSaldoGPCapital             As Currency
    Dim curValorActual          As Currency, curMontoRenta                  As Currency
    Dim curMontoContable        As Currency, curMontoMovimientoMN           As Currency
    Dim curMontoMovimientoME    As Currency, curSaldoValorizar              As Currency
    Dim curMontoAjusteVAC       As Currency, curMontoInteresVAC             As Currency
    Dim intCantRegistros        As Integer, intContador                     As Integer
    Dim intRegistro             As Integer, intBaseCalculo                  As Integer
    Dim intDiasPlazo            As Integer, intDiasDeRenta                  As Integer
    Dim strNumAsiento           As String, strDescripAsiento                As String
    Dim strDescripMovimiento    As String, strIndDebeHaber                  As String
    Dim strCodCuenta            As String, strFiles                         As String
    Dim strCodFile              As String, strModalidadInteres              As String
    Dim strCodTasa              As String, strIndCuponCero                  As String
    Dim strCodDetalleFile       As String, strNemonico                      As String
    Dim strFechaGrabar          As String, strFechaEmision                  As String
    Dim strCodTipoAjuste        As String, strCodPeriodoPago                As String
    Dim strCodIndiceInicial     As String, strCodIndiceFinal                As String
    Dim strCodBaseCalculo       As String
    Dim dblTipoCambioCierre     As Double
    Dim dblValorNominalActual   As Double
    
    Dim dblValorTipoCambio          As Double, strTipoDocumento             As String
    Dim strNumDocumento             As String, strCodPersonaContraparte     As String
    Dim strTipoPersonaContraparte   As String

    Dim strIndContracuenta          As String, strCodContracuenta           As String
    Dim strCodFileContracuenta      As String, strCodAnaliticaContracuenta  As String
    Dim strIndUltimoMovimiento      As String
    
    '*** Rentabilidad de Valores de Renta Fija Largo Plazo ***
    frmMainMdi.stbMdi.Panels(3).Text = "Calculando rentabilidad de Valores de Renta Fija Largo Plazo..."
       
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT IK.CodTitulo,II.CodMoneda,IK.CodAnalitica,IK.CodFile,SaldoFinal,SaldoAmortizacion,CodDetalleFile,CodSubDetalleFile,CuponCalculo,CodTipoVac," & _
            "ValorNominal,CodTipoTasa,BaseAnual,CodTasa,TasaInteres,DiasPlazo,IndCuponCero,FechaEmision,CodTipoAjuste,TirPromedio,Nemotecnico,PeriodoPago,II.CodEmisor " & _
            "FROM InversionKardex IK JOIN InstrumentoInversion II " & _
            "ON(II.CodTitulo=IK.CodTitulo) " & _
            "WHERE IK.CodFile IN ('005') AND FechaOperacion < '" & strFechaSiguiente & "' AND SaldoFinal > 0 AND " & _
            "IK.CodAdministradora='" & gstrCodAdministradora & "' AND IK.CodFondo='" & strCodFondo & "' AND IK.NumKardex = dbo.uf_IVObtenerUltimoMovimientoKardexValor(IK.CodFondo,IK.CodAdministradora,IK.CodTitulo,'" & strFechaCierre & "')"
'            "IndUltimoMovimiento='X'"
        Set adoRegistro = .Execute
    
        Do Until adoRegistro.EOF
            strCodFile = Trim(adoRegistro("CodFile"))
            strCodDetalleFile = Trim(adoRegistro("CodDetalleFile"))
            strCodTasa = Trim(adoRegistro("CodTipoTasa"))
            strCodBaseCalculo = Trim(adoRegistro("BaseAnual"))
            dblTasaInteres = CDbl(adoRegistro("TasaInteres"))
            intDiasPlazo = CInt(adoRegistro("DiasPlazo"))
            strNemonico = Trim(adoRegistro("Nemotecnico"))
            strIndCuponCero = Trim(adoRegistro("IndCuponCero"))
            strCodIndiceFinal = Trim(adoRegistro("CuponCalculo"))
            strCodPeriodoPago = Trim(adoRegistro("PeriodoPago"))
            strCodIndiceInicial = Trim(adoRegistro("CodTipoVac"))
            strCodTipoAjuste = Trim(adoRegistro("CodTipoAjuste"))
            curSaldoValorizar = CCur(adoRegistro("SaldoFinal")) 'SaldoAmortizacion

            intBaseCalculo = 365
            Select Case strCodBaseCalculo
                Case Codigo_Base_30_360: intBaseCalculo = 360
                Case Codigo_Base_Actual_365: intBaseCalculo = 365
                Case Codigo_Base_Actual_360: intBaseCalculo = 360
                Case Codigo_Base_30_365: intBaseCalculo = 365
            End Select
            
            Set adoConsulta = New ADODB.Recordset
                        
            '*** Verificar Dinamica Contable ***
            .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
                "WHERE TipoOperacion='" & Codigo_Dinamica_Provision & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"

            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                If CInt(adoConsulta("NumRegistros")) > 0 Then
                    intCantRegistros = CInt(adoConsulta("NumRegistros"))
                Else
                    MsgBox "NO EXISTE Dinámica Contable para la valorización", vbCritical
                    adoConsulta.Close: Set adoConsulta = Nothing
                    Exit Sub
                End If
            End If
            adoConsulta.Close
                        
            dblValorNominalActual = adoRegistro("ValorNominal")
            
            .CommandText = "SELECT dbo.uf_IVObtenerValorNominalCupon('" & adoRegistro("CodTitulo") & "','" & strFechaCierre & "') AS 'ValorNominal'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                dblValorNominalActual = adoConsulta("ValorNominal")
            End If
            adoConsulta.Close
                        
                        
            '*** Obtener Ultimo Precio de Cierre registrado ***
            .CommandText = "{ call up_IVSelDatoInstrumentoInversion(2,'" & Trim(adoRegistro("CodTitulo")) & "','" & strFechaCierre & "') }"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                dblPrecioCierre = CDbl(adoConsulta("PrecioCierre"))
                dblTirCierre = CDbl(adoConsulta("TirCierre"))
                dblPrecioPromedio = CDbl(adoConsulta("PrecioPromedio"))
            End If
            adoConsulta.Close
            
            '*** Obtener el factor diario del cupón ***
            .CommandText = "SELECT FactorDiario,FechaInicio,FechaVencimiento FROM InstrumentoInversionCalendario " & _
                "WHERE CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "' AND FechaVencimiento>='" & strFechaCierre & "' ORDER BY FechaVencimiento"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                dblFactorDiarioCupon = CDbl(adoConsulta("FactorDiario"))
                intDiasDeRenta = DateDiff("d", CVDate(adoConsulta("FechaInicio")), gdatFechaActual) + 2
            End If
            adoConsulta.Close
            
            '*** Obtener las cuentas de inversión ***
            Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, adoRegistro("CodMoneda"), adoRegistro("CodSubDetalleFile"))
            
            '*** Obtener tipo de cambio ***
            'dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
            dblTipoCambioCierre = ObtenerValorTipoCambio(adoRegistro("CodMoneda"), Codigo_Moneda_Local, strFechaCierre, strFechaCierre, gstrCodClaseTipoCambioOperacionFondo, Codigo_Valor_TipoCambioCompra)
            
            
            '*** Obtener Saldo de Inversión ***
            If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaInversion & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoInversion = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Interés Corrido ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaInteresCorrido & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoInteresCorrido = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Interés VAC Corrido ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaVacCorrido & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoVacCorrido = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
                                
            '*** Obtener Saldo de Provisión ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvInteres & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoFluctuacion = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Provisión VAC***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvInteresVac & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoFluctuacionVac = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Provisión Reajuste Capital Vac ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvReajusteK & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoFluctuacionReajuste = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Provisión G/P Capital ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvFlucK & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoGPCapital = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Fluctuación Mercado ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvFlucMercado & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoFluctuacionMercado = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
                        
            curValorAnterior = curSaldoInteresCorrido + curSaldoFluctuacion
            
            If Trim(adoRegistro("CodSubDetalleFile")) <> Valor_Caracter Then strModalidadInteres = Trim(adoRegistro("CodSubDetalleFile"))
            
            '*** REVISAR ***
            '*** Cálculo de Provisión de Intereses ***

                curValorActual = CalculoInteresCorrido(adoRegistro("CodTitulo"), CDbl(curSaldoValorizar) * dblValorNominalActual, adoRegistro("FechaEmision"), DateAdd("d", 1, dtpFechaCierre.Value), strCodIndiceFinal, strCodTipoAjuste, strCodTasa, strCodPeriodoPago, strCodIndiceInicial, strCodBaseCalculo, intBaseCalculo)
                
                curMontoProvision = Round(curValorActual - curValorAnterior, 2)

                curValorAnterior = curSaldoVacCorrido + curSaldoFluctuacionReajuste + curSaldoFluctuacionVac
                
'                curValorActual = CalculoAjusteVAC(adoRegistro("CodTitulo"), CDbl(curSaldoValorizar), adoRegistro("FechaEmision"), DateAdd("d", 1, dtpFechaCierre.Value), strCodIndiceFinal, strCodTasa, strCodPeriodoPago, strCodIndiceInicial, intBaseCalculo, intDiasDeRenta)
                curValorActual = CStr(CalculoVacCorrido(adoRegistro("CodTitulo"), CDbl(curSaldoValorizar), adoRegistro("FechaEmision"), DateAdd("d", 1, dtpFechaCierre.Value), strCodIndiceFinal, strCodTipoAjuste, strCodTasa, strCodPeriodoPago, strCodIndiceInicial, intBaseCalculo))

                curMontoInteresVAC = Round(curValorActual - curValorAnterior, 2)
                

            
            '*** Cálculo Provisión G/P Capital ***
            'If strOrigen = "L" Then
                '*** VAN AL DIA ANTERIOR AL CIERRE ***
                curValorAnterior = curSaldoInversion + curSaldoInteresCorrido + curSaldoFluctuacion + curSaldoFluctuacionVac + curSaldoFluctuacionReajuste + curSaldoVacCorrido + curSaldoGPCapital
    
                '*** CALCULO DEL VAN A LA FECHA DE CIERRE ***
                If CDbl(adoRegistro("TirPromedio")) <> 0 Then
                    Dim datFechaGP  As Date, datFechaSiguienteGP    As Date
                    Dim dblValorTir As Double
    
                    'vntFeccie = CVar(Right$(strFeccie, 2) + "/" + Mid$(strFeccie, 5, 2) + "/" + Left$(strFeccie, 4))
                    'vntFeccieMas1Dia = CVar(Right$(strFechaSiguiente, 2) + "/" + Mid$(strFechaSiguiente, 5, 2) + "/" + Left$(strFechaSiguiente, 4))
                    datFechaGP = Convertddmmyyyy(strFechaCierre)
                    datFechaSiguienteGP = Convertddmmyyyy(strFechaSiguiente)
    
                    dblValorTir = CDbl(adoRegistro("TirPromedio"))
                    
                    curValorActual = VNANoPer(adoRegistro("CodTitulo"), datFechaSiguienteGP, datFechaSiguienteGP, curSaldoValorizar * dblValorNominalActual, curSaldoValorizar, dblValorTir, adoRegistro("CodTipoAjuste"), strCodIndiceInicial, strCodIndiceFinal)
    
                    '*** CALCULO DEL MONTO DE GANANCIA/PERDIDA DE curCapital ***
                    curMontoProvisionCapital = Round(curValorActual - curValorAnterior - curMontoProvision, 2)
                Else
                    curMontoProvisionCapital = 0
                End If
    
            'End If
            
            '*** Cálculo Fluctuación Mercado ***
            If dblTirCierre > 0 Then
                curValorAnterior = curSaldoInversion + curSaldoInteresCorrido + curSaldoFluctuacion + curSaldoFluctuacionVac + curSaldoFluctuacionReajuste + curSaldoVacCorrido + curSaldoGPCapital + curSaldoFluctuacionMercado
                
                curValorActual = VNANoPer(adoRegistro("CodTitulo"), datFechaSiguienteGP, datFechaSiguienteGP, curSaldoValorizar * dblValorNominalActual, curSaldoValorizar, dblTirCierre, adoRegistro("CodTipoAjuste"), strCodIndiceInicial, strCodIndiceFinal)
             
                curMontoFluctuacionMercado = Round(curValorActual - curValorAnterior - curMontoProvision, 2)
                'inicio agregado ACR: 16/11/2012
                curMontoProvisionCapital = 0
                'fin agregado ACR: 16/11/2012
            Else
                curMontoFluctuacionMercado = 0
                'inicio agregado ACR: 16/11/2012
                curMontoFluctuacionMercado = curMontoProvisionCapital
                curMontoProvisionCapital = 0
                'fin agregado ACR: 16/11/2012
            End If
                        
            '*** Contabilización ***
            If curMontoProvision <> 0 Or curMontoProvisionCapital <> 0 Or curMontoFluctuacionMercado <> 0 Then
                strDescripAsiento = "Valorización" & Space(1) & strNemonico
                strDescripMovimiento = "Pérdida"
                If curMontoProvision > 0 Then strDescripMovimiento = "Ganancia"
                                                
                .CommandType = adCmdStoredProc
                '*** Obtener el número del parámetro **
                .CommandText = "up_ACObtenerUltNumero"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "up_ACObtenerUltNumeroTmp"  '*** Simulación ***
                
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
                
'                .CommandText = "BEGIN TRAN ProcAsiento"
'                adoConn.Execute .CommandText
                
                'On Error GoTo Ctrl_Error
                
                '*** Contabilizar ***
                strFechaGrabar = strFechaCierre & Space(1) & Format(Time, "hh:mm")
                
                '*** Cabecera ***
                .CommandText = "{ call up_ACAdicAsientoContable('"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableTmp('"  '*** Simulación ***
                
                .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
                    strFechaGrabar & "','" & _
                    gstrPeriodoActual & "','" & gstrMesActual & "','" & _
                    "','" & _
                    strDescripAsiento & "','" & Trim(adoRegistro("CodMoneda")) & "','" & _
                    Codigo_Moneda_Local & "','',''," & _
                    CDec(curMontoProvision + curMontoProvisionCapital + curMontoFluctuacionMercado) & ",'" & Estado_Activo & "'," & _
                    intCantRegistros & ",'" & _
                    strFechaCierre & Space(1) & Format(Time, "hh:ss") & "','" & _
                    strCodModulo & "',''," & _
                    dblTipoCambioCierre & ",'" & _
                    "','','" & _
                    strDescripAsiento & "','" & _
                    "','X','') }"
                adoConn.Execute .CommandText
                
                If curMontoProvision < 0 Then curMontoProvision = 0
                
                '*** Detalle ***
                .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
                    "WHERE TipoOperacion='" & Codigo_Dinamica_Provision & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodSubDetalleFile = '" & adoRegistro("CodSubDetalleFile") & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                    IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"
                Set adoConsulta = .Execute
        
                Do While Not adoConsulta.EOF
                    
                    curMontoMovimientoMN = 0
                
                    Select Case Trim(adoConsulta("TipoCuentaInversion"))
                        Case Codigo_CtaInversion
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvInteres
                            curMontoMovimientoMN = curMontoProvision
'                            strDescripMovimiento = "Pérdida"
'                            If curMontoProvision > 0 Then strDescripMovimiento = "Ganancia"
                            
                        Case Codigo_CtaInteres
                            curMontoMovimientoMN = curMontoProvision
                            
                        Case Codigo_CtaCosto
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaIngresoOperacional
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresVencido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaVacCorrido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaXPagar
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaXCobrar
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresCorrido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvReajusteK
                            curMontoMovimientoMN = curMontoAjusteVAC
                            
                        Case Codigo_CtaReajusteK
                            curMontoMovimientoMN = curMontoAjusteVAC
                            
                        Case Codigo_CtaProvInteresVac
                            curMontoMovimientoMN = curMontoInteresVAC
                            
                        Case Codigo_CtaInteresVac
                            curMontoMovimientoMN = curMontoInteresVAC
                            
                        Case Codigo_CtaIntCorridoK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvFlucK
                            curMontoMovimientoMN = curMontoProvisionCapital
'                            strDescripMovimiento = "Pérdida"
'                            If curMontoProvisionCapital > 0 Then strDescripMovimiento = "Ganancia"
                            
                        Case Codigo_CtaFlucK
                            curMontoMovimientoMN = curMontoProvisionCapital
                            
                        Case Codigo_CtaInversionTransito
                            curMontoMovimientoMN = curMontoRenta
                            
                        'ACR: 16/11/2012
                        Case Codigo_CtaProvFlucMercado
                            If curMontoFluctuacionMercado > 0 Then curMontoMovimientoMN = curMontoFluctuacionMercado
                            
                        Case Codigo_CtaFlucMercado
                            If curMontoFluctuacionMercado > 0 Then curMontoMovimientoMN = curMontoFluctuacionMercado
                            
                        Case Codigo_CtaProvFlucMercado_Perdida
                            If curMontoFluctuacionMercado < 0 Then curMontoMovimientoMN = curMontoFluctuacionMercado
                            
                        Case Codigo_CtaFlucMercado_Perdida
                            If curMontoFluctuacionMercado < 0 Then curMontoMovimientoMN = curMontoFluctuacionMercado
                        'ACR: 16/11/2012
                            
                    End Select
                                                        
                    strIndDebeHaber = Trim(adoConsulta("IndDebeHaber"))
                    
                    If strIndDebeHaber = "H" Then
                        'curMontoMovimientoMN = curMontoMovimientoMN * -1
                        If curMontoMovimientoMN > 0 Then
                            curMontoMovimientoMN = curMontoMovimientoMN * -1
                        End If
                    ElseIf strIndDebeHaber = "D" Then
                        If curMontoMovimientoMN < 0 Then
                            curMontoMovimientoMN = curMontoMovimientoMN * -1
                        End If
                    End If
                    
                    If strIndDebeHaber = "T" Then
                        If curMontoMovimientoMN > 0 Then
                            strIndDebeHaber = "D"
                        Else
                            strIndDebeHaber = "H"
                        End If
                    End If
                    
                    strDescripMovimiento = Trim(adoConsulta("DescripDinamica"))
                    curMontoMovimientoME = 0
                    curMontoContable = curMontoMovimientoMN
        
                    dblValorTipoCambio = 1
                    
                    If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
                        curMontoContable = Round(curMontoMovimientoMN * dblTipoCambioCierre, 2)
                        curMontoMovimientoME = curMontoMovimientoMN
                        curMontoMovimientoMN = 0
                        dblValorTipoCambio = dblTipoCambioCierre
                    End If
                    
                    strTipoDocumento = ""
                    strNumDocumento = ""
                    strTipoPersonaContraparte = ""
                    strCodPersonaContraparte = ""
                    strIndContracuenta = ""
                    strCodContracuenta = ""
                    strCodFileContracuenta = ""
                    strCodAnaliticaContracuenta = ""
                    strIndUltimoMovimiento = ""
                                
                    If curMontoContable <> 0 Then
                        '*** Movimiento ***
                        .CommandText = "{ call up_ACAdicAsientoContableDetalle('"
                        If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableDetalleTmp('"
    
                        .CommandText = .CommandText & strNumAsiento & "','" & strCodFondo & "','" & _
                            gstrCodAdministradora & "'," & _
                            CInt(adoConsulta("NumSecuencial")) & ",'" & _
                            strFechaGrabar & "','" & _
                            gstrPeriodoActual & "','" & _
                            gstrMesActual & "','" & _
                            strDescripMovimiento & "','" & _
                            strIndDebeHaber & "','" & _
                            Trim(adoConsulta("CodCuenta")) & "','" & _
                            Trim(adoRegistro("CodMoneda")) & "'," & _
                            CDec(curMontoMovimientoMN) & "," & _
                            CDec(curMontoMovimientoME) & "," & _
                            CDec(curMontoContable) & "," & _
                            dblValorTipoCambio & ",'" & _
                            Trim(adoRegistro("CodFile")) & "','" & _
                            Trim(adoRegistro("CodAnalitica")) & "','" & _
                            strTipoDocumento & "','" & _
                            strNumDocumento & "','" & _
                            strTipoPersonaContraparte & "','" & _
                            strCodPersonaContraparte & "','" & _
                            strIndContracuenta & "','" & _
                            strCodContracuenta & "','" & _
                            strCodFileContracuenta & "','" & _
                            strCodAnaliticaContracuenta & "','" & _
                            strIndUltimoMovimiento & "') }"
    
                        adoConn.Execute .CommandText
                    
                        '*** Validar valor de cuenta contable ***
                        If Trim(adoConsulta("CodCuenta")) = Valor_Caracter Then
                            MsgBox "Registro Nro. " & CStr(intContador) & " de Asiento Contable no tiene cuenta asignada", vbCritical, "Valorización Depósitos"
                            gblnRollBack = True
                            Exit Sub
                        End If
                    
                    End If
                    
                    adoConsulta.MoveNext
                Loop
                adoConsulta.Close: Set adoConsulta = Nothing
                                
                '*** Actualizar el número del parámetro **
                .CommandText = "{ call up_ACActUltNumero('"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNActUltNumeroTmp('"
                
                .CommandText = .CommandText & strCodFondo & "','" & _
                    gstrCodAdministradora & "','" & _
                    Valor_NumComprobante & "','" & _
                    strNumAsiento & "') }"
                    
                adoConn.Execute .CommandText
        
            End If
                                    
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
                
    End With
    
    Exit Sub
       
Ctrl_Error:

    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault

   
End Sub

Private Sub ValorizacionPacto(strTipoCierre As String)

    Dim adoRegistro             As ADODB.Recordset, adoConsulta             As ADODB.Recordset
    Dim dblPrecioCierre         As Double, dblPrecioPromedio                As Double
    Dim dblTirCierre            As Double, dblFactorDiario                  As Double
    Dim dblTasaInteres          As Double, dblFactorDiarioCupon             As Double
    Dim curSaldoInversion       As Currency, curSaldoInteresCorrido         As Currency
    Dim curSaldoFluctuacion     As Currency, curValorAnterior               As Currency
    Dim curMontoProvision       As Currency, curMontoFluctuacionMercado     As Currency
    Dim curSaldoVacCorrido      As Currency, curSaldoFluctuacionMercado     As Currency
    Dim curSaldoFluctuacionVac  As Currency, curSaldoFluctuacionReajuste    As Currency
    Dim curMontoProvisionCapital As Currency, curSaldoGPCapital             As Currency
    Dim curValorActual          As Currency, curMontoRenta                  As Currency
    Dim curMontoContable        As Currency, curMontoMovimientoMN           As Currency
    Dim curMontoMovimientoME    As Currency, curSaldoValorizar              As Currency
    Dim intCantRegistros        As Integer, intContador                     As Integer
    Dim intRegistro             As Integer, intBaseCalculo                  As Integer
    Dim intDiasPlazo            As Integer, intDiasDeRenta                  As Integer
    Dim strNumAsiento           As String, strDescripAsiento                As String
    Dim strDescripMovimiento    As String, strIndDebeHaber                  As String
    Dim strCodCuenta            As String, strFiles                         As String
    Dim strCodFile              As String, strModalidadInteres              As String
    Dim strCodTasa              As String, strIndCuponCero                  As String
    Dim strCodDetalleFile       As String, strFechaGrabar          As String
    Dim dblTipoCambioCierre     As Double
    
    
    '*** Rentabilidad de Valores de Renta Fija Largo Plazo ***
    frmMainMdi.stbMdi.Panels(3).Text = "Calculando rentabilidad de Pactos..."
       
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT IK.CodTitulo,II.CodMoneda,IK.CodAnalitica,IK.CodFile,SaldoFinal,CodDetalleFile,CodSubDetalleFile,FechaVencimiento," & _
            "ValorNominal,CodTipoTasa,BaseAnual,CodTasa,TasaInteres,DiasPlazo,IndCuponCero,FechaEmision,CodTipoVac,TirPromedio,ValorMFL1,ValorMFL2,CodGarantia " & _
            "FROM InversionKardex IK JOIN InstrumentoInversion II " & _
            "ON(II.CodTitulo=IK.CodTitulo) " & _
            "WHERE IK.CodFile IN ('009') AND FechaOperacion <='" & strFechaCierre & "' AND SaldoFinal > 0 AND " & _
            "IK.CodAdministradora='" & gstrCodAdministradora & "' AND IK.CodFondo='" & strCodFondo & "' AND IndUltimoMovimiento='X'"
            
        Set adoRegistro = .Execute
    
        Do Until adoRegistro.EOF
            strCodFile = Trim(adoRegistro("CodFile"))
            strCodDetalleFile = Trim(adoRegistro("CodDetalleFile"))
            strCodTasa = Trim(adoRegistro("CodTasa"))
            dblTasaInteres = CDbl(adoRegistro("TasaInteres"))
            intDiasPlazo = CInt(adoRegistro("DiasPlazo"))
            strIndCuponCero = Trim(adoRegistro("IndCuponCero"))
            curSaldoValorizar = CCur(adoRegistro("SaldoFinal"))
            
            Set adoConsulta = New ADODB.Recordset
                        
            '*** Verificar Dinamica Contable ***
            .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
                "WHERE TipoOperacion='" & Codigo_Dinamica_Provision & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                If CInt(adoConsulta("NumRegistros")) > 0 Then
                    intCantRegistros = CInt(adoConsulta("NumRegistros"))
                Else
                    MsgBox "NO EXISTE Dinámica Contable para la valorización", vbCritical
                    adoConsulta.Close: Set adoConsulta = Nothing
                    Exit Sub
                End If
            End If
            adoConsulta.Close
                        
            '*** Obtener Ultimo Precio de Cierre registrado ***
            .CommandText = "{ call up_IVSelDatoInstrumentoInversion(2,'" & _
                Trim(adoRegistro("CodTitulo")) & "') }"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                dblPrecioCierre = CDbl(adoConsulta("PrecioCierre"))
                dblTirCierre = CDbl(adoConsulta("TirCierre"))
                dblPrecioPromedio = CDbl(adoConsulta("PrecioPromedio"))
            End If
            adoConsulta.Close
            
            '*** Obtener el factor diario del cupón ***
            .CommandText = "SELECT FactorDiario,FechaInicio,FechaVencimiento FROM InstrumentoInversionCalendario " & _
                "WHERE CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "' AND IndVigente='X'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                dblFactorDiarioCupon = CDbl(adoConsulta("FactorDiario"))
                intDiasDeRenta = DateDiff("d", CVDate(adoConsulta("FechaInicio")), gdatFechaActual) + 1
            End If
            adoConsulta.Close
            
            '*** Obtener las cuentas de inversión ***
            Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, Trim(adoRegistro("CodMoneda")))
            
            '*** Obtener tipo de cambio ***
            'dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
            dblTipoCambioCierre = ObtenerValorTipoCambio(adoRegistro("CodMoneda"), Codigo_Moneda_Local, strFechaCierre, strFechaCierre, gstrCodClaseTipoCambioOperacionFondo, Codigo_Valor_TipoCambioCompra)
            
            
            '*** Obtener Saldo de Inversión ***
            If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaInversion & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoInversion = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Interés Corrido ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaInteresCorrido & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoInteresCorrido = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Interés VAC Corrido ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaVacCorrido & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoVacCorrido = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
                                
            '*** Obtener Saldo de Provisión ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvInteres & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoFluctuacion = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Provisión VAC***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvInteresVac & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoFluctuacionVac = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Provisión Reajuste Capital Vac ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvReajusteK & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoFluctuacionReajuste = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Provisión G/P Capital ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvFlucK & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoGPCapital = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Fluctuación Mercado ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvFlucMercado & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoFluctuacionMercado = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
                        
            curValorAnterior = curSaldoInteresCorrido + curSaldoFluctuacion
            
            If adoRegistro("BaseAnual") = Codigo_Base_Actual_365 Or adoRegistro("BaseAnual") = Codigo_Base_30_365 Or adoRegistro("BaseAnual") = Codigo_Base_Actual_Actual Then
                intBaseCalculo = 365
            Else
                intBaseCalculo = 360
            End If
            
            If Trim(adoRegistro("CodSubDetalleFile")) <> Valor_Caracter Then strModalidadInteres = Trim(adoRegistro("CodSubDetalleFile"))
            
            '*** REVISAR ***
            '*** Cálculo de Provisión de Intereses ***
            If Trim(adoRegistro("CodTipoVac")) = Codigo_Tipo_Ajuste_Vac Then
                If strCodTasa = Codigo_Tipo_Tasa_Efectiva Then
                    dblFactorDiario = ((1 + CDbl(((1 + (dblTasaInteres / 100) ^ (intDiasPlazo / intBaseCalculo))) - 1)) ^ (1 / intDiasPlazo)) - 1
                    curValorActual = (curSaldoValorizar * (((1 + dblFactorDiario) ^ (intDiasDeRenta)) - 1))
                Else
                    dblFactorDiario = (CDbl(((1 + (dblTasaInteres / 100)) / intBaseCalculo)))
                    curValorActual = (curSaldoValorizar * (dblFactorDiario * intDiasDeRenta))
                End If
                curValorAnterior = 0
                curValorActual = 0

            Else

                    If strCodTasa = Codigo_Tipo_Tasa_Efectiva Then
                        dblFactorDiario = ((1 + CDbl(((1 + (dblTasaInteres / 100)) ^ (intDiasPlazo / intBaseCalculo)) - 1)) ^ (1 / intDiasPlazo)) - 1
                    Else
                        dblFactorDiario = (CDbl(((1 + (dblTasaInteres / 100)) / intBaseCalculo)))
                    End If
                'End If
                If strCodTasa = Codigo_Tipo_Tasa_Efectiva Then
                    curValorActual = (curSaldoValorizar * (((1 + dblFactorDiario) ^ (intDiasDeRenta)) - 1))
                Else
                    curValorActual = (curSaldoValorizar * (dblFactorDiario * intDiasDeRenta))
                End If
            End If
                        
            curMontoProvision = Round(curValorActual - curValorAnterior, 2)
            
            '*** Cálculo Provisión G/P Capital ***
            'If strOrigen = "L" Then
                '*** VAN AL DIA ANTERIOR AL CIERRE ***
                curValorAnterior = curSaldoInversion + curSaldoInteresCorrido + curSaldoFluctuacion + curSaldoFluctuacionVac + curSaldoFluctuacionReajuste + curSaldoVacCorrido + curSaldoGPCapital
    
                '*** CALCULO DEL VAN A LA FECHA DE CIERRE ***
                If CDbl(adoRegistro("TirPromedio")) <> 0 Then
                    Dim datFechaGP  As Date, datFechaSiguienteGP    As Date
                    Dim dblValorTir As Double
    
                    datFechaGP = Convertddmmyyyy(strFechaCierre)
                    datFechaSiguienteGP = Convertddmmyyyy(strFechaSiguiente)
    
                    dblValorTir = CDbl(adoRegistro("TirPromedio"))
                    
                    curValorActual = VNANoPerPlazo(adoRegistro("CodGarantia"), datFechaSiguienteGP, datFechaSiguienteGP, adoRegistro("FechaVencimiento"), curSaldoValorizar, curSaldoValorizar, dblValorTir, adoRegistro("CodTipoVac"), adoRegistro("ValorMFL2"))
    
                    '*** CALCULO DEL MONTO DE GANANCIA/PERDIDA DE curCapital ***
                    curMontoProvisionCapital = Round(curValorActual - curValorAnterior - curMontoProvision, 2)
                Else
                    curMontoProvisionCapital = 0
                End If
    


            curMontoProvision = curMontoProvision + curMontoProvisionCapital
                        
            '*** Contabilización ***
            If curMontoProvision <> 0 Then
                strDescripAsiento = "Valorización" & Space(1) & "(" & Trim(adoRegistro("CodFile")) & "-" & Trim(adoRegistro("CodAnalitica")) & ")"
                strDescripMovimiento = "Pérdida"
                If curMontoProvision > 0 Then strDescripMovimiento = "Ganancia"
                                                
                .CommandType = adCmdStoredProc
                '*** Obtener el número del parámetro **
                .CommandText = "up_ACObtenerUltNumero"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "up_ACObtenerUltNumeroTmp"  '*** Simulación ***
                
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
                
                .CommandText = "BEGIN TRAN ProcAsiento"
                adoConn.Execute .CommandText
                
                On Error GoTo Ctrl_Error
                
                '*** Contabilizar ***
                strFechaGrabar = strFechaCierre & Space(1) & Format(Time, "hh:mm")
                
                '*** Cabecera ***
                .CommandText = "{ call up_ACAdicAsientoContable('"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableTmp('"  '*** Simulación ***
                
                .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
                    strFechaGrabar & "','" & _
                    gstrPeriodoActual & "','" & gstrMesActual & "','" & _
                    "','" & _
                    strDescripAsiento & "','" & Trim(adoRegistro("CodMoneda")) & "','" & _
                    Codigo_Moneda_Local & "','',''," & _
                    CDec(curMontoProvision + curMontoProvisionCapital + curMontoFluctuacionMercado) & ",'" & Estado_Activo & "'," & _
                    intCantRegistros & ",'" & _
                    strFechaCierre & Space(1) & Format(Time, "hh:ss") & "','" & _
                    strCodModulo & "',''," & _
                    dblTipoCambioCierre & ",'" & _
                    "','','" & _
                    strDescripAsiento & "','" & _
                    "','X','') }"
                adoConn.Execute .CommandText
                
                '*** Detalle ***
                .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
                    "WHERE TipoOperacion='" & Codigo_Dinamica_Provision & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                    IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"
                Set adoConsulta = .Execute
        
                Do While Not adoConsulta.EOF
                
                    Select Case Trim(adoConsulta("TipoCuentaInversion"))
                        Case Codigo_CtaInversion
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvInteres
                            curMontoMovimientoMN = curMontoProvision
                            strDescripMovimiento = "Pérdida"
                            If curMontoProvision > 0 Then strDescripMovimiento = "Ganancia"
                            
                        Case Codigo_CtaInteres
                            curMontoMovimientoMN = curMontoProvision
                            
                        Case Codigo_CtaCosto
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaIngresoOperacional
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresVencido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaVacCorrido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaXPagar
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaXCobrar
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresCorrido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvReajusteK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaReajusteK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvFlucMercado
                            curMontoMovimientoMN = curMontoFluctuacionMercado
                            strDescripMovimiento = "Pérdida"
                            If curMontoFluctuacionMercado > 0 Then strDescripMovimiento = "Ganancia"
                            
                        Case Codigo_CtaFlucMercado
                            curMontoMovimientoMN = curMontoFluctuacionMercado
                            
                        Case Codigo_CtaProvInteresVac
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresVac
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaIntCorridoK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvFlucK
                            curMontoMovimientoMN = curMontoProvisionCapital
                            strDescripMovimiento = "Pérdida"
                            If curMontoProvisionCapital > 0 Then strDescripMovimiento = "Ganancia"
                            
                        Case Codigo_CtaFlucK
                            curMontoMovimientoMN = curMontoProvisionCapital
                            
                        Case Codigo_CtaInversionTransito
                            curMontoMovimientoMN = curMontoRenta
                            
                    End Select
                
                    strIndDebeHaber = Trim(adoConsulta("IndDebeHaber"))
                    
                    If strIndDebeHaber = "H" Then
                        curMontoMovimientoMN = curMontoMovimientoMN * -1
                        If curMontoMovimientoMN > 0 Then strIndDebeHaber = "D"
                    ElseIf strIndDebeHaber = "D" Then
                        If curMontoMovimientoMN < 0 Then strIndDebeHaber = "H"
                    End If
                    
                    If strIndDebeHaber = "T" Then
                        If curMontoMovimientoMN > 0 Then
                            strIndDebeHaber = "D"
                        Else
                            strIndDebeHaber = "H"
                        End If
                    End If
                    strDescripMovimiento = Trim(adoConsulta("DescripDinamica"))
                    curMontoMovimientoME = 0
                    curMontoContable = curMontoMovimientoMN
        
                    If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
                        curMontoContable = Round(curMontoMovimientoMN * dblTipoCambioCierre, 2)
                        curMontoMovimientoME = curMontoMovimientoMN
                        curMontoMovimientoMN = 0
                    End If
                                
                    '*** Movimiento ***
                    .CommandText = "{ call up_ACAdicAsientoContableDetalle('"
                    If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableDetalleTmp('"

                    .CommandText = .CommandText & strNumAsiento & "','" & strCodFondo & "','" & _
                        gstrCodAdministradora & "'," & _
                        CInt(adoConsulta("NumSecuencial")) & ",'" & _
                        strFechaGrabar & "','" & _
                        gstrPeriodoActual & "','" & _
                        gstrMesActual & "','" & _
                        strDescripMovimiento & "','" & _
                        strIndDebeHaber & "','" & _
                        Trim(adoConsulta("CodCuenta")) & "','" & _
                        Trim(adoRegistro("CodMoneda")) & "'," & _
                        CDec(curMontoMovimientoMN) & "," & _
                        CDec(curMontoMovimientoME) & "," & _
                        CDec(curMontoContable) & ",'" & _
                        Trim(adoRegistro("CodFile")) & "','" & _
                        Trim(adoRegistro("CodAnalitica")) & "','','') }"

                    adoConn.Execute .CommandText
                
                                    
                    '*** Validar valor de cuenta contable ***
                    If Trim(adoConsulta("CodCuenta")) = Valor_Caracter Then
                        MsgBox "Registro Nro. " & CStr(intContador) & " de Asiento Contable no tiene cuenta asignada", vbCritical, "Valorización Depósitos"
                        gblnRollBack = True
                        Exit Sub
                    End If
                    
                    adoConsulta.MoveNext
                Loop
                adoConsulta.Close: Set adoConsulta = Nothing
                                
                '*** Actualizar el número del parámetro **
                .CommandText = "{ call up_ACActUltNumero('"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNActUltNumeroTmp('"
                
                .CommandText = .CommandText & strCodFondo & "','" & _
                    gstrCodAdministradora & "','" & _
                    Valor_NumComprobante & "','" & _
                    strNumAsiento & "') }"
                    
                adoConn.Execute .CommandText
                                
                .CommandText = "COMMIT TRAN ProcAsiento"
                adoConn.Execute .CommandText
        
            End If
                                    
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
                
    End With
    
    Exit Sub
       
Ctrl_Error:
    adoComm.CommandText = "ROLLBACK TRAN ProcAsiento"
    adoConn.Execute adoComm.CommandText
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
   
End Sub

Private Sub RenLetHipotecarias()

   
End Sub

Private Sub RenLetPag()
    
End Sub


Private Sub RenRenFija()

End Sub

Private Sub ValorizacionReportes(strTipoCierre As String)

    Dim adoRegistro             As ADODB.Recordset, adoConsulta     As ADODB.Recordset
    Dim dblPrecioCierre         As Double, dblPrecioPromedio        As Double
    Dim dblTirCierre            As Double, dblFactorDiario          As Double
    Dim dblTasaInteres          As Double, dblFactorDiarioCupon     As Double
    Dim curSaldoInversion       As Currency, curSaldoInteresCorrido As Currency
    Dim curSaldoFluctuacion     As Currency, curValorAnterior       As Currency
    Dim curValorActual          As Currency, curMontoRenta          As Currency
    Dim curMontoContable        As Currency, curMontoMovimientoMN   As Currency
    Dim curMontoMovimientoME    As Currency, curSaldoValorizar      As Currency
    Dim intCantRegistros        As Integer, intContador             As Integer
    Dim intRegistro             As Integer, intBaseCalculo          As Integer
    Dim intDiasPlazo            As Integer, intDiasDeRenta          As Integer
    Dim strNumAsiento           As String, strDescripAsiento        As String
    Dim strDescripMovimiento    As String, strIndDebeHaber          As String
    Dim strCodCuenta            As String, strFiles                 As String
    Dim strCodFile              As String, strModalidadInteres      As String
    Dim strCodTasa              As String, strIndCuponCero          As String
    Dim strCodDetalleFile       As String, strNemonico              As String
    Dim strFechaGrabar          As String
    Dim dblTipoCambioCierre     As Double, strCtaCostoInversion     As String
    Dim curMontoRentaContable   As Currency
    
    Dim dblValorTipoCambio          As Double, strTipoDocumento As String
    Dim strNumDocumento             As String, strCodPersonaContraparte As String
    Dim strTipoPersonaContraparte   As String
    
    Dim strIndContracuenta          As String, strCodContracuenta           As String
    Dim strCodFileContracuenta      As String, strCodAnaliticaContracuenta  As String
    Dim strIndUltimoMovimiento      As String
    
    
    '*** Rentabilidad de Reportes ***
    frmMainMdi.stbMdi.Panels(3).Text = "Calculando rentabilidad de Operaciones de Reporte..."
       
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT IK.CodTitulo,II.CodMoneda,IK.CodAnalitica,IK.CodFile,SaldoFinal,MontoSaldo,CodDetalleFile,CodSubDetalleFile," & _
            "ValorNominal,CodTipoTasa,BaseAnual,CodTasa,TasaInteres,DiasPlazo,IndCuponCero,FechaEmision,Nemotecnico,II.CodEmisor " & _
            "FROM InversionKardex IK LEFT JOIN InstrumentoInversion II " & _
            "ON(II.CodFondo=IK.CodFondo AND II.CodAdministradora=IK.CodAdministradora AND II.CodTitulo=IK.CodTitulo) " & _
            "WHERE IK.CodFile IN ('008') AND FechaOperacion <='" & strFechaCierre & "' AND SaldoFinal > 0 AND " & _
            "IK.CodAdministradora='" & gstrCodAdministradora & "' AND IK.CodFondo='" & strCodFondo & "' AND IndUltimoMovimiento='X'"
            
        Set adoRegistro = .Execute
    
        Do Until adoRegistro.EOF
            strCodFile = Trim(adoRegistro("CodFile"))
            strCodDetalleFile = Trim(adoRegistro("CodDetalleFile"))
            strNemonico = Trim(adoRegistro("Nemotecnico"))
            strModalidadInteres = Trim(adoRegistro("CodDetalleFile"))
            strCodTasa = Trim(adoRegistro("CodTasa"))
            dblTasaInteres = CDbl(adoRegistro("TasaInteres"))
            intDiasPlazo = CInt(adoRegistro("DiasPlazo"))
            strIndCuponCero = Trim(adoRegistro("IndCuponCero"))
            curSaldoValorizar = CCur(adoRegistro("MontoSaldo"))
            intDiasDeRenta = DateDiff("d", CVDate(adoRegistro("FechaEmision")), gdatFechaActual) + 1
            
            dblPrecioCierre = 0
            dblTirCierre = 0
            dblPrecioPromedio = 0
            curSaldoInversion = 0
            curSaldoInteresCorrido = 0
            curSaldoFluctuacion = 0
            
            
            Set adoConsulta = New ADODB.Recordset
                        
            '*** Verificar Dinamica Contable ***
            .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
                "WHERE TipoOperacion='" & Codigo_Dinamica_Provision & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"

            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                If CInt(adoConsulta("NumRegistros")) > 0 Then
                    intCantRegistros = CInt(adoConsulta("NumRegistros"))
                Else
                    MsgBox "NO EXISTE Dinámica Contable para la valorización", vbCritical
                    adoConsulta.Close: Set adoConsulta = Nothing
                    Exit Sub
                End If
            End If
            adoConsulta.Close
                        
            '*** Obtener Ultimo Precio de Cierre registrado ***
            .CommandText = "{ call up_IVSelDatoInstrumentoInversion(2,'" & Trim(adoRegistro("CodTitulo")) & "','" & strFechaCierre & "') }"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                dblPrecioCierre = CDbl(adoConsulta("PrecioCierre"))
                dblTirCierre = CDbl(adoConsulta("TirCierre"))
                dblPrecioPromedio = CDbl(adoConsulta("PrecioPromedio"))
            End If
            adoConsulta.Close
            
            '*** Obtener el factor diario del cupón ***
            .CommandText = "SELECT FactorDiario, ValorInteres + ValorAmortizacion AS SaldoValorizar FROM InstrumentoInversionCalendario " & _
                "WHERE CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                dblFactorDiario = CDbl(adoConsulta("FactorDiario"))
                curSaldoValorizar = CDbl(adoConsulta("SaldoValorizar")) 'MFL2
            End If
            adoConsulta.Close
            
            '*** Obtener las cuentas de inversión ***
            Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, Trim(adoRegistro("CodMoneda")))
            
            '*** Obtener tipo de cambio ***
            'dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
            dblTipoCambioCierre = ObtenerValorTipoCambio(adoRegistro("CodMoneda"), Codigo_Moneda_Local, strFechaCierre, strFechaCierre, gstrCodClaseTipoCambioOperacionFondo, Codigo_Valor_TipoCambioCompra)
            
            
            '*** Obtener Saldo de Inversión ***
            If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                .CommandText = "SELECT ISNULL(SUM(SaldoFinalContable),0) AS Saldo "
            Else
                .CommandText = "SELECT ISNULL(SUM(SaldoFinalME),0) AS Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            strCtaCostoInversion = "('" & strCtaInversion & "','" & strCtaInversionCostoSAB & "','" & _
                                   strCtaInversionCostoBVL & "','" & strCtaInversionCostoCavali & "','" & _
                                   strCtaInversionCostoFondoGarantia & "','" & strCtaInversionCostoConasev & "','" & _
                                   strCtaInversionCostoIGV & "','" & strCtaInversionCostoCompromiso & "','" & _
                                   strCtaInversionCostoResponsabilidad & "','" & strCtaInversionCostoFondoLiquidacion & "','" & _
                                   strCtaInversionCostoComisionEspecial & "','" & strCtaInversionCostoGastosBancarios & "')"
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta IN " & strCtaCostoInversion & " AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoInversion = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Interés Corrido ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT ISNULL(SaldoFinalContable,0) Saldo "
            Else
                .CommandText = "SELECT ISNULL(SaldoFinalME,0) Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaInteresCorrido & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoInteresCorrido = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Provisión ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT ISNULL(SaldoFinalContable,0) Saldo "
            Else
                .CommandText = "SELECT ISNULL(SaldoFinalME,0) Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvInteres & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoFluctuacion = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
                        
            curValorAnterior = curSaldoInversion + curSaldoInteresCorrido + curSaldoFluctuacion
            
            If adoRegistro("BaseAnual") = Codigo_Base_Actual_365 Or adoRegistro("BaseAnual") = Codigo_Base_30_365 Or adoRegistro("BaseAnual") = Codigo_Base_Actual_Actual Then
                intBaseCalculo = 365
            Else
                intBaseCalculo = 360
            End If
            
            If Trim(adoRegistro("CodSubDetalleFile")) <> Valor_Caracter Then strModalidadInteres = Trim(adoRegistro("CodSubDetalleFile"))
                        
            curValorAnterior = curSaldoValorizar / ((1 + dblFactorDiario) ^ (intDiasPlazo - intDiasDeRenta + 1))
            
            curValorActual = curSaldoValorizar / ((1 + dblFactorDiario) ^ (intDiasPlazo - intDiasDeRenta))
                        
            curMontoRenta = Round(curValorActual - curValorAnterior, 2)
            
            '*** Ganancia/Pérdida ***
            If curMontoRenta <> 0 Then
                strDescripAsiento = "Valorización" & Space(1) & strNemonico
                strDescripMovimiento = "Pérdida"
                If curMontoRenta > 0 Then strDescripMovimiento = "Ganancia"
                                                
                If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
                    curMontoRentaContable = Round(curMontoRenta * dblTipoCambioCierre, 2)
                Else
                    curMontoRentaContable = curMontoRenta
                End If
                                                
                .CommandType = adCmdStoredProc
                '*** Obtener el número del parámetro **
                .CommandText = "up_ACObtenerUltNumero"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "up_ACObtenerUltNumeroTmp"  '*** Simulación ***
                
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
                
'                .CommandText = "BEGIN TRAN ProcAsiento"
'                adoConn.Execute .CommandText
                
                On Error GoTo Ctrl_Error
                
                '*** Contabilizar ***
                strFechaGrabar = strFechaCierre & Space(1) & Format(Time, "hh:mm")
                
                '*** Cabecera ***
                .CommandText = "{ call up_ACAdicAsientoContable('"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableTmp('"  '*** Simulación ***
                
                .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
                    strFechaGrabar & "','" & _
                    gstrPeriodoActual & "','" & gstrMesActual & "','" & Valor_Caracter & "','" & _
                    strDescripAsiento & "','" & Trim(adoRegistro("CodMoneda")) & "','" & _
                    Codigo_Moneda_Local & "','" & _
                    "','" & _
                    "'," & _
                    CDec(curMontoRenta) & ",'" & Estado_Activo & "'," & _
                    intCantRegistros & ",'" & _
                    strFechaCierre & Space(1) & Format(Time, "hh:ss") & "','" & _
                    strCodModulo & "','" & _
                    "'," & _
                    dblTipoCambioCierre & ",'" & _
                    "','" & _
                    "','" & _
                    strDescripAsiento & "','" & _
                    "','" & _
                    "X','') }"
                adoConn.Execute .CommandText
                
                '*** Detalle ***
                .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
                    "WHERE TipoOperacion='" & Codigo_Dinamica_Provision & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                    IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"

                Set adoConsulta = .Execute
        
                Do While Not adoConsulta.EOF
                
                    Select Case Trim(adoConsulta("TipoCuentaInversion"))
                        Case Codigo_CtaInversion
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvInteres
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteres
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaCosto
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaIngresoOperacional
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresVencido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaVacCorrido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaXPagar
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaXCobrar
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresCorrido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvReajusteK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaReajusteK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvFlucMercado
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaFlucMercado
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvInteresVac
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresVac
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaIntCorridoK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvFlucK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaFlucK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInversionTransito
                            curMontoMovimientoMN = curMontoRenta
                            
                    End Select
                    
                    strIndDebeHaber = Trim(adoConsulta("IndDebeHaber"))
                    If strIndDebeHaber = "H" Then
                        curMontoMovimientoMN = curMontoMovimientoMN * -1
                        If curMontoMovimientoMN > 0 Then strIndDebeHaber = "D"
                    ElseIf strIndDebeHaber = "D" Then
                        If curMontoMovimientoMN < 0 Then strIndDebeHaber = "H"
                    End If
                    
                    If strIndDebeHaber = "T" Then
                        If curMontoMovimientoMN > 0 Then
                            strIndDebeHaber = "D"
                        Else
                            strIndDebeHaber = "H"
                        End If
                    End If
                    strDescripMovimiento = Trim(adoConsulta("DescripDinamica"))
                    curMontoMovimientoME = 0
                    curMontoContable = curMontoMovimientoMN
        
                    dblValorTipoCambio = 1

                    If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
                        curMontoContable = Round(curMontoMovimientoMN * dblTipoCambioCierre, 2)
                        curMontoMovimientoME = curMontoMovimientoMN
                        curMontoMovimientoMN = 0
                        dblValorTipoCambio = dblTipoCambioCierre
                    End If

                    strTipoDocumento = ""
                    strNumDocumento = ""
                    strTipoPersonaContraparte = ""
                    strCodPersonaContraparte = ""
                    strIndContracuenta = ""
                    strCodContracuenta = ""
                    strCodFileContracuenta = ""
                    strCodAnaliticaContracuenta = ""
                    strIndUltimoMovimiento = ""
                                
                    '*** Movimiento ***
                    .CommandText = "{ call up_ACAdicAsientoContableDetalle('"
                    If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableDetalleTmp('"

                    .CommandText = .CommandText & strNumAsiento & "','" & strCodFondo & "','" & _
                        gstrCodAdministradora & "'," & _
                        CInt(adoConsulta("NumSecuencial")) & ",'" & _
                        strFechaGrabar & "','" & _
                        gstrPeriodoActual & "','" & _
                        gstrMesActual & "','" & _
                        strDescripMovimiento & "','" & _
                        strIndDebeHaber & "','" & _
                        Trim(adoConsulta("CodCuenta")) & "','" & _
                        Trim(adoRegistro("CodMoneda")) & "'," & _
                        CDec(curMontoMovimientoMN) & "," & _
                        CDec(curMontoMovimientoME) & "," & _
                        CDec(curMontoContable) & "," & _
                        dblValorTipoCambio & ",'" & _
                        Trim(adoRegistro("CodFile")) & "','" & _
                        Trim(adoRegistro("CodAnalitica")) & "','" & _
                        strTipoDocumento & "','" & _
                        strNumDocumento & "','" & _
                        strTipoPersonaContraparte & "','" & _
                        strCodPersonaContraparte & "','" & _
                        strIndContracuenta & "','" & _
                        strCodContracuenta & "','" & _
                        strCodFileContracuenta & "','" & _
                        strCodAnaliticaContracuenta & "','" & _
                        strIndUltimoMovimiento & "') }"

                    adoConn.Execute .CommandText
                                                    
                    '*** Validar valor de cuenta contable ***
                    If Trim(adoConsulta("CodCuenta")) = Valor_Caracter Then
                        MsgBox "Registro Nro. " & CStr(intContador) & " de Asiento Contable no tiene cuenta asignada", vbCritical, "Valorización Depósitos"
                        gblnRollBack = True
                        Exit Sub
                    End If
                    
                    
                    '*** Insertar en up_GNManInversionValorizacionDiaria **
                    .CommandText = "{ call up_GNManInversionValorizacionDiaria('"
                    If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNManInversionValorizacionDiariaTmp('"
                
                    .CommandText = .CommandText & strCodFondo & "','" & _
                        gstrCodAdministradora & "','" & _
                        strFechaCierre & "','" & _
                        Trim(adoRegistro("CodTitulo")) & "','" & _
                        strCodFile & "','" & _
                        Trim(adoRegistro("CodAnalitica")) & "','" & _
                        Trim(adoRegistro("Nemotecnico")) & "','" & _
                        Trim(adoRegistro("CodMoneda")) & "','" & _
                        Codigo_Moneda_Local & "'," & _
                        dblTasaInteres & "," & _
                        0 & "," & _
                        dblFactorDiario & "," & _
                        adoRegistro("SaldoFinal") & "," & _
                        intDiasDeRenta & "," & _
                        adoRegistro("ValorNominal") & "," & _
                        curValorAnterior & "," & _
                        curValorActual & "," & _
                        curMontoRenta & "," & _
                        0 & "," & _
                        0 & "," & _
                        0 & "," & _
                        curMontoRentaContable & "," & _
                        0 & "," & _
                        0 & "," & _
                        0 & ",'" & gstrCodClaseTipoCambioOperacionFondo & "'," & dblTipoCambioCierre & ",'" & adoRegistro("CodEmisor") & "') }"
                    adoConn.Execute .CommandText
                    
                    
                    adoConsulta.MoveNext
                Loop
                adoConsulta.Close: Set adoConsulta = Nothing
                                
                '*** Actualizar el número del parámetro **
                .CommandText = "{ call up_ACActUltNumero('"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNActUltNumeroTmp('"
                
                .CommandText = .CommandText & strCodFondo & "','" & _
                    gstrCodAdministradora & "','" & _
                    Valor_NumComprobante & "','" & _
                    strNumAsiento & "') }"
                    
                adoConn.Execute .CommandText
                                
'                .CommandText = "COMMIT TRAN ProcAsiento"
'                adoConn.Execute .CommandText
        
            End If
                                    
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
                
    End With
    
    Exit Sub
       
Ctrl_Error:
'    adoComm.CommandText = "ROLLBACK TRAN ProcAsiento"
'    adoConn.Execute adoComm.CommandText
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault


End Sub

Private Sub ValorizacionAcreencias(strTipoCierre As String)

    Dim adoRegistro                 As ADODB.Recordset, adoConsulta         As ADODB.Recordset
    Dim dblPrecioCierre             As Double, dblPrecioPromedio            As Double
    Dim dblTirCierre                As Double, dblFactorDiario              As Double
    Dim dblTasaInteres              As Double, dblFactorDiarioCupon         As Double
    Dim curSaldoGPCapital           As Currency, curSaldoFluctuacionMercado As Currency
    Dim curSaldoInversion           As Currency, curSaldoInteresCorrido     As Currency
    Dim curSaldoFluctuacion         As Currency, curValorAnterior           As Currency
    Dim curValorActual              As Currency, curMontoRenta              As Currency
    Dim curMontoContable            As Currency, curMontoMovimientoMN       As Currency
    Dim curMontoMovimientoME        As Currency, curSaldoValorizar          As Currency
    Dim curMontoProvisionCapital    As Currency, curMontoFluctuacionMercado As Currency
    Dim intCantRegistros            As Integer, intContador                 As Integer
    Dim intRegistro                 As Integer, intBaseCalculo              As Integer
    Dim intDiasPlazo                As Integer, intDiasDeRenta              As Integer
    Dim strNumAsiento               As String, strDescripAsiento            As String
    Dim strDescripMovimiento        As String, strIndDebeHaber              As String
    Dim strCodCuenta                As String, strFiles                     As String
    Dim strCodFile                  As String, strModalidadInteres          As String
    Dim strCodAnalitica             As String, strCodTitulo                 As String
    Dim strNumCuota                 As String, strNumOperacion              As String
    Dim strCodTasa                  As String, strIndCuponCero              As String
    Dim strCodDetalleFile           As String, strNemonico                  As String
    Dim strCodIndiceInicial         As String, strCodIndiceFinal            As String
    Dim strFechaGrabar              As String, strBaseAnual                 As String
    Dim dblTipoCambioCierre         As Double
    
    Dim dblValorTipoCambio          As Double, strTipoDocumento As String
    Dim strNumDocumento             As String
    
    Dim strTipoPersonaContraparte   As String, strCodPersonaContraparte As String
    
    Dim strIndContracuenta          As String, strCodContracuenta As String
    Dim strCodFileContracuenta      As String, strCodAnaliticaContracuenta As String
    Dim strIndUltimoMovimiento      As String
    
    
    '*** Rentabilidad de Valores de Renta Fija Corto Plazo ***
    frmMainMdi.stbMdi.Panels(3).Text = "Calculando rentabilidad de Valores de Renta Fija Corto Plazo..."
       
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT CC.NumOperacionOrig as NumOperacion, CC.CodFile, CC.CodAnalitica, CC.CodTitulo, CC.NumCuota, CC.NumSecuencial, IO.FechaVencimiento  " & _
            "FROM InversionKardex IK " & _
            "JOIN InstrumentoInversion II ON (II.CodFondo = IK.CodFondo and II.CodAdministradora = IK.CodAdministradora " & _
            "AND II.CodFile=IK.CodFile AND II.CodAnalitica=IK.CodAnalitica and IK.CodTitulo = II.CodTitulo) " & _
            "JOIN InversionOperacion IO ON (IO.CodFondo = IK.CodFondo AND IO.CodAdministradora = IK.CodAdministradora AND IO.NumOperacion = IK.NumOperacion) " & _
            "JOIN InversionOperacionCalendarioCuota CC ON (CC.CodFondo = IK.CodFondo AND CC.CodAdministradora = IK.CodAdministradora AND " & _
            "CC.NumOperacionOrig = dbo.uf_IVObtenerNumOperacionOrigen(IK.CodFondo,IK.CodAdministradora,IK.CodFile,IK.CodAnalitica) " & _
            "AND CC.CodFile = IK.CodFile AND CC.CodAnalitica = IK.CodAnalitica) " & _
            "WHERE " & _
            "IK.CodAdministradora = '" & gstrCodAdministradora & "' AND IK.CodFondo = '" & strCodFondo & "' AND " & _
            "IK.CodFile IN ('006','010','012','014','015') AND IK.SaldoFinal > 0 AND " & _
            "IK.NumKardex = dbo.uf_IVObtenerUltimoMovimientoKardexValor(IK.CodFondo,IK.CodAdministradora,IK.CodTitulo,'" & strFechaCierre & "') AND " & _
            "CC.NumSecuencial = dbo.uf_IVObtenerUltimoCalendarioCuotaVigente(CC.CodFondo,CC.CodAdministradora,CC.CodFile, CC.CodAnalitica,CC.NumCuota,'" & strFechaCierre & "') ORDER BY IK.CodFile, IK.CodAnalitica"
        Set adoRegistro = .Execute
    
        Do Until adoRegistro.EOF
            strNumOperacion = Trim(adoRegistro("NumOperacion"))
            strCodFile = Trim(adoRegistro("CodFile"))
            strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
            strCodTitulo = Trim(adoRegistro("CodTitulo"))
            strNumCuota = Trim$(adoRegistro("NumCuota"))
        
            '*** Grabamos Devengado ***
            .CommandText = "{ call up_IVAdicInversionDevengado('" & strCodFondo & "','" & gstrCodAdministradora & _
                            "','" & strCodFile & "','" & strCodAnalitica & "','" & strCodTitulo & _
                            "','" & strNumCuota & "','" & strFechaCierre & "','" & strTipoCierre & "') }"
            .Execute
            
            If Convertddmmyyyy(strFechaCierre) < CDate(adoRegistro("FechaVencimiento")) Then
                '*** Contabilizamos Devengado ***
                .CommandText = "{ call up_ACProcContabilizarOperacion('" & strCodFondo & "','" & gstrCodAdministradora & _
                                "','" & strFechaCierre & "','" & strCodFile & "','" & strNumOperacion & "', '" & Codigo_Caja_Provision & "') }"
                .Execute
            Else
                '*** Contabilizamos Devengado Adicional ***
                .CommandText = "{ call up_ACProcContabilizarOperacion('" & strCodFondo & "','" & gstrCodAdministradora & _
                                "','" & strFechaCierre & "','" & strCodFile & "','" & strNumOperacion & "', '" & Codigo_Caja_Provision_Intereses_Adicionales & "') }"
                .Execute
                
                '*** Contabilizamos Devengado Moratorio ***
                .CommandText = "{ call up_ACProcContabilizarOperacion('" & strCodFondo & "','" & gstrCodAdministradora & _
                                "','" & strFechaCierre & "','" & strCodFile & "','" & strNumOperacion & "', '" & Codigo_Caja_Provision_Intereses_Moratorios & "') }"
                .Execute
            End If
            
            adoRegistro.MoveNext
        Loop
        
    End With
       
    'El método resultó corto dado que toda la lógica de cálculo y escritura en tablas está dentro del
    'stored procedure
       
    Exit Sub
       
Ctrl_Error:
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Resume Next
    Me.MousePointer = vbDefault
    
End Sub

Private Sub ValorizacionPrestamos(strTipoCierre As String)

    Dim adoRegistro                 As ADODB.Recordset
    Dim strCodFile As String
    Dim strCodAnalitica             As String
    Dim strNumCuota                 As String, strNumOperacion              As String
    
    '*** Rentabilidad de Valores de Renta Fija Corto Plazo ***
    frmMainMdi.stbMdi.Panels(3).Text = "Calculando rentabilidad de Valores de Renta Fija Corto Plazo..."
       
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "Select CC.NumOperacionOrig as NumOperacion, CC.CodFile, CC.CodAnalitica, CC.NumCuota, CC.NumSecuencial, CC.FechaVencimientoCuota   " & _
                        "FROM FinanciamientoOperacionCalendarioCuota CC  " & _
                        "JOIN FinanciamientoKardex FK on (FK.CodFondo = CC.CodFondo and FK.CodAdministradora = CC.CodAdministradora  " & _
                        "and FK.CodFile = CC.CodFile and FK.CodAnalitica = CC.CodAnalitica)  " & _
                        "JOIN FinanciamientoOperacion FO on (FO.CodFondo = CC.CodFondo and FO.CodAdministradora = CC.CodAdministradora  " & _
                        "and FO.CodFile = CC.CodFile and FO.CodAnalitica = CC.CodAnalitica)  " & _
                        "WHERE CC.CodAdministradora = '" & gstrCodAdministradora & "' AND CC.CodFondo = '" & gstrCodFondoContable & "' " & _
                        "AND CC.CodFile IN ('" & CodFile_Financiamiento_Prestamos & "') " & _
                        "AND FK.SaldoFinal > 0 " & _
                        "AND FK.NumKardex = dbo.uf_FIObtenerUltimoMovimientoKardexValor(FK.CodFondo,FK.CodAdministradora,FK.CodFile,FK.CodAnalitica,'" & strFechaCierre & "') " & _
                        "AND FO.TipoOperacion = '01' AND   " & _
                        "CC.NumSecuencial = dbo.uf_FIObtenerUltimoCalendarioCuotaVigente(CC.CodFondo,CC.CodAdministradora,CC.NumOperacionOrig,CC.NumCuota,'" & strFechaCierre & "') " & _
                        "ORDER BY CC.CodFile, CC.CodAnalitica, CC.NumCuota"

        Set adoRegistro = .Execute
    
        Do Until adoRegistro.EOF
            strNumOperacion = Trim$(adoRegistro("NumOperacion"))
            strCodFile = Trim$(adoRegistro("CodFile"))
            strCodAnalitica = Trim$(adoRegistro("CodAnalitica"))
            
            strNumCuota = Trim$(adoRegistro("NumCuota"))
        
            '*** Grabamos Devengado ***
            .CommandText = "{ call up_FIAdicFinanciamientoDevengado('" & strCodFondo & "','" & gstrCodAdministradora & _
                            "','" & strCodFile & "','" & strCodAnalitica & _
                            "','" & strNumCuota & "','" & strFechaCierre & "','" & strTipoCierre & "') }"
            .Execute
            
            If Convertddmmyyyy(strFechaCierre) < CDate(adoRegistro("FechaVencimientoCuota")) _
            And CDate(adoRegistro("FechaVencimientoCuota")) < Convertddmmyyyy("29990101") Then
                '*** Contabilizamos Devengado ***
                .CommandText = "{ call up_ACProcContabilizarOperacion('" & strCodFondo & "','" & gstrCodAdministradora & _
                                "','" & strFechaCierre & "','" & strCodFile & "','" & strNumOperacion & "', '" & Codigo_Caja_Provision & "', 2) }"
                .Execute
            Else
                '*** Contabilizamos Devengado Adicional ***
                .CommandText = "{ call up_ACProcContabilizarOperacion('" & strCodFondo & "','" & gstrCodAdministradora & _
                                "','" & strFechaCierre & "','" & strCodFile & "','" & strNumOperacion & "', '" & Codigo_Caja_Provision_Intereses_Adicionales & "', 2) }"
                .Execute
            End If
            
            adoRegistro.MoveNext
        Loop
        
    End With
       
    'El método resultó corto dado que toda la lógica de cálculo y escritura en tablas está dentro del
    'stored procedure
       
    Exit Sub
       
Ctrl_Error:
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Resume Next
    Me.MousePointer = vbDefault
    
End Sub

Private Sub ValorizacionFlujos(strTipoCierre As String)

    Dim adoRegistro                 As ADODB.Recordset
    Dim strCodFile As String
    Dim strCodAnalitica             As String, strCodTitulo                 As String
    Dim strNumCuota                 As String, strNumOperacion              As String
    
    '*** Rentabilidad de Valores de Renta Fija Corto Plazo ***
    frmMainMdi.stbMdi.Panels(3).Text = "Calculando rentabilidad de Valores de Renta Fija Corto Plazo..."
       
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT IO.NumOperacion as NumOperacion, CC.CodFile, CC.CodAnalitica, CC.CodTitulo, CC.NumCuota, CC.NumSecuencial, CC.FechaVencimientoCuota " & _
                        "FROM InversionOperacionCalendarioCuota CC " & _
                        "JOIN InstrumentoInversion II ON (CC.CodFondo = II.CodFondo and CC.CodAdministradora = II.CodAdministradora " & _
                        "and CC.CodFile = II.CodFile and CC.CodAnalitica = II.CodAnalitica and II.CodTitulo=CC.CodTitulo) " & _
                        "JOIN InversionKardex IK on (IK.CodFondo = CC.CodFondo and IK.CodAdministradora = CC.CodAdministradora " & _
                        "and IK.CodFile = CC.CodFile and IK.CodAnalitica = CC.CodAnalitica and IK.CodTitulo = CC.CodTitulo) " & _
                        "JOIN InversionOperacion IO on (IO.CodFondo = CC.CodFondo and IO.CodAdministradora = CC.CodAdministradora " & _
                        "and IO.CodFile = CC.CodFile and IO.CodAnalitica = CC.CodAnalitica and IO.CodTitulo=CC.CodTitulo) " & _
                        "WHERE CC.CodAdministradora = '" & gstrCodAdministradora & "' AND CC.CodFondo = '" & strCodFondo & "' " & _
                        "AND CC.CodFile IN ('016') " & _
                        "AND IK.SaldoFinal > 0 " & _
                        "AND IK.NumKardex = dbo.uf_IVObtenerUltimoMovimientoKardexValor(IK.CodFondo,IK.CodAdministradora,IK.CodTitulo,'" & strFechaCierre & "') " & _
                        "AND CC.NumDesembolso = 0 " & _
                        "AND IO.TipoOperacion = '01' AND  " & _
                        "CC.NumSecuencial = dbo.uf_IVObtenerUltimoCalendarioCuotaVigente(CC.CodFondo,CC.CodAdministradora,CC.CodFile, CC.CodAnalitica,CC.NumCuota,'" & strFechaCierre & "') " & _
                        "ORDER BY CC.CodFile, CC.CodAnalitica, CC.NumCuota"

        Set adoRegistro = .Execute
    
        Do Until adoRegistro.EOF
            strNumOperacion = Trim$(adoRegistro("NumOperacion"))
            strCodFile = Trim$(adoRegistro("CodFile"))
            strCodAnalitica = Trim$(adoRegistro("CodAnalitica"))
            strCodTitulo = Trim$(adoRegistro("CodTitulo"))
            strNumCuota = Trim$(adoRegistro("NumCuota"))
        
            '*** Grabamos Devengado ***
            .CommandText = "{ call up_IVAdicInversionDevengado('" & strCodFondo & "','" & gstrCodAdministradora & _
                            "','" & strCodFile & "','" & strCodAnalitica & "','" & strCodTitulo & _
                            "','" & strNumCuota & "','" & strFechaCierre & "','" & strTipoCierre & "') }"
            .Execute
            
            If Convertddmmyyyy(strFechaCierre) < CDate(adoRegistro("FechaVencimientoCuota")) _
            And CDate(adoRegistro("FechaVencimientoCuota")) < Convertddmmyyyy("29990101") Then
                '*** Contabilizamos Devengado ***
                .CommandText = "{ call up_ACProcContabilizarOperacion('" & strCodFondo & "','" & gstrCodAdministradora & _
                                "','" & strFechaCierre & "','" & strCodFile & "','" & strNumOperacion & "', '" & Codigo_Caja_Provision & "') }"
                .Execute
            Else
                '*** Contabilizamos Devengado Adicional ***
                .CommandText = "{ call up_ACProcContabilizarOperacion('" & strCodFondo & "','" & gstrCodAdministradora & _
                                "','" & strFechaCierre & "','" & strCodFile & "','" & strNumOperacion & "', '" & Codigo_Caja_Provision_Intereses_Adicionales & "') }"
                .Execute
                
                '*** Contabilizamos Devengado Moratorio ***
                .CommandText = "{ call up_ACProcContabilizarOperacion('" & strCodFondo & "','" & gstrCodAdministradora & _
                                "','" & strFechaCierre & "','" & strCodFile & "','" & strNumOperacion & "', '" & Codigo_Caja_Provision_Intereses_Moratorios & "') }"
                .Execute

            End If
            
            adoRegistro.MoveNext
        Loop
        
    End With
       
    'El método resultó corto dado que toda la lógica de cálculo y escritura en tablas está dentro del
    'stored procedure
       
    Exit Sub
       
Ctrl_Error:
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Resume Next
    Me.MousePointer = vbDefault
    
End Sub

Private Sub ValorizacionRentaVariable(strTipoCierre As String)

    Dim adoRegistro                 As ADODB.Recordset, adoConsulta     As ADODB.Recordset
    Dim dblPrecioCierre             As Double, dblPrecioPromedio        As Double
    Dim dblTirCierre                As Double, curSaldoInversion        As Currency
    Dim curSaldoFluctuacion         As Currency, curValorAnterior       As Currency
    Dim curValorActual              As Currency, curMontoRenta          As Currency
    Dim curMontoContable            As Currency, curMontoMovimientoMN   As Currency
    Dim curMontoMovimientoME        As Currency, dblTipoCambioCierre    As Double
    Dim intCantRegistros            As Integer, intContador             As Integer
    Dim intRegistro                 As Integer
    Dim strNumAsiento               As String, strDescripAsiento        As String
    Dim strDescripMovimiento        As String, strIndDebeHaber          As String
    Dim strCodFile                  As String, strCodDetalleFile        As String
    Dim strCodCuenta                As String, strNemonico              As String
    Dim strFechaGrabar              As String
    Dim strTipoAuxiliar             As String
    Dim strCodAuxiliar              As String
    Dim strCtaCostoInversion        As String
    Dim curMontoRentaContable       As Currency
    
    Dim dblValorTipoCambio          As Double, strTipoDocumento As String
    Dim strNumDocumento             As String
    
    Dim strTipoPersonaContraparte   As String, strCodPersonaContraparte As String
    
    Dim strIndContracuenta          As String, strCodContracuenta As String
    Dim strCodFileContracuenta      As String, strCodAnaliticaContracuenta As String
    Dim strIndUltimoMovimiento      As String
    
    
    '*** Rentabilidad de Valores de Renta Variable ***
    frmMainMdi.stbMdi.Panels(3).Text = "Calculando rentabilidad de Valores de Renta Variable..."
       
    Set adoRegistro = New ADODB.Recordset
    
    strTipoAuxiliar = "01" 'Inversiones
    
    With adoComm
 
        .CommandText = "SELECT IK.CodTitulo,II.CodMoneda,IK.CodAnalitica,IK.CodFile,SaldoFinal,MontoSaldo,MontoMovimiento,MontoComision,CodDetalleFile,CodSubDetalleFile," & _
            "ValorNominal,CodTipoTasa,BaseAnual,CodTasa,TasaInteres,DiasPlazo,IndCuponCero,FechaEmision,Nemotecnico, II.CodEmisor " & _
            "FROM InversionKardex IK LEFT JOIN InstrumentoInversion II " & _
            "ON(II.CodTitulo=IK.CodTitulo) " & _
            "WHERE IK.CodFile IN ('004','017') AND FechaOperacion < '" & strFechaSiguiente & "' AND SaldoFinal > 0 AND " & _
            "IK.CodAdministradora='" & gstrCodAdministradora & "' AND IK.CodFondo='" & strCodFondo & "' " & _
            "AND IK.NumKardex = dbo.uf_IVObtenerUltimoMovimientoKardexValor(IK.CodFondo,IK.CodAdministradora,IK.CodTitulo,'" & strFechaCierre & "')"
        
        Set adoRegistro = .Execute
    
        Do Until adoRegistro.EOF
            strCodFile = Trim(adoRegistro("CodFile"))
            strCodDetalleFile = Trim(adoRegistro("CodDetalleFile"))
            strNemonico = Trim(adoRegistro("Nemotecnico"))
            
            curSaldoInversion = 0
            curSaldoFluctuacion = 0
            dblPrecioCierre = 0
            dblTirCierre = 0
            dblPrecioPromedio = 0
                        
            '*** Obtener Tipo de Cambio ***
            'dblTipoCambioCierre = ObtenerTipoCambio(gstrCodClaseTipoCambioFondo, gstrValorTipoCambioCierre, dtpFechaCierre.Value, Trim(adoRegistro("CodMoneda")))
            
            Set adoConsulta = New ADODB.Recordset
            
            '*** Verificar Dinamica Contable ***
            .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
                "WHERE TipoOperacion='" & Codigo_Dinamica_Provision & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                If CInt(adoConsulta("NumRegistros")) > 0 Then
                    intCantRegistros = CInt(adoConsulta("NumRegistros"))
                Else
                    MsgBox "NO EXISTE Dinámica Contable para la valorización", vbCritical
                    adoConsulta.Close: Set adoConsulta = Nothing
                    Exit Sub
                End If
            End If
            adoConsulta.Close
            
            '*** Obtener Ultimo Precio de Cierre registrado ***
            .CommandText = "{ call up_IVSelDatoInstrumentoInversion(2,'" & Trim(adoRegistro("CodTitulo")) & "','" & strFechaCierre & "') }"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                dblPrecioCierre = CDbl(adoConsulta("PrecioCierre"))
                dblTirCierre = CDbl(adoConsulta("TirCierre"))
                dblPrecioPromedio = CDbl(adoConsulta("PrecioPromedio"))
            End If
            adoConsulta.Close
            
            '*** Obtener las cuentas de inversi?n ***
            Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, Trim(adoRegistro("CodMoneda")))
            
            '*** Obtener tipo de cambio ***
            'dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
            dblTipoCambioCierre = ObtenerValorTipoCambio(adoRegistro("CodMoneda"), Codigo_Moneda_Local, strFechaCierre, strFechaCierre, gstrCodClaseTipoCambioOperacionFondo, Codigo_Valor_TipoCambioCompra)
            
            
            '*** Obtener Saldo de Inversi?n ***
            If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                .CommandText = "SELECT ISNULL(SUM(SaldoFinalContable),0) Saldo "
            Else
                .CommandText = "SELECT ISNULL(SUM(SaldoFinalME),0) Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            strCtaCostoInversion = "('" & strCtaInversion & "','" & strCtaInversionCostoSAB & "','" & _
                                   strCtaInversionCostoBVL & "','" & strCtaInversionCostoCavali & "','" & _
                                   strCtaInversionCostoFondoGarantia & "','" & strCtaInversionCostoConasev & "','" & _
                                   strCtaInversionCostoIGV & "','" & strCtaInversionCostoCompromiso & "','" & _
                                   strCtaInversionCostoResponsabilidad & "','" & strCtaInversionCostoFondoLiquidacion & "','" & _
                                   strCtaInversionCostoComisionEspecial & "','" & strCtaInversionCostoGastosBancarios & "')"
                        
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta IN " & strCtaCostoInversion & " AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoInversion = CDbl(adoConsulta("Saldo"))
            Else
                curSaldoInversion = 0
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Provision (GANANCIA Y PERDIDA) ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT ISNULL(SUM(SaldoFinalContable),0) Saldo "
            Else
                .CommandText = "SELECT ISNULL(SUM(SaldoFinalME),0) Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
                        
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvFlucMercado & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoFluctuacion = CDbl(adoConsulta("Saldo"))
            Else
                curSaldoFluctuacion = 0
            End If
            adoConsulta.Close
            
            strCodAuxiliar = strCodFile + adoRegistro("CodAnalitica")
            
'            If adoRegistro("CodAnalitica") = "00000001" Then
'                MsgBox "hola"
'            End If
            
            curValorAnterior = curSaldoInversion + curSaldoFluctuacion
            curValorActual = CCur(adoRegistro("SaldoFinal")) * dblPrecioCierre
            curMontoRenta = curValorActual - curValorAnterior
            
            '*** Ganancia/P?rdida ***
            If curMontoRenta <> 0 Then
                strDescripAsiento = "Valorización" & Space(1) & strNemonico
                strDescripMovimiento = "Pérdida"
                If curMontoRenta > 0 Then strDescripMovimiento = "Ganancia"
                                
                If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
                    curMontoRentaContable = Round(curMontoRenta * dblTipoCambioCierre, 2)
                Else
                    curMontoRentaContable = curMontoRenta
                End If
                                
                .CommandType = adCmdStoredProc
                '*** Obtener el n?mero del par?metro **
                .CommandText = "up_ACObtenerUltNumero"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "up_ACObtenerUltNumeroTmp"  '*** Simulaci?n ***
                
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
                                                
'                .CommandText = "BEGIN TRAN ProcAsiento"
'                adoConn.Execute .CommandText
                
                'On Error GoTo Ctrl_Error
                
                '*** Contabilizar ***
                strFechaGrabar = strFechaCierre & Space(1) & Format(Time, "hh:mm")
                
                '*** Cabecera ***
                .CommandText = "{ call up_ACAdicAsientoContable('"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableTmp('"  '*** Simulaci?n ***
               
                .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
                    strFechaGrabar & "','" & _
                    gstrPeriodoActual & "','" & gstrMesActual & "','" & Valor_Caracter & "','" & _
                    strDescripAsiento & "','" & Trim(adoRegistro("CodMoneda")) & "','" & _
                    Codigo_Moneda_Local & "','',''," & _
                    CDec(curMontoRenta) & ",'" & Estado_Activo & "'," & _
                    intCantRegistros & ",'" & _
                    strFechaCierre & Space(1) & Format(Time, "hh:ss") & "','" & _
                    strCodModulo & "',''," & _
                    dblTipoCambioCierre & ",'','','" & _
                    strDescripAsiento & "','','X','') }"
                adoConn.Execute .CommandText
                
                '*** Detalle ***
                .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
                    "WHERE TipoOperacion='" & Codigo_Dinamica_Provision & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                    IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"
                Set adoConsulta = .Execute
        
                Do While Not adoConsulta.EOF
                
                    curMontoMovimientoMN = 0
                
                    Select Case Trim(adoConsulta("TipoCuentaInversion"))
                        Case Codigo_CtaInversion
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvInteres
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteres
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaCosto
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaIngresoOperacional
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresVencido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaVacCorrido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaXPagar
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaXCobrar
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresCorrido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvReajusteK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaReajusteK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvFlucMercado
                            If curMontoRenta > 0 Then curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaFlucMercado
                            If curMontoRenta > 0 Then curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvFlucMercado_Perdida
                            If curMontoRenta < 0 Then curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaFlucMercado_Perdida
                            If curMontoRenta < 0 Then curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvInteresVac
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresVac
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaIntCorridoK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvFlucK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaFlucK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInversionTransito
                            curMontoMovimientoMN = curMontoRenta
                            
                    End Select
                    
                    strIndDebeHaber = Trim(adoConsulta("IndDebeHaber"))
                    
                    If strIndDebeHaber = "H" Then
                        'curMontoMovimientoMN = curMontoMovimientoMN * -1
                        If curMontoMovimientoMN > 0 Then
                            curMontoMovimientoMN = curMontoMovimientoMN * -1
                        End If
                    ElseIf strIndDebeHaber = "D" Then
                        If curMontoMovimientoMN < 0 Then
                            curMontoMovimientoMN = curMontoMovimientoMN * -1
                        End If
                    End If
                    
                    If strIndDebeHaber = "T" Then
                        If curMontoMovimientoMN > 0 Then
                            strIndDebeHaber = "D"
                        Else
                            strIndDebeHaber = "H"
                        End If
                    End If
                    
                    strDescripMovimiento = Trim(adoConsulta("DescripDinamica"))
                    curMontoMovimientoME = 0
                    curMontoContable = curMontoMovimientoMN
        
                    dblValorTipoCambio = 1
                    
                    If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
                        curMontoContable = Round(curMontoMovimientoMN * dblTipoCambioCierre, 2)
                        curMontoMovimientoME = curMontoMovimientoMN
                        curMontoMovimientoMN = 0
                        dblValorTipoCambio = dblTipoCambioCierre
                    End If
                                
                    strTipoDocumento = ""
                    strNumDocumento = ""
                    strTipoPersonaContraparte = ""
                    strCodPersonaContraparte = ""
                    strIndContracuenta = ""
                    strCodContracuenta = ""
                    strCodFileContracuenta = ""
                    strCodAnaliticaContracuenta = ""
                    strIndUltimoMovimiento = ""
                    
                    If curMontoContable <> 0 Then
                        '*** Movimiento ***
                        
                        .CommandText = "{ call up_ACAdicAsientoContableDetalle('"
                        If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableDetalleTmp('"
    
                        .CommandText = .CommandText & strNumAsiento & "','" & strCodFondo & "','" & _
                            gstrCodAdministradora & "'," & _
                            CInt(adoConsulta("NumSecuencial")) & ",'" & _
                            strFechaGrabar & "','" & _
                            gstrPeriodoActual & "','" & _
                            gstrMesActual & "','" & _
                            strDescripMovimiento & "','" & _
                            strIndDebeHaber & "','" & _
                            Trim(adoConsulta("CodCuenta")) & "','" & _
                            Trim(adoRegistro("CodMoneda")) & "'," & _
                            CDec(curMontoMovimientoMN) & "," & _
                            CDec(curMontoMovimientoME) & "," & _
                            CDec(curMontoContable) & "," & _
                            dblValorTipoCambio & ",'" & _
                            Trim(adoRegistro("CodFile")) & "','" & _
                            Trim(adoRegistro("CodAnalitica")) & "','" & _
                            strTipoDocumento & "','" & _
                            strNumDocumento & "','" & _
                            strTipoPersonaContraparte & "','" & _
                            strCodPersonaContraparte & "','" & _
                            strIndContracuenta & "','" & _
                            strCodContracuenta & "','" & _
                            strCodFileContracuenta & "','" & _
                            strCodAnaliticaContracuenta & "','" & _
                            strIndUltimoMovimiento & "') }"
    
    
                        adoConn.Execute .CommandText
                                                          
                        '*** Validar valor de cuenta contable ***
                        If Trim(adoConsulta("CodCuenta")) = Valor_Caracter Then
                            MsgBox "Registro Nro. " & CStr(intContador) & " de Asiento Contable no tiene cuenta asignada", vbCritical, "Valorización Depósitos"
                            gblnRollBack = True
                            Exit Sub
                        End If
                    
                    End If

                   
                    '*** Insertar en up_GNManInversionValorizacionDiaria **
                    .CommandText = "{ call up_GNManInversionValorizacionDiaria('"
                    If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNManInversionValorizacionDiariaTmp('"
                
                    .CommandText = .CommandText & strCodFondo & "','" & _
                        gstrCodAdministradora & "','" & _
                        strFechaCierre & "','" & _
                        Trim(adoRegistro("CodTitulo")) & "','" & _
                        strCodFile & "','" & _
                        Trim(adoRegistro("CodAnalitica")) & "','" & _
                        Trim(adoRegistro("Nemotecnico")) & "','" & _
                        Trim(adoRegistro("CodMoneda")) & "','" & _
                        Codigo_Moneda_Local & "'," & _
                        0 & "," & _
                        dblPrecioCierre & "," & _
                        0 & "," & _
                        adoRegistro("SaldoFinal") & "," & _
                        0 & "," & _
                        adoRegistro("ValorNominal") & "," & _
                        curValorAnterior & "," & _
                        curValorActual & "," & _
                        0 & "," & _
                        0 & "," & _
                        0 & "," & _
                        curMontoRenta & "," & _
                        0 & "," & _
                        0 & "," & _
                        0 & "," & _
                        curMontoRentaContable & ",'" & gstrCodClaseTipoCambioOperacionFondo & "'," & dblTipoCambioCierre & ",'" & adoRegistro("CodEmisor") & "') }"
                    adoConn.Execute .CommandText
                    
                    adoConsulta.MoveNext
                Loop
                adoConsulta.Close: Set adoConsulta = Nothing
                
                
                '*** Actualizar el n?mero del par?metro **
                .CommandText = "{ call up_ACActUltNumero('"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNActUltNumeroTmp('"
                
                .CommandText = .CommandText & strCodFondo & "','" & _
                    gstrCodAdministradora & "','" & _
                    Valor_NumComprobante & "','" & _
                    strNumAsiento & "') }"
                adoConn.Execute .CommandText
                
'                .CommandText = "COMMIT TRAN ProcAsiento"
'                adoConn.Execute .CommandText
        
            End If
                                    
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
                
    End With
    
    Exit Sub
       
Ctrl_Error:
'    adoComm.CommandText = "ROLLBACK TRAN ProcAsiento"
'    adoConn.Execute adoComm.CommandText
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
      
End Sub



Private Function TraeCta(ByVal Tipo As String, ByVal codi As String, ByVal Mone As String) As String

    Dim sensql As String, adoresultTmp As New Recordset
    
    sensql = "SELECT * FROM FMCTADEF"
    sensql = sensql + " WHERE TIP_DEFI='" + Tipo + "'"
    sensql = sensql + " AND COD_DEFI='" + codi + "'"
    adoComm.CommandText = sensql
    Set adoresultTmp = adoComm.Execute
    If Mone = "S" Then
      TraeCta = adoresultTmp!cod_ctan
    Else
      TraeCta = adoresultTmp!cod_ctax
    End If
    adoresultTmp.Close: Set adoresultTmp = Nothing
  
End Function


Private Sub ValFlucPrecioBonos()
    
End Sub

Private Sub ValFlucPrecioLetrasH()
    
End Sub

Private Sub CorteEventoCorporativo()
       
    Dim adoRegistro         As ADODB.Recordset, adoRegistroMov  As ADODB.Recordset
    Dim curLiberadas        As Currency, curDividendos              As Currency
    Dim strNumOrdenEvento   As String

    '*** Verificar vencimientos de entregas de acciones ***
    frmMainMdi.stbMdi.Panels(3).Text = "Verificando vencimientos y entregas de acciones..."

    Set adoRegistro = New ADODB.Recordset
    Set adoRegistroMov = New ADODB.Recordset
    
    With adoComm
        '*** Verificar tambi?n a la fecha de Corte para ver la cantidad de Acciones ***
        '*** que tienen Derecho ***
        .CommandText = "SELECT * FROM EventoCorporativoAcuerdo " & _
            "WHERE (FechaCorte>='" & strFechaCierre & "' AND FechaCorte<'" & strFechaSiguiente & "') AND " & _
            "CodAdministradora='" & gstrCodAdministradora & "' AND EstadoEvento='" & Estado_Acuerdo_Ingresado & "'"
        Set adoRegistro = .Execute
        
        Do Until adoRegistro.EOF
            'OBTIENE KARDEX DEL TITULO REFERENCIA
            .CommandText = "SELECT * FROM InversionKardex " & _
                "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodTitulo='" & adoRegistro("CodTituloReferencia") & "' AND " & _
                "FechaMovimiento<='" & Convertyyyymmdd(adoRegistro("FechaCorte")) & "' AND IndUltimoMovimiento='X'"
            Set adoRegistroMov = .Execute
            
            If Not adoRegistroMov.EOF Then
                '*** Ver liberadas ***
                curLiberadas = 0
                curDividendos = 0
                
                If adoRegistro("TipoAcuerdo") = Codigo_Evento_Liberacion Or adoRegistro("TipoAcuerdo") = Codigo_Evento_Preferente Then
                    curLiberadas = CLng(CDbl(adoRegistro("PorcenAccionesLiberadas")) * 0.01 * adoRegistroMov("SaldoFinal")) 'Al tanto por ciento
                End If
                
                If adoRegistro("TipoAcuerdo") = Codigo_Evento_Nominal Then
                    curLiberadas = CLng(CDbl(adoRegistro("PorcenAccionesLiberadas")) * adoRegistroMov("SaldoFinal")) 'Al tanto por 1
                    curLiberadas = curLiberadas - adoRegistroMov("SaldoFinal")
                End If
                
                If curLiberadas <> 0 Then
                    '*** Obtener Secuencial ***
                    strNumOrdenEvento = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumEntregaEvento)

                    .CommandText = "{ call up_GNAdicEventoCorporativoOrden('" & _
                        strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        adoRegistro("CodTitulo") & "'," & CInt(adoRegistro("NumAcuerdo")) & "," & _
                        CLng(strNumOrdenEvento) & ",'" & adoRegistro("CodFile") & "','" & _
                        adoRegistro("CodAnalitica") & "','" & adoRegistro("CodTituloReferencia") & "','" & _
                        adoRegistro("CodFileReferencia") & "','" & adoRegistro("CodAnaliticaReferencia") & "','" & strFechaSiguiente & "'," & _
                        CDec(adoRegistroMov("SaldoFinal")) & "," & CDec(curLiberadas) & "," & _
                        "0,0,0,0,0,'" & strFechaSiguiente & "','" & Convertyyyymmdd(adoRegistro("FechaEntrega")) & "','" & _
                        Trim(adoRegistro("DescripAcuerdo")) & "','','" & Estado_Entrega_Generado & "'," & _
                        "0,'" & adoRegistro("TipoAcuerdo") & "','" & _
                        gstrLogin & "','" & strFechaCierre & "','" & _
                        gstrLogin & "','" & strFechaCierre & "') }"
                    adoConn.Execute .CommandText
                    
                    '*** Actualiza Secuencial **
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumEntregaEvento & "','" & strNumOrdenEvento & "') }"
                    adoConn.Execute .CommandText
                End If
                    

                '*** Ver dividendos ***
                If adoRegistro("TipoAcuerdo") = Codigo_Evento_Dividendo Then
                    curDividendos = CCur(adoRegistro("PorcenDividendoEfectivo") * adoRegistroMov("SaldoFinal"))
                End If
                
                
                '*** Obtener Secuencial ***
                If curDividendos > 0 Then
                
                    strNumOrdenEvento = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumEntregaEvento)

                    .CommandText = "{ call up_GNAdicEventoCorporativoOrden('" & _
                        strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        adoRegistro("CodTitulo") & "'," & CInt(adoRegistro("NumAcuerdo")) & "," & _
                        CLng(strNumOrdenEvento) & ",'" & adoRegistro("CodFile") & "','" & _
                        adoRegistro("CodAnalitica") & "','" & adoRegistro("CodTituloReferencia") & "','" & _
                        adoRegistro("CodFileReferencia") & "','" & adoRegistro("CodAnaliticaReferencia") & "','" & strFechaSiguiente & "'," & _
                        CDec(adoRegistroMov("SaldoFinal")) & ",0," & _
                        curDividendos & ",0,0,0,0,'" & strFechaSiguiente & "','" & Convertyyyymmdd(adoRegistro("FechaEntrega")) & "','" & _
                        Trim(adoRegistro("DescripAcuerdo")) & "','','" & Estado_Entrega_Generado & "'," & _
                        "0,'" & adoRegistro("TipoAcuerdo") & "','" & _
                        gstrLogin & "','" & strFechaCierre & "','" & _
                        gstrLogin & "','" & strFechaCierre & "') }"
                    adoConn.Execute .CommandText
                    
                    '*** Actualiza Secuencial **
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumEntregaEvento & "','" & strNumOrdenEvento & "') }"
                    adoConn.Execute .CommandText
                End If
            End If
            adoRegistroMov.Close
            
            adoRegistro.MoveNext
        Loop
        Set adoRegistroMov = Nothing
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
   
End Sub


Private Sub VerPagCtaCte()

    
End Sub

Private Sub VerPagIntAho()


End Sub

Private Sub VencimientoRentaFijaLargoPlazo()

    Dim adoRegistro             As ADODB.Recordset, adoConsulta     As ADODB.Recordset
    Dim dblPrecioCierre         As Double, dblPrecioPromedio        As Double
    Dim dblTirCierre            As Double, dblFactorDiario          As Double
    Dim dblTasaInteres          As Double, dblFactorDiarioCupon     As Double
    Dim dblPrecioUnitario       As Double, dblValorPromedioKardex   As Double
    Dim dblInteresCorridoPromedio As Double, dblTirOperacionKardex  As Double
    Dim dblTirPromedioKardex    As Double, dblTirNetaKardex         As Double
    Dim dblValorCupon           As Double, dblValorInteres          As Double
    Dim dblValorAmortizacion    As Double, dblPorcenAmortizacion    As Double
    Dim curSaldoInversion       As Currency, curSaldoInteresCorrido As Currency
    Dim curSaldoFluctuacion     As Currency, curValorAnterior       As Currency
    Dim curValorActual          As Currency, curMontoRenta          As Currency
    Dim curMontoContable        As Currency, curMontoMovimientoMN   As Currency
    Dim curMontoMovimientoME    As Currency, curSaldoValorizar      As Currency
    Dim curCantMovimiento       As Currency, curKarValProm          As Currency
    Dim curValorMovimiento      As Currency, curSaldoInicialKardex  As Currency
    Dim curSaldoFinalKardex     As Currency, curValorSaldoKardex    As Currency
    Dim curValComi              As Currency, curVacCorrido          As Currency
    Dim curSaldoAmortizacion    As Currency
    Dim intCantRegistros        As Integer, intContador             As Integer
    Dim intRegistro             As Integer, intBaseCalculo          As Integer
    Dim intDiasPlazo            As Integer, intDiasDeRenta          As Integer
    Dim strNumAsiento           As String, strDescripAsiento        As String
    Dim strNumOperacion         As String, strNumKardex             As String
    Dim strNumCaja              As String, strFechaPago             As String
    Dim strCodTitulo            As String, strCodEmisor             As String
    Dim strDescripMovimiento    As String, strIndDebeHaber          As String
    Dim strCodCuenta            As String, strFiles                 As String
    Dim strCodFile              As String, strModalidadInteres      As String
    Dim strCodTasa              As String, strIndCuponCero          As String
    Dim strCodDetalleFile       As String, strCodAnalitica          As String
    Dim strCodSubDetalleFile    As String, strFechaGrabar           As String
    Dim strSQLOperacion         As String, strSQLKardex             As String
    Dim strSQLOrdenCaja         As String, strSQLOrdenCajaDetalle   As String
    Dim strIndUltimoMovimiento  As String, strTipoMovimientoKardex  As String
    Dim blnVenceTitulo          As Boolean, blnVenceCupon           As Boolean
    Dim dblTipoCambioCierre     As Double

    '*** Verificación de Vencimiento de Valores de Depósito ***
    frmMainMdi.stbMdi.Panels(3).Text = "Verificando Vencimiento de Valores de Renta Fija Largo Plazo..."
       
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT IK.CodTitulo,II.CodMoneda,IK.CodAnalitica,IK.CodFile,SaldoFinal,CodDetalleFile,CodSubDetalleFile," & _
            "ValorNominal,CodTipoTasa,BaseAnual,CodTasa,TasaInteres,DiasPlazo,IndCuponCero,FechaEmision,FechaVencimiento," & _
            "II.CodEmisor,IK.PrecioUnitario,IK.MontoMovimiento,IK.SaldoInteresCorrido,IK.MontoSaldo,IK.ValorPromedio " & _
            "FROM InversionKardex IK JOIN InstrumentoInversion II " & _
            "ON(II.CodTitulo=IK.CodTitulo) " & _
            "WHERE IK.CodFile IN ('005','007') AND FechaOperacion <='" & strFechaCierre & "' AND SaldoFinal > 0 AND " & _
            "IK.CodAdministradora='" & gstrCodAdministradora & "' AND IK.CodFondo='" & strCodFondo & "' AND IndUltimoMovimiento='X'"
            
        Set adoRegistro = .Execute
    
        Do Until adoRegistro.EOF
            blnVenceTitulo = False: blnVenceCupon = False
            
            '*** Obtener Secuenciales ***
            strNumAsiento = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumComprobante)
            strNumOperacion = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOperacion)
            strNumKardex = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumKardex)
            strNumCaja = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOrdenCaja)
            
            '*** Fecha de Vencimiento del Título = Fecha de Cierre Más 1 Día ***
            If Convertyyyymmdd(adoRegistro("FechaVencimiento")) = strFechaSiguiente Then blnVenceTitulo = True
            
            '*** Obtener el cupón vigente ***
            .CommandText = "SELECT ValorCupon,ValorInteres,ValorAmortizacion,PorcenAmortizacion,FechaVencimiento,FechaPago " & _
                "FROM InstrumentoInversionCalendario " & _
                "WHERE CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "' AND IndVigente='X'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                dblValorCupon = CDbl(adoConsulta("ValorCupon"))
                dblValorInteres = CDbl(adoConsulta("ValorInteres"))
                dblValorAmortizacion = CDbl(adoConsulta("ValorAmortizacion"))
                dblPorcenAmortizacion = CDbl(adoConsulta("PorcenAmortizacion"))
                strFechaPago = Convertyyyymmdd(adoConsulta("FechaPago"))
                If Convertyyyymmdd(adoConsulta("FechaVencimiento")) = strFechaSiguiente Then blnVenceCupon = True
            End If
            adoConsulta.Close
                        
            '*** Si vence el título o el cupón ***
            If blnVenceTitulo Or blnVenceCupon Then
                strCodTitulo = Trim(adoRegistro("CodTitulo"))
                strCodFile = Trim(adoRegistro("CodFile"))
                strCodDetalleFile = Trim(adoRegistro("CodDetalleFile"))
                strCodSubDetalleFile = Trim(adoRegistro("CodSubDetalleFile"))
                strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
                strCodEmisor = Trim(adoRegistro("CodEmisor"))
                strModalidadInteres = Trim(adoRegistro("CodDetalleFile"))
                strCodTasa = Trim(adoRegistro("CodTasa"))
                dblTasaInteres = CDbl(adoRegistro("TasaInteres"))
                intDiasPlazo = CInt(adoRegistro("DiasPlazo"))
                strIndCuponCero = Trim(adoRegistro("IndCuponCero"))
                curSaldoValorizar = CCur(adoRegistro("SaldoFinal"))
                curKarValProm = CDbl(adoRegistro("ValorPromedio"))
                intDiasDeRenta = DateDiff("d", CVDate(adoRegistro("FechaEmision")), gdatFechaActual) + 1
            
                Set adoConsulta = New ADODB.Recordset
                        
                '*** Verificar Dinamica Contable ***
                .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
                    "WHERE TipoOperacion='" & Codigo_Dinamica_Vencimiento & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                    IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"

                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    If CInt(adoConsulta("NumRegistros")) > 0 Then
                        intCantRegistros = CInt(adoConsulta("NumRegistros"))
                    Else
                        MsgBox "NO EXISTE Dinámica Contable para la valorización", vbCritical
                        adoConsulta.Close: Set adoConsulta = Nothing
                        Exit Sub
                    End If
                End If
                adoConsulta.Close
            
                '*** Obtener las cuentas de inversión ***
                Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, Trim(adoRegistro("CodMoneda")))
                
                '*** Obtener tipo de cambio ***
                'dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
                dblTipoCambioCierre = ObtenerValorTipoCambio(adoRegistro("CodMoneda"), Codigo_Moneda_Local, strFechaCierre, strFechaCierre, gstrCodClaseTipoCambioOperacionFondo, Codigo_Valor_TipoCambioCompra)
                
                
                '*** Obtener Saldo de Inversión ***
                If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaInversion & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                    "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
                
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curSaldoInversion = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                '*** Obtener Saldo de Interés Corrido ***
                If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaInteresCorrido & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                    "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
                
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curSaldoInteresCorrido = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                '*** Obtener Saldo de Provisión ***
                If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaProvInteres & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                    "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
                
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curSaldoFluctuacion = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                If blnVenceTitulo Then
                    '*** Calculos ***
                    curCtaXCobrar = Round(curSaldoInversion + curSaldoInteresCorrido + curSaldoFluctuacion, 2)
                    curCtaInversion = curSaldoInversion
                    curCtaCosto = curSaldoInversion
                    curCtaInteresCorrido = curSaldoInteresCorrido
                    curCtaProvInteres = curSaldoFluctuacion
                    curCtaIngresoOperacional = curCtaXCobrar - curCtaProvInteres
                    
                    curCantMovimiento = CCur(adoRegistro("SaldoFinal"))
                    dblPrecioUnitario = curKarValProm
                    curValorMovimiento = dblPrecioUnitario * curCantMovimiento * -1
                    curSaldoInicialKardex = CCur(adoRegistro("SaldoFinal"))
                    curSaldoFinalKardex = 0
                    curValorSaldoKardex = 0
                    dblValorPromedioKardex = 0
                    dblInteresCorridoPromedio = 0
                    curValComi = 0
                    curVacCorrido = 0
                    dblTirOperacionKardex = 0
                    dblTirPromedioKardex = 0
                    dblTirNetaKardex = 0
                    curSaldoAmortizacion = 0
                End If
                
                '************************
                '*** Armar sentencias ***
                '************************
                strDescripAsiento = "Vencimiento" & Space(1) & "(" & strCodFile & "-" & strCodAnalitica & ")"
                '*** Operación ***
                strSQLOperacion = "{ call up_IVAdicInversionOperacion('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strNumOperacion & "','" & strFechaSiguiente & "','" & strCodTitulo & "','" & Left(strFechaSiguiente, 4) & "','" & _
                    Mid(strFechaSiguiente, 5, 2) & "','','" & Estado_Activo & "','" & strCodAnalitica & "','" & _
                    strCodFile & "','" & strCodAnalitica & "','" & strCodDetalleFile & "','" & strCodSubDetalleFile & "','" & _
                    Codigo_Caja_Vencimiento & "','','','" & strDescripAsiento & "','" & strCodEmisor & "','" & _
                    "','','','" & strFechaSiguiente & "','" & strFechaSiguiente & "','" & _
                    strFechaSiguiente & "','" & adoRegistro("CodMoneda") & "'," & CDec(adoRegistro("SaldoFinal")) & "," & CDec(gdblTipoCambio) & "," & _
                    CDec(adoRegistro("ValorNominal")) & "," & CDec(adoRegistro("PrecioUnitario")) & "," & CDec(adoRegistro("MontoMovimiento")) & "," & CDec(adoRegistro("SaldoInteresCorrido")) & "," & _
                    "0,0,0,0,0,0,0," & CDec(curCtaXCobrar) & ",0,0,0,0,0,0,0,0,0," & _
                    "0,0,0,0,'X','" & strNumAsiento & "','','','" & _
                    "','','','',0,'','','','',''," & CDec(dblTasaInteres) & "," & _
                    "0,0,'','','','" & gstrLogin & "') }"
                                
                strIndUltimoMovimiento = "X"
                strTipoMovimientoKardex = "S"
                '*** Kardex ***
                strSQLKardex = "{ call up_IVAdicInversionKardex('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strCodTitulo & "','" & strNumKardex & "','" & strFechaSiguiente & "','" & Left(strFechaSiguiente, 4) & "','" & _
                    Mid(strFechaSiguiente, 5, 2) & "','','" & strNumOperacion & "','" & strCodEmisor & "','','O','" & _
                    strFechaSiguiente & "','" & strTipoMovimientoKardex & "','O'," & curCantMovimiento & ",'" & adoRegistro("CodMoneda") & "'," & _
                    dblPrecioUnitario & "," & curValorMovimiento & "," & curValComi & "," & curSaldoInicialKardex & "," & _
                    curSaldoFinalKardex & "," & curValorSaldoKardex & ",'" & strDescripAsiento & "'," & dblValorPromedioKardex & ",'" & _
                    strIndUltimoMovimiento & "','" & strCodFile & "','" & strCodAnalitica & "'," & dblInteresCorridoPromedio & "," & _
                    curSaldoInteresCorrido & "," & dblTirOperacionKardex & "," & dblTirPromedioKardex & "," & curVacCorrido & "," & _
                    dblTirNetaKardex & "," & curSaldoAmortizacion & ") }"
    
                '*** Orden de Cobro/Pago ***
                strSQLOrdenCaja = "{ call up_ACAdicMovimientoFondo('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strNumCaja & "','" & strFechaSiguiente & "','" & Trim(frmMainMdi.Tag) & "','" & strNumOperacion & "','" & strFechaPago & "','" & _
                    strNumAsiento & "','','E','" & strCtaXCobrar & "'," & CDec(curCtaXCobrar) & ",'" & _
                    strCodFile & "','" & strCodAnalitica & "','" & adoRegistro("CodMoneda") & "','" & _
                    strDescripAsiento & "','" & Codigo_Caja_Vencimiento & "','','" & Estado_Caja_NoConfirmado & "','','','','','','','','','','" & gstrLogin & "') }"
                
                '*** Orden de Cobro/Pago Detalle ***
                strSQLOrdenCajaDetalle = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strNumCaja & "','" & strFechaSiguiente & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripAsiento & "','" & _
                    "H','" & strCtaXCobrar & "'," & CDec(curCtaXCobrar) * -1 & ",'" & _
                    strCodFile & "','" & strCodAnalitica & "','" & adoRegistro("CodMoneda") & "','') }"
                                
                '*** Monto Orden ***
                If curCtaXCobrar > 0 Then
                                                                                            
                    On Error GoTo Ctrl_Error
                    
'                    .CommandText = "BEGIN TRANSACTION ProcAsiento"
'                    adoConn.Execute .CommandText
                                                            
                    '*** Actualizar indicador de último movimiento en Kardex ***
                    .CommandText = "UPDATE InversionKardex SET IndUltimoMovimiento='' " & _
                        "WHERE CodAnalitica='" & strCodAnalitica & "' AND CodFile='" & strCodFile & "' AND " & _
                        "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                        "IndUltimoMovimiento='X'"
        
                    adoConn.Execute .CommandText
        
                    '*** Inserta movimiento en el kardex ***
                    adoConn.Execute strSQLKardex
                    
                    '*** Operación ***
                    adoConn.Execute strSQLOperacion
                    
                    '*** Contabilizar ***
                    strFechaGrabar = strFechaSiguiente & Space(1) & Format(Time, "hh:mm")
                    
                    '*** Cabecera ***
                    .CommandText = "{ call up_ACAdicAsientoContable('"
                    .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
                        strFechaGrabar & "','" & _
                        Left(strFechaSiguiente, 4) & "','" & Mid(strFechaSiguiente, 5, 2) & "','" & _
                        "','" & _
                        strDescripAsiento & "','" & Trim(adoRegistro("CodMoneda")) & "','" & _
                        Codigo_Moneda_Local & "','" & _
                        "','" & _
                        "'," & _
                        CDec(curCtaXCobrar) & ",'" & Estado_Activo & "'," & _
                        intCantRegistros & ",'" & _
                        strFechaSiguiente & Space(1) & Format(Time, "hh:ss") & "','" & _
                        strCodModulo & "','" & _
                        "'," & _
                        dblTipoCambioCierre & ",'" & _
                        "','" & _
                        "','" & _
                        strDescripAsiento & "','" & _
                        "','" & _
                        "X','') }"
                    adoConn.Execute .CommandText
                    
                    '*** Detalle ***
                    .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
                        "WHERE TipoOperacion='" & Codigo_Dinamica_Vencimiento & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                        strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                        IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"

                    Set adoConsulta = .Execute
            
                    Do While Not adoConsulta.EOF
                    
                        Select Case Trim(adoConsulta("TipoCuentaInversion"))
                            Case Codigo_CtaInversion
                                curMontoMovimientoMN = curCtaInversion
                                
                            Case Codigo_CtaProvInteres
                                curMontoMovimientoMN = curCtaProvInteres
                                
                            Case Codigo_CtaInteres
                                curMontoMovimientoMN = curCtaInteres
                                
                            Case Codigo_CtaCosto
                                curMontoMovimientoMN = curCtaCosto
                                
                            Case Codigo_CtaIngresoOperacional
                                curMontoMovimientoMN = curCtaIngresoOperacional
                                
                            Case Codigo_CtaInteresVencido
                                curMontoMovimientoMN = curCtaInteresVencido
                                
                            Case Codigo_CtaVacCorrido
                                curMontoMovimientoMN = curCtaVacCorrido
                                
                            Case Codigo_CtaXPagar
                                curMontoMovimientoMN = curCtaXPagar
                                
                            Case Codigo_CtaXCobrar
                                curMontoMovimientoMN = curCtaXCobrar
                                
                            Case Codigo_CtaInteresCorrido
                                curMontoMovimientoMN = curCtaInteresCorrido
                                
                            Case Codigo_CtaProvReajusteK
                                curMontoMovimientoMN = curCtaProvReajusteK
                                
                            Case Codigo_CtaReajusteK
                                curMontoMovimientoMN = curCtaReajusteK
                                
                            Case Codigo_CtaProvFlucMercado
                                curMontoMovimientoMN = curCtaProvFlucMercado
                                
                            Case Codigo_CtaFlucMercado
                                curMontoMovimientoMN = curCtaFlucMercado
                                
                            Case Codigo_CtaProvInteresVac
                                curMontoMovimientoMN = curCtaProvInteresVac
                                
                            Case Codigo_CtaInteresVac
                                curMontoMovimientoMN = curCtaInteresVac
                                
                            Case Codigo_CtaIntCorridoK
                                curMontoMovimientoMN = curCtaIntCorridoK
                                
                            Case Codigo_CtaProvFlucK
                                curMontoMovimientoMN = curCtaProvFlucK
                                
                            Case Codigo_CtaFlucK
                                curMontoMovimientoMN = curCtaFlucK
                                
                            Case Codigo_CtaInversionTransito
                                curMontoMovimientoMN = curCtaInversionTransito
                                
                        End Select
                        
                        strIndDebeHaber = Trim(adoConsulta("IndDebeHaber"))
                        If strIndDebeHaber = "H" Then
                            curMontoMovimientoMN = curMontoMovimientoMN * -1
                            If curMontoMovimientoMN > 0 Then strIndDebeHaber = "D"
                        ElseIf strIndDebeHaber = "D" Then
                            If curMontoMovimientoMN < 0 Then strIndDebeHaber = "H"
                        End If
                        
                        If strIndDebeHaber = "T" Then
                            If curMontoMovimientoMN > 0 Then
                                strIndDebeHaber = "D"
                            Else
                                strIndDebeHaber = "H"
                            End If
                        End If
                        strDescripMovimiento = Trim(adoConsulta("DescripDinamica"))
                        curMontoMovimientoME = 0
                        curMontoContable = curMontoMovimientoMN
            
                        If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
                            curMontoContable = Round(curMontoMovimientoMN * dblTipoCambioCierre, 2)
                            curMontoMovimientoME = curMontoMovimientoMN
                            curMontoMovimientoMN = 0
                        End If
                                    
                        '*** Movimiento ***
                        .CommandText = "{ call up_ACAdicAsientoContableDetalle('"
                        .CommandText = .CommandText & strNumAsiento & "','" & strCodFondo & "','" & _
                            gstrCodAdministradora & "'," & _
                            CInt(adoConsulta("NumSecuencial")) & ",'" & _
                            strFechaGrabar & "','" & _
                            Left(strFechaSiguiente, 4) & "','" & _
                            Mid(strFechaSiguiente, 5, 2) & "','" & _
                            strDescripMovimiento & "','" & _
                            strIndDebeHaber & "','" & _
                            Trim(adoConsulta("CodCuenta")) & "','" & _
                            Trim(adoRegistro("CodMoneda")) & "'," & _
                            CDec(curMontoMovimientoMN) & "," & _
                            CDec(curMontoMovimientoME) & "," & _
                            CDec(curMontoContable) & ",'" & _
                            Trim(adoRegistro("CodFile")) & "','" & _
                            Trim(adoRegistro("CodAnalitica")) & "','','') }"
                        adoConn.Execute .CommandText
                    
                        '*** Validar valor de cuenta contable ***
                        If Trim(adoConsulta("CodCuenta")) = Valor_Caracter Then
                            MsgBox "Registro Nro. " & CStr(intContador) & " de Asiento Contable no tiene cuenta asignada", vbCritical, "Valorización Depósitos"
                            gblnRollBack = True
                            Exit Sub
                        End If
                        
                        adoConsulta.MoveNext
                    Loop
                    adoConsulta.Close: Set adoConsulta = Nothing
                                    
                    '*** Orden de Cobro ***
                    adoConn.Execute strSQLOrdenCaja
                    adoConn.Execute strSQLOrdenCajaDetalle
        
                    '*** Actualizar Secuenciales **
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumComprobante & "','" & strNumAsiento & "') }"
                    adoConn.Execute .CommandText
                    
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumOperacion & "','" & strNumOperacion & "') }"
                    adoConn.Execute .CommandText
                    
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumKardex & "','" & strNumKardex & "') }"
                    adoConn.Execute .CommandText
                    
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumOrdenCaja & "','" & strNumCaja & "') }"
                    adoConn.Execute .CommandText
                                                                            
'                    .CommandText = "COMMIT TRANSACTION ProcAsiento"
'                    adoConn.Execute .CommandText
            
                End If
            
            End If

            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
                
    End With
    
    Exit Sub
       
Ctrl_Error:
'    adoComm.CommandText = "ROLLBACK TRANSACTION ProcAsiento"
'    adoConn.Execute adoComm.CommandText
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
    

End Sub


Private Sub VencimientoDepositos()

    Dim adoRegistro             As ADODB.Recordset, adoConsulta     As ADODB.Recordset
    Dim dblPrecioCierre         As Double, dblPrecioPromedio        As Double
    Dim dblTirCierre            As Double, dblFactorDiario          As Double
    Dim dblTasaInteres          As Double, dblFactorDiarioCupon     As Double
    Dim dblPrecioUnitario       As Double, dblValorPromedioKardex   As Double
    Dim dblInteresCorridoPromedio As Double, dblTirOperacionKardex  As Double
    Dim dblTirPromedioKardex    As Double, dblTirNetaKardex         As Double
    Dim curSaldoInversion       As Currency, curSaldoInteresCorrido As Currency
    Dim curSaldoFluctuacion     As Currency, curValorAnterior       As Currency
    Dim curValorActual          As Currency, curMontoRenta          As Currency
    Dim curMontoContable        As Currency, curMontoMovimientoMN   As Currency
    Dim curMontoMovimientoME    As Currency, curSaldoValorizar      As Currency
    Dim curCantMovimiento       As Currency, curKarValProm          As Currency
    Dim curValorMovimiento      As Currency, curSaldoInicialKardex  As Currency
    Dim curSaldoFinalKardex     As Currency, curValorSaldoKardex    As Currency
    Dim curValComi              As Currency, curVacCorrido          As Currency
    Dim curSaldoAmortizacion    As Currency
    Dim intCantRegistros        As Integer, intContador             As Integer
    Dim intRegistro             As Integer, intBaseCalculo          As Integer
    Dim intDiasPlazo            As Integer, intDiasDeRenta          As Integer
    Dim strNumAsiento           As String, strDescripAsiento        As String
    Dim strNumOperacion         As String, strNumKardex             As String
    Dim strNumCaja              As String, strFechaPago             As String
    Dim strCodTitulo            As String, strCodEmisor             As String
    Dim strDescripMovimiento    As String, strIndDebeHaber          As String
    Dim strCodCuenta            As String, strFiles                 As String
    Dim strCodFile              As String, strModalidadInteres      As String
    Dim strCodBaseAnual         As String, strCodTipoTasa           As String
    Dim strCodTasa              As String, strIndCuponCero          As String
    Dim strCodDetalleFile       As String, strCodAnalitica          As String
    Dim strCodSubDetalleFile    As String, strFechaGrabar           As String
    Dim strSQLOperacion         As String, strSQLKardex             As String
    Dim strSQLContabilizar      As String, strSQLOrdenCajaDetalle   As String
    Dim strSQLOrdenCaja         As String, strSQLOrdenCajaDetalle2  As String
    Dim strIndUltimoMovimiento  As String, strTipoMovimientoKardex  As String
    Dim blnVenceTitulo          As Boolean, blnVenceCupon           As Boolean
    Dim dblTipoCambioCierre     As Double, strNumOperacionOrig      As String
    
    Dim dblValorTipoCambio          As Double, strTipoDocumento             As String
    Dim strNumDocumento             As String, strTipoPersonaContraparte    As String
    Dim strCodPersonaContraparte    As String
    Dim strIndContracuenta          As String, strCodContracuenta           As String
    Dim strCodFileContracuenta      As String, strCodAnaliticaContracuenta  As String
    
    '*** Verificación de Vencimiento de Valores de Depósito ***
    frmMainMdi.stbMdi.Panels(3).Text = "Verificando Vencimiento de Valores de Depósito..."
       
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT IK.CodTitulo,IOP.NumOperacion,II.CodMoneda,IK.CodAnalitica,IK.CodFile,SaldoFinal,II.CodDetalleFile,II.CodSubDetalleFile," & _
            "II.ValorNominal,IOP.MontoVencimiento,IOP.TipoTasa,IOP.BaseAnual,CodTasa,II.TasaInteres,DiasPlazo,IndCuponCero,II.FechaEmision,II.FechaVencimiento," & _
            "II.CodEmisor,IK.PrecioUnitario,IK.MontoMovimiento,IK.SaldoInteresCorrido,IK.MontoSaldo,IOP.ModoCobroInteres,IOP.CantDocumAnexo,IOP.MontoInteres " & _
            "FROM InversionKardex IK JOIN InstrumentoInversion II ON (II.CodFondo=IK.CodFondo and II.CodTitulo=IK.CodTitulo) " & _
            "JOIN InversionOperacion IOP ON (IOP.CodFondo=IK.CodFondo and IOP.CodTitulo=IK.CodTitulo and IOP.TipoOperacion = '01') " & _
            "WHERE IK.CodFile IN ('003','011') AND IK.FechaOperacion <='" & strFechaCierre & "' AND SaldoFinal > 0 AND " & _
            "IK.CodAdministradora='" & gstrCodAdministradora & "' AND IK.CodFondo='" & strCodFondo & "' AND " & _
            "IK.NumKardex = dbo.uf_IVObtenerUltimoMovimientoKardexValor(IK.CodFondo,IK.CodAdministradora,IK.CodTitulo,'" & strFechaCierre & "')"
        
        Set adoRegistro = .Execute
    
        'IndUltimoMovimiento='X'"
    
        Do Until adoRegistro.EOF
            blnVenceTitulo = False: blnVenceCupon = False
            
            '*** Obtener Secuenciales ***
            strNumAsiento = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumComprobante)
            strNumOperacion = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOperacion)
            strNumKardex = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumKardex)
            strNumCaja = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOrdenCaja)
            
            '*** Fecha de Vencimiento del Título = Fecha de Cierre Más 1 Día ***
            If Convertyyyymmdd(adoRegistro("FechaVencimiento")) = strFechaSiguiente Then blnVenceTitulo = True
            
            '*** Si vence el título o el cupón ***
            If blnVenceTitulo Or blnVenceCupon Then
                strCodTitulo = Trim(adoRegistro("CodTitulo"))
                strNumOperacionOrig = Trim(adoRegistro("NumOperacion"))
                strCodFile = Trim(adoRegistro("CodFile"))
                strCodDetalleFile = Trim(adoRegistro("CodDetalleFile"))
                strCodSubDetalleFile = Trim(adoRegistro("CodSubDetalleFile"))
                strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
                strCodEmisor = Trim(adoRegistro("CodEmisor"))
                strModalidadInteres = Trim(adoRegistro("CodDetalleFile"))
                strCodTasa = Trim(adoRegistro("CodTasa"))
                strCodTipoTasa = Trim(adoRegistro("TipoTasa"))
                strCodBaseAnual = Trim(adoRegistro("BaseAnual"))
                
                dblTasaInteres = CDbl(adoRegistro("TasaInteres"))
                intDiasPlazo = CInt(adoRegistro("DiasPlazo"))
                strIndCuponCero = Trim(adoRegistro("IndCuponCero"))
                curSaldoValorizar = CCur(adoRegistro("SaldoFinal"))
                'curKarValProm = CDbl(adoRegistro("ValorPromedio"))
                intDiasDeRenta = DateDiff("d", CVDate(adoRegistro("FechaEmision")), gdatFechaActual) + 1
            
                Set adoConsulta = New ADODB.Recordset
                        
'                '*** Verificar Dinamica Contable ***
'                .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
'                    "WHERE TipoOperacion='" & Codigo_Dinamica_Vencimiento & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
'                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
'                    IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"
'
'                Set adoConsulta = .Execute
'
'                If Not adoConsulta.EOF Then
'                    If CInt(adoConsulta("NumRegistros")) > 0 Then
'                        intCantRegistros = CInt(adoConsulta("NumRegistros"))
'                    Else
'                        MsgBox "NO EXISTE Dinámica Contable para la valorización", vbCritical
'                        adoConsulta.Close: Set adoConsulta = Nothing
'                        Exit Sub
'                    End If
'                End If
'                adoConsulta.Close
                
                '*** Obtener la Fecha de Pago ***
                .CommandText = "SELECT FechaPago FROM InstrumentoInversionCalendario " & _
                    "WHERE CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    strFechaPago = Convertyyyymmdd(adoConsulta("FechaPago"))
                End If
                adoConsulta.Close
            
                '*** Obtener las cuentas de inversión ***
'                Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, Trim(adoRegistro("CodMoneda")))
                
                '*** Obtener tipo de cambio ***
                'dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
                dblTipoCambioCierre = ObtenerValorTipoCambio(adoRegistro("CodMoneda"), Codigo_Moneda_Local, strFechaCierre, strFechaCierre, gstrCodClaseTipoCambioOperacionFondo, Codigo_Valor_TipoCambioCompra)

                
                '*** Obtener Saldo de Inversión ***
'                If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
'                    .CommandText = "SELECT SaldoFinalContable Saldo "
'                Else
'                    .CommandText = "SELECT SaldoFinalME Saldo "
'                End If
'                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
'                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
'                    "CodCuenta='" & strCtaInversion & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
'                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                    "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
'                    "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
'
'                Set adoConsulta = .Execute
'
'                If Not adoConsulta.EOF Then
'                    curSaldoInversion = CDbl(adoConsulta("Saldo"))
'                End If
'                adoConsulta.Close
'
'                '*** Obtener Saldo de Interés Corrido ***
'                If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
'                    .CommandText = "SELECT SaldoFinalContable Saldo "
'                Else
'                    .CommandText = "SELECT SaldoFinalME Saldo "
'                End If
'                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
'                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
'                    "CodCuenta='" & strCtaInteresCorrido & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
'                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                    "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
'                    "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
'
'                Set adoConsulta = .Execute
'
'                If Not adoConsulta.EOF Then
'                    curSaldoInteresCorrido = CDbl(adoConsulta("Saldo"))
'                End If
'                adoConsulta.Close
'
'                '*** Obtener Saldo de Provisión ***
'                If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
'                    .CommandText = "SELECT SaldoFinalContable Saldo "
'                Else
'                    .CommandText = "SELECT SaldoFinalME Saldo "
'                End If
'                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
'                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
'                    "CodCuenta='" & strCtaProvInteres & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
'                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                    "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
'                    "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
'
'                Set adoConsulta = .Execute
'
'                If Not adoConsulta.EOF Then
'                    curSaldoFluctuacion = CDbl(adoConsulta("Saldo"))
'                End If
'                adoConsulta.Close
                
                If blnVenceTitulo Then
                    '*** Calculos ***
                    curCtaXCobrar = Round(curSaldoInversion + curSaldoInteresCorrido + curSaldoFluctuacion, 2)
                    curCtaInversion = curSaldoInversion
                    curCtaCosto = curSaldoInversion
                    curCtaInteresCorrido = curSaldoInteresCorrido
                    curCtaProvInteres = curSaldoFluctuacion
                    curCtaIngresoOperacional = curCtaXCobrar - curCtaProvInteres
                    
                    curCantMovimiento = CCur(adoRegistro("SaldoFinal"))
                    dblPrecioUnitario = curKarValProm
                    curValorMovimiento = dblPrecioUnitario * curCantMovimiento * CCur(adoRegistro("ValorNominal")) * -1
                    curSaldoInicialKardex = CCur(adoRegistro("SaldoFinal"))
                    curSaldoFinalKardex = 0
                    curValorSaldoKardex = 0
                    dblValorPromedioKardex = 0
                    dblInteresCorridoPromedio = 0
                    curValComi = 0
                    curVacCorrido = 0
                    dblTirOperacionKardex = 0
                    dblTirPromedioKardex = 0
                    dblTirNetaKardex = 0
                    curSaldoAmortizacion = 0
                End If
                
                '************************
                '*** Armar sentencias ***
                '************************
                strDescripAsiento = "Vencimiento" & Space(1) & "(" & strCodFile & "-" & strCodAnalitica & ")"
                '*** Operación ***
                strSQLOperacion = "{ call up_IVAdicInversionOperacion('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strNumOperacion & "','" & strFechaSiguiente & "','" & strCodTitulo & "','" & Left(strFechaSiguiente, 4) & "','" & _
                    Mid(strFechaSiguiente, 5, 2) & "','','" & Estado_Activo & "','" & strCodAnalitica & "','" & _
                    strCodFile & "','" & strCodAnalitica & "','" & strCodDetalleFile & "','" & strCodSubDetalleFile & "','" & _
                    Codigo_Caja_Vencimiento & "','','','" & strDescripAsiento & "','" & strCodEmisor & "','',0,'" & _
                    "','','" & strFechaSiguiente & "','" & strFechaSiguiente & "','" & _
                    strFechaSiguiente & "','" & CStr(adoRegistro("CodMoneda")) & "','" & CStr(adoRegistro("CodMoneda")) & "'," & CDec(adoRegistro("SaldoFinal")) & ",''," & CDec(gdblTipoCambio) & "," & _
                    CDec(adoRegistro("ValorNominal")) & "," & CDec(adoRegistro("ValorNominal")) & "," & CDec(adoRegistro("PrecioUnitario")) & "," & CDec(adoRegistro("MontoMovimiento")) & "," & CDec(adoRegistro("PrecioUnitario")) & "," & CDec(adoRegistro("SaldoInteresCorrido")) & "," & _
                    "0,0,0,0,0,0,0,0,0," & CDec(curCtaXCobrar) & "," & CDec(curCtaXCobrar) & ",0,0,0,0,0,0,0,0,0," & _
                    "0,0,0,0,0,0," & CDec(adoRegistro("MontoVencimiento")) & "," & CInt(adoRegistro("DiasPlazo")) & ",'X','" & strNumAsiento & "','','','" & _
                    "','','','" & strCodEmisor & "','" & strCodEmisor & "','','',0,'','','','" & strCodTipoTasa & "','" & strCodBaseAnual & "'," & CDec(dblTasaInteres) & "," & CDec(dblTasaInteres) & "," & CDec(dblTasaInteres) & "," & CDec(dblTasaInteres) & "," & _
                    "'','','','01','X','07','X','" & gstrLogin & "','" & gstrFechaActual & "','" & gstrLogin & "','" & gstrFechaActual & "','" & strCodTitulo & "','" & CStr(adoRegistro("ModoCobroInteres")) & "'," & CDec(adoRegistro("MontoInteres")) & ",0,0,0,0,'01',0," & _
                    "0,0,0,0,0,0,0,0,0,0,'','','','','','','','','',''," & CDec(adoRegistro("ValorNominal")) & "," & CInt(adoRegistro("CantDocumAnexo")) & ",0," & CDec(adoRegistro("MontoVencimiento")) & ",0) }"
                                                
                strIndUltimoMovimiento = "X"
                strTipoMovimientoKardex = "S"
                '*** Kardex ***
                strSQLKardex = "{ call up_IVAdicInversionKardex('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strCodTitulo & "','" & strNumKardex & "','" & strFechaSiguiente & "','" & Left(strFechaSiguiente, 4) & "','" & _
                    Mid(strFechaSiguiente, 5, 2) & "','','" & strNumOperacion & "','" & strCodEmisor & "','','O','" & _
                    strFechaSiguiente & "','" & strTipoMovimientoKardex & "','O'," & curCantMovimiento & ",'" & CStr(adoRegistro("CodMoneda")) & "'," & _
                    dblPrecioUnitario & "," & dblPrecioUnitario & ",0,0,0,0,0,0,0," & curValorMovimiento & "," & curValComi & "," & curSaldoInicialKardex & "," & _
                    curSaldoFinalKardex & "," & curValorSaldoKardex & ",0,0,'" & Convertyyyymmdd(Valor_Fecha) & "',0,0,'" & strDescripAsiento & "','" & _
                    strIndUltimoMovimiento & "','" & strCodFile & "','" & strCodAnalitica & "'," & dblInteresCorridoPromedio & "," & _
                    curSaldoInteresCorrido & "," & curVacCorrido & "," & curSaldoAmortizacion & ",'01') }"
                
                '*** Contabilizar ***
                strSQLContabilizar = "{ call up_ACProcContabilizarOperacion('" & strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaSiguiente & "','" & strCodFile & "','" & strNumOperacionOrig & "','" & Codigo_Caja_Vencimiento & "') }"
                
                strDescripAsiento = "Abono en Cuenta por Vencimiento de Depósito a Plazo" & Space(1) & "(" & strCodFile & "-" & strCodAnalitica & ")"
                
                '*** Orden de Cobro/Pago ***
                strSQLOrdenCaja = "{ call up_ACAdicMovimientoFondo('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strNumCaja & "','" & strFechaSiguiente & "','" & Trim(frmMainMdi.Tag) & "','" & Valor_NumOperacion & "','" & strNumOperacion & "','" & strFechaPago & "','" & _
                    strNumAsiento & "','','','','E',''," & CDec(adoRegistro("MontoVencimiento")) & ",'" & _
                    strCodFile & "','" & strCodAnalitica & "','" & CStr(adoRegistro("CodMoneda")) & "','" & _
                    strDescripAsiento & "','" & Codigo_Caja_Vencimiento & "','','" & Estado_Caja_NoConfirmado & "','','','','','','','','','','0','" & adoRegistro("CodEmisor") & "','02','" & gstrLogin & "') }"
                
                '*** Orden de Cobro/Pago Detalle ***
                strSQLOrdenCajaDetalle = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strNumCaja & "','" & strFechaSiguiente & "',1,'" & Trim(frmMainMdi.Tag) & "','CUENTA POR COBRAR - DEPÓSITO A PLAZO','" & _
                    CStr(adoRegistro("CodMoneda")) & "'," & (CDec(adoRegistro("ValorNominal")) * -1) & ",'','',0,'','','" & _
                    "H','" & IIf(CStr(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, "168231", "168232") & "','" & _
                    strCodFile & "','" & strCodAnalitica & "','','','','','X') }"
                    
                strSQLOrdenCajaDetalle2 = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strNumCaja & "','" & strFechaSiguiente & "',2,'" & Trim(frmMainMdi.Tag) & "','INTERESES DIFERIDOS - DEP. A PLAZO','" & _
                    CStr(adoRegistro("CodMoneda")) & "'," & (CDec(adoRegistro("MontoInteres")) * -1) & ",'','',0,'','','" & _
                    "H','" & IIf(CStr(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, "496115", "496215") & "','" & _
                    strCodFile & "','" & strCodAnalitica & "','','','','','X') }"
                                
                '*** Monto Orden ***
'                If curCtaXCobrar > 0 Then
                                                                                            
'                    On Error GoTo Ctrl_Error
                    
'                    .CommandText = "BEGIN TRANSACTION ProcAsiento"
'                    adoConn.Execute .CommandText
                                                            
                    '*** Actualizar indicador de último movimiento en Kardex ***
                    .CommandText = "UPDATE InversionKardex SET IndUltimoMovimiento='' " & _
                        "WHERE CodAnalitica='" & strCodAnalitica & "' AND CodFile='" & strCodFile & "' AND " & _
                        "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                        "IndUltimoMovimiento='X'"
        
                    adoConn.Execute .CommandText
        
                    '*** Inserta movimiento en el kardex ***
                    adoConn.Execute strSQLKardex
                    
                    '*** Operación ***
                    adoConn.Execute strSQLOperacion
                    
                    '*** Contabilizar ***
                    adoConn.Execute strSQLContabilizar
'                    strFechaGrabar = strFechaSiguiente & Space(1) & Format(Time, "hh:mm")
'
'                    '*** Cabecera ***
'                    .CommandText = "{ call up_ACAdicAsientoContable('"
'                    .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
'                        strFechaGrabar & "','" & _
'                        Left(strFechaSiguiente, 4) & "','" & Mid(strFechaSiguiente, 5, 2) & "','" & _
'                        "','" & _
'                        strDescripAsiento & "','" & Trim(adoRegistro("CodMoneda")) & "','" & _
'                        Codigo_Moneda_Local & "','" & _
'                        "','" & _
'                        "'," & _
'                        CDec(curCtaXCobrar) & ",'" & Estado_Activo & "'," & _
'                        intCantRegistros & ",'" & _
'                        strFechaSiguiente & Space(1) & Format(Time, "hh:ss") & "','" & _
'                        strCodModulo & "','" & _
'                        "'," & _
'                        dblTipoCambioCierre & ",'" & _
'                        "','" & _
'                        "','" & _
'                        strDescripAsiento & "','" & _
'                        "','" & _
'                        "X','') }"
'                    adoConn.Execute .CommandText
'
'                    '*** Detalle ***
'                    .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
'                        "WHERE TipoOperacion='" & Codigo_Dinamica_Vencimiento & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
'                        strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
'                        IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"
'                    Set adoConsulta = .Execute
'
'                    Do While Not adoConsulta.EOF
'
'                        Select Case Trim(adoConsulta("TipoCuentaInversion"))
'                            Case Codigo_CtaInversion
'                                curMontoMovimientoMN = curCtaInversion
'
'                            Case Codigo_CtaProvInteres
'                                curMontoMovimientoMN = curCtaProvInteres
'
'                            Case Codigo_CtaInteres
'                                curMontoMovimientoMN = curCtaInteres
'
'                            Case Codigo_CtaCosto
'                                curMontoMovimientoMN = curCtaCosto
'
'                            Case Codigo_CtaIngresoOperacional
'                                curMontoMovimientoMN = curCtaIngresoOperacional
'
'                            Case Codigo_CtaInteresVencido
'                                curMontoMovimientoMN = curCtaInteresVencido
'
'                            Case Codigo_CtaVacCorrido
'                                curMontoMovimientoMN = curCtaVacCorrido
'
'                            Case Codigo_CtaXPagar
'                                curMontoMovimientoMN = curCtaXPagar
'
'                            Case Codigo_CtaXCobrar
'                                curMontoMovimientoMN = curCtaXCobrar
'
'                            Case Codigo_CtaInteresCorrido
'                                curMontoMovimientoMN = curCtaInteresCorrido
'
'                            Case Codigo_CtaProvReajusteK
'                                curMontoMovimientoMN = curCtaProvReajusteK
'
'                            Case Codigo_CtaReajusteK
'                                curMontoMovimientoMN = curCtaReajusteK
'
'                            Case Codigo_CtaProvFlucMercado
'                                curMontoMovimientoMN = curCtaProvFlucMercado
'
'                            Case Codigo_CtaFlucMercado
'                                curMontoMovimientoMN = curCtaFlucMercado
'
'                            Case Codigo_CtaProvInteresVac
'                                curMontoMovimientoMN = curCtaProvInteresVac
'
'                            Case Codigo_CtaInteresVac
'                                curMontoMovimientoMN = curCtaInteresVac
'
'                            Case Codigo_CtaIntCorridoK
'                                curMontoMovimientoMN = curCtaIntCorridoK
'
'                            Case Codigo_CtaProvFlucK
'                                curMontoMovimientoMN = curCtaProvFlucK
'
'                            Case Codigo_CtaFlucK
'                                curMontoMovimientoMN = curCtaFlucK
'
'                            Case Codigo_CtaInversionTransito
'                                curMontoMovimientoMN = curCtaInversionTransito
'
'                        End Select
'
'                        strIndDebeHaber = Trim(adoConsulta("IndDebeHaber"))
'                        If strIndDebeHaber = "H" Then
'                            curMontoMovimientoMN = curMontoMovimientoMN * -1
'                            If curMontoMovimientoMN > 0 Then strIndDebeHaber = "D"
'                        ElseIf strIndDebeHaber = "D" Then
'                            If curMontoMovimientoMN < 0 Then strIndDebeHaber = "H"
'                        End If
'
'                        If strIndDebeHaber = "T" Then
'                            If curMontoMovimientoMN > 0 Then
'                                strIndDebeHaber = "D"
'                            Else
'                                strIndDebeHaber = "H"
'                            End If
'                        End If
'
'                        strDescripMovimiento = Trim(adoConsulta("DescripDinamica"))
'                        curMontoMovimientoME = 0
'                        curMontoContable = curMontoMovimientoMN
'
'                        dblValorTipoCambio = 1
'
'                        If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
'                            curMontoContable = Round(curMontoMovimientoMN * dblTipoCambioCierre, 2)
'                            curMontoMovimientoME = curMontoMovimientoMN
'                            curMontoMovimientoMN = 0
'                            dblValorTipoCambio = dblTipoCambioCierre
'                        End If
'
'                        strTipoDocumento = ""
'                        strNumDocumento = ""
'                        strTipoPersonaContraparte = ""
'                        strCodPersonaContraparte = ""
'                        strIndContracuenta = ""
'                        strCodContracuenta = ""
'                        strCodFileContracuenta = ""
'                        strCodAnaliticaContracuenta = ""
'                        strIndUltimoMovimiento = ""
'
'                        '*** Movimiento ***
'                        .CommandText = "{ call up_ACAdicAsientoContableDetalle('"
'                        .CommandText = .CommandText & strNumAsiento & "','" & strCodFondo & "','" & _
'                            gstrCodAdministradora & "'," & _
'                            CInt(adoConsulta("NumSecuencial")) & ",'" & _
'                            strFechaGrabar & "','" & _
'                            Left(strFechaSiguiente, 4) & "','" & _
'                            Mid(strFechaSiguiente, 5, 2) & "','" & _
'                            strDescripMovimiento & "','" & _
'                            strIndDebeHaber & "','" & _
'                            Trim(adoConsulta("CodCuenta")) & "','" & _
'                            Trim(adoRegistro("CodMoneda")) & "'," & _
'                            CDec(curMontoMovimientoMN) & "," & _
'                            CDec(curMontoMovimientoME) & "," & _
'                            CDec(curMontoContable) & "," & _
'                            dblValorTipoCambio & ",'" & _
'                            Trim(adoRegistro("CodFile")) & "','" & _
'                            Trim(adoRegistro("CodAnalitica")) & "','" & _
'                            strTipoDocumento & "','" & _
'                            strNumDocumento & "','" & _
'                            strTipoPersonaContraparte & "','" & _
'                            strCodPersonaContraparte & "','" & _
'                            strIndContracuenta & "','" & _
'                            strCodContracuenta & "','" & _
'                            strCodFileContracuenta & "','" & _
'                            strCodAnaliticaContracuenta & "','" & _
'                            strIndUltimoMovimiento & "') }"
'                        adoConn.Execute .CommandText
'
'                        '*** Saldos ***
''                        .CommandText = "{ call up_ACGenPartidaContableSaldos('"
''                        .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & _
''                            Left(strFechaSiguiente, 4) & "','" & Mid(strFechaSiguiente, 5, 2) & "','" & _
''                            Trim(adoConsulta("CodCuenta")) & "','" & _
''                            Trim(adoRegistro("CodFile")) & "','" & _
''                            Trim(adoRegistro("CodAnalitica")) & "','" & _
''                            strFechaSiguiente & "','" & _
''                            strFechaSubSiguiente & "'," & _
''                            CDec(curMontoMovimientoMN) & "," & _
''                            CDec(curMontoMovimientoME) & "," & _
''                            CDec(curMontoContable) & ",'" & _
''                            strIndDebeHaber & "','" & _
''                            Trim(adoRegistro("CodMoneda")) & "') }"
''                        adoConn.Execute .CommandText
'
'                        '*** Validar valor de cuenta contable ***
'                        If Trim(adoConsulta("CodCuenta")) = Valor_Caracter Then
'                            MsgBox "Registro Nro. " & CStr(intContador) & " de Asiento Contable no tiene cuenta asignada", vbCritical, "Valorización Depósitos"
'                            gblnRollBack = True
'                            Exit Sub
'                        End If
'
'                        adoConsulta.MoveNext
'                    Loop
'                    adoConsulta.Close: Set adoConsulta = Nothing
                                    
                    '*** Orden de Cobro ***
                    adoConn.Execute strSQLOrdenCaja
                    adoConn.Execute strSQLOrdenCajaDetalle
                    adoConn.Execute strSQLOrdenCajaDetalle2
        
                    '*** Actualizar Secuenciales **
'                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
'                        Valor_NumComprobante & "','" & strNumAsiento & "') }"
'                    adoConn.Execute .CommandText
                    
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumOperacion & "','" & strNumOperacion & "') }"
                    adoConn.Execute .CommandText
                    
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumKardex & "','" & strNumKardex & "') }"
                    adoConn.Execute .CommandText
                    
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumOrdenCaja & "','" & strNumCaja & "') }"
                    adoConn.Execute .CommandText
                                                                            
'                    .CommandText = "COMMIT TRANSACTION ProcAsiento"
'                    adoConn.Execute .CommandText
            
'                End If
            
            End If

            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
                
    End With
    
    Exit Sub
       
Ctrl_Error:
'    adoComm.CommandText = "ROLLBACK TRANSACTION ProcAsiento"
'    adoConn.Execute adoComm.CommandText
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
    
End Sub

Private Sub VencimientoRentaFijaCortoPlazo()

    Dim adoRegistro             As ADODB.Recordset, adoConsulta     As ADODB.Recordset
    Dim dblPrecioCierre         As Double, dblPrecioPromedio        As Double
    Dim dblTirCierre            As Double, dblFactorDiario          As Double
    Dim dblTasaInteres          As Double, dblFactorDiarioCupon     As Double
    Dim dblPrecioUnitario       As Double, dblValorPromedioKardex   As Double
    Dim dblInteresCorridoPromedio As Double, dblTirOperacionKardex  As Double
    Dim dblTirPromedioKardex    As Double, dblTirNetaKardex         As Double
    Dim curSaldoInversion       As Currency, curSaldoInteresCorrido As Currency
    Dim curSaldoFluctuacion     As Currency, curValorAnterior       As Currency
    Dim curValorActual          As Currency, curMontoRenta          As Currency
    Dim curMontoContable        As Currency, curMontoMovimientoMN   As Currency
    Dim curMontoMovimientoME    As Currency, curSaldoValorizar      As Currency
    Dim curCantMovimiento       As Currency, curKarValProm          As Currency
    Dim curValorMovimiento      As Currency, curSaldoInicialKardex  As Currency
    Dim curSaldoFinalKardex     As Currency, curValorSaldoKardex    As Currency
    Dim curValComi              As Currency, curVacCorrido          As Currency
    Dim curSaldoAmortizacion    As Currency, curSaldoGPCapital      As Currency
    Dim curSaldoFluctuacionMercado  As Currency
    Dim intCantRegistros        As Integer, intContador             As Integer
    Dim intRegistro             As Integer, intBaseCalculo          As Integer
    Dim intDiasPlazo            As Integer, intDiasDeRenta          As Integer
    Dim strNumAsiento           As String, strDescripAsiento        As String
    Dim strNumOperacion         As String, strNumKardex             As String
    Dim strNumCaja              As String, strFechaPago             As String
    Dim strCodTitulo            As String, strCodEmisor             As String
    Dim strDescripMovimiento    As String, strIndDebeHaber          As String
    Dim strCodCuenta            As String, strFiles                 As String
    Dim strCodFile              As String, strModalidadInteres      As String
    Dim strCodTasa              As String, strIndCuponCero          As String
    Dim strCodDetalleFile       As String, strCodAnalitica          As String
    Dim strCodSubDetalleFile    As String, strFechaGrabar           As String
    Dim strSQLOperacion         As String, strSQLKardex             As String
    Dim strSQLOrdenCaja         As String, strSQLOrdenCajaDetalle   As String
    Dim strIndUltimoMovimiento  As String, strTipoMovimientoKardex  As String
    Dim blnVenceTitulo          As Boolean, blnVenceCupon           As Boolean
    Dim dblTipoCambioCierre     As Double

    Dim dblValorTipoCambio          As Double, strTipoDocumento             As String
    Dim strNumDocumento             As String, strTipoPersonaContraparte    As String
    Dim strCodPersonaContraparte    As String
    Dim strIndContracuenta          As String, strCodContracuenta           As String
    Dim strCodFileContracuenta      As String, strCodAnaliticaContracuenta  As String

    '*** Verificación de Vencimiento de Valores de Depósito ***
    frmMainMdi.stbMdi.Panels(3).Text = "Verificando Vencimiento de Valores de Renta Fija Corto Plazo..."
       
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT IK.CodTitulo,II.CodMoneda,IK.CodAnalitica,IK.CodFile,SaldoFinal,CodDetalleFile,CodSubDetalleFile," & _
            "ValorNominal,CodTipoTasa,BaseAnual,CodTasa,TasaInteres,DiasPlazo,IndCuponCero,FechaEmision,FechaVencimiento," & _
            "II.CodEmisor,IK.PrecioUnitario,IK.MontoMovimiento,IK.SaldoInteresCorrido,IK.MontoSaldo,IK.ValorPromedio " & _
            "FROM InversionKardex IK JOIN InstrumentoInversion II " & _
            "ON(II.CodTitulo=IK.CodTitulo) " & _
            "WHERE IK.CodFile IN ('006','010','012') AND FechaOperacion <='" & strFechaCierre & "' AND SaldoFinal > 0 AND " & _
            "IK.CodAdministradora='" & gstrCodAdministradora & "' AND IK.CodFondo='" & strCodFondo & "' AND " & _
            "IK.NumKardex = dbo.uf_IVObtenerUltimoMovimientoKardexValor(IK.CodFondo,IK.CodAdministradora,IK.CodTitulo,'" & strFechaCierre & "')"
        Set adoRegistro = .Execute
    
        Do Until adoRegistro.EOF
            blnVenceTitulo = False: blnVenceCupon = False
            
            '*** Obtener Secuenciales ***
            strNumAsiento = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumComprobante)
            strNumOperacion = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOperacion)
            strNumKardex = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumKardex)
            strNumCaja = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOrdenCaja)
            
            '*** Fecha de Vencimiento del Título = Fecha de Cierre Más 1 Día ***
            If Convertyyyymmdd(adoRegistro("FechaVencimiento")) = strFechaSiguiente Then blnVenceTitulo = True
            
            '*** Si vence el título o el cupón ***
            If blnVenceTitulo Or blnVenceCupon Then
                strCodTitulo = Trim(adoRegistro("CodTitulo"))
                strCodFile = Trim(adoRegistro("CodFile"))
                strCodDetalleFile = Trim(adoRegistro("CodDetalleFile"))
                strCodSubDetalleFile = Trim(adoRegistro("CodSubDetalleFile"))
                strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
                strCodEmisor = Trim(adoRegistro("CodEmisor"))
                strModalidadInteres = Trim(adoRegistro("CodDetalleFile"))
                strCodTasa = Trim(adoRegistro("CodTasa"))
                dblTasaInteres = CDbl(adoRegistro("TasaInteres"))
                intDiasPlazo = CInt(adoRegistro("DiasPlazo"))
                strIndCuponCero = Trim(adoRegistro("IndCuponCero"))
                curSaldoValorizar = CCur(adoRegistro("SaldoFinal"))
                curKarValProm = CDbl(adoRegistro("ValorPromedio"))
                intDiasDeRenta = DateDiff("d", CVDate(adoRegistro("FechaEmision")), gdatFechaActual) + 1
            
                Set adoConsulta = New ADODB.Recordset
                        
                '*** Verificar Dinamica Contable ***
                .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
                    "WHERE TipoOperacion='" & Codigo_Dinamica_Vencimiento & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                    IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"

                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    If CInt(adoConsulta("NumRegistros")) > 0 Then
                        intCantRegistros = CInt(adoConsulta("NumRegistros"))
                    Else
                        MsgBox "NO EXISTE Dinámica Contable para la valorización", vbCritical
                        adoConsulta.Close: Set adoConsulta = Nothing
                        Exit Sub
                    End If
                End If
                adoConsulta.Close
                
                '*** Obtener la Fecha de Pago ***
                .CommandText = "SELECT FechaPago FROM InstrumentoInversionCalendario " & _
                    "WHERE CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    strFechaPago = Convertyyyymmdd(adoConsulta("FechaPago"))
                End If
                adoConsulta.Close
            
                '*** Obtener las cuentas de inversión ***
                Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, Trim(adoRegistro("CodMoneda")))
                
                '*** Obtener tipo de cambio ***
                'dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
                dblTipoCambioCierre = ObtenerValorTipoCambio(adoRegistro("CodMoneda"), Codigo_Moneda_Local, strFechaCierre, strFechaCierre, gstrCodClaseTipoCambioOperacionFondo, Codigo_Valor_TipoCambioCompra)
                
                
                '*** Obtener Saldo de Inversión ***
                If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaInversion & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                    "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curSaldoInversion = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                '*** Obtener Saldo de Interés Corrido ***
                If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaInteresCorrido & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                    "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curSaldoInteresCorrido = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                '*** Obtener Saldo de Provisión ***
                If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaProvInteres & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                    "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curSaldoFluctuacion = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                '*** Obtener Saldo de Provisión G/P Capital ***
                If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                                
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaProvFlucK & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                    "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curSaldoGPCapital = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                '*** Obtener Saldo de Fluctuación Mercado ***
                If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                                
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaProvFlucMercado & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                    "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curSaldoFluctuacionMercado = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                If blnVenceTitulo Then
                    '*** Calculos ***
                    curCtaXCobrar = Round(curSaldoInversion + curSaldoInteresCorrido + curSaldoFluctuacion, 2)
                    curCtaInversion = curSaldoInversion
                    curCtaCosto = curSaldoInversion
                    curCtaInteresCorrido = curSaldoInteresCorrido
                    curCtaProvInteres = curSaldoFluctuacion
                    curCtaIngresoOperacional = curCtaXCobrar - curCtaProvInteres
                    
                    curCtaProvFlucK = curSaldoGPCapital
                    curCtaFlucK = curSaldoGPCapital * -1
                    curCtaProvFlucMercado = curSaldoFluctuacionMercado
                    curCtaFlucMercado = curSaldoFluctuacionMercado * -1
                    
                    curCantMovimiento = CCur(adoRegistro("SaldoFinal"))
                    dblPrecioUnitario = curKarValProm
                    curValorMovimiento = dblPrecioUnitario * curCantMovimiento * -1
                    curSaldoInicialKardex = CCur(adoRegistro("SaldoFinal"))
                    curSaldoFinalKardex = 0
                    curValorSaldoKardex = 0
                    dblValorPromedioKardex = 0
                    dblInteresCorridoPromedio = 0
                    curValComi = 0
                    curVacCorrido = 0
                    dblTirOperacionKardex = 0
                    dblTirPromedioKardex = 0
                    dblTirNetaKardex = 0
                    curSaldoAmortizacion = 0
                End If
                
                '************************
                '*** Armar sentencias ***
                '************************
                strDescripAsiento = "Vencimiento" & Space(1) & "(" & strCodFile & "-" & strCodAnalitica & ")"
                '*** Operación ***
                strSQLOperacion = "{ call up_IVAdicInversionOperacion('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strNumOperacion & "','" & strFechaSiguiente & "','" & strCodTitulo & "','" & Left(strFechaSiguiente, 4) & "','" & _
                    Mid(strFechaSiguiente, 5, 2) & "','','" & Estado_Activo & "','" & strCodAnalitica & "','" & _
                    strCodFile & "','" & strCodAnalitica & "','" & strCodDetalleFile & "','" & strCodSubDetalleFile & "','" & _
                    Codigo_Caja_Vencimiento & "','','','" & strDescripAsiento & "','" & strCodEmisor & "','" & _
                    "','','','" & strFechaSiguiente & "','" & strFechaSiguiente & "','" & _
                    strFechaSiguiente & "','" & adoRegistro("CodMoneda") & "','" & adoRegistro("CodMoneda") & "','" & adoRegistro("CodMoneda") & "'," & CDec(adoRegistro("SaldoFinal")) & "," & CDec(gdblTipoCambio) & "," & CDec(gdblTipoCambio) & "," & _
                    CDec(adoRegistro("ValorNominal")) & "," & CDec(adoRegistro("PrecioUnitario")) & "," & CDec(adoRegistro("MontoMovimiento")) & "," & CDec(adoRegistro("MontoMovimiento")) & "," & CDec(adoRegistro("SaldoInteresCorrido")) & "," & _
                    "0,0,0,0,0,0,0,0,0," & CDec(curCtaXCobrar) & "," & CDec(curCtaXCobrar) & ",0,0,0,0,0,0,0,0,0," & _
                    "0,0,0,0,0,0,'X','" & strNumAsiento & "','','','" & _
                    "','','','',0,'','','','',''," & CDec(dblTasaInteres) & "," & _
                    "0,0,'','','','" & gstrLogin & "') }"
                
                strIndUltimoMovimiento = "X"
                strTipoMovimientoKardex = "S"
                '*** Kardex ***
                strSQLKardex = "{ call up_IVAdicInversionKardex('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strCodTitulo & "','" & strNumKardex & "','" & strFechaSiguiente & "','" & Left(strFechaSiguiente, 4) & "','" & _
                    Mid(strFechaSiguiente, 5, 2) & "','','" & strNumOperacion & "','" & strCodEmisor & "','','O','" & _
                    strFechaSiguiente & "','" & strTipoMovimientoKardex & "','O'," & curCantMovimiento & ",'" & adoRegistro("CodMoneda") & "'," & _
                    dblPrecioUnitario & "," & curValorMovimiento & "," & curValComi & "," & curSaldoInicialKardex & "," & _
                    curSaldoFinalKardex & "," & curValorSaldoKardex & ",'" & strDescripAsiento & "'," & dblValorPromedioKardex & ",'" & _
                    strIndUltimoMovimiento & "','" & strCodFile & "','" & strCodAnalitica & "'," & dblInteresCorridoPromedio & "," & _
                    curSaldoInteresCorrido & "," & dblTirOperacionKardex & "," & dblTirPromedioKardex & "," & curVacCorrido & "," & _
                    dblTirNetaKardex & "," & curSaldoAmortizacion & ") }"
    
                '*** Orden de Cobro/Pago ***
                strSQLOrdenCaja = "{ call up_ACAdicMovimientoFondo('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strNumCaja & "','" & strFechaSiguiente & "','" & Trim(frmMainMdi.Tag) & "','" & strNumOperacion & "','" & strFechaPago & "','" & _
                    strNumAsiento & "','','E','" & strCtaXCobrar & "'," & CDec(curCtaXCobrar) & ",'" & _
                    strCodFile & "','" & strCodAnalitica & "','" & adoRegistro("CodMoneda") & "','" & _
                    strDescripAsiento & "','" & Codigo_Caja_Vencimiento & "','','" & Estado_Caja_NoConfirmado & "','','','','','','','','','','" & gstrLogin & "') }"
                
                '*** Orden de Cobro/Pago Detalle ***
                strSQLOrdenCajaDetalle = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strNumCaja & "','" & strFechaSiguiente & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripAsiento & "','" & _
                    "H','" & strCtaXCobrar & "'," & CDec(curCtaXCobrar) * -1 & ",'" & _
                    strCodFile & "','" & strCodAnalitica & "','" & adoRegistro("CodMoneda") & "','') }"
                                
                '*** Monto Orden ***
                If curCtaXCobrar > 0 Then
                                                                                            
                    On Error GoTo Ctrl_Error
                    
'                    .CommandText = "BEGIN TRANSACTION ProcAsiento"
'                    adoConn.Execute .CommandText
                                                            
                    '*** Actualizar indicador de último movimiento en Kardex ***
                    .CommandText = "UPDATE InversionKardex SET IndUltimoMovimiento='' " & _
                        "WHERE CodAnalitica='" & strCodAnalitica & "' AND CodFile='" & strCodFile & "' AND " & _
                        "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                        "IndUltimoMovimiento='X'"
        
                    adoConn.Execute .CommandText
        
                    '*** Inserta movimiento en el kardex ***
                    adoConn.Execute strSQLKardex
                    
                    '*** Operación ***
                    adoConn.Execute strSQLOperacion
                    
                    '*** Contabilizar ***
                    strFechaGrabar = strFechaSiguiente & Space(1) & Format(Time, "hh:mm")
                    
                    '*** Cabecera ***
                    .CommandText = "{ call up_ACAdicAsientoContable('"
                    .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
                        strFechaGrabar & "','" & _
                        Left(strFechaSiguiente, 4) & "','" & Mid(strFechaSiguiente, 5, 2) & "','" & _
                        "','" & _
                        strDescripAsiento & "','" & Trim(adoRegistro("CodMoneda")) & "','" & _
                        Codigo_Moneda_Local & "','',''," & _
                        CDec(curCtaXCobrar) & ",'" & Estado_Activo & "'," & _
                        intCantRegistros & ",'" & _
                        strFechaSiguiente & Space(1) & Format(Time, "hh:ss") & "','" & _
                        strCodModulo & "',''," & _
                        dblTipoCambioCierre & ",'','','" & _
                        strDescripAsiento & "','','X','') }"
                    adoConn.Execute .CommandText
                    
                    '*** Detalle ***
                    .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
                        "WHERE TipoOperacion='" & Codigo_Dinamica_Vencimiento & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                        strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                        IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"

                    Set adoConsulta = .Execute
            
                    Do While Not adoConsulta.EOF
                    
                        Select Case Trim(adoConsulta("TipoCuentaInversion"))
                            Case Codigo_CtaInversion
                                curMontoMovimientoMN = curCtaInversion
                                
                            Case Codigo_CtaProvInteres
                                curMontoMovimientoMN = curCtaProvInteres
                                
                            Case Codigo_CtaInteres
                                curMontoMovimientoMN = curCtaInteres
                                
                            Case Codigo_CtaCosto
                                curMontoMovimientoMN = curCtaCosto
                                
                            Case Codigo_CtaIngresoOperacional
                                curMontoMovimientoMN = curCtaIngresoOperacional
                                
                            Case Codigo_CtaInteresVencido
                                curMontoMovimientoMN = curCtaInteresVencido
                                
                            Case Codigo_CtaVacCorrido
                                curMontoMovimientoMN = curCtaVacCorrido
                                
                            Case Codigo_CtaXPagar
                                curMontoMovimientoMN = curCtaXPagar
                                
                            Case Codigo_CtaXCobrar
                                curMontoMovimientoMN = curCtaXCobrar
                                
                            Case Codigo_CtaInteresCorrido
                                curMontoMovimientoMN = curCtaInteresCorrido
                                
                            Case Codigo_CtaProvReajusteK
                                curMontoMovimientoMN = curCtaProvReajusteK
                                
                            Case Codigo_CtaReajusteK
                                curMontoMovimientoMN = curCtaReajusteK
                                
                            Case Codigo_CtaProvFlucMercado
                                curMontoMovimientoMN = curCtaProvFlucMercado
                                
                            Case Codigo_CtaFlucMercado
                                curMontoMovimientoMN = curCtaFlucMercado
                                
                            Case Codigo_CtaProvInteresVac
                                curMontoMovimientoMN = curCtaProvInteresVac
                                
                            Case Codigo_CtaInteresVac
                                curMontoMovimientoMN = curCtaInteresVac
                                
                            Case Codigo_CtaIntCorridoK
                                curMontoMovimientoMN = curCtaIntCorridoK
                                
                            Case Codigo_CtaProvFlucK
                                curMontoMovimientoMN = curCtaProvFlucK
                                
                            Case Codigo_CtaFlucK
                                curMontoMovimientoMN = curCtaFlucK
                                
                            Case Codigo_CtaInversionTransito
                                curMontoMovimientoMN = curCtaInversionTransito
                                
                        End Select
                        
                        strIndDebeHaber = Trim(adoConsulta("IndDebeHaber"))
                        If strIndDebeHaber = "H" Then
                            curMontoMovimientoMN = curMontoMovimientoMN * -1
                            If curMontoMovimientoMN > 0 Then strIndDebeHaber = "D"
                        ElseIf strIndDebeHaber = "D" Then
                            If curMontoMovimientoMN < 0 Then strIndDebeHaber = "H"
                        End If
                        
                        If strIndDebeHaber = "T" Then
                            If curMontoMovimientoMN > 0 Then
                                strIndDebeHaber = "D"
                            Else
                                strIndDebeHaber = "H"
                            End If
                        End If
                        strDescripMovimiento = Trim(adoConsulta("DescripDinamica"))
                        curMontoMovimientoME = 0
                        curMontoContable = curMontoMovimientoMN
            
                        If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
                            curMontoContable = Round(curMontoMovimientoMN * dblTipoCambioCierre, 2)
                            curMontoMovimientoME = curMontoMovimientoMN
                            curMontoMovimientoMN = 0
                        End If
                                    
                        strTipoDocumento = ""
                        strNumDocumento = ""
                        strTipoPersonaContraparte = ""
                        strCodPersonaContraparte = ""
                        strIndContracuenta = ""
                        strCodContracuenta = ""
                        strCodFileContracuenta = ""
                        strCodAnaliticaContracuenta = ""
                        strIndUltimoMovimiento = ""
                                    
                        '*** Movimiento ***
                        .CommandText = "{ call up_ACAdicAsientoContableDetalle('"
                        .CommandText = .CommandText & strNumAsiento & "','" & strCodFondo & "','" & _
                            gstrCodAdministradora & "'," & _
                            CInt(adoConsulta("NumSecuencial")) & ",'" & _
                            strFechaGrabar & "','" & _
                            Left(strFechaSiguiente, 4) & "','" & _
                            Mid(strFechaSiguiente, 5, 2) & "','" & _
                            strDescripMovimiento & "','" & _
                            strIndDebeHaber & "','" & _
                            Trim(adoConsulta("CodCuenta")) & "','" & _
                            Trim(adoRegistro("CodMoneda")) & "'," & _
                            CDec(curMontoMovimientoMN) & "," & _
                            CDec(curMontoMovimientoME) & "," & _
                            CDec(curMontoContable) & ",'" & _
                            Trim(adoRegistro("CodFile")) & "','" & _
                            Trim(adoRegistro("CodAnalitica")) & "','','') }"
                        adoConn.Execute .CommandText
                    
                        '*** Saldos ***
'                        .CommandText = "{ call up_ACGenPartidaContableSaldos('"
'                        .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & _
'                            Left(strFechaSiguiente, 4) & "','" & Mid(strFechaSiguiente, 5, 2) & "','" & _
'                            Trim(adoConsulta("CodCuenta")) & "','" & _
'                            Trim(adoRegistro("CodFile")) & "','" & _
'                            Trim(adoRegistro("CodAnalitica")) & "','" & _
'                            strFechaSiguiente & "','" & _
'                            strFechaSubSiguiente & "'," & _
'                            CDec(curMontoMovimientoMN) & "," & _
'                            CDec(curMontoMovimientoME) & "," & _
'                            CDec(curMontoContable) & ",'" & _
'                            strIndDebeHaber & "','" & _
'                            Trim(adoRegistro("CodMoneda")) & "') }"
'                        adoConn.Execute .CommandText
                                        
                        '*** Validar valor de cuenta contable ***
                        If Trim(adoConsulta("CodCuenta")) = Valor_Caracter Then
                            MsgBox "Registro Nro. " & CStr(intContador) & " de Asiento Contable no tiene cuenta asignada", vbCritical, "Valorización Depósitos"
                            gblnRollBack = True
                            Exit Sub
                        End If
                        
                        adoConsulta.MoveNext
                    Loop
                    adoConsulta.Close: Set adoConsulta = Nothing
                                    
                    '*** Orden de Cobro ***
                    adoConn.Execute strSQLOrdenCaja
                    adoConn.Execute strSQLOrdenCajaDetalle
        
                    '*** Actualizar Secuenciales **
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumComprobante & "','" & strNumAsiento & "') }"
                    adoConn.Execute .CommandText
                    
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumOperacion & "','" & strNumOperacion & "') }"
                    adoConn.Execute .CommandText
                    
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumKardex & "','" & strNumKardex & "') }"
                    adoConn.Execute .CommandText
                    
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumOrdenCaja & "','" & strNumCaja & "') }"
                    adoConn.Execute .CommandText
                                                                            
'                    .CommandText = "COMMIT TRANSACTION ProcAsiento"
'                    adoConn.Execute .CommandText
            
                End If
            
            End If

            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
                
    End With
    
    Exit Sub
       
Ctrl_Error:
'    adoComm.CommandText = "ROLLBACK TRANSACTION ProcAsiento"
'    adoConn.Execute adoComm.CommandText
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
    
End Sub

Private Sub VencimientoReporteRentaVariable()

    Dim adoRegistro             As ADODB.Recordset, adoConsulta     As ADODB.Recordset
    Dim dblPrecioCierre         As Double, dblPrecioPromedio        As Double
    Dim dblTirCierre            As Double, dblFactorDiario          As Double
    Dim dblTasaInteres          As Double, dblFactorDiarioCupon     As Double
    Dim dblPrecioUnitario       As Double, dblValorPromedioKardex   As Double
    Dim dblInteresCorridoPromedio As Double, dblTirOperacionKardex  As Double
    Dim dblTirPromedioKardex    As Double, dblTirNetaKardex         As Double
    Dim curSaldoInversion       As Currency, curSaldoInteresCorrido As Currency
    Dim curSaldoFluctuacion     As Currency, curValorAnterior       As Currency
    Dim curValorActual          As Currency, curMontoRenta          As Currency
    Dim curMontoContable        As Currency, curMontoMovimientoMN   As Currency
    Dim curMontoMovimientoME    As Currency, curSaldoValorizar      As Currency
    Dim curCantMovimiento       As Currency, curKarValProm          As Currency
    Dim curValorMovimiento      As Currency, curSaldoInicialKardex  As Currency
    Dim curSaldoFinalKardex     As Currency, curValorSaldoKardex    As Currency
    Dim curValComi              As Currency, curVacCorrido          As Currency
    Dim curSaldoAmortizacion    As Currency, curSaldoGPCapital      As Currency
    Dim curSaldoFluctuacionMercado  As Currency
    Dim intCantRegistros        As Integer, intContador             As Integer
    Dim intRegistro             As Integer, intBaseCalculo          As Integer
    Dim intDiasPlazo            As Integer, intDiasDeRenta          As Integer
    Dim strNumAsiento           As String, strDescripAsiento        As String
    Dim strNumOperacion         As String, strNumKardex             As String
    Dim strNumCaja              As String, strFechaGrabar           As String
    Dim strCodTitulo            As String, strCodEmisor             As String
    Dim strDescripMovimiento    As String, strIndDebeHaber          As String
    Dim strCodCuenta            As String, strFiles                 As String
    Dim strCodFile              As String, strModalidadInteres      As String
    Dim strCodTasa              As String, strIndCuponCero          As String
    Dim strCodDetalleFile       As String, strCodAnalitica          As String
    Dim strCodSubDetalleFile    As String, strIndGasto              As String
    Dim strFechaPago            As String
    Dim strSQLOperacion         As String, strSQLKardex             As String
    Dim strSQLOrdenCaja         As String, strSQLOrdenCajaDetalle   As String
    Dim strIndUltimoMovimiento  As String, strTipoMovimientoKardex  As String
    Dim blnVenceTitulo          As Boolean, blnVenceCupon           As Boolean
    Dim dblTipoCambioCierre     As Double

    '*** Verificación de Vencimiento de Valores de Depósito ***
    frmMainMdi.stbMdi.Panels(3).Text = "Verificando Vencimiento de Valores de Renta Fija Corto Plazo..."
       
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT IK.CodTitulo,II.CodMoneda,IK.CodAnalitica,IK.CodFile,SaldoFinal,CodDetalleFile,CodSubDetalleFile," & _
            "ValorNominal,CodTipoTasa,BaseAnual,CodTasa,TasaInteres,DiasPlazo,IndCuponCero,FechaEmision,FechaVencimiento," & _
            "II.CodEmisor,IK.PrecioUnitario,IK.MontoMovimiento,IK.SaldoInteresCorrido,IK.MontoSaldo,IK.ValorPromedio " & _
            "FROM InversionKardex IK JOIN InstrumentoInversion II " & _
            "ON(II.CodTitulo=IK.CodTitulo) " & _
            "WHERE IK.CodFile IN ('008') AND FechaOperacion <='" & strFechaCierre & "' AND SaldoFinal > 0 AND " & _
            "IK.CodAdministradora='" & gstrCodAdministradora & "' AND IK.CodFondo='" & strCodFondo & "' AND IndUltimoMovimiento='X'"
        Set adoRegistro = .Execute
    
        Do Until adoRegistro.EOF
            blnVenceTitulo = False: blnVenceCupon = False
            
            '*** Obtener Secuenciales ***
            strNumAsiento = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumComprobante)
            strNumOperacion = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOperacion)
            strNumKardex = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumKardex)
            strNumCaja = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOrdenCaja)
            
            '*** Fecha de Vencimiento del Título = Fecha de Cierre Más 1 Día ***
            If Convertyyyymmdd(adoRegistro("FechaVencimiento")) = strFechaSiguiente Then blnVenceTitulo = True
            
            '*** Si vence el título o el cupón ***
            If blnVenceTitulo Or blnVenceCupon Then
                strCodTitulo = Trim(adoRegistro("CodTitulo"))
                strCodFile = Trim(adoRegistro("CodFile"))
                strCodDetalleFile = Trim(adoRegistro("CodDetalleFile"))
                strCodSubDetalleFile = Trim(adoRegistro("CodSubDetalleFile"))
                strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
                strCodEmisor = Trim(adoRegistro("CodEmisor"))
                strModalidadInteres = Trim(adoRegistro("CodDetalleFile"))
                strCodTasa = Trim(adoRegistro("CodTasa"))
                dblTasaInteres = CDbl(adoRegistro("TasaInteres"))
                intDiasPlazo = CInt(adoRegistro("DiasPlazo"))
                strIndCuponCero = Trim(adoRegistro("IndCuponCero"))
                curSaldoValorizar = CCur(adoRegistro("SaldoFinal"))
                curKarValProm = CDbl(adoRegistro("ValorPromedio"))
                intDiasDeRenta = DateDiff("d", CVDate(adoRegistro("FechaEmision")), gdatFechaActual) + 1
            
                Set adoConsulta = New ADODB.Recordset
                        
                '*** Verificar Dinamica Contable ***
                .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
                    "WHERE TipoOperacion='" & Codigo_Dinamica_Vencimiento & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                    IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"

                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    If CInt(adoConsulta("NumRegistros")) > 0 Then
                        intCantRegistros = CInt(adoConsulta("NumRegistros"))
                    Else
                        MsgBox "NO EXISTE Dinámica Contable para el Vencimiento", vbCritical
                        adoConsulta.Close: Set adoConsulta = Nothing
                        Exit Sub
                    End If
                End If
                adoConsulta.Close
                
                '*** Obtener la Fecha de Pago ***
                .CommandText = "SELECT FechaPago FROM InstrumentoInversionCalendario " & _
                    "WHERE CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    strFechaPago = Convertyyyymmdd(adoConsulta("FechaPago"))
                End If
                adoConsulta.Close
            
                strIndGasto = Valor_Caracter
                
                .CommandText = "SELECT IndPorcenPrecio,IndGasto FROM InversionFile " & _
                    "WHERE CodFile='" & Trim(adoRegistro("CodFile")) & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    strIndGasto = Trim(adoConsulta("IndGasto"))
                End If
                adoConsulta.Close
            
                '*** Obtener las cuentas de inversión ***
                Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, Trim(adoRegistro("CodMoneda")))
                
                '*** Obtener tipo de cambio ***
                'dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
                dblTipoCambioCierre = ObtenerValorTipoCambio(adoRegistro("CodMoneda"), Codigo_Moneda_Local, strFechaCierre, strFechaCierre, gstrCodClaseTipoCambioOperacionFondo, Codigo_Valor_TipoCambioCompra)

                
                '*** Obtener Saldo de Inversión ***
                If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaInversion & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curSaldoInversion = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                '*** Obtener Saldo de Interés Corrido ***
                If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaInteresCorrido & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curSaldoInteresCorrido = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                '*** Obtener Saldo de Provisión ***
                If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaProvInteres & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curSaldoFluctuacion = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                '*** Obtener Saldo de Provisión G/P Capital ***
                If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                                
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaProvFlucK & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curSaldoGPCapital = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                '*** Obtener Saldo de Fluctuación Mercado ***
                If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                                
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaProvFlucMercado & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curSaldoFluctuacionMercado = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                '*** Obtener Saldo Costo SAB ***
                If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaCostoSAB & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curCtaCostoSAB = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                '*** Obtener Saldo Costo BVL ***
                If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaCostoBVL & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curCtaCostoBVL = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                '*** Obtener Saldo Costo Cavali ***
                If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaCostoCavali & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curCtaCostoCavali = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                '*** Obtener Saldo Costo Conasev ***
                If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaCostoConasev & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curCtaCostoConasev = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                '*** Obtener Saldo Costo Fondo Garantía ***
                If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaCostoFondoGarantia & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curCtaCostoFondoGarantia = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                '*** Obtener Saldo IGV ***
                If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                    .CommandText = "SELECT SaldoFinalContable Saldo "
                Else
                    .CommandText = "SELECT SaldoFinalME Saldo "
                End If
                .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                    "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                    "CodCuenta='" & strCtaImpuesto & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                    "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    curCtaImpuesto = CDbl(adoConsulta("Saldo"))
                End If
                adoConsulta.Close
                
                If blnVenceTitulo Then
                    '*** Calculos ***
                    If strIndGasto = Valor_Indicador Then
                        curCtaXCobrar = Round(curSaldoInversion + curSaldoInteresCorrido + curSaldoFluctuacion + curCtaCostoBVL + curCtaCostoCavali + curCtaCostoConasev + curCtaCostoFondoGarantia + curCtaCostoSAB + curCtaImpuesto, 2)
                    Else
                        curCtaXCobrar = Round(curSaldoInversion + curSaldoInteresCorrido + curSaldoFluctuacion, 2)
                    End If
                    
                    curCtaInversion = curSaldoInversion
                    curCtaCosto = curSaldoInversion
                    curCtaInteresCorrido = curSaldoInteresCorrido
                    curCtaProvInteres = curSaldoFluctuacion
                    curCtaIngresoOperacional = curCtaXCobrar - curCtaProvInteres
                    
                    curCtaProvFlucK = curSaldoGPCapital
                    curCtaFlucK = curSaldoGPCapital * -1
                    curCtaProvFlucMercado = curSaldoFluctuacionMercado
                    curCtaFlucMercado = curSaldoFluctuacionMercado * -1
                    
                    curCantMovimiento = CCur(adoRegistro("SaldoFinal"))
                    dblPrecioUnitario = curKarValProm
                    curValorMovimiento = Round(dblPrecioUnitario * curCantMovimiento * -1, 2)
                    curSaldoInicialKardex = CCur(adoRegistro("SaldoFinal"))
                    curSaldoFinalKardex = 0
                    curValorSaldoKardex = 0
                    dblValorPromedioKardex = 0
                    dblInteresCorridoPromedio = 0
                    curValComi = 0
                    curVacCorrido = 0
                    dblTirOperacionKardex = 0
                    dblTirPromedioKardex = 0
                    dblTirNetaKardex = 0
                    curSaldoAmortizacion = 0
                End If
                
                '************************
                '*** Armar sentencias ***
                '************************
                strDescripAsiento = "Vencimiento" & Space(1) & "(" & strCodFile & "-" & strCodAnalitica & ")"
                '*** Operación ***
                strSQLOperacion = "{ call up_IVAdicInversionOperacion('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strNumOperacion & "','" & strFechaSiguiente & "','" & strCodTitulo & "','" & Left(strFechaSiguiente, 4) & "','" & _
                    Mid(strFechaSiguiente, 5, 2) & "','','" & Estado_Activo & "','" & strCodAnalitica & "','" & _
                    strCodFile & "','" & strCodAnalitica & "','" & strCodDetalleFile & "','" & strCodSubDetalleFile & "','" & _
                    Codigo_Caja_Vencimiento & "','','','" & strDescripAsiento & "','" & strCodEmisor & "','" & _
                    "','','','" & strFechaSiguiente & "','" & strFechaSiguiente & "','" & _
                    strFechaSiguiente & "','" & adoRegistro("CodMoneda") & "'," & CDec(adoRegistro("SaldoFinal")) & "," & CDec(gdblTipoCambio) & "," & _
                    CDec(adoRegistro("ValorNominal")) & "," & CDec(adoRegistro("PrecioUnitario")) & "," & CDec(adoRegistro("MontoMovimiento")) & "," & CDec(adoRegistro("SaldoInteresCorrido")) & "," & _
                    "0,0,0,0,0,0,0," & CDec(curCtaXCobrar) & ",0,0,0,0,0,0,0,0,0," & _
                    "0,0,0,0,'X','" & strNumAsiento & "','','','" & _
                    "','','','',0,'','','','',''," & CDec(dblTasaInteres) & "," & _
                    "0,0,'','','','" & gstrLogin & "') }"
                
                strIndUltimoMovimiento = "X"
                strTipoMovimientoKardex = "S"
                '*** Kardex ***
                strSQLKardex = "{ call up_IVAdicInversionKardex('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strCodTitulo & "','" & strNumKardex & "','" & strFechaSiguiente & "','" & Left(strFechaSiguiente, 4) & "','" & _
                    Mid(strFechaSiguiente, 5, 2) & "','','" & strNumOperacion & "','" & strCodEmisor & "','','O','" & _
                    strFechaSiguiente & "','" & strTipoMovimientoKardex & "','O'," & curCantMovimiento & ",'" & adoRegistro("CodMoneda") & "'," & _
                    dblPrecioUnitario & "," & curValorMovimiento & "," & curValComi & "," & curSaldoInicialKardex & "," & _
                    curSaldoFinalKardex & "," & curValorSaldoKardex & ",'" & strDescripAsiento & "'," & dblValorPromedioKardex & ",'" & _
                    strIndUltimoMovimiento & "','" & strCodFile & "','" & strCodAnalitica & "'," & dblInteresCorridoPromedio & "," & _
                    curSaldoInteresCorrido & "," & dblTirOperacionKardex & "," & dblTirPromedioKardex & "," & curVacCorrido & "," & _
                    dblTirNetaKardex & "," & curSaldoAmortizacion & ") }"
    
                '*** Orden de Cobro/Pago ***
                strSQLOrdenCaja = "{ call up_ACAdicMovimientoFondo('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strNumCaja & "','" & strFechaSiguiente & "','" & Trim(frmMainMdi.Tag) & "','" & strNumOperacion & "','" & strFechaPago & "','" & _
                    strNumAsiento & "','','E','" & strCtaXCobrar & "'," & CDec(curCtaXCobrar) & ",'" & _
                    strCodFile & "','" & strCodAnalitica & "','" & adoRegistro("CodMoneda") & "','" & _
                    strDescripAsiento & "','" & Codigo_Caja_Vencimiento & "','','" & Estado_Caja_NoConfirmado & "','','','','','','','','','','" & gstrLogin & "') }"
                
                '*** Orden de Cobro/Pago Detalle ***
                strSQLOrdenCajaDetalle = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strNumCaja & "','" & strFechaSiguiente & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripAsiento & "','" & _
                    "H','" & strCtaXCobrar & "'," & CDec(curCtaXCobrar) * -1 & ",'" & _
                    strCodFile & "','" & strCodAnalitica & "','" & adoRegistro("CodMoneda") & "','') }"
                                
                '*** Monto Orden ***
                If curCtaXCobrar > 0 Then
                                                                                            
                    On Error GoTo Ctrl_Error
                    
'                    .CommandText = "BEGIN TRANSACTION ProcAsiento"
'                    adoConn.Execute .CommandText
                                                            
                    '*** Actualizar indicador de último movimiento en Kardex ***
                    .CommandText = "UPDATE InversionKardex SET IndUltimoMovimiento='' " & _
                        "WHERE CodAnalitica='" & strCodAnalitica & "' AND CodFile='" & strCodFile & "' AND " & _
                        "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                        "IndUltimoMovimiento='X'"
        
                    adoConn.Execute .CommandText
        
                    '*** Inserta movimiento en el kardex ***
                    adoConn.Execute strSQLKardex
                    
                    '*** Operación ***
                    adoConn.Execute strSQLOperacion
                    
                    '*** Contabilizar ***
                    strFechaGrabar = strFechaSiguiente & Space(1) & Format(Time, "hh:mm")
                    
                    '*** Cabecera ***
                    .CommandText = "{ call up_ACAdicAsientoContable('"
                    .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
                        strFechaGrabar & "','" & _
                        Left(strFechaSiguiente, 4) & "','" & Mid(strFechaSiguiente, 5, 2) & "','" & _
                        "','" & _
                        strDescripAsiento & "','" & Trim(adoRegistro("CodMoneda")) & "','" & _
                        Codigo_Moneda_Local & "','" & _
                        "','" & _
                        "'," & _
                        CDec(curCtaXCobrar) & ",'" & Estado_Activo & "'," & _
                        intCantRegistros & ",'" & _
                        strFechaSiguiente & Space(1) & Format(Time, "hh:ss") & "','" & _
                        strCodModulo & "','" & _
                        "'," & _
                        dblTipoCambioCierre & ",'" & _
                        "','" & _
                        "','" & _
                        strDescripAsiento & "','" & _
                        "','" & _
                        "X','') }"
                    adoConn.Execute .CommandText
                    
                    '*** Detalle ***
                    .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
                        "WHERE TipoOperacion='" & Codigo_Dinamica_Vencimiento & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                        strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                        IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"

                    Set adoConsulta = .Execute
            
                    Do While Not adoConsulta.EOF
                    
                        Select Case Trim(adoConsulta("TipoCuentaInversion"))
                            Case Codigo_CtaInversion
                                curMontoMovimientoMN = curCtaInversion
                                
                            Case Codigo_CtaProvInteres
                                curMontoMovimientoMN = curCtaProvInteres
                                
                            Case Codigo_CtaInteres
                                curMontoMovimientoMN = curCtaInteres
                                
                            Case Codigo_CtaCosto
                                curMontoMovimientoMN = curCtaCosto
                                
                            Case Codigo_CtaIngresoOperacional
                                curMontoMovimientoMN = curCtaIngresoOperacional
                                
                            Case Codigo_CtaInteresVencido
                                curMontoMovimientoMN = curCtaInteresVencido
                                
                            Case Codigo_CtaVacCorrido
                                curMontoMovimientoMN = curCtaVacCorrido
                                
                            Case Codigo_CtaXPagar
                                curMontoMovimientoMN = curCtaXPagar
                                
                            Case Codigo_CtaXCobrar
                                curMontoMovimientoMN = curCtaXCobrar
                                
                            Case Codigo_CtaInteresCorrido
                                curMontoMovimientoMN = curCtaInteresCorrido
                                
                            Case Codigo_CtaProvReajusteK
                                curMontoMovimientoMN = curCtaProvReajusteK
                                
                            Case Codigo_CtaReajusteK
                                curMontoMovimientoMN = curCtaReajusteK
                                
                            Case Codigo_CtaProvFlucMercado
                                curMontoMovimientoMN = curCtaProvFlucMercado
                                
                            Case Codigo_CtaFlucMercado
                                curMontoMovimientoMN = curCtaFlucMercado
                                
                            Case Codigo_CtaProvInteresVac
                                curMontoMovimientoMN = curCtaProvInteresVac
                                
                            Case Codigo_CtaInteresVac
                                curMontoMovimientoMN = curCtaInteresVac
                                
                            Case Codigo_CtaIntCorridoK
                                curMontoMovimientoMN = curCtaIntCorridoK
                                
                            Case Codigo_CtaProvFlucK
                                curMontoMovimientoMN = curCtaProvFlucK
                                
                            Case Codigo_CtaFlucK
                                curMontoMovimientoMN = curCtaFlucK
                                
                            Case Codigo_CtaInversionTransito
                                curMontoMovimientoMN = curCtaInversionTransito
                        End Select
                        
                        strIndDebeHaber = Trim(adoConsulta("IndDebeHaber"))
                        If strIndDebeHaber = "H" Then
                            curMontoMovimientoMN = curMontoMovimientoMN * -1
                            If curMontoMovimientoMN > 0 Then strIndDebeHaber = "D"
                        ElseIf strIndDebeHaber = "D" Then
                            If curMontoMovimientoMN < 0 Then strIndDebeHaber = "H"
                        End If
                        
                        If strIndDebeHaber = "T" Then
                            If curMontoMovimientoMN > 0 Then
                                strIndDebeHaber = "D"
                            Else
                                strIndDebeHaber = "H"
                            End If
                        End If
                        strDescripMovimiento = Trim(adoConsulta("DescripDinamica"))
                        curMontoMovimientoME = 0
                        curMontoContable = curMontoMovimientoMN
            
                        If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
                            curMontoContable = Round(curMontoMovimientoMN * dblTipoCambioCierre, 2)
                            curMontoMovimientoME = curMontoMovimientoMN
                            curMontoMovimientoMN = 0
                        End If
                                    
                        '*** Movimiento ***
                        .CommandText = "{ call up_ACAdicAsientoContableDetalle('"
                        .CommandText = .CommandText & strNumAsiento & "','" & strCodFondo & "','" & _
                            gstrCodAdministradora & "'," & _
                            CInt(adoConsulta("NumSecuencial")) & ",'" & _
                            strFechaGrabar & "','" & _
                            Left(strFechaSiguiente, 4) & "','" & _
                            Mid(strFechaSiguiente, 5, 2) & "','" & _
                            strDescripMovimiento & "','" & _
                            strIndDebeHaber & "','" & _
                            Trim(adoConsulta("CodCuenta")) & "','" & _
                            Trim(adoRegistro("CodMoneda")) & "'," & _
                            CDec(curMontoMovimientoMN) & "," & _
                            CDec(curMontoMovimientoME) & "," & _
                            CDec(curMontoContable) & ",'" & _
                            Trim(adoRegistro("CodFile")) & "','" & _
                            Trim(adoRegistro("CodAnalitica")) & "','','') }"
                        adoConn.Execute .CommandText
                    
                        '*** Saldos ***
'                        .CommandText = "{ call up_ACGenPartidaContableSaldos('"
'                        .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & _
'                            Left(strFechaSiguiente, 4) & "','" & Mid(strFechaSiguiente, 5, 2) & "','" & _
'                            Trim(adoConsulta("CodCuenta")) & "','" & _
'                            Trim(adoRegistro("CodFile")) & "','" & _
'                            Trim(adoRegistro("CodAnalitica")) & "','" & _
'                            strFechaSiguiente & "','" & _
'                            strFechaSubSiguiente & "'," & _
'                            CDec(curMontoMovimientoMN) & "," & _
'                            CDec(curMontoMovimientoME) & "," & _
'                            CDec(curMontoContable) & ",'" & _
'                            strIndDebeHaber & "','" & _
'                            Trim(adoRegistro("CodMoneda")) & "') }"
'                        adoConn.Execute .CommandText
                                        
                        '*** Validar valor de cuenta contable ***
                        If Trim(adoConsulta("CodCuenta")) = Valor_Caracter Then
                            MsgBox "Registro Nro. " & CStr(intContador) & " de Asiento Contable no tiene cuenta asignada", vbCritical, "Valorización Depósitos"
                            gblnRollBack = True
                            Exit Sub
                        End If
                        
                        adoConsulta.MoveNext
                    Loop
                    adoConsulta.Close: Set adoConsulta = Nothing
                                    
                    '*** Orden de Cobro ***
                    adoConn.Execute strSQLOrdenCaja
                    adoConn.Execute strSQLOrdenCajaDetalle
        
                    '*** Actualizar Secuenciales **
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumComprobante & "','" & strNumAsiento & "') }"
                    adoConn.Execute .CommandText
                    
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumOperacion & "','" & strNumOperacion & "') }"
                    adoConn.Execute .CommandText
                    
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumKardex & "','" & strNumKardex & "') }"
                    adoConn.Execute .CommandText
                    
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumOrdenCaja & "','" & strNumCaja & "') }"
                    adoConn.Execute .CommandText
                                                                            
'                    .CommandText = "COMMIT TRANSACTION ProcAsiento"
'                    adoConn.Execute .CommandText
            
                End If
            
            End If

            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
                
    End With
    
    Exit Sub
       
Ctrl_Error:
'    adoComm.CommandText = "ROLLBACK TRANSACTION ProcAsiento"
'    adoConn.Execute adoComm.CommandText
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
    
End Sub



Private Sub VerVenLetPag()
    
End Sub




Private Sub lblRentabilidad_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblRentabilidad(Index), Decimales_Tasa)
    
End Sub

Private Sub lblValorAIR_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblValorAIR(Index), Decimales_ValorCuota_Cierre)
    
End Sub

Private Sub lblValorDIR_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblValorDIR(Index), Decimales_ValorCuota_Cierre)
    
End Sub

Private Sub BuscarTipoCambio()

    Dim strSQL As String
    Dim strFechaConsulta As String
    
    strFechaConsulta = Convertyyyymmdd(dtpFechaCierreHasta.Value)

    strSQL = "{ call up_ACObtieneTipoCambioFecha('" & gstrCodClaseTipoCambioFondo & "','" & strFechaConsulta & "') }"
                                            
                      
    With adoTipoCambioCierre
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With
        
    tdgTipoCambioCierre.Refresh
    
End Sub
Private Sub BuscarFondoSeries(strTipoCierre As String)

    Dim strSQL As String
   
    Dim adoRegistro As ADODB.Recordset
   
    Set adoRegistro = New ADODB.Recordset
    
    strSQL = "{ call up_GNObtenerFondoSerieValorCuotaActual('" & _
        strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaCierre & "','" & strTipoCierre & "')}"
    
    With adoRegistro
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    dbgFondoSeries.DataSource = adoRegistro
    
End Sub

Private Function ObtenerValorTCCierre(ByVal strpCodMoneda As String) As Double

    ObtenerValorTCCierre = 0
    
    If adoTipoCambioCierre.Recordset.EOF And adoTipoCambioCierre.Recordset.BOF Then Exit Function
    
    adoTipoCambioCierre.Recordset.MoveFirst
'    adoTipoCambioCierre.Recordset.Find ("CodMoneda='" & strpCodMoneda & "'")
    
    If strpCodMoneda = Codigo_Moneda_Local Or IsNull(tdgTipoCambioCierre.Columns(2).Value) Then
        ObtenerValorTCCierre = 1
    Else
        ObtenerValorTCCierre = CDbl(tdgTipoCambioCierre.Columns(2).Value)
    End If
    
End Function
Sub ActualizarFechasCierre(datFecha As Date)

    dtpFechaEntrega.Value = DateAdd("d", gintDiasPagoRescate, datFecha)
                
    strFechaCierre = Convertyyyymmdd(datFecha)
    gstrPeriodoActual = Format(Year(datFecha), "0000")
    gstrMesActual = Format(Month(datFecha), "00")
    gstrDiaActual = Format(Day(datFecha), "00")
    strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, datFecha))
    strFechaAnterior = Convertyyyymmdd(DateAdd("d", -1, datFecha))
    strFechaAnteAnterior = Convertyyyymmdd(DateAdd("d", -2, datFecha))
    strFechaSubSiguiente = Convertyyyymmdd(DateAdd("d", 2, datFecha))
    
    gdatFechaActual = datFecha
    gstrFechaActual = Convertyyyymmdd(gdatFechaActual)

    lblDescrip(6).Caption = CStr(datFecha)

End Sub

Sub CierreDiario()

     '*** Inicializar Variables de Trabajo ***
    Dim TimeInip                As Variant, TimeFinp                    As Variant
    Dim adoFondo                As ADODB.Recordset, adoConsulta         As ADODB.Recordset
    Dim adoAuxiliar             As ADODB.Recordset
    Dim strMensaje              As String, strIndPagoParcial            As String
    Dim lngNumCom               As Long, lngNumCaj                      As Long
    Dim lngNumEnt               As Long, lngNumKar                      As Long
    Dim lngNumOpe               As Long
    Dim dblSaldoTotal           As Double, dblTasa                      As Double
    Dim strCodVariable          As String
    Dim strIndCobrar            As String
    Dim dblTCCierre             As Double
    Dim strTablaFondo           As String
    Dim strCodigoCierre         As String
        
    '*** Cierre Diario ***
    If TodoOK() And VerificaOrdenPendienteFacturacion(strCodFondo) Then
        frmMainMdi.stbMdi.Panels(3).Text = "Acción..."
        
        '*** Pedir Confirmación de Datos ***
        If chkSimulacion.Value Then
            strMensaje = "Para el proceso de SIMULACION confirme lo siguiente : " & vbNewLine
        Else
            strMensaje = "Para el proceso de CIERRE confirme lo siguiente: " & vbNewLine
        End If
        
        strMensaje = strMensaje & " Fondo >> " & Trim(cboFondo.Text) & vbNewLine & _
            " Fecha Inicial >> " & CStr(dtpFechaCierre.Value) & vbNewLine & _
            " Fecha Final >> " & CStr(dtpFechaCierreHasta.Value) & vbNewLine & _
            "¿ Seguro de continuar ?"
        If MsgBox(strMensaje, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then  'No desea continuar
            GoTo cmdCierre_fin
        End If
                        
        '*** Inicio del proceso ***
        If chkSimulacion.Value Then
            '*** Inicio de la SIMULACION ***
            frmMainMdi.stbMdi.Panels(3).Text = "Iniciando Simulación del Valor Cuota..."
            strCodigoCierre = Codigo_Cierre_Simulacion
        Else
            frmMainMdi.stbMdi.Panels(3).Text = "Inicio del Proceso..."
            strCodigoCierre = Codigo_Cierre_Definitivo
        End If
                
Fondo_Reproceso:
                
        Set adoFondo = New ADODB.Recordset
                
        '*** Tipo de Cambio ***
        adoComm.CommandText = "SELECT CodMoneda FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND " & _
            "CodFondo='" & strCodFondo & "'"
        Set adoFondo = adoComm.Execute
    
        strIndCobrar = Valor_Caracter
        If Not adoFondo.EOF Then
            Me.MousePointer = vbDefault
            
            '*** Obtener tipo de cambio ***
            'dblTCCierre = ObtenerValorTCCierre(adoFondo("CodMoneda"))
            dblTCCierre = ObtenerValorTipoCambio(adoFondo("CodMoneda"), Codigo_Moneda_Local, strFechaCierre, strFechaCierre, gstrCodClaseTipoCambioOperacionFondo, Codigo_Valor_TipoCambioCompra)
                        
            If dblTCCierre = 0 Then
                MsgBox "El Tipo de Cambio para la fecha de cierre NO ESTA REGISTRADO.", vbCritical, Me.Caption
                adoFondo.Close: Set adoFondo = Nothing
                Exit Sub
            End If
        End If
        adoFondo.Close
                        
        Me.MousePointer = vbHourglass
        
        strTablaFondo = "Fondo"
        
        '*** Prepara Tablas para la Simulación ***
        If chkSimulacion.Value Then
            frmMainMdi.stbMdi.Panels(3).Text = "Preparando Tablas para la Simulación..."
            
            strTablaFondo = "FondoTmp"
            
            adoComm.CommandText = "{ call up_GNProcPrepararTablasSimulacion('" & _
                strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strFechaCierre & "','" & strFechaSiguiente & "') }"
            adoConn.Execute adoComm.CommandText
            Sleep 0&
        End If
      
        
        Set adoFondo = Nothing: Set adoFondo = New ADODB.Recordset
       
        '*** Comisión de Administración ***
        adoComm.CommandText = "SELECT CodFondo FROM " & strTablaFondo & " WHERE CodAdministradora='" & gstrCodAdministradora & "' AND " & _
            "CodFondo='" & strCodFondo & "' AND (IndComision=NULL OR IndComision='')"
        Set adoFondo = adoComm.Execute
        
        strIndCobrar = Valor_Caracter
        
        If Not adoFondo.EOF Then
            Me.MousePointer = vbDefault
            If MsgBox("¿ Se calculará las comisiones de la empresa a partir de hoy ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                adoComm.CommandText = "UPDATE " & strTablaFondo & " SET FechaInicioEtapaOperativa='" & strFechaCierre & "',FechaPagoAdministradora='" & strFechaCierre & "', IndComision='X' " & _
                    "WHERE CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "'"
                adoConn.Execute adoComm.CommandText
               
                strIndCobrar = Valor_Indicador
                
            End If
        End If
        adoFondo.Close: Set adoFondo = Nothing
                
        cmdCierre.Enabled = False
        TimeInip = Time
        Me.Refresh
        Me.MousePointer = vbHourglass

        '*** Actualizar cuotas del fondo ***
        frmMainMdi.stbMdi.Panels(3).Text = "Actualizando Cantidad Final de Cuotas..."
        
        adoComm.CommandText = "{ call up_GNActCuotasCierre('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & _
            strFechaCierre & "','" & strFechaSiguiente & "','" & strCodigoCierre & "') }"
        
        adoConn.Execute adoComm.CommandText
        Sleep 0&

        
        '*** Rentabilidad de Depósitos de Ahorro ***
        frmMainMdi.stbMdi.Panels(3).Text = "Valuación de Depósitos de Ahorro..."
        
        adoComm.CommandText = "{ call up_GNProcValorizacionAhorros('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & _
            strFechaCierre & "','" & strFechaSiguiente & "','" & _
            Codigo_Tipo_Cuenta_Ahorro & "'," & Replace(dblTCCierre, ",", ".") & ",'" & _
            Codigo_Dinamica_Provision & "','" & strCodigoCierre & "') }"
        
        adoConn.Execute adoComm.CommandText
        Sleep 0&
                                                                                                                                
        '*** Rentabilidad de Ctas.Ctes. ***
        frmMainMdi.stbMdi.Panels(3).Text = "Valuación de Ctas.Ctes..."
        
        adoComm.CommandText = "{ call up_GNProcValorizacionCuentaCorriente('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & _
            strFechaCierre & "','" & strFechaSiguiente & "','" & _
            Codigo_Tipo_Cuenta_Corriente & "'," & Replace(dblTCCierre, ",", ".") & ",'" & _
            Codigo_Dinamica_Provision & "','" & strCodigoCierre & "') }"
        
        adoConn.Execute adoComm.CommandText
        Sleep 0&
        
        '*** Rentabilidad de Valores de Renta Variable ***
        frmMainMdi.stbMdi.Panels(3).Text = "Valuación de Acciones..."
        
        Call ValorizacionRentaVariable(strCodigoCierre)
        
        
        '*** Rentabilidad de Depósitos a Plazo y Certificados Bancarios ***
        frmMainMdi.stbMdi.Panels(3).Text = "Valuación de Valores de Depósitos..."
           
        Call ValorizacionDepositos(strCodigoCierre)
                                    
        
        '*** Rentabilidad de Bonos ***
        frmMainMdi.stbMdi.Panels(3).Text = "Valuación de Bonos..."
        
        Call ValorizacionRentaFijaLargoPlazo(strCodigoCierre)
                
        
        '*** Rentabilidad de Pactos ***
        frmMainMdi.stbMdi.Panels(3).Text = "Valuación de Pactos..."
        
        'FALTA REVISAR
        Call ValorizacionPacto(strCodigoCierre)
            
        '*** Rentabilidad de Operaciones de Reporte ***
        Call ValorizacionReportes(strCodigoCierre)
        
        '*** Rentabilidad de Valores de Renta Fija Corto Plazo ***
        frmMainMdi.stbMdi.Panels(3).Text = "Valuación de Valores de Renta Fija Corto Plazo..."

'        If chkSimulacion.Value Then 'FALTA REVISAR
'           Call ValorizacionRentaFijaCortoplazo(Codigo_Cierre_Simulacion)
'        Else
'           Call ValorizacionRentaFijaCortoplazo(Codigo_Cierre_Definitivo)
'        End If
        
        '*** Acreencias
        Call ValorizacionAcreencias(strCodigoCierre)
        
        '*** Flujos
        Call ValorizacionFlujos(strCodigoCierre)
        
        '*** Préstamos
        Call ValorizacionPrestamos(strCodigoCierre)

        '*** Rentabilidad de Valores Coberturados ***
        frmMainMdi.stbMdi.Panels(3).Text = "Valuación de Valores Coberturados..."

        'FALTA REVISAR
        Call ValorizacionCobertura(strCodigoCierre)
               
        Set adoFondo = New ADODB.Recordset
        
        '*** Extraer datos del fondo y calcular patrimonio y comisión adm. cartera ***
        adoComm.CommandText = "SELECT ValorCuotaNominal,TipoValuacion,IndComision,CodMoneda,CantPartesPagoSuscripcion FROM " & strTablaFondo & _
            " WHERE CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "'"
        Set adoFondo = adoComm.Execute
        
        dblValorCuotaNominal = CDbl(adoFondo("ValorCuotaNominal"))
        
        
        '*** Verifica y provisiona Comisiones de los comisionistas del fondo ***
        frmMainMdi.stbMdi.Panels(3).Text = "Verifica y provisiona Comisiones de los Comisionistas del Fondo..."
        
        adoComm.CommandText = "{ call up_GNCalcularDevengoComisionista('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & _
            strFechaCierre & "','" & strCodigoCierre & "') }"
        
        adoConn.Execute adoComm.CommandText
        Sleep 0&
        
        '*** Verifica y provisiona Comisiones de los comisionistas del fondo ***
        'Call ProvisionComisionesParticipacion(gstrCodFondoContable, strFechaCierre, strCodigoCierre)
        
        '*** Verifica y provisiona Comisiones de las Inversiones del fondo ***  pruebas REA 2015-04-27
        'Call ProvisionComisionesInversion(gstrCodFondoContable, strFechaCierre, strCodigoCierre)
        
        '*** Verifica y provisiona gastos del fondo - no incluye comision por adm de cartera ***
        'Call ProvisionGastosFondo(strCodigoCierre, Valor_Caracter)
        Call ProvisionGastosFondo(gstrCodFondoContable, "G", strFechaCierre, strFechaSiguiente, strCodMoneda, strCodModulo, strCodigoCierre, Valor_Caracter)
        
        '*** PreCierre - Traslado de SALDOS CONTABLES DE PRECIERRE >> PartidaContablePreSaldos ***
        frmMainMdi.stbMdi.Panels(3).Text = "Traslado de Saldos Al PreCierre..."
        
        adoComm.CommandText = "{ call up_GNProcTrasladoSaldosAlPreCierre('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & _
            strFechaCierre & "','" & strCodigoCierre & "') }"
        
        adoConn.Execute adoComm.CommandText
        Sleep 0&
                
        '*** Cálculo de Ganancias y Pérdidas al PreCierre ***
        frmMainMdi.stbMdi.Panels(3).Text = "Calculando Pérdidas y Ganancias al Precierre..."
        
        adoComm.CommandText = "{ call up_GNProcCalcGPAlPreCierre('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & _
            strFechaCierre & "','" & strFechaCierre & "'," & dblTCCierre & ",'" & strCodigoCierre & "') }"
        
        adoConn.Execute adoComm.CommandText
        Sleep 0&
        
        '*** Verifica y provisiona gastos del fondo - solo incluye comisiones que afectan balance de cierre ***
        Call ProvisionGastosFondo(gstrCodFondoContable, "G", strFechaCierre, strFechaSiguiente, strCodMoneda, strCodModulo, strCodigoCierre, Valor_Indicador)
        
        '*** Actualización Inicial de Saldos Finales y Monto de Ajuste Contable ***
        frmMainMdi.stbMdi.Panels(3).Text = "Actualizando Saldos Finales y Montos de Ajuste..."
        
        adoComm.CommandText = "{ call up_GNActMontoAjusteContable('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & _
            strFechaCierre & "','" & strFechaSiguiente & "','" & _
            gstrCodClaseTipoCambioFondo & "','" & _
            gstrValorTipoCambioCierre & "','" & strCodigoCierre & "') }"
        
        adoConn.Execute adoComm.CommandText
        Sleep 0&
                        
        '*** Asientos de Pérdida/Ganancia por Tipo de Cambio ***
        frmMainMdi.stbMdi.Panels(3).Text = "Registrando Asientos Contables por Ajuste en el Tipo de Cambio..."
        
        adoComm.CommandText = "{ call up_GNProcAjusteTipoCambio('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & _
            strFechaCierre & "','" & strFechaSiguiente & "'," & _
            Replace(dblTCCierre, ",", ".") & ",'" & gstrLogin & "','" & strCodigoCierre & "') }"
        
        adoConn.Execute adoComm.CommandText
        Sleep 0&
        
        '*** Cierre Pérdidas y Ganancias ***
         frmMainMdi.stbMdi.Panels(3).Text = "Calculando Resultados del Ejercicio al Cierre..."
        
        adoComm.CommandText = "{ call up_GNProcCalcGPAlCierre('" & _
                strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strFechaCierre & "','" & strFechaCierre & "'," & dblTCCierre & ",'" & strCodigoCierre & "') }"
        
        adoConn.Execute adoComm.CommandText
        Sleep 0&
        
        'NUEVA RUTINA DE CALCULO DE VALOR DE CUOTA
        '*** Calculo de Nuevo Patrimonio y Nueva Cuota ***
        frmMainMdi.stbMdi.Panels(3).Text = "Calculando Nuevo Patrimonio y Valor de Cuota..."
                
        strIndPagoParcial = Valor_Caracter
        If CInt(adoFondo("CantPartesPagoSuscripcion")) > 1 Then strIndPagoParcial = Valor_Indicador
              
        adoComm.CommandText = "{ call up_GNProcCalcValorCuota('" & _
                            strCodFondo & "','" & gstrCodAdministradora & "','" & _
                            strFechaCierre & "','" & strFechaCierre & "','" & _
                            CStr(adoFondo("IndComision")) & "','" & strIndPagoParcial & "','" & _
                            CStr(adoFondo("CodMoneda")) & "'," & dblTCCierre & ",'" & strCodigoCierre & "') }"
        
        adoConn.Execute adoComm.CommandText
        Sleep 0&
        
            
        Call BuscarFondoSeries(strCodigoCierre)
        
      
        'PROVISIONALMENTE ACA TERMINA LA SIMULACION
        If chkSimulacion.Value Then
            
            TimeFinp = Time
            frmMainMdi.stbMdi.Panels(3).Text = "Duración : " & Format((TimeFinp - TimeInip), "hh:mm:ss")
           
            MsgBox "Proceso de SIMULACION de Valor Cuota terminado Exitosamente.", vbInformation

            GoTo cmdCierre_fin
            
        End If
       
       
        '*** Realiza provision del IR a los participes ***
        frmMainMdi.stbMdi.Panels(3).Text = "Provisiona Impuesto a la Renta de los Participes..."

        '*** Realiza suscripciones a valor desconocido ***
        frmMainMdi.stbMdi.Panels(3).Text = "Procesa Operaciones de Participación a Valor Desconocido..."
                
        adoComm.CommandText = "{ call up_GNProcCierreParticipes('" & strCodFondo & "','" & gstrCodAdministradora & "'," & _
                    dblTCCierre & ",'" & strFechaCierre & "','" & Convertyyyymmdd(dtpFechaEntrega.Value) & "','" & _
                    gstrLogin & "','','" & strCodigoCierre & "') }"
        
        adoConn.Execute adoComm.CommandText
        
        
        'ASIENTO DE DES-ATRIBUCION SI APLICA
        frmMainMdi.stbMdi.Panels(3).Text = "Procesa Asiento de Des-atribucion..."

        '*** Actualizar patrimonio y activo en tabla de kardex de cuotas ***
        frmMainMdi.stbMdi.Panels(3).Text = "Actualizando Patrimonio y Activo Final..."
                
           
        adoComm.CommandText = "{ call up_GNActPatrimonioActivo('" & _
                strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strFechaCierre & "','" & strFechaCierre & "'," & dblTCCierre & ",'" & strCodMoneda & "','" & strCodigoCierre & "') }"
        
        adoConn.Execute adoComm.CommandText
        
        '*** Limites de Inversión ***
        'FALTA --ACR
        
        '*** Actualizar número de partícipes en tabla FondoValorCuota ***
        frmMainMdi.stbMdi.Panels(3).Text = "Actualizando número de partícipes..."
        
        adoComm.CommandText = "{ call up_GNActNumParticipes('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaCierre & "','" & strCodigoCierre & "') }"
        
        adoConn.Execute adoComm.CommandText
        
        Me.Refresh
            
        '*** Pase de Saldos Finales del día como Saldos Iniciales del día siguiente ***
        ' Esto se realiza siempre y cuando no estemos realizando el cierre del ultimo dia del periodo contable (año)
        frmMainMdi.stbMdi.Panels(3).Text = "Pasando Saldos Contables..."
        
        Set adoConsulta = New ADODB.Recordset
        
        adoComm.CommandText = "SELECT FechaFinal FROM PeriodoContable WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND PeriodoContable='" & gstrPeriodoActual & "' AND MesContable='99'"
        Set adoConsulta = adoComm.Execute
        If Not adoConsulta.EOF Then

            If Convertyyyymmdd(adoConsulta("FechaFinal")) <> strFechaCierre Then
                
                '*** Pase de saldos al dia siguiente
                adoComm.CommandText = "{ call up_GNProcTrasladoSaldosInicialesDiaSiguiente('"
                adoComm.CommandText = adoComm.CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strFechaCierre & "','" & Codigo_Cierre_Definitivo & "') }"
                adoConn.Execute adoComm.CommandText
                
                '*** Deshabilitar dia de hoy ***
                adoComm.CommandText = "{ call up_GNActIndFechaHabil('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaCierre & "'," & _
                    "'','" & Codigo_Cierre_Definitivo & "') }"
                adoConn.Execute adoComm.CommandText
            
                '*** Habilitar dia siguiente ***
                adoComm.CommandText = "{ call up_GNActIndFechaHabil('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaSiguiente & "'," & _
                    "'X','" & Codigo_Cierre_Definitivo & "') }"
                adoConn.Execute adoComm.CommandText
                
                '*** Actualiza kardex de cuotas al dia siguiente
                adoComm.CommandText = "{ call up_GNActKardexInicialCuotasFondo('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strFechaSiguiente & "'," & _
                    dblTCCierre & "," & dblTasa & ",'','','X','X','"
                adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Definitivo & "') }"
                adoConn.Execute adoComm.CommandText
                Sleep 0&
            
            End If
        End If
        
        adoConsulta.Close
                
        '*** Procesar Cartera Inversión ***
        adoComm.CommandText = "{ call up_GNGeneraCarteraInversion('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaCierre & "','" & _
            strFechaSiguiente & "') }"
        adoConn.Execute adoComm.CommandText
        
        '*** Calcular Duration ***
        Dim strCodTitulo        As String
        Dim curMontoSubTotal    As Currency, curMontoInteres    As Currency
        Dim curValorNominal     As Currency, curValorTitulo     As Currency
        Dim dblTasaMercado      As Double, dblDuration          As Double
        
        adoComm.CommandText = "SELECT * FROM InversionValorizacion " & _
            "WHERE (FechaValorizacion>='" & strFechaCierre & "' AND FechaValorizacion<'" & strFechaSiguiente & "') AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
            "CodFile='005'"
        Set adoConsulta = adoComm.Execute
        
        Do While Not adoConsulta.EOF
            strCodTitulo = adoConsulta("CodTitulo")
            curMontoSubTotal = CCur(adoConsulta("ValorInversion"))
            curMontoInteres = CCur(adoConsulta("ValorInteresCorrido")) + CCur(adoConsulta("ValorFluctuacionInteres")) + CCur(adoConsulta("ValorFluctuacionVac")) + CCur(adoConsulta("ValorGPCapital")) + CCur(adoConsulta("ValorFluctuacionMercado"))
            curValorNominal = CCur(adoConsulta("ValorNominal"))
            curValorTitulo = CCur(adoConsulta("ValorTitulo"))
            dblTasaMercado = CDbl(adoConsulta("ValorTasaMercado"))
            
            dblDuration = Duration(strCodTitulo, gdatFechaActual, gdatFechaActual, curMontoSubTotal, curMontoInteres, curValorNominal, curValorTitulo, dblTasaMercado, "", "", "")
            
            '*** Actualizar ***
            adoComm.CommandText = "UPDATE InversionValorizacion SET ValorDuration=" & dblDuration & " " & _
                "WHERE (FechaValorizacion>='" & strFechaCierre & "' AND FechaValorizacion<'" & strFechaSiguiente & "') AND " & _
                "CodTitulo='" & strCodTitulo & "' AND CodFondo='" & strCodFondo & "' AND " & _
                "CodAdministradora='" & gstrCodAdministradora & "'"
            adoConn.Execute adoComm.CommandText
            
            adoConsulta.MoveNext
        Loop
        adoConsulta.Close: Set adoConsulta = Nothing
               
        '*** Verificar corte de entrega de acciones ***
        Call CorteEventoCorporativo
         
        '*** Veriticar vencimiento de depósitos a plazo ***
        Call VencimientoDepositos
         
        adoComm.CommandText = "{ call up_IVProcInversionVencimiento('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaCierre & "') }"
        adoConn.Execute adoComm.CommandText
                
        adoComm.CommandText = "{ call up_IVProcEventoCorporativoVencimiento('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaCierre & "') }"
           
        '*** Factura Cuotas de Flujos por Vencimiento ***
        adoComm.CommandText = "{ call up_IVProcFacturacionFlujos('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaSiguiente & "','" & Valor_Indicador & "','" & Valor_Caracter & "') }"
        adoConn.Execute adoComm.CommandText
        
        '*** Factura Cuotas de Préstamos por Vencimiento ***
        adoComm.CommandText = "{ call up_FIProcFacturacionPrestamos('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaSiguiente & "','" & Valor_Indicador & "','" & Valor_Caracter & "') }"
        adoConn.Execute adoComm.CommandText

        '*** Verificar Pago de Cuotas de Suscripción ***
        If CInt(adoFondo("CantPartesPagoSuscripcion")) > 1 Then
            Call GenOrdenPagoCuotaSuscripcion(adoFondo("TipoValuacion"))
        End If
                        
        '*** Cierra datos del fondo ***
        adoFondo.Close: Set adoFondo = Nothing
                                                    
        '*** Es reproceso ***
        If Me.Tag = "R" Then
            
            datFechaCierre = DateAdd("d", 1, datFechaCierre)
            
            If datFechaCierre < dtpFechaCierreHasta Then
                '*** Reprocesar siguiente día ***
                Call ActualizarFechasCierre(datFechaCierre)
                GoTo Fondo_Reproceso
            Else
                '*** Actualizar estado del registro de Reproceso ***
                adoComm.CommandText = "UPDATE FondoReproceso SET Estado='" & Estado_Inactivo & "' " & _
                    "WHERE CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND NumReproceso=" & intNumReproceso
                adoConn.Execute adoComm.CommandText
            End If
        End If
                                                    
                                                    
        TimeFinp = Time
        frmMainMdi.stbMdi.Panels(3).Text = "Duración : " & Format((TimeFinp - TimeInip), "hh:mm:ss")
        MsgBox "Proceso de Cierre culminado exitosamente.", vbInformation
        Sleep 0&: Me.Refresh
    End If
    
cmdCierre_fin:
   Me.MousePointer = vbDefault
    cboFondo_Click
    frmMainMdi.stbMdi.Panels(3).Text = "Acción..."
   cmdCierre.Enabled = True
   chkSimulacion.Enabled = True
   Exit Sub
   
cmdCierre_error:
   strMensaje = "Error   : " & Str$(err) & Chr$(10)
   strMensaje = strMensaje & "Detalle : " & Error$ & Chr$(10)
   strMensaje = strMensaje & "SQL     : " & adoComm.CommandText
   
   MsgBox strMensaje, vbCritical
   
   Resume cmdCierre_fin

End Sub

Sub CierreMensual()
          
            
    '*** Inicializar Variables de Trabajo ***
    Dim TimeInip                As Variant, TimeFinp                    As Variant
    Dim adoFondo                As ADODB.Recordset, adoConsulta         As ADODB.Recordset
    Dim adoAuxiliar             As ADODB.Recordset
    Dim strMensaje              As String, strIndPagoParcial            As String
    Dim lngNumCom               As Long, lngNumCaj                      As Long
    Dim lngNumEnt               As Long, lngNumKar                      As Long
    Dim lngNumOpe               As Long
    Dim dblSaldoTotal           As Double, dblTasa                      As Double
    Dim dblTasaAdministracion   As Double, dblTasaValorCartera          As Double
    Dim strCodComision          As String, strCodVariable               As String
    Dim strTipo                 As String
    
          
    dblValNuevaCuota = 0
    dblValNuevaCuotaReal = 0
        
    '*** Cierre Diario ***

        frmMainMdi.stbMdi.Panels(3).Text = "Acción..."
        
        '*** Pedir Confirmación de Datos ***
        If chkSimulacion.Value Then
            strMensaje = "Para el proceso de SIMULACION confirme lo siguiente : " & vbNewLine
        Else
            strMensaje = "Para el proceso de CIERRE confirme lo siguiente : " & vbNewLine
        End If
        
        strMensaje = strMensaje & " Fondo >> " & Trim(cboFondo.Text) & vbNewLine & _
            " Fecha Inicial >> " & CStr(dtpFechaCierre.Value) & vbNewLine & _
            " Fecha Final >> " & CStr(dtpFechaCierreHasta.Value) & vbNewLine & _
            "¿ Seguro de continuar ?"
        If MsgBox(strMensaje, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then  'No desea continuar
            GoTo cmdCierre_fin
        End If
                        
        '*** Inicio del proceso ***
        If chkSimulacion.Value Then
            '*** Inicio de la SIMULACION ***
            frmMainMdi.stbMdi.Panels(3).Text = "Iniciando Simulación del Valor Cuota..."
        Else
            frmMainMdi.stbMdi.Panels(3).Text = "Inicio del Proceso..."
        End If
        
        '*** Prepara Tablas para la Simulación ***
        If chkSimulacion.Value Then
            frmMainMdi.stbMdi.Panels(3).Text = "Preparando Tablas para la Simulación..."
            
            adoComm.CommandText = "{ call up_GNProcPrepararTablasSimulacion('" & _
                strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strFechaCierreDesde & "','" & strFechaCierreHasta & "') }"
            adoConn.Execute adoComm.CommandText
            Sleep 0&
        
        End If
               
Fondo_Proceso:
                
        cmdCierre.Enabled = False
        TimeInip = Time
        Me.Refresh
        Me.MousePointer = vbHourglass

        frmMainMdi.stbMdi.Panels(3).Text = "Actualizando Saldos Finales..."

        adoComm.CommandText = "{ call up_GNActPartidaContableSaldos('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & _
            strFechaCierre & "','"

        If chkSimulacion.Value Then
           adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Simulacion & "') }"
        Else
           adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Definitivo & "') }"
        End If
        adoConn.Execute adoComm.CommandText
        Sleep 0&

        '*** Verifica y provisiona gastos del fondo - no incluye comision por adm de cartera ***
        If chkSimulacion.Value Then
           Call ProvisionGastosFondoMensual(Codigo_Cierre_Simulacion, Valor_Caracter)
        Else
           Call ProvisionGastosFondoMensual(Codigo_Cierre_Definitivo, Valor_Caracter)
        End If
                
        
        If strFechaCierre = strFechaCierreHasta Then
        
            '*** PreCierre - Traslado de SALDOS CONTABLES DE PRECIERRE >> PartidaContablePreSaldos ***
            frmMainMdi.stbMdi.Panels(3).Text = "Traslado de Saldos Al PreCierre..."
            
            adoComm.CommandText = "{ call up_GNProcTrasladoSaldosAlPreCierre('" & _
                strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strFechaCierreHasta & "','"
                
            If chkSimulacion.Value Then
                adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Simulacion & "') }"
            Else
                adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Definitivo & "') }"
            End If
            adoConn.Execute adoComm.CommandText
            Sleep 0&
                    
            '*** Cálculo de Ganancias y Pérdidas al PreCierre ***
            frmMainMdi.stbMdi.Panels(3).Text = "Calculando Pérdidas y Ganancias al Precierre..."
            
            adoComm.CommandText = "{ call up_GNProcCalcGPAlPreCierre('" & _
                strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strFechaCierreDesde & "','" & strFechaCierreHasta & "'," & Replace(dblTCCierre, ",", ".") & ""
                ','" & strCodMoneda & "','"
                
            If chkSimulacion.Value Then
                adoComm.CommandText = adoComm.CommandText & ", '" & Codigo_Cierre_Simulacion & "') }"
            Else
                adoComm.CommandText = adoComm.CommandText & ", '" & Codigo_Cierre_Definitivo & "') }"
            End If
            adoConn.Execute adoComm.CommandText
            Sleep 0&
            
            '*** Verifica y Provisiona Gastos del Fondo - Comision de Administración ***
            frmMainMdi.stbMdi.Panels(3).Text = "Verifica y Provisiona Gastos del Fondo - Comision de Administración..."
            
            If chkSimulacion.Value Then
               Call ProvisionGastosFondoMensual(Codigo_Cierre_Simulacion, Valor_Indicador)
            Else
               Call ProvisionGastosFondoMensual(Codigo_Cierre_Definitivo, Valor_Indicador)
            End If
            
            '*** Actualización Inicial de Saldos Finales y Monto de Ajuste Contable ***
            frmMainMdi.stbMdi.Panels(3).Text = "Actualizando Saldos Finales y Montos de Ajuste..."
            
            adoComm.CommandText = "{ call up_GNActMontoAjusteContable('" & _
                strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strFechaCierre & "','" & strFechaSiguiente & "','" & _
                gstrCodClaseTipoCambioFondo & "','" & _
                gstrValorTipoCambioCierre & "','"
                    
            If chkSimulacion.Value Then
               adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Simulacion & "') }"
            Else
               adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Definitivo & "') }"
            End If
            adoConn.Execute adoComm.CommandText
            Sleep 0&
                            
            '*** Asientos de Pérdida/Ganancia por Tipo de Cambio ***
            frmMainMdi.stbMdi.Panels(3).Text = "Registrando Asientos Contables por Ajuste en el Tipo de Cambio..."
            
            adoComm.CommandText = "{ call up_GNProcAjusteTipoCambio('" & _
                strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strFechaCierre & "','" & strFechaSiguiente & "'," & _
                Replace(dblTCCierre, ",", ".") & ",'" & gstrLogin & "','"
                    
            If chkSimulacion.Value Then
               adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Simulacion & "') }"
            Else
               adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Definitivo & "') }"
            End If
            adoConn.Execute adoComm.CommandText
            Sleep 0&

            '*** Cierre Pérdidas y Ganancias ***
            frmMainMdi.stbMdi.Panels(3).Text = "Calculando Resultados del Ejercicio al Cierre..."
            
            adoComm.CommandText = "{ call up_GNProcCalcGPAlCierre('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strFechaCierreDesde & "','" & strFechaCierreHasta & "','"
            
            If chkSimulacion.Value Then
               adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Simulacion & "') }"
            Else
               adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Definitivo & "') }"
            End If
            adoConn.Execute adoComm.CommandText
            Sleep 0&
            
            '*** Calculo de Valor de Cuota ***
            frmMainMdi.stbMdi.Panels(3).Text = "Calculando el Valor de Cuota"
            
            adoComm.CommandText = "{ call up_GNProcCalcValorCuota('" & _
                                strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                strFechaCierre & "','" & strFechaCierre & "','" & _
                                "','','" & _
                                strCodMoneda & "'," & _
                                dblTCCierre & ",'" & Trim(gstrCodClaseTipoCambioFondo) & "','" & _
                                Trim(gstrValorTipoCambioCierre) & "','"
                                
            If chkSimulacion.Value Then
               adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Simulacion & "') }"
            Else
               adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Definitivo & "') }"
            End If
            
            adoConn.Execute adoComm.CommandText
            Sleep 0&
                          
            If chkSimulacion.Value Then
                Call BuscarFondoSeries(Codigo_Cierre_Simulacion)
            Else
                Call BuscarFondoSeries(Codigo_Cierre_Definitivo)
            End If
                   
            '*** Actualizar patrimonio y activo en tabla de kardex de cuotas ***
            frmMainMdi.stbMdi.Panels(3).Text = "Actualizando Patrimonio y Activo..."
                    
            adoComm.CommandText = "{ call up_GNActPatrimonioActivo('" & _
                strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaCierre & "','" & strFechaCierre & "'," & _
                   CDbl(tdgTipoCambioCierre.Columns(2).Value) & ",'" & strCodMoneda & "','"
            
            If chkSimulacion.Value Then
               adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Simulacion & "') }"
            Else
               adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Definitivo & "') }"
            End If
            adoConn.Execute adoComm.CommandText
            Sleep 0&
        
        End If
        
       
        ' Esto se realiza siempre y cuando no estemos realizando el cierre del ultimo dia del periodo contable (año)
        frmMainMdi.stbMdi.Panels(3).Text = "Pasando Saldos Contables..."
        
        Set adoConsulta = New ADODB.Recordset
        
        adoComm.CommandText = "SELECT FechaFinal FROM PeriodoContable WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND PeriodoContable='" & gstrPeriodoActual & "' AND MesContable='99'"
        Set adoConsulta = adoComm.Execute
        If Not adoConsulta.EOF Then

            If Convertyyyymmdd(adoConsulta("FechaFinal")) <> strFechaCierre Then
                
                '*** Pase de saldos al dia siguiente
                adoComm.CommandText = "{ call up_GNProcTrasladoSaldosInicialesDiaSiguiente('"
                adoComm.CommandText = adoComm.CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strFechaCierre & "','"
                
                If chkSimulacion.Value Then
                   adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Simulacion & "') }"
                Else
                   adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Definitivo & "') }"
                End If
                adoConn.Execute adoComm.CommandText
                Sleep 0&

                
                '*** Actualiza kardex de cuotas al dia siguiente
                adoComm.CommandText = "{ call up_GNActKardexInicialCuotasFondo('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strFechaSiguiente & "'," & _
                    dblTCCierre & "," & dblTasa & ",'','','X','X','"
            
                If chkSimulacion.Value Then
                   adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Simulacion & "') }"
                Else
                   adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Definitivo & "') }"
                End If
                adoConn.Execute adoComm.CommandText
                Sleep 0&
            
                '*** Verifica apertura de fecha
                If strFechaCierre = strFechaCierreHasta Then
                    
                    '*** Deshabilitar periodo actual ***
                    adoComm.CommandText = "{ call up_GNActIndPeriodoHabil('" & _
                        strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaCierre & "'," & _
                        "'','X','"
                    
                    If chkSimulacion.Value Then
                       adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Simulacion & "') }"
                    Else
                       adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Definitivo & "') }"
                    End If
                    adoConn.Execute adoComm.CommandText
                    Sleep 0&
                
                
                    '*** Habilitar periodo siguiente ***
                    adoComm.CommandText = "{ call up_GNActIndPeriodoHabil('" & _
                        strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaSiguiente & "'," & _
                        "'X','','"
                    
                    If chkSimulacion.Value Then
                       adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Simulacion & "') }"
                    Else
                       adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Definitivo & "') }"
                    End If
                    adoConn.Execute adoComm.CommandText
                    Sleep 0&
                    
                    
                    '*** Procesa el cambio de fecha ***
                    adoComm.CommandText = "{ call up_CNProcControlFecha('" & _
                        strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        strFechaCierre & "','" & _
                        strFechaSiguiente & "') }"
                    adoConn.Execute adoComm.CommandText
                    
                    
                End If
            
            Else
            
            
                '*** Verifica apertura de fecha 31/12/
                If strFechaCierre = strFechaCierreHasta Then
                    
                    '*** Deshabilitar periodo actual al 31/12***
                    adoComm.CommandText = "{ call up_GNActIndPeriodoHabil('" & _
                        strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaCierre & "'," & _
                        "'','X','"
                    
                    If chkSimulacion.Value Then
                       adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Simulacion & "') }"
                    Else
                       adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Definitivo & "') }"
                    End If
                    adoConn.Execute adoComm.CommandText
                    Sleep 0&
                
                
                    '*** Habilitar periodo siguiente para tipo 99***
                    adoComm.CommandText = "{ call up_GNActIndPeriodoHabil('" & _
                        strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaCierre & "'," & _
                        "'X','','"
                    
                    If chkSimulacion.Value Then
                       adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Simulacion & "','99') }"
                    Else
                       adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Definitivo & "','99') }"
                    End If
                    adoConn.Execute adoComm.CommandText
                    Sleep 0&
                    
                    
                End If
            
            
            
            
            End If
        End If

        adoConsulta.Close: Set adoConsulta = Nothing
       
        '*** Verificar corte de entrega de acciones ***
        Call CorteEventoCorporativo
                        
        '*** Cierra datos del fondo ***
        'adoFondo.Close: Set adoFondo = Nothing
                                                    
        datFechaCierre = DateAdd("d", 1, datFechaCierre)
            
        If datFechaCierre <= dtpFechaCierreHasta.Value Then
            '*** Reprocesar siguiente día ***
            Call ActualizarFechasCierre(datFechaCierre)
            GoTo Fondo_Proceso
        Else
             frmMainMdi.txtFechaSistema = CStr(datFechaCierre)
        End If
                                                    
        TimeFinp = Time
        frmMainMdi.stbMdi.Panels(3).Text = "Duración : " & Format((TimeFinp - TimeInip), "hh:mm:ss")
        MsgBox "Proceso de Cierre culminado exitosamente.", vbInformation
        Sleep 0&: Me.Refresh
 
    
cmdCierre_fin:
   Me.MousePointer = vbDefault
   frmMainMdi.stbMdi.Panels(3).Text = "Acción..."
   cmdCierre.Enabled = True
   chkSimulacion.Enabled = True
   Exit Sub
   
cmdCierre_error:
   strMensaje = "Error   : " & Str$(err) & Chr$(10)
   strMensaje = strMensaje & "Detalle : " & Error$ & Chr$(10)
   strMensaje = strMensaje & "SQL     : " & adoComm.CommandText
   
   MsgBox strMensaje, vbCritical
   
   Resume cmdCierre_fin

End Sub


Private Function TodoOKMensual() As Boolean
                
    Dim adoConsulta As ADODB.Recordset
    Dim strMensaje  As String
    Dim adoFondo As ADODB.Recordset
    
    TodoOKMensual = False
                
    If cboFondo.ListCount = 0 Then
        MsgBox "No existen fondos definidos...", vbCritical, Me.Caption
        Exit Function
    End If
    
    If strMesContable = "00" Or strMesContable = "99" Then
        MsgBox "El Cierre Mensual no está permitido para el periodo de Apertura/Cierre Contable", vbCritical, Me.Caption
        Exit Function
    End If

    '*** Cerrar periodo ? ***
    adoComm.CommandText = "SELECT IndCierre FROM PeriodoContable " & _
        "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
        "MesContable='99' AND FechaFinal='" & strFechaAnterior & "'"
    Set adoConsulta = adoComm.Execute
    
    If Not adoConsulta.EOF Then
        If adoConsulta("IndCierre") = Valor_Caracter Then
            MsgBox "Por favor procese el Cierre Anual.", vbCritical, Me.Caption
            adoConsulta.Close: Set adoConsulta = Nothing
            Exit Function
        End If
    End If
    adoConsulta.Close

    '*** Existe un descuadre en la contabilidad ? ***
    adoComm.CommandText = "{ call up_ACValidaCuadreContabilidad('" & _
        strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaCierre & "','" & _
        strFechaSiguiente & "') }"
    Set adoConsulta = adoComm.Execute
    
    If Not adoConsulta.EOF Then
        If Not IsNull(adoConsulta("MontoContable")) Then
            If CCur(adoConsulta("MontoContable")) <> 0 Then
                If MsgBox("Existe un descuadre contable de " & CStr(adoConsulta("MontoContable")) & " para el día " & CStr(dtpFechaCierre.Value) & "." & vbNewLine & vbNewLine & "Desea Continuar ?.", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                    adoConsulta.Close: Set adoConsulta = Nothing
                    Exit Function
                End If
            End If
        End If
    End If
    adoConsulta.Close

    '*** Cierre en fecha aún no abierta para el Fondo ***
    adoComm.CommandText = "{ call up_GNValidaFechaNoAbierta('" & _
        strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaCierreHasta & "','" & _
        strFechaCierreHastaSiguiente & "') }"
    Set adoConsulta = adoComm.Execute
    
    If Not adoConsulta.EOF Then
        If Trim(adoConsulta("IndAbierto")) = Valor_Caracter Then
            MsgBox "El Día " & CStr(dtpFechaCierreHasta.Value) & " aún no ha sido abierto.", vbCritical, Me.Caption
            adoConsulta.Close: Set adoConsulta = Nothing
            Exit Function
        End If
    End If
    adoConsulta.Close
        
    '*** Verificar si existen valores sin precio o tir de mercado en cartera ***
    adoComm.CommandText = "SELECT II.Nemotecnico,isnull(PrecioCierre,0) PrecioCierre,isnull(TirCierre,0) TirCierre " & _
        "FROM InversionKardex IK JOIN InstrumentoInversion II ON(II.CodTitulo=IK.CodTitulo) " & _
        "LEFT JOIN InstrumentoPrecioTir IPT ON(IPT.CodTitulo=IK.CodTitulo AND IPT.IndUltimoPrecio='X') " & _
        "JOIN InversionFile IVF ON(IVF.CodFile=IK.CodFile AND (IndTir='X' OR IndPrecio='X')) " & _
        "WHERE IK.SaldoFinal>0 AND IK.IndUltimoMovimiento='X' AND IK.CodFondo= '" & gstrCodFondoContable & "' " & _
        "ORDER BY II.Nemotecnico"
    Set adoConsulta = adoComm.Execute
    
    strMensaje = Valor_Caracter
    Do While Not adoConsulta.EOF
        If adoConsulta("PrecioCierre") = 0 And adoConsulta("TirCierre") = 0 Then
            strMensaje = strMensaje & adoConsulta("Nemotecnico") & vbNewLine
        End If
        
        adoConsulta.MoveNext
    Loop
    adoConsulta.Close: Set adoConsulta = Nothing
    
    If strMensaje <> Valor_Caracter Then
        strMensaje = "Existen los siguientes valores sin Precio o Tir de Mercado :" & vbNewLine & vbNewLine & strMensaje
        If MsgBox(strMensaje & vbNewLine & vbNewLine & "Desea Continuar ?.", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            Exit Function
        End If
    End If
    
    Set adoConsulta = New ADODB.Recordset
            
    '*** Tipo de Cambio ***
    adoComm.CommandText = "SELECT CodMoneda FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND " & _
        "CodFondo='" & strCodFondo & "'"
    Set adoConsulta = adoComm.Execute

    strIndCobrar = Valor_Caracter
    If Not adoConsulta.EOF Then
        Me.MousePointer = vbDefault
        'º
        '*** Obtener tipo de cambio ***
        dblTCCierre = ObtenerValorTCCierre(adoConsulta("CodMoneda"))
        
        If dblTCCierre = 0 Then
            MsgBox "El Tipo de Cambio para la fecha de cierre NO ESTA REGISTRADO.", vbCritical, Me.Caption
            
'            adoFondo.Close:
            Set adoFondo = Nothing
            Exit Function
        End If
    End If
    adoConsulta.Close
    
    TodoOKMensual = True
  
End Function

Sub Ocultar_Recursos()


    If strTipoFondoFrecuencia = strFrecuenciaValorizacionMensual And strTipoFondo = Administradora_Fondos Then
    
        lblDescrip(2).Visible = False
        lblDescrip(3).Visible = False
        lblDescrip(4).Visible = False
        lblDescrip(5).Visible = False
        lblDescrip(6).Visible = False
        lblDescrip(7).Visible = False
        dtpFechaEntrega.Visible = False
        
        lblValorAIR(0).Visible = False
        lblValorAIR(1).Visible = False
        
        lblValorDIR(0).Visible = False
        lblValorDIR(1).Visible = False
        
        lblRentabilidad(0).Visible = False
        lblRentabilidad(1).Visible = False
        
        frmCierreDiario.Caption = "Cierre Mensual"
        
    Else
         frmCierreDiario.Caption = "Cierre Diario"
    End If


End Sub


Private Sub ProvisionGastosFondoMensual(strTipoCierre As String, strIndNoIncluyeEnPreCierre As String)

    Dim adoRegistro             As ADODB.Recordset
    Dim adoConsulta             As ADODB.Recordset
    Dim strCodFile              As String, strCodDetalleFile            As String
    Dim strNumAsiento           As String, strDescripAsiento            As String
    Dim strIndDebeHaber         As String, strDescripMovimiento         As String
    Dim strDescripGasto         As String, strFechaGrabar               As String
    Dim intDiasProvision        As Integer, intCantRegistros            As Integer
    Dim intContador             As Integer, intDiasCorridos             As Integer
    Dim curMontoRenta           As Currency, curSaldoProvision          As Currency
    Dim curMontoMovimientoMN    As Currency, curMontoMovimientoME       As Currency
    Dim curMontoContable        As Currency, curValorAnterior           As Currency
    Dim curValorActual          As Currency
    Dim dblValorTipoCambio      As Double
    Dim dblValorAjusteProv      As Double
    Dim curValorTotal           As Currency
    Dim intNumDiasPeriodo       As Integer
    Dim intDiasProvision1       As Integer
    Dim strTipoAuxiliar         As String
    Dim strCodAuxiliar          As String

    frmMainMdi.stbMdi.Panels(3).Text = "Provisionando Gastos del Fondo..."
    
    Set adoRegistro = New ADODB.Recordset
    Set adoConsulta = New ADODB.Recordset
    
    With adoComm
    
       
        If strTipoCierre = Codigo_Cierre_Definitivo Then
            .CommandText = "SELECT * FROM FondoGasto FG " & _
                 "JOIN FondoGastoPeriodo FGP ON (FG.CodFondo = FGP.CodFondo AND FG.CodAdministradora = FGP.CodAdministradora AND FG.NumGasto = FGP.NumGasto) " & _
                 "WHERE FGP.FechaInicio <= '" & strFechaCierre & "' AND FGP.FechaVencimiento >= '" & strFechaCierre & "' AND FG.CodFondo='" & strCodFondo & "' AND " & _
                "FG.CodAdministradora='" & gstrCodAdministradora & "' AND FG.IndVigente='X' AND FG.IndNoIncluyeEnBalancePreCierre = '" & strIndNoIncluyeEnPreCierre & "'" & _
                " AND FG.CodAplicacionDevengo = '" & Codigo_Aplica_Devengo_Periodica & "'"
        Else
            .CommandText = "SELECT * FROM FondoGastoTmp FG " & _
                 "JOIN FondoGastoPeriodoTmp FGP ON (FG.CodFondo = FGP.CodFondo AND FG.CodAdministradora = FGP.CodAdministradora AND FG.NumGasto = FGP.NumGasto) " & _
                 "WHERE FGP.FechaInicio <= '" & strFechaCierre & "' AND FGP.FechaVencimiento >= '" & strFechaCierre & "' AND FG.CodFondo='" & strCodFondo & "' AND " & _
                "FG.CodAdministradora='" & gstrCodAdministradora & "' AND FG.IndVigente='X' AND FG.IndNoIncluyeEnBalancePreCierre = '" & strIndNoIncluyeEnPreCierre & "'" & _
                " AND FG.CodAplicacionDevengo = '" & Codigo_Aplica_Devengo_Periodica & "'"
        End If
                        
        Set adoRegistro = .Execute
        
        Do While Not adoRegistro.EOF
            strCodFile = Trim(adoRegistro("CodFile"))
            
            strTipoAuxiliar = "02"
            strCodAuxiliar = adoRegistro("TipoProveedor") & adoRegistro("CodProveedor")
            
            .CommandText = "SELECT CodDetalleFile FROM InversionDetalleFile " & _
                "WHERE CodFile='" & strCodFile & "' AND DescripDetalleFile='" & adoRegistro("CodCuenta") & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                strCodDetalleFile = adoConsulta("CodDetalleFile")
            End If
            adoConsulta.Close
            Set adoConsulta = New ADODB.Recordset
            
            If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
                'Por defecto obtener el tipo de cambio SUNAT de la fecha del documento indicada en el registro de compras
                .CommandText = "SELECT FechaComprobante, CodMonedaPago " & _
                    " FROM RegistroCompra RC " & _
                    " WHERE RC.NumGasto = " & CInt(adoRegistro("NumGasto")) & " AND RC.CodFondo = '" & strCodFondo & "' AND " & _
                    " RC.CodAdministradora = '" & gstrCodAdministradora & "' AND RC.FechaPago = '" & strFechaCierre & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    'Es tipo de cambio SUNAT de la fecha de emision del documento si el documento es factura
                    dblValorTipoCambio = ObtenerTipoCambioMoneda(Codigo_TipoCambio_Sunat, Codigo_Valor_TipoCambioVenta, adoConsulta("FechaComprobante"), Codigo_Moneda_Local, adoRegistro("CodMoneda"))
                Else
                    dblValorTipoCambio = ObtenerTipoCambioMoneda(Codigo_TipoCambio_Sunat, Codigo_Valor_TipoCambioVenta, dtpFechaCierre.Value, Codigo_Moneda_Local, Trim(adoRegistro("CodMoneda")))
                End If
                adoConsulta.Close
                Set adoConsulta = New ADODB.Recordset
            Else
                dblValorTipoCambio = 1
            End If
                        
            '*** Verificar Dinamica Contable ***
            .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
                "WHERE TipoOperacion='" & Codigo_Dinamica_Gasto & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                If CInt(adoConsulta("NumRegistros")) > 0 Then
                    intCantRegistros = CInt(adoConsulta("NumRegistros"))
                Else
                    MsgBox "NO EXISTE Dinámica Contable para la provisión", vbCritical
                    adoConsulta.Close: Set adoConsulta = Nothing
                    Exit Sub
                End If
            End If
            adoConsulta.Close
                        
            '*** Obtener Descripción del Gasto ***
            .CommandText = "SELECT DescripCuenta FROM PlanContable WHERE CodCuenta='" & adoRegistro("CodCuenta") & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                strDescripGasto = Trim(adoConsulta("DescripCuenta"))
            End If
            adoConsulta.Close
            
            '*** Obtener las cuentas de inversión ***
            Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, adoRegistro("CodMoneda"))
            
            '*** Obtener Saldo de Inversión ***
            If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvGasto & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoProvision = CDbl(adoConsulta("Saldo"))
            Else
                curSaldoProvision = 0
            End If
            adoConsulta.Close
                                
            intDiasProvision = DateDiff("d", adoRegistro("FechaInicio"), adoRegistro("FechaVencimiento")) + 1
            intDiasCorridos = DateDiff("d", adoRegistro("FechaInicio"), gdatFechaActual) + 1
            
            curValorAnterior = curSaldoProvision
            
            If adoRegistro("CodAplicacionDevengo") = Codigo_Aplica_Devengo_Periodica Then
                Set adoConsulta = New ADODB.Recordset
        
                '*** Obtener el número de días del peridodo de pago ***
                .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' AND CodParametro='" & adoRegistro("CodFrecuenciaDevengo") & "'"
                Set adoConsulta = adoComm.Execute
        
                If Not adoConsulta.EOF Then
                    intNumDiasPeriodo = CInt(adoConsulta("ValorParametro")) '*** Días del periodo  ***
                Else
                    intNumDiasPeriodo = 0
                End If
                adoConsulta.Close: Set adoConsulta = Nothing
            Else
                intNumDiasPeriodo = 0
            End If
            
            If adoRegistro("CodTipoValor") = Codigo_Tipo_Costo_Porcentaje Then
                curValorTotal = CalculoInteres(adoRegistro("PorcenGasto"), adoRegistro("CodTipoTasa"), adoRegistro("CodPeriodoTasa"), adoRegistro("CodBaseAnual"), adoRegistro("MontoBaseCalculo"), adoRegistro("FechaInicio"), adoRegistro("FechaVencimiento"))
            Else
                curValorTotal = adoRegistro("MontoGasto")
            End If
            
            'UltimoDiaMes
            
            'Para el calculo prorratea sobre la base de Actual/x --osea sobre el numero real de dias del mes!
            If adoRegistro("CodAplicacionDevengo") = Codigo_Aplica_Devengo_Periodica Then
                If intNumDiasPeriodo <> 0 Then
                    If intDiasCorridos Mod intNumDiasPeriodo = 0 Then
                        curMontoRenta = Round(curValorTotal / (intDiasProvision / intNumDiasPeriodo), 2)
                    Else
                        curMontoRenta = 0
                    End If
                Else
                    curMontoRenta = 0
                End If
                curValorActual = adoRegistro("MontoDevengo") + curMontoRenta
            Else 'No Porratea, es inmediato
                curMontoRenta = curValorTotal
                curValorActual = curMontoRenta
            End If
           
            
            'Control de remanentes
            If adoRegistro("FechaVencimiento") = gdatFechaActual Then
                If (curValorTotal - curValorActual) <> 0 Then
                    dblValorAjusteProv = (curValorTotal - curValorActual)
                    curMontoRenta = curMontoRenta + dblValorAjusteProv
                    curValorActual = curValorActual + dblValorAjusteProv
                End If
            End If
                                    
            '*** Provisión ***
            If curMontoRenta <> 0 Then
                strDescripAsiento = "Provisión" & Space(1) & strDescripGasto
                strDescripMovimiento = strDescripGasto
                If curMontoRenta > 0 Then strDescripMovimiento = strDescripGasto
                                                
                .CommandType = adCmdStoredProc
                '*** Obtener el número del parámetro **
                .CommandText = "up_ACObtenerUltNumero"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "up_ACObtenerUltNumeroTmp"  '*** Simulación ***
                
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
                
                '*** Cabecera ***
                .CommandText = "{ call up_ACAdicAsientoContable('"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableTmp('"  '*** Simulación ***
                
                .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
                    strFechaGrabar & "','" & _
                    gstrPeriodoActual & "','" & gstrMesActual & "','','" & _
                    strDescripAsiento & "','" & Trim(adoRegistro("CodMoneda")) & "','" & _
                    Codigo_Moneda_Local & "','',''," & _
                    CDec(curMontoRenta) & ",'" & Estado_Activo & "'," & _
                    intCantRegistros & ",'" & _
                    strFechaCierre & Space(1) & Format(Time, "hh:ss") & "','" & _
                    strCodModulo & "',''," & _
                    CDec(dblValorTipoCambio) & ",'','','" & _
                    strDescripAsiento & "','','X','') }"
                adoConn.Execute .CommandText
                
                '*** Detalle ***
                .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
                    "WHERE TipoOperacion='" & Codigo_Dinamica_Gasto & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' " & _
                    "ORDER BY NumSecuencial"
                Set adoConsulta = .Execute
        
                Do While Not adoConsulta.EOF
                
                    Select Case Trim(adoConsulta("TipoCuentaInversion"))
                        Case Codigo_CtaInversion
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvInteres
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteres
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresVencido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaVacCorrido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaXPagar
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaXCobrar
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresCorrido
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvReajusteK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaReajusteK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvFlucMercado
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaFlucMercado
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvInteresVac
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInteresVac
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaIntCorridoK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvFlucK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaFlucK
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaInversionTransito
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaProvGasto
                            curMontoMovimientoMN = curMontoRenta
                            
                        Case Codigo_CtaIngresoOperacional
                            curMontoMovimientoMN = curMontoRenta
                        
                        Case Codigo_CtaCosto
                            curMontoMovimientoMN = curMontoRenta
                        
                        Case Codigo_CtaCostoSAB
                            curMontoMovimientoMN = curMontoRenta
                    
                        Case Codigo_CtaCostoBVL
                            curMontoMovimientoMN = curMontoRenta
                        
                        Case Codigo_CtaCostoCavali
                            curMontoMovimientoMN = curMontoRenta

                        Case Codigo_CtaCostoFondoLiquidacion
                            curMontoMovimientoMN = curMontoRenta

                        Case Codigo_CtaCostoGastosBancarios
                            curMontoMovimientoMN = curMontoRenta

                        Case Codigo_CtaCostoComisionEspecial
                            curMontoMovimientoMN = curMontoRenta
                    
                        Case Codigo_CtaCostoFondoGarantia
                            curMontoMovimientoMN = curMontoRenta

                        Case Codigo_CtaCostoConasev
                            curMontoMovimientoMN = curMontoRenta
                        
                        Case Codigo_CtaComision
                            curMontoMovimientoMN = curMontoRenta
                    
                    End Select
                    
                    strIndDebeHaber = Trim(adoConsulta("IndDebeHaber"))
                    If strIndDebeHaber = "H" Then
                        curMontoMovimientoMN = curMontoMovimientoMN * -1
                        If curMontoMovimientoMN > 0 Then strIndDebeHaber = "D"
                    ElseIf strIndDebeHaber = "D" Then
                        If curMontoMovimientoMN < 0 Then strIndDebeHaber = "H"
                    End If
                    
                    If strIndDebeHaber = "T" Then
                        If curMontoMovimientoMN > 0 Then
                            strIndDebeHaber = "D"
                        Else
                            strIndDebeHaber = "H"
                        End If
                    End If
                    strDescripMovimiento = Trim(adoConsulta("DescripDinamica"))
                    curMontoMovimientoME = 0
                    curMontoContable = curMontoMovimientoMN
        
                    If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
                        curMontoContable = Round(curMontoMovimientoMN * CDec(dblValorTipoCambio), 2)
                        curMontoMovimientoME = curMontoMovimientoMN
                        curMontoMovimientoMN = 0
                    End If
                                
                    '*** Movimiento ***
                    .CommandText = "{ call up_ACAdicAsientoContableDetalle('"
                    If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNAdicAsientoContableDetalleTmp('"

                    .CommandText = .CommandText & strNumAsiento & "','" & strCodFondo & "','" & _
                        gstrCodAdministradora & "'," & _
                        CInt(adoConsulta("NumSecuencial")) & ",'" & _
                        strFechaGrabar & "','" & _
                        gstrPeriodoActual & "','" & _
                        gstrMesActual & "','" & _
                        strDescripMovimiento & "','" & _
                        strIndDebeHaber & "','" & _
                        Trim(adoConsulta("CodCuenta")) & "','" & _
                        Trim(adoRegistro("CodMoneda")) & "'," & _
                        CDec(curMontoMovimientoMN) & "," & _
                        CDec(curMontoMovimientoME) & "," & _
                        CDec(curMontoContable) & ",'" & _
                        Trim(adoRegistro("CodFile")) & "','" & _
                        Trim(adoRegistro("CodAnalitica")) & "','" & _
                        strTipoAuxiliar & "','" & _
                        strCodAuxiliar & "') }"
                    adoConn.Execute .CommandText
                
                    '*** Saldos ***
                    .CommandText = "{ call up_ACGenPartidaContableSaldos('"
                    If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNGenPartidaContableSaldosTmp('"

                    .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        gstrPeriodoActual & "','" & gstrMesActual & "','" & _
                        Trim(adoConsulta("CodCuenta")) & "','" & _
                        Trim(adoRegistro("CodFile")) & "','" & _
                        Trim(adoRegistro("CodAnalitica")) & "','" & _
                        strFechaCierre & "','" & _
                        strFechaSiguiente & "'," & _
                        CDec(curMontoMovimientoMN) & "," & _
                        CDec(curMontoMovimientoME) & "," & _
                        CDec(curMontoContable) & ",'" & _
                        strIndDebeHaber & "','" & _
                        Trim(adoRegistro("CodMoneda")) & "') }"
                    adoConn.Execute .CommandText
                                    
                    '*** Validar valor de cuenta contable ***
                    If Trim(adoConsulta("CodCuenta")) = Valor_Caracter Then
                        MsgBox "Registro Nro. " & CStr(intContador) & " de Asiento Contable no tiene cuenta asignada", vbCritical, "Valorización Depósitos"
                        gblnRollBack = True
                        Exit Sub
                    End If
                    
                    adoConsulta.MoveNext
                Loop
                adoConsulta.Close: Set adoConsulta = Nothing
                                
                '*** Actualizar el número del parámetro **
                .CommandText = "{ call up_ACActUltNumero('"
                If strTipoCierre = Codigo_Cierre_Simulacion Then .CommandText = "{ call up_GNActUltNumeroTmp('"
                
                .CommandText = .CommandText & strCodFondo & "','" & _
                    gstrCodAdministradora & "','" & _
                    Valor_NumComprobante & "','" & _
                    strNumAsiento & "') }"
                adoConn.Execute .CommandText
                
                If strTipoCierre = Codigo_Cierre_Definitivo Then
                    .CommandText = "UPDATE FondoGasto SET "
                Else
                    .CommandText = "UPDATE FondoGastoTmp SET "
                End If
                
                If Convertyyyymmdd(adoRegistro("FechaVencimiento")) = strFechaCierre Then
                    .CommandText = .CommandText & "IndVigente='',"
                End If
                
                .CommandText = .CommandText & "MontoDevengo=" & curValorActual & _
                " WHERE NumGasto=" & adoRegistro("NumGasto") & " AND " & _
                       "CodCuenta='" & Trim(adoRegistro("CodCuenta")) & "' AND CodFondo='" & strCodFondo & "' AND " & _
                       "CodAdministradora='" & gstrCodAdministradora & "'"
                
                adoConn.Execute .CommandText
                
            End If
                                            
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    Exit Sub
  
Ctrl_Error:
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
  
End Sub

