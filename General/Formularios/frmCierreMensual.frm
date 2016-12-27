VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmCierreMensual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre Contable"
   ClientHeight    =   4860
   ClientLeft      =   1440
   ClientTop       =   1695
   ClientWidth     =   6615
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
   Icon            =   "frmCierreMensual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4860
   ScaleWidth      =   6615
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   4920
      Picture         =   "frmCierreMensual.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCierre 
      Caption         =   "&Procesar"
      Height          =   735
      Left            =   3480
      Picture         =   "frmCierreMensual.frx":09C4
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3960
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3795
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6694
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Parametros de Cierre"
      TabPicture(0)   =   "frmCierreMensual.frx":0F2C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDescrip(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraFechaCierre"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkSimulacion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboFondo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Tipos de Cambio"
      TabPicture(1)   =   "frmCierreMensual.frx":0F48
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tdgTipoCambioCierre"
      Tab(1).Control(1)=   "lblDescrip(8)"
      Tab(1).ControlCount=   2
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
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   810
         Width           =   5115
      End
      Begin VB.CheckBox chkSimulacion 
         Caption         =   "Simular el Cierre Contable"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         ToolTipText     =   "Marcar para proceso de simulación"
         Top             =   3270
         Width           =   2895
      End
      Begin VB.Frame fraFechaCierre 
         Caption         =   "Periodo de Cierre"
         Height          =   1665
         Left            =   360
         TabIndex        =   1
         Top             =   1410
         Width           =   5625
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
            Left            =   1320
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   420
            Width           =   3975
         End
         Begin MSComCtl2.DTPicker dtpFechaCierre 
            Height          =   345
            Left            =   1290
            TabIndex        =   2
            Top             =   1050
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
            Format          =   176160769
            CurrentDate     =   38068
         End
         Begin MSComCtl2.DTPicker dtpFechaCierreHasta 
            Height          =   345
            Left            =   3930
            TabIndex        =   3
            Top             =   1050
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
            Format          =   176160769
            CurrentDate     =   38068
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Periodo"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   10
            Left            =   360
            TabIndex        =   9
            Top             =   480
            Width           =   885
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Desde"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   5
            Top             =   1110
            Width           =   765
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Hasta"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   3030
            TabIndex        =   4
            Top             =   1110
            Width           =   765
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgTipoCambioCierre 
         Bindings        =   "frmCierreMensual.frx":0F64
         Height          =   1545
         Left            =   -73320
         OleObjectBlob   =   "frmCierreMensual.frx":0F86
         TabIndex        =   25
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label lblDescrip 
         Alignment       =   2  'Center
         Caption         =   "Tipo de Cambio Oficial"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   8
         Left            =   -73860
         TabIndex        =   10
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fondo"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   855
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc adoTipoCambioCierre 
      Height          =   375
      Left            =   -450
      Top             =   1110
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
   Begin MSComCtl2.DTPicker dtpFechaEntrega 
      Height          =   315
      Left            =   3330
      TabIndex        =   11
      Top             =   7320
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   176160769
      CurrentDate     =   38068
   End
   Begin VB.Label lblDescrip 
      Alignment       =   2  'Center
      Caption         =   "Rentabilidad"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   5
      Left            =   5265
      TabIndex        =   23
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDescrip 
      Alignment       =   2  'Center
      Caption         =   "Valor D.I.R."
      ForeColor       =   &H00800000&
      Height          =   165
      Index           =   4
      Left            =   3345
      TabIndex        =   22
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDescrip 
      Alignment       =   2  'Center
      Caption         =   "Valor A.I.R."
      ForeColor       =   &H00800000&
      Height          =   165
      Index           =   3
      Left            =   1425
      TabIndex        =   21
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDescrip 
      Caption         =   "Fecha de Entrega de Redenciones"
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   2
      Left            =   195
      TabIndex        =   20
      Top             =   7365
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Label lblRentabilidad 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.0000"
      Height          =   285
      Index           =   0
      Left            =   5220
      TabIndex        =   19
      Top             =   8040
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label lblValorDIR 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00000000"
      Height          =   285
      Index           =   0
      Left            =   3300
      TabIndex        =   18
      Top             =   8040
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblValorAIR 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100.00000000"
      Height          =   285
      Index           =   0
      Left            =   1380
      TabIndex        =   17
      Top             =   8040
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblDescrip 
      Caption         =   "dd/mm/yyyy"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   6
      Left            =   150
      TabIndex        =   16
      Top             =   8040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblDescrip 
      Caption         =   "dd/mm/yyyy"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   7
      Left            =   150
      TabIndex        =   15
      Top             =   8400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblValorAIR 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100.00000000"
      ForeColor       =   &H8000000E&
      Height          =   285
      Index           =   1
      Left            =   1380
      TabIndex        =   14
      Top             =   8400
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblValorDIR 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00000000"
      ForeColor       =   &H8000000E&
      Height          =   285
      Index           =   1
      Left            =   3300
      TabIndex        =   13
      Top             =   8400
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblRentabilidad 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.0000"
      ForeColor       =   &H8000000E&
      Height          =   285
      Index           =   1
      Left            =   5220
      TabIndex        =   12
      Top             =   8400
      Visible         =   0   'False
      Width           =   1635
   End
End
Attribute VB_Name = "frmCierreMensual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()              As String, arrPeriodo()                 As String

Dim strCodFondo             As String, strFechaCierre               As String
Dim strFechaAnterior        As String, strFechaSiguiente            As String
Dim strFechaAnteAnterior    As String, strFechaSubSiguiente         As String
Dim strCodMoneda            As String, strCodModulo                 As String
Dim strSQL                  As String, datFechaCierre               As Date
Dim strFechaCierreHasta     As String, strFechaCierreDesde          As String
Dim strFechaCierreHastaSiguiente As String, strMesContable          As String

Dim dblValNuevaCuota        As Double, dblValorCuotaNominal         As Double
Dim dblValNuevaCuotaReal    As Double

'*** Variables para los códigos de cuentas contables ***
Dim strCodCuentaValuacion   As String, strCodCuentaResultados       As String
Dim intNumReproceso         As Integer, strPeriodoContable          As String
Dim strIndCobrar            As String, dblTCCierre                  As Double
Dim adoConsultaTipoCambio   As ADODB.Recordset


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
            
           ' Call ObtenerCuentasInversion(strCodFile, adoRegistro("CodDetalleFile"))
            
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
            
            'Call ObtenerCuentasInversion(strCodFile, adoRegistro("CodDetalleFile"))
            
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
            
           ' Call ObtenerCuentasInversion(strCodFile, adoRegistro("CodDetalleFile"))
            
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
    
    '*** Rentabilidad de Valores de Renta Fija Corto Plazo ***
    frmMainMdi.stbMdi.Panels(3).Text = "Calculando rentabilidad de Valores de Renta Fija Corto Plazo..."
       
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT IK.MontoSaldo,IK.CodTitulo,II.CodMoneda,IK.CodAnalitica,IK.CodFile,SaldoFinal,CodDetalleFile,CodSubDetalleFile,CodTipoVac,CuponCalculo," & _
            "ValorNominal,CodTipoTasa,BaseAnual,CodTasa,TasaInteres,DiasPlazo,IndCuponCero,FechaEmision,TirPromedio,CodTipoAjuste,Nemotecnico " & _
            "FROM InversionKardex IK JOIN InstrumentoInversion II " & _
            "ON(II.CodTitulo=IK.CodTitulo) " & _
            "WHERE IK.CodFile IN ('006','010','012','014','015','016') AND FechaOperacion <='" & strFechaCierre & "' AND SaldoFinal > 0 AND " & _
            "IK.CodAdministradora='" & gstrCodAdministradora & "' AND IK.CodFondo='" & strCodFondo & "' AND IndUltimoMovimiento='X'"
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
            curSaldoValorizar = CCur(adoRegistro("MontoSaldo")) '("SaldoFinal")
            intDiasDeRenta = DateDiff("d", CVDate(adoRegistro("FechaEmision")), gdatFechaActual) + 1
            
            If strBaseAnual = Codigo_Base_30_360 Or strBaseAnual = Codigo_Base_30_365 Then intDiasDeRenta = Dias360(CVDate(adoRegistro("FechaEmision")), gdatFechaActual, True) + 1
            
            Set adoConsulta = New ADODB.Recordset
                        
            '*** Verificar Dinamica Contable ***
            .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
                "WHERE TipoOperacion='" & Codigo_Dinamica_Provision & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "'"
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
            .CommandText = "SELECT FactorDiario FROM InstrumentoInversionCalendario " & _
                "WHERE CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                dblFactorDiarioCupon = CDbl(adoConsulta("FactorDiario"))
            End If
            adoConsulta.Close
            
            '*** Obtener las cuentas de inversión ***
           ' Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile)
            
            '*** Obtener tipo de cambio ***
            dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
            
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
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
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
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
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
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
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
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvFlucMercado & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "'"
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
                

                        
            curMontoRenta = Round(curValorActual - curValorAnterior, 2)
            
            '*** Cálculo Provisión G/P Capital ***
            'If strOrigen = "L" Then
                '*** VAN AL DIA ANTERIOR AL CIERRE ***
                curValorAnterior = curSaldoInversion + curSaldoInteresCorrido + curSaldoFluctuacion + curSaldoGPCapital
    
                '*** CALCULO DEL VAN A LA FECHA DE CIERRE ***
                If CDbl(adoRegistro("TirPromedio")) <> 0 Then
                    If strModalidadInteres = Codigo_Interes_Descuento Then curSaldoValorizar = curSaldoInversion
                    
                    Dim datFechaGP  As Date, datFechaSiguienteGP    As Date
                    Dim dblValorTir As Double
   
                    datFechaGP = Convertddmmyyyy(strFechaCierre)
                    datFechaSiguienteGP = Convertddmmyyyy(strFechaSiguiente)
    
                    dblValorTir = CDbl(adoRegistro("TirPromedio"))
                    
                    curValorActual = VNANoPer(adoRegistro("CodTitulo"), datFechaSiguienteGP, datFechaSiguienteGP, curSaldoValorizar, curSaldoValorizar, dblValorTir, adoRegistro("CodTipoAjuste"), Valor_Caracter, Valor_Caracter)
    
                    '*** CALCULO DEL MONTO DE GANANCIA/PERDIDA DE curCapital ***
                    curMontoProvisionCapital = Round(curValorActual - curValorAnterior - curMontoRenta, 2)
                Else
                    curMontoProvisionCapital = 0
                End If
    
            'End If
            
            '*** Cálculo Fluctuación Mercado ***
            If dblTirCierre > 0 Then
                curValorAnterior = curSaldoInversion + curSaldoInteresCorrido + curSaldoFluctuacion + curSaldoGPCapital + curSaldoFluctuacionMercado
                
                'curValorActual= CDbl(VNANoPer(Trim$(adoresult!COD_FILE), Trim$(adoresult!COD_ANAL), vntFeccie, vntFeccieMas1Dia, CDbl(curCapital), SldAmort, dblTirHoy, strTipVac))
                curValorActual = VNANoPer(adoRegistro("CodTitulo"), datFechaSiguienteGP, datFechaSiguienteGP, curSaldoValorizar, curSaldoValorizar, dblTirCierre, adoRegistro("CodTipoAjuste"), Valor_Caracter, Valor_Caracter)
                
                curMontoFluctuacionMercado = Round(curValorActual - curValorAnterior - curMontoRenta - curMontoProvisionCapital, 2)
            Else
                curMontoFluctuacionMercado = 0
            End If
            
            '*** Contabilización ***
            If curMontoRenta <> 0 Or curMontoProvisionCapital <> 0 Or curMontoFluctuacionMercado <> 0 Then
                'strDescripAsiento = "Valorización" & Space(1) & "(" & Trim(adoRegistro("CodFile")) & "-" & Trim(adoRegistro("CodAnalitica")) & ")"
                strDescripAsiento = "Valorización" & Space(1) & strNemonico
                strDescripMovimiento = "Pérdida"
                If curMontoRenta > 0 Then strDescripMovimiento = "Ganancia"
                                                
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
                        Trim(adoRegistro("CodAnalitica")) & "') }"

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
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "'"
                    
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
                'Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile)
                
                '*** Obtener tipo de cambio ***
                dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
                
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
                    "CodFondo='" & strCodFondo & "'"
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
                    "CodFondo='" & strCodFondo & "'"
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
                        strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' " & _
                        "ORDER BY NumSecuencial"
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
                            Trim(adoRegistro("CodAnalitica")) & "') }"
                        adoConn.Execute .CommandText
                    
                        '*** Saldos ***
                        .CommandText = "{ call up_ACGenPartidaContableSaldos('"
                        .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                            Left(strFechaSiguiente, 4) & "','" & Mid(strFechaSiguiente, 5, 2) & "','" & _
                            Trim(adoConsulta("CodCuenta")) & "','" & _
                            Trim(adoRegistro("CodFile")) & "','" & _
                            Trim(adoRegistro("CodAnalitica")) & "','" & _
                            strFechaSiguiente & "','" & _
                            strFechaSubSiguiente & "'," & _
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
            
                End If
            
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
               ' Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile)
            
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
            
            '*** Periodo Actual ***
            strSQL = "{ call up_CNSelPeriodoContableVigente ('" & strCodFondo & "','" & gstrCodAdministradora & "') }"
            CargarControlLista strSQL, cboPeriodoContable, arrPeriodo(), ""
        
            If cboPeriodoContable.ListCount > 0 Then cboPeriodoContable.ListIndex = 0
            
            strCodMoneda = adoRegistro("CodMoneda")
            lblValorAIR(0).Caption = CStr(adoRegistro("ValorCuotaInicial"))
            lblValorDIR(0).Caption = "0"
            lblRentabilidad(1).Caption = "0"
            
            frmMainMdi.txtFechaSistema.Text = CStr(adoRegistro("FechaCuota"))
            cmdCierre.Enabled = True
        Else
            cmdCierre.Enabled = False
            MsgBox "Periodo contable no vigente para este fondo! Debe aperturar primero un periodo contable para este fondo!", vbExclamation + vbOKOnly, Me.Caption
            Exit Sub
        End If
        adoRegistro.Close
        
        .CommandText = " SELECT ValorCuotaInicial,ValorCuotaFinal FROM FondoValorCuota " & _
            "WHERE (FechaCuota >='" & strFechaAnterior & "' AND FechaCuota <'" & strFechaCierre & "') AND " & _
            "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            lblValorAIR(1).Caption = CStr(adoRegistro("ValorCuotaInicial"))
            lblValorDIR(1).Caption = CStr(adoRegistro("ValorCuotaFinal"))
            
            If CDbl(lblValorAIR(0).Caption) > 0 Then
                lblRentabilidad(0).Caption = CStr((((CDbl(lblValorDIR(0).Caption) / CDbl(lblValorAIR(0).Caption)) ^ 365) - 1) * 100)
            End If
        Else
            lblValorAIR(1).Caption = "0"
            lblValorDIR(1).Caption = "0"
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
        lblDescrip(7).Caption = Convertddmmyyyy(strFechaAnterior)


    End With
    
End Sub

Private Sub ValidarFechas()
    
    If EsDiaUtil(dtpFechaEntrega.Value) Then
      dtpFechaEntrega.Value = dtpFechaEntrega
    Else
      dtpFechaEntrega.Value = ProximoDiaUtil(dtpFechaEntrega)
    End If
    
End Sub

Private Sub cboPeriodoContableActual_Change()

End Sub

Private Sub cboPeriodoContable_Click()

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

    'Call ValidarFechas
    
    Call BuscarTipoCambio
        


End Sub

Private Sub cmdCierre_Click()

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
    If TodoOK() Then
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

        'Ajuste por tipo de cambio
        If strFechaCierre = strFechaCierreHasta Then
            '*** Actualización Inicial de Saldos Finales y Monto de Ajuste Contable ***
            frmMainMdi.stbMdi.Panels(3).Text = "Actualizando Montos de Ajuste por Tipo de Cambio..."
            
            Dim strFecSiguiente As String
            
            strFecSiguiente = DateAdd("d", 1, datFechaCierre)
            
            adoComm.CommandText = "{ call up_GNActMontoAjusteContable('" & _
                strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strFechaCierre & "','" & strFecSiguiente & "','" & _
                gstrCodClaseTipoCambioFondo & "','','"

            If chkSimulacion.Value Then
               adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Simulacion & "') }"
            Else
               adoComm.CommandText = adoComm.CommandText & Codigo_Cierre_Definitivo & "') }"
            End If
            adoConn.Execute adoComm.CommandText
            Sleep 0&
                            
            '*** Asientos de Pérdida/Ganancia por Tipo de Cambio ***
            frmMainMdi.stbMdi.Panels(3).Text = "Registrando Asientos Contables por Ajuste en el Tipo de Cambio..."
            
            adoComm.CommandText = "{ call up_GNProcAjusteTipoCambio1('" & _
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
        End If

        '*** Verifica y provisiona gastos del fondo - no incluye comision por adm de cartera ***
        If chkSimulacion.Value Then
           Call ProvisionGastosFondo(Codigo_Cierre_Simulacion, Valor_Caracter)
        Else
           Call ProvisionGastosFondo(Codigo_Cierre_Definitivo, Valor_Caracter)
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
                          
            '*** Actualizar datos de Cuotas ***
            frmMainMdi.stbMdi.Panels(3).Text = "Actualizando Datos de Cuotas..."
            
            adoComm.CommandText = "{ call up_GNActKardexFinalCuotasFondo('" & _
                strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strFechaCierre & "'," & _
                dblValNuevaCuota & "," & dblValNuevaCuotaReal & "," & _
                dblTCCierre & "," & dblTasa & ",'X','X','','',"
            
            If chkSimulacion.Value Then
               adoComm.CommandText = adoComm.CommandText & " '" & Codigo_Cierre_Simulacion & "') }"
            Else
               adoComm.CommandText = adoComm.CommandText & " '" & Codigo_Cierre_Definitivo & "') }"
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
    End If
    
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

Private Function TodoOK() As Boolean
                
    Dim adoConsulta As ADODB.Recordset
    Dim strMensaje  As String
    Dim adoFondo As ADODB.Recordset
    
    TodoOK = False
                
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
    
    TodoOK = True
  
End Function
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



Private Sub Form_Load()
  
    Call InicializarValores
    Call CargarListas
    Call DarFormato
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
        
    CentrarForm Me
      
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
      
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
Private Sub CargarListas()
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(29,'" & gstrCodAdministradora & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
            
End Sub
Private Sub InicializarValores()

    dtpFechaCierre.Value = gdatFechaActual
    dtpFechaEntrega.Value = DateAdd("d", gintDiasPagoRescate, dtpFechaCierre.Value)
    
    Call ValidarFechas
  
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
            dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
                        
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
                    strCodFile & "','" & strCodAnalitica & "') }"
                adoConn.Execute .CommandText
        
                '*** Saldos ***
                .CommandText = "{ call up_ACGenPartidaContableSaldos('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    gstrPeriodoActual & "','" & gstrMesActual & "','" & _
                    strCodCuenta & "','" & _
                    strCodFile & "','" & strCodAnalitica & "','" & _
                    strFechaCierre & "','" & strFechaSiguiente & "'," & _
                    CDec(curMontoMN) & "," & _
                    CDec(curMontoME) & "," & _
                    CDec(curMontoContable) & ",'" & _
                    strIndDebeHaber & "','" & strCodMonedaOperacion & "') }"
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

Private Sub ProvisionGastosFondo(strTipoCierre As String, strIndNoIncluyeEnPreCierre As String)

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
    Dim dblTipoCambioCierre     As Double
    
    '*** Rentabilidad de Valores de Depósito ***
    frmMainMdi.stbMdi.Panels(3).Text = "Calculando rentabilidad de Valores de Depósito..."
       
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT IK.CodTitulo,II.CodMoneda,IK.CodAnalitica,IK.CodFile,SaldoFinal,CodDetalleFile,CodSubDetalleFile," & _
            "ValorNominal,CodTipoTasa,BaseAnual,CodTasa,TasaInteres,DiasPlazo,IndCuponCero,FechaEmision,Nemotecnico " & _
            "FROM InversionKardex IK JOIN InstrumentoInversion II " & _
            "ON(II.CodTitulo=IK.CodTitulo) " & _
            "WHERE IK.CodFile IN ('003','011') AND FechaOperacion <='" & strFechaCierre & "' AND SaldoFinal > 0 AND " & _
            "IK.CodAdministradora='" & gstrCodAdministradora & "' AND IK.CodFondo='" & strCodFondo & "' AND IndUltimoMovimiento='X'"
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
            curSaldoValorizar = CCur(adoRegistro("SaldoFinal"))
            intDiasDeRenta = DateDiff("d", CVDate(adoRegistro("FechaEmision")), gdatFechaActual) + 1
            
            If strBaseAnual = Codigo_Base_30_360 Or strBaseAnual = Codigo_Base_30_365 Then intDiasDeRenta = Dias360(CVDate(adoRegistro("FechaEmision")), gdatFechaActual, True) + 1
            
            Set adoConsulta = New ADODB.Recordset
                        
            '*** Verificar Dinamica Contable ***
            .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
                "WHERE TipoOperacion='" & Codigo_Dinamica_Provision & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "'"
                
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
            .CommandText = "SELECT FactorDiario FROM InstrumentoInversionCalendario " & _
                "WHERE CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                dblFactorDiarioCupon = CDbl(adoConsulta("FactorDiario"))
            End If
            adoConsulta.Close
            
            '*** Obtener las cuentas de inversión ***
            Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, adoRegistro("CodMoneda"))
            
            '*** Obtener tipo de cambio ***
            dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
            
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
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
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
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvInteres & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoFluctuacion = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
                        
            curValorAnterior = curSaldoInteresCorrido + curSaldoFluctuacion
            
            If strBaseAnual = Codigo_Base_Actual_365 Or strBaseAnual = Codigo_Base_30_365 Or strBaseAnual = Codigo_Base_Actual_Actual Then
                intBaseCalculo = 365
            Else
                intBaseCalculo = 360
            End If
            
            If Trim(adoRegistro("CodSubDetalleFile")) <> Valor_Caracter Then strModalidadInteres = Trim(adoRegistro("CodSubDetalleFile"))
            
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
'                End If
                
                curValorActual = Round(curSaldoValorizar * dblFactorDiario, 2)
           
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
                
                '.CommandText = "BEGIN TRAN ProcAsiento"
                'adoConn.Execute .CommandText
                
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
                        CDec(curMontoContable) & ",'" & _
                        Trim(adoRegistro("CodFile")) & "','" & _
                        Trim(adoRegistro("CodAnalitica")) & "') }"

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
                                
                '.CommandText = "COMMIT TRAN ProcAsiento"
                'adoConn.Execute .CommandText
        
            End If
                                    
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
                
    End With
    
    Exit Sub
       
Ctrl_Error:
    'adoComm.CommandText = "ROLLBACK TRAN ProcAsiento"
    'adoConn.Execute adoComm.CommandText
    
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
                strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "'"
                
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
            'Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile)
            
            '*** Obtener tipo de cambio ***
            dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
            
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
                "CodFondo='" & strCodFondo & "'"
            
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
                
                '.CommandText = "BEGIN TRAN ProcAsiento"
                'adoConn.Execute .CommandText
                
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
                        CDec(curMontoContable) & ",'" & _
                        Trim(adoRegistro("CodFile")) & "','" & _
                        Trim(adoRegistro("CodAnalitica")) & "') }"

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
                                
                '.CommandText = "COMMIT TRAN ProcAsiento"
                'adoConn.Execute .CommandText
        
            End If
                                    
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
                
    End With
    
    Exit Sub
       
Ctrl_Error:
    'adoComm.CommandText = "ROLLBACK TRAN ProcAsiento"
    'adoConn.Execute adoComm.CommandText
    
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
    
    '*** Rentabilidad de Valores de Renta Fija Largo Plazo ***
    frmMainMdi.stbMdi.Panels(3).Text = "Calculando rentabilidad de Valores de Renta Fija Largo Plazo..."
       
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT IK.CodTitulo,II.CodMoneda,IK.CodAnalitica,IK.CodFile,SaldoFinal,SaldoAmortizacion,CodDetalleFile,CodSubDetalleFile,CuponCalculo,CodTipoVac," & _
            "ValorNominal,CodTipoTasa,BaseAnual,CodTasa,TasaInteres,DiasPlazo,IndCuponCero,FechaEmision,CodTipoAjuste,TirPromedio,Nemotecnico,PeriodoPago " & _
            "FROM InversionKardex IK JOIN InstrumentoInversion II " & _
            "ON(II.CodTitulo=IK.CodTitulo) " & _
            "WHERE IK.CodFile IN ('005') AND FechaOperacion <='" & strFechaCierre & "' AND SaldoFinal > 0 AND " & _
            "IK.CodAdministradora='" & gstrCodAdministradora & "' AND IK.CodFondo='" & strCodFondo & "' AND IndUltimoMovimiento='X'"
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
            curSaldoValorizar = CCur(adoRegistro("SaldoAmortizacion"))

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
                strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "'"
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
            .CommandText = "{ call up_IVSelDatoInstrumentoInversion(2,'" & Trim(adoRegistro("CodTitulo")) & "') }"
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
           ' Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile)
            
            '*** Obtener tipo de cambio ***
            dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
            
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
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaInteresCorrido & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "'"
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
                "CodFondo='" & strCodFondo & "'"
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
                "CodFondo='" & strCodFondo & "'"
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
                "CodFondo='" & strCodFondo & "'"
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
                "CodFondo='" & strCodFondo & "'"
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
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvFlucMercado & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoFluctuacionMercado = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
                        
            curValorAnterior = curSaldoInteresCorrido + curSaldoFluctuacion
            
            If Trim(adoRegistro("CodSubDetalleFile")) <> Valor_Caracter Then strModalidadInteres = Trim(adoRegistro("CodSubDetalleFile"))
            
            '*** REVISAR ***
            '*** Cálculo de Provisión de Intereses ***

                curValorActual = CalculoInteresCorrido(adoRegistro("CodTitulo"), CDbl(curSaldoValorizar), adoRegistro("FechaEmision"), DateAdd("d", 1, dtpFechaCierre.Value), strCodIndiceFinal, strCodTipoAjuste, strCodTasa, strCodPeriodoPago, strCodIndiceInicial, strCodBaseCalculo, intBaseCalculo)
                
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
                    
                    curValorActual = VNANoPer(adoRegistro("CodTitulo"), datFechaSiguienteGP, datFechaSiguienteGP, curSaldoValorizar, curSaldoValorizar, dblValorTir, adoRegistro("CodTipoAjuste"), strCodIndiceInicial, strCodIndiceFinal)
    
                    '*** CALCULO DEL MONTO DE GANANCIA/PERDIDA DE curCapital ***
                    curMontoProvisionCapital = Round(curValorActual - curValorAnterior - curMontoProvision, 2)
                Else
                    curMontoProvisionCapital = 0
                End If
    
            'End If
            
            '*** Cálculo Fluctuación Mercado ***
            If dblTirCierre > 0 Then
                curValorAnterior = curSaldoInversion + curSaldoInteresCorrido + curSaldoFluctuacion + curSaldoFluctuacionVac + curSaldoFluctuacionReajuste + curSaldoVacCorrido + curSaldoGPCapital + curSaldoFluctuacionMercado
                
                curValorActual = VNANoPer(adoRegistro("CodTitulo"), datFechaSiguienteGP, datFechaSiguienteGP, curSaldoValorizar, curSaldoValorizar, dblTirCierre, adoRegistro("CodTipoAjuste"), strCodIndiceInicial, strCodIndiceFinal)
                
                curMontoFluctuacionMercado = Round(curValorActual - curValorAnterior - curMontoProvision - curMontoProvisionCapital, 2)
            Else
                curMontoFluctuacionMercado = 0
            End If

                        
            '*** Contabilización ***
            If curMontoProvision <> 0 Or curMontoProvisionCapital <> 0 Or curMontoFluctuacionMercado <> 0 Then
                strDescripAsiento = "Valorización" & Space(1) & strNemonico
                strDescripMovimiento = "Pérdida"
                If curMontoProvision > 0 Then strDescripMovimiento = "Ganancia"
                                                
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
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' " & _
                    "ORDER BY NumSecuencial"
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
                            curMontoMovimientoMN = curMontoAjusteVAC
                            
                        Case Codigo_CtaReajusteK
                            curMontoMovimientoMN = curMontoAjusteVAC
                            
                        Case Codigo_CtaProvFlucMercado
                            curMontoMovimientoMN = curMontoFluctuacionMercado
                            strDescripMovimiento = "Pérdida"
                            If curMontoFluctuacionMercado > 0 Then strDescripMovimiento = "Ganancia"
                            
                        Case Codigo_CtaFlucMercado
                            curMontoMovimientoMN = curMontoFluctuacionMercado
                            
                        Case Codigo_CtaProvInteresVac
                            curMontoMovimientoMN = curMontoInteresVAC
                            
                        Case Codigo_CtaInteresVac
                            curMontoMovimientoMN = curMontoInteresVAC
                            
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
                        Trim(adoRegistro("CodAnalitica")) & "') }"

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
        
            End If
                                    
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
                
    End With
    
    Exit Sub
       
Ctrl_Error:

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
                strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "'"
                
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
            'Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile)
            
            '*** Obtener tipo de cambio ***
            dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
            
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
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaInteresCorrido & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "'"
            
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
                "CodFondo='" & strCodFondo & "'"
            
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
                "CodFondo='" & strCodFondo & "'"
            
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
                "CodFondo='" & strCodFondo & "'"
            
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
                "CodFondo='" & strCodFondo & "'"
            
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
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvFlucMercado & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "'"
            
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
                'If strIndCuponCero = Valor_Indicador Then
'                    dblFactorDiario = dblFactorDiarioCupon
                'Else
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
    
                    'vntFeccie = CVar(Right$(strFeccie, 2) + "/" + Mid$(strFeccie, 5, 2) + "/" + Left$(strFeccie, 4))
                    'vntFeccieMas1Dia = CVar(Right$(strFechaSiguiente, 2) + "/" + Mid$(strFechaSiguiente, 5, 2) + "/" + Left$(strFechaSiguiente, 4))
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
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' " & _
                    "ORDER BY NumSecuencial"
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
                        Trim(adoRegistro("CodAnalitica")) & "') }"

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
    Dim dblTipoCambioCierre     As Double
    
    '*** Rentabilidad de Reportes ***
    frmMainMdi.stbMdi.Panels(3).Text = "Calculando rentabilidad de Operaciones de Reporte..."
       
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT IK.CodTitulo,II.CodMoneda,IK.CodAnalitica,IK.CodFile,SaldoFinal,MontoSaldo,CodDetalleFile,CodSubDetalleFile," & _
            "ValorNominal,CodTipoTasa,BaseAnual,CodTasa,TasaInteres,DiasPlazo,IndCuponCero,FechaEmision,Nemotecnico " & _
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
            
            Set adoConsulta = New ADODB.Recordset
                        
            '*** Verificar Dinamica Contable ***
            .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
                "WHERE TipoOperacion='" & Codigo_Dinamica_Provision & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "'"
                
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
            .CommandText = "SELECT FactorDiario FROM InstrumentoInversionCalendario " & _
                "WHERE CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "'"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                dblFactorDiarioCupon = CDbl(adoConsulta("FactorDiario"))
            End If
            adoConsulta.Close
            
            '*** Obtener las cuentas de inversión ***
            'Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile)
            
            '*** Obtener tipo de cambio ***
            dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
            
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
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
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
            
            If strTipoCierre = Codigo_Cierre_Simulacion Then
                .CommandText = .CommandText & "FROM PartidaContableSaldosTmp "
            Else
                .CommandText = .CommandText & "FROM PartidaContableSaldos "
            End If
            
            .CommandText = .CommandText & "WHERE (FechaSaldo >='" & strFechaCierre & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvInteres & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoFluctuacion = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
                        
            curValorAnterior = curSaldoInteresCorrido + curSaldoFluctuacion
            
            If adoRegistro("BaseAnual") = Codigo_Base_Actual_365 Or adoRegistro("BaseAnual") = Codigo_Base_30_365 Or adoRegistro("BaseAnual") = Codigo_Base_Actual_Actual Then
                intBaseCalculo = 365
            Else
                intBaseCalculo = 360
            End If
            
            If Trim(adoRegistro("CodSubDetalleFile")) <> Valor_Caracter Then strModalidadInteres = Trim(adoRegistro("CodSubDetalleFile"))
            
            If strModalidadInteres = Codigo_Interes_Descuento Then
                If strCodTasa = Codigo_Tipo_Tasa_Efectiva Then
                    dblFactorDiario = ((1 + CDbl(((1 + (dblTasaInteres / 100) ^ (intDiasPlazo / intBaseCalculo))) - 1)) ^ (1 / intDiasPlazo)) - 1
                    curValorActual = (curSaldoValorizar * (((1 + dblFactorDiario) ^ (intDiasDeRenta)) - 1))
                Else
                    dblFactorDiario = (CDbl(((1 + (dblTasaInteres / 100)) / intBaseCalculo)))
                    curValorActual = (curSaldoValorizar * (dblFactorDiario * intDiasDeRenta))
                End If
            Else
                If strIndCuponCero = Valor_Indicador Then
                    dblFactorDiario = dblFactorDiarioCupon
                Else
                    If strCodTasa = Codigo_Tipo_Tasa_Efectiva Then
                        dblFactorDiario = ((1 + CDbl(((1 + (dblTasaInteres / 100)) ^ (intDiasPlazo / intBaseCalculo)) - 1)) ^ (1 / intDiasPlazo)) - 1
                    Else
                        dblFactorDiario = (CDbl(((1 + (dblTasaInteres / 100)) / intBaseCalculo)))
                    End If
                End If
                If strCodTasa = Codigo_Tipo_Tasa_Efectiva Then
                    curValorActual = (curSaldoValorizar * (((1 + dblFactorDiario) ^ (intDiasDeRenta)) - 1))
                Else
                    curValorActual = (curSaldoValorizar * (dblFactorDiario * intDiasDeRenta))
                End If
            End If
                        
            curMontoRenta = Round(curValorActual - curValorAnterior, 2)
            
            '*** Ganancia/Pérdida ***
            If curMontoRenta <> 0 Then
                strDescripAsiento = "Valorización" & Space(1) & strNemonico
                strDescripMovimiento = "Pérdida"
                If curMontoRenta > 0 Then strDescripMovimiento = "Ganancia"
                                                
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
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND DCodAdministradora='" & gstrCodAdministradora & "' " & _
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
                        CDec(curMontoContable) & ",'" & _
                        Trim(adoRegistro("CodFile")) & "','" & _
                        Trim(adoRegistro("CodAnalitica")) & "') }"

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
Private Sub ValorizacionRentaVariable(strTipoCierre As String)

    Dim adoRegistro             As ADODB.Recordset, adoConsulta     As ADODB.Recordset
    Dim dblPrecioCierre         As Double, dblPrecioPromedio        As Double
    Dim dblTirCierre            As Double, curSaldoInversion        As Currency
    Dim curSaldoFluctuacion     As Currency, curValorAnterior       As Currency
    Dim curValorActual          As Currency, curMontoRenta          As Currency
    Dim curMontoContable        As Currency, curMontoMovimientoMN   As Currency
    Dim curMontoMovimientoME    As Currency, dblTipoCambioCierre    As Double
    Dim intCantRegistros        As Integer, intContador             As Integer
    Dim intRegistro             As Integer
    Dim strNumAsiento           As String, strDescripAsiento        As String
    Dim strDescripMovimiento    As String, strIndDebeHaber          As String
    Dim strCodFile              As String, strCodDetalleFile        As String
    Dim strCodCuenta            As String, strNemonico              As String
    Dim strFechaGrabar          As String
    Dim strTipoAuxiliar         As String
    Dim strCodAuxiliar         As String
    Dim strCtaCostoInversion    As String
    
    '*** Rentabilidad de Valores de Renta Variable ***
    frmMainMdi.stbMdi.Panels(3).Text = "Calculando rentabilidad de Valores de Renta Variable..."
       
    Set adoRegistro = New ADODB.Recordset
    
    strTipoAuxiliar = "01" 'Inversiones
    
    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT IK.CodTitulo,II.CodMoneda,IK.CodAnalitica,IK.CodFile,SaldoFinal,MontoSaldo,MontoMovimiento,MontoComision,CodDetalleFile,CodSubDetalleFile," & _
            "ValorNominal,CodTipoTasa,BaseAnual,CodTasa,TasaInteres,DiasPlazo,IndCuponCero,FechaEmision,Nemotecnico " & _
            "FROM InversionKardex IK LEFT JOIN InstrumentoInversion II " & _
            "ON(II.CodTitulo=IK.CodTitulo) " & _
            "WHERE IK.CodFile IN ('004') AND FechaOperacion <='" & strFechaCierre & "' AND SaldoFinal > 0 AND " & _
            "IK.CodAdministradora='" & gstrCodAdministradora & "' AND IK.CodFondo='" & strCodFondo & "' AND IndUltimoMovimiento='X'"
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
                strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "'"
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
            .CommandText = "{ call up_IVSelDatoInstrumentoInversion(2,'" & Trim(adoRegistro("CodTitulo")) & "') }"
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                dblPrecioCierre = CDbl(adoConsulta("PrecioCierre"))
                dblTirCierre = CDbl(adoConsulta("TirCierre"))
                dblPrecioPromedio = CDbl(adoConsulta("PrecioPromedio"))
            End If
            adoConsulta.Close
            
            '*** Obtener las cuentas de inversión ***
            'Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile)
            
            '*** Obtener tipo de cambio ***
            dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
            
            '*** Obtener Saldo de Inversión ***
            If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                .CommandText = "SELECT SUM(SaldoFinalContable) Saldo "
            Else
                .CommandText = "SELECT SUM(SaldoFinalME) Saldo "
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
                "CodFondo='" & strCodFondo & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoInversion = CDbl(adoConsulta("Saldo"))
            Else
                curSaldoInversion = 0
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
                "CodCuenta='" & strCtaProvFlucMercado & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "'"
            
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
            
            '*** Ganancia/Pérdida ***
            If curMontoRenta <> 0 Then
                strDescripAsiento = "Valorización" & Space(1) & strNemonico
                strDescripMovimiento = "Pérdida"
                If curMontoRenta > 0 Then strDescripMovimiento = "Ganancia"
                                
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
                    gstrPeriodoActual & "','" & gstrMesActual & "','','" & _
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
    Dim lngLiberadas        As Long, curDividendos              As Currency
    Dim strNumOrdenEvento   As String

    '*** Verificar vencimientos de entregas de acciones ***
    frmMainMdi.stbMdi.Panels(3).Text = "Verificando vencimientos y entregas de acciones..."

    Set adoRegistro = New ADODB.Recordset
    Set adoRegistroMov = New ADODB.Recordset
    
    With adoComm
        '*** Verificar también a la fecha de Corte para ver la cantidad de Acciones ***
        '*** que tienen Derecho ***
        .CommandText = "SELECT * FROM EventoCorporativoAcuerdo " & _
            "WHERE (FechaCorte>='" & strFechaCierre & "' AND FechaCorte<'" & strFechaSiguiente & "') AND " & _
            "CodAdministradora='" & gstrCodAdministradora & "' AND EstadoEvento='" & Estado_Acuerdo_Ingresado & "'"
        Set adoRegistro = .Execute
        
        Do Until adoRegistro.EOF
            .CommandText = "SELECT * FROM InversionKardex " & _
                "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodTitulo='" & adoRegistro("CodTitulo") & "' AND " & _
                "FechaMovimiento<='" & Convertyyyymmdd(adoRegistro("FechaCorte")) & "'"
            Set adoRegistroMov = .Execute
            
            If Not adoRegistroMov.EOF Then
                '*** Ver liberadas ***
                If CDbl(adoRegistro("PorcenAccionesLiberadas")) > 0 Then
                    lngLiberadas = CLng(CDbl(adoRegistro("PorcenAccionesLiberadas")) * 0.01 * adoRegistroMov("SaldoFinal"))
                    '*** Obtener Secuencial ***
                    strNumOrdenEvento = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumEntregaEvento)
                .CommandText = "{ call up_GNAdicEventoCorporativoOrden('" & _
                        strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        adoRegistro("CodTitulo") & "'," & CInt(adoRegistro("NumAcuerdo")) & "," & _
                        CLng(strNumOrdenEvento) & ",'" & adoRegistro("CodFile") & "','" & _
                        adoRegistro("CodAnalitica") & "','','','','" & strFechaSiguiente & "'," & _
                        CInt(adoRegistroMov("SaldoFinal")) & "," & lngLiberadas & "," & _
                        "0,0,0,0,0,'" & strFechaCierre & "','" & adoRegistro("FechaEntrega") & "','" & _
                        "No Contabilizado" & "','','" & Estado_Entrega_Generado & "'," & _
                        "0,'" & Codigo_Evento_Liberacion & "','" & _
                        gstrLogin & "','" & strFechaCierre & "','" & _
                        gstrLogin & "','" & strFechaCierre & "') }"
                    adoConn.Execute .CommandText
                    
                    '*** Actualiza Secuencial **
                    .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        Valor_NumEntregaEvento & "','" & strNumOrdenEvento & "') }"
                    adoConn.Execute .CommandText
                End If

                '*** Ver dividendos ***
                If (CDbl(adoRegistro("PorcenDividendoEfectivo")) * CDbl(adoRegistroMov("SaldoFinal"))) > 0 Then
                    curDividendos = CCur(adoRegistro("PorcenDividendoEfectivo") * adoRegistroMov("SaldoFinal"))
                    '*** Obtener Secuencial ***
                    strNumOrdenEvento = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumEntregaEvento)

                        .CommandText = "{ call up_GNAdicEventoCorporativoOrden('" & _
                        strCodFondo & "','" & gstrCodAdministradora & "','" & _
                        adoRegistro("CodTitulo") & "'," & CInt(adoRegistro("NumAcuerdo")) & "," & _
                        CLng(strNumOrdenEvento) & ",'" & adoRegistro("CodFile") & "','" & _
                        adoRegistro("CodAnalitica") & "','','','','" & strFechaSiguiente & "'," & _
                        CInt(adoRegistroMov("SaldoFinal")) & ",0," & _
                        CInt(adoRegistro("PorcenDividendoEfectivo") * adoRegistroMov("SaldoFinal")) & ",0,0,0,0,'" & strFechaCierre & "','" & adoRegistro("FechaEntrega") & "','" & _
                        "No Contabilizado" & "','','" & Estado_Entrega_Generado & "'," & _
                        "0,'" & Codigo_Evento_Dividendo & "','" & _
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
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "'"
                    
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
               ' Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile)
                
                '*** Obtener tipo de cambio ***
                dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
                
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
                        strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' " & _
                        "ORDER BY NumSecuencial"
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
                            Trim(adoRegistro("CodAnalitica")) & "') }"
                        adoConn.Execute .CommandText
                    
                        '*** Saldos ***
                        .CommandText = "{ call up_ACGenPartidaContableSaldos('"
                        .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                            Left(strFechaSiguiente, 4) & "','" & Mid(strFechaSiguiente, 5, 2) & "','" & _
                            Trim(adoConsulta("CodCuenta")) & "','" & _
                            Trim(adoRegistro("CodFile")) & "','" & _
                            Trim(adoRegistro("CodAnalitica")) & "','" & _
                            strFechaSiguiente & "','" & _
                            strFechaSubSiguiente & "'," & _
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

            
                End If
            
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
        .CommandText = "SELECT IK.CodTitulo,II.CodMoneda,IK.CodAnalitica,IK.CodFile,SaldoFinal,CodDetalleFile,CodSubDetalleFile," & _
            "ValorNominal,CodTipoTasa,BaseAnual,CodTasa,TasaInteres,DiasPlazo,IndCuponCero,FechaEmision,FechaVencimiento," & _
            "II.CodEmisor,IK.PrecioUnitario,IK.MontoMovimiento,IK.SaldoInteresCorrido,IK.MontoSaldo,IK.ValorPromedio " & _
            "FROM InversionKardex IK JOIN InstrumentoInversion II " & _
            "ON(II.CodTitulo=IK.CodTitulo) " & _
            "WHERE IK.CodFile IN ('003','011') AND FechaOperacion <='" & strFechaCierre & "' AND SaldoFinal > 0 AND " & _
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
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "'"
                    
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
                'Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile)
                
                '*** Obtener tipo de cambio ***
                dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
                
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
                        strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' " & _
                        "ORDER BY NumSecuencial"
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
                            Trim(adoRegistro("CodAnalitica")) & "') }"
                        adoConn.Execute .CommandText
                    
                        '*** Saldos ***
                        .CommandText = "{ call up_ACGenPartidaContableSaldos('"
                        .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                            Left(strFechaSiguiente, 4) & "','" & Mid(strFechaSiguiente, 5, 2) & "','" & _
                            Trim(adoConsulta("CodCuenta")) & "','" & _
                            Trim(adoRegistro("CodFile")) & "','" & _
                            Trim(adoRegistro("CodAnalitica")) & "','" & _
                            strFechaSiguiente & "','" & _
                            strFechaSubSiguiente & "'," & _
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
            
                End If
            
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
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "'"
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
                'Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile)
                
                '*** Obtener tipo de cambio ***
                dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
                
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
                        strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' " & _
                        "ORDER BY NumSecuencial"
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
                            Trim(adoRegistro("CodAnalitica")) & "') }"
                        adoConn.Execute .CommandText
                    
                        '*** Saldos ***
                        .CommandText = "{ call up_ACGenPartidaContableSaldos('"
                        .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                            Left(strFechaSiguiente, 4) & "','" & Mid(strFechaSiguiente, 5, 2) & "','" & _
                            Trim(adoConsulta("CodCuenta")) & "','" & _
                            Trim(adoRegistro("CodFile")) & "','" & _
                            Trim(adoRegistro("CodAnalitica")) & "','" & _
                            strFechaSiguiente & "','" & _
                            strFechaSubSiguiente & "'," & _
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

                End If
            
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
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "'"
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
                'Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile)
                
                '*** Obtener tipo de cambio ***
                dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
                
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
                        strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' " & _
                        "ORDER BY NumSecuencial"
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
                            Trim(adoRegistro("CodAnalitica")) & "') }"
                        adoConn.Execute .CommandText
                    
                        '*** Saldos ***
                        .CommandText = "{ call up_ACGenPartidaContableSaldos('"
                        .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                            Left(strFechaSiguiente, 4) & "','" & Mid(strFechaSiguiente, 5, 2) & "','" & _
                            Trim(adoConsulta("CodCuenta")) & "','" & _
                            Trim(adoRegistro("CodFile")) & "','" & _
                            Trim(adoRegistro("CodAnalitica")) & "','" & _
                            strFechaSiguiente & "','" & _
                            strFechaSubSiguiente & "'," & _
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
            
                End If
            
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
   
    Set adoConsultaTipoCambio = New ADODB.Recordset
    
    strFechaConsulta = Convertyyyymmdd(dtpFechaCierreHasta.Value)
                                                                     
    strSQL = "{ call up_ACObtieneTipoCambioFecha('" & gstrCodClaseTipoCambioFondo & "','" & strFechaConsulta & "') }"
                      
    With adoConsultaTipoCambio
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
  
    tdgTipoCambioCierre.DataSource = adoConsultaTipoCambio
        
    tdgTipoCambioCierre.Refresh
        
    'tdgTipoCambioCierre.Refresh
    
End Sub


Private Function ObtenerValorTCCierre(ByVal strpCodMoneda As String) As Double

    ObtenerValorTCCierre = 0
    
    If adoConsultaTipoCambio.EOF And adoConsultaTipoCambio.BOF Then Exit Function
    
    adoConsultaTipoCambio.MoveFirst
    'adoConsultaTipoCambio.Find ("CodMoneda='" & strpCodMoneda & "'")
    
    If strpCodMoneda = Codigo_Moneda_Local Then
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

