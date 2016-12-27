VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBandejaInversion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bandeja de Inversiones"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   13245
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   13245
      _ExtentX        =   23363
      _ExtentY        =   1111
      ButtonWidth     =   2355
      ButtonHeight    =   1005
      ImageList       =   "imlBandeja"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refrescar"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aprobar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Formalizar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Condic. Financ."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cuponera"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Desembolso"
            Description     =   "Desembolso"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCriterio 
      Caption         =   "Criterios de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   60
      TabIndex        =   0
      Top             =   630
      Width           =   13065
      Begin VB.ComboBox cboEstado 
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1050
         Width           =   4785
      End
      Begin VB.ComboBox cboTipoInstrumento 
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   690
         Width           =   4785
      End
      Begin VB.ComboBox cboFondo 
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   4785
      End
      Begin MSComCtl2.DTPicker dtpFechaOrdenDesde 
         Height          =   315
         Left            =   9120
         TabIndex        =   4
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
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
         CheckBox        =   -1  'True
         Format          =   175570945
         CurrentDate     =   38785
      End
      Begin MSComCtl2.DTPicker dtpFechaOrdenHasta 
         Height          =   315
         Left            =   11385
         TabIndex        =   5
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
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
         CheckBox        =   -1  'True
         Format          =   175570945
         CurrentDate     =   38785
      End
      Begin MSComCtl2.DTPicker dtpFechaLiquidacionDesde 
         Height          =   315
         Left            =   9120
         TabIndex        =   6
         Top             =   660
         Width           =   1455
         _ExtentX        =   2566
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
         CheckBox        =   -1  'True
         Format          =   175570945
         CurrentDate     =   38785
      End
      Begin MSComCtl2.DTPicker dtpFechaLiquidacionHasta 
         Height          =   315
         Left            =   11385
         TabIndex        =   7
         Top             =   660
         Width           =   1455
         _ExtentX        =   2566
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
         CheckBox        =   -1  'True
         Format          =   175570945
         CurrentDate     =   38785
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
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
         Index           =   23
         Left            =   210
         TabIndex        =   16
         Top             =   1125
         Width           =   600
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Instrumento"
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
         Index           =   22
         Left            =   210
         TabIndex        =   15
         Top             =   765
         Width           =   1005
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   21
         Left            =   10710
         TabIndex        =   14
         Top             =   315
         Width           =   510
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   20
         Left            =   8370
         TabIndex        =   13
         Top             =   345
         Width           =   555
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Fondo"
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
         Index           =   19
         Left            =   210
         TabIndex        =   12
         Top             =   375
         Width           =   540
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Solicitud"
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
         Index           =   43
         Left            =   6480
         TabIndex        =   11
         Top             =   345
         Width           =   1335
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Liquidación"
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
         Index           =   44
         Left            =   6480
         TabIndex        =   10
         Top             =   735
         Width           =   1560
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   45
         Left            =   8370
         TabIndex        =   9
         Top             =   735
         Width           =   555
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   46
         Left            =   10710
         TabIndex        =   8
         Top             =   705
         Width           =   510
      End
   End
   Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
      Bindings        =   "frmBandejaInversion.frx":0000
      Height          =   5205
      Left            =   60
      OleObjectBlob   =   "frmBandejaInversion.frx":001A
      TabIndex        =   18
      Top             =   2190
      Width           =   13065
   End
   Begin MSComctlLib.ImageList imlBandeja 
      Left            =   330
      Top             =   1860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBandejaInversion.frx":B05A
            Key             =   "NUEVO"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBandejaInversion.frx":CD64
            Key             =   "GUARDAR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBandejaInversion.frx":D63E
            Key             =   "BUSCAR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBandejaInversion.frx":DA92
            Key             =   "CONSULTAR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBandejaInversion.frx":F79C
            Key             =   "IMPRIMIR"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBandejaInversion.frx":114A6
            Key             =   "ELIMINAR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBandejaInversion.frx":122F8
            Key             =   "BLOQUEAR"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBandejaInversion.frx":1314A
            Key             =   "REPORTES"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBandejaInversion.frx":13F9C
            Key             =   "AYUDA"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBandejaInversion.frx":142B6
            Key             =   "CANCELAR"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBandejaInversion.frx":145D0
            Key             =   "PRIMERO"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBandejaInversion.frx":14EAA
            Key             =   "ANTERIOR"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBandejaInversion.frx":15784
            Key             =   "SIGUIENTE"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBandejaInversion.frx":1605E
            Key             =   "ULTIMO"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBandejaInversion.frx":16938
            Key             =   "REFRESCAR"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   1710
      Top             =   6900
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
End
Attribute VB_Name = "frmBandejaInversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()            As String
Dim arrTipoInstrumento()  As String
Dim arrEstado()           As String

Dim strCodFondo           As String, strCodigosFile              As String
Dim strCodTipoInstrumento As String
Dim strCodEstado          As String, strSQL                       As String

Public Sub Buscar()

    Dim strFechaSolicitudDesde   As String, strFechaSolicitudHasta        As String
    Dim strFechaLiquidacionDesde As String, strFechaLiquidacionHasta  As String
    Dim datFechaSiguiente        As Date

    Me.MousePointer = vbHourglass
    
    If Not IsNull(dtpFechaOrdenDesde.Value) And Not IsNull(dtpFechaOrdenHasta.Value) Then
        strFechaSolicitudDesde = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaOrdenHasta.Value)
        strFechaSolicitudHasta = Convertyyyymmdd(datFechaSiguiente)
    End If
    
    If Not IsNull(dtpFechaLiquidacionDesde.Value) And Not IsNull(dtpFechaLiquidacionHasta.Value) Then
        strFechaLiquidacionDesde = Convertyyyymmdd(dtpFechaLiquidacionDesde.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaLiquidacionHasta.Value)
        strFechaLiquidacionHasta = Convertyyyymmdd(datFechaSiguiente)
    End If
        
    strSQL = "SELECT IOR.NumSolicitud,FechaSolicitud,FechaLiquidacion,CodTitulo,EstadoSolicitud,IOR.CodFile,CodAnalitica,TipoSolicitud,IOR.CodMoneda," & "DescripSolicitud,MontoSolicitud,MontoAprobado, " & "CodSigno DescripMoneda, IOR.CodDetalleFile, IOR.CodSubDetalleFile, IOR.CodFondo, " & "IOR.CodEmisor, IP1.DescripPersona DesEmisor,EstadoSolicitud, EST.DescripParametro AS DescripEstado " & "FROM InversionSolicitud IOR JOIN AuxiliarParametro EST ON(EST.CodParametro=IOR.EstadoSolicitud AND EST.CodTipoParametro = 'ESTSCF') " & "JOIN Moneda MON ON(MON.CodMoneda=IOR.CodMoneda) " & "LEFT JOIN InstitucionPersona IP1 ON (IP1.CodPersona = IOR.CodEmisor AND IP1.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "') " & "WHERE IOR.CodAdministradora='" & gstrCodAdministradora & "' AND IOR.CodFondo='" & strCodFondo & "' "
        
    If strCodTipoInstrumento <> Valor_Caracter Then
        strSQL = strSQL & "AND IOR.CodFile='" & strCodTipoInstrumento & "' "
    Else
        strSQL = strSQL & "AND IOR.CodFile IN " & strCodigosFile & " "
    End If
    
    If Not IsNull(dtpFechaOrdenDesde.Value) And Not IsNull(dtpFechaOrdenHasta.Value) Then
        strSQL = strSQL & "AND (FechaSolicitud >='" & strFechaSolicitudDesde & "' AND FechaSolicitud <'" & strFechaSolicitudHasta & "') "
    End If
    
    If Not IsNull(dtpFechaLiquidacionDesde.Value) And Not IsNull(dtpFechaLiquidacionHasta.Value) Then
        strSQL = strSQL & "AND (FechaLiquidacion >='" & strFechaLiquidacionDesde & "' AND FechaLiquidacion <'" & strFechaLiquidacionHasta & "') "
    End If
    
    If strCodEstado <> Valor_Caracter Then
        strSQL = strSQL & "AND EstadoSolicitud='" & strCodEstado & "' "
    End If
    
    '    If (strLineaClienteLista <> Valor_Caracter) And (cboTipoInstrumento.ListIndex > 0) Then
    '        strSQL = strSQL & "AND IOR.CodLimiteCli='" & strLineaClienteLista & "' AND IOR.CodEstructura ='" & Codigo_LimiteRE_Cliente & "' "
    '    End If
    
    strSQL = strSQL & "ORDER BY IOR.NumSolicitud"
    
    With adoConsulta
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With

    tdgConsulta.Refresh

    Me.MousePointer = vbDefault
    
End Sub

Private Sub cboEstado_Click()
    strCodEstado = Valor_Caracter

    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strCodEstado = Trim$(arrEstado(cboEstado.ListIndex))
    
    Call Buscar
End Sub

Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondo = Valor_Caracter

    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim$(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Moneda ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "','000') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = CVDate(adoRegistro("FechaCuota"))
            dtpFechaOrdenDesde.Value = gdatFechaActual
            dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
            'strCodMoneda = trim$(adoRegistro("CodMoneda"))
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        End If

        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & _
            "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_CortoPlazo & _
            "' AND IndInstrumento='X' AND IndVigente='X' AND " & "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & _
            "' AND  FIF.CodFile = '" & CodFile_Descuento_Flujos_Dinerarios & "' " & _
            " ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumento, arrTipoInstrumento(), Sel_Todos
    
    If cboTipoInstrumento.ListCount > 0 Then cboTipoInstrumento.ListIndex = 0
        
End Sub

Private Sub cboTipoInstrumento_Click()
    strCodTipoInstrumento = Valor_Caracter

    If cboTipoInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumento = Trim$(arrTipoInstrumento(cboTipoInstrumento.ListIndex))
    
    Call Buscar

End Sub

Private Sub dtpFechaLiquidacionDesde_Click()

    If IsNull(dtpFechaLiquidacionDesde.Value) Then
        dtpFechaLiquidacionHasta.Value = Null
    Else
        dtpFechaLiquidacionDesde.Value = gdatFechaActual
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    End If

End Sub

Private Sub dtpFechaLiquidacionHasta_Click()

    If IsNull(dtpFechaLiquidacionHasta.Value) Then
        dtpFechaLiquidacionDesde.Value = Null
    Else
        dtpFechaLiquidacionDesde.Value = gdatFechaActual
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    End If

End Sub

Private Sub dtpFechaOrdenDesde_Click()

    If IsNull(dtpFechaOrdenDesde.Value) Then
        dtpFechaOrdenHasta.Value = Null
    Else
        dtpFechaOrdenDesde.Value = gdatFechaActual
        dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    End If

End Sub

Private Sub dtpFechaOrdenHasta_Click()

    If IsNull(dtpFechaOrdenHasta.Value) Then
        dtpFechaOrdenDesde.Value = Null
    Else
        dtpFechaOrdenDesde.Value = gdatFechaActual
        dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    End If

End Sub

Private Sub Form_Activate()
    frmMainMdi.stbMdi.Panels(3).Text = Me.Caption
    
    Call Buscar
End Sub

Private Sub Form_Load()
    Call InicializarValores
    Call CargarListas
    Call Buscar
    
    Call ValidarPermisoUsoControl(Trim$(gstrLogin), Me, Trim$(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)

    CentrarForm Me
    
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
End Sub

Private Sub tdgConsulta_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim strMostrarBotones As String
    Dim strBotones()      As String
    Dim i                 As Integer
    
    strMostrarBotones = Trim$(traerCampo("AuxiliarParametro", "ValorParametro", "CodParametro", "" & tdgConsulta.Columns("EstadoSolicitud").Value, " CodTipoParametro = 'ESTSCF'"))
    
    For i = 2 To Toolbar1.Buttons.Count - 1
        Toolbar1.Buttons(i).Visible = False
    Next
    
    strBotones = Split(strMostrarBotones, ",")
    
    For i = 0 To UBound(strBotones)
        Toolbar1.Buttons(CInt(strBotones(i))).Visible = True
    Next

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim strMsgError As String
    On Error GoTo err
    Dim adoAux As ADODB.Recordset

    Select Case Button.Index

        Case 1 'Refrescar
            Call Buscar

        Case 2 'Aprobar

            If MsgBox("¿Seguro de Aprobar la Solicitud?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                cambiarEstadoSolicitud "03"
            End If

        Case 3 'Formalizar
            frmSolicitudDescuentoContratos.setFondo (strCodFondo)
            frmSolicitudDescuentoContratos.cmdAccion.Visible = True
            frmSolicitudDescuentoContratos.mostrarForm Trim$(tdgConsulta.Columns(0).Value)
           

        Case 4 'Condiciones Financieras
            frmCronograma.beneficiario = Trim$(tdgConsulta.Columns("DesEmisor").Value)
            frmCronograma.codigoUnico = traerCampo("InversionSolicitud", "CodTitulo", "NumSolicitud", Trim$(tdgConsulta.Columns(0).Value), " CodFondo = '" & strCodFondo & "' AND CodAdministradora = '" & gstrCodAdministradora & "'")
            
            adoComm.CommandText = "Select count (*) as counter from InstrumentoInversionCondicionesFinancieras where CodTitulo = '" & frmCronograma.codigoUnico & "'"
            Set adoAux = adoComm.Execute
            If Not adoAux.EOF Then
                frmCronograma.flagVisor = adoAux("counter").Value = 1
            Else
                frmCronograma.flagVisor = False
            End If
            frmCronograma.codSolicitud = Trim$(tdgConsulta.Columns(0).Value)
    
            frmCronograma.Show

        Case 5 'Cuponera
           
            frmVisorCronograma.codigoUnico = traerCampo("InversionSolicitud", "CodTitulo", "NumSolicitud", Trim$(tdgConsulta.Columns(0).Value), " CodFondo = '" & strCodFondo & "' AND CodAdministradora = '" & gstrCodAdministradora & "'")
            frmVisorCronograma.strNumSolicitud = Trim$(tdgConsulta.Columns(0).Value)
            frmVisorCronograma.Source = 0
            frmVisorCronograma.setFondo (strCodFondo)
            frmVisorCronograma.Show
            
        Case 6 'Desembolso
            frmOrdenDescuentoContratos.setFondo (strCodFondo)
            frmOrdenDescuentoContratos.mostrarForm Trim$(tdgConsulta.Columns(0).Value)

        Case 7 'Salir
            Unload Me
    End Select

    Exit Sub
err:

    If strMsgError = "" Then strMsgError = err.Description
  
End Sub

Private Sub InicializarValores()
    
    Dim adoRegistro As ADODB.Recordset
    
    '*** Valores Iniciales ***
    dtpFechaOrdenDesde.Value = gdatFechaActual
    dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    dtpFechaLiquidacionDesde.Value = Null
    dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    
    Set adoRegistro = New ADODB.Recordset

    With adoComm
        .CommandText = "SELECT CodFile FROM InversionFile " & "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_CortoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' " & "ORDER BY DescripFile"
        Set adoRegistro = .Execute
                
        strCodigosFile = Valor_Caracter

        Do While Not adoRegistro.EOF

            If strCodigosFile <> Valor_Caracter Then strCodigosFile = strCodigosFile & ",'"
            
            strCodigosFile = strCodigosFile & Trim$(adoRegistro("CodFile")) & "'"
        
            adoRegistro.MoveNext
        Loop

        adoRegistro.Close: Set adoRegistro = Nothing
                
        strCodigosFile = "('" & strCodigosFile & ",'009')"
    End With
    
End Sub

Private Sub CargarListas()

    Dim intRegistro As Integer
    
    '*** Fondos ***
    '*** Fondos ***
    strSQL = "SELECT CodFondo CODIGO,DescripFondo DESCRIP FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND Estado='01' and CodFondo = '" & gstrCodFondoContable & "' ORDER BY DescripFondo"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
         
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0

            
    '*** Estados de la Solicitud ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE Estado = '01' AND CodTipoParametro='ESTSCF' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Sel_Todos
    
    intRegistro = ObtenerItemLista(arrEstado(), "02")

    If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
    
    '*** Carga Lista para el tab de formas de pago ***
    '    indCargaPantalla = True

End Sub

Private Sub cambiarEstadoSolicitud(ByVal strEstado As String)
    
    adoComm.CommandText = "UPDATE InversionSolicitud SET EstadoSolicitud='" & strEstado & "'," & "UsuarioEdicion='" & gstrLogin & "',FechaEdicion=getdate() " & "WHERE NumSolicitud='" & Trim$(tdgConsulta.Columns(0)) & "' AND CodFondo='" & strCodFondo & "' AND " & "CodAdministradora='" & gstrCodAdministradora & "'"

    adoConn.Execute adoComm.CommandText
        
    MsgBox "Cambio de estado realizado satisfactoriamente", vbInformation, App.Title
    
    Buscar
End Sub
