VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLimite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Limites"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   7110
   Begin VB.CommandButton Command1 
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
      Left            =   5160
      Picture         =   "frmLimite.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4680
      Width           =   1200
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   1920
      Top             =   4920
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
   Begin TabDlg.SSTab tabLimites 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7858
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmLimite.frx":0582
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dgdConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraLimite"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmLimite.frx":059E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraLimite 
         Caption         =   "Criterios de Búsqueda"
         Height          =   1360
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   6375
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   800
            Width           =   4815
         End
         Begin VB.ComboBox cboTipoLimite 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   360
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dtpFechaRegistro 
            Height          =   315
            Left            =   4560
            TabIndex        =   3
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   293404673
            CurrentDate     =   38790
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   3840
            TabIndex        =   8
            Top             =   380
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   7
            Top             =   375
            Width           =   315
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   6
            Top             =   825
            Width           =   450
         End
      End
      Begin MSDataGridLib.DataGrid dgdConsulta 
         Height          =   2295
         Left            =   240
         TabIndex        =   1
         Top             =   1920
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   4048
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Limites"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmLimite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Ficha de Mantenimento General de Limites"
Option Explicit

Dim arrTipoLimite()     As String, arrFondo()       As String

Dim strCodTipoLimite    As String, strCodFondo      As String
Dim strEstado           As String, strSQL           As String
Dim dblLimiteMinimo     As Double, dblLimiteMaximo  As Double

Public Sub Cancelar()

    Call Buscar
    
End Sub

Private Sub CargarDatos()
    
'    Dim intRes As Integer
'
'    With adoComm
'
'        Select Case strTipLimite
'            Case "01" '*** Patrimonio ***
'                .CommandText = "SELECT tblCtrlLimites.COD_FILE COD_LIMI, DSC_FILE DSC_LIMI,tblCtrlLimites.MIN_LIMI, tblCtrlLimites.MAX_LIMI, ' ' FLG_MODI "
'                .CommandText = .CommandText & "FROM tblCtrlLimites JOIN tblCodigosFile ON(tblCodigosFile.COD_FILE=tblCtrlLimites.COD_FILE) "
'                .CommandText = .CommandText & "WHERE COD_LIMI='" & strTipLimite & "' AND COD_FOND='" & strLimites & "' AND FLG_ULTI='X' "
'                .CommandText = .CommandText & "ORDER BY DSC_FILE"
'            Case "02" '*** Crédito Vigente ***
'                .CommandText = "SELECT tblCtrlLimites.COD_EMPR COD_LIMI, RAZ_SOCI DSC_LIMI,tblCtrlLimites.MIN_LIMI, tblCtrlLimites.MAX_LIMI, ' ' FLG_MODI "
'                .CommandText = .CommandText & "FROM tblCtrlLimites JOIN FMPERSON ON(FMPERSON.COD_PERS=tblCtrlLimites.COD_EMPR) "
'                .CommandText = .CommandText & "WHERE COD_LIMI='" & strTipLimite & "' AND COD_FOND='00' AND TIP_PERS='EM' AND FLG_VIGE='X' AND FLG_ULTI='X' "
'                .CommandText = .CommandText & "ORDER BY RAZ_SOCI"
'            Case "03" '*** Clasificación de Riesgo ***
'                .CommandText = "SELECT COD_FILE COD_LIMI, (RTRIM(VAL_PARA) + ' ' + RTRIM(DSC_PARA)) DSC_LIMI, MIN_LIMI, MAX_LIMI, ' ' FLG_MODI "
'                .CommandText = .CommandText & "FROM tblCtrlLimites JOIN tblParametros ON(tblParametros.COD_PARA=tblCtrlLimites.COD_FILE) "
'                .CommandText = .CommandText & "WHERE COD_LIMI='" & strTipLimite & "' AND COD_FOND='00' AND COD_TIPP='TIPRIE' AND FLG_ULTI='X'"
'
'            Case "04" '*** Instrumentos de Inversión ***
'                .CommandText = "SELECT tblCtrlLimites.COD_FILE COD_LIMI, DSC_FILE DSC_LIMI,tblCtrlLimites.MIN_LIMI, tblCtrlLimites.MAX_LIMI, ' ' FLG_MODI "
'                .CommandText = .CommandText & "FROM tblCtrlLimites JOIN tblCodigosFile ON(tblCodigosFile.COD_FILE=tblCtrlLimites.COD_FILE) "
'                .CommandText = .CommandText & "WHERE COD_LIMI='" & strTipLimite & "' AND COD_FOND='00' AND FLG_ULTI='X' "
'                .CommandText = .CommandText & "ORDER BY DSC_FILE"
'
'            Case "05" '*** Activo ***
'                .CommandText = "SELECT tblCtrlLimites.COD_FILE COD_LIMI, 'Participación Emisor (%)' DSC_LIMI,tblCtrlLimites.MIN_LIMI, tblCtrlLimites.MAX_LIMI, ' ' FLG_MODI "
'                .CommandText = .CommandText & "FROM tblCtrlLimites "
'                .CommandText = .CommandText & "WHERE COD_LIMI='" & strTipLimite & "' AND COD_FOND='00' AND COD_FILE='01' AND FLG_ULTI='X' "
'
'        End Select
'        adoRecCONS.Open .CommandText, adoConn
'
'
'        'LlenarGrid grdListaLimites, adoRecCONS, AGrdCnf(), aDirReg(), frmMainMdi.pgbMain
'
'
'    End With
        
End Sub


Private Sub LlenarCampos(Modo As String)

'    Dim intCont As Integer, intRes As Integer
'
'    Select Case Modo
'        Case "ADICIONAR"
'            strAccion = "ADICIONAR"
'
'            dtpFechaRegistro = gdatFechaActual
'
'        Case "CONSULTAR", "MODIFICAR"
'            If Modo = "CONSULTAR" Then
'                strAccion = "CONSULTAR"
'
'                cboTipoLimite.Enabled = False: cboLimite.Enabled = False
'
'            ElseIf Modo = "MODIFICAR" Then
'                strAccion = "MODIFICAR"
'
'            End If
'
'            If Modo = "CONSULTAR" Then
'            With adoComm
'                Set adoRecCONS = New ADODB.Recordset
'                .CommandType = adCmdText
'
'                .CommandText = "SELECT * FROM tblEventosCorporativos "
'                .CommandText = .CommandText & "WHERE NRO_ACUE=" & CInt(dgdConsulta.Columns(0)) & " AND "
'                .CommandText = .CommandText & "FCH_OPER='" & Convertyyyymmdd(CVDate(dgdConsulta.Columns(1))) & "' AND "
'                .CommandText = .CommandText & "COD_TITU='" & strTipLimite & "' AND FLG_STAT<>'A'"
'
'                Set adoRecCONS = .Execute
'                If Not adoRecCONS.EOF Then
'                    dtpFechaRegistro = Convertddmmyyyy(adoRecCONS("FCH_OPER"))
'
'                End If
'                adoRecCONS.Close: Set adoRecCONS = Nothing
'            End With
'            End If
'        Case Else
'
'    End Select
'    Me.Refresh
    
End Sub


Private Sub cboFondo_Click()

    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
End Sub


Private Sub cboTipoLimite_Click()

    strCodTipoLimite = Valor_Caracter
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    If cboTipoLimite.ListIndex < 0 Then Exit Sub
    
    strCodTipoLimite = Trim(arrTipoLimite(cboTipoLimite.ListIndex))
    
    If strCodTipoLimite = Codigo_Limite_Patrimonio Then
        cboFondo.Enabled = True
    Else
        cboFondo.Enabled = False
    End If
    
    CargarDatos
    
End Sub


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()

    frmMainMdi.stbMdi.Panels(3).Text = Me.Caption
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
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
    
'    With adoComm
'        Set adoRecCONS = New ADODB.Recordset
'
'        '*** Control de Limites ***
'        .CommandText = "SELECT COD_PARA CODIGO,DSC_PARA DESCRIP FROM tblParametros WHERE COD_TIPP='CTRLIM' ORDER BY DSC_PARA"
'        Set adoRecCONS = .Execute
'
'        'LlenaCombo adoRecCONS, cboTipoLimite, aTipLimite(), ""
'        adoRecCONS.Close: Set adoRecCONS = Nothing
'    End With
'
'    '*** Completar la grilla ***
'    ReDim AGrdCnf(1 To 5)
'
'    AGrdCnf(1).TitDes = "Código"
'    AGrdCnf(1).TitJus = vbCenter
'    AGrdCnf(1).DatNom = "COD_LIMI"
'    AGrdCnf(1).DatAnc = 130 * 4
'    AGrdCnf(1).DatJus = vbLeftJustify
'
'    AGrdCnf(2).TitDes = "Descripción"
'    AGrdCnf(2).TitJus = vbCenter
'    AGrdCnf(2).DatNom = "DSC_LIMI"
'    AGrdCnf(2).DatAnc = 130 * 13
'    AGrdCnf(2).DatJus = vbLeftJustify
'
'    AGrdCnf(3).TitDes = "Minimo"
'    AGrdCnf(3).TitJus = vbCenter
'    AGrdCnf(3).DatNom = "MIN_LIMI"
'    AGrdCnf(3).DatAnc = 130 * 10
'    AGrdCnf(3).DatJus = vbRightJustify
'    AGrdCnf(3).DatFmt = "D"
'
'    AGrdCnf(4).TitDes = "Máximo"
'    AGrdCnf(4).TitJus = vbCenter
'    AGrdCnf(4).DatNom = "MAX_LIMI"
'    AGrdCnf(4).DatAnc = 130 * 10
'    AGrdCnf(4).DatJus = vbRightJustify
'    AGrdCnf(4).DatFmt = "D"
'
'    AGrdCnf(5).TitDes = "Modificado"
'    AGrdCnf(5).TitJus = vbCenter
'    AGrdCnf(5).DatNom = "FLG_MODI"
'    AGrdCnf(5).DatAnc = 130 * 3
'    AGrdCnf(5).DatJus = vbCenter
'
'    strAccion = "CONSULTAR"
'    dtpFechaRegistro.Value = gdatFechaActual
    
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
            
End Sub
Public Sub Buscar()

'    Dim strSQL  As String
'
'    strSQL = "SELECT CodAdministradora,CodTitulo,NumAcuerdo,FechaOperacion,CodFile,CodAnalitica,FechaCorte,CodMoneda," & _
'        "CASE PorcenAccionesLiberadas WHEN 0 THEN (CASE PorcenDividendoEfectivo WHEN 0 THEN ValorNominal ELSE PorcenDividendoEfectivo END) " & _
'        "ELSE PorcenAccionesLiberadas END ValorEvento " & _
'        "FROM EventoCorporativoAcuerdo " & _
'        "WHERE CodTitulo='" & strCodTituloCriterio & "' AND TipoAcuerdo='" & strCodEventoCriterio & "' AND " & _
'        "CodAdministradora='" & gstrCodAdministradora & "'"
'
'    strEstado = Reg_Defecto
'    With adoConsulta
'        .ConnectionString = gstrConnectConsulta
'        .RecordSource = strSQL
'        .Refresh
'    End With
'
'    tdgConsulta.Refresh
'
'    If adoConsulta.Recordset.RecordCount > 0 Then strEstado = Reg_Consulta
'
'    Me.MousePointer = vbDefault
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Vista Activa"
    
End Sub
Private Sub CargarListas()
                    
    '*** Control de Limites ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro " & _
        "WHERE CodTipoParametro='CTRLIM' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoLimite, arrTipoLimite(), Sel_Defecto
    
    If cboTipoLimite.ListCount > 0 Then cboTipoLimite.ListIndex = 0
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatos(8) }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Sel_Defecto
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
        
End Sub
Private Sub InicializarValores()
    
    strEstado = Reg_Defecto
    tabLimites.Tab = 0
                    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmLimite = Nothing
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub

Private Function ValidaDatosGenerales() As Boolean

    Dim blnReturn As Boolean
    
    blnReturn = False
            
    
    
    '*** Si todo paso ok ***
    ValidaDatosGenerales = blnReturn
  
End Function
Private Function ValidaAccion() As Boolean

    Dim blnReturn As Boolean
    
    blnReturn = False
    
    If cboTipoLimite.ListIndex < 0 Then
        MsgBox "¡ Seleccione el Título !", vbCritical, Me.Caption
        cboTipoLimite.SetFocus
        blnReturn = True
        
    ElseIf cboFondo.ListIndex < 0 Then
        MsgBox "¡ Seleccione el Evento !", vbCritical, Me.Caption
        cboFondo.SetFocus
        blnReturn = True
      
'    ElseIf cboTipOper.ListIndex < 0 Then
'        MsgBox "¡ Seleccione el Plazo del Concepto del Costo !", vbCritical, Me.Caption
'        cboTipOper.SetFocus
'        blnReturn = True
        
    End If
        
    ValidaAccion = blnReturn

End Function

Private Sub UTlbMant_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.IButtonMenu)

'    On Error GoTo Ctrl_Error
'
'    Dim strTmpSel As String
'
'    If cboTipoLimite.ListIndex < 0 Then
'        MsgBox "¡ Seleccione el Tipo de Limite !", vbCritical, Me.Caption
'        cboTipoLimite.SetFocus
'        Exit Sub
'    End If
'
'    If cboLimite.Visible Then
'        If cboLimite.ListIndex < 0 Then
'            MsgBox "¡ Seleccione el Fondo !", vbCritical, Me.Caption
'            cboLimite.SetFocus
'            Exit Sub
'        End If
'    End If
'
'    gstrNameRepo = "FLLIMITES"
'    gstrFchDel = "": gstrFchAl = ""
'
'    Select Case ButtonMenu.Index
'        Case 1
'            strTmpSel = ""
'            gstrSelFrml = strTmpSel
'        Case 2
'            strTmpSel = "{tblCtrlLimites.FCH_MOVI} IN 'Fch1' TO 'Fch2'"
'            gstrSelFrml = strTmpSel
'            frmRngFecha.Show vbModal
'
'            If gstrSelFrml <> "0" Then
'            Else
'                Exit Sub
'            End If
'    End Select
'
'    Me.Refresh
'    Set gobjReport = New CrystalReport
'    With gobjReport
'        .Connect = gstrConnectReport
'        .ReportFileName = gstrRptPath & gstrNameRepo & ".RPT"
'
'        .Formulas(0) = "User='" & gstrUID & "'"
'        .Formulas(1) = "FchProcIni='" & gstrFchDel & "'"
'        .Formulas(2) = "FchProcFin='" & gstrFchAl & "'"
'        If cboLimite.Visible Then
'            .Formulas(3) = "DscLimite='" & Trim(cboTipoLimite.Text) & "-" & Trim(cboLimite.Text) & "'"
'        Else
'            .Formulas(3) = "DscLimite='" & Trim(cboTipoLimite.Text) & "'"
'        End If
'        .Formulas(4) = "Hora='" & Format(Time, "hh:mm") & "'"
'
'        .StoredProcParam(0) = strTipLimite
'        .StoredProcParam(1) = strTipLimite
'        .StoredProcParam(2) = strLimites
'
'        Select Case Left(ButtonMenu.Key, 3)
'            Case "PRE" '*** Pantalla ***
'                .Destination = crptToWindow
'                .WindowState = crptNormal
'                .WindowState = crptMaximized
'                .WindowTop = 0
'                .WindowLeft = 0
'            Case "IMP" '*** Impresora ***
'                .Destination = crptToPrinter
'        End Select
'
'        .Action = 1
'
'    End With
'    Exit Sub
'
'Ctrl_Error:
'
'    With Err
'        Select Case .Number
'            Case 20504
'                MsgBox "! Reporte NO EXISTE !" & vbNewLine & vbNewLine & _
'                        "CONSULTE con el Administrador del Sistema.", vbCritical, Me.Caption
'            Case Else
'                MsgBox " Error " & Str$(.Number) & .Description, vbCritical, Me.Caption
'        End Select
'    End With
'    Me.MousePointer = vbDefault
'    Resume Next
    
End Sub

Private Sub GrabarLimites()

'    Dim intCont As Integer, intRec As Integer
'
'    With adoComm
'        For intCont = 1 To (grdListaLimites.Rows - 1)
'            grdListaLimites.Row = intCont
'            grdListaLimites.Col = 5
'            If Trim(grdListaLimites.Text) = "X" Then
'                Select Case strTipLimite
'                    Case "01"
'                        grdListaLimites.Col = 3
'                        .CommandText = "UPDATE tblCtrlLimites SET MIN_LIMI=" & CDbl(grdListaLimites.Text) & ","
'                        grdListaLimites.Col = 4
'                        .CommandText = .CommandText & "MAX_LIMI=" & CDbl(grdListaLimites.Text) & ","
'                        .CommandText = .CommandText & "USR_ACTU='" & gstrUID & "',FCH_ACTU='" & strFchOper & "',"
'                        .CommandText = .CommandText & "HOR_ACTU='" & Format(Time, "hh:mm") & "' "
'                        grdListaLimites.Col = 1
'                        .CommandText = .CommandText & "WHERE FCH_CREA='" & strFchOper & "' AND COD_FILE='" & Trim(grdListaLimites.Text) & "' AND "
'                        .CommandText = .CommandText & "COD_LIMI='" & strTipLimite & "' AND COD_EMPR='000000' AND "
'                        .CommandText = .CommandText & "COD_FOND='" & strLimites & "'"
'
'                        gadoConexion.Execute .CommandText, intRec
'
'                        If intRec = 0 Then
'                            .CommandText = "UPDATE tblCtrlLimites SET FLG_ULTI='' "
'                            grdListaLimites.Col = 1
'                            .CommandText = .CommandText & "WHERE COD_FILE='" & Trim(grdListaLimites.Text) & "' AND "
'                            .CommandText = .CommandText & "COD_LIMI='" & strTipLimite & "' AND COD_EMPR='000000' AND "
'                            .CommandText = .CommandText & "COD_FOND='" & strLimites & "' AND FLG_ULTI='X'"
'
'                            gadoConexion.Execute .CommandText
'
'
'                            .CommandText = "INSERT INTO tblCtrlLimites VALUES ('" & strLimites & "','" & strTipLimite & "','"
'                            .CommandText = .CommandText & Trim(grdListaLimites.Text) & "','000000',"
'                            grdListaLimites.Col = 3
'                            .CommandText = .CommandText & CDbl(grdListaLimites.Text) & ","
'                            grdListaLimites.Col = 4
'                            .CommandText = .CommandText & CDbl(grdListaLimites.Text) & ",'" & gstrUID & "','"
'                            .CommandText = .CommandText & strFchOper & "','" & Format(Time, "hh:mm") & "','X','','','')"
'
'                            gadoConexion.Execute .CommandText
'
'                        End If
'                    Case "02"
'                        grdListaLimites.Col = 3
'                        .CommandText = "UPDATE tblCtrlLimites SET MIN_LIMI=" & CDbl(grdListaLimites.Text) & ","
'                        grdListaLimites.Col = 4
'                        .CommandText = .CommandText & "MAX_LIMI=" & CDbl(grdListaLimites.Text) & ","
'                        .CommandText = .CommandText & "USR_ACTU='" & gstrUID & "',FCH_ACTU='" & strFchOper & "',HOR_ACTU='" & Format(Time, "hh:mm") & "' "
'                        grdListaLimites.Col = 1
'                        .CommandText = .CommandText & "WHERE FCH_CREA='" & strFchOper & "' AND COD_FILE='00' AND "
'                        .CommandText = .CommandText & "COD_LIMI='" & strTipLimite & "' AND COD_EMPR='" & Trim(grdListaLimites.Text) & "' AND "
'                        .CommandText = .CommandText & "COD_FOND='00'"
'
'                        gadoConexion.Execute .CommandText, intRec
'
'                        If intRec = 0 Then
'                            .CommandText = "UPDATE tblCtrlLimites SET FLG_ULTI='' "
'                            grdListaLimites.Col = 1
'                            .CommandText = .CommandText & "WHERE COD_FILE='00' AND "
'                            .CommandText = .CommandText & "COD_LIMI='" & strTipLimite & "' AND COD_EMPR='" & Trim(grdListaLimites.Text) & "' AND "
'                            .CommandText = .CommandText & "COD_FOND='00' AND FLG_ULTI='X'"
'
'                            gadoConexion.Execute .CommandText
'
'
'                            .CommandText = "INSERT INTO tblCtrlLimites VALUES ('00','" & strTipLimite & "','"
'                            .CommandText = .CommandText & "00','" & Trim(grdListaLimites.Text) & "',"
'                            grdListaLimites.Col = 3
'                            .CommandText = .CommandText & CDbl(grdListaLimites.Text) & ","
'                            grdListaLimites.Col = 4
'                            .CommandText = .CommandText & CDbl(grdListaLimites.Text) & ",'" & gstrUID & "','"
'                            .CommandText = .CommandText & strFchOper & "','" & Format(Time, "hh:mm") & "','X','','','')"
'
'                            gadoConexion.Execute .CommandText
'
'                        End If
'
'                    Case "03"
'                        grdListaLimites.Col = 3
'                        .CommandText = "UPDATE tblCtrlLimites SET MIN_LIMI=" & CDbl(grdListaLimites.Text) & ","
'                        grdListaLimites.Col = 4
'                        .CommandText = .CommandText & "MAX_LIMI=" & CDbl(grdListaLimites.Text) & ","
'                        .CommandText = .CommandText & "USR_ACTU='" & gstrUID & "',FCH_ACTU='" & strFchOper & "',HOR_ACTU='" & Format(Time, "hh:mm") & "' "
'                        grdListaLimites.Col = 1
'                        .CommandText = .CommandText & "WHERE FCH_CREA='" & strFchOper & "' AND COD_FILE='" & Trim(grdListaLimites.Text) & "' AND "
'                        .CommandText = .CommandText & "COD_LIMI='" & strTipLimite & "' AND COD_EMPR='000000' AND "
'                        .CommandText = .CommandText & "COD_FOND='00'"
'
'                        gadoConexion.Execute .CommandText, intRec
'
'                        If intRec = 0 Then
'                            .CommandText = "UPDATE tblCtrlLimites SET FLG_ULTI='' "
'                            grdListaLimites.Col = 1
'                            .CommandText = .CommandText & "WHERE COD_FILE='" & Trim(grdListaLimites.Text) & "' AND "
'                            .CommandText = .CommandText & "COD_LIMI='" & strTipLimite & "' AND COD_EMPR='000000' AND "
'                            .CommandText = .CommandText & "COD_FOND='00' AND FLG_ULTI='X'"
'
'                            gadoConexion.Execute .CommandText
'
'
'                            .CommandText = "INSERT INTO tblCtrlLimites VALUES ('00','" & strTipLimite & "','"
'                            .CommandText = .CommandText & Trim(grdListaLimites.Text) & "','000000',"
'                            grdListaLimites.Col = 3
'                            .CommandText = .CommandText & CDbl(grdListaLimites.Text) & ","
'                            grdListaLimites.Col = 4
'                            .CommandText = .CommandText & CDbl(grdListaLimites.Text) & ",'" & gstrUID & "','"
'                            .CommandText = .CommandText & strFchOper & "','" & Format(Time, "hh:mm") & "','X','','','')"
'
'                            gadoConexion.Execute .CommandText
'
'                        End If
'
'
'                    Case "04"
'                        grdListaLimites.Col = 3
'                        .CommandText = "UPDATE tblCtrlLimites SET MIN_LIMI=" & CDbl(grdListaLimites.Text) & ","
'                        grdListaLimites.Col = 4
'                        .CommandText = .CommandText & "MAX_LIMI=" & CDbl(grdListaLimites.Text) & ","
'                        .CommandText = .CommandText & "USR_ACTU='" & gstrUID & "',FCH_ACTU='" & strFchOper & "',HOR_ACTU='" & Format(Time, "hh:mm") & "' "
'                        grdListaLimites.Col = 1
'                        .CommandText = .CommandText & "WHERE FCH_CREA='" & strFchOper & "' AND COD_FILE='" & Trim(grdListaLimites.Text) & "' AND "
'                        .CommandText = .CommandText & "COD_LIMI='" & strTipLimite & "' AND COD_EMPR='000000' AND "
'                        .CommandText = .CommandText & "COD_FOND='00'"
'
'                        gadoConexion.Execute .CommandText, intRec
'
'                        If intRec = 0 Then
'                            .CommandText = "UPDATE tblCtrlLimites SET FLG_ULTI='' "
'                            grdListaLimites.Col = 1
'                            .CommandText = .CommandText & "WHERE COD_FILE='" & Trim(grdListaLimites.Text) & "' AND "
'                            .CommandText = .CommandText & "COD_LIMI='" & strTipLimite & "' AND COD_EMPR='000000' AND "
'                            .CommandText = .CommandText & "COD_FOND='00' AND FLG_ULTI='X'"
'
'                            gadoConexion.Execute .CommandText
'
'
'                            .CommandText = "INSERT INTO tblCtrlLimites VALUES ('00','" & strTipLimite & "','"
'                            .CommandText = .CommandText & Trim(grdListaLimites.Text) & "','000000',"
'                            grdListaLimites.Col = 3
'                            .CommandText = .CommandText & CDbl(grdListaLimites.Text) & ","
'                            grdListaLimites.Col = 4
'                            .CommandText = .CommandText & CDbl(grdListaLimites.Text) & ",'" & gstrUID & "','"
'                            .CommandText = .CommandText & strFchOper & "','" & Format(Time, "hh:mm") & "','X','','','')"
'
'                            gadoConexion.Execute .CommandText
'
'                        End If
'
'                    Case "05"
'                        grdListaLimites.Col = 3
'                        .CommandText = "UPDATE tblCtrlLimites SET MIN_LIMI=" & CDbl(grdListaLimites.Text) & ","
'                        grdListaLimites.Col = 4
'                        .CommandText = .CommandText & "MAX_LIMI=" & CDbl(grdListaLimites.Text) & ","
'                        .CommandText = .CommandText & "USR_ACTU='" & gstrUID & "',FCH_ACTU='" & strFchOper & "',"
'                        .CommandText = .CommandText & "HOR_ACTU='" & Format(Time, "hh:mm") & "' "
'                        grdListaLimites.Col = 1
'                        .CommandText = .CommandText & "WHERE FCH_CREA='" & strFchOper & "' AND COD_FILE='" & Trim(grdListaLimites.Text) & "' AND "
'                        .CommandText = .CommandText & "COD_LIMI='" & strTipLimite & "' AND COD_EMPR='000000' AND "
'                        .CommandText = .CommandText & "COD_FOND='00'"
'
'                        gadoConexion.Execute .CommandText, intRec
'
'                        If intRec = 0 Then
'                            .CommandText = "UPDATE tblCtrlLimites SET FLG_ULTI='' "
'                            grdListaLimites.Col = 1
'                            .CommandText = .CommandText & "WHERE COD_FILE='" & Trim(grdListaLimites.Text) & "' AND "
'                            .CommandText = .CommandText & "COD_LIMI='" & strTipLimite & "' AND COD_EMPR='000000' AND "
'                            .CommandText = .CommandText & "COD_FOND='00' AND FLG_ULTI='X'"
'
'                            gadoConexion.Execute .CommandText
'
'
'                            .CommandText = "INSERT INTO tblCtrlLimites VALUES ('00','" & strTipLimite & "','"
'                            .CommandText = .CommandText & Trim(grdListaLimites.Text) & "','000000',"
'                            grdListaLimites.Col = 3
'                            .CommandText = .CommandText & CDbl(grdListaLimites.Text) & ","
'                            grdListaLimites.Col = 4
'                            .CommandText = .CommandText & CDbl(grdListaLimites.Text) & ",'" & gstrUID & "','"
'                            .CommandText = .CommandText & strFchOper & "','" & Format(Time, "hh:mm") & "','X','','','')"
'
'                            gadoConexion.Execute .CommandText
'
'                        End If
'                End Select
'            End If
'        Next
'
'    End With
    
End Sub
