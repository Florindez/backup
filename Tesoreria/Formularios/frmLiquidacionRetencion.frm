VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmLiquidacionRetencion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operaciones con Retención"
   ClientHeight    =   5280
   ClientLeft      =   525
   ClientTop       =   1875
   ClientWidth     =   8355
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5280
   ScaleWidth      =   8355
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   6600
      TabIndex        =   13
      Top             =   4440
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   4800
      Top             =   4560
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
      Caption         =   "adoConsulta"
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
   Begin TabDlg.SSTab tabRetencion 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Operaciones con Retención"
      TabPicture(0)   =   "frmLiquidacionRetencion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblValorCuota"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDescrip(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDescrip(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDescrip(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDescrip(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblDescrip(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dtpFechaLiquidacion"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboFondo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmd_Acc"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtTipoCambio"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkConfirmados"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dgdConsulta"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmLiquidacionRetencion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDetalle"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraDetalle 
         Height          =   3135
         Left            =   -74640
         TabIndex        =   14
         Top             =   720
         Width           =   7335
         Begin VB.CommandButton cmd_AccIte 
            Caption         =   "&Actualizar"
            Height          =   735
            Index           =   0
            Left            =   1320
            Picture         =   "frmLiquidacionRetencion.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox txtNumCheque 
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
            TabIndex        =   18
            Top             =   1425
            Width           =   1695
         End
         Begin VB.CommandButton cmd_AccIte 
            Cancel          =   -1  'True
            Caption         =   "A&nular"
            Height          =   735
            Index           =   1
            Left            =   3000
            Picture         =   "frmLiquidacionRetencion.frx":0574
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CommandButton cmd_AccIte 
            Caption         =   "&Cancelar"
            Default         =   -1  'True
            Height          =   735
            Index           =   2
            Left            =   4680
            Picture         =   "frmLiquidacionRetencion.frx":064B
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   2280
            Width           =   1200
         End
         Begin VB.ComboBox cboBanco 
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
            TabIndex        =   15
            Top             =   1050
            Width           =   3855
         End
         Begin MSComCtl2.DTPicker dtpFechaRetencion 
            Height          =   315
            Left            =   2430
            TabIndex        =   20
            Top             =   1770
            Width           =   1305
            _ExtentX        =   2302
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
            Format          =   175964161
            CurrentDate     =   38068
         End
         Begin VB.Label lblNumFolio 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2430
            TabIndex        =   27
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblDescripParticipe 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2430
            TabIndex        =   26
            Top             =   705
            Width           =   3855
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Num. Folio"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   6
            Left            =   960
            TabIndex        =   25
            Top             =   375
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Partícipe"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   7
            Left            =   960
            TabIndex        =   24
            Top             =   735
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Banco"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   8
            Left            =   960
            TabIndex        =   23
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Num. Cheque"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   9
            Left            =   960
            TabIndex        =   22
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Retención"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   10
            Left            =   960
            TabIndex        =   21
            Top             =   1785
            Width           =   975
         End
      End
      Begin MSDataGridLib.DataGrid dgdConsulta 
         Bindings        =   "frmLiquidacionRetencion.frx":0BAD
         Height          =   1935
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3413
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin VB.CheckBox chkConfirmados 
         Caption         =   "Incluir confirmados"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtTipoCambio 
         Alignment       =   1  'Right Justify
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
         Left            =   6480
         TabIndex        =   10
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Acc 
         Caption         =   "&Liquidar"
         Height          =   735
         Left            =   6480
         Picture         =   "frmLiquidacionRetencion.frx":0BC7
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   1200
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   3945
      End
      Begin MSComCtl2.DTPicker dtpFechaLiquidacion 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   855
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
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
         Format          =   175964161
         CurrentDate     =   38068
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Tipo de Cambio"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   5
         Left            =   4920
         TabIndex        =   8
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Valor Cuota"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   4
         Left            =   2520
         TabIndex        =   7
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Moneda"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   3480
         TabIndex        =   6
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fecha"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   875
         Width           =   855
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fondo"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblValorCuota 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0000"
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
         Left            =   3660
         TabIndex        =   3
         Top             =   1560
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmLiquidacionRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim aMapFonMut()
Dim aMapCodBan()
Dim adoOper      As New Recordset

Dim adirreg()    As String
Dim aBookmark()  As Variant
Dim strCodFon    As String
Dim strCodMon    As String
Dim dblValcuo    As Double
Dim strCodProTit As String
Dim strFecDia    As String
Dim strNewNroCom As String
Dim strNewNroCrt As String
Dim strEstado As String

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vQuery
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

Public Sub Abrir()

End Sub


Public Sub Adicionar()

End Sub

Public Sub Anterior()

End Sub

Public Sub Ayuda()

End Sub

Public Sub Buscar()

    MousePointer = vbHourglass
    LDoGrid
    MousePointer = vbDefault
            
End Sub

Public Sub Cancelar()

End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("REPORTES").ButtonMenus("REPO1").Visible = True
    frmMainMdi.tlbMdi.Buttons("REPORTES").ButtonMenus("REPO1").Text = "Vista Activa"
    frmMainMdi.tlbMdi.Buttons("REPORTES").ButtonMenus("REPO2").Visible = True
    frmMainMdi.tlbMdi.Buttons("REPORTES").ButtonMenus("REPO2").Text = "Cheques Confirmados"
    frmMainMdi.tlbMdi.Buttons("REPORTES").ButtonMenus("REPO3").Visible = True
    frmMainMdi.tlbMdi.Buttons("REPORTES").ButtonMenus("REPO3").Text = "Cheques con Retención"
    
End Sub

Private Sub Deshabilita()

End Sub

Public Sub Eliminar()

End Sub

Public Sub Exportar()

End Sub

Public Sub Grabar()

End Sub

Public Sub Importar()

End Sub

Public Sub Imprimir()

End Sub

Public Sub Modificar()

End Sub

Public Sub OrdenarAZ()

End Sub

Public Sub OrdenarZA()

End Sub

Public Sub Primero()

End Sub

Public Sub Refrescar()

End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub Seguridad()

End Sub

Public Sub Siguiente()

End Sub

Public Sub SubImprimir(index As Integer)

'    Dim strTmpSel As String
'
'    gstrNameRepo = "FLCHQLIB2"
'
'    Dim frmRpt As frmReportViewer
'    Set frmRpt = New frmReportViewer
'
'    Dim aReportParamS(), aReportParamF(), aReportParamFn()
'
'    ReDim aReportParamS(1)
'    ReDim aReportParamFn(1)
'    ReDim aReportParamF(1)
'
'    aReportParamFn(0) = "User"
'    aReportParamFn(1) = "Hora"
'
'    aReportParamF(0) = gstrLogin
'    aReportParamF(1) = Format(Time, "hh:mm:ss")
'
'    If Index = 0 Then
'        aReportParamS(0) = "L"
'        aReportParamS(1) = gstrFechaAct
'    Else
'        aReportParamS(0) = "R"
'        aReportParamS(1) = gstrFechaAct
'    End If
'
'    frmRpt.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"
'
'    Call frmRpt.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())
'
'    frmRpt.Caption = "Reporte - (" & gstrNameRepo & ")"
'    frmRpt.Show vbModal
'
'    Set frmRpt = Nothing
'
'    Screen.MousePointer = vbNormal
    
End Sub

Public Sub Ultimo()

End Sub


Private Sub cmb_FonMut_Click()

'    Dim adoRecord As ADODB.Recordset
'    Dim intRes As Integer
'    Dim strSQL As String
'
'    If cmb_FonMut.ListIndex >= 0 Then
'        strCodFon = Mid(aMapFonMut(cmb_FonMut.ListIndex), 1, 2)
'        strCodMon = Mid(aMapFonMut(cmb_FonMut.ListIndex), 3, 1)
'        Rea_TipCam.Text = IIf(strCodMon = "S", 0, gdblTipoCamb)
'        lblDescrip(3).Visible = IIf(strCodMon = "S", False, True)
'        Rea_TipCam.Visible = IIf(strCodMon = "S", False, True)
'        lblDescrip(3).Caption = IIf(strCodMon = "D", "US$ Dólares", "Soles")
'
'        Set adoRecord = New ADODB.Recordset
'        '*** Fecha disponible para el fondo ***
'        strSQL = "SELECT FMCUOTAS.VAL_CUOT,FCH_CUOT, TIP_VALU, VAL_CUO2 FROM FMCUOTAS, FMFONDOS"
'        strSQL = strSQL & " WHERE FMCUOTAS.COD_FOND='" & strCodFon & "'"
'        strSQL = strSQL & " AND FLG_ABIE='X'"
'        strSQL = strSQL & " AND FMFONDOS.COD_FOND = FMCUOTAS.COD_FOND"
'        strSQL = strSQL & " ORDER BY FCH_CUOT"
'        adoComm.CommandText = strSQL
'        Set adoRecord = adoComm.Execute
'
'        If Not adoRecord.EOF Then
'            lbl_ValCuo.Caption = Format(adoRecord("Val_cuot"), "0.0000")
'            dblValcuo = Format(adoRecord("Val_cuot"), "0.00000000")
'
'            Dat_FchLiq.Value = FmtFec(adoRecord("fch_cuot"), "yyyymmdd", "win", intRes)
'            strFecDia = adoRecord("fch_cuot")
'            If Not LEsDiaUtil(Dat_FchLiq.Value) Then
'                Dat_FchLiq.Value = LProxDiaUtil(Dat_FchLiq.Value)
'            End If
'            'gstrFechaAct = dat_FchLiq.Text
'            Dat_FecDev.Value = gstrFechaAct
'            Dat_FecDev.Value = DateAdd("d", gintDiasPDev, Dat_FecDev.Value)
'            If Not LEsDiaUtil(Dat_FecDev.Value) Then
'                Dat_FecDev.Value = LProxDiaUtil(Dat_FecDev.Value)
'            End If
'        Else
'            MsgBox "No hay fecha disponible para el Fondo " & cmb_FonMut.Text
'            lbl_ValCuo.Caption = Format(0, "0.0000")
'        End If
'        adoRecord.Close: Set adoRecord = Nothing
'    End If
   
End Sub

Private Sub cmd_Acc_Click()
   
'    'Liquidar
'    If MsgBox("Desea Procesar las Operaciones con retención pendientes y cuya fecha de liberación sea " & Dat_FchLiq.Value, 36) <> 6 Then
'        Exit Sub
'    End If
'    If strCodMon = "D" And CDbl(Rea_TipCam.Text) <= 0 Then
'        MsgBox "El Tipo de cambio debe ser mayor a cero!", vbExclamation
'        Exit Sub
'    End If
'
'    'Iniciar Proceso
'    LConfRete
'    MsgBox "Fin de Proceso", vbInformation
    
End Sub

Private Sub cmd_AccIte_Click(index As Integer)

'    Dim strFchRet As String
'    Dim intRes    As Integer
'
'    Select Case Index
'        Case 0 'Actualiza cambios
'                MousePointer = 11
'                '*** Aqui Actualizar FMSOLICI, FMOPERAC, FMMOVCTA ***
'                strFchRet = FmtFec(dat_FchRet.Value, "win", "yyyymmdd", intRes)
'                adoComm.CommandText = "UPDATE FMSOLICI SET FCH_FRET = '" + strFchRet + "', "
'                adoComm.CommandText = adoComm.CommandText + "COD_BANC = '" & aMapCodBan(cmb_CodBan.ListIndex) & "', "
'                adoComm.CommandText = adoComm.CommandText + "NRO_CHEQ = '" + txt_NroChe.Text + "'"
'                adoComm.CommandText = adoComm.CommandText + " WHERE COD_FOND = '" & adoOper!COD_FOND & "'"
'                adoComm.CommandText = adoComm.CommandText + " AND NRO_FOLI = '" & adoOper!nro_foli & "'"
'                adoConn.Execute adoComm.CommandText
'
'                adoComm.CommandText = "UPDATE FMOPERAC SET FCH_FRET = '" + strFchRet + "', "
'                adoComm.CommandText = adoComm.CommandText + "COD_BANC = '" & aMapCodBan(cmb_CodBan.ListIndex) & "', "
'                adoComm.CommandText = adoComm.CommandText + "NRO_CHEQ = '" + txt_NroChe.Text + "'"
'                adoComm.CommandText = adoComm.CommandText + " WHERE COD_FOND = '" & adoOper!COD_FOND & "'"
'                adoComm.CommandText = adoComm.CommandText + " AND NRO_FOLI = '" & adoOper!nro_foli & "'"
'                adoConn.Execute adoComm.CommandText
'
'                adoComm.CommandText = "UPDATE FMMOVCTA SET FCH_OBLI = '" + strFchRet + "', "
'                adoComm.CommandText = adoComm.CommandText + "COD_BAND = '" & aMapCodBan(cmb_CodBan.ListIndex) & "', "
'                adoComm.CommandText = adoComm.CommandText + "NRO_CHED = '" + txt_NroChe.Text + "' "
'                adoComm.CommandText = adoComm.CommandText + "WHERE COD_FOND = '" & adoOper!COD_FOND & "' "
'                adoComm.CommandText = adoComm.CommandText + "AND NRO_FOLI = '" & adoOper!nro_foli & "'"
'                adoConn.Execute adoComm.CommandText
'                LDoGrid
'                MousePointer = 0
'
'        Case 1 'Anular Operación
'                If MsgBox("Esta Ud. seguro de anular esta operación, si elige anular la operación no podrá recuperarla, verifique su información antes de aceptar", vbInformation + vbYesNo, gstrNombreEmpresa) = vbYes Then
'                    LAnuPagRet
'                End If
'
'        Case 2 'Salir
'    End Select
'    Fra_DesCam(0).Visible = False
'    Fra_DesCam(0).Enabled = False
    
End Sub

Private Sub cmd_salir_Click()
   
   Unload Me

End Sub

Private Sub dat_FchLiq_LostFocus()
   
   'If Not LEsDiaUtil(dat_FchLiq) Then
   '  dat_FchLiq.Text = LProxDiaUtil(dat_FchLiq)
   'End If
  
End Sub

Private Sub dat_FchRet_LostFocus()
   
   'If Not LEsDiaUtil(dat_FchRet) Then
   '   dat_FchRet.Text = LProxDiaUtil(dat_FchRet)
   'End If

End Sub

Private Sub Dat_FecDev_LostFocus()

'    If Not LEsDiaUtil(Dat_FecDev.Value) Then
'        Dat_FecDev.Value = LProxDiaUtil(Dat_FecDev.Value)
'    End If
    
End Sub

Private Sub Form_Load()

'    Dim strSentencia As String
'
'    adoComm.CommandTimeout = 7200
'
'    '*** Fondos ***
'    strSentencia = "SELECT (COD_FOND + COD_MONE) CODIGO, DSC_FOND DESCRIP FROM FMFONDOS ORDER BY DSC_FOND"
'    Call LCmbLoad(strSentencia, cmb_FonMut, aMapFonMut(), "")
'    If cmb_FonMut.ListCount > 0 Then cmb_FonMut.ListIndex = 0
'
'    '*** Bancos ***
'    strSentencia = "SELECT COD_PERS CODIGO, DSC_PERS DESCRIP FROM FMPERSON WHERE TIP_PERS='EM' AND COD_SECT='02' AND FLG_VIGE='X' ORDER BY DSC_PERS"
'    Call LCmbLoad(strSentencia, cmb_CodBan, aMapCodBan(), "")
'
'    'Configuración de Grilla
'    ReDim aGrdCnf(1 To 9)
'    aGrdCnf(1).TitDes = "Nro.Oper.."
'    aGrdCnf(1).DatNom = "NRO_OPER"
'    aGrdCnf(1).DatAnc = 130 * 8
'
'    aGrdCnf(2).TitDes = "Tip.Pago."
'    aGrdCnf(2).DatNom = "TIP_PAGO"
'    aGrdCnf(2).DatAnc = 130 * 4
'
'    aGrdCnf(3).TitDes = "Partícipe"
'    aGrdCnf(3).DatNom = "DSC_PART"
'    aGrdCnf(3).DatAnc = 130 * 30
'
'    aGrdCnf(4).TitDes = "Banco"
'    aGrdCnf(4).DatNom = "DSC_PERS"
'    aGrdCnf(4).DatAnc = 130 * 20
'
'    aGrdCnf(5).TitDes = "Nro.Cta."
'    aGrdCnf(5).DatNom = "NRO_CTA"
'    aGrdCnf(5).DatAnc = 130 * 10
'    aGrdCnf(5).DatJus = 1
'
'    aGrdCnf(6).TitDes = "Nro.Cheque"
'    aGrdCnf(6).DatNom = "NRO_CHEQ"
'    aGrdCnf(6).DatAnc = 130 * 10
'    aGrdCnf(6).DatJus = 1
'
'    aGrdCnf(7).TitDes = "Monto"
'    aGrdCnf(7).DatNom = "VAL_MOVI"
'    aGrdCnf(7).DatAnc = 130 * 20
'    aGrdCnf(7).DatFmt = "D"
'    aGrdCnf(7).DatJus = 1
'
'    aGrdCnf(8).TitDes = "Fec.Oper"
'    aGrdCnf(8).DatNom = "FCH_OPER"
'    aGrdCnf(8).DatAnc = 130 * 10
'    aGrdCnf(8).DatFmt = "F"
'
'    aGrdCnf(9).TitDes = "Hor.Oper"
'    aGrdCnf(9).DatNom = "HOR_OPER"
'    aGrdCnf(9).DatAnc = 130 * 10
'
'    strEstado = ""
'    tabRetencion.Tab = 0
    Call Buscar
    CentrarForm Me
    Call CargarReportes
    
    Set cmdSalir.FormularioActivo = Me
        
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
End Sub
   
Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmLiquidacionRetencion = Nothing
    
End Sub


Private Sub grd_OpeRet_Click()
  
'  If grd_OpeRet.Rows = 1 Then Exit Sub
'  grd_OpeRet.SelStartRow = grd_OpeRet.Row: grd_OpeRet.SelStartCol = 1
'  grd_OpeRet.SelEndRow = grd_OpeRet.Row: grd_OpeRet.SelEndCol = grd_OpeRet.Cols - 1

End Sub

Private Sub grd_OpeRet_DblClick()

'    Dim vntIte As Variant
'    Dim intRes As Integer
'
'    If grd_OpeRet.Row = 0 Then
'        Exit Sub
'    End If
'    grd_OpeRet.Col = 1
'    If Trim(grd_OpeRet.Text) = "" Then
'        Exit Sub
'    End If
'
'    adoOper.Bookmark = aBookmark(grd_OpeRet.Row)
'    Fra_DesCam(0).Enabled = True
'    Fra_DesCam(0).Visible = True
'
'    '*** Setear datos ***
'    dat_FchRet.Value = FmtFec(adoOper!FCH_FRET, "yyyymmdd", "win", intRes)
'    lbl_NroFoli.Caption = adoOper!nro_foli
'    lbl_Dscpart.Caption = adoOper!DSC_PART
'    vntIte = adoOper!COD_BAND
'    cmb_CodBan.ListIndex = LBsqIteArr(aMapCodBan(), vntIte)
'    txt_NroChe.Text = adoOper!nro_cheq
'    cmd_AccIte(1).Enabled = IIf(adoOper!FLG_CONF = "X", False, True)
'    tabRetencion.Tab = 1

End Sub

Private Sub LAnuPagRet()

'Dim adoRecord As New Recordset
'Dim CabAsi    As RCabasicon
'Dim intRes    As Integer
'Dim intIdx    As Integer
'
'
'    Static DetAsi(1 To 2) As RDetasicon
'    adoOper.Bookmark = aBookmark(grd_OpeRet.Row)
'    DoEvents
'    strNewNroCom = Format$(NewNro(strCodFon, "COM"), "00000000")
'    gblnRollBack = False
'    On Error GoTo tag_Error:
'    adoConn.Execute "BEGIN TRANSACTION OpeRet"
'
'   '*** Asiento contable
'   '*** Cabecera del Comprobante
'   adoComm.CommandText = "SELECT * FROM FMCOMPRO WHERE COD_FOND = '" & adoOper!COD_FOND & "' AND NRO_COMP = '" & adoOper!NRO_COMP & "'" ' AND NRO_OPER = '" & adoOper!NRO_OPER & "'"
'   Set adoRecord = adoComm.Execute
'
'   '*** inicializa ***
'   LIniRCabAsiCon CabAsi
'
'   CabAsi.NRO_COMP = strNewNroCom
'   CabAsi.CNT_MOVI = 2
'   CabAsi.COD_FOND = adoRecord!COD_FOND
'   CabAsi.COD_MONC = adoRecord!COD_MONC
'   CabAsi.COD_MONE = adoRecord!COD_MONE
'   CabAsi.DSL_COMP = "Rechazo de Suscripción con Retención"
'   CabAsi.FCH_COMP = gstrFechaAct
'   CabAsi.FCH_CONT = CabAsi.FCH_COMP
'   CabAsi.FLG_AUTO = adoRecord!FLG_AUTO
'   CabAsi.FLG_CONT = adoRecord!FLG_CONT
'   CabAsi.GLO_COMP = "Rechazo de Suscripción con Retención"
'   CabAsi.HOR_COMP = Time
'   CabAsi.NRO_OPER = adoRecord!NRO_OPER
'   CabAsi.prd_cont = Mid(CabAsi.FCH_COMP, 1, 4)
'   CabAsi.MES_CONT = Mid(CabAsi.FCH_COMP, 5, 2)
'   CabAsi.SUB_SIST = adoRecord!SUB_SIST
'   CabAsi.VAL_COMP = adoRecord!VAL_COMP
'
'   '*** Detalle del Comprobante
'   adoComm.CommandText = "SELECT * FROM FMMOVCON WHERE COD_FOND = '" & adoOper!COD_FOND & "' AND NRO_COMP = '" & adoOper!NRO_COMP & "' AND FCH_MOVI = '" & adoRecord!FCH_COMP & "'"
'   Set adoRecord = adoComm.Execute
'   intIdx = 0
'
'   Do While Not adoRecord.EOF
'      intIdx = intIdx + 1
'      LIniRDetAsiCon DetAsi(intIdx)  'Inicializar
'      DetAsi(intIdx).COD_FILE = adoRecord!COD_FILE
'      DetAsi(intIdx).COD_ANAL = adoRecord!COD_ANAL
'      DetAsi(intIdx).COD_FOND = adoRecord!COD_FOND
'      DetAsi(intIdx).COD_MONE = adoRecord!COD_MONE
'      DetAsi(intIdx).FCH_MOVI = CabAsi.FCH_COMP
'      DetAsi(intIdx).FLG_PROC = "X"
'      DetAsi(intIdx).MES_COMP = Mid(DetAsi(intIdx).FCH_MOVI, 5, 2)
'      DetAsi(intIdx).NRO_COMP = strNewNroCom
'      DetAsi(intIdx).prd_cont = Mid(DetAsi(intIdx).FCH_MOVI, 1, 4)
'      DetAsi(intIdx).SEC_MOVI = Format$(intIdx, "000")
'      DetAsi(intIdx).TIP_GENR = adoRecord!TIP_GENR
'      DetAsi(intIdx).DSC_MOVI = adoRecord!DSC_MOVI
'      DetAsi(intIdx).COD_CTA = adoRecord!COD_CTA
'      DetAsi(intIdx).FLG_DEHA = IIf(adoRecord!FLG_DEHA = "D", "H", "D")
'      DetAsi(intIdx).VAL_CONT = Format(adoRecord!VAL_CONT * -1, "0.00")
'      DetAsi(intIdx).VAL_MOVN = Format(adoRecord!VAL_MOVN * -1, "0.00")
'      DetAsi(intIdx).VAL_MOVX = Format(adoRecord!VAL_MOVX * -1, "0.00")
'      adoRecord.MoveNext
'   Loop
'   LGraAsiCont CabAsi, DetAsi() 'Grabar el asiento
'
'   '** Actualiza Nuevo Nro de Comprobante Contable
'   Call UpdNewNro(strCodFon, "COM", strNewNroCom)
'   '** Anula Operación
'   adoComm.CommandText = "UPDATE FMOPERAC SET FLG_CONT = 'A', FLG_CONF = 'A',FLG_RETE = 'A' WHERE COD_FOND = '" & adoOper!COD_FOND & "' AND NRO_OPER = '" & adoOper!NRO_OPER & "'"
'   adoConn.Execute adoComm.CommandText
'
'   '** Anula Movim. Caja
'   adoComm.CommandText = "UPDATE FMMOVCTA SET FLG_NVIG = 'X', FLG_CONT = 'A', FLG_CONF = 'X' WHERE COD_FOND = '" & adoOper!COD_FOND & "' AND SUB_SIST = 'P' AND NRO_MCTA = '" & adoOper!NRO_MCTA & "'"
'   adoConn.Execute adoComm.CommandText
'   If gblnRollBack Then
'      adoConn.Execute "ROLLBACK TRAN OpeRet"
'      MsgBox "Transacción Cancelada, verifique información e intente nuevamente la operación."
'   Else
'      adoConn.Execute "COMMIT TRAN OpeRet"
'      MsgBox "Transacción Culminada Exitosamente."
'      LDoGrid
'   End If
'
'   Exit Sub
'tag_Error:
'   gblnRollBack = True
'   Resume Next
End Sub


Private Sub LConfRete()

'Dim adoMovCta  As New Recordset
'Dim adoRecord  As New Recordset
'Dim adoRecord1 As New Recordset
'Dim strCodCtaN As String * 8
'Dim strCodCtaX As String * 8
'Dim WComCon    As RCabasicon        'Comprobante contable
'Dim WMovCon()  As RDetasicon        'Movimientos del comprobante contable
'Dim strCodFil  As String * 2
'Dim strCodAna  As String * 6
'Dim strNroCom  As String * 8
'Dim strNroTmp  As String * 8
'Dim strFchOpe  As String * 8
'Dim strHorOpe  As String
'Dim strPrdCon  As String * 4
'Dim strMesCon  As String * 2
'
'Dim strCtaDif  As String * 8
'Dim strCtaApo  As String * 8
'Dim strCtaSus  As String * 8
'Dim strCtaBpa  As String * 8
'Dim strCtaSpa  As String * 8
'Dim strCtaIgv  As String * 8
'Dim strCtaSuD  As String * 8
'Dim strTipCom  As String
'Dim strTipDoc  As String
'Dim strNroDoc  As String
'Dim strDslCom  As String
'Dim strGloCom  As String
'Dim strStaCom  As String
'Dim strCodCta  As String * 8
'Dim strDscMov  As String
'Dim strDscOrd  As String
'Dim strTipMov  As String
'Dim strSecMov  As String * 3
'
'Dim dblCntCuo    As Double
'Dim curValTot    As Currency
'Dim curMtoNom    As Currency
'Dim curMtoCom    As Currency
'Dim curMtoDif    As Currency
'Dim n_MtoIgv     As Currency
'Dim curValAju    As Double
'Dim curValNomCuo As Currency
'Dim curValSPa    As Currency
'Dim curValBPa    As Currency
'Dim curValMov    As Currency
'Dim aFlgDeha()   As String * 1
'Dim aCodCtas()   As String * 8
'Dim aValMovi()   As Currency
'Dim aDscMovi()   As String
'Dim aCodMone()   As String * 1
'Dim curTasOper   As Currency
'Dim strFlgExtr   As String * 1
'Dim strFlgCust   As String * 1
'Dim strClsPers   As String * 1
'Dim strCodProm   As String * 6
'Dim strCodPart   As String * 15
'Dim strNroOper   As String * 8
'Dim strNroFoli   As String * 11
'Dim strTipOper   As String * 2
'Dim curValMovi   As Currency
'Dim dblSinCuo    As Double
'
''Retención
'Dim strCodCtaAct As String * 8
'Dim strCodFilAct As String * 2
'Dim strCodAnaAct As String * 6
'Dim strCodCta16N As String * 8
'Dim strCodCta16X As String * 8
'Dim WCertif      As RCertif        'Certificados de Participación
'
''Para el Vuelto
'Dim strTipPagd As String * 1
'Dim strCodBand As String
'Dim strTipCtad As String
'Dim strNroCtad As String
'Dim strCodAged As String * 6
'Dim strCodSucd As String * 3
'
''Control de Transacciones
'Dim intTotOK   As Integer
'Dim intTotER   As Integer
'Dim intTotTr   As Integer
'Dim strTipValu As String * 1
'Dim strMensaje As String
'Dim dblTipCam  As Double
'Dim strFchRet  As String
'Dim intRes     As Integer
'Dim strCodMone As String
'Dim strCodBanc As String
'Dim intNumDet  As Integer
'
'
'    dblTipCam = CDbl(Rea_TipCam.Text)
'
'    '*** Del Fondo
'    adoComm.CommandText = "SELECT VAL_CUOT, TIP_VALU FROM FMFONDOS WHERE COD_FOND = '" + strCodFon + "'"
'    Set adoRecord = adoComm.Execute
'    curValNomCuo = adoRecord!Val_cuot
'    strTipValu = adoRecord!TIP_VALU
'    strNewNroCom = Format$(NewNro(strCodFon, "COM"), "00000000")
'    strNroTmp = Format$(NewNro(strCodFon, "TMP"), "00000000")
'
'    '*** Inicializaciones ***
'    intTotOK = 0: intTotER = 0: intTotTr = 0
'    strNroCom = ""
'    strFchOpe = strFecDia ' Luciano Salazar 11/06/1998
'    strFchRet = strFecDia
'
'    strHorOpe = Format(Time, "hh:mm")
'    strPrdCon = Format(Year(FmtFec(strFchOpe, "yyyymmdd", "WIN", intRes)), "0000")
'    strMesCon = Format(Month(FmtFec(strFchOpe, "yyyymmdd", "WIN", intRes)), "00")
'
'    '** Procesar operaciones de suscripción a valor desconocido ***
'    With adoComm
'        .CommandText = "SELECT FMMOVCTA.COD_FOND, FMMOVCTA.NRO_MCTA, FMMOVCTA.FCH_CREA, FMMOVCTA.SUB_SIST, FMMOVCTA.NRO_OPER, "
'        .CommandText = .CommandText & "FMMOVCTA.FCH_OBLI, FMMOVCTA.FCH_CONT, FMMOVCTA.PRD_CONT, FMMOVCTA.MES_CONT, FMMOVCTA.NRO_COMP, "
'        .CommandText = .CommandText & "FMMOVCTA.COD_BANC, FMMOVCTA.NRO_CHEQ, FMMOVCTA.TIP_MOVI, FMMOVCTA.COD_CTA, FMMOVCTA.COD_FILE, "
'        .CommandText = .CommandText & "FMMOVCTA.COD_ANAL, FMMOVCTA.COD_MONE, FMMOVCTA.VAL_MOVI, FMMOVCTA.COM_ORIG, FMMOVCTA.FLG_NVIG, "
'        .CommandText = .CommandText & "FMMOVCTA.TIP_OPER, FMMOVCTA.FLG_CONT, FMMOVCTA.FLG_CONF, FMMOVCTA.COD_PART, FMOPERAC.NRO_FOLI, "
'        .CommandText = .CommandText & "FMMOVCTA.TIP_PAGD, FMMOVCTA.COD_BAND, FMMOVCTA.NRO_CHED, FMMOVCTA.TIP_CTAD, FMMOVCTA.NRO_CTAD,"
'        .CommandText = .CommandText & "FMMOVCTA.COD_SUCD, FMMOVCTA.COD_AGED, FMMOVCTA.COD_FILD, FMMOVCTA.COD_ANAD, FMMOVCTA.VAL_TCMB "
'        .CommandText = .CommandText & "FROM FMMOVCTA, FMOPERAC WHERE FMMOVCTA.COD_FOND = FMOPERAC.COD_FOND AND FMMOVCTA.NRO_OPER = FMOPERAC.NRO_OPER "
'        .CommandText = .CommandText & " AND SUB_SIST='P' AND (FMMOVCTA.TIP_OPER='SD' OR FMMOVCTA.TIP_OPER='SC') AND FMMOVCTA.FLG_CONT='' "
'        .CommandText = .CommandText & " AND FMMOVCTA.FLG_CONF='' AND FMOPERAC.FLG_RETE='X' "
'        .CommandText = .CommandText & " AND FMOPERAC.COD_FOND='" & strCodFon & "' AND FCH_OBLI='" & strFchOpe & "'"
'
'        Set adoMovCta = .Execute
'    End With
'    Do While Not adoMovCta.EOF
'        If (adoMovCta!FCH_CREA = strFchOpe) And (adoMovCta!FCH_OBLI = strFchOpe) Then
'            MsgBox "La operación con número de folio " & Trim$(adoMovCta!nro_foli) & " no se va a procesar!. No se puede liberar cheque de suscripción ingresado el mismo día.", vbCritical
'            adoMovCta.MoveNext
'        End If
'
'        '*** Obtiene el detalle de la operación de suscripción a valor desconocido
'        strNroOper = adoMovCta!NRO_OPER
'        strNroFoli = adoMovCta!nro_foli
'        adoComm.CommandText = "SELECT * FROM FMOPERAC "
'        adoComm.CommandText = adoComm.CommandText & "WHERE COD_FOND='" & strCodFon & "' AND NRO_OPER='" & strNroOper & "' "
'        Set adoRecord = adoComm.Execute
'        curTasOper = adoRecord!TAS_OPER
'        If IsNull(adoRecord!FLG_EXTR) Then
'            strFlgExtr = ""
'        Else
'            strFlgExtr = adoRecord!FLG_EXTR
'        End If
'        strClsPers = adoRecord!CLS_PERS
'        strCodPart = adoRecord!Cod_part
'        strCodMone = adoRecord!COD_MONE
'        strCodSucd = adoRecord!COD_SUCU
'        strCodAged = adoRecord!COD_AGEN
'        strCodProm = adoRecord!COD_PROM
'        strTipOper = adoRecord!TIP_OPER
'        curValMovi = adoMovCta!VAL_MOVI
'
'        adoRecord.Close: Set adoRecord = Nothing
'        adoComm.CommandText = "SELECT FLG_CUST,COD_PROM FROM FMPARTIC "
'        adoComm.CommandText = adoComm.CommandText & "WHERE COD_PART='" & strCodPart & "'"
'        Set adoRecord = adoComm.Execute
'        If IsNull(adoRecord!FLG_CUST) Then
'            strFlgCust = ""
'        Else
'            strFlgCust = adoRecord!FLG_CUST
'        End If
'        If IsNull(adoRecord!COD_PROM) Then
'            strCodProTit = ""
'        Else
'            strCodProTit = adoRecord!COD_PROM
'        End If
'
'        adoRecord.Close: Set adoRecord = Nothing
'
''        '*** Cuenta del partícipe si este tuviera
''        adoComm.CommandText = "SELECT * FROM FMNROCTA WHERE COD_PART = '" & strCodPart & "' "
''        If strCodMone = "S" Then
''            adoComm.CommandText = adoComm.CommandText + " AND (TIP_CTA='0' OR TIP_CTA='2' )"
''        Else
''            adoComm.CommandText = adoComm.CommandText + " AND (TIP_CTA='1' OR TIP_CTA='3' )"
''        End If
''        adoComm.CommandText = adoComm.CommandText + " ORDER BY FLG_DEFA DESC"
'
'        '*** Si existe una Por defecto la asume, si no asume la sgte ***
''        Set adoRecord = adoComm.Execute
''        If Not adoRecord.EOF Then
''            strTipPagd = "U"
''            strCodBand = adoRecord!COD_BANC
''            strTipCtad = adoRecord!TIP_CTA
''            strNroCtad = adoRecord!NRO_CTA
''        Else
'            strTipPagd = "C"
'            strCodBand = adoMovCta!COD_BANC
'            strTipCtad = ""
'            strNroCtad = ""
''        End If
''        adoRecord.Close: Set adoRecord = Nothing
'
'        '*** Obtiene el número de cuotas a valor desconocido ***
'        strCodMon = strCodMone
'        curValAju = dblValcuo * (curTasOper / 100) * (1 + gdblpTasIGV) + dblValcuo
'        dblCntCuo = Format(curValMovi / curValAju, "0.00000")
'        intRes = DoEvents()
'
'        adoComm.CommandText = "SELECT COD_BANC FROM FMCAJA WHERE COD_FOND = '" + strCodFon + "' AND COD_MONE = '" + strCodMon + "' AND COD_BANC = '" + strCodBand + "' AND FLG_VIGE = 'X'"
'        Set adoRecord = adoComm.Execute
'        If Not adoRecord.EOF Then
'            strCodBanc = adoRecord!COD_BANC
'        Else
'            strCodBanc = gstrBancDefa
'        End If
'        adoRecord.Close: Set adoRecord = Nothing
'
'        '*** Obtiene los montos del aporte y capital adicional ***
'        curMtoNom = dblCntCuo * curValNomCuo
'        curValSPa = Format(Abs(IIf((dblValcuo - curValNomCuo) > 0, (dblValcuo - curValNomCuo) * dblCntCuo, 0)), "0.00")
'        curValBPa = Format(Abs(IIf((curValNomCuo - dblValcuo) > 0, (dblValcuo - curValNomCuo) * dblCntCuo, 0)) * -1, "0.00")
'
'        '*** Obtiene los montos de las comisiones, de los tributos y del vuelto ***
'        curMtoCom = Format(dblValcuo * dblCntCuo * curTasOper / 100, "0.00")
'        n_MtoIgv = Format(curMtoCom * gdblpTasIGV, "0.00")
'        curMtoDif = Format(curValMovi - (curMtoNom + curValSPa + curValBPa), "0.00")
'
'        '*** Obtiene el valor total por cobrar ***
'        curValTot = curMtoNom + curValSPa + curValBPa + curMtoCom + n_MtoIgv
'
'        '**** Inicio - Modificación E.H.F. Cuotas Fraccionadas ****
'        If curValTot < curValMovi Then
'            curValSPa = Format(Abs(IIf((dblValcuo - curValNomCuo) > 0, (dblValcuo - curValNomCuo) * dblCntCuo + (curValMovi - curValTot), 0)), "0.00")
'            curValBPa = Format(Abs(IIf((curValNomCuo - dblValcuo) > 0, (dblValcuo - curValNomCuo) * dblCntCuo - (curValMovi - curValTot), 0)) * -1, "0.00")
'        Else
'            curValSPa = Format(Abs(IIf((dblValcuo - curValNomCuo) > 0, (dblValcuo - curValNomCuo) * dblCntCuo - (curValMovi - curValTot), 0)), "0.00")
'            curValBPa = Format(Abs(IIf((curValNomCuo - dblValcuo) > 0, (dblValcuo - curValNomCuo) * dblCntCuo + (curValMovi - curValTot), 0)) * -1, "0.00")
'        End If
'        curValTot = Format(curMtoNom + curValSPa + curValBPa + curMtoCom + n_MtoIgv, "0.00")
'        '**** Fin - Modificación E.H.F. Cuotas Fraccionadas ****
'
'        '*** Obtención de cuentas
'        Call LGetCtaDef("001", "R", strCodCtaN, strCodCtaX)
'        strCtaSus = strCodCtaN
'        Call LGetCtaDef("027", "R", strCodCtaN, strCodCtaX)
'        strCtaSuD = strCodCtaN
'        Call LGetCtaDef("019", "R", strCodCtaN, strCodCtaX)
'        strCtaDif = strCodCtaN
'        If strClsPers = "N" Then    'Naturales
'            If Trim(strFlgExtr) = "" Then  'Nacional
'                Call LGetCtaDef("006", "C", strCodCtaN, strCodCtaX)
'            Else                   'Extranjera
'                Call LGetCtaDef("007", "C", strCodCtaN, strCodCtaX)
'            End If
'        Else                      'Juridicas
'            If Trim(strFlgExtr) = "" Then    'Nacional
'                Call LGetCtaDef("008", "C", strCodCtaN, strCodCtaX)
'            Else                     'Extranjera
'                Call LGetCtaDef("009", "C", strCodCtaN, strCodCtaX)
'            End If
'        End If
'        strCtaApo = strCodCtaN
'        If strClsPers = "N" Then    'Naturales
'            If Trim(strFlgExtr) = "" Then  'Nacional
'                Call LGetCtaDef("010", "C", strCodCtaN, strCodCtaX)
'            Else                   'Extranjera
'                Call LGetCtaDef("011", "C", strCodCtaN, strCodCtaX)
'            End If
'        Else                      'Juridicas
'            If Trim(strFlgExtr) = "" Then    'Nacional
'                Call LGetCtaDef("012", "C", strCodCtaN, strCodCtaX)
'            Else                     'Extranjera
'                Call LGetCtaDef("013", "C", strCodCtaN, strCodCtaX)
'            End If
'        End If
'        strCtaBpa = strCodCtaN
'        If strClsPers = "N" Then    'Naturales
'            If Trim(strFlgExtr) = "" Then  'Nacional
'                Call LGetCtaDef("014", "C", strCodCtaN, strCodCtaX)
'            Else                   'Extranjera
'                Call LGetCtaDef("015", "C", strCodCtaN, strCodCtaX)
'            End If
'        Else                      'Juridicas
'            If Trim(strFlgExtr) = "" Then    'Nacional
'                Call LGetCtaDef("016", "C", strCodCtaN, strCodCtaX)
'            Else                     'Extranjera
'                Call LGetCtaDef("017", "C", strCodCtaN, strCodCtaX)
'            End If
'        End If
'        strCtaSpa = strCodCtaN
'        Call LGetCtaDef("001", "R", strCodCtaN, strCodCtaX) '*** 02/01/98 ***
'        strCtaIgv = strCodCtaN
'
'        '*** Empieza la parte contable ***
'        strNroCom = Format$(NewNro(strCodFon, "COM"), "00000000")
'        strTipCom = "": strTipDoc = "": strNroDoc = ""
'        strDslCom = "S.Desc >> " + IIf(strCodMon = "S", "S/. ", "US$ ") & curValMovi
'        strGloCom = strDslCom
'        strStaCom = ""
'
'        '*** Cta de Suscripciones con Fondos por Confirmar ***
'        adoComm.CommandText = "SP_S_Cuentasb '" & strCodFon & "', '" & strCodBand & "', 'S'"  'LUCIANO SALAZAR 12/11/1998
'        Set adoRecord1 = adoComm.Execute
'        strCodCta16N = adoRecord1!NRO_CTAR
'        adoRecord1.Close: Set adoRecord1 = Nothing
'
'        strCodCtaAct = adoMovCta!COD_CTA
'        If Not IsNull(adoMovCta!COD_FILE) Then
'            strCodFilAct = adoMovCta!COD_FILE
'        Else
'            strCodFilAct = ""
'        End If
'        If Not IsNull(adoMovCta!COD_FILE) Then
'            strCodAnaAct = adoMovCta!COD_ANAL
'        Else
'            strCodAnaAct = ""
'        End If
'
'        gblnRollBack = False
'        On Error GoTo tag_ErrConf:
'        adoConn.Execute "BEGIN TRANSACTION OpeRet"
'
'        '*** Actualiza la operación ***
'        '*** Actualizar a SC ***
'        With adoComm
'            .CommandText = "UPDATE FMOPERAC SET "
'            .CommandText = .CommandText & "CNT_CUOT=" & dblCntCuo & ",VAL_CUOT=" & dblValcuo & ","
'            .CommandText = .CommandText & "FLG_CONT='X',FLG_CONF='X',TIP_OPER='SC',TIP_VCUO='C',"
'            .CommandText = .CommandText & "VAL_COMI=" & curMtoCom & ","
'            .CommandText = .CommandText & "VAL_IGV=" & n_MtoIgv & ","
'            .CommandText = .CommandText & "VAL_TOTA=" & curValTot & ","
'            .CommandText = .CommandText & "VAL_DEVO=" & (curValMovi - curMtoCom - curMtoNom - curValSPa - curValBPa - n_MtoIgv) & ","
'            .CommandText = .CommandText & "TIP_PAGV='" & strTipPagd & "',"
'            .CommandText = .CommandText & "COD_BANV='" & strCodBand & "',"
'            .CommandText = .CommandText & "TIP_CTAV='" & strTipCtad & "',"
'            .CommandText = .CommandText & "NRO_CTAV='" & strNroCtad & "',"
'            .CommandText = .CommandText & "FCH_OPER='" & strFchOpe & "',"
'            .CommandText = .CommandText & "HOR_OPER='A.I.R.',"
'            .CommandText = .CommandText & "HOR_ACTU='A.I.R.',"
'            .CommandText = .CommandText & "FLG_RETE='C' "
'            .CommandText = .CommandText & "WHERE COD_FOND='" & strCodFon & "' and NRO_OPER='" & strNroOper & "'"
'
'            adoConn.Execute .CommandText
'            intRes = DoEvents()
'
'            '*** Saldo Inicial de Cuotas ***
'            .CommandText = "SELECT SUM(CNT_CUOT) SUMCRT FROM FMCERTIF WHERE COD_FOND = '" + strCodFon + "' AND COD_PART = '" + strCodPart + "' AND FLG_VIGE = 'X' AND FCH_SUSC <= '" + strFchOpe + "'"
'            Set adoRecord = .Execute
'            If Not IsNull(adoRecord!SUMCRT) Then
'                dblSinCuo = adoRecord!SUMCRT
'            Else
'                dblSinCuo = 0
'            End If
'            adoRecord.Close: Set adoRecord = Nothing
'
'            '**** Actualiza el detalle de la operación ***
'            .CommandText = "UPDATE FMDETOPE SET "
'            .CommandText = .CommandText & "CNT_CUOT=" & dblCntCuo & ",VAL_CUOT=" & dblValcuo & ", VAL_COMI=" & curMtoCom & ","
'            .CommandText = .CommandText & "VAL_IGV=" & n_MtoIgv & ",VAL_APOR=" & curMtoNom & ",VAL_SPAR=" & curValSPa & ","
'            .CommandText = .CommandText & "VAL_BPAR=" & curValBPa & ",VAL_TOTA= " & curValTot & ", "
'            .CommandText = .CommandText & "SIN_CUOT=" & dblSinCuo
'            .CommandText = .CommandText & "WHERE COD_FOND='" & strCodFon & "' and NRO_OPER='" & strNroOper & "' "
'            adoConn.Execute .CommandText
'            intRes = DoEvents()
'
'            '** Nuevo Nro de Certificado ***
'            strNewNroCrt = Format$(NewNro(strCodFon, "CRT"), "00000000")
'
'            '**** Aqui se debe crear recien el certificado ***
'            WCertif.COD_FOND = strCodFon
'            WCertif.Cod_part = strCodPart
'            WCertif.NRO_CERT = strNewNroCrt
'            WCertif.FCH_CREA = strFchOpe
'            WCertif.FCH_SUSC = strFchOpe
'            WCertif.FCH_REDE = ""
'            WCertif.TIP_OPER = "SC"
'            WCertif.NRO_OPER = strNroOper
'            WCertif.TIP_VCUO = "C"
'            WCertif.CNT_CUOT = dblCntCuo
'            WCertif.Val_cuot = dblValcuo
'            WCertif.COD_MONE = strCodMon
'            WCertif.CLS_PERS = strClsPers
'            WCertif.COD_PROM = strCodProTit
'            WCertif.ORI_OPER = "SC"
'            WCertif.ORI_TOPE = ""
'            WCertif.FIN_OPER = ""
'            WCertif.FIN_TOPE = ""
'            WCertif.NRO_CELI = ""
'            WCertif.FLG_EXTR = strFlgExtr
'            WCertif.FLG_VIGE = "X"
'            WCertif.FLG_EMIS = ""
'            WCertif.FLG_ENTR = ""
'            WCertif.FLG_CONT = "X"
'            WCertif.FLG_CUST = strFlgCust
'            WCertif.FLG_GARA = ""
'            WCertif.FLG_BLOQ = ""
'            Call LGraCertif(WCertif)
'
'            '*** Actualiza el número de cuotas en el Kardex ***
'            Call LGrabKarCuo(strCodFon, strFchOpe, "SC", CDbl(dblCntCuo), " ", strTipValu)
'
'            '*** Inicio Comprobante Contable ***
'            WComCon.CNT_MOVI = 8
'            WComCon.COD_FOND = adoMovCta!COD_FOND
'            WComCon.NRO_COMP = strNewNroCom
'            WComCon.DSL_COMP = "Conf. Susc. con Retención"
'            WComCon.GLO_COMP = "Conf. Suscripción Fondos x Conf."
'            WComCon.COD_MONC = "S"
'            WComCon.COD_MONE = adoMovCta!COD_MONE
'            WComCon.FCH_COMP = strFchOpe
'            WComCon.FCH_CONT = strFchOpe
'            WComCon.FLG_AUTO = ""
'            WComCon.FLG_CONT = "X"
'            WComCon.GEN_COMP = "X"
'            WComCon.HOR_COMP = Format(Time, "hh:mm")
'            WComCon.NRO_DOCU = ""
'            WComCon.NRO_OPER = adoMovCta!NRO_OPER
'            WComCon.PER_DIGI = gstrLogin
'            WComCon.PER_REVI = ""
'            WComCon.MES_CONT = Mid(WComCon.FCH_COMP, 5, 2)
'            WComCon.prd_cont = Mid(WComCon.FCH_COMP, 1, 4)
'            WComCon.STA_COMP = ""
'            WComCon.SUB_SIST = "P"
'            WComCon.TIP_CAMB = IIf(adoMovCta!COD_MONE = "S", 0, CDbl(Rea_TipCam.Text))
'            WComCon.TIP_COMP = ""
'            WComCon.TIP_DOCU = ""
'            WComCon.VAL_COMP = Format(adoMovCta!VAL_MOVI, "0.00")
'
'            '*** Detalle del Comprobante contable ***
'            '*** 1er. Item del Comprobante
'
'            ReDim WMovCon(WComCon.CNT_MOVI)
'            For intNumDet = 1 To WComCon.CNT_MOVI
'                WMovCon(intNumDet).SEC_MOVI = CVar(intNumDet)
'                WMovCon(intNumDet).COD_FOND = strCodFon
'                WMovCon(intNumDet).COD_MONE = adoMovCta!COD_MONE
'                WMovCon(intNumDet).FCH_MOVI = WComCon.FCH_COMP
'                WMovCon(intNumDet).FLG_PROC = "X"
'                WMovCon(intNumDet).NRO_COMP = strNewNroCom
'                WMovCon(intNumDet).prd_cont = WComCon.prd_cont
'                WMovCon(intNumDet).MES_COMP = WComCon.MES_CONT
'                WMovCon(intNumDet).STA_MOVI = "X"
'                WMovCon(intNumDet).TIP_GENR = ""
'                WMovCon(intNumDet).CTA_AMAR = ""
'                WMovCon(intNumDet).CTA_AUTO = ""
'                WMovCon(intNumDet).CTA_ORIG = ""
'                WMovCon(intNumDet).COD_FILE = "00"
'                WMovCon(intNumDet).COD_ANAL = "000000"
'            Next
'
'            '*** Inicio detalle de Comprobante Contable
'            '*** Flag debe haber
'            WMovCon(1).FLG_DEHA = "D"
'            WMovCon(2).FLG_DEHA = "H"
'            WMovCon(3).FLG_DEHA = "D"
'            WMovCon(4).FLG_DEHA = "H"
'            WMovCon(5).FLG_DEHA = "H"
'            WMovCon(6).FLG_DEHA = IIf(curValSPa > 0, "H", "D")
'            WMovCon(7).FLG_DEHA = "H"
'            WMovCon(8).FLG_DEHA = "H"
'
'            '*** Cuenta Contable ***
'            WMovCon(1).COD_CTA = strCodCtaAct
'            WMovCon(2).COD_CTA = strCodCta16N
'            WMovCon(3).COD_CTA = strCtaSuD
'            WMovCon(4).COD_CTA = strCtaSus
'            WMovCon(5).COD_CTA = strCtaApo
'            WMovCon(6).COD_CTA = IIf(curValSPa > 0, strCtaSpa, strCtaBpa)
'            WMovCon(7).COD_CTA = strCtaDif
'            WMovCon(8).COD_CTA = strCtaIgv
'
'            '*** Montos ***
'            WMovCon(1).VAL_MOVN = IIf(adoMovCta!COD_MONE = "S", Format(Abs(curValMovi), "0.00"), 0)
'            WMovCon(1).VAL_MOVX = IIf(adoMovCta!COD_MONE = "S", 0, Format(Abs(curValMovi), "0.00"))
'            WMovCon(1).VAL_CONT = Format(Abs(curValMovi) * IIf(adoMovCta!COD_MONE = "S", 1, dblTipCam), "0.00")
'
'            WMovCon(2).VAL_MOVN = WMovCon(1).VAL_MOVN * -1
'            WMovCon(2).VAL_MOVX = WMovCon(1).VAL_MOVX * -1
'            WMovCon(2).VAL_CONT = WMovCon(1).VAL_CONT * -1
'
'            WMovCon(3).VAL_MOVN = WMovCon(1).VAL_MOVN
'            WMovCon(3).VAL_MOVX = WMovCon(1).VAL_MOVX
'            WMovCon(3).VAL_CONT = WMovCon(1).VAL_CONT
'
'            WMovCon(4).VAL_MOVN = IIf(adoMovCta!COD_MONE = "S", Format(Abs(curMtoCom) * -1, "0.00"), 0)
'            WMovCon(4).VAL_MOVX = IIf(adoMovCta!COD_MONE = "S", 0, Format(Abs(curMtoCom) * -1, "0.00"))
'            WMovCon(4).VAL_CONT = Format(Abs(curMtoCom) * -1 * IIf(adoMovCta!COD_MONE = "S", 1, dblTipCam), "0.00")
'
'            WMovCon(5).VAL_MOVN = IIf(adoMovCta!COD_MONE = "S", Format(Abs(curMtoNom) * -1, "0.00"), 0)
'            WMovCon(5).VAL_MOVX = IIf(adoMovCta!COD_MONE = "S", 0, Format(Abs(curMtoNom) * -1, "0.00"))
'            WMovCon(5).VAL_CONT = Format(Abs(curMtoNom) * -1 * IIf(adoMovCta!COD_MONE = "S", 1, dblTipCam), "0.00")
'
'            WMovCon(6).VAL_MOVN = IIf(adoMovCta!COD_MONE = "S", Format(IIf(Abs(curValSPa) > 0, Abs(curValSPa) * -1, Abs(curValBPa)), "0.00"), 0)
'            WMovCon(6).VAL_MOVX = IIf(adoMovCta!COD_MONE = "S", 0, Format(IIf(Abs(curValSPa) > 0, Abs(curValSPa) * -1, Abs(curValBPa)), "0.00"))
'            WMovCon(6).VAL_CONT = Format(IIf(Abs(curValSPa) > 0, Abs(curValSPa) * -1, Abs(curValBPa)) * IIf(adoMovCta!COD_MONE = "S", 1, dblTipCam), "0.00")
'
'            WMovCon(8).VAL_MOVN = IIf(adoMovCta!COD_MONE = "S", Format(Abs(n_MtoIgv) * -1, "0.00"), 0)
'            WMovCon(8).VAL_MOVX = IIf(adoMovCta!COD_MONE = "S", 0, Format(Abs(n_MtoIgv) * -1, "0.00"))
'            WMovCon(8).VAL_CONT = Format(Abs(n_MtoIgv) * -1 * IIf(adoMovCta!COD_MONE = "S", 1, dblTipCam), "0.00")
'
'            WMovCon(7).VAL_MOVN = IIf(adoMovCta!COD_MONE = "S", Format((WMovCon(3).VAL_MOVN + WMovCon(4).VAL_MOVN + WMovCon(5).VAL_MOVN + WMovCon(6).VAL_MOVN + WMovCon(8).VAL_MOVN) * -1, "0.00"), 0)
'            WMovCon(7).VAL_MOVX = IIf(adoMovCta!COD_MONE = "S", 0, Format((WMovCon(3).VAL_MOVX + WMovCon(4).VAL_MOVX + WMovCon(5).VAL_MOVX + WMovCon(6).VAL_MOVX + WMovCon(8).VAL_MOVX) * -1, "0.00"))
'            WMovCon(7).VAL_CONT = Format((WMovCon(3).VAL_CONT + WMovCon(4).VAL_CONT + WMovCon(5).VAL_CONT + WMovCon(6).VAL_CONT + WMovCon(8).VAL_CONT) * -1, "0.00") ' YA SE ESTAN UTILIZANDO VALORES CONTABLES * IIf(adoMovCta!COD_MONE = "S", 1, dblTipCam)
'
'            '*** Inicio - Modificación E.H.F. Cuotas Fraccionadas ***
'            If (WMovCon(7).VAL_MOVN < 0) Or (WMovCon(7).VAL_MOVX < 0) Or (WMovCon(7).VAL_CONT < 0) Then
'                'No hay modificación
'            Else
'                WMovCon(7).VAL_MOVN = IIf(adoMovCta!COD_MONE = "S", Format((WMovCon(3).VAL_MOVN + WMovCon(4).VAL_MOVN + WMovCon(5).VAL_MOVN + WMovCon(6).VAL_MOVN + WMovCon(8).VAL_MOVN), "0.00"), 0)
'                WMovCon(7).VAL_MOVX = IIf(adoMovCta!COD_MONE = "S", 0, Format((WMovCon(3).VAL_MOVX + WMovCon(4).VAL_MOVX + WMovCon(5).VAL_MOVX + WMovCon(6).VAL_MOVX + WMovCon(8).VAL_MOVX), "0.00"))
'                WMovCon(7).VAL_CONT = Format((WMovCon(3).VAL_CONT + WMovCon(4).VAL_CONT + WMovCon(5).VAL_CONT + WMovCon(6).VAL_CONT + WMovCon(8).VAL_CONT), "0.00") ' YA SE ESTAN UTILIZANDO VALORES CONTABLES * IIf(adoMovCta!COD_MONE = "S", 1, dblTipCam)
'
'                WMovCon(6).VAL_MOVN = WMovCon(6).VAL_MOVN + (WMovCon(7).VAL_MOVN * -2)
'                WMovCon(6).VAL_MOVX = WMovCon(6).VAL_MOVX + (WMovCon(7).VAL_MOVX * -2)
'                WMovCon(6).VAL_CONT = WMovCon(6).VAL_CONT + (WMovCon(7).VAL_CONT * -2)
'            End If
'            '*** Fin - Modificación E.H.F. Cuotas Fraccionadas ***
'
'            WMovCon(8).VAL_MOVN = IIf(adoMovCta!COD_MONE = "S", Format(Abs(n_MtoIgv) * -1, "0.00"), 0)
'            WMovCon(8).VAL_MOVX = IIf(adoMovCta!COD_MONE = "S", 0, Format(Abs(n_MtoIgv) * -1, "0.00"))
'            WMovCon(8).VAL_CONT = Format(Abs(n_MtoIgv) * -1 * IIf(adoMovCta!COD_MONE = "S", 1, dblTipCam), "0.00")
'
'            '*** Descripciones ***
'            WMovCon(1).DSC_MOVI = "Cta Corriente"
'            WMovCon(2).DSC_MOVI = "Susc.XConfirmar"
'            WMovCon(3).DSC_MOVI = "Suscripción Fondos Conf."
'            WMovCon(4).DSC_MOVI = "Comisiones"
'            WMovCon(5).DSC_MOVI = "Aporte Fijo"
'            WMovCon(6).DSC_MOVI = "Aporte Adicional"
'            WMovCon(7).DSC_MOVI = "Devolución"
'            WMovCon(8).DSC_MOVI = "IGV"
'
'            '*** Cta Corriente ***
'            WMovCon(1).COD_FILE = strCodFilAct: WMovCon(1).COD_ANAL = strCodAnaAct
'
'            '*** Cta X Cobrar (16) inicialmente afectada ***
'            WMovCon(2).COD_FILE = strCodFilAct: WMovCon(2).COD_ANAL = strCodAnaAct
'
'            Call LGraAsiCont(WComCon, WMovCon())
'            Call UpdNewNro(strCodFon, "COM", strNewNroCom)
'            Call UpdNewNro(strCodFon, "CRT", strNewNroCrt)
'
'            '*** Actualiza en FMMOVCTA ***
'            .CommandText = "UPDATE FMMOVCTA SET "
'            .CommandText = .CommandText & "NRO_COMP='" & strNewNroCom & "',"
'            .CommandText = .CommandText & "COD_CTA='" & strCodCtaAct & "',"
'            .CommandText = .CommandText & "COD_FILE='" & strCodFilAct & "' , "
'            .CommandText = .CommandText & "COD_ANAL='" & strCodAnaAct & "', "
'            .CommandText = .CommandText & "COD_BANC='" & strCodBand & "', "
'            .CommandText = .CommandText & "FCH_CONT='" & WComCon.FCH_COMP & "', "
'            .CommandText = .CommandText & "MES_CONT='" & WComCon.MES_CONT & "', "
'            .CommandText = .CommandText & "PRD_CONT='" & WComCon.prd_cont & "', "
'            .CommandText = .CommandText & "NRO_FOLI='" & adoMovCta!nro_foli & "', "
'            .CommandText = .CommandText & "FLG_CONF='X', "
'            .CommandText = .CommandText & "FLG_CONT='X' "
'            .CommandText = .CommandText & "WHERE COD_FOND='" & strCodFon & "' AND NRO_MCTA='" & adoMovCta!NRO_MCTA & "'"
'            .CommandText = .CommandText & ""
'            adoConn.Execute .CommandText
'            '*** Fin Comprobante Contable ***
'
'        '*** Creación de la información de la devolución ***
''        If Format(Abs(WMovCon(7).VAL_CONT), "0.00") > 0 Then
''            '*** Si Tiene Cuenta, Nro de Cta y Banco al que se va a depositar ***
''            '*** Crear movimiento de cuenta ***
''            Call UpdNewNro(strCodFon, "TMP", strNroTmp)
''            curValMov = Format(Abs(IIf(strCodMon = "S", WMovCon(7).VAL_MOVN, WMovCon(7).VAL_MOVX)), "0.00")
''            strDscOrd = "Dev.SDes>> " & IIf(strCodMon = "S", "S/. ", "US$ ") & IIf(strCodMon = "S", WMovCon(3).VAL_MOVN, WMovCon(3).VAL_MOVX) & " =" & strFchOpe & "="
''            strTipMov = "S"
''
''            adoComm.CommandText = "INSERT INTO FMMOVCTA "
''            adoComm.CommandText = adoComm.CommandText & "(COD_FOND,FCH_CREA,FCH_OBLI,NRO_MCTA,SUB_SIST,NRO_OPER,TIP_OPER,"
''            adoComm.CommandText = adoComm.CommandText & "TIP_MOVI,COD_FILE,COD_ANAL,COD_MONE,VAL_MOVI,COM_ORIG,COD_PART,COD_BANC,"
''            adoComm.CommandText = adoComm.CommandText & "TIP_PAGD,COD_BAND,NRO_CHED,TIP_CTAD,NRO_CTAD,COD_SUCD,COD_AGED,FLG_CONT,FLG_CONF,NRO_FOLI) VALUES ("
''            adoComm.CommandText = adoComm.CommandText & "'" & strCodFon & "',"
''            adoComm.CommandText = adoComm.CommandText & "'" & strFchOpe & "',"
''            If strTipPagd = "U" Then 'Fecha VCTO
''                adoComm.CommandText = adoComm.CommandText & "'" & strFchRet & "',"
''            Else 'Pago al último día del año
''                adoComm.CommandText = adoComm.CommandText & "'" & strFchRet & "',"
''            End If
''            adoComm.CommandText = adoComm.CommandText & "'" & strNroTmp & "',"
''            adoComm.CommandText = adoComm.CommandText & "'P',"
''            adoComm.CommandText = adoComm.CommandText & "'" & strNroOper & "',"
''            adoComm.CommandText = adoComm.CommandText & "'DD',"
''            adoComm.CommandText = adoComm.CommandText & "'" & strTipMov & "'"
''            adoComm.CommandText = adoComm.CommandText & ",'" & strCodFilAct & "',"
''            adoComm.CommandText = adoComm.CommandText & "'" & strCodAnaAct & "',"
''            adoComm.CommandText = adoComm.CommandText & "'" & strCodMon & "',"
''            adoComm.CommandText = adoComm.CommandText & curValMov * -1 & ","
''            adoComm.CommandText = adoComm.CommandText & "'" & strDscOrd & "',"
''            adoComm.CommandText = adoComm.CommandText & "'" & strCodPart & "',"
''            adoComm.CommandText = adoComm.CommandText & "'" & strCodBanc & "',"
''
''            '*** Información sobre el Destino del Pago ***
''            adoComm.CommandText = adoComm.CommandText & "'" & strTipPagd & "',"    '*** Tipo Pago Destino
''            adoComm.CommandText = adoComm.CommandText & "'" & strCodBand & "',"  '*** Banco Destino
''            adoComm.CommandText = adoComm.CommandText & "'" & "" & "',"   '*** Cheque Destino
''            adoComm.CommandText = adoComm.CommandText & "'" & strTipCtad & "',"     '*** Tipo Cta Destino
''            adoComm.CommandText = adoComm.CommandText & "'" & strNroCtad & "',"     '*** Número Cta Destino
''            adoComm.CommandText = adoComm.CommandText & "'" & strCodSucd & "',"             '*** Sucursal Destino
''            adoComm.CommandText = adoComm.CommandText & "'" & strCodAged & "','','','" & strNroFoli & "')" '*** Agencia Destino, FLG_CONT,FLG_CONF
''            adoConn.Execute adoComm.CommandText
''            intRes = DoEvents()
''
''            '*** Crear movimientos contables temporales ***
''            strCodFil = "00"
''            strCodAna = "000000"
''            strCodCta = strCtaDif
''            strDscMov = "Dev.Suscrip. Confirmada"
''            strSecMov = "001"
''            adoComm.CommandText = "INSERT INTO FMMOVTMP "
''            adoComm.CommandText = adoComm.CommandText & "( COD_FOND,SUB_SIST,NRO_MCTA,SEC_MOVI,DSC_MOVI,"
''            adoComm.CommandText = adoComm.CommandText & "FLG_DEHA,COD_CTA,COD_FILE,COD_ANAL,VAL_MOVI,COD_MONE ) "
''            adoComm.CommandText = adoComm.CommandText & "VALUES ('" & strCodFon & "','P','" & strNroTmp & "','" & strSecMov & "','" & Mid(strDscMov, 1, 20) & "',"
''            adoComm.CommandText = adoComm.CommandText & "'D','" & strCodCta & "','" & strCodFil & "','" & strCodAna & "'," & curValMov & ",'" & strCodMon & "' )"
''            adoConn.Execute adoComm.CommandText
''            intRes = DoEvents()
''        End If
'        End With
'        If Not gblnRollBack Then
'            intTotOK = intTotOK + 1
'            adoConn.Execute "COMMIT TRAN OpeRet"
'            strNewNroCom = Format$(Val(strNewNroCom) + 1, "00000000")
'        End If
'tag_SgtOpeRet:
'        '*** Procesar el siguiente
'        adoMovCta.MoveNext
'        intRes = DoEvents()
'    Loop
'    adoMovCta.Close: Set adoMovCta = Nothing
'    intRes = DoEvents()
'    intTotTr = intTotOK + intTotER
'    strMensaje = "Fin de Proceso!!" & Chr(13) & Chr(10)
'    strMensaje = strMensaje & "Transacciones Canceladas :" & intTotER & Chr(13) & Chr(10)
'    strMensaje = strMensaje & "Transacciones Procesadas :" & intTotOK & Chr(13) & Chr(10)
'    strMensaje = strMensaje & "Total Transacciones      :" & intTotTr
'    MsgBox strMensaje, vbInformation
'    Exit Sub
'
'tag_ErrConf:
'    gblnRollBack = True
'    intTotER = intTotER + 1
'    adoConn.Execute "ROLLBACK TRANSACTION OpeRet"
'    GoTo tag_SgtOpeRet

End Sub

Private Sub LDoGrid()

'    Dim intRes  As Integer
'    Dim intCont As Integer
'    Dim strSQL As String
'
'    If adoOper.State = 1 Then
'        adoOper.Close: Set adoOper = Nothing
'    End If
'    adoConn.CommandTimeout = 360
'    strSQL = "SELECT FMOPERAC.FCH_FRET, FMMOVCTA.FLG_CONF, FMMOVCTA.NRO_MCTA, FMOPERAC.COD_FOND, FMOPERAC.NRO_OPER, FMPARTIC.DSC_PART, FMOPERAC.TIP_PAGO, FMPERSON.DSC_PERS, FMOPERAC.NRO_CTA, FMOPERAC.NRO_CHEQ, FMOPERAC.FCH_OPER, FMOPERAC.HOR_OPER, FMMOVCTA.VAL_MOVI, FMOPERAC.NRO_COMP, FMOPERAC.NRO_FOLI, FMMOVCTA.COD_BAND, FMOPERAC.COD_PART, FMOPERAC.NRO_OPER from FMMOVCTA, FMOPERAC, FMPERSON, FMPARTIC WHERE "
'    strSQL = strSQL & "FMMOVCTA.COD_FOND = FMOPERAC.COD_FOND AND FMMOVCTA.NRO_OPER = FMOPERAC.NRO_OPER AND FMMOVCTA.NRO_FOLI = FMOPERAC.NRO_FOLI AND FMMOVCTA.COD_BAND *= FMPERSON.COD_PERS AND FMMOVCTA.COD_PART = FMOPERAC.COD_PART AND FMMOVCTA.COD_PART = FMPARTIC.COD_PART AND FMOPERAC.COD_PART = FMPARTIC.COD_PART AND "
'    strSQL = strSQL & "(TIP_PAGO = 'C' OR TIP_PAGO = 'T') AND (FMMOVCTA.TIP_OPER = 'SD' OR FMMOVCTA.TIP_OPER = 'SC') AND "
'    strSQL = strSQL & IIf(chk_FlgCon.Value, "", "(FMMOVCTA.FLG_CONF = '' OR FMMOVCTA.FLG_CONF = NULL) AND ")
'    strSQL = strSQL & "FMMOVCTA.COD_FOND = '" & strCodFon & "' AND "
'    strSQL = strSQL & "FCH_FRET = '" & FmtFec(Dat_FchLiq.Value, "WIN", "yyyymmdd", intRes) & "'"
'    strSQL = strSQL & " ORDER BY FMOPERAC.NRO_OPER"
'    adoComm.CommandText = strSQL
'
'    adoOper.CursorLocation = adUseClient
'    adoOper.Open adoComm.CommandText, adoConn, adOpenStatic, adLockBatchOptimistic
'
'    strEstado = ""
'    If Not adoOper.EOF Then
'        Call LlenarGrid(grd_OpeRet, adoOper, aGrdCnf(), adirreg())
'        strEstado = "Consulta"
'    End If
'
'    grd_OpeRet.ColWidth(0) = 660
'
'    grd_OpeRet.Col = 0: grd_OpeRet.Row = 1
'    If Trim(grd_OpeRet.Text) = "" Then Exit Sub
'    adoOper.MoveFirst: intCont = 1
'    If adoOper.EOF Then Exit Sub
'    Do While Not adoOper.EOF
'        ReDim Preserve aBookmark(1 To intCont)
'        aBookmark(intCont) = adoOper.Bookmark
'        adoOper.MoveNext: intCont = intCont + 1
'    Loop
        
End Sub

Private Sub mnu_SOpc_Click(index As Integer)

    Select Case index
        Case 0  'Anular
            LAnuPagRet
            'Comprobante contable Inverso a Ingreso de Cheque
   
        Case 1 'Raya
    
        Case 2 'Terminar
            Unload Me
   
   End Select
End Sub


Private Sub mnu_SSRep1_Click(index As Integer)

'    Dim strTmpSel As String
'
'    gstrNameRepo = "FLCHQLIB2"
'    Set gobjReport = CreateObject("Crystal.CrystalReport")
'    gobjReport.Connect = gstrRptConnectODBC
'    gobjReport.ReportFileName = gstrRptPath & gstrNameRepo & ".RPT"
'    gobjReport.Formulas(0) = "User = '" + gstrLogin + "'"
'    gobjReport.Formulas(1) = "Hora = '" + Format(Time, "HH:MM") + "'"
'
'    If Index = 0 Then
'        gobjReport.StoredProcParam(0) = "L"
'        gobjReport.StoredProcParam(1) = gstrFechaAct
'    Else
'        gobjReport.StoredProcParam(0) = "R"
'        gobjReport.StoredProcParam(1) = gstrFechaAct
'    End If
'        gobjReport.Destination = 0
'        gobjReport.WindowTitle = "(" & gstrNameRepo & ") "
'        gobjReport.WindowState = 2
'        gobjReport.Action = 1
'        DoEvents
End Sub

Private Sub Rea_TipCam_KeyPress(KeyAscii As Integer)
   
'   gdblTipoCamb = CDbl(Rea_TipCam.Text)

End Sub

Private Sub tabRetencion_Click(PreviousTab As Integer)

'    Select Case tabRetencion.Tab
'        Case 0
'            If cmb_FonMut.ListIndex < 0 Then strEstado = ""
'            If cmb_FonMut.ListIndex >= 0 Then strEstado = "Consulta"
'            frmSYSMainMdi.stbMdi.Panels(5).Text = "Acción"
'
'        Case 1
'            If strEstado = "Consulta" Then grd_OpeRet_DblClick
'            If strEstado = "Consulta" Then Call Deshabilita
'            If strEstado = "" Then tabRetencion.Tab = 0
'
'    End Select
    
End Sub

