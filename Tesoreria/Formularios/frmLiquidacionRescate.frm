VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLiquidacionRescate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidación de Operaciones de Rescate "
   ClientHeight    =   5250
   ClientLeft      =   1965
   ClientTop       =   1905
   ClientWidth     =   9195
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
   ScaleHeight     =   5250
   ScaleWidth      =   9195
   Begin VB.Frame fraCambioFecha 
      Caption         =   "Cambio de fecha de liquidación"
      Height          =   1695
      Left            =   4200
      TabIndex        =   12
      Top             =   3000
      Width           =   3375
      Begin VB.CommandButton cmd_Ite 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   735
         Index           =   1
         Left            =   1740
         Picture         =   "frmLiquidacionRescate.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   810
         Width           =   1200
      End
      Begin VB.CommandButton cmd_Ite 
         Appearance      =   0  'Flat
         Caption         =   "&Aceptar"
         Default         =   -1  'True
         Height          =   735
         Index           =   0
         Left            =   375
         Picture         =   "frmLiquidacionRescate.frx":0562
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   810
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker dat_FecChg 
         Height          =   315
         Left            =   1545
         TabIndex        =   15
         Top             =   360
         Width           =   1275
         _ExtentX        =   2249
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
      Begin VB.Label lbl_descam 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nueva fecha"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   16
         Top             =   390
         Width           =   1110
      End
   End
   Begin VB.Frame fraLiquidacion 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.ComboBox cboTipoPago 
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
         TabIndex        =   21
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox Txt_CodUnico 
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
         Left            =   1650
         MaxLength       =   18
         TabIndex        =   18
         Top             =   2325
         Width           =   1305
      End
      Begin VB.CommandButton Cmd_Bsq 
         Caption         =   "Par&tícipe"
         Height          =   735
         Left            =   360
         Picture         =   "frmLiquidacionRescate.frx":09E7
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2160
         Width           =   1200
      End
      Begin VB.CheckBox chk_OptPrn 
         Caption         =   "Individual"
         Height          =   255
         Left            =   5415
         TabIndex        =   11
         Top             =   780
         Width           =   1215
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
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   5055
      End
      Begin VB.CommandButton cmd_Acc 
         Caption         =   "&Estimar"
         Height          =   735
         Index           =   1
         Left            =   6120
         Picture         =   "frmLiquidacionRescate.frx":0FE1
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1200
         Width           =   1200
      End
      Begin VB.CommandButton cmd_Acc 
         Caption         =   "&Procesar"
         Height          =   735
         Index           =   0
         Left            =   7440
         Picture         =   "frmLiquidacionRescate.frx":14E5
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   1200
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
         Left            =   4335
         TabIndex        =   1
         Top             =   780
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpFechaLiquidacion 
         Height          =   315
         Left            =   885
         TabIndex        =   2
         Top             =   720
         Width           =   1275
         _ExtentX        =   2249
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
      Begin VB.Label lbl_descam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Pago"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   20
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label lblDescripParticipe 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   255
         Left            =   3120
         TabIndex        =   19
         Top             =   2325
         Width           =   4455
      End
      Begin VB.Label lbl_descam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   195
         Index           =   6
         Left            =   315
         TabIndex        =   10
         Top             =   750
         Width           =   540
      End
      Begin VB.Label lbl_descam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fondo"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   9
         Top             =   405
         Width           =   540
      End
      Begin VB.Label lbl_descam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cambio"
         Height          =   195
         Index           =   2
         Left            =   2655
         TabIndex        =   8
         Top             =   735
         Width           =   1335
      End
      Begin VB.Label lblTotPag 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999,999,999.99"
         Height          =   255
         Left            =   3420
         TabIndex        =   7
         Top             =   1140
         Width           =   2490
      End
      Begin VB.Label Lbl_DesTot 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Estimado a Pagar (Rescates)"
         Height          =   195
         Left            =   315
         TabIndex        =   6
         Top             =   1140
         Width           =   3060
      End
   End
End
Attribute VB_Name = "frmLiquidacionRescate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim aMapCodFon()
Dim aMapLibAho()
Dim aMapCtaCte()

Dim strCodFon    As String
Dim strCodMon    As String
Dim strComi      As String
Dim strCodPart   As String
Dim strNroFoli   As String
Dim strDscPart   As String

Dim adirreg()    As String
Dim aBookmark()  As Variant
Dim vntDiaAbie   As Variant
Dim strFchDia    As String
Dim strCodPar    As String

Dim intCntOK      As Integer
Dim intCntER      As Integer
Dim intCntTO      As Integer
Dim adoMovCta     As New Recordset
Dim strNewNroCom  As String

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

End Sub

Private Sub calcular()

Dim dblValMovi As Double
Dim intCont    As Integer
      
'    If Cmb_Cod_fond.ListIndex < 0 Then Exit Sub
'
'    '*** Total pago Rescates ***
''    adoComm.CommandText = "SELECT SUM (VAL_MOVI) SumVal FROM FMMOVCTA (Index = PK_FMMOVCTA) WHERE "
''    adoComm.CommandText = adoComm.CommandText + " COD_FOND='" + strCodFon + "' AND SUB_SIST='P' AND (FLG_CONF='' OR FLG_CONF=NULL) "
''    adoComm.CommandText = adoComm.CommandText + " AND TIP_MOVI='S'"
''    adoComm.CommandText = adoComm.CommandText + " AND TIP_OPER LIKE '" & IIf(opt_TipOpe(0).Value, "R_", "D_") & "' "
''    adoComm.CommandText = adoComm.CommandText + " AND TIP_PAGD LIKE '" & IIf(opt_TipPag(0).Value, "U", "C") & "' "
''    If opt_TipPag(0).Value Then 'Efectivo Seleccionar Partícipe
''        adoComm.CommandText = adoComm.CommandText + " AND FCH_OBLI<='" & FmtFec(dat_FchLiq, "WIN", "yyyymmdd", intRes) & "' "
''    Else
''        If strCodPar <> "" Then
''            adoComm.CommandText = adoComm.CommandText + " AND FMMOVCTA.COD_PART = '" & strCodPar & "'"
''        End If
''    End If
'
'    lblTotPag.Caption = Format(0, "Standard")
'    If Grd_MovCta.Rows - 1 > 0 Then
'        Grd_MovCta.Row = 1: Grd_MovCta.Col = 1
'        If Trim(Grd_MovCta.Text) = "" Then Exit Sub
'    End If
'
'    For intCont = 1 To Grd_MovCta.Rows - 1
'        Grd_MovCta.Col = 5: Grd_MovCta.Row = intCont
'        dblValMovi = dblValMovi + Format(IIf(Grd_MovCta.Text = "", "0", Grd_MovCta.Text), "###,###,##0.00")
'    Next
'    lblTotPag.Caption = Format(dblValMovi, "###,###,##0.00")
'    Lbl_DesTot.Caption = "Monto estimado a pagar (" & IIf(opt_TipOpe(0).Value, "Rescates", "Devoluciones") & ")"
    
End Sub

Public Sub Cancelar()

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

Public Sub Ultimo()

End Sub

Private Sub Cmb_Cod_fond_Click()

'    Dim adoRecord As New Recordset
'    Dim strComi   As String * 1
'    Dim intRes    As Integer
'
'    strComi = "'"
'    If Cmb_Cod_fond.ListIndex >= 0 Then
'        strCodFon = Mid(aMapCodFon(Cmb_Cod_fond.ListIndex), 1, 2)
'        strCodMon = Mid(aMapCodFon(Cmb_Cod_fond.ListIndex), 3, 1)
'        lbl_descam(2).Visible = IIf(strCodMon = "S", False, True)
'        Rea_TipCam.Visible = IIf(strCodMon = "S", False, True)
'        Rea_TipCam.Text = IIf(strCodMon = "S", 0, gdblTipoCamb)
'
'        'Frm_DatLiq.Caption = "Datos de la Liquidación " + IIf(strCodMon = "S", "Soles (S/.)", "Dólares (US$)") + " :"
'        '** Fecha disponible para el fondo
'        adoComm.CommandText = "SELECT VAL_CUOT, FCH_CUOT FROM FMCUOTAS WHERE COD_FOND = '" & strCodFon & "' AND FLG_ABIE = 'X' ORDER BY FCH_CUOT"
'        Set adoRecord = adoComm.Execute
'        If Not adoRecord.EOF Then
'            Dat_FchLiq.Value = FmtFec(adoRecord!fch_cuot, "yyyymmdd", "win", intRes)
'            strFchDia = adoRecord!fch_cuot
'            vntDiaAbie = Dat_FchLiq.Value
'        End If
'        adoRecord.Close: Set adoRecord = Nothing
'
'        MousePointer = vbHourglass
'        LDoGrid
'        calcular
'        MousePointer = vbDefault
'    End If

End Sub

Private Sub cmd_Acc_Click(index As Integer)

'    Dim strNroMcta As String
'    Dim strMsg     As String
'    Dim intCntOrd  As Integer
'    Dim intCont    As Integer
'    Dim adoRecord  As New Recordset
'
'    intCntOrd = 0
'
'    Select Case Index
'        Case 0
'            '---------------------------
'            'Procesar Ordenes por pagar
'            '---------------------------
'            For intCont = 1 To Grd_MovCta.Rows - 1
'                Grd_MovCta.Row = intCont
'                Grd_MovCta.Col = 1
'                If Grd_MovCta.Text = "X" Then intCntOrd = intCntOrd + 1
'            Next intCont
'
'            strMsg = "Hay " & intCntOrd & " ordenes por liquidar" & Chr(13) & Chr(10)
'            strMsg = strMsg & "Desea continuar?"
'            If MsgBox(strMsg, 292) <> 6 Then
'                Exit Sub
'            End If
'
'            intCntOK = 0: intCntER = 0: intCntTO = 0
'            For intCont = 1 To Grd_MovCta.Rows - 1
'                Grd_MovCta.Row = intCont
'                Grd_MovCta.Col = 1
'                If Grd_MovCta.Text = "X" Then    'Activada
'                    Grd_MovCta.Col = 2
'                    adoComm.CommandText = "Select *, (select dsc_part from fmpartic where fmpartic.cod_part = fmmovcta.cod_part) dsc_part  from fmmovcta where cod_fond = '" & Mid(aMapCodFon(Cmb_Cod_fond.ListIndex), 1, 2) & "' and nro_foli = '" & Trim(Grd_MovCta.Text) & "'"
'                    Grd_MovCta.Col = 6
'                    adoComm.CommandText = adoComm.CommandText & " and cod_part = '" & Trim(Grd_MovCta.Text) & "'"
'                    Set adoRecord = adoComm.Execute
'
'                    '*** Agregado por CMalpartida 28/10/1998 ***
'                    '*** Se agrego el COD_PART, NRO_FOLI ***
'                    strCodPart = adoRecord!Cod_part
'                    strNroFoli = adoRecord!nro_foli
'                    strDscPart = adoRecord!DSC_PART
'                    '***********
'                    MousePointer = vbHourglass
'                    'LLiqOrdPag strNroMcta
'                    LLiqOrdPag
'                    MousePointer = vbDefault
'                End If
'            Next intCont
'
'            MsgBox "Total Ordenes :" & intCntOrd & Chr(13) & Chr(10) & "Total Ordenes OK:" & intCntOK & Chr(13) & Chr(10) & "Total Ordenes Error:" & intCntER, 16
'            'MousePointer = 11
'            'LDoGrid
'            'calcular
'            'MousePointer = 0
'
'        Case 1
'            '---------------------------
'            'Estimar Monto a Pagar
'            '---------------------------
'            calcular
'    End Select
'    MousePointer = vbHourglass
'    LDoGrid
'    calcular
'    MousePointer = vbDefault

End Sub


Private Sub Cmd_Bsq_Click()

    Dim adoRecord    As New Recordset
    Dim strTmpVar    As String
    Dim strCodProTit As String

    '** Datos del Partícipe
    adoComm.CommandText = "select COD_PART, DSC_PART, RUC_PART, NRO_CUST, COD_KEY, COD_PAIS, CLS_PART, FLG_CUST, FLG_DIROK, COD_PROM from FMPARTIC where COD_UNICO = '" & Trim(Txt_CodUnico.Text) & "'"
    Set adoRecord = adoComm.Execute
    If adoRecord.EOF Then
        If MsgBox("No se ha encontrado al partícipe identificado con código único " & Txt_CodUnico.Text & Chr$(13) & Chr$(10) & "Desea realizar una búsqueda en el Maestro de Partícipes?", 36) = 6 Then
            strTmpVar = LBsqPar()
            strCodPar = Mid(strTmpVar, 1, 15)
            If Trim(strTmpVar) <> "" Then
                adoComm.CommandText = "select COD_PART, COD_UNICO, DSC_PART, RUC_PART, NRO_CUST, COD_KEY, COD_PAIS, CLS_PART, FLG_CUST, FLG_DIROK, COD_PROM from FMPARTIC where COD_PART = '" & strCodPar & "'"
                Set adoRecord = adoComm.Execute
                Txt_CodUnico.Text = adoRecord("COD_UNICO")
            Else
                Txt_CodUnico.SetFocus
                Exit Sub
            End If
        Else
            Txt_CodUnico.SetFocus
            Exit Sub
        End If
    End If
    strCodPar = adoRecord("Cod_part")
    strCodProTit = adoRecord("COD_PROM")
    lblDescripParticipe.Caption = adoRecord("DSC_PART")
    MousePointer = vbHourglass
    LDoGrid
    calcular
    MousePointer = vbDefault
    adoRecord.Close: Set adoRecord = Nothing

End Sub

Private Sub cmd_Ite_Click(index As Integer)

'    Dim intRes    As Integer
'    Dim vntFecOld As Variant
'    Dim adoRecord As New Recordset
'
'    Select Case Index
'        Case 0 'Cambiar fecha de liquidación
'            'Validar Fecha Correcta
'
'            adoComm.CommandText = "Select * from fmmovcta where cod_fond = '" & Mid(aMapCodFon(Cmb_Cod_fond.ListIndex), 1, 2) & "' and cod_part = '" & strCodPart & "' and nro_foli = '" & strNroFoli & "'"
'            Set adoRecord = adoComm.Execute
'
'            vntFecOld = FmtFec(adoRecord!FCH_OBLI, "yyyymmdd", "win", intRes)
'            If DateDiff("d", vntDiaAbie, dat_FecChg.Value) < 0 Then
'                MsgBox "La nueva fecha de liquidación no puede ser menor a " & vntDiaAbie
'                Exit Sub
'            End If
'
'            '*** Se le quito en el Where NRO_MCTA y se agrego COD_PART y NRO_FOLI
'            adoComm.CommandText = "UPDATE FMMOVCTA SET FCH_OBLI = '" & FmtFec(dat_FecChg.Value, "win", "yyyymmdd", intRes) & "' "
'            adoComm.CommandText = adoComm.CommandText + " WHERE COD_FOND ='" & adoRecord!COD_FOND & "'"
'            adoComm.CommandText = adoComm.CommandText + " AND COD_PART ='" & strCodPart & "'"
'            adoComm.CommandText = adoComm.CommandText + " AND NRO_FOLI ='" & strNroFoli & "'"
'            adoConn.Execute adoComm.CommandText
'
'        Case 1 'Cancela opción
'
'        Case Else
'    End Select
'
'    frm_DatFon.Enabled = True
'    fra_FecLiq.Visible = False 'Frame Hide

End Sub

Private Sub cmd_Print_Click()

'    On Error GoTo tag_ErrPrn
'
'    Dim intCont   As Integer
'    Dim strTmpSel As String
'
'    gstrNameRepo = "FLPAGVUE"
'    If chk_OptPrn.Value Then 'Individual
'        gstrSelFrml = " {FMMOVCTA.TIP_OPER} = 'DD' AND {FMMOVCTA.COD_FOND} = '" + strCodFon + "' AND {FMMOVCTA.FLG_CONF} = '' AND {FMMOVCTA.TIP_PAGD} = 'C' "
'        gstrSelFrml = gstrSelFrml + "  AND ("
'        For intCont = 1 To UBound(aBookmark)
'            adoMovCta.Bookmark = aBookmark(intCont)
'            gstrSelFrml = gstrSelFrml + "{FMMOVCTA.NRO_MCTA}='" & adoMovCta!NRO_MCTA & "' OR "
'        Next intCont
'
'        'SACO EL ULTIMO OR
'        gstrSelFrml = Left(gstrSelFrml, Len(gstrSelFrml) - 4)
'        gstrSelFrml = gstrSelFrml + ")"
'
'        Set gobjReport = CreateObject("Crystal.CrystalReport")
'        gobjReport.Connect = gstrRptConnectODBC
'        gobjReport.Formulas(0) = "User='" + gstrLogin + "'"
'        gobjReport.Formulas(1) = "CodFond = '" & strCodFon & "'"
'        gobjReport.Formulas(2) = "DscFond ='" + Left$(Cmb_Cod_fond + Space(40), 40) + "'"
'        gobjReport.SelectionFormula = ""
'        gobjReport.SelectionFormula = gstrSelFrml
'        gobjReport.ReportFileName = gstrRptPath & gstrNameRepo & ".RPT"
'        gobjReport.Destination = 0
'        gobjReport.WindowTitle = "(" & gstrNameRepo & ") "
'        gobjReport.WindowState = 2
'        gobjReport.Action = 1
'        DoEvents
'    Else 'Masiva
'        strTmpSel = "{FMMOVCTA.FCH_CREA} IN 'Fch1' TO 'Fch2'"
'        gstrSelFrml = strTmpSel
'        frmSYSRngFch.Show 1
'        If gstrSelFrml <> "0" Then
'            gstrSelFrml = gstrSelFrml + " AND {FMMOVCTA.TIP_OPER}='DD' AND {FMMOVCTA.COD_FOND}='" + strCodFon + "' AND {FMMOVCTA.FLG_CONF}='' AND {FMMOVCTA.TIP_PAGD}='C' "
'
'            Set gobjReport = CreateObject("Crystal.CrystalReport")
'            gobjReport.Connect = gstrRptConnectODBC
'            gobjReport.Formulas(0) = "User='" + gstrLogin + "'"
'            gobjReport.Formulas(1) = "CodFond = '" & strCodFon & "'"
'            gobjReport.Formulas(2) = "DscFond='" + Left$(Cmb_Cod_fond + Space(40), 40) + "'"
'            gobjReport.SelectionFormula = ""
'            gobjReport.SelectionFormula = gstrSelFrml
'            gobjReport.ReportFileName = gstrRptPath & gstrNameRepo & ".RPT"
'            gobjReport.Destination = 0
'            gobjReport.WindowTitle = "(" & gstrNameRepo & ") "
'            gobjReport.WindowState = 2
'            gobjReport.Action = 1
'            DoEvents
'        End If
'    End If
'    Exit Sub
'
'tag_ErrPrn:
'    If Err = 9 Then 'Si no hay registros
'        Exit Sub
'    End If

End Sub

Private Sub cmd_salir_Click()
   
   Unload Me

End Sub


Private Sub Form_Load()

'    Dim strSentencia As String
'
'    adoComm.CommandTimeout = 360
'
'    strComi = "'"
'
'    strSentencia = "SELECT COD_FOND + COD_MONE CODIGO, DSC_FOND DESCRIP FROM FMFONDOS ORDER BY DSC_FOND"
'    Call LCmbLoad(strSentencia, Cmb_Cod_fond, aMapCodFon(), "")
'    If Cmb_Cod_fond.ListCount > 0 Then Cmb_Cod_fond.ListIndex = 0
'
'    '---------------------------------------------------------------------
'    'Configurar Grilla de Movimientos
'    '---------------------------------------------------------------------
'    ReDim aGrdCnf(1 To 6)
'    aGrdCnf(1).TitDes = "Conf."
'    aGrdCnf(1).DatNom = "FLG_CONF"
'    aGrdCnf(1).DatAnc = 130 * 2
'
'    aGrdCnf(2).TitDes = "Nro.Folio"
'    aGrdCnf(2).DatNom = "NRO_FOLI"
'    aGrdCnf(2).DatAnc = 130 * 10
'
'    aGrdCnf(3).TitDes = "Partícipe."
'    aGrdCnf(3).DatNom = "DSC_PART"
'    aGrdCnf(3).DatAnc = 130 * 21
'
'    aGrdCnf(4).TitDes = "Fch.Soli"
'    aGrdCnf(4).DatNom = "FCH_CREA"
'    aGrdCnf(4).DatAnc = 130 * 8
'
'    aGrdCnf(5).TitDes = "Monto"
'    aGrdCnf(5).DatNom = "VAL_MOVI"
'    aGrdCnf(5).DatAnc = 130 * 10
'    aGrdCnf(5).DatJus = 1
'    aGrdCnf(5).DatFmt = "D"
'
'    aGrdCnf(6).TitDes = "Cod_Part"
'    aGrdCnf(6).DatNom = "COD_PART"
'    aGrdCnf(6).DatAnc = 1
    
    CentrarForm Me
    
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmLiquidacionRescate = Nothing
    
End Sub

Private Sub Grd_MovCta_Click()

'    Grd_MovCta.SelStartRow = Grd_MovCta.Row: Grd_MovCta.SelStartCol = 1
'    Grd_MovCta.SelEndRow = Grd_MovCta.Row: Grd_MovCta.SelEndCol = Grd_MovCta.Cols - 1
'    If Grd_MovCta.Col = 1 Then
'        '----------------------------------------------
'        'Actualizar
'        '----------------------------------------------
'        If Grd_MovCta.Text = "" Then
'            Grd_MovCta.Text = "X"
'        Else
'            Grd_MovCta.Text = ""
'        End If
'        LFlgCnf 1
'        '----------------------------------------------
'    End If
    
End Sub

Private Sub Grd_MovCta_DblClick()

'    Dim intRes    As Integer
'    Dim adoRecord As New Recordset
'
'    If Grd_MovCta.Row > 0 Then
'        Grd_MovCta.Col = 1
'        If Trim(Grd_MovCta.Text) = "" Then
'        Exit Sub
'    End If
'
'    Grd_MovCta.Col = 2
'    adoComm.CommandText = "Select * from fmmovcta where cod_fond = '" & Mid(aMapCodFon(Cmb_Cod_fond.ListIndex), 1, 2) & "' and nro_foli = '" & Trim(Grd_MovCta.Text) & "'"
'    Grd_MovCta.Col = 6
'    adoComm.CommandText = adoComm.CommandText & " and cod_part = '" & Trim(Grd_MovCta.Text) & "'"
'    Set adoRecord = adoComm.Execute
'
'    'If adorecord!TIP_PAGD = "U" Then
'        dat_FecChg.Value = FmtFec(adoRecord!FCH_OBLI, "yyyymmdd", "win", intRes)
'        '*** Cambio de fecha de liqidación **
'        frm_DatFon.Enabled = False
'        fra_FecLiq.Visible = True 'Frame Hide
'        strCodPart = adoRecord!Cod_part
'        Grd_MovCta.Col = 2
'        strNroFoli = Grd_MovCta.Text
'        End If
'    'End If
'    adoRecord.Close: Set adoRecord = Nothing
    
End Sub

Private Function LBsqPar() As String

'    Dim adoRecord As New Recordset
'
'    frmBusquedaParticipe.Show vbModal
'
'    If Trim$(gstrLlamaSoli) = "" Then
'        strCodPar = ""
'        LBsqPar = ""
'    Else
'      adoComm.CommandText = "Sp_S_LiqPart '01', '" & Trim(gstrLlamaSoli) & "'"
'      Set adoRecord = adoComm.Execute
'      LBsqPar = adoRecord!Cod_part & adoRecord!DSC_PART
'      adoRecord.Close: Set adoRecord = Nothing
'    End If
    
End Function

Private Sub LDoGrid()

'    Dim intRes    As Integer
'    Dim intCont   As Integer
'
'    If Cmb_Cod_fond.ListIndex = -1 Then Exit Sub
'
'    If opt_TipPag(1).Value And strCodPar <> "" Then
'        adoComm.CommandText = "Sp_S_LiqPart '02', '" & strCodFon & "', '" & IIf(opt_TipPag(0).Value, "U", "CT") & "', '" & strCodPar & "'"
''        adoComm.CommandText = "Select COD_FOND, ' ' NRO_MCTA, FCH_CREA, ' ' NRO_OPER, FCH_OBLI, SUM(VAL_MOVI) VAL_MOVI, ' ' COM_ORIG, FLG_NVIG, ' ' TIP_OPER,  FLG_CONT, FLG_CONF, FMMOVCTA.COD_PART, NRO_FOLI, TIP_PAGD, DSC_PART FROM FMMOVCTA (Index = PK_FMMOVCTA), FMPARTIC WHERE "
''        adoComm.CommandText = adoComm.CommandText + " FMMOVCTA.COD_PART = FMPARTIC.COD_PART "
''        adoComm.CommandText = adoComm.CommandText + " AND COD_FOND = '" & strCodFon & "' "
''        adoComm.CommandText = adoComm.CommandText + " AND FLG_CONF = '' "
''        adoComm.CommandText = adoComm.CommandText + " AND TIP_OPER LIKE '" & IIf(opt_TipOpe(0).Value, "R_", "D_") & "' "
''        adoComm.CommandText = adoComm.CommandText + " AND TIP_PAGD = '" & IIf(opt_TipPag(0).Value, "U", "C") & "' "
''        adoComm.CommandText = adoComm.CommandText + " AND FMMOVCTA.COD_PART = '" & strCodPar & "'"
'    End If
'
'    If (opt_TipPag(1).Value = True And Trim(strCodPar) = "") Or (opt_TipPag(0).Value = True And Trim(strCodPar) = "") Then 'Efectivo Seleccionar Partícipe
'        adoComm.CommandText = "Sp_S_LiqPart '03', '" & FmtFec(Dat_FchLiq.Value, "win", "yyyymmdd", intRes) & "', '" & IIf(opt_TipPag(0).Value, "U", "CT") & "', '" & strCodFon & "'"
'
''        adoComm.CommandText = "Select COD_FOND, ' ' NRO_MCTA, FCH_CREA, ' ' NRO_OPER, FCH_OBLI, SUM(VAL_MOVI) VAL_MOVI, ' ' COM_ORIG, FLG_NVIG, ' ' TIP_OPER,  FLG_CONT, FLG_CONF, FMMOVCTA.COD_PART, NRO_FOLI, TIP_PAGD, DSC_PART FROM FMMOVCTA (Index = PK_FMMOVCTA), FMPARTIC WHERE "
''        adoComm.CommandText = adoComm.CommandText + " FMMOVCTA.COD_PART = FMPARTIC.COD_PART "
''        adoComm.CommandText = adoComm.CommandText + " AND COD_FOND = '" & strCodFon & "' "
''        adoComm.CommandText = adoComm.CommandText + " AND FLG_CONF = '' "
''        adoComm.CommandText = adoComm.CommandText + " AND TIP_OPER LIKE '" & IIf(opt_TipOpe(0).Value, "R_", "D_") & "' "
''        adoComm.CommandText = adoComm.CommandText + " AND TIP_PAGD = '" & IIf(opt_TipPag(0).Value, "U", "C") & "' "
''        adoComm.CommandText = adoComm.CommandText + " Group By NRO_FOLI, COD_FOND, FCH_CREA, FCH_OBLI, FLG_NVIG, FLG_CONT, FLG_CONF, FMMOVCTA.COD_PART, NRO_FOLI, TIP_PAGD, DSC_PART "
''        adoComm.CommandText = adoComm.CommandText + " ORDER BY FCH_OBLI ASC, DSC_PART ASC"
'    ElseIf opt_TipPag(0).Value And Trim(strCodPar) = "" Then
''        adoComm.CommandText = "Select COD_FOND, ' ' NRO_MCTA, FCH_CREA, ' ' NRO_OPER, FCH_OBLI, SUM(VAL_MOVI) VAL_MOVI, ' ' COM_ORIG, FLG_NVIG, ' ' TIP_OPER,  FLG_CONT, FLG_CONF, FMMOVCTA.COD_PART, NRO_FOLI, TIP_PAGD, DSC_PART FROM FMMOVCTA (Index = PK_FMMOVCTA), FMPARTIC WHERE "
''        adoComm.CommandText = adoComm.CommandText + " FMMOVCTA.COD_PART = FMPARTIC.COD_PART "
''        adoComm.CommandText = adoComm.CommandText + " AND COD_FOND = '" & strCodFon & "' "
''        adoComm.CommandText = adoComm.CommandText + " AND FLG_CONF = '' "
''        adoComm.CommandText = adoComm.CommandText + " AND TIP_OPER LIKE '" & IIf(opt_TipOpe(0).Value, "R_", "D_") & "' "
''        adoComm.CommandText = adoComm.CommandText + " AND TIP_PAGD = '" & IIf(opt_TipPag(0).Value, "U", "C") & "' "
''        adoComm.CommandText = adoComm.CommandText + " AND FCH_OBLI ='" & FmtFec(dat_FchLiq, "WIN", "yyyymmdd", intRes) & "' "
''        adoComm.CommandText = adoComm.CommandText + " Group By NRO_FOLI, COD_FOND, FCH_CREA, FCH_OBLI, FLG_NVIG, FLG_CONT, FLG_CONF, FMMOVCTA.COD_PART, NRO_FOLI, TIP_PAGD, DSC_PART "
''        adoComm.CommandText = adoComm.CommandText + " ORDER BY FCH_OBLI ASC, DSC_PART ASC"
'    End If
'
'    adoConn.CommandTimeout = 360
'    adoMovCta.CursorLocation = adUseClient
'    Set adoMovCta = adoComm.Execute
'    If Not adoMovCta.EOF Then
'        Call LlenarGrid(Grd_MovCta, adoMovCta, aGrdCnf(), adirreg())
'        LFlgCnf 0
'    End If
'    adoMovCta.Close: Set adoMovCta = Nothing
    
End Sub

Private Sub LFlgCnf(index As Integer)
    '0 Activa Todos con "X" y Sumariza
    '1 Sumariza aquellos activados con "X"

'    Dim dblTotRes As Double
'    Dim intCont  As Integer
'
'    dblTotRes = 0
'    For intCont = 1 To Grd_MovCta.Rows - 1
'        Grd_MovCta.Row = intCont
'        Grd_MovCta.Col = 2  'NroMcta
'        If Index = 0 And Grd_MovCta.Text <> "" Then
'            Grd_MovCta.Col = 1
'            Grd_MovCta.Text = "X"
'        End If
'        Grd_MovCta.Col = 1
'        If Grd_MovCta.Text = "X" Then
'            Grd_MovCta.Col = 5
'            If IsNumeric(Grd_MovCta.Text) Then
'               dblTotRes = dblTotRes + CDbl(Grd_MovCta.Text)
'            End If
'        End If
'    Next
'    lblTotPag.Caption = Format(dblTotRes, "###,###,##0.00")

End Sub

Private Sub LLiqOrdPag()

'    On Error GoTo tag_ErrLiq:
'
'    Dim adoCom         As New Recordset
'    Dim adoDet         As New Recordset
'    Dim intRes         As Integer
'    Dim adoRecord      As New Recordset
'    Dim WComCon        As RCabasicon         'Comprobante contable
'    Dim WMovCon()      As RDetasicon         'Movimientos del comprobante contable
'    Dim intNumDet      As Integer
'    Dim dblTipCam      As Double
'    Dim strFilCaj      As String * 2
'    Dim strAnaCaj      As String * 6
'    Dim dblValDep      As Double
'    Dim strCtaAct      As String
'    Dim strBcoPag      As String
'    Dim strDscBan      As String
'    Dim intCntReg      As Integer
'    Dim adoRecord1     As New Recordset
'    Dim blnPagoRescate As Boolean
'    Dim adoNewNro      As New Recordset
'
'    '** Pagar Ordenes por Cobrar pendientes de confirmación
'    '** Aquellas ordenes pendientes generadas por subsistema "P"
'    '** Definir
'    dblTipCam = CDbl(Rea_TipCam.Text)
'
'    '*** Inicializa Acc. de Parametro ***
'    blnPagoRescate = False
'
'    '** Nuevo Nro de Comprobante Contable ***
'    adoComm.CommandText = "Select NRO_ULTI_SOLI from FMPARAME where COD_FOND = '" & strCodFon & "' AND COD_PARA = 'COM'"
'    Set adoNewNro = adoComm.Execute
'    strNewNroCom = Format(adoNewNro!NRO_ULTI_SOLI + 1, "00000000")
'    adoNewNro.Close: Set adoNewNro = Nothing
'
'    adoConn.CursorLocation = adUseClient
'    adoComm.CommandText = "SELECT * FROM FMMOVCTA (Index = PK_FMMOVCTA) WHERE COD_FOND = '" & strCodFon & "' AND NRO_FOLI = '" & strNroFoli & "' AND COD_PART = '" & strCodPart & "' AND FLG_CONF <> 'X'"
'    Set adoCom = adoComm.Execute
'    If adoCom.EOF Then
'        MsgBox "No se encontró la orden " & adoCom!NRO_MCTA & "."
'        Exit Sub
'    End If
'
'    intCntReg = adoCom.RecordCount
'    intCntTO = 0
'    Do While Not adoCom.EOF
'        intCntTO = intCntTO + 1
'
'        '*** Inicio de Transacción ***
'        gblnRollBack = False
'        adoComm.CommandText = "BEGIN TRANSACTION OrdPagMas"
'        adoComm.Execute
'
'        '*** Si el Pago es con cheque de Otro Banco ***
'        If (Not IsNull(adoCom!COD_BANC) Or Trim(adoCom!COD_BANC) <> "") Then
'            strBcoPag = adoCom!COD_BANC
'        Else
'            strBcoPag = gstrBancDefa
'        End If
'
'        '*** Identificar Con que Cuenta se va a pagar ***
'        adoComm.CommandText = "SP_S_Cuentasb '" & strCodFon & "', '" & strBcoPag & "', 'R'"   'LUCIANO SALAZAR 12/11/1998
'        Set adoRecord1 = adoComm.Execute
'
'        If Not adoRecord1.EOF Then
'            strFilCaj = adoRecord1!COD_FILE: strAnaCaj = adoRecord1!COD_ANAL: strCtaAct = adoRecord1!NRO_CTAC
'        Else
'            adoComm.CommandText = "SP_S_Cuentasb '" & strCodFon & "', '" & strBcoPag & "', 'R'"   'LUCIANO SALAZAR 12/11/1998
'            Set adoRecord1 = adoComm.Execute
'
'            strFilCaj = "": strAnaCaj = "": strCtaAct = ""
'            If Not adoRecord1.EOF Then
'                strFilCaj = adoRecord1!COD_FILE: strAnaCaj = adoRecord1!COD_ANAL: strCtaAct = adoRecord1!NRO_CTAC
'            Else
'                strFilCaj = "": strAnaCaj = "": strCtaAct = ""
'                MsgBox "Error... El Banco de la operación con Nro. de Folio : " & adoCom!nro_foli & " es el Banco Santander", vbInformation, gstrNombreEmpresa
'                gblnRollBack = True
'                GoTo tag_sgteMovCta:
'            End If
'        End If
'        If adoRecord1.State = 1 Then
'            adoRecord1.Close: Set adoRecord1 = Nothing
'        End If
'        '*** Fin Identificar Cuenta ***
'
'        dblValDep = 0
'        WComCon.COD_FOND = strCodFon
'        WComCon.NRO_COMP = strNewNroCom
'        WComCon.DSL_COMP = adoCom!com_orig
'        WComCon.GLO_COMP = adoCom!com_orig
'        WComCon.COD_MONC = "S"
'        WComCon.COD_MONE = adoCom!COD_MONE
'        WComCon.FCH_COMP = strFchDia ' Modificado por L.S. 10/06/98
'        WComCon.FCH_CONT = strFchDia
'        WComCon.FLG_AUTO = ""
'        WComCon.FLG_CONT = "X"
'        WComCon.GEN_COMP = "X"
'        WComCon.HOR_COMP = Format(Time, "hh:mm")
'        WComCon.NRO_DOCU = ""
'        WComCon.NRO_OPER = adoCom!NRO_OPER
'        WComCon.PER_DIGI = gstrLogin
'        WComCon.PER_REVI = ""
'        WComCon.MES_CONT = Mid(WComCon.FCH_COMP, 5, 2)
'        WComCon.prd_cont = Mid(WComCon.FCH_COMP, 1, 4)
'        WComCon.STA_COMP = ""
'        WComCon.SUB_SIST = "P"
'        WComCon.TIP_CAMB = IIf(adoCom!COD_MONE = "S", 0, CDbl(Rea_TipCam.Text))  'rea_TipCam.
'        WComCon.TIP_COMP = ""
'        WComCon.TIP_DOCU = ""
'        WComCon.VAL_COMP = Format(adoCom!VAL_MOVI, "0.00")
'
'        '** Detalle del Comprobante contable ***
'        '** 1er. Item del Comprobante ***
'        intNumDet = 1
'        ReDim WMovCon(intNumDet)
'        WMovCon(intNumDet).SEC_MOVI = CVar(intNumDet)
'        WMovCon(intNumDet).COD_FOND = strCodFon
'        WMovCon(intNumDet).COD_MONE = adoCom!COD_MONE
'        WMovCon(intNumDet).FCH_MOVI = WComCon.FCH_COMP
'        WMovCon(intNumDet).FLG_PROC = "X"
'        WMovCon(intNumDet).NRO_COMP = strNewNroCom
'        WMovCon(intNumDet).prd_cont = WComCon.prd_cont
'        WMovCon(intNumDet).MES_COMP = WComCon.MES_CONT
'        WMovCon(intNumDet).STA_MOVI = "X"
'        WMovCon(intNumDet).TIP_GENR = ""
'        WMovCon(intNumDet).CTA_AMAR = ""
'        WMovCon(intNumDet).CTA_AUTO = ""
'        WMovCon(intNumDet).CTA_ORIG = ""
'        WMovCon(intNumDet).COD_FILE = strFilCaj
'        WMovCon(intNumDet).COD_ANAL = strAnaCaj
'        WMovCon(intNumDet).DSC_MOVI = strDscPart
'        WMovCon(intNumDet).FLG_DEHA = IIf(adoCom!TIP_MOVI = "E", "D", "H")
'        WMovCon(intNumDet).COD_CTA = strCtaAct
'        WMovCon(intNumDet).VAL_MOVN = Format(IIf(adoCom!COD_MONE = "S", adoCom!VAL_MOVI, 0), "0.00")
'        WMovCon(intNumDet).VAL_MOVX = Format(IIf(adoCom!COD_MONE = "S", 0, adoCom!VAL_MOVI), "0.00")
'        WMovCon(intNumDet).VAL_CONT = Format(adoCom!VAL_MOVI * IIf(adoCom!COD_MONE = "S", 1, dblTipCam), "0.00")
'
'        adoComm.CommandText = "SELECT * FROM FMMOVTMP WHERE COD_FOND = '" & strCodFon & "' AND NRO_MCTA = '" & adoCom!NRO_MCTA & "' ORDER BY SEC_MOVI"
'        Set adoDet = adoComm.Execute
'        Do While Not adoDet.EOF
'            intNumDet = intNumDet + 1
'            ReDim Preserve WMovCon(intNumDet)
'            WMovCon(intNumDet).SEC_MOVI = CVar(intNumDet)
'            WMovCon(intNumDet).COD_FOND = adoDet!COD_FOND
'            WMovCon(intNumDet).COD_MONE = adoDet!COD_MONE
'            WMovCon(intNumDet).FCH_MOVI = WComCon.FCH_COMP
'            WMovCon(intNumDet).FLG_PROC = "X"
'            WMovCon(intNumDet).NRO_COMP = strNewNroCom
'            WMovCon(intNumDet).prd_cont = WComCon.prd_cont
'            WMovCon(intNumDet).MES_COMP = WComCon.MES_CONT
'            WMovCon(intNumDet).STA_MOVI = "X"
'            WMovCon(intNumDet).TIP_GENR = ""
'            WMovCon(intNumDet).CTA_AMAR = ""
'            WMovCon(intNumDet).CTA_AUTO = ""
'            WMovCon(intNumDet).CTA_ORIG = ""
'            WMovCon(intNumDet).COD_FILE = adoDet!COD_FILE
'            WMovCon(intNumDet).COD_ANAL = adoDet!COD_ANAL
'            WMovCon(intNumDet).DSC_MOVI = strDscPart
'            WMovCon(intNumDet).FLG_DEHA = adoDet!FLG_DEHA
'            WMovCon(intNumDet).COD_CTA = adoDet!COD_CTA
'            WMovCon(intNumDet).VAL_MOVN = Format(IIf(adoDet!COD_MONE = "S", adoDet!VAL_MOVI, 0), "0.00")
'            WMovCon(intNumDet).VAL_MOVX = Format(IIf(adoDet!COD_MONE = "S", 0, adoDet!VAL_MOVI), "0.00")
'            WMovCon(intNumDet).VAL_CONT = Format(adoDet!VAL_MOVI * IIf(adoDet!COD_MONE = "S", 1, dblTipCam), "0.00")
'            adoDet.MoveNext
'        Loop
'        adoDet.Close: Set adoDet = Nothing
'
'        adoComm.CommandText = "UPDATE FMMOVCTA SET "
'        adoComm.CommandText = adoComm.CommandText & "NRO_COMP='" & strNewNroCom & "',"
'        adoComm.CommandText = adoComm.CommandText & "COD_CTA='" & strCtaAct & "',"
'        adoComm.CommandText = adoComm.CommandText & "COD_FILE='" & strFilCaj & "' , "
'        adoComm.CommandText = adoComm.CommandText & "COD_ANAL='" & strAnaCaj & "', "
'        adoComm.CommandText = adoComm.CommandText & "COD_BANC='" & strBcoPag & "', "
'        adoComm.CommandText = adoComm.CommandText & "FCH_CONT='" & WComCon.FCH_COMP & "', "
'        adoComm.CommandText = adoComm.CommandText & "MES_CONT='" & WComCon.MES_CONT & "', "
'        adoComm.CommandText = adoComm.CommandText & "PRD_CONT='" & WComCon.prd_cont & "', "
'        adoComm.CommandText = adoComm.CommandText & "NRO_CHEQ='" & txt_NroChe.Text & "', "
'        adoComm.CommandText = adoComm.CommandText & "NRO_CHED='" & txt_NroChe.Text & "', "
'        adoComm.CommandText = adoComm.CommandText & "FLG_CONF='X', "
'        adoComm.CommandText = adoComm.CommandText & "FLG_CONT='X' "
'        adoComm.CommandText = adoComm.CommandText & "WHERE COD_FOND='" & strCodFon & "' AND NRO_MCTA='" & adoCom!NRO_MCTA & "'"
'        adoConn.Execute adoComm.CommandText
'
'        WComCon.CNT_MOVI = intNumDet
'        Call LGraAsiMsg(WComCon, WMovCon())
'        strNewNroCom = Format(Val(strNewNroCom) + 1, "00000000")
'
'tag_sgteMovCta:
'        If gblnRollBack Then
'            adoComm.CommandText = "ROLLBACK TRAN OrdPagMas"
'            adoComm.Execute
'            If intCntTO <= intCntReg Then
'                intCntER = intCntER + 1
'             End If
'        Else
'            adoComm.CommandText = "COMMIT TRAN OrdPagMas"
'            adoComm.Execute
'            If intCntTO = intCntReg Then
'                intCntOK = intCntOK + 1
'            End If
'        End If
'        adoCom.MoveNext
'    Loop
'    adoCom.Close: Set adoCom = Nothing
'
'    '** Controlar Transacción
'    '** Actualiza Nuevo Nro de Comprobante Contable
'    Call UpdNewNro(strCodFon, "COM", strNewNroCom)
'
'    '----------------------------------------------
'    'Imprime pago de Vueltos
'    '----------------------------------------------
'    PanMen.Caption = ""
'    Exit Sub
'
'tag_ErrLiq:
'    gblnRollBack = True
'    MsgBox Err.Description
'    GoTo tag_sgteMovCta:
   
End Sub

Private Sub opt_TipOpe_Click(index As Integer, Value As Integer)
    
    MousePointer = vbHourglass
    LDoGrid
    calcular
    MousePointer = vbDefault
    
End Sub

Private Sub opt_TipPag_Click(index As Integer, Value As Integer)

'    Select Case Index
'        Case 0 'Abono en Cta
'            'Esconde Botón
'            chk_OptPrn.Visible = False
'            'cmd_Print.Visible = False
'            'cmd_Print.Enabled = False
'            pic_BsqPar.Enabled = False
'            pic_BsqPar.Visible = False
'            Fra_DesCam(1).Width = 2115
'
'        Case 1 'Efectivo Cheque
'            'Visualizar Botón
'            chk_OptPrn.Visible = True
'            'cmd_Print.Enabled = True
'            'cmd_Print.Visible = True
'            pic_BsqPar.Enabled = True
'            pic_BsqPar.Visible = True
'            txt_NroChe.Text = ""
'            Fra_DesCam(1).Width = 3375
'            'Inicializa info Partícipe
'            strCodPar = ""
'            Txt_CodUnico.Text = " ": Txt_CodUnico.Text = ""
'            lblDescripParticipe.Caption = ""
'            Txt_CodUnico.SetFocus
'    End Select
'    frmPROLiqPart.Refresh
'    MousePointer = vbHourglass
'    LDoGrid
'    calcular
'    MousePointer = vbDefault

End Sub

Private Sub Rea_TipCam_LostFocus()

'    gdblTipoCamb = CDbl(Rea_TipCam.Text)
    
End Sub

Private Sub Txt_CodUnico_Change()

    Cmd_Bsq.Default = True
    
End Sub

Private Sub Txt_CodUnico_LostFocus()

    If Trim(Txt_CodUnico.Text) = "" Then strCodPart = ""
    
End Sub


