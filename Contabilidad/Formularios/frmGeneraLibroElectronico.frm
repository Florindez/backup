VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmGeneraLibroElectronico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generar Archivos Regulatorios"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8970
   Icon            =   "frmGeneraLibroElectronico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraListaArchivos 
      Caption         =   "Lista de Archivos"
      Height          =   2895
      Left            =   30
      TabIndex        =   12
      Top             =   1890
      Width           =   8865
      Begin VB.ListBox lst_DscRepo 
         Height          =   2010
         Left            =   360
         MultiSelect     =   2  'Extended
         TabIndex        =   13
         Top             =   450
         Width           =   7875
      End
   End
   Begin VB.CommandButton cmdSalr 
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
      Left            =   7020
      Picture         =   "frmGeneraLibroElectronico.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Generar Archivos"
      Top             =   4860
      Width           =   1200
   End
   Begin VB.CommandButton CmdGenerarArchivos 
      Caption         =   "&Generar"
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
      Left            =   5670
      Picture         =   "frmGeneraLibroElectronico.frx":058E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Generar Archivos"
      Top             =   4860
      Width           =   1200
   End
   Begin VB.Frame fraCriterio 
      Caption         =   "Criterios para la Generación"
      Height          =   1785
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   8865
      Begin VB.CommandButton cmd_Listar 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   315
         Left            =   7740
         TabIndex        =   11
         Top             =   1200
         Width           =   315
      End
      Begin VB.TextBox txtDestino 
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "C:\"
         Top             =   1200
         Width           =   6165
      End
      Begin VB.ComboBox cboFondo 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   6555
      End
      Begin MSComCtl2.DTPicker dtpFechaDesde 
         Height          =   315
         Left            =   1500
         TabIndex        =   6
         Top             =   780
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         Format          =   182059009
         CurrentDate     =   38779
      End
      Begin MSComCtl2.DTPicker dtpFechaHasta 
         Height          =   315
         Left            =   6510
         TabIndex        =   7
         Top             =   780
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         Format          =   182059009
         CurrentDate     =   38779
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   5460
         TabIndex        =   9
         Top             =   840
         Width           =   825
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   900
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Destino"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   390
         TabIndex        =   4
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fondo"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   390
         TabIndex        =   3
         Top             =   420
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmGeneraLibroElectronico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim strCodFondo     As String, arrFondo()      As String
Dim aCodRep() As String 'AGREGADO POR RRG
Dim s_FchRepo As String 'AGREGADO POR RRG
Dim s_FchRepoWin As String 'AGREGADO POR RRG
Dim NL$ 'AGREGADO POR RRG
Dim ID_File As String 'AGREGADO POR RRG
Dim RptPath As String 'AGREGADO POR RRG
Dim dd_File As String 'AGREGADO POR RRG
Dim mm_File As String 'AGREGADO POR RRG
Dim EXT_File As String 'AGREGADO POR RRG
Dim strCodAdministradoraCNS As String 'AGREGADO POR RRG
Dim s_fecha As String 'AGREGADO POR RRG
'Dim dd_File As String
'Dim mm_File As String
Dim aa_File As String
Dim s_Rec As String
Dim n_NroRepo As Integer
Dim n_Cont  As Integer
Dim s_FonMut As String
Dim s_CODFONDO As String
Dim n_PorcCom As Integer
Dim s_FchInic As String

''''INCLUSION ACR ACR
Dim amap_FonMut()
'Dim s_FonMut As String * 2
'Dim s_CODFONDO As String * 4
Dim s_CodMone As String * 2
Dim s_FchVcto As String * 8

'H.R.P.>> 10/02/98
'Dim EXT_File As String
'Dim ID_File As String
'Dim dd_File As String
'Dim mm_File As String
'Dim aa_File As String
'Dim s_fecha As String

'Dim aCodRep() As String
Dim aranRep() As String
Dim aclsRep() As String
Dim ACodFon() As String
Dim TipRang As String * 1
Dim ClsRepo As String * 1

'Dim RptPath As String
'Dim NL$


Dim n_CntRegi As Integer
Dim strNombreArchivo As String

'Dim n_NroRepo As Integer
'Dim Dn_Var As Dynaset
'Dim s_FchRepo As String
'Dim s_FchRepoWin As String





Private Sub cboFondo_Click()
'
'  strCodFondo = Valor_Caracter
'    If cboFondo.ListIndex < 0 Then Exit Sub
'
'    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    '******************
    '******************
    'MODIFICADO POR RRG
    'Call Buscar
    
    Dim adoRegistro As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = adoRegistro("FechaCuota")
            gdblTipoCambio = adoRegistro("ValorTipoCambio")
            gstrCodMoneda = adoRegistro("CodMoneda")
            dtpFechaDesde.Value = DateAdd("d", -1, gdatFechaActual)
            dtpFechaHasta.Value = dtpFechaDesde.Value
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    
End Sub

Private Sub cmd_Listar_Click()

   frm_ListaDir.Show 1
   If gs_FormName <> "" And gs_FormName <> "c:\" Then
      txtDestino.Text = gs_FormName + "\"
   End If
   If gs_FormName = "c:\" Then
      txtDestino.Text = gs_FormName
   End If

End Sub

Private Sub CmdGenerarArchivos_Click()
 
    'Dim Separador As String
    Dim n_dias, res As Integer, gn As Integer
    Dim NomFile As String
    Dim i As Integer
    
    'Separador = Trim(Me.txtSeparador.Text)
    
    'On Error GoTo CmdGenerarArchivos_Exit
    
    If Not ValidForm() Then
       Exit Sub
    End If
    
    For gn = 0 To lst_DscRepo.ListCount - 1
        If lst_DscRepo.Selected(gn) Then
           lst_DscRepo.ListIndex = gn
           NomFile = aCodRep(lst_DscRepo.ListIndex)
           n_dias = DateDiff("d", dtpFechaDesde.Value, dtpFechaHasta.Value)
           'For i = 0 To n_dias
               's_FchRepo = fmtfec(DateAdd("d", I, dtpFechaDesde.Value), "win", "yyyymmdd", res)
               s_FchRepoWin = DateAdd("d", i, dtpFechaDesde.Value)
               Call GenerarArchivos(dtpFechaDesde.Value, NomFile)
           'Next i
        End If
        lst_DscRepo.Selected(gn) = False
        lst_DscRepo.Refresh
    Next gn
    
    MsgBox "Archivo(s) generado(s) correctamente.", 48
    Exit Sub
    
CmdGenerarArchivos_Exit:
    Me.MousePointer = 0
    MsgBox Error(err)
    MsgBox "Error de generación de archivos.", 48
    Exit Sub

    'GeneraArchivoRegulatorio_Participes (Separador)
    'GeneraArchivoRegulatorio_SaldosContables (Separador)
    'GeneraArchivoRegulatorio_ValorCuotayCuentasPrincipales (Separador)

    'MsgBox "Se genero el Archivo"

End Sub


'Private Sub GeneraArchivoRegulatorio_Participes() '(ByVal Separador As String)
'
'    Dim strFechaDesde As String
'
'    'Open Trim(txtDestino.Text) & "1Conasev-" & Format(Date, "dd-MM-yy") & ".txt" For Output As #1
'    Open Trim(txtDestino.Text) For Output As #1
'
'    Me.MousePointer = vbHourglass
'
'    Dim adoRegistro As ADODB.Recordset
'
'    Set adoRegistro = New ADODB.Recordset
'
'    strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
'
'    With adoComm
'
''        .CommandText = " up_ACParametroCONASEV (1,'" & strCodFondo & "','" & gstrCodAdministradora & "','','')"
'        .CommandText = "{call up_CNGeneraArchivoParticipesCNSV ('" & strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaDesde & "')}"
'
'        Set adoRegistro = .Execute
'
'        'FormatoConasev_X(Trim(adoRegistro.Fields(4).Value), 20) & Separador & _
'
'        Do While Not adoRegistro.EOF
'            Print #1, _
'
'                    Trim (adoRegistro.Fields(0).Value) & Separador & _
'                    Trim(adoRegistro.Fields(1).Value) & Separador & _
'                    Trim(adoRegistro.Fields(2).Value) & Separador & _
'                    Format(Trim(adoRegistro.Fields(3).Value), "ddMMyyyy") & Separador & _
'                    FormatoConasev_9_V9(Trim(adoRegistro.Fields(4).Value), 20) & Separador & _
'                    Trim(adoRegistro.Fields(5).Value) & Separador & _
'                    Trim(adoRegistro.Fields(6).Value) & Separador & _
'                    FormatoConasev_9_V9(Format(Trim(adoRegistro.Fields(7).Value), "."), 12, 8) & Separador & _
'                    FormatoConasev_9_V9(Format(Trim(adoRegistro.Fields(8).Value), "."), 12, 8) & Separador & _
'                    FormatoConasev_9_V9(Format(Trim(adoRegistro.Fields(9).Value), "."), 12, 8) & Separador & _
'                    FormatoConasev_9_V9(Format(Trim(adoRegistro.Fields(10).Value), "."), 12, 8) & Separador & _
'                    FormatoConasev_9_V9(Format(Trim(adoRegistro.Fields(11).Value), "."), 12, 8) & Separador & _
'                    FormatoConasev_9_V9(Format(Trim(adoRegistro.Fields(12).Value), "."), 12, 8)
'
'            adoRegistro.MoveNext
'        Loop
'        adoRegistro.Close: Set adoRegistro = Nothing
'    End With
'
'    Close #1
'    Me.MousePointer = vbDefault
'
'End Sub
Private Sub GeneraArchivoRegulatorio(strCodRegistro As String, strCodAdministradora As String, datFechaReporte As Date)
 
    Dim numCampos As Integer
    Dim n As Integer
    Dim strRegistro As String
    Dim strFechaRegistro As String
    Dim numNumRegistroControl As Long
    Dim numCantRegistros As Long
    Dim numOut As Integer
    
    Me.MousePointer = vbHourglass
    
    strFechaRegistro = Convertyyyymmdd(datFechaReporte)
    
    numCantRegistros = 0
    
    
    Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
    
    numCampos = 0
    strRegistro = ""
    
    With adoComm
        If strCodRegistro = "RC" Then
          .CommandText = "{call up_RepRegistroCompraElectronico "
        ElseIf strCodRegistro = "RV" Then
            .CommandText = "{call up_RepRegistroVentaElectronico "
        End If
        
        .CommandText = .CommandText & "('" & strCodFondo & "','" & gstrCodAdministradora & "','" & Convertyyyymmdd(dtpFechaDesde.Value) & "','" & Convertyyyymmdd(dtpFechaHasta.Value) & "')}"
        
        Set adoRegistro = .Execute
        

        If Not adoRegistro.EOF Then
            strNombreArchivo = strNombreArchivo & "111.TXT"
            numCampos = adoRegistro.Fields.Count
        Else
            strNombreArchivo = strNombreArchivo & "011.TXT"
        End If

        numOut = FreeFile
        Open strNombreArchivo For Binary Access Read Write As numOut


        Do While Not adoRegistro.EOF
            
            For n = 0 To numCampos - 1
                strRegistro = strRegistro & adoRegistro.Fields(n).Value & "|"
            Next n
                
            strRegistro = strRegistro + Chr(13) + Chr(10)
                
            'Print #1, strRegistro
            Put numOut, , strRegistro
        
            strRegistro = ""
            
            numCantRegistros = numCantRegistros + 1
        
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
'    'ACTUALIZAR REGISTRO DE CONTROL
'    strRegistro = ""
'    strRegistro = strRegistro + "C"
'    strRegistro = strRegistro + Format(Trim(strCodAdministradoraCNS), "@@@@@@")
'    strRegistro = strRegistro + strCodRegistro
'    strRegistro = strRegistro + Right(strFechaRegistro, 2) + Mid(strFechaRegistro, 5, 2) + Left(strFechaRegistro, 4)
'    strRegistro = strRegistro + Format(numNumRegistroControl, "0000")
'    strRegistro = strRegistro + Format(numCantRegistros, "00000")
'
    Put numOut, 1, strRegistro
    
    Close #1
    'Libro 8.2 vacio:
    If strCodRegistro = "RC" Then
        strNombreArchivo = Left(strNombreArchivo, Len(strNombreArchivo) - 15) & "80200001011.TXT"
        numOut = FreeFile
        Open strNombreArchivo For Binary Access Read Write As numOut
    
    '    Call ActualizaRegistroControl(strCodRegistro, gstrCodAdministradora, strFechaRegistro, numNumRegistroControl, numCantRegistros)
        Put numOut, 1, strRegistro
        
        Close #1
    End If
    
    Me.MousePointer = vbDefault

End Sub

Private Sub GeneraArchivoRegulatorio_SaldosContables(ByVal Separador As String)
 
    Open Trim(txtDestino.Text) & "2Conasev-" & Format(Date, "dd-MM-yy") & ".txt" For Output As #1
    
    Me.MousePointer = vbHourglass
    Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = " up_ACParametroCONASEV (2,'" & strCodFondo & "','" & gstrCodAdministradora & "','','')"
        Set adoRegistro = .Execute

        'FormatoConasev_X

        Do While Not adoRegistro.EOF
            Print #1, _
                    Trim(adoRegistro.Fields(0).Value) & Separador & _
                    Trim(adoRegistro.Fields(1).Value) & Separador & _
                    Trim(adoRegistro.Fields(2).Value) & Separador & _
                    Format(Trim(adoRegistro.Fields(3).Value), "ddMMyyyy") & Separador & _
                    Trim(adoRegistro.Fields(4).Value) & Separador & _
                    FormatoConasev_9_V9(Trim(adoRegistro.Fields(5).Value), 6) & Separador & _
                    FormatoConasev_9_V9(Format(Trim(adoRegistro.Fields(6).Value), "."), 12, 8)
                    
            
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    Close #1
    Me.MousePointer = vbDefault

End Sub

Private Sub GeneraArchivoRegulatorio_ValorCuotayCuentasPrincipales(ByVal Separador As String)

    Open Trim(txtDestino.Text) & "3Conasev-" & Format(Date, "dd-MM-yy") & ".txt" For Output As #1

    Me.MousePointer = vbHourglass
    Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = " up_ACParametroCONASEV  3,'" & strCodFondo & "','" & gstrCodAdministradora & "','" & Convertyyyymmdd(dtpFechaDesde.Value) & "','" & Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value)) & "'"
        Set adoRegistro = .Execute

        Do While Not adoRegistro.EOF
            Print #1, _
                    Trim(adoRegistro.Fields(0).Value) & Separador & _
                    Trim(adoRegistro.Fields(1).Value) & Separador & _
                    Trim(adoRegistro.Fields(2).Value) & Separador & _
                    Format(Trim(adoRegistro.Fields(3).Value), "ddMMyyyy") & Separador & _
                    FormatoConasev_9_V9(Format(Trim(adoRegistro.Fields(4).Value), "."), 12, 8) & _
                    FormatoConasev_9_V9(Format(Trim(adoRegistro.Fields(5).Value), "."), 12, 8) & _
                    FormatoConasev_9_V9(Format(Trim(adoRegistro.Fields(6).Value), "."), 12, 8) & _
                    FormatoConasev_9_V9(Format(Trim(adoRegistro.Fields(7).Value), "."), 12, 8) & _
                    FormatoConasev_9_V9(Format(Trim(adoRegistro.Fields(8).Value), "."), 12, 8) & _
                    FormatoConasev_9(Format(Trim(adoRegistro.Fields(9).Value), "."), 6) & _
                    FormatoConasev_9_V9(Format(Trim(adoRegistro.Fields(10).Value), "."), 8, 2) & _
                    FormatoConasev_9_V9(Format(Trim(adoRegistro.Fields(11).Value), "."), 2, 8) & _
                    FormatoConasev_9_V9(Format(Trim(adoRegistro.Fields(12).Value), "."), 5, 8) & _
                    FormatoConasev_9_V9(Format(Trim(adoRegistro.Fields(13).Value), "."), 8, 8)
                    
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    Close #1
    Me.MousePointer = vbDefault

End Sub

Public Sub Buscar()
    
    Dim strSQL  As String
    Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
    strSQL = "SELECT * FROM ArchivoRegulatorio WHERE CodAdministradora='" & gstrCodAdministradora & "' AND EstadoArchivo = '01' AND CodArchivo >= 12 ORDER BY CodRegistro"
    
    With adoComm
        .CommandText = strSQL
        Set adoRegistro = .Execute
        
        'lst_DscRepo.Clear
        
        Do While Not adoRegistro.EOF
        
            lst_DscRepo.AddItem Trim(adoRegistro.Fields(3).Value) 'Nombre del Archivo
            ReDim Preserve aCodRep(lst_DscRepo.ListCount - 1)
            aCodRep(lst_DscRepo.ListCount - 1) = Trim(adoRegistro.Fields(2).Value) 'Código del Archivo
        
        adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
    
    End With
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub cmdSalr_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    Call Buscar
    Call CargarListas
    Call CodAdministradoraCNS
    
    'InicializarValores
    
'    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
'    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
'
    CentrarForm Me

End Sub

Private Sub CargarListas()

    Dim strSQL  As String
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Sel_Todos
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 1
    
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
    
        Case vSearch
            Call Buscar
        Case vReport
'            Call Imprimir
        Case vExit
            Call Salir
    
    End Select
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub
'Private Sub InicializarValores()
                      
    '*** Valores Iniciales ***
    '*** Ancho por defecto de las columnas de la grilla ***
'    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 10
'    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 60
   
'End Sub

'**********************************************************************
'********************** AGREGADO POR RRG
Private Function ValidForm() As Boolean
  Dim i As Integer
  Dim s As Integer
  Dim Msg As String
  Dim nRetVal As Boolean
  
  nRetVal = False
  
  If cboFondo.ListIndex < 0 Then
     Msg = "Debe Elegir el Fondo Mutuo para Generar el Archivo."
     GoTo FinCnsv
  End If
  'validar que seleccione por lo menos un reporte
  s = 0
  For i = 0 To lst_DscRepo.ListCount - 1
      If Not lst_DscRepo.Selected(i) Then
         s = s + 1
      End If
  Next i
  'si no ha seleccionado ningun reporte...
  If s = lst_DscRepo.ListCount Then
     Msg = "Debe Seleccionar un Registro en la Lista."
     GoTo FinCnsv
  End If
  
  nRetVal = True
  
FinCnsv:
   If Msg <> "" Then
      MsgBox (Msg)
      Exit Function
   End If
   ValidForm = nRetVal
End Function



Private Sub GenerarArchivos(FECHA As Date, File As String)
    Dim adoAux As ADODB.Recordset
    Dim res As Integer
    Dim strAnio As String
    Dim strMes As String
    Dim strDia As String
    Dim strName2 As String
    
    '----------------------------------------------------
    'DR 04/05/99 NOMENCLATURA DEL ARCHIVO A GENERAR
    '----------------------------------------------------
    RptPath = Trim(txtDestino.Text)
    EXT_File = Trim(Mid$(strCodAdministradoraCNS, 1, 1) + Mid$(strCodAdministradoraCNS, 5, 6))
    
    adoComm.CommandText = "select NumRucFondo from Fondo where CodFondo = '" & strCodFondo & "'"
    Set adoAux = adoComm.Execute
    
    strAnio = Year(FECHA)
    strMes = Format(Month(FECHA), "00")
    strDia = Format(Day(FECHA), "00")
    
    s_fecha = strAnio + strMes '+ strDia
    
    dd_File = strDia
    mm_File = strMes
    aa_File = strAnio
    
    NL$ = Chr(13) + Chr(10)
    
    If File = "RC" Then
        strName2 = "0008010000"
    ElseIf File = "RV" Then
        strName2 = "0014010000"
    End If
    
    'ID_File = Mid$(Trim$(File), 1, 2)
    '----------------------------------------------------------
    'GENERANDO EL NOMBRE DEL ARCHIVO
    '----------------------------------------------------------
    strNombreArchivo = RptPath + "LE" + Trim$(adoAux("NumRucFondo")) + s_fecha + strName2 + "1"
    
    Call GeneraArchivoRegulatorio(File, gstrCodAdministradora, FECHA)
    
    
'    Select Case Trim$(File)
'        '**********************
'        '     Primera Fase
'        '**********************
'        Case "VC"
'            'Call Genera_VC ' Archivo de Valor Cuota
'        Case "CO"
'            'Call Genera_COaammdd ' Archivo de Cobertura
'        Case "SC"
'            'Call Genera_SC ' Archivo de Saldos Contables
'        Case "PT"   ' Archivo de Participes
'            Call GeneraArchivoRegulatorio_Participes     'Call Genera_PTaammdd
'        Case "PS"   ' Archivo de Personas
'            'Call Genera_PS
'        Case "VI"   ' Archivo de Personas Relacionadas
'            'Call Genera_VIaammdd
'        Case "TD"   ' Registro de Tabla de Desarrollo
'            'Call Genera_TDaammdd
'        Case "VL"   ' Archivo de Valorizacion
'            'Call Genera_VLaammdd
'        Case "OF"   ' Archivo de Oficinas
'            'Call Genera_OFaammdd
'        Case "VA"   ' Archivo de Valores No Inscritos
'            'Call Genera_VAaammdd
'
'    End Select

End Sub

Private Sub CodAdministradoraCNS()
    
    Dim strSQL As String
    Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
    strSQL = "SELECT CodConasev FROM Administradora WHERE CodAdministradora='" & gstrCodAdministradora & "'"
    
    With adoComm
        .CommandText = strSQL
        Set adoRegistro = .Execute
        
        strCodAdministradoraCNS = adoRegistro.Fields(0).Value
        
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub

Private Sub Genera_VC()
    
    ' Generacion del archivo de valor de cuota.
    Dim strSQL As String
    Dim n_Out As Integer, res As Integer
    Dim s_out As String * 176 '(174 longitud del registro)
    Dim s_con As String * 27, s_Espacio As String * 1
    Dim n_ACTPCIE As Double, n_PATPCIE As Double, n_RESEJER As Double
    Dim s_CtaRem As String, n_RemDIAR  As Double, n_RemAcum As Double
    Dim n_Cont As Integer, s_CodMone As String
    Dim NL$
    Dim adoFondos As ADODB.Recordset
    Dim adoCuota As ADODB.Recordset
    Dim adoCuenta As ADODB.Recordset
    Dim adoVar As ADODB.Recordset
    'Dim adoCuenta As ADODB.Recordset

    n_ACTPCIE = 0: n_PATPCIE = 0: n_RESEJER = 0: s_CtaRem = "": n_RemDIAR = 0: n_RemAcum = 0
    
    NL$ = Chr(13) + Chr(10)
    n_Cont = 0
    
    strSQL = "SELECT CodFondo,CodConasev,CodMoneda FROM Fondo "
    'sensql = sensql + "WHERE TipoFondo = '0' ORDER BY CodConasev ASC"
    '****************************
    'PREGUNTAR A ACR SI ES TipoFondo o CodAdministradora
    '****************************
    strSQL = strSQL + "WHERE CodAdministradora = gstrCodAdministradora ORDER BY CodConasev ASC"
    
    Set adoFondos = New ADODB.Recordset
    
    With adoComm
        .CommandText = strSQL
        Set adoFondos = .Execute
        
        If adoFondos.EOF Then
            MsgBox "No existen Fondos Mutuos disponibles para el proceso.", 48
            Exit Sub
        Else
            adoFondos.MoveLast
            adoFondos.MoveFirst
        End If
    End With
    
    'OBTENER NUMERO DE ARCHIVO
    'Call GeneraNroArchivoCnsv("VC", gstrCodAdministradora, s_fecha)
    
    n_Out = FreeFile
    Open txtDestino For Binary Access Read Write As n_Out
    
    '*********************
    'POR EVALUAR
    ' REGISTRO DE CONTROL
    s_Rec = ""
    s_Rec = s_Rec + "C"
    s_Rec = s_Rec + Format(strCodAdministradoraCNS, "@@@@@@")
    s_Rec = s_Rec + "VC"
    s_Rec = s_Rec + Right(s_fecha, 2) + Mid(s_fecha, 5, 2) + Left(s_fecha, 4)
    s_Rec = s_Rec + Format$(n_NroRepo, "0000")
    s_Rec = s_Rec + Format$(n_Cont, "00000")
    s_Rec = s_Rec + Format(Left$(s_Espacio, 148), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    s_Rec = s_Rec + NL$
    s_out = s_Rec
    Put n_Out, , s_out
    '**************************
    '**************************
    
    Do Until adoFondos.EOF
        
        ' PARAMETROS PARA LA GENERACION
        s_FonMut = adoFondos("CodFondo"): s_CODFONDO = Trim$(adoFondos("CodConasev")): s_CodMone = adoFondos("CodMoneda")
        
        strSQL = "SELECT * FROM FondoValorCuota WHERE CodFondo='" + s_FonMut + "' AND "
        strSQL = strSQL + "IndCierre='X' AND FechaCuota='" + s_fecha + "' "
        Set adoCuota = New ADODB.Recordset
    
        With adoComm
            .CommandText = strSQL
            Set adoCuota = .Execute
        
            If adoCuota.EOF Then
                MsgBox "Fecha del Fondo No Disponible para el proceso.", 48
                dtpFechaDesde.SetFocus
                Exit Sub
            Else
                adoCuota.MoveLast
                adoCuota.MoveFirst
            End If
        End With
                
        n_PorcCom = ((((adoCuota("TasaAdministracion") / 100) + 1) ^ (1 / 360)) - 1) * 100 * (1 + gdblTasaIgv)
        n_PorcCom = Format$(n_PorcCom, "0.00000000")

        Do Until adoCuota.EOF
            
            strSQL = "SELECT CodCuenta FROM DinamicaContable "
            strSQL = strSQL + "WHERE TipoOperacion = '18' AND CodFile = '098' AND CodAdministradora = '" + gstrCodAdministradora + "' AND"
            strSQL = strSQL + "CodDetalleFile='001'"
            
            Set adoCuenta = New ADODB.Recordset
    
            With adoCuenta
                '.CommandText = strSQL
                'Set adoCuenta = .Execute
        
                If adoCuenta.EOF Then
                    s_CtaRem = ""
                Else
                    s_CtaRem = adoCuenta("CodCuenta")
                End If
                adoCuenta.Close: Set adoCuenta = Nothing
            End With
                        
            
            'REMUNERACION DIARIA Y ACUMULADA A LA SAFM
            'sensql = "SELECT COD_CTAN, COD_CTAX FROM FMCTADEF WHERE TIP_DEFI='R' AND COD_DEFI='003'"
            'Set Dn_Var = GBdFm.CreateDynaset(sensql, 64)
            'If Dn_Var.EOF Then
            '    s_CtaRem = ""
            'Else
            '    If s_CodMone = "S" Then
            '       s_CtaRem = Trim$(Dn_Var!cod_ctan)
            '    Else
            '       s_CtaRem = Trim$(Dn_Var!cod_ctax)
            '    End If
            'End If
            'Dn_Var.Close
        
            
            strSQL = "SELECT SUM(SaldoParcialContable) REM_DIAR FROM PartidaContableSaldos WHERE "
            strSQL = strSQL + "CodFondo='" + s_FonMut + "' AND FechaSaldo='" + s_fecha + "' AND "
            strSQL = strSQL + "CodCuenta = '" + s_CtaRem + "'"
            
            Set adoVar = New ADODB.Recordset
    
            With adoVar
                '.CommandText = strSQL
                'Set adoVar = .Execute
        
                If adoVar.EOF Then
                    n_RemDIAR = Format(0, "0.00")
                Else
                    n_RemDIAR = Format(adoVar("REM_DIAR") * (1 + gdblTasaIgv), "0.00")
                End If
                adoVar.Close: Set adoVar = Nothing
            End With
    
            s_Rec = ""
            s_Rec = s_Rec + "D"
            s_Rec = s_Rec + Format(strCodAdministradoraCNS, "@@@@@@")
            s_Rec = s_Rec + Format(s_CODFONDO, "@@@@")
            's_Rec = s_Rec + yyyymmdd(adoCuota("FechaCuota"), "yyyymmdd", "ddmmyyyy", res)
            
            'Movimiento de cuotas en el día
            s_Rec = s_Rec + Left$(Format$(CDbl(adoCuota("CantCuotaSuscripcionConocida")), "000000000000.00000000"), 12) + Right$(Format$(CDbl(adoCuota("CantCuotaSuscripcionConocida")), "000000000000.00000000"), 8)
            s_Rec = s_Rec + Left$(Format$(CDbl(Abs(adoCuota("CantCuotaRedencionConocida"))), "000000000000.00000000"), 12) + Right$(Format$(CDbl(Abs(adoCuota("CantCuotaRedencionConocida"))), "000000000000.00000000"), 8)
            s_Rec = s_Rec + Left$(Format$(CDbl(adoCuota("CantCuotaSuscripcionDesconocida")), "000000000000.00000000"), 12) + Right$(Format$(CDbl(adoCuota("CantCuotaSuscripcionDesconocida")), "000000000000.00000000"), 8)
            s_Rec = s_Rec + Left$(Format$(CDbl(Abs(adoCuota("CantCuotaRedencionDesconocida"))), "000000000000.00000000"), 12) + Right$(Format$(CDbl(Abs(adoCuota("CantCuotaRedencionDesconocida"))), "000000000000.00000000"), 8)
            s_Rec = s_Rec + Left$(Format$(CDbl(adoCuota("CantCuotaFinal")), "000000000000.00000000"), 12) + Right$(Format$(CDbl(adoCuota("CantCuotaFinal")), "000000000000.00000000"), 8)
            s_Rec = s_Rec + Left$(Format$(CDbl(adoCuota("CantParticipe")), "000000"), 6)
            
            'Remuneración diaria
            s_Rec = s_Rec + Left$(Format$(Abs(n_RemDIAR), "00000000.00"), 8) + Right$(Format$(Abs(n_RemDIAR), "00000000.00"), 2)
            
            'Porcentaje total de comisiones del día
            s_Rec = s_Rec + Left$(Format$(n_PorcCom, "00.00000000"), 2) + Right$(Format$(n_PorcCom, "00.00000000"), 8)

            'Ingresa tipo de cambio y valor cuota del día
            s_Rec = s_Rec + Left$(Format$(CDbl(adoCuota("ValorTipoCambio")), "00000.00000000"), 5) + Right$(Format$(CDbl(adoCuota("ValorTipoCambio")), "00000.00000000"), 8)
            s_Rec = s_Rec + Left$(Format$(CDbl(adoCuota("ValorCuotaFinal")), "00000000.00000000"), 8) + Right$(Format$(CDbl(adoCuota("ValorCuotaFinal")), "00000000.00000000"), 8)
            s_Rec = s_Rec + NL$
            s_out = s_Rec
            Put n_Out, , s_out
            n_Cont = n_Cont + 1
            adoCuota.MoveNext
        Loop
        adoCuota.Close
        adoFondos.MoveNext
    Loop

    '***************************
    '    Registro de Control
    '***************************
    s_Rec = ""
    s_Rec = s_Rec + "C"
    s_Rec = s_Rec + Format(Format(strCodAdministradoraCNS, "@@@@@@"))
    s_Rec = s_Rec + "VC"
    s_Rec = s_Rec + Right(s_fecha, 2) + Mid(s_fecha, 5, 2) + Left(s_fecha, 4)
    s_Rec = s_Rec + Format$(n_NroRepo, "0000")
    s_Rec = s_Rec + Format$(n_Cont, "00000")
    s_con = s_Rec
    
    Put n_Out, 1, s_con
    Close n_Out
    adoFondos.Close
    Set adoFondos = Nothing

    'Call ActualizaRegcontrol("VC", gstrCodAdministradora, n_NroRepo, n_Cont)

End Sub

Private Sub Genera_SC()
    
    Dim strSQL As String
    Dim n_Out As Integer, res As Integer
    Dim s_Espacio As String * 1
    Dim s_out As String * 43
    Dim s_con As String * 27 '(26 longitud del registro)
    Dim n_SALDO As Double, n_SLDCIER As Double
    Dim n_SLDPCIE As Double, s_Cuenta As String
    Dim s_CodCta As String
    Dim dn_Fondos As ADODB.Recordset
    Dim dn_plncta As ADODB.Recordset
    Dim n_Cont As Integer
    Dim s_Rec As String
    Dim s_TipBaln As String
    Dim s_FchSald As String
    Dim s_CodFond As String
    Dim s_TipEEFF As String
    Dim NL$
    Dim s_CodGcta As String, s_CodCom As String, s_SfiTemp As Double
    Dim adoFondos As ADODB.Recordset
    Dim adoVar As ADODB.Recordset
    Dim adoPlncta As ADODB.Recordset
    Dim adoCtaCnsv As ADODB.Recordset
    Dim adoCuota As ADODB.Recordset
    Dim s_CtaCns As String
    Dim s_FchAper As String
    Dim s_FlgCnsv As String
      
    
    n_Cont = 0
    s_CodGcta = ""
    s_CodCom = ""
    NL$ = Chr(13) + Chr(10)
    
    strSQL = "SELECT CodFondo,CodConasev,CodMoneda FROM Fondo "
    strSQL = strSQL + "WHERE CodAdministradora = gstrCodAdministradora ORDER BY CodConasev ASC"
    
    Set adoFondos = New ADODB.Recordset
    
    With adoComm
        .CommandText = strSQL
        Set adoFondos = .Execute
        
        If adoFondos.EOF Then
            MsgBox "No existen Fondos Mutuos disponibles para el proceso.", 48
            Exit Sub
        Else
            adoFondos.MoveLast
            adoFondos.MoveFirst
        End If
    End With
            
    s_FchInic = Year(dtpFechaDesde.Value) + Format(Month(dtpFechaDesde.Value), "00") + Format(Day(dtpFechaDesde.Value), "00")
    s_FchSald = dd_File + mm_File + aa_File
    
    'OBTENER NUMERO DE ARCHIVO
    'Call GeneraNroArchivoCnsv("SC", gstrCodAdministradora, s_fecha)
    
    n_Out = FreeFile
    Open txtDestino For Binary Access Read Write As n_Out

    'REGISTRO DE CONTROL
    s_Rec = ""
    s_Rec = s_Rec + "C"
    s_Rec = s_Rec + Format(strCodAdministradoraCNS, "@@@@@@")
    s_Rec = s_Rec + "SC"
    s_Rec = s_Rec + Right(s_fecha, 2) + Mid(s_fecha, 5, 2) + Left(s_fecha, 4)
    s_Rec = s_Rec + Format$(n_NroRepo, "0000")
    s_Rec = s_Rec + Format$(n_Cont, "00000")
    s_Rec = s_Rec + Format(Left$(s_Espacio, 15), "@@@@@@@@@@@@@@@!")
    s_Rec = s_Rec + NL$
    s_out = s_Rec
    Put n_Out, , s_out

    Do Until adoFondos.EOF
       
       ' PARAMETROS PARA LA GENERACION
       s_FonMut = adoFondos("CodFondo"): s_CODFONDO = Trim$(adoFondos("CodConasev"))
       
       'Determinar el inicio del periodo contable de la fecha seleccionada y el fondo
       strSQL = "SELECT FechaInicio FROM PeriodoContable WHERE "
       strSQL = strSQL + "CodFondo = '" + s_FonMut + "' AND "
       strSQL = strSQL + "MesContable = '00' AND "
       strSQL = strSQL + "PeriodoContable = '" + aa_File + "'"
       
       Set adoVar = New ADODB.Recordset
       
       With adoVar
            '.CommandText = strSQL
            'Set adoVar = .Execute
        
            'SI ES NULO CONSIDERO SOLO EL MES DE LA FECHA...
            's_FchAper = IIf(IsNull(adoVar("FechaInicio")), aa_File + mm_File + "01", adoVar("FechaInicio"))
             
            adoVar.Close
            Set adoVar = Nothing
        End With
       
       'strSQL = "SELECT CodCuenta, FLG_PCIE, TIP_EEFF, FLG_CNSV FROM PlanContable "
       strSQL = "SELECT CodCuenta, CodTipoEEFF, IndConasev FROM PlanContable "
       strSQL = strSQL + "WHERE IndConasev = 'X' "
       strSQL = strSQL + "ORDER BY CodCuenta "
       
       Set adoPlncta = New ADODB.Recordset
       
       With adoComm
            .CommandText = strSQL
            Set adoPlncta = .Execute
        
            If Not adoPlncta.EOF Then
                adoPlncta.MoveLast
                adoPlncta.MoveFirst
            End If
       End With
       
       Do Until adoPlncta.EOF
          
          s_CodCta = Trim$(adoPlncta("CodCuenta"))
          s_TipEEFF = adoPlncta("CodTipoEEFF")
          
          n_SLDCIER = 0: n_SLDPCIE = 0: n_SALDO = 0: s_Cuenta = ""
          s_TipBaln = "3"
       
          'Guarda el código del grupo de cuentas
          If s_CodGcta <> Left(s_CodCta, 1) Then
             s_CodGcta = Left(s_CodCta, 1)
             s_CodCom = ""
             
             strSQL = "SELECT CodCuenta FROM PlanContable WHERE "
             strSQL = strSQL + "CodCuenta LIKE '" + s_CodGcta + "%' AND "
             strSQL = strSQL + "IndConasev= 'X' "
             strSQL = strSQL + "ORDER BY CodCuenta "
             
             '*******************************************
             'FALTA MARCAR CON X EN LA TABLA PlanContable
             '*******************************************
             Set adoCtaCnsv = New ADODB.Recordset
             With adoComm
                .CommandText = strSQL
                Set adoCtaCnsv = .Execute
            
                If Not adoCtaCnsv.EOF Then
                    adoCtaCnsv.MoveLast
                    adoCtaCnsv.MoveFirst
                End If
            End With
             
            Do Until adoCtaCnsv.EOF
                s_CtaCns = s_CtaCns + "AND CodCuenta NOT LIKE '" + Trim(adoCtaCnsv("CodCuenta")) + "%' "
                adoCtaCnsv.MoveNext
             Loop
             adoCtaCnsv.Close
             Set adoCtaCnsv = Nothing
                      
             strSQL = "SELECT COALESCE(SUM(SaldoFinalContable),0) SFICONT FROM PartidaContableSaldos WHERE "
             strSQL = strSQL + "CodFondo = '" + s_FonMut + "' AND "
                          
             If s_TipEEFF = "1" Then  'SI LA CUENTA ES DE BALANCE
                strSQL = strSQL + "FechaSaldo = '" + s_FchInic + "' AND "
             Else                     'SI LA CUENTA ES DE RESULTADOS
                strSQL = strSQL + "FechaSaldo BETWEEN '" + s_FchAper + "' AND '" + s_FchInic + "' AND "
             End If
             strSQL = strSQL + "SaldoFinalContable <> 0 AND "
             strSQL = strSQL + "CodCuenta LIKE '" + s_CodGcta + "%' "
             strSQL = strSQL + s_CtaCns
             
             Set adoCtaCnsv = New ADODB.Recordset
             With adoComm
                .CommandText = strSQL
                Set adoCtaCnsv = .Execute
                
                s_SfiTemp = adoCtaCnsv("SaldoFinalContable")
                adoCtaCnsv.Close
                Set adoCtaCnsv = Nothing
            End With
             
          End If

          If IsNull(adoPlncta("IndConasev")) Then
             s_FlgCnsv = ""
          Else
             s_FlgCnsv = Trim(adoPlncta("IndConasev"))
          End If

          'If Trim$(dn_plncta!FLG_CNSV) = "X" Then
          If s_FlgCnsv = "X" Then
          
             '-------------------------------------------------------------------
             '1.- SELECCIONAR LOS SALDOS DE CIERRE AGRUPADOS POR CUENTA CONTABLE
             '-------------------------------------------------------------------
             strSQL = "SELECT COALESCE(SUM(SaldoFinalContable),0) SFICONT FROM PartidaContableSaldos WHERE "
             strSQL = strSQL + "CodFondo = '" + s_FonMut + "' AND "
             If s_TipEEFF = "1" Then  'SI LA CUENTA ES DE BALANCE
                strSQL = strSQL + "FechaSaldo = '" + s_FchInic + "' AND "
             Else                     'SI LA CUENTA ES DE RESULTADOS
                strSQL = strSQL + "FechaSaldo BETWEEN '" + s_FchAper + "' AND '" + s_FchInic + "' AND "
             End If
             strSQL = strSQL + "SaldoFinalContable <> 0 AND "
             strSQL = strSQL + "CodCuenta LIKE '" + Trim$(s_CodCta) + "%' "
             
             Set adoCuota = New ADODB.Recordset
             
             With adoComm
                .CommandText = strSQL
                Set adoCuota = .Execute
        
                If Not adoCuota.EOF Then
                    adoCuota.MoveLast
                    adoCuota.MoveFirst
                    n_SLDCIER = Format(adoCuota("SaldoFinalContable"), "0.00")
                End If
             End With
            
             adoCuota.Close
             Set adoCuota = Nothing
             
             '**************************************************************
             '**************************************************************
             'PREGUNTAR A ACR CUAL ES EL EQUIVALENTE DEL CAMPO FLG_PCIE
             '**************************************************************
             '**************************************************************
             If dn_plncta!FLG_PCIE = "X" Then
                '-------------------------------------------------------------------
                '2.- BUSCAR EL SALDO DE PRECIERRE DE LA CUENTA CORRESPONDIENTE
                '-------------------------------------------------------------------
                strSQL = "SELECT COALESCE(SUM(SaldoFinalContable),0) SFICONT FROM PartidaContablePreSaldos WHERE "
                strSQL = strSQL + "CodFondo = '" + s_FonMut + "' AND "
                strSQL = strSQL + "FechaSaldo = '" + s_FchInic + "' AND "
                strSQL = strSQL + "SaldoFinalContable <> 0 AND "
                strSQL = strSQL + "CodCuenta LIKE '" + Trim$(s_CodCta) + "%' "
                
                Set adoCuota = New ADODB.Recordset
             
                With adoComm
                    .CommandText = strSQL
                    Set adoCuota = .Execute
        
                    If Not adoCuota.EOF Then
                        n_SLDPCIE = Format(adoCuota("SaldoFinalContable"), "0.00")
                    End If
                End With
                adoCuota.Close
                Set adoCuota = Nothing
                
                s_TipBaln = "1"
                If n_SLDPCIE <> 0 Then
                   s_Rec = InsertarSaldosContables("D", strCodAdministradoraCNS, s_CODFONDO, s_FchSald, s_TipBaln, s_CodCta, n_SLDPCIE)
                   Put n_Out, , s_Rec
                   n_Cont = n_Cont + 1
                End If
                s_TipBaln = "2"
             End If
                                       
             If s_CodGcta = "1" Or s_CodGcta = "3" Or s_CodGcta = "5" Or s_CodGcta = "8" Then
                If Mid(Trim$(s_CodCta), 2, 1) = "9" And s_CodCom = "" Then
                   s_CodCom = Trim$(s_CodCta)
                   n_SLDCIER = n_SLDCIER + s_SfiTemp
                End If
             Else
                If s_CodGcta = "6" And s_CodCom = "" And Trim$(s_CodCta) = "679" Then
                   s_CodCom = Trim$(s_CodCta)
                   n_SLDCIER = n_SLDCIER + s_SfiTemp
                End If
                If s_CodGcta = "7" And s_CodCom = "" And Trim$(s_CodCta) = "779" Then
                   s_CodCom = Trim$(s_CodCta)
                   n_SLDCIER = n_SLDCIER + s_SfiTemp
                End If
                If s_CodGcta = "4" And Mid(Trim$(s_CodCta), 3, 1) = "9" And s_CodCom = "" Then
                   s_CodCom = Trim$(s_CodCta)
                   n_SLDCIER = n_SLDCIER + s_SfiTemp
                End If
             End If
             
             If n_SLDCIER <> 0 Then
                s_Rec = InsertarSaldosContables("D", strCodAdministradoraCNS, s_CODFONDO, s_FchSald, s_TipBaln, s_CodCta, n_SLDCIER)
                Put n_Out, , s_Rec
                n_Cont = n_Cont + 1
             End If

          End If
          adoPlncta.MoveNext
       Loop
       
       adoPlncta.Close
       adoFondos.MoveNext
    
    Loop
    
    'REGISTRO DE CONTROL
    s_Rec = ""
    s_Rec = s_Rec + "C"
    s_Rec = s_Rec + Format(strCodAdministradoraCNS, "@@@@@@")
    s_Rec = s_Rec + "SC"
    s_Rec = s_Rec + Right(s_fecha, 2) + Mid(s_fecha, 5, 2) + Left(s_fecha, 4)
    s_Rec = s_Rec + Format(n_NroRepo, "0000")
    s_Rec = s_Rec + Format$(n_Cont, "00000")
    s_con = s_Rec
    
    Put n_Out, 1, s_con
    Close #1
    adoFondos.Close

    'Call ActualizaRegcontrol("SC", gstrCodAdministradora, n_NroRepo, n_Cont)

    Me.MousePointer = 0
'    MsgBox "Archivo de Saldos Contables generado Exitosamente.", 48

End Sub


Private Function GeneraNumRegistroControl(strTipo As String, strCodAdministradora As String, strFecha As String)

    Dim adoNroReporte As ADODB.Recordset
    Dim strSQL As String
    Dim strNumRegistro As String
    
    Set adoNroReporte = New ADODB.Recordset
    
    GeneraNumRegistroControl = ""
    
    strSQL = "{ call up_CNObtenerNumRegistroControl ('" + strTipo + "', '" + strCodAdministradora + "','" + strFecha + "') }"
    
    With adoComm
        .CommandText = strSQL
        Set adoNroReporte = .Execute
          
        strNumRegistro = adoNroReporte("NumReporte")
           
        adoNroReporte.Close: Set adoNroReporte = Nothing
    End With
    
    GeneraNumRegistroControl = strNumRegistro

End Function

Private Sub ActualizaRegistroControl(strCodArchivo As String, strCodAdministradora As String, strFechaReporte As String, numNumReporte As Long, numCantRegistros As Long)

Dim res As Integer
Dim strSQL As String
Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset

    strSQL = "{ call up_CNActualizarRegistroControl ('" & strCodArchivo & "','"
    strSQL = strSQL & strCodAdministradora & "','"
    strSQL = strSQL & strFechaReporte & "'," & numNumReporte & "," & numCantRegistros & ") }"
    
    With adoComm
        .CommandText = strSQL
        .Execute
    End With

End Sub

Private Function InsertarSaldosContables(s_TipRegi As String, s_CodSAFM As String, s_CodFond As String, s_FchSald As String, s_TipBaln As String, s_CodCta As String, n_ValSald As Double) As String
Dim s_Rec As String

    s_Rec = ""
    s_Rec = s_Rec + s_TipRegi
    s_Rec = s_Rec + Format(s_CodSAFM, "@@@@@@")
    s_Rec = s_Rec + Format(s_CodFond, "@@@@")
    s_Rec = s_Rec + s_FchSald
    s_Rec = s_Rec + s_TipBaln    'Tipo de Balance: 1:Precierre, 2:Cierre y 3:Ambos
    
    s_CodCta = Trim(s_CodCta)

    If s_CodCta = "572" Then
       s_CodCta = "571"
    End If

    If s_CodCta = "18242" Or s_CodCta = "18243" Or s_CodCta = "19242" Or s_CodCta = "19243" Then
       s_CodCta = Left(s_CodCta, 4) + "0" + Right(s_CodCta, 1)
    End If

    s_Rec = s_Rec + Format(Left$(s_CodCta, 6), "@@@@@@!")
    If n_ValSald >= 0 Then
        s_Rec = s_Rec + "+" + Left$(Format$(n_ValSald, "000000000000.00"), 12) + Right$(Format$(n_ValSald, "000000000000.00"), 2)
    Else
        s_Rec = s_Rec + "-" + Left$(Format$(Abs(n_ValSald), "000000000000.00"), 12) + Right$(Format$(Abs(n_ValSald), "000000000000.00"), 2)
    End If
    s_Rec = s_Rec + NL$
    
    InsertarSaldosContables = s_Rec

End Function


'Private Sub Genera_PSaammdd()
'    '----------------------------------------------------------------
'    ' GENERACION DEL ARCHIVO : PS_ddmm.iii
'    ' ARCHIVO DE PERSONAS
'    ' EL ARCHIVO ES UNICO PARA LOS TRES FONDOS MUTUOS DE LA SAFM
'    '----------------------------------------------------------------
'    Dim sensql As String
'    Dim n_Out, res As Integer
'    Dim s_Espacio As String * 1
'    Dim s_out As String * 166 '(164 longitud del registro)
'    Dim dn_Fondos As ADODB.Recordset, dn_cuota As ADODB.Recordset, dn_person As ADODB.Recordset
'    Dim s_con As String
'
'    PnlMsg.Caption = "Generando Archivo de Personas del " + s_FchRepoWin
'
'
'    adoComm.CommandText = " up_ACParametroCONASEV (1,'" & strCodFondo & "','" & gstrCodAdministradora & "','','')"
'    Set adoRegistro = .Execute
'
'
'    sensql = "EXEC SP_R_ObtenerNroArchivoCnsv 'PS', '" + s_fecha + "' "
'
'    Set Dn_Var = GBdFm.CreateDynaset(sensql, 64)
'    n_NroRepo = Dn_Var!nro_repo
'
'    Dn_Var.Close
'
'    sensql = "SELECT COD_FOND, COD_CNSV FROM FMFONDOS "
'    sensql = sensql + "WHERE COD_TFON = '0' ORDER BY COD_CNSV ASC"
'    Set dn_Fondos = GBdFm.CreateDynaset(sensql, 64)
'    If dn_Fondos.EOF Then
'       MsgBox "No existen Fondos Mutuos disponibles para el proceso.", 48
'       Exit Sub
'    End If
'
'    'Borra el archivo si existe
'    n_Out = FreeFile
'    Open txt_NomFil For Binary Access Read Write As n_Out
'
'    dn_Fondos.MoveLast
'    dn_Fondos.MoveFirst
'    Do Until dn_Fondos.EOF
'       '-------------------------------------------------------------------
'       ' PARAMETROS PARA LA GENERACION
'       '-------------------------------------------------------------------
'       s_FonMut = dn_Fondos!COD_FOND: s_CODFONDO = Trim$(dn_Fondos!COD_CNSV)
'
'       sensql = "SELECT * FROM FMCUOTAS WHERE "
'       sensql = sensql + "COD_FOND='" + s_FonMut + "' AND "
'       sensql = sensql + "FLG_CIER='X' AND "
'       sensql = sensql + "FCH_CUOT='" + s_fecha + "' "
'       Set dn_cuota = GBdFm.CreateDynaset(sensql, 64)
'
'       If dn_cuota.EOF Then
'          MsgBox "Fecha del Fondo no Disponible para el Proceso.", 48
'          Dat_fecini.SetFocus
'          dn_cuota.Close
'          Exit Sub
'       End If
'       dn_cuota.Close
'       dn_Fondos.MoveNext
'    Loop
'    dn_Fondos.Close
'    Me.MousePointer = 11
'
'    'Sensql = "EXEC SP_R_ArchivoPersonas '" + s_fecha + "' "
'    sensql = "EXEC SP_R_ArchivoPersonas '" + s_fecha + "','" + s_fecha + "'"
'
'    Set dn_person = GBdFm.CreateDynaset(sensql, 64)
'
'    If Not dn_person.EOF Then
'       dn_person.MoveLast
'       dn_person.MoveFirst
'    End If
'
'    n_CntRegi = dn_person.RecordCount
'
'    '***************************
'    '    Registro de Control
'    '***************************
'
'    s_Rec = ""
'    s_Rec = s_Rec + "C"
'    s_Rec = s_Rec + Format(g_CODSAFM, "@@@@@@")
'    s_Rec = s_Rec + "PS"
'    s_Rec = s_Rec + Right(s_fecha, 2) + Mid(s_fecha, 5, 2) + Left(s_fecha, 4)
'    s_Rec = s_Rec + Format$(n_NroRepo, "0000")
'    s_Rec = s_Rec + Format$(n_CntRegi, "00000")
'    s_Rec = s_Rec + NL$
'    s_con = s_Rec
'    Put n_Out, , s_con
'
'    Do Until dn_person.EOF
'       s_Rec = ""
'       s_Rec = s_Rec + "D"
'       s_Rec = s_Rec + Format(g_CODSAFM, "@@@@@@")
'       s_Rec = s_Rec + Format(dn_person!COD_NACI, "@@@")
'       s_Rec = s_Rec + Format(dn_person!TIP_IDEN, "@@")
'       If dn_person!TIP_IDEN = "03" Then  'RUC
'          s_Rec = s_Rec + Format(Left$(dn_person!NRO_IDEN, 11), "@@@@@@@@@@@!") + "0"
'       Else
'          s_Rec = s_Rec + Format(Left$(dn_person!NRO_IDEN, 10), "@@@@@@@@@@!") + "00"
'       End If
'       s_Rec = s_Rec + Format(Left$(dn_person!app_pers, 30), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
'       s_Rec = s_Rec + Format(Left$(dn_person!apm_pers, 30), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
'       s_Rec = s_Rec + Format(Left$(dn_person!nom_pers, 30), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
'       s_Rec = s_Rec + Format(Left$(dn_person!RAZ_SOCI, 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
'       s_Rec = s_Rec + NL$
'       s_out = s_Rec
'       Put n_Out, , s_out
'       dn_person.MoveNext
'    Loop
'    Close n_Out
'    dn_person.Close
'
'    Call ActualizaRegcontrol("PS", n_NroRepo, n_CntRegi)
'
'    Me.MousePointer = 0
''    MsgBox "Archivo de Personas generado Exitosamente.", 48
'
'End Sub

Private Sub lst_DscRepo_Click()

End Sub
