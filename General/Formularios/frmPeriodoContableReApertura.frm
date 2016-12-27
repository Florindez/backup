VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPeriodoContableReApertura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reapertura de Periodo Contable"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6735
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6735
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
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
      Left            =   5100
      Picture         =   "frmPeriodoContableReApertura.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4170
      Width           =   1200
   End
   Begin VB.ComboBox cboFondo 
      Height          =   315
      Left            =   1275
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1260
      Width           =   4995
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
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
      Left            =   3600
      Picture         =   "frmPeriodoContableReApertura.frx":0562
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4170
      Width           =   1200
   End
   Begin VB.Frame fraControlfecha 
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
      Height          =   3885
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.ComboBox cboPeriodoContableApertura 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2880
         Width           =   3585
      End
      Begin MSComCtl2.DTPicker dtpFechaActual 
         Height          =   315
         Left            =   2640
         TabIndex        =   5
         Top             =   2220
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   175898625
         CurrentDate     =   40830
      End
      Begin VB.ComboBox cboPeriodoContableActual 
         Height          =   315
         Left            =   2640
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1800
         Width           =   3585
      End
      Begin MSComCtl2.DTPicker dtpFechaApertura 
         Height          =   315
         Left            =   2640
         TabIndex        =   8
         Top             =   3300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   175898625
         CurrentDate     =   40830
      End
      Begin VB.Label lblDescrip 
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
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   1515
      End
      Begin VB.Line Line4 
         X1              =   360
         X2              =   6240
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   6210
         Y1              =   2670
         Y2              =   2670
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Periodo Contable Apertura"
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
         Height          =   225
         Index           =   3
         Left            =   330
         TabIndex        =   11
         Top             =   2910
         Width           =   2265
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fecha Apertura"
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
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   3330
         Width           =   1515
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fecha Actual"
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
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   2310
         Width           =   1515
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Periodo Contable Actual"
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
         Height          =   225
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   1830
         Width           =   2175
      End
      Begin VB.Label lblFondo 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   300
         TabIndex        =   1
         Top             =   290
         Width           =   5925
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   5850
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5850
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmPeriodoContableReApertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFechaOriginal As String
Dim arrPeriodo() As String, arrFondo() As String, arrPeriodoReApertura() As String
Dim strCodFondo As String
Dim strPeriodoActual As String
Dim strMesActual As String
Dim strPeriodoReApertura As String
Dim strMesReApertura As String
Dim strFechaReApertura As String
Dim strMesAnterior As String

Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset
    
    Dim strSQL      As String, intRegistro As Integer
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    lblFondo.Caption = cboFondo.Text
    
    Set adoRegistro = New ADODB.Recordset
           
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            
            strFechaOriginal = adoRegistro("FechaCuota")
            gdblTipoCambio = 1 'adoRegistro("ValorTipoCambio")
            gstrCodMoneda = adoRegistro("CodMoneda")
            
            dtpFechaActual.Value = strFechaOriginal
            
            '*** Periodo Actual ***
            strSQL = "{ call up_CNSelPeriodoContableVigente('" & strCodFondo & "','" & gstrCodAdministradora & "') }"
            CargarControlLista strSQL, cboPeriodoContableActual, arrPeriodo(), ""
           
            If cboPeriodoContableActual.ListCount > 0 Then cboPeriodoContableActual.ListIndex = 0
            
            '*** Periodos candidatos a reapertura ***
            strSQL = "{ call up_CNSelPeriodoContableReApertura('" & strCodFondo & "','" & gstrCodAdministradora & "','" & strPeriodoActual & "','" & strMesActual & "' ) }"
            CargarControlLista strSQL, cboPeriodoContableApertura, arrPeriodoReApertura(), ""
           
            intRegistro = ObtenerItemLista(arrPeriodoReApertura(), strPeriodoActual + strMesAnterior)
            If intRegistro >= 0 Then cboPeriodoContableApertura.ListIndex = intRegistro
            
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)

End Sub


Private Sub cboPeriodoContableActual_Click()

    strPeriodoActual = Valor_Caracter
    If cboPeriodoContableActual.ListIndex < 0 Then Exit Sub
    
    strPeriodoActual = Mid(Trim(arrPeriodo(cboPeriodoContableActual.ListIndex)), 1, 4)
    strMesActual = Mid(Trim(arrPeriodo(cboPeriodoContableActual.ListIndex)), 5, 2)
    
    If strMesActual = "01" Then Exit Sub
    
    strMesAnterior = Mid(CStr((CInt(Mid(Trim(arrPeriodo(cboPeriodoContableActual.ListIndex)), 5, 2)) - 1) + 100), 2, 2)

End Sub

Private Sub cboPeriodoContableApertura_Click()

    Dim adoRegistro As ADODB.Recordset, strSQL As String
    
    Set adoRegistro = New ADODB.Recordset
    
    strPeriodoReApertura = Valor_Caracter
    If cboPeriodoContableApertura.ListIndex < 0 Then Exit Sub
    
    strPeriodoReApertura = Mid(Trim(arrPeriodoReApertura(cboPeriodoContableApertura.ListIndex)), 1, 4)
    strMesReApertura = Mid(Trim(arrPeriodoReApertura(cboPeriodoContableApertura.ListIndex)), 5, 2)
       
    strFechaReApertura = DateSerial(CInt(strPeriodoReApertura), CInt(strMesReApertura), UltimoDiaMes(CInt(strMesReApertura), CInt(strPeriodoReApertura)))
       
'    dtpFechaApertura.MinDate = strFechaReApertura
'    dtpFechaApertura.MaxDate = strFechaReApertura
         dtpFechaApertura.Value = strFechaReApertura
       
    strSQL = "{ call up_CNSelPeriodoContableFecha ('" & _
                        strCodFondo & "','" & _
                        gstrCodAdministradora & "','" & _
                        strPeriodoReApertura & "','" & _
                        strMesReApertura & "') }"
    
    With adoRegistro
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
        
        If Not .EOF Then
            strFechaReApertura = adoRegistro.Fields("FechaFinal").Value
'            dtpFechaApertura.MinDate = adoRegistro.Fields("FechaFinal").Value
'            dtpFechaApertura.MaxDate = adoRegistro.Fields("FechaFinal").Value
            
            dtpFechaApertura.Value = adoRegistro.Fields("FechaFinal").Value

        End If
        
        .Close
        
    End With
    
    Set adoRegistro = Nothing
        
    dtpFechaApertura.Value = strFechaReApertura
        
    
End Sub

Private Sub cmdAceptar_Click()


    Dim adoFondo As ADODB.Recordset, strSQL As String
    Dim strFechaNueva As String
    Dim strFrecuenciaValorizacionMensual    As String
        
    strFechaNueva = dtpFechaApertura.Value
    
     Set adoFondo = New ADODB.Recordset
    
    
    strFrecuenciaValorizacionMensual = "05"
    
    With adoComm
    
        .CommandText = " SELECT FrecuenciaValorizacion,TipoFondo FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "'"
        Set adoFondo = .Execute
        
         If adoFondo("FrecuenciaValorizacion") = strFrecuenciaValorizacionMensual And adoFondo("TipoFondo") = Administradora_Fondos Then
                If strFechaNueva <> strFechaOriginal Or (strFechaNueva = strFechaOriginal And (strMesActual = "99" Or strMesActual = "00")) Then
                    'Cambia la fecha
                    
                    If MsgBox("Desea re-aperturar el mes " & strMesReApertura & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            
                        '*** Procesa el cambio de fecha ***
                        adoComm.CommandText = "{ call up_CNProcControlFechaReApertura('" & _
                            strCodFondo & "','" & gstrCodAdministradora & "','" & _
                            Convertyyyymmdd(strFechaOriginal) & "','" & _
                            Convertyyyymmdd(strFechaNueva) & "','" & strMesActual & "') }"
                        adoConn.Execute adoComm.CommandText
                        
                        gdatFechaActual = strFechaNueva
                        gstrFechaActual = Convertyyyymmdd(strFechaNueva)
                        
                        gstrPeriodoActual = CStr(Year(gdatFechaActual))
                        gstrMesActual = Mid(CStr(Month(gdatFechaActual) + 100), 2, 2)
                        
                        frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
                        
                        Unload Me
                   
                    End If
                            
                Else
                    MsgBox "Las fechas no tienen que ser iguales", vbInformation + vbOKOnly, Me.Caption
                End If
          Else
          
                'MsgBox "No se puede aplicar el siguiento Proceso a un Fondo con Frecuencia de Valorización Diaria", vbCritical, Me.Caption
                  MsgBox "No se puede aplicar el siguiento Proceso a un Tipo de Fondo diferente a Administradora de Fondos", vbCritical, Me.Caption
          End If
    End With

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call DarFormato

    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
 

End Sub

Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strFechaOriginal = Valor_Caracter
    CentrarForm Me
  
End Sub

Private Sub CargarListas()

    Dim strSQL  As String
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
     CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
        
End Sub
Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
            
End Sub


