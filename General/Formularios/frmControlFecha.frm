VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmControlFecha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Fechas"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6540
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6540
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
      Left            =   4860
      Picture         =   "frmControlFecha.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2730
      Width           =   1200
   End
   Begin VB.ComboBox cboFondo 
      Height          =   315
      Left            =   1130
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1040
      Width           =   5115
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
      Left            =   3390
      Picture         =   "frmControlFecha.frx":0562
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2730
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
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin MSComCtl2.DTPicker dtpFechaActual 
         Height          =   315
         Left            =   2640
         TabIndex        =   5
         Top             =   2130
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   146669569
         CurrentDate     =   40830
      End
      Begin VB.ComboBox cboPeriodoContableActual 
         Height          =   315
         Left            =   2640
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1680
         Width           =   3585
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
         Height          =   225
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   6240
         Y1              =   1520
         Y2              =   1520
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
         Left            =   420
         TabIndex        =   3
         Top             =   2160
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
         Top             =   1710
         Width           =   1515
      End
      Begin VB.Label lblFondo 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ARICSA"
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
         Top             =   360
         Width           =   5925
      End
   End
End
Attribute VB_Name = "frmControlFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFechaOriginal As String
Dim arrPeriodo() As String, arrFondo() As String
Dim strCodFondo As String
Dim strPeriodoActual As String
Dim strMesActual As String
Dim strFechaActual As String


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
            
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)


End Sub


Private Sub cboPeriodoContableActual_Click()

    'SETEANDO LOS LIMITES DE FECHA DEL PERIODO
    Dim adoRegistro As ADODB.Recordset, strSQL As String
    
    Set adoRegistro = New ADODB.Recordset
    
    strPeriodoActual = Valor_Caracter
    If cboPeriodoContableActual.ListIndex < 0 Then Exit Sub
    
    strPeriodoActual = Mid(Trim(arrPeriodo(cboPeriodoContableActual.ListIndex)), 1, 4)
    strMesActual = Mid(Trim(arrPeriodo(cboPeriodoContableActual.ListIndex)), 5, 2)
       
    dtpFechaActual.MinDate = strFechaOriginal
    dtpFechaActual.MaxDate = strFechaOriginal
       
    strSQL = "{ call up_CNSelPeriodoContableFecha ('" & _
                        strCodFondo & "','" & _
                        gstrCodAdministradora & "','" & _
                        strPeriodoActual & "','" & _
                        strMesActual & "') }"
    
    With adoRegistro
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
        
        If Not .EOF Then
            strFechaActual = adoRegistro.Fields("FechaFinal").Value
            dtpFechaActual.MinDate = adoRegistro.Fields("FechaInicio").Value
            dtpFechaActual.MaxDate = adoRegistro.Fields("FechaFinal").Value
        End If
        
        .Close
        
    End With
    
    Set adoRegistro = Nothing
        

End Sub

Private Sub cmdAceptar_Click()

    Dim adoFondo As ADODB.Recordset, strSQL As String
    Dim strFechaNueva As String
    Dim strFrecuenciaValorizacionMensual    As String
    
    
     strFechaNueva = dtpFechaActual.Value
    
     Set adoFondo = New ADODB.Recordset
    
    
    strFrecuenciaValorizacionMensual = "05"
    
    With adoComm
    
        .CommandText = " SELECT FrecuenciaValorizacion FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "'"
        Set adoFondo = .Execute
        
        If adoFondo("FrecuenciaValorizacion") = strFrecuenciaValorizacionMensual Then
    
               If strFechaNueva <> strFechaOriginal Then
                   'Cambia la fecha
                   
                   If MsgBox("Desea cambiar la fecha actual del " & strFechaOriginal & " al " & strFechaNueva & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                       
                       Me.MousePointer = vbHourglass
                       
                       '*** Procesa el cambio de fecha ***
                       adoComm.CommandText = "{ call up_CNProcControlFecha('" & _
                           strCodFondo & "','" & gstrCodAdministradora & "','" & _
                           Convertyyyymmdd(strFechaOriginal) & "','" & _
                           Convertyyyymmdd(strFechaNueva) & "') }"
                       adoConn.Execute adoComm.CommandText
                       
                       Me.MousePointer = vbDefault
                       
                   End If
                   
               End If
               
               gdatFechaActual = strFechaNueva
               gstrFechaActual = Convertyyyymmdd(strFechaNueva)
               
               gstrPeriodoActual = CStr(Year(gdatFechaActual))
               gstrMesActual = Mid(CStr(Month(gdatFechaActual) + 100), 2, 2)
               
               frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
               
               Unload Me
        Else

                 MsgBox "No se puede aplicar el siguiento Proceso a un Fondo con Frecuencia de Valorización Diaria", vbCritical, Me.Caption

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


