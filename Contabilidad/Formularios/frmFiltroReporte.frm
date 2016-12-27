VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmFiltroReporte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtro"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
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
      Left            =   4200
      Picture         =   "frmFiltroReporte.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4110
      Width           =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
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
      Left            =   2670
      Picture         =   "frmFiltroReporte.frx":0562
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4110
      Width           =   1200
   End
   Begin VB.Frame fraFiltroReporte 
      Caption         =   "Filtro del Reporte"
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
      Height          =   3945
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   7695
      Begin VB.TextBox txtNumAsiento 
         Height          =   885
         Left            =   3450
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   2850
         Width           =   3645
      End
      Begin VB.TextBox txtCodCuenta 
         Height          =   1245
         Left            =   3450
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1410
         Width           =   3645
      End
      Begin VB.ComboBox cboMonedaContable 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3450
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   450
         Width           =   3645
      End
      Begin VB.CommandButton cmdBusquedaCuenta 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7110
         TabIndex        =   8
         Top             =   1410
         Width           =   405
      End
      Begin VB.CheckBox chkOpcionFiltro 
         Caption         =   "Por Numero de Asiento"
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
         Height          =   345
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Top             =   2880
         Value           =   1  'Checked
         Width           =   2445
      End
      Begin VB.CheckBox chkOpcionFiltro 
         Caption         =   "Por Cuenta(s) Especifica(s)"
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
         Height          =   345
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   1380
         Width           =   2955
      End
      Begin VB.CheckBox chkOpcionFiltro 
         Caption         =   "Por Rango de Fechas"
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
         Height          =   345
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   900
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtpFechaInicial 
         Height          =   315
         Left            =   3840
         TabIndex        =   1
         Top             =   930
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   50397185
         CurrentDate     =   38068
      End
      Begin MSComCtl2.DTPicker dtpFechaFinal 
         Height          =   315
         Left            =   5700
         TabIndex        =   2
         Top             =   930
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   50397185
         CurrentDate     =   38068
      End
      Begin VB.Label lblMonedaContableReporte 
         Caption         =   "Moneda Contable"
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
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Width           =   1680
      End
      Begin VB.Label lblRango 
         Caption         =   "Al"
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
         Index           =   1
         Left            =   5340
         TabIndex        =   4
         Top             =   960
         Width           =   570
      End
      Begin VB.Label lblRango 
         Caption         =   "Del"
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
         Left            =   3450
         TabIndex        =   3
         Top             =   960
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmFiltroReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrMonedaContable()         As String
Public strCodFondo              As String
Public strCodAdministradora     As String


Private Sub cboMonedaContable_Click()

    gstrCodMonedaReporte = Valor_Caracter
    If cboMonedaContable.ListIndex < 0 Then Exit Sub
    
    gstrCodMonedaReporte = Trim(arrMonedaContable(cboMonedaContable.ListIndex))

End Sub

Private Sub chkOpcionFiltro_Click(Index As Integer)

    Select Case Index
    
        Case 0

            If chkOpcionFiltro(Index).Value = vbChecked Then
                dtpFechaInicial.Value = gdatFechaActual
                dtpFechaFinal.Value = gdatFechaActual
                dtpFechaInicial.Visible = True
                dtpFechaFinal.Visible = True
                lblRango(0).Visible = True
                lblRango(1).Visible = True
            
                chkOpcionFiltro(2).Value = vbUnchecked
               ' Call chkOpcionFiltro_Click(2)
            
            Else
                dtpFechaInicial.Visible = False
                dtpFechaFinal.Visible = False
                lblRango(0).Visible = False
                lblRango(1).Visible = False
            
                chkOpcionFiltro(2).Value = vbChecked
                'Call chkOpcionFiltro_Click(2)
            
            End If
       
       
       Case 1

            If chkOpcionFiltro(Index).Value = vbUnchecked Then
                txtCodCuenta.Text = ""
                cmdBusquedaCuenta.Enabled = False
                txtCodCuenta.Enabled = False
            Else
                txtCodCuenta.Enabled = True
                cmdBusquedaCuenta.Enabled = True
            End If
    
       
       Case 2
            If chkOpcionFiltro(Index).Value = vbUnchecked Then
                txtNumAsiento.Text = ""
                txtNumAsiento.Enabled = False
                
                chkOpcionFiltro(0).Value = vbChecked
               ' Call chkOpcionFiltro_Click(0)
            Else
                txtNumAsiento.Enabled = True
            
                chkOpcionFiltro(0).Value = vbUnchecked
               ' Call chkOpcionFiltro_Click(0)
            End If
            
            

    End Select



End Sub

Private Sub cmdAceptar_Click()
      
    Dim strFchDe As String, strFchAl As String
    
    If ValidaFormulario() Then
        strFchDe = CStr(dtpFechaInicial.Value)
        strFchAl = CStr(dtpFechaFinal.Value)
        gstrSelFrml = strtran(gstrSelFrml, "Fch1", strFchDe)
        gstrSelFrml = strtran(gstrSelFrml, "Fch2", strFchAl)
        gstrFchDel = CStr(dtpFechaInicial.Value)
        gstrFchAl = CStr(dtpFechaFinal.Value)
              
        If chkOpcionFiltro(1).Value = vbUnchecked Then
            gstrCodCuenta = "%"
        Else
            gstrCodCuenta = Trim(txtCodCuenta.Text)
        End If
               
        If chkOpcionFiltro(2).Value = vbUnchecked Then
            gstrNumAsiento = ""
        Else
            gstrNumAsiento = Trim(txtNumAsiento.Text)
        End If
        
        Unload Me
        DoEvents
        
    End If

End Sub

Private Sub cmdBusquedaCuenta_Click()

    gstrFormulario = "frmFiltroReporte"
    frmBusquedaCuenta.Show vbModal

End Sub

Private Sub cmdCancelar_Click()

    gstrSelFrml = "0"  '** Cancela
    Unload Me

End Sub

Private Sub Form_Load()


    Dim strSQL As String

    dtpFechaInicial.Value = gdatFechaActual
    dtpFechaFinal.Value = gdatFechaActual
    gstrFchDel = Valor_Caracter: gstrFchAl = Valor_Caracter
    If Trim(gstrSelFrml) = Valor_Caracter Then gstrSelFrml = "0"
    gindMonedaContable = Valor_Caracter
    
    strSQL = "{ call up_ACSelDatosParametro('70','" & strCodFondo & "','" & strCodAdministradora & "') }"
    CargarControlLista strSQL, cboMonedaContable, arrMonedaContable(), Valor_Caracter
    If cboMonedaContable.ListCount > 0 Then cboMonedaContable.ListIndex = 0


End Sub

Function ValidaFormulario() As Integer
    
    Dim intlOk As Integer
    Dim strMsg As String
    Dim r As Integer
    
    strMsg = ""
    intlOk = False


    If chkOpcionFiltro(1).Value = vbChecked And Trim(txtCodCuenta.Text) = "" Then
        MsgBox "Debe seleccionar la cuenta!", vbExclamation + vbOKOnly, Me.Caption
        GoTo ErrFicha
    End If
 
    If chkOpcionFiltro(0).Value = vbChecked Then
        If Not IsDate(dtpFechaInicial.Value) Then
            strMsg = dtpFechaInicial & " no es una fecha valida."
            GoTo ErrFicha
        End If
    
        If Not IsDate(dtpFechaFinal.Value) Then
            strMsg = dtpFechaFinal.Value & " no es una fecha valida."
            GoTo ErrFicha
        End If
    
        If DateDiff("d", dtpFechaInicial.Value, dtpFechaFinal.Value) < 0 Then
            strMsg = "Fecha Final debe ser posterior a " & dtpFechaInicial.Value & "."
            GoTo ErrFicha
        End If
        
    End If
    
    intlOk = True
    
ErrFicha:
    If strMsg <> "" Then r = MsgBox(strMsg, 0)
    ValidaFormulario = intlOk
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    If gstrFchDel = Valor_Caracter Then gstrSelFrml = "0"
    Set frmRangoFecha = Nothing
End Sub
