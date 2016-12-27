VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRangoFecha 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rango de Selección"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRango 
      Caption         =   "Fechas"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
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
         Left            =   360
         Picture         =   "frmRangoFecha.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   1200
      End
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
         Left            =   1890
         Picture         =   "frmRangoFecha.frx":0485
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker dtpFechaInicial 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   480
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   52101121
         CurrentDate     =   38068
      End
      Begin MSComCtl2.DTPicker dtpFechaFinal 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   960
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   52101121
         CurrentDate     =   38068
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
         Left            =   630
         TabIndex        =   6
         Top             =   495
         Width           =   570
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
         Left            =   630
         TabIndex        =   5
         Top             =   990
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmRangoFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim intNemotecnicoInd As Integer
'Dim strNemotecnicoVal As String
Function ValidaFormulario() As Integer
    
    Dim intlOk As Integer
    Dim strMsg As String
    Dim r As Integer
    
    strMsg = ""
    intlOk = False

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
    intlOk = True
    
ErrFicha:
    If strMsg <> "" Then r = MsgBox(strMsg, 0)
    ValidaFormulario = intlOk
    
End Function

Private Sub cmdAceptar_Click()

    Dim strFchDe As String, strFchAl As String
        
    If ValidaFormulario() Then
       strFchDe = CStr(dtpFechaInicial.Value)
       strFchAl = CStr(dtpFechaFinal.Value)
       gstrSelFrml = strtran(gstrSelFrml, "Fch1", strFchDe)
       gstrSelFrml = strtran(gstrSelFrml, "Fch2", strFchAl)
       gstrFchDel = CStr(dtpFechaInicial.Value)
       gstrFchAl = CStr(dtpFechaFinal.Value)
       Unload Me
       DoEvents
    End If
  
    
End Sub

Private Sub cmdCancelar_Click()
    
    gstrSelFrml = "0"  '** Cancela
    Unload Me

End Sub

Private Sub dtpFechaInicial_Change()

    'dtpFechaFinal.Value = dtpFechaInicial.Value

End Sub

Private Sub Form_Load()
        
    dtpFechaInicial.Value = gdatFechaActual
    dtpFechaFinal.Value = gdatFechaActual
    gstrFchDel = Valor_Caracter: gstrFchAl = Valor_Caracter
    If Trim(gstrSelFrml) = Valor_Caracter Then gstrSelFrml = "0"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If gstrFchDel = Valor_Caracter Then gstrSelFrml = "0"
    Set frmRangoFecha = Nothing
    
End Sub

