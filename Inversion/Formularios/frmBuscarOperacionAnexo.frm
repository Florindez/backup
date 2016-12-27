VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBuscarOperacionAnexo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buscar Registro"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   17985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   17985
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option1 
      Caption         =   "Buscar Flujo"
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
      Left            =   4110
      TabIndex        =   8
      Top             =   180
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame frmBusqueda 
      Height          =   555
      Left            =   180
      TabIndex        =   5
      Top             =   4830
      Width           =   6195
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1140
         TabIndex        =   7
         Top             =   180
         Width           =   4965
      End
      Begin VB.Label lblBusqueda 
         Caption         =   "Búsqueda"
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
         Left            =   150
         TabIndex        =   6
         Top             =   210
         Width           =   915
      End
   End
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
      Height          =   465
      Left            =   16380
      TabIndex        =   4
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "Seleccionar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   14790
      TabIndex        =   3
      Top             =   4920
      Width           =   1455
   End
   Begin VB.OptionButton rbBuscarAnexo 
      Caption         =   "Buscar Anexo"
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
      Left            =   2310
      TabIndex        =   2
      Top             =   180
      Width           =   1575
   End
   Begin VB.OptionButton rbBuscarOperacion 
      Caption         =   "Buscar Operación"
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
      Left            =   210
      TabIndex        =   1
      Top             =   180
      Value           =   -1  'True
      Width           =   1875
   End
   Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
      Bindings        =   "frmBuscarOperacionAnexo.frx":0000
      Height          =   4245
      Left            =   180
      OleObjectBlob   =   "frmBuscarOperacionAnexo.frx":001A
      TabIndex        =   0
      Top             =   540
      Width           =   17625
   End
End
Attribute VB_Name = "frmBuscarOperacionAnexo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoConsulta           As ADODB.Recordset
Dim strSQL                As String

Public strCodFondo        As String
Public numOperacion       As String
Public numAnexo           As String
Public indAnexo           As Boolean
Public strCodEmisor       As String
Public blnValSeleccionado As Boolean

Private Sub cmdCancelar_Click()
    Unload Me
    
End Sub

Private Sub cmdSeleccionar_Click()
    
    If tdgConsulta.SelBookmarks.Count < 1 Then
        MsgBox "No se ha seleccionado ningún registro", vbCritical, Me.Caption
        Exit Sub
    End If
    
    indAnexo = rbBuscarAnexo.Value
    
    numAnexo = tdgConsulta.Columns(0)
    numOperacion = tdgConsulta.Columns(1)
    
    strCodEmisor = tdgConsulta.Columns(3)
    blnValSeleccionado = True
    
    Unload Me
End Sub

Private Sub Form_Load()
    Call Buscar
    blnValSeleccionado = False
    CentrarForm Me
    
End Sub

Private Sub Buscar()
    Set adoConsulta = New ADODB.Recordset
    Me.MousePointer = vbHourglass

    If rbBuscarOperacion.Value = True Then
        strSQL = "SELECT IO.NumAnexo, IO.NumOperacion, IO.CodEmisor, IP1.DescripPersona as DescEmisor, IP2.DescripPersona as DescObligado, " & _
                    "II.Nemotecnico, IO.NumDocumentoFisico, IO.FechaEmision,IO.FechaVencimiento,IO.ValorNominalDscto, " & _
                    "dbo.uf_ACCalcularDeudaTotal(IO.CodFondo,IO.CodAdministradora,IO.NumOperacion,'" & gstrFechaActual & "')  as Deuda "
    Else
        strSQL = "SELECT IO.NumAnexo, '<blank>' as NumOperacion, IO.CodEmisor, IP1.DescripPersona as DescEmisor, '<blank>' as DescObligado, " & _
                    "'<blank>' as Nemotecnico, '<blank>' as NumDocumentoFisico,IO.FechaEmision,'<blank>' as FechaVencimiento, " & _
                    "sum(IO.ValorNominalDscto) as ValorNominalDscto, " & _
                    "sum(dbo.uf_ACCalcularDeudaTotal(IO.CodFondo,IO.CodAdministradora,IO.NumOperacion,'" & gstrFechaActual & "')) as Deuda "
    End If
            
    strSQL = strSQL & " from InversionOperacion IO " & "join InversionKardex IK on (IK.CodFile = IO.CodFile AND IK.CodAnalitica = IO.CodAnalitica " & _
            "AND IK.SaldoFinal <> 0 AND " & "IK.CodFondo = IO.CodFondo AND IK.IndUltimoMovimiento ='X') " & _
            "join InstitucionPersona IP1 on (IO.CodEmisor = IP1.CodPersona and IP1.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "') " & _
            "join InstitucionPersona IP2 on (IO.CodObligado = IP2.CodPersona and IP2.TipoPersona = '" & Codigo_Tipo_Persona_Obligado & "') " & _
            "join InstrumentoInversion II on (IO.CodFile = II.CodFile  and IO.CodAnalitica = II.CodAnalitica and IO.CodFondo = II.CodFondo) " & _
            "where IO.CodFondo = '" & strCodFondo & "' " & "AND IO.CodAdministradora = '002' " & "AND IO.CodFile in ('014','015') " & _
            "and IO.TipoOperacion = '01' "
                                
    If rbBuscarOperacion.Value = False Then
        strSQL = strSQL & " group by IO.NumAnexo, IO.CodEmisor, IP1.DescripPersona,IO.FechaEmision "
    End If

    strSQL = strSQL & "order by 4,1,2"

    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With

    tdgConsulta.Refresh
 
    tdgConsulta.DataSource = adoConsulta
    
    tdgConsulta.Refresh
    Me.MousePointer = vbDefault

End Sub

Private Sub rbBuscarAnexo_Click()
    tdgConsulta.Columns.Item(1).Visible = False
    tdgConsulta.Columns.Item(4).Visible = False
    tdgConsulta.Columns.Item(5).Visible = False
    tdgConsulta.Columns.Item(6).Visible = False
    tdgConsulta.Columns.Item(8).Visible = False
    tdgConsulta.MultiSelect = dbgMultiSelectNone
    tdgConsulta.Width = 9355
    cmdSeleccionar.Left = 6460
    cmdCancelar.Left = 8050
    Me.Width = 9805
    Call Buscar
End Sub

Private Sub rbBuscarOperacion_Click()
    tdgConsulta.Columns.Item(1).Visible = True
    tdgConsulta.Columns.Item(4).Visible = True
    tdgConsulta.Columns.Item(5).Visible = True
    tdgConsulta.Columns.Item(6).Visible = True
    tdgConsulta.Columns.Item(8).Visible = True
    tdgConsulta.MultiSelect = dbgMultiSelectExtended
    tdgConsulta.Width = 17625
    cmdSeleccionar.Left = 14790
    cmdCancelar.Left = 16380
    Me.Width = 18075
    Call Buscar
End Sub
