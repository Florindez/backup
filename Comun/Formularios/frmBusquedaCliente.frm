VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBusquedaCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "B�squeda de Clientes "
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   8730
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Buscar"
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
      Index           =   2
      Left            =   3750
      Picture         =   "frmBusquedaCliente.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Buscar"
      Top             =   3300
      Width           =   1200
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Seleccionar"
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
      Index           =   1
      Left            =   2250
      Picture         =   "frmBusquedaCliente.frx":00EA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Seleccionar"
      Top             =   3300
      Width           =   1200
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Cerrar"
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
      Index           =   0
      Left            =   5250
      Picture         =   "frmBusquedaCliente.frx":0695
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cerrar Ventana"
      Top             =   3300
      Width           =   1200
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   5820
      Top             =   3300
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
   Begin VB.Frame fraBusquedaCliente 
      Caption         =   "Criterios de B�squeda"
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
      Height          =   1440
      Left            =   135
      TabIndex        =   4
      Top             =   75
      Width           =   8475
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2550
         TabIndex        =   3
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox txtNumDocumento 
         Height          =   285
         Left            =   2550
         TabIndex        =   1
         Top             =   420
         Width           =   3615
      End
      Begin VB.OptionButton optCriterio 
         Caption         =   "Descripci�n"
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
         Left            =   435
         TabIndex        =   2
         Top             =   870
         Width           =   1785
      End
      Begin VB.OptionButton optCriterio 
         Caption         =   "Num. Documento"
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
         Left            =   435
         TabIndex        =   0
         Top             =   450
         Value           =   -1  'True
         Width           =   1830
      End
   End
   Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
      Bindings        =   "frmBusquedaCliente.frx":0C17
      Height          =   1605
      Left            =   135
      OleObjectBlob   =   "frmBusquedaCliente.frx":0C31
      TabIndex        =   8
      Top             =   1590
      Width           =   8445
   End
End
Attribute VB_Name = "frmBusquedaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strEstado As String

Private Sub Buscar()
        
    Dim strSql As String
    Dim adoresultAux1 As ADODB.Recordset
                                                                                    
    Me.MousePointer = vbHourglass
                

    strSql = "SELECT CodUnico,DescripParametro TipoIdentidad,NumIdentidad,DescripCliente,FechaIngreso,TipoIdentidad CodIdentidad "
    strSql = strSql & "FROM Cliente JOIN AuxiliarParametro ON(AuxiliarParametro.CodParametro=Cliente.TipoIdentidad AND AuxiliarParametro.CodTipoParametro='TIPIDE') "
    If Trim(txtNumDocumento.Text) <> "" And optCriterio(0).Value Then
        strSql = strSql & "WHERE NumIdentidad='" & Trim(txtNumDocumento.Text) & "'"
    ElseIf Trim(txtDescripcion.Text) <> "" And optCriterio(1).Value Then
        strSql = strSql & "WHERE DescripCliente LIKE '%" & Trim(txtDescripcion.Text) & "%'"
    End If
    
    With adoConsulta
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSql
        .Refresh
    End With
    
    tdgConsulta.Refresh
    
    If adoConsulta.Recordset.RecordCount > 0 Then strEstado = Reg_Consulta
        
    Me.MousePointer = vbDefault
                                    
End Sub

Private Sub Salir()

    Unload Me
    
End Sub

Private Sub cmdOpcion_Click(Index As Integer)

    Select Case Index
        Case 0 '*** Cancelar ***
            Call Salir
        Case 1 '*** Seleccionar ***
            Call Modificar
        Case 2 '*** Buscar ***
            Call Buscar
    End Select
    
End Sub

Private Sub Form_Load()

    Call InicializarValores
    
    CentrarForm Me
    Call Buscar
End Sub

Private Sub InicializarValores()

    strEstado = Reg_Defecto
    
    optCriterio(0).Value = vbUnchecked
    optCriterio(1).Value = vbUnchecked
    optCriterio(0).Value = vbChecked
    
End Sub
Private Sub Modificar()

    Dim intRegistro As Integer
    
    If strEstado = Reg_Consulta Then
        Select Case gstrFormulario
            Case "frmMancomunado"
                intRegistro = ObtenerItemLista(garrTipoDocumento(), Trim(tdgConsulta.Columns("TipoIdentidad")))
                If intRegistro >= 0 Then frmMancomunado.cboTipoDocumento.ListIndex = intRegistro
                
                frmMancomunado.txtNumDocumento = Trim(tdgConsulta.Columns(2))
                frmMancomunado.lblDescripCliente = Trim(tdgConsulta.Columns(3))
                frmMancomunado.lblCodCliente = Trim(tdgConsulta.Columns(0))
                
            Case "frmRepresentanteParticipe"
                intRegistro = ObtenerItemLista(garrTipoDocumento(), Trim(tdgConsulta.Columns(5)))
                If intRegistro >= 0 Then frmRepresentanteParticipe.cboTipoDocumento.ListIndex = intRegistro
                
                frmRepresentanteParticipe.txtNumDocumento = Trim(tdgConsulta.Columns(2))
                frmRepresentanteParticipe.lblDescripCliente = Trim(tdgConsulta.Columns(3))
                frmRepresentanteParticipe.lblCodCliente = Trim(tdgConsulta.Columns(0))
                
            Case "frmContratoParticipe"
                intRegistro = ObtenerItemLista(garrTipoDocumento(), Trim(tdgConsulta.Columns(5)))
                If intRegistro >= 0 Then frmContratoParticipe.cboTipoDocumento.ListIndex = intRegistro
                
                frmContratoParticipe.txtNumDocumentoCliente = Trim(tdgConsulta.Columns(2))
                frmContratoParticipe.lblDescripParticipe = Trim(tdgConsulta.Columns(3))
                frmContratoParticipe.lblDescripTitular = Trim(tdgConsulta.Columns(3))
                frmContratoParticipe.lblCodCliente = Trim(tdgConsulta.Columns(0))
        
        End Select
        
        Call Salir
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmBusquedaCliente = Nothing
    
End Sub

Private Sub optCriterio_Click(Index As Integer)

    If Index = 0 Then
        txtNumDocumento.Enabled = True
        txtDescripcion.Enabled = False
        txtNumDocumento.Text = ""
    Else
        txtDescripcion.Enabled = True
        txtNumDocumento.Enabled = False
        txtDescripcion.Text = ""
    End If
    
    Call Buscar
    
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub
