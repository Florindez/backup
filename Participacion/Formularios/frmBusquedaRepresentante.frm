VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBusquedaRepresentante 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Búsqueda de Representantes"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraBusquedaCliente 
      Caption         =   "Criterios de Búsqueda"
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
      Left            =   210
      TabIndex        =   4
      Top             =   180
      Width           =   6640
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
         TabIndex        =   8
         Top             =   510
         Width           =   1830
      End
      Begin VB.OptionButton optCriterio 
         Caption         =   "Descripción"
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
         TabIndex        =   7
         Top             =   930
         Width           =   1785
      End
      Begin VB.TextBox txtNumDocumento 
         Height          =   285
         Left            =   2550
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   2550
         TabIndex        =   5
         Top             =   900
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label lblCodClienteParticipe 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4230
         TabIndex        =   9
         Top             =   150
         Visible         =   0   'False
         Width           =   2295
      End
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
      Left            =   3495
      Picture         =   "frmBusquedaRepresentante.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cerrar Ventana"
      Top             =   3495
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
      Left            =   495
      Picture         =   "frmBusquedaRepresentante.frx":0582
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Seleccionar"
      Top             =   3495
      Width           =   1200
   End
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
      Left            =   1995
      Picture         =   "frmBusquedaRepresentante.frx":0B2D
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Buscar"
      Top             =   3495
      Width           =   1200
   End
   Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
      Bindings        =   "frmBusquedaRepresentante.frx":0C17
      Height          =   1455
      Left            =   210
      OleObjectBlob   =   "frmBusquedaRepresentante.frx":0C31
      TabIndex        =   0
      Top             =   1830
      Width           =   6645
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   5175
      Top             =   3495
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
End
Attribute VB_Name = "frmBusquedaRepresentante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strEstado As String



Private Sub CargarRapido()
        
    Dim strSql As String
    Dim adoresultAux1 As ADODB.Recordset
                                                                                    
    Me.MousePointer = vbHourglass
    
        strSql = "Select CR.CodRepresentante,AP.DescripParametro as TipoIdentidad,CR.TipoDocumento CodTipoIdentidad,CR.NroDocumento NumIdentidad,CR.Nombres, CR.Apellidos DescripCliente,CR.Vinculacion "
        strSql = strSql & "FROM ClienteRepresentantes CR JOIN AuxiliarParametro AP ON(AP.CodParametro=CR.TipoDocumento AND AP.CodTipoParametro='TIPIDE') "
        
        
        strSql = strSql & "WHERE CR.CodCliente='" & Trim(lblCodClienteParticipe.Caption) & "'"
       
        With adoConsulta
            .ConnectionString = gstrConnectConsulta
            .RecordSource = strSql
            .Refresh
        End With
        
        tdgConsulta.Refresh
    
        
    
    Me.MousePointer = vbDefault
                                    
End Sub

Private Sub Buscar()
        
    Dim strSql As String
    Dim adoresultAux1 As ADODB.Recordset
                                                                                    
    Me.MousePointer = vbHourglass
                
    If Trim(txtNumDocumento.Text) <> "" Or Trim(txtDescripcion.Text) <> "" Then
        'strSQL = "SELECT CodUnico,DescripParametro TipoIdentidad,NumIdentidad,DescripCliente,FechaIngreso,TipoIdentidad CodIdentidad "
        'strSQL = strSQL & "FROM Cliente JOIN AuxiliarParametro ON(AuxiliarParametro.CodParametro=Cliente.TipoIdentidad AND AuxiliarParametro.CodTipoParametro='TIPIDE') "
        strSql = "Select CR.CodRepresentante,AP.DescripParametro as TipoIdentidad,CR.TipoDocumento CodTipoIdentidad,CR.NroDocumento NumIdentidad,CR.Nombres, CR.Apellidos DescripCliente,CR.Vinculacion "
        strSql = strSql & "FROM ClienteRepresentantes CR JOIN AuxiliarParametro AP ON(AP.CodParametro=CR.TipoDocumento AND AP.CodTipoParametro='TIPIDE') "
        If optCriterio(0).Value Then
            'strSQL = strSQL & "WHERE NumIdentidad='" & Trim(txtNumDocumento.Text) & "'"
            strSql = strSql & "WHERE CR.CodCliente='" & Trim(lblCodClienteParticipe.Caption) & "' and CR.NumDocumento='" & Trim(txtNumDocumento.Text) & "'"
        Else
            'strSQL = strSQL & "WHERE DescripCliente LIKE '%" & Trim(txtDescripcion.Text) & "%'"
            strSql = strSql & "WHERE CR.CodCliente='" & Trim(lblCodClienteParticipe.Caption) & "' and (CR.Nombres+' '+Apellidos) LIKE '%" & Trim(txtDescripcion.Text) & "%'"
        End If
        
        With adoConsulta
            .ConnectionString = gstrConnectConsulta
            .RecordSource = strSql
            .Refresh
        End With
        
        tdgConsulta.Refresh
        
        If adoConsulta.Recordset.RecordCount > 0 Then strEstado = Reg_Consulta
        
    End If
    
    Me.MousePointer = vbDefault
                                    
End Sub

Private Sub Salir()

    Unload Me
    
End Sub

Private Sub cmdOpcion_Click(index As Integer)

    Select Case index
        Case 0 '*** Cancelar ***
            Call Salir
        Case 1 '*** Seleccionar ***
            Call Modificar
        Case 2 '*** Buscar ***
            Call Buscar
    End Select
    
End Sub

Private Sub Form_Activate()

    Call CargarRapido
    
End Sub

Private Sub Form_Load()

    Call InicializarValores
    
    CentrarForm Me
    
End Sub

Private Sub InicializarValores()

    strEstado = Reg_Defecto
    
    
    optCriterio(0).Value = vbUnchecked
    optCriterio(1).Value = vbUnchecked
    optCriterio(0).Value = vbChecked
    
End Sub
Private Sub Modificar()

    Dim intRegistro As Integer
    
    'If strEstado = Reg_Consulta Then
        Select Case gstrFormulario
            
            Case "frmRepresentanteParticipe"
                intRegistro = ObtenerItemLista(garrTipoDocumento(), Trim(tdgConsulta.Columns(5)))
                If intRegistro >= 0 Then frmRepresentanteParticipe.cboTipoDocumento.ListIndex = intRegistro
                
                frmRepresentanteParticipe.txtNumDocumento = Trim(tdgConsulta.Columns(2))
                frmRepresentanteParticipe.lblDescripCliente = Trim(tdgConsulta.Columns(3))
                frmRepresentanteParticipe.lblCodCliente = Trim(tdgConsulta.Columns(0))
                frmRepresentanteParticipe.txtTipoDocumento = Trim(tdgConsulta.Columns(1))
                frmRepresentanteParticipe.txtCodTipoDocumento = Trim(tdgConsulta.Columns(5))
                 
            Case "frmContratoParticipe"
                intRegistro = ObtenerItemLista(garrTipoDocumento(), Trim(tdgConsulta.Columns(5)))
                If intRegistro >= 0 Then frmContratoParticipe.cboTipoDocumento.ListIndex = intRegistro
                
                frmContratoParticipe.txtNumDocumentoCliente = Trim(tdgConsulta.Columns(2))
                frmContratoParticipe.lblDescripParticipe = Trim(tdgConsulta.Columns(3))
                frmContratoParticipe.lblDescripTitular = Trim(tdgConsulta.Columns(3))
                frmContratoParticipe.lblCodCliente = Trim(tdgConsulta.Columns(0))
                frmContratoParticipe.txtNumDocumento = Trim(tdgConsulta.Columns(2))
        End Select
        
        Call Salir
    'End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmBusquedaRepresentante = Nothing
    
End Sub

Private Sub optCriterio_Click(index As Integer)

    If index = 0 Then
        txtNumDocumento.Visible = True
        txtDescripcion.Visible = False
        txtNumDocumento.Text = ""
    Else
        txtDescripcion.Visible = True
        txtNumDocumento.Visible = False
        txtDescripcion.Text = ""
    End If
    
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub
