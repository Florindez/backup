VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBusquedaParticipe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Búsqueda de Contratos"
   ClientHeight    =   4830
   ClientLeft      =   1065
   ClientTop       =   2205
   ClientWidth     =   7035
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
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4830
   ScaleWidth      =   7035
   Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
      Bindings        =   "frmBusquedaParticipe.frx":0000
      Height          =   1575
      Left            =   195
      OleObjectBlob   =   "frmBusquedaParticipe.frx":001A
      TabIndex        =   10
      Top             =   2160
      Width           =   6645
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Cerrar"
      Height          =   735
      Index           =   0
      Left            =   3510
      Picture         =   "frmBusquedaParticipe.frx":469A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cerrar Ventana"
      Top             =   3960
      Width           =   1200
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Seleccionar"
      Height          =   735
      Index           =   1
      Left            =   480
      Picture         =   "frmBusquedaParticipe.frx":4C1C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Seleccionar"
      Top             =   3960
      Width           =   1200
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Buscar"
      Height          =   735
      Index           =   2
      Left            =   1980
      Picture         =   "frmBusquedaParticipe.frx":51C7
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Buscar"
      Top             =   3960
      Width           =   1200
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   5160
      Top             =   4080
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
   Begin VB.Frame fraBusquedaCliente 
      Caption         =   "Criterios de Búsqueda"
      ForeColor       =   &H00800000&
      Height          =   1800
      Left            =   200
      TabIndex        =   0
      Top             =   200
      Width           =   6640
      Begin VB.OptionButton optCriterio 
         Caption         =   "Descripción"
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   2
         Left            =   435
         TabIndex        =   6
         Top             =   1300
         Width           =   1785
      End
      Begin VB.TextBox txtDescripcion 
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
         Left            =   2550
         TabIndex        =   5
         Top             =   1280
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.OptionButton optCriterio 
         Caption         =   "Código Partícipe"
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   0
         Left            =   435
         TabIndex        =   4
         Top             =   450
         Width           =   1830
      End
      Begin VB.OptionButton optCriterio 
         Caption         =   "Num. Documento"
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   1
         Left            =   435
         TabIndex        =   3
         Top             =   870
         Width           =   1785
      End
      Begin VB.TextBox txtCodParticipe 
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
         Left            =   2550
         TabIndex        =   2
         Top             =   420
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox txtNumDocumento 
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
         Left            =   2550
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmBusquedaParticipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strEstado As String

Private Sub Salir()

    Unload Me
    
End Sub

Private Sub Modificar()

    Dim intRegistro As Integer
    
    If strEstado = Reg_Consulta Then
        Select Case gstrFormulario

            Case "frmAbonoRetiroCtaCliente"
            
                If frmAbonoRetiroCtaCliente.tabPagos.Tab = 0 Then
            
                intRegistro = ObtenerItemLista(garrTipoDocumento(), Trim(tdgConsulta.Columns(5)))
                frmAbonoRetiroCtaCliente.lblDescripParticipeBusqueda = Trim(tdgConsulta.Columns(3))
                frmAbonoRetiroCtaCliente.txtCodParticipeBusqueda = Trim(tdgConsulta.Columns(0))
                gstrCodParticipe = Trim(tdgConsulta.Columns(0))
            
                Else
            
                intRegistro = ObtenerItemLista(garrTipoDocumento(), Trim(tdgConsulta.Columns(5)))
                frmAbonoRetiroCtaCliente.lblDescripParticipe = Trim(tdgConsulta.Columns(3))
                frmAbonoRetiroCtaCliente.txtCodParticipe = Trim(tdgConsulta.Columns(0))
                gstrCodParticipe = Trim(tdgConsulta.Columns(0))
                
                End If
                
              Case "frmMovimientoCambiario"
                
                intRegistro = ObtenerItemLista(garrTipoDocumento(), Trim(tdgConsulta.Columns(5)))
                frmMovimientoCambiario.lblNombreCliente = Trim(tdgConsulta.Columns(3))
                frmMovimientoCambiario.txtCodParticipe = Trim(tdgConsulta.Columns(0))
                gstrCodParticipe = Trim(tdgConsulta.Columns(0))
                
            
            Case "frmMovimientoCambiarioBusqueda"
                
                intRegistro = ObtenerItemLista(garrTipoDocumento(), Trim(tdgConsulta.Columns(5)))
                frmMovimientoCambiario.lblDescripParticipeBusqueda = Trim(tdgConsulta.Columns(3))
                frmMovimientoCambiario.txtCodParticipeBusqueda = Trim(tdgConsulta.Columns(0))
                gstrCodParticipe = Trim(tdgConsulta.Columns(0))
            
            
            
        End Select
        
        Call Salir
    End If
    
End Sub

Private Sub InicializarValores()

    strEstado = Reg_Defecto
    
    optCriterio(0).Value = vbUnchecked
    optCriterio(1).Value = vbUnchecked
    optCriterio(2).Value = vbUnchecked
    optCriterio(0).Value = vbChecked
    
End Sub

Public Sub Buscar()
        
    Dim strSQL As String
    Dim adoresultAux1 As ADODB.Recordset
                                                                                    
    Me.MousePointer = vbHourglass
                
    If Trim(txtCodParticipe.Text) <> "" Or Trim(txtNumDocumento.Text) <> "" Or Trim(txtDescripcion.Text) <> "" Then
        strSQL = "SELECT CodParticipe,AP1.DescripParametro TipoIdentidad,NumIdentidad,DescripParticipe,FechaIngreso,TipoIdentidad CodTipoIdentidad,AP2.DescripParametro TipoMancomuno "
        strSQL = strSQL & "FROM ParticipeContrato JOIN AuxiliarParametro AP1 ON(AP1.CodParametro=ParticipeContrato.TipoIdentidad AND AP1.CodTipoParametro='TIPIDE') "
        strSQL = strSQL & "JOIN AuxiliarParametro AP2 ON(AP2.CodParametro=ParticipeContrato.TipoMancomuno AND AP2.CodTipoParametro='TIPMAN') "
        If optCriterio(0).Value Then
            strSQL = strSQL & "WHERE CodParticipe='" & Trim(txtCodParticipe.Text) & "'"
        ElseIf optCriterio(1).Value Then
            strSQL = strSQL & "WHERE NumIdentidad='" & Trim(txtNumDocumento.Text) & "'"
        Else
        
            strSQL = strSQL & "WHERE DescripParticipe LIKE '%" & Trim(txtDescripcion.Text) & "%'"
        End If
        
        With adoConsulta
            .ConnectionString = gstrConnectConsulta
            .RecordSource = strSQL
            .Refresh
        End With
        
        tdgConsulta.Refresh
        
        If adoConsulta.Recordset.RecordCount > 0 Then strEstado = Reg_Consulta
        
    End If
    
    Me.MousePointer = vbDefault
                                    
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

Private Sub Form_Load()

    Call InicializarValores
    
    CentrarForm Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmBusquedaParticipe = Nothing
    
End Sub

Private Sub optCriterio_Click(index As Integer)

    If index = 0 Then
        txtCodParticipe.Visible = True
        txtNumDocumento.Visible = False
        txtDescripcion.Visible = False
        txtCodParticipe.Text = ""
    ElseIf index = 1 Then
        txtNumDocumento.Visible = True
        txtCodParticipe.Visible = False
        txtDescripcion.Visible = False
        txtNumDocumento.Text = ""
    Else
        txtDescripcion.Visible = True
        txtCodParticipe.Visible = False
        txtNumDocumento.Visible = False
        txtDescripcion.Text = ""
    End If
    
End Sub

Private Sub txtCodParticipe_LostFocus()

    txtCodParticipe.Text = Format(txtCodParticipe.Text, "00000000000000000000")
    
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

