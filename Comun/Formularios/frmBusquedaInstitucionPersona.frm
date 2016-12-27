VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBusquedaInstitucionPersona 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de Proveedores"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7980
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
      Left            =   360
      Picture         =   "frmBusquedaInstitucionPersona.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Buscar"
      Top             =   3360
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
      Left            =   1920
      Picture         =   "frmBusquedaInstitucionPersona.frx":00EA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Seleccionar"
      Top             =   3360
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
      Left            =   3480
      Picture         =   "frmBusquedaInstitucionPersona.frx":01BA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cerrar Ventana"
      Top             =   3360
      Width           =   1200
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   5160
      Top             =   3480
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
      Left            =   200
      TabIndex        =   4
      Top             =   165
      Width           =   7605
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   3030
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtNumDocumento 
         Height          =   285
         Left            =   3030
         TabIndex        =   1
         Top             =   420
         Visible         =   0   'False
         Width           =   3735
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
         Left            =   675
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
         Left            =   675
         TabIndex        =   0
         Top             =   450
         Width           =   1830
      End
   End
   Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
      Bindings        =   "frmBusquedaInstitucionPersona.frx":073C
      Height          =   1455
      Left            =   200
      OleObjectBlob   =   "frmBusquedaInstitucionPersona.frx":0756
      TabIndex        =   8
      Top             =   1800
      Width           =   7605
   End
   Begin VB.Label lblTipoInstitucion 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6720
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmBusquedaInstitucionPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strEstado As String


Private Sub Buscar()
        
    Dim strSQL As String
    Dim adoresultAux1 As ADODB.Recordset
                                                                                    
    Me.MousePointer = vbHourglass
                
    If Trim(txtNumDocumento.Text) <> Valor_Caracter Or Trim(txtDescripcion.Text) <> Valor_Caracter Then
        strSQL = "SELECT CodPersona,DescripParametro TipoIdentidad,NumIdentidad,DescripPersona,TipoIdentidad CodIdentidad, Direccion1 + Direccion2 Direccion " & _
            "FROM InstitucionPersona IP JOIN AuxiliarParametro AP ON(AP.CodParametro=IP.TipoIdentidad AND AP.CodTipoParametro='TIPIDE') "
        
        Select Case lblTipoInstitucion.Caption
            Case Codigo_Tipo_Persona_Agente: strSQL = strSQL & "WHERE TipoPersona='" & Codigo_Tipo_Persona_Agente & "' AND "
            Case Codigo_Tipo_Persona_Emisor: strSQL = strSQL & "WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND "
            Case Codigo_Tipo_Persona_Proveedor: strSQL = strSQL & "WHERE TipoPersona='" & Codigo_Tipo_Persona_Proveedor & "' AND "
            Case Codigo_Tipo_Persona_Relacionado: strSQL = strSQL & "WHERE TipoPersona='" & Codigo_Tipo_Persona_Relacionado & "' AND "
        End Select
        
        If optCriterio(0).Value Then
            strSQL = strSQL & "NumIdentidad='" & Trim(txtNumDocumento.Text) & "'"
        Else
            strSQL = strSQL & "DescripPersona LIKE '%" & Trim(txtDescripcion.Text) & "%'"
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
            Case "frmRegistroCompras"
                frmRegistroCompras.lblCodProveedor = Trim(tdgConsulta.Columns(0))
                frmRegistroCompras.lblProveedor = Trim(tdgConsulta.Columns(3))
                frmRegistroCompras.lblNumDocID = Trim(tdgConsulta.Columns(2))
                frmRegistroCompras.lblDireccion = Trim(tdgConsulta.Columns(4))
        End Select
        
        Call Salir
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmBusquedaInstitucionPersona = Nothing
    
End Sub

Private Sub optCriterio_Click(Index As Integer)

    If Index = 0 Then
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
