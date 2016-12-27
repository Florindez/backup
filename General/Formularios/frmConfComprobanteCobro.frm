VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form frmConfComprobanteCobro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Impresion - Comprobante Cobro"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmnSalir 
      Caption         =   "Salir"
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
      Left            =   8130
      MaskColor       =   &H00800000&
      TabIndex        =   10
      Top             =   540
      Width           =   1065
   End
   Begin VB.CommandButton cmbListar 
      Caption         =   "Listar"
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
      Left            =   8130
      MaskColor       =   &H00800000&
      TabIndex        =   9
      Top             =   150
      Width           =   1065
   End
   Begin VB.TextBox txtSerieComprobante 
      Height          =   315
      Left            =   1710
      MaxLength       =   3
      TabIndex        =   7
      Top             =   900
      Width           =   615
   End
   Begin VB.ComboBox cboTipoComprobante 
      Height          =   315
      Left            =   1710
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   510
      Width           =   5955
   End
   Begin VB.ComboBox cboFondo 
      Height          =   315
      Left            =   1710
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   150
      Width           =   5955
   End
   Begin DXDBGRIDLibCtl.dxDBGrid gCabecera 
      Height          =   2565
      Left            =   120
      OleObjectBlob   =   "frmConfComprobanteCobro.frx":0000
      TabIndex        =   0
      Top             =   1290
      Width           =   9060
   End
   Begin DXDBGRIDLibCtl.dxDBGrid gDetalle 
      Height          =   2325
      Left            =   120
      OleObjectBlob   =   "frmConfComprobanteCobro.frx":3867
      TabIndex        =   1
      Top             =   3915
      Width           =   9060
   End
   Begin DXDBGRIDLibCtl.dxDBGrid gTotales 
      Height          =   2295
      Left            =   120
      OleObjectBlob   =   "frmConfComprobanteCobro.frx":6C2F
      TabIndex        =   2
      Top             =   6300
      Width           =   9060
   End
   Begin VB.Label lblDescrip 
      AutoSize        =   -1  'True
      Caption         =   "Serie:"
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
      Index           =   3
      Left            =   60
      TabIndex        =   8
      Top             =   960
      Width           =   510
   End
   Begin VB.Label lblDescrip 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Comprobante:"
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
      Index           =   11
      Left            =   60
      TabIndex        =   6
      Top             =   570
      Width           =   1620
   End
   Begin VB.Label lblDescrip 
      AutoSize        =   -1  'True
      Caption         =   "Fondo:"
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
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   180
      Width           =   600
   End
End
Attribute VB_Name = "frmConfComprobanteCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()                  As String
Dim arrTipoComprobante()        As String
Dim strCodFondo                 As String
Dim strCodTipoComprobante       As String

Private Sub cmbListar_Click()
listar
End Sub

Private Sub cmnSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim strSQL As String

    CentrarForm Me

    Call ValidarPermisoUsoControl(Trim(gstrLogin), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    ConfGrid gCabecera, True, False, False, False
    ConfGrid gDetalle, True, False, False, False
    ConfGrid gTotales, True, False, False, False
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(29,'" & gstrCodAdministradora & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    '*** Tipo de Comprobante Sunat ***
    strSQL = "SELECT CodTipoComprobantePago CODIGO,DescripTipoComprobantePago DESCRIP From TipoComprobantePago ORDER BY DescripTipoComprobantePago"
    CargarControlLista strSQL, cboTipoComprobante, arrTipoComprobante(), Sel_Defecto
End Sub

Private Sub listar()
Dim strSQL As String

    strSQL = "SELECT Identificador,GlsCampo, tipoDato, decimales, indImprime, impX, impY, impLongitud, GlsObs " & _
             "FROM objRegistroVenta " & _
             "WHERE CodAdministradora = '" & gstrCodAdministradora & "' " & _
               "AND CodFondo = '" & strCodFondo & "' " & _
               "AND CodTipoComprobante = '" & strCodTipoComprobante & "' " & _
               "AND SerieComprobante = '" & txtSerieComprobante.Text & "' "
             
    With gCabecera
        .DefaultFields = False
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = gstrConnectConsulta
        .Dataset.ADODataset.CommandText = strSQL & "AND tipoObj = 'C' "
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Open
        .Dataset.Active = True
        .KeyField = "Identificador"
    End With
    
    'csql = "Select Identificador,etiqueta,indImprime,impX,impY,impLongitud From objdocventas Where idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' AND idDocumento = '" & txtCod_Documento.Text & "'  AND idSerie = '" & txt_Serie.Text & "' and trim(GlsCampo) <> ''"
    With gDetalle
        .DefaultFields = False
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = gstrConnectConsulta
        .Dataset.ADODataset.CommandText = strSQL & "AND tipoObj = 'D' "
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Open
        .Dataset.Active = True
        .KeyField = "Identificador"
    End With
    
    'csql = "Select Identificador,GlsObs,indImprime,impX,impY,impLongitud From objdocventas Where idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'T' AND idDocumento = '" & txtCod_Documento.Text & "' AND idSerie = '" & txt_Serie.Text & "' and trim(GlsCampo) <> ''"
    With gTotales
        .DefaultFields = False
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = gstrConnectConsulta
        .Dataset.ADODataset.CommandText = strSQL & "AND tipoObj = 'T' "
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Open
        .Dataset.Active = True
        .KeyField = "Identificador"
    End With
    
End Sub

Private Sub cboFondo_Click()
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
End Sub

Private Sub cboTipoComprobante_Click()

    strCodTipoComprobante = Valor_Caracter
    If cboTipoComprobante.ListIndex < 0 Then Exit Sub
    
    strCodTipoComprobante = arrTipoComprobante(cboTipoComprobante.ListIndex)
   
End Sub

Private Sub gCabecera_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
If KeyCode = 13 Then gCabecera.Dataset.Post
End Sub

Private Sub gDetalle_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
If KeyCode = 13 Then gDetalle.Dataset.Post
End Sub

Private Sub gTotales_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
If KeyCode = 13 Then gTotales.Dataset.Post
End Sub

Private Sub txtSerieComprobante_LostFocus()
    txtSerieComprobante.Text = Format(txtSerieComprobante.Text, "000")
End Sub
