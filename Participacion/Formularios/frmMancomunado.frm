VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmMancomunado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mancomunados"
   ClientHeight    =   6240
   ClientLeft      =   1170
   ClientTop       =   3975
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   12135
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   5400
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1296
      Buttons         =   2
      Caption0        =   "Con&sultar"
      Tag0            =   "3"
      Visible0        =   0   'False
      ToolTipText0    =   "Consultar"
      Caption1        =   "&Cerrar"
      Tag1            =   "9"
      Visible1        =   0   'False
      ToolTipText1    =   "Cerrar Ventana"
      UserControlWidth=   2700
   End
   Begin VB.Frame fraMancomunado 
      Height          =   5205
      Left            =   30
      TabIndex        =   4
      Top             =   60
      Width           =   12075
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Agregar detalle"
         Top             =   3060
         Width           =   375
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Quitar detalle"
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "A"
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   2640
         Width           =   375
      End
      Begin TAMControls.TAMTextBox txtPorcenParticipacion 
         Height          =   315
         Left            =   2250
         TabIndex        =   14
         Top             =   2040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmMancomunado.frx":0000
         Decimales       =   6
         ColorEnfoque    =   8454143
         Apariencia      =   1
         Borde           =   1
         DecimalesValue  =   6
         MaximoValor     =   100
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Height          =   2205
         Left            =   930
         OleObjectBlob   =   "frmMancomunado.frx":001C
         TabIndex        =   13
         Top             =   2640
         Width           =   10545
      End
      Begin VB.CommandButton cmdBusqueda 
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
         Height          =   315
         Left            =   10875
         TabIndex        =   2
         ToolTipText     =   "Búsqueda de Cliente"
         Top             =   800
         Width           =   375
      End
      Begin VB.ComboBox cboTipoDocumento 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2250
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1230
         Width           =   3255
      End
      Begin VB.TextBox txtNumDocumento 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2250
         TabIndex        =   1
         Top             =   1635
         Width           =   3255
      End
      Begin VB.Label lblDescrip 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   5
         Left            =   3990
         TabIndex        =   15
         Top             =   2070
         Width           =   240
      End
      Begin VB.Label lblCodCliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5850
         TabIndex        =   12
         Top             =   1230
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblDescripCliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2250
         TabIndex        =   11
         Top             =   795
         Width           =   8595
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Num. Documento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   1695
         Width           =   1500
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Participe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   375
         Width           =   1500
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Tipo Documento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   1290
         Width           =   1500
      End
      Begin VB.Label lblParticipe 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2250
         TabIndex        =   7
         Top             =   360
         Width           =   8985
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   3
         Left            =   360
         TabIndex        =   6
         Top             =   825
         Width           =   1500
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Porcen. Particip."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   2085
         Width           =   1740
      End
   End
End
Attribute VB_Name = "frmMancomunado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrTipoMancomuno()      As String
Dim strCodTipoDocumento     As String
Dim strCodCliente           As String, intNumSecuencial         As Integer
Dim strEstado               As String
Public strCodTipoMancomuno  As String
Dim adoRegistroAux          As ADODB.Recordset
Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vModify
            Call Modificar
        Case vDelete
            Call Eliminar
        Case vSearch
            Call Buscar
        Case vReport
            Call Imprimir
        Case vSave
            Call Grabar
        Case vCancel
            Call Cancelar
        Case vExit
            Call Salir
        
    End Select
    
End Sub

Public Sub Adicionar()
    
End Sub

Public Sub Anterior()

End Sub

Public Sub Ayuda()

End Sub

Public Sub Buscar()

    Dim adoRegistro As ADODB.Recordset
    Dim adoField As ADODB.Field
    
    Dim strSql As String
    
    Set adoRegistro = New ADODB.Recordset
        
    Call ConfiguraRecordsetAuxiliar
    
    'If strEstado = Reg_Edicion Then
        
        strSql = "{ call up_ACSelDatosParametro(14,'" & gstrCodParticipe & "') }"

        With adoRegistro
            .ActiveConnection = gstrConnectConsulta
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open strSql
        
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    adoRegistroAux.AddNew
                    For Each adoField In adoRegistroAux.Fields
                        adoRegistroAux.Fields(adoField.Name) = adoRegistro.Fields(adoField.Name)
                    Next
                    adoRegistroAux.Update
                    adoRegistro.MoveNext
                Loop
                adoRegistroAux.MoveFirst
            End If
            
        End With
    
    'End If
    
    tdgConsulta.DataSource = adoRegistroAux
    
    tdgConsulta.Refresh
    
    If adoRegistroAux.RecordCount > 0 Then strEstado = Reg_Consulta
        
    Me.MousePointer = vbDefault
    
                
End Sub

Public Sub Cancelar()

    Call Salir
    
End Sub

Public Sub Eliminar()
    
End Sub

Public Sub Grabar()
                
End Sub

Public Sub Imprimir()

End Sub

Public Sub Modificar()

    Dim intRegistro As Integer
    
    If strEstado = Reg_Consulta Then
        'intNumSecuencial = CInt(tdgConsulta.Columns(0))
        lblCodCliente.Caption = Trim(tdgConsulta.Columns(1))
        
        intRegistro = ObtenerItemLista(garrTipoDocumento(), Trim(tdgConsulta.Columns(4)))
        If intRegistro >= 0 Then cboTipoDocumento.ListIndex = intRegistro
        
        txtNumDocumento.Text = Trim(tdgConsulta.Columns(4))
        lblDescripCliente.Caption = Trim(tdgConsulta.Columns(5))
        
'        intRegistro = ObtenerItemLista(arrTipoMancomuno(), Trim(tdgConsulta.Columns(6)))
'        If intRegistro >= 0 Then cboTipoMancomuno.ListIndex = intRegistro
        
    End If
        
End Sub

Private Sub ObtenerDatosCliente()

    Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
            
    adoComm.CommandText = "{ call up_ACSelDatosParametro(36,'" & strCodTipoDocumento & "','" & Trim(txtNumDocumento.Text) & "','01') }"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        lblCodCliente.Caption = Trim(adoRegistro("CodUnico"))
        lblDescripCliente.Caption = Trim(adoRegistro("DescripCliente"))
    End If
    adoRegistro.Close: Set adoRegistro = Nothing
    
End Sub

Public Sub Primero()

End Sub

Public Sub Salir()

    Dim adoRegistro                         As ADODB.Recordset
    Dim objParticipeContratoDetalleXML      As DOMDocument60
    Dim strParticipeContratoDetalleXML      As String
    Dim strMsgError                         As String
    
    Call XMLADORecordset(objParticipeContratoDetalleXML, "ParticipeContratoDetalle", "Detalle", adoRegistroAux, strMsgError)
    strParticipeContratoDetalleXML = objParticipeContratoDetalleXML.xml 'CrearXMLDetalle(objTipoCambioReemplazoXML)

    Set adoRegistro = New ADODB.Recordset
    
    adoComm.CommandText = "{ call up_PRProcParticipeContratoDetalle('" & gstrCodParticipe & "','" & strParticipeContratoDetalleXML & "') }"
    adoComm.Execute
    
    adoComm.CommandText = "{ call up_ACSelDatosParametro(12,'" & gstrCodParticipe & "') }"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        lblParticipe.Caption = Trim(adoRegistro("DescripParticipe"))
        frmContratoParticipe.lblDescripParticipe = Trim(adoRegistro("DescripParticipe"))
        frmContratoParticipe.lblTipoMancomuno = Trim(adoRegistro("DescripTipoMancomuno"))
    End If
    adoRegistro.Close: Set adoRegistro = Nothing

    Unload Me
    
End Sub

Public Sub Seguridad()

End Sub

Public Sub Siguiente()

End Sub

Public Sub Ultimo()

End Sub

Private Sub cboTipoDocumento_Click()

    strCodTipoDocumento = ""
    If cboTipoDocumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoDocumento = Trim(garrTipoDocumento(cboTipoDocumento.ListIndex))
    
End Sub

'Private Sub cboTipoMancomuno_Click()
'
'    strCodTipoMancomuno = ""
'    If cboTipoMancomuno.ListIndex < 0 Then Exit Sub
'
'    strCodTipoMancomuno = Trim(arrTipoMancomuno(cboTipoMancomuno.ListIndex))
'
''    If strCodTipoMancomuno = Codigo_Tipo_Mancomuno_Indistinto Then
''        txtPorcenParticipacion.Visible = True
''        lblDescrip(5).Visible = True
''    Else
''        txtPorcenParticipacion.Visible = False
''        txtPorcenParticipacion.Text = "0"
''        lblDescrip(5).Visible = False
''    End If
'
''Codigo_Tipo_Mancomuno_Individual = "01"
''Codigo_Tipo_Mancomuno_Indistinto = "03"
''Codigo_Tipo_Mancomuno_Conjunto = "02"
'
'End Sub

Private Sub cmdActualizar_Click()


    'VALIDAR QUE EXISTA REGISTRO
    If adoRegistroAux.RecordCount = 0 Then
        MsgBox "No puede editar un movimiento si no existen registros en el detalle del asiento!", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If adoRegistroAux.EOF Then
        MsgBox "Debe seleccionar un movimiento para editar!", vbInformation, Me.Caption
        Exit Sub
    End If
    
    
    adoRegistroAux.Fields("NumSecuencial") = 0
    adoRegistroAux.Fields("CodCliente") = Trim(lblCodCliente.Caption)
    adoRegistroAux.Fields("TipoIdentidad") = strCodTipoDocumento
    adoRegistroAux.Fields("DescripTipoIdentidad") = cboTipoDocumento.List(cboTipoDocumento.ListIndex)
    adoRegistroAux.Fields("NumIdentidad") = Trim(txtNumDocumento.Text)
    adoRegistroAux.Fields("DescripCliente") = Trim(lblDescripCliente.Caption)
    adoRegistroAux.Fields("TipoMancomuno") = strCodTipoMancomuno
    adoRegistroAux.Fields("DescripTipoMancomuno") = Valor_Caracter
    adoRegistroAux.Fields("PorcenParticipacion") = CDbl(txtPorcenParticipacion.Value)
    

End Sub

Private Sub cmdAgregar_Click()

    Dim adoRegistro As ADODB.Recordset
    Dim dblBookmark As Double
    
    'If strEstado = Reg_Consulta Then Exit Sub
    
    'If strEstado = Reg_Adicion Or strEstado = Reg_Edicion Then
        If TodoOK() Then
           
           
            adoRegistroAux.AddNew
            adoRegistroAux.Fields("NumSecuencial") = 0
            adoRegistroAux.Fields("CodCliente") = Trim(lblCodCliente.Caption)
            adoRegistroAux.Fields("TipoIdentidad") = strCodTipoDocumento
            adoRegistroAux.Fields("DescripTipoIdentidad") = Trim(cboTipoDocumento.List(cboTipoDocumento.ListIndex))
            adoRegistroAux.Fields("NumIdentidad") = Trim(txtNumDocumento.Text)
            adoRegistroAux.Fields("DescripCliente") = Trim(lblDescripCliente.Caption)
            adoRegistroAux.Fields("TipoMancomuno") = strCodTipoMancomuno
            adoRegistroAux.Fields("DescripTipoMancomuno") = Valor_Caracter
            adoRegistroAux.Fields("PorcenParticipacion") = CDbl(txtPorcenParticipacion.Value)
            
            adoRegistroAux.Update
            
            dblBookmark = adoRegistroAux.Bookmark
            
            tdgConsulta.Refresh
            
            Call NumerarRegistros
            
            Call ActualizaDescripParticipe
            
            adoRegistroAux.Bookmark = dblBookmark
            
            cmdQuitar.Enabled = True
        
        End If
    'End If
    
    
End Sub
Private Sub ActualizaDescripParticipe()


    Dim n As Long
    Dim adoConsulta             As ADODB.Recordset
    Dim strDescripTipoMancomuno As String
    Dim strDescripParticipe     As String
    
    n = 1
    
    If Not (adoRegistroAux.EOF And adoRegistroAux.BOF) Then
        adoRegistroAux.MoveFirst
    End If
    
    Set adoConsulta = New ADODB.Recordset
    
    'Obtiene tipo de mancomuno
    strDescripTipoMancomuno = Valor_Caracter

    adoComm.CommandText = "SELECT ValorParametro AS DescripTipoMancomuno FROM AuxiliarParametro " & _
                          "WHERE " & _
                          "CodTipoParametro = 'TIPMAN' AND " & _
                          "CodParametro = '" & strCodTipoMancomuno & "'"  'adoRegistroAux.Fields("TipoMancomuno") & "'"
    Set adoConsulta = adoComm.Execute
    
    If Not adoConsulta.EOF Then
        strDescripTipoMancomuno = Trim(adoConsulta.Fields("DescripTipoMancomuno"))
    End If
    
    adoConsulta.Close
    
    If adoRegistroAux.RecordCount = 1 Then
            
        adoComm.CommandText = "SELECT ApellidoPaterno, ApellidoMaterno, Nombres FROM Cliente " & _
                              "WHERE CodUnico = '" & adoRegistroAux.Fields("CodCliente") & "'"
        Set adoConsulta = adoComm.Execute
    
        If Not adoConsulta.EOF Then
            strDescripParticipe = Trim(adoConsulta.Fields("ApellidoPaterno")) + " " + Trim(adoConsulta.Fields("ApellidoMaterno")) + " " + Trim(adoConsulta.Fields("Nombres"))
        End If
        adoConsulta.Close

    Else
    
        While Not adoRegistroAux.EOF
            
            adoComm.CommandText = "SELECT ApellidoPaterno, ApellidoMaterno, Nombres FROM Cliente " & _
                                  "WHERE CodUnico = '" & adoRegistroAux.Fields("CodCliente") & "'"
            Set adoConsulta = adoComm.Execute
        
            If Not adoConsulta.EOF Then
                If adoRegistroAux.Fields("NumSecuencial") = 1 Then
                    strDescripParticipe = Trim(adoConsulta.Fields("ApellidoPaterno")) + " " + Mid(adoConsulta.Fields("ApellidoMaterno"), 1, 1) + ". " + Trim(adoConsulta.Fields("Nombres"))
                ElseIf adoRegistroAux.Fields("NumSecuencial") > 1 Then
                    strDescripParticipe = strDescripParticipe + " " + strDescripTipoMancomuno + " " + Trim(adoConsulta.Fields("ApellidoPaterno")) + " " + Mid(adoConsulta.Fields("ApellidoMaterno"), 1, 1) + ". " + Trim(adoConsulta.Fields("Nombres"))
                End If
            End If
            adoConsulta.Close
            
            adoRegistroAux.MoveNext
        
        Wend

    End If
    
    lblParticipe.Caption = Trim(strDescripParticipe)
    

End Sub

Private Sub NumerarRegistros()

    Dim n As Long
    
    n = 1
    
    If Not adoRegistroAux.EOF And Not adoRegistroAux.BOF Then
        adoRegistroAux.MoveFirst
    End If
    
    While Not adoRegistroAux.EOF
        adoRegistroAux.Fields("NumSecuencial") = n
        adoRegistroAux.Update
        n = n + 1
        adoRegistroAux.MoveNext
    Wend


End Sub
Private Function TodoOK() As Boolean

    TodoOK = False
            
    Dim adoRegistro As ADODB.Recordset
    
    If Trim(txtNumDocumento.Text) = Valor_Caracter Then
        MsgBox "Debe seleccionar el Cliente.", vbCritical
        cmdBusqueda.SetFocus
        Exit Function
    End If
        
'    If cboTipoMancomuno.ListIndex = 0 Then
'        MsgBox "Debe seleccionar el Tipo de Mancomunado.", vbCritical
'        cboTipoMancomuno.SetFocus
'        Exit Function
'    End If
    
    Set adoRegistro = New ADODB.Recordset
    
'    adoComm.CommandText = "{ call up_ACSelDatosParametro(15,'" & gstrCodParticipe & "','" & Trim(lblCodCliente.Caption) & "') }"
'    Set adoRegistro = adoComm.Execute
'
'    If Not adoRegistro.EOF Then
'        MsgBox "Cliente ya se encuentra registrado.", vbCritical, gstrNombreEmpresa
'        Call InicializarValores
'        adoRegistro.Close: Set adoRegistro = Nothing
'        Exit Function
'    End If
'
'    adoRegistro.Close
    
    '/**/
    
    adoComm.CommandText = "SELECT * FROM Cliente where ClaseCliente='02' and CodUnico='" + Trim(lblCodCliente.Caption) + "'"
    Set adoRegistro = adoComm.Execute
    If Not adoRegistro.EOF Then
        MsgBox "No se puede hacer un Mancomunado con una Persona Juridica.", vbCritical, gstrNombreEmpresa
        adoRegistro.Close: Set adoRegistro = Nothing
        Exit Function
    End If
    
    '/**/
                                                        
    '*** Si todo paso OK ***
    TodoOK = True

End Function
Private Sub cmdBusqueda_Click()

    intNumSecuencial = 0
    gstrFormulario = "frmMancomunado"
    frmBusquedaCliente.Show vbModal
    
End Sub

Private Sub cmdQuitar_Click()

    Dim dblBookmark As Double
    
    If adoRegistroAux.RecordCount > 0 Then
    
        If CInt(tdgConsulta.Columns(0)) > 1 Then
    
            dblBookmark = adoRegistroAux.Bookmark
        
            adoRegistroAux.Delete adAffectCurrent
            
            If adoRegistroAux.EOF Then
                adoRegistroAux.MovePrevious
                tdgConsulta.MovePrevious
            End If
                
            adoRegistroAux.Update
            
            If adoRegistroAux.RecordCount = 0 Then cmdQuitar.Enabled = False
    
            If adoRegistroAux.RecordCount > 0 And Not adoRegistroAux.BOF And Not adoRegistroAux.EOF And dblBookmark > 1 Then adoRegistroAux.Bookmark = dblBookmark - 1
            
            Call NumerarRegistros
            
            Call ActualizaDescripParticipe
            
            If adoRegistroAux.RecordCount > 0 And Not adoRegistroAux.BOF And Not adoRegistroAux.EOF Then adoRegistroAux.Bookmark = dblBookmark - 1
       
            tdgConsulta.Refresh
        Else
            MsgBox "No se puede eliminar el titular", vbCritical, gstrNombreEmpresa
        End If
    
    End If
    
    
End Sub

Private Sub dgdConsulta_Click()

End Sub

Private Sub Form_Deactivate()

    'ReDim garrTipoDocumento(0)
    'Call Salir
    
End Sub

Private Sub Form_Load()
   
    Call InicializarValores
    Call CargarListas
    Call CargarReportes
    Call Buscar
    Call DarFormato
    
    CentrarForm Me
        
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
            
End Sub
Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = ""
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = ""
    
End Sub
Private Sub CargarListas()

    Dim strSql  As String
    
    '*** Tipo Documento Identidad  - Naturales ***
    strSql = "{ call up_ACSelDatosParametro(4,'" & Codigo_Persona_Natural & "') }"
    CargarControlLista strSql, cboTipoDocumento, garrTipoDocumento(), Sel_Defecto
    
    If cboTipoDocumento.ListCount > 0 Then cboTipoDocumento.ListIndex = 0
        
    '*** Tipo de Mancomuno ***
'    strSql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPMAN' AND CodParametro<>'01' ORDER BY DescripParametro"
'    CargarControlLista strSql, cboTipoMancomuno, arrTipoMancomuno(), Sel_Defecto
'
'    If cboTipoMancomuno.ListCount > 0 Then cboTipoMancomuno.ListIndex = 0
        
End Sub
Private Sub InicializarValores()

    strEstado = Reg_Defecto
        
    cboTipoDocumento.ListIndex = -1
    If cboTipoDocumento.ListCount > 0 Then cboTipoDocumento.ListIndex = 0
    
    txtNumDocumento.Text = ""
    lblDescripCliente.Caption = ""
    lblCodCliente.Caption = ""
    
    'cboTipoMancomuno.ListIndex = -1
    'If cboTipoMancomuno.ListCount > 0 Then cboTipoMancomuno.ListIndex = 0
    
    '*** Verificando Nivel de Acceso de Usuario ***
'    strNivAcceso = AccesoForm(gstrNomOpc, gstrNumInd)

    Set cmdOpcion.FormularioActivo = Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set frmMancomunado = Nothing
        
End Sub



Private Sub tdgConsulta_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    Dim intRegistro As Integer

    If Not (adoRegistroAux.EOF And adoRegistroAux.BOF) Then

        lblCodCliente.Caption = adoRegistroAux.Fields("CodCliente")
        txtNumDocumento.Text = adoRegistroAux.Fields("NumIdentidad")
        lblDescripCliente.Caption = adoRegistroAux.Fields("DescripCliente")
        txtPorcenParticipacion.Text = CStr(adoRegistroAux.Fields("PorcenParticipacion"))
        
        intRegistro = ObtenerItemLista(garrTipoDocumento(), adoRegistroAux.Fields("TipoIdentidad"))
        If intRegistro >= 0 Then cboTipoDocumento.ListIndex = intRegistro
    End If
    

End Sub

Private Sub txtNumDocumento_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call ObtenerDatosCliente
    End If
    
End Sub

Private Sub ConfiguraRecordsetAuxiliar()

    Set adoRegistroAux = New ADODB.Recordset

    With adoRegistroAux
       .CursorLocation = adUseClient
       .Fields.Append "NumSecuencial", adInteger, 10
       .Fields.Append "CodCliente", adVarChar, 20
       .Fields.Append "TipoIdentidad", adChar, 2
       .Fields.Append "DescripTipoIdentidad", adVarChar, 30
       .Fields.Append "NumIdentidad", adVarChar, 15
       .Fields.Append "DescripCliente", adVarChar, 75
       .Fields.Append "TipoMancomuno", adChar, 2
       .Fields.Append "DescripTipoMancomuno", adVarChar, 30
       .Fields.Append "PorcenParticipacion", adDecimal, 8
       .LockType = adLockBatchOptimistic
    End With

    With adoRegistroAux.Fields.Item("PorcenParticipacion")
        .Precision = 8
        .NumericScale = 4
    End With
    
    adoRegistroAux.Open

End Sub
