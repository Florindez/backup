VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmRepresentanteParticipe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Representante Contrato"
   ClientHeight    =   6000
   ClientLeft      =   1050
   ClientTop       =   4560
   ClientWidth     =   9270
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9270
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   600
      TabIndex        =   15
      Top             =   5160
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1296
      Buttons         =   2
      Caption0        =   "Con&sultar"
      Tag0            =   "3"
      ToolTipText0    =   "Consultar"
      Caption1        =   "&Cerrar"
      Tag1            =   "9"
      ToolTipText1    =   "Cerrar Ventana"
      UserControlWidth=   2700
   End
   Begin VB.Frame fraRepresentanteContrato 
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   200
      Width           =   8775
      Begin VB.TextBox txtTipoDocumento 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2130
         TabIndex        =   23
         Top             =   1860
         Width           =   2205
      End
      Begin VB.TextBox txtCodTipoDocumento 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4410
         TabIndex        =   22
         Top             =   1860
         Width           =   795
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmRepresentanteParticipe.frx":0000
         Height          =   1335
         Left            =   930
         OleObjectBlob   =   "frmRepresentanteParticipe.frx":001A
         TabIndex        =   20
         Top             =   3165
         Width           =   7575
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
         Left            =   8040
         TabIndex        =   18
         ToolTipText     =   "Búsqueda de Representante"
         Top             =   690
         Width           =   375
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   ">"
         Height          =   375
         Left            =   330
         TabIndex        =   14
         Top             =   3285
         Width           =   375
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "<"
         Height          =   375
         Left            =   330
         TabIndex        =   13
         Top             =   3945
         Width           =   375
      End
      Begin VB.TextBox txtNumDocumento 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2130
         TabIndex        =   2
         Top             =   2235
         Width           =   3045
      End
      Begin VB.ComboBox cboTipoDocumento 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6180
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1710
         Width           =   2325
      End
      Begin MSComCtl2.DTPicker dtpFechaIngreso 
         Height          =   315
         Left            =   2130
         TabIndex        =   3
         Top             =   1095
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         _Version        =   393216
         Format          =   102629377
         CurrentDate     =   38069
      End
      Begin MSComCtl2.DTPicker dtpFechaSalida 
         Height          =   315
         Left            =   2130
         TabIndex        =   4
         Top             =   1455
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         _Version        =   393216
         Format          =   102629377
         CurrentDate     =   38069
      End
      Begin VB.Label lblCodClienteParticipe 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6180
         TabIndex        =   21
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lblCodCliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6180
         TabIndex        =   19
         Top             =   1410
         Width           =   2295
      End
      Begin VB.Label lblCargo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2130
         TabIndex        =   17
         Top             =   2600
         Width           =   6375
      End
      Begin VB.Label lblDescripCliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2130
         TabIndex        =   16
         Top             =   690
         Width           =   5760
      End
      Begin VB.Label lblParticipe 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2130
         TabIndex        =   12
         Top             =   300
         Width           =   6285
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Partícipe"
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
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fecha Ingreso"
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
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   1110
         Width           =   1635
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fecha Salida"
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
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1470
         Width           =   1635
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Cargo"
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
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   2620
         Width           =   1635
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Tipo Documento"
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
         Index           =   4
         Left            =   210
         TabIndex        =   7
         Top             =   1890
         Width           =   1635
      End
      Begin VB.Label lblDescrip 
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
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   6
         Top             =   2265
         Width           =   1635
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Representante"
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
         Index           =   6
         Left            =   240
         TabIndex        =   5
         Top             =   750
         Width           =   1635
      End
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   6360
      Top             =   5160
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepresentanteParticipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strCodTipoDocumento     As String, intNumSecuencial        As Integer
Dim strEstado               As String
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

    Dim strSql As String
                                                                                    
    Me.MousePointer = vbHourglass
                    
    strSql = "{ call up_ACSelDatosParametro(37,'" & gstrCodParticipe & "') }"
    
    With adoConsulta
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSql
        .Refresh
    End With
    
    tdgConsulta.Refresh
    
    If adoConsulta.Recordset.RecordCount > 0 Then strEstado = Reg_Consulta
        
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
        
        txtNumDocumento.Text = Trim(tdgConsulta.Columns(3))
        lblDescripCliente.Caption = Trim(tdgConsulta.Columns(5))
                        
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
        lblCargo.Caption = Trim(adoRegistro("CargoCliente"))
    End If
    adoRegistro.Close: Set adoRegistro = Nothing
    
End Sub

Public Sub Primero()

End Sub

Public Sub Refrescar()

End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub Seguridad()

End Sub

Public Sub Siguiente()

End Sub

Public Sub Ultimo()

End Sub

'Private Sub cboTipoDocumento_Click()

'    strCodTipoDocumento = ""
'    If cboTipoDocumento.ListIndex < 0 Then Exit Sub
    
'    strCodTipoDocumento = Trim(garrTipoDocumento(cboTipoDocumento.ListIndex))
    
'End Sub

Private Function TodoOK() As Boolean

    TodoOK = False
            
    If Trim(txtNumDocumento.Text) = Valor_Caracter Then
        MsgBox "Debe seleccionar el Representante.", vbCritical
        cmdBusqueda.SetFocus
        Exit Function
    End If
            
    '*** Si todo paso OK ***
    TodoOK = True

End Function


Private Sub cmdAgregar_Click()

    Dim adoRegistro As ADODB.Recordset
    Dim intRegistro As Integer
    
    If TodoOK() Then
        Set adoRegistro = New ADODB.Recordset
        
        adoComm.CommandText = "{ call up_ACSelDatosParametro(38,'" & gstrCodParticipe & "','" & Trim(lblCodCliente.Caption) & "') }"
        Set adoRegistro = adoComm.Execute
        
        If Not adoRegistro.EOF Then
            MsgBox "Cliente ya se encuentra registrado.", vbCritical, gstrNombreEmpresa
            Call InicializarValores
            adoRegistro.Close: Set adoRegistro = Nothing
            Exit Sub
        End If
        adoRegistro.Close
                
        With adoComm
            .CommandText = "{ call up_PRManRepresentanteParticipe('"
            .CommandText = .CommandText & gstrCodParticipe & "',"
            .CommandText = .CommandText & intNumSecuencial & ",'"
            .CommandText = .CommandText & Trim(lblCodCliente.Caption) & "','"
            .CommandText = .CommandText & Convertyyyymmdd(dtpFechaIngreso.Value) & "','"
            .CommandText = .CommandText & Convertyyyymmdd(dtpFechaSalida.Value) & "','"
            .CommandText = .CommandText & Trim(txtCodTipoDocumento.Text) & "','"
            .CommandText = .CommandText & Trim(txtNumDocumento.Text) & "','"
            .CommandText = .CommandText & "I') }"
            'MsgBox .CommandText, vbCritical
            adoConn.Execute .CommandText
                                                            
        End With
            
        Call Buscar
    End If
    
End Sub

Private Sub cmdBusqueda_Click()

    intNumSecuencial = 0
    gstrFormulario = "frmRepresentanteParticipe"
    frmBusquedaRepresentante.Caption = "Búsqueda de Representantes"
    frmBusquedaRepresentante.lblCodClienteParticipe = Trim(lblCodClienteParticipe.Caption)
    frmBusquedaRepresentante.Show vbModal
End Sub

Private Sub cmdQuitar_Click()
On Error GoTo Error1            '/**/ HMC Habilitamos la rutina de Errores.
    
    With adoComm
        
        .CommandText = "DELETE RepresentanteParticipe "
        .CommandText = .CommandText & "WHERE CodParticipe='" & gstrCodParticipe & "' AND NumSecuencial=" & CInt(tdgConsulta.Columns(0))
        
        adoConn.Execute .CommandText
                                               
    End With
    
    Call Buscar
    
On Error GoTo 0                  '/**/
Exit Sub                         '/**/
Error1:     MsgBox DescripcionError & vbNewLine & DescripcionTecnica & err.Description, vbExclamation, TituloError ' Mostrar Error
        
End Sub


Private Sub dgdConsulta_Click()

End Sub

Private Sub Form_Deactivate()

    ReDim garrTipoDocumento(0)
    Call Salir
    
End Sub

Private Sub Form_Load()
    
    Call InicializarValores
    Call CargarListas
    Call CargarReportes
    Call Buscar
    Call DarFormato
    
    CentrarForm Me
    
'    aGrdCnf(1).TitDes = "Nro. Relac"
'    aGrdCnf(1).DatNom = "NRO_RELA"
'
'    aGrdCnf(2).TitDes = "Fec. Ingreso"
'    aGrdCnf(2).DatNom = "FCH_IREL"
'
'    aGrdCnf(3).TitDes = "Fec. Salida"
'    aGrdCnf(3).DatNom = "FCH_FREL"
'
'    aGrdCnf(4).TitDes = "Tipo Documento"
'    aGrdCnf(4).DatNom = "TIP_IDEN"
'
'    aGrdCnf(5).TitDes = "Nro. Documento"
'    aGrdCnf(5).DatNom = "NRO_IDEN"
'
'    aGrdCnf(6).TitDes = "Descrip. Cargo"
'    aGrdCnf(6).DatNom = "DSC_CARG"
'
'    aGrdCnf(7).TitDes = "Descrip. Pers"
'    aGrdCnf(7).DatNom = "DSC_PERS"

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
                    
End Sub
Private Sub InicializarValores()

    strEstado = Reg_Defecto
        
    dtpFechaIngreso.Value = gdatFechaActual
    dtpFechaSalida.Value = gdatFechaActual
    
    cboTipoDocumento.ListIndex = -1
    If cboTipoDocumento.ListCount > 0 Then cboTipoDocumento.ListIndex = 0
    
    txtNumDocumento.Text = ""
    
    '*** Verificando Nivel de Acceso de Usuario ***
'    strNivAcceso = AccesoForm(gstrNomOpc, gstrNumInd)

    Set cmdOpcion.FormularioActivo = Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Unload Me

End Sub


Private Sub txtNumDocumento_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call ObtenerDatosCliente
    End If
   
End Sub

