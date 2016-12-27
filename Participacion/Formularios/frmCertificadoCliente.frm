VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmCertificadoCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emisión y Control de Certificados de Participación"
   ClientHeight    =   6690
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   10095
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   7680
      TabIndex        =   6
      Top             =   5760
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      Visible0        =   0   'False
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   840
      TabIndex        =   5
      Top             =   5760
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1296
      Buttons         =   2
      Caption0        =   "&Modificar"
      Tag0            =   "3"
      Visible0        =   0   'False
      ToolTipText0    =   "Modificar"
      Caption1        =   "&Buscar"
      Tag1            =   "5"
      Visible1        =   0   'False
      ToolTipText1    =   "Buscar"
      UserControlWidth=   2700
   End
   Begin TabDlg.SSTab tabBloqueo 
      Height          =   5460
      Left            =   165
      TabIndex        =   7
      Top             =   165
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   9631
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmCertificadoCliente.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCertificado"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmCertificadoCliente.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(1)=   "fraOperacion"
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -69000
         TabIndex        =   20
         Top             =   4560
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   1296
         Buttons         =   2
         Caption0        =   "&Guardar"
         Tag0            =   "2"
         Visible0        =   0   'False
         ToolTipText0    =   "Guardar"
         Caption1        =   "&Cancelar"
         Tag1            =   "8"
         Visible1        =   0   'False
         ToolTipText1    =   "&Cancelar"
         UserControlWidth=   2700
      End
      Begin VB.Frame fraOperacion 
         Caption         =   "Fondo"
         Height          =   3975
         Left            =   -74760
         TabIndex        =   13
         Top             =   480
         Width           =   9255
         Begin VB.ComboBox cboEstadoOperacion 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   3300
            Width           =   2535
         End
         Begin VB.TextBox txtObservacion 
            Height          =   1095
            Left            =   1860
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   1995
            Width           =   7095
         End
         Begin VB.ComboBox cboTipoBloqueo 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1470
            Width           =   2535
         End
         Begin VB.Label lblCantCuotas 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6400
            TabIndex        =   35
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cuotas"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   4920
            TabIndex        =   34
            Top             =   960
            Width           =   495
         End
         Begin VB.Label lblNumCertificado 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1860
            TabIndex        =   33
            Top             =   975
            Width           =   2535
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Num. Certificado"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   240
            TabIndex        =   32
            Top             =   995
            Width           =   1170
         End
         Begin VB.Label lblDescripParticipeDetalle 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1860
            TabIndex        =   31
            Top             =   480
            Width           =   7095
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Partícipe"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   30
            Top             =   500
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Estado"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   10
            Left            =   240
            TabIndex        =   18
            Top             =   3320
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Observaciones"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   9
            Left            =   240
            TabIndex        =   16
            Top             =   2015
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Bloqueo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   15
            Top             =   1490
            Width           =   945
         End
      End
      Begin VB.Frame fraCertificado 
         Height          =   2655
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   9225
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
            Left            =   8360
            TabIndex        =   1
            ToolTipText     =   "Búsqueda de Partícipe"
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox txtNumDocumento 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2100
            TabIndex        =   22
            Top             =   1458
            Width           =   3375
         End
         Begin VB.ComboBox cboTipoDocumento 
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1077
            Width           =   3375
         End
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   315
            Width           =   6660
         End
         Begin VB.ComboBox cboEstado 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   2160
            Width           =   3375
         End
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   285
            Left            =   7245
            TabIndex        =   3
            Top             =   1458
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   503
            _Version        =   393216
            Format          =   175570945
            CurrentDate     =   38069
         End
         Begin MSComCtl2.DTPicker dtpFechaHasta 
            Height          =   285
            Left            =   7245
            TabIndex        =   4
            Top             =   1815
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   503
            _Version        =   393216
            Format          =   175570945
            CurrentDate     =   38069
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Partícipe"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   28
            Top             =   735
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Partícipe"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   27
            Top             =   1835
            Width           =   1005
         End
         Begin VB.Label lblDescripTipoParticipe 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2100
            TabIndex        =   26
            Top             =   1815
            Width           =   3375
         End
         Begin VB.Label lblDescripParticipe 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2100
            TabIndex        =   25
            Top             =   720
            Width           =   6225
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Documento"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   360
            TabIndex        =   24
            Top             =   1097
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Num.Documento"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   360
            TabIndex        =   23
            Top             =   1478
            Width           =   1200
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fondo"
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
            Height          =   285
            Index           =   4
            Left            =   360
            TabIndex        =   12
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Hasta"
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
            Height          =   285
            Index           =   1
            Left            =   6180
            TabIndex        =   11
            Top             =   1835
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Desde"
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
            Height          =   285
            Index           =   0
            Left            =   6180
            TabIndex        =   10
            Top             =   1478
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Estado"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   7
            Left            =   360
            TabIndex        =   9
            Top             =   2180
            Width           =   1455
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmCertificadoCliente.frx":0038
         Height          =   1695
         Left            =   240
         OleObjectBlob   =   "frmCertificadoCliente.frx":0052
         TabIndex        =   29
         Top             =   3240
         Width           =   9225
      End
   End
End
Attribute VB_Name = "frmCertificadoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()              As String, arrEstado()          As String
Dim arrEstadoOperacion()    As String, arrTipoBloqueo()     As String

Dim strCodFondo             As String, strCodEstado         As String
Dim strCodestadoOperacion   As String, strCodTipoBloqueo    As String
Dim strCodTipoDocumento     As String, strCodMoneda         As String
Dim strFechaDesde           As String, strFechaHasta        As String
Dim strEstado               As String, strSql               As String
Dim adoConsulta             As ADODB.Recordset
Dim adoCertificados         As ADODB.Recordset
Dim indSortAsc              As Boolean, indSortDesc         As Boolean

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

Private Sub ActualizaDatosCertificado(strSql As String, strNumCertificado As String)

    With adoComm
        .CommandText = "UPDATE ParticipeCertificado SET " & strSql & _
            "FechaActualiza='" & Convertyyyymmdd(gdatFechaActual) & "'," & _
            "UsuarioEdicion='" & gstrLogin & "'," & _
            "FechaEdicion='" & Convertyyyymmdd(gdatFechaActual) & "' " & _
            "WHERE (FechaSuscripcion >='" & strFechaDesde & "' AND FechaSuscripcion <'" & strFechaHasta & "') AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
            "NumCertificado='" & strNumCertificado & "'"
            
        adoConn.Execute .CommandText
    End With
    
End Sub

Public Sub Adicionar()
        
End Sub

Private Sub Deshabilita()

    cboEstadoOperacion.Enabled = False
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim strSql As String
    Dim intRegistro As Integer
    
    Select Case strModo
    
        Case Reg_Adicion
            txtObservacion.Text = Valor_Caracter
            
            cboTipoBloqueo.ListIndex = -1
            If cboTipoBloqueo.ListCount > 0 Then cboTipoBloqueo.ListIndex = 0

            cboEstadoOperacion.ListIndex = -1
            intRegistro = ObtenerItemLista(arrEstadoOperacion(), Estado_Activo)
            If intRegistro >= 0 Then cboEstadoOperacion.ListIndex = intRegistro
                                                                                    
            cboTipoBloqueo.SetFocus
                        
        Case Reg_Edicion
        
            Dim adoRegistro As ADODB.Recordset
            Dim adoRegistro2 As ADODB.Recordset
            Dim strFechaInicio  As String, strFechaFin  As String
            
            
            Set adoRegistro2 = New ADODB.Recordset

                adoComm.CommandText = "Select * from ParticipeCertificadoBloqueo " & _
                "where CodParticipe ='" & gstrCodParticipe & "' AND NumCertificado = '" & Trim(tdgConsulta.Columns(0)) & "' and " & _
                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            
            Set adoRegistro2 = adoComm.Execute

            If adoRegistro2.EOF = False Then
                cboEstadoOperacion.ListIndex = -1
                intRegistro = ObtenerItemLista(arrEstadoOperacion(), adoRegistro2("Estado").Value)
                If intRegistro >= 0 Then cboEstadoOperacion.ListIndex = intRegistro
                cboEstadoOperacion.Enabled = False
            End If
            adoRegistro2.Close: Set adoRegistro2 = Nothing
            
            
            '*** Valores Iniciales ***
            cboTipoBloqueo.Enabled = True
            cboTipoBloqueo.ListIndex = -1
            If cboTipoBloqueo.ListCount > 0 Then cboTipoBloqueo.ListIndex = 0
            
            txtObservacion.Text = Valor_Caracter
            txtObservacion.Enabled = False
            lblDescripParticipeDetalle.Caption = Trim(lblDescripParticipe.Caption)
            lblNumCertificado.Caption = Trim(tdgConsulta.Columns(0))
            lblCantCuotas.Caption = CStr(tdgConsulta.Columns(2))
            cboTipoBloqueo.SetFocus
            
            
            If Trim(tdgConsulta.Columns(8)) <> Valor_Caracter Then
                Set adoRegistro = New ADODB.Recordset
                
                strFechaInicio = Convertyyyymmdd(CVDate(tdgConsulta.Columns(8)))
                strFechaFin = Convertyyyymmdd(DateAdd("d", 1, CVDate(tdgConsulta.Columns(8))))

'            adoComm.CommandText = "SELECT PCB.*,DescripParticipe,AP1.DescripParametro TipoIdentidad,PCD.NumIdentidad,PCD.TipoIdentidad CodIdentidad,AP2.DescripParametro DescripMancomuno,PC.TipoMancomuno " & _
'                "FROM ParticipeCertificadoBloqueo PCB JOIN ParticipeContrato PC ON(PC.CodParticipe=PCB.CodParticipe) " & _
'                "JOIN ParticipeContratoDetalle PCD ON(PCD.CodParticipe=PC.CodParticipe) " & _
'                "JOIN AuxiliarParametro AP1 ON(AP1.CodParametro=PCD.TipoIdentidad AND AP1.CodTipoParametro='TIPIDE') " & _
'                "JOIN AuxiliarParametro AP2 ON(AP2.CodParametro=PC.TipoMancomuno AND AP2.CodTipoParametro='TIPMAN') " & _
'                "WHERE (PCB.FechaBloqueo >='" & strFechaInicio & "' AND PCB.FechaBloqueo <'" & strFechaFin & "') AND " & _
'                "PCB.CodParticipe='" & gstrCodParticipe & "' AND " & _
'                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' "

                adoComm.CommandText = "SELECT * " & _
                    "FROM ParticipeCertificadoBloqueo " & _
                    "WHERE (FechaBloqueo >='" & strFechaInicio & "' AND FechaBloqueo <'" & strFechaFin & "') AND " & _
                    "CodParticipe='" & gstrCodParticipe & "' AND NumCertificado='" & lblNumCertificado.Caption & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' "
                Set adoRegistro = adoComm.Execute

                If Not adoRegistro.EOF Then
                    cboTipoBloqueo.ListIndex = -1
                    intRegistro = ObtenerItemLista(arrTipoBloqueo(), adoRegistro("TipoBloqueo"))
                    If intRegistro >= 0 Then cboTipoBloqueo.ListIndex = intRegistro
                    cboTipoBloqueo.Enabled = False
    
                    txtObservacion.Text = adoRegistro("DescripMotivo")
    
                    cboEstadoOperacion.ListIndex = -1
                    intRegistro = ObtenerItemLista(arrEstadoOperacion(), adoRegistro("Estado"))
                    If intRegistro >= 0 Then cboEstadoOperacion.ListIndex = intRegistro
                    cboEstadoOperacion.Enabled = True
    
'                    txtObservacion.SetFocus
                End If
                adoRegistro.Close: Set adoRegistro = Nothing
            End If
    End Select
            
End Sub
            


Public Sub Ayuda()

End Sub

Public Sub Buscar()
        
    strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
    strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value))
            
    strSql = "SELECT PC.NumCertificado,FechaSuscripcion,FechaOperacion,ValorCuota,PC.CantCuotas,PC.CantCuotasPagadas,FechaModifica=CASE Convert(char,PCB.FechaDesbloqueo,112) WHEN '19000101' THEN PCB.FechaBloqueo ELSE PCB.FechaDesbloqueo END," & _
        "CASE IndCustodia WHEN 'X' THEN 'SI' ELSE 'NO' END DescripCustodia,IndCustodia,DescripParametro,PCB.FechaBloqueo " & _
        "FROM ParticipeCertificado PC LEFT JOIN ParticipeCertificadoBloqueo PCB " & _
        "ON(PCB.NumCertificado=PC.NumCertificado AND PCB.CodParticipe=PC.CodParticipe AND PCB.CodFondo=PC.CodFondo AND PCB.CodAdministradora=PC.CodAdministradora) " & _
        "LEFT JOIN AuxiliarParametro AP ON(AP.CodParametro=PCB.TipoBloqueo AND CodTipoParametro='TIPBLO') " & _
        "WHERE PC.CodParticipe='" & gstrCodParticipe & "' AND PC.CodFondo='" & strCodFondo & "' AND " & _
        "PC.CodAdministradora='" & gstrCodAdministradora & "' "
        
    If strCodEstado <> Valor_Caracter Then
        If strCodEstado = Estado_Certificado_Vigente Then
            strSql = strSql & "AND IndVigente='X' "
        Else
            strSql = strSql & "AND IndVigente='' "
        End If
    End If
    strSql = strSql & "ORDER BY FechaSuscripcion"
    
    Set adoCertificados = New ADODB.Recordset
    
    strEstado = Reg_Defecto
    With adoCertificados
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSql
    End With
        
    tdgConsulta.DataSource = adoCertificados
    
    If adoCertificados.RecordCount > 0 Then strEstado = Reg_Consulta
                
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabBloqueo
        .TabEnabled(0) = True
        .Tab = 0
    End With
    Call Buscar
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Emisión de Certificados"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Lista de Certificados Bloqueados"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Lista de Certificados"
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Visible = True
    
End Sub

Public Sub Eliminar()

End Sub

Public Sub Grabar()

   Dim intRegistro         As Integer, intContador     As Integer
   Dim strFechaInicio      As String, strFechaFin      As String
   Dim adoRegistro As ADODB.Recordset '
                
    If strEstado = Reg_Defecto Then Exit Sub
                
    If strEstado = Reg_Edicion Then
        If TodoOK() Then
            If MsgBox(Mensaje_Edicion, vbQuestion + vbYesNo, gstrNombreEmpresa) = vbNo Then Exit Sub

            Me.MousePointer = vbHourglass

            '*** Actualizar Operación de Bloqueo ***
            With adoComm
            
                strFechaInicio = Convertyyyymmdd(CVDate(tdgConsulta.Columns(9)))
                strFechaFin = Convertyyyymmdd(DateAdd("d", 1, CVDate(tdgConsulta.Columns(9))))

                If strCodTipoBloqueo = Codigo_Tipo_Bloqueo_Emision Then
                    .CommandText = "UPDATE ParticipeCertificado SET IndBloqueo='X',IndCustodia='' "
                Else
                    .CommandText = "UPDATE ParticipeCertificado SET IndBloqueo='X' "
                
                End If
                
                .CommandText = .CommandText & "WHERE (FechaOperacion>='" & strFechaInicio & "' AND FechaOperacion<'" & strFechaFin & "') AND " & _
                    "NumCertificado='" & Trim(tdgConsulta.Columns(0)) & "' AND CodParticipe='" & gstrCodParticipe & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                adoConn.Execute .CommandText

                    '/**/HMC
                    Set adoRegistro = New ADODB.Recordset
                    
                    With adoComm
                            .CommandText = "Select * from ParticipeCertificadoBloqueo " & _
                                           "where CodParticipe ='" & gstrCodParticipe & "' AND NumCertificado = '" & Trim(tdgConsulta.Columns(0)) & "' and " & _
                                           "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                        Set adoRegistro = .Execute
                        
                        If Not adoRegistro.EOF Then
                            .CommandText = "UPDATE ParticipeCertificadoBloqueo SET Estado ='" & strCodestadoOperacion & "'" & _
                                "where CodParticipe ='" & gstrCodParticipe & "' AND NumCertificado = '" & Trim(tdgConsulta.Columns(0)) & "' and " & _
                                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                        Else
                            .CommandText = "INSERT INTO ParticipeCertificadoBloqueo VALUES('" & _
                                strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                gstrCodParticipe & "','" & Trim(tdgConsulta.Columns(0)) & "','" & _
                                Convertyyyymmdd(gdatFechaActual) & "','" & Convertyyyymmdd(Valor_Fecha) & "'," & _
                                CDec(lblCantCuotas.Caption) & ",'" & Trim(txtObservacion.Text) & "','" & _
                                strCodTipoBloqueo & "','" & strCodestadoOperacion & "')"
                        End If
                            .Execute
                        'End If
                        adoRegistro.Close: Set adoRegistro = Nothing
                    End With
                    '/**/
            End With

                Me.MousePointer = vbDefault
                MsgBox Mensaje_Edicion_Exitosa, vbExclamation
                frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
                cmdOpcion.Visible = True
                    With tabBloqueo
                        .TabEnabled(0) = True
                        .Tab = 0
                    End With
                Call Buscar
            End If
        End If

End Sub

Private Function TodoOK() As Boolean

    TodoOK = False
        
'    If cboTipoDocumento.ListIndex = 0 Then
'        MsgBox "Seleccione el Tipo de Documento.", vbCritical
'        cboTipoDocumento.SetFocus
'        Exit Function
'    End If

    If Trim(txtNumDocumento.Text) = Valor_Caracter Then
        MsgBox "El Campo Número de Documento no es Válido!.", vbCritical
        txtNumDocumento.SetFocus
        Exit Function
    End If
    
    If Trim(lblDescripParticipe.Caption) = Valor_Caracter Then
        MsgBox "El Campo Descripción no es Válido!, presione ENTER en el campo Número de Documento.", vbCritical
        txtNumDocumento.SetFocus
        Exit Function
    End If
    
'    If CDbl(lblCuotas.Caption) = 0 Then
'        MsgBox "No se ha seleccionado ningún certificado", vbCritical
'        tdgCertificado.SetFocus
'        Exit Function
'    End If
    
    If cboTipoBloqueo.ListIndex = 0 Then
        MsgBox "Seleccione el Tipo de Bloqueo.", vbCritical
        cboTipoBloqueo.SetFocus
        Exit Function
    End If
    
    '*** Si todo paso OK ***
    TodoOK = True

End Function
Public Sub Imprimir()

End Sub

Public Sub Modificar()
        
    If strEstado = Reg_Consulta Then
        frmMainMdi.stbMdi.Panels(3).Text = "Modificar condición del certificado..."
        
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabBloqueo
            .TabEnabled(0) = False
            .Tab = 1
        End With
        'Call Habilita
    End If

End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub Seguridad()

End Sub

Public Sub SubImprimir(Index As Integer)
    
    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim adoConsulta             As ADODB.Recordset
    Dim strCodParticipe         As String

    If tabBloqueo.Tab = 1 Then Exit Sub
    
    Select Case Index
        Case 1
            gstrNameRepo = "CertificadoParticipacion"
            
            Set adoConsulta = New ADODB.Recordset
            Set frmReporte = New frmVisorReporte
            
            With adoComm
                .CommandText = "SELECT CodParticipe FROM ParticipeContrato WHERE DescripParticipe='" & Trim(lblDescripParticipe.Caption) & "'"
                Set adoConsulta = .Execute

                If Not adoConsulta.EOF Then
                    strCodParticipe = Trim(adoConsulta("CodParticipe"))
                End If
                adoConsulta.Close
            End With


            ReDim aReportParamS(4)
            ReDim aReportParamFn(1)
            ReDim aReportParamF(1)

            strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
            strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value))
            
            aReportParamFn(0) = "Fondo"
            aReportParamFn(1) = "Lugar"
            
            aReportParamF(0) = Trim(cboFondo.Text)
            aReportParamF(1) = "San Isidro"
                        
            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora
            aReportParamS(2) = strFechaDesde
            aReportParamS(3) = strFechaHasta
            aReportParamS(4) = strCodParticipe
            
        Case 2:
        
        'gstrNameRepo = "CertificadoParticipacion_Bloqueados"
            gstrNameRepo = "ListaCertificados"
                        
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(1)
            ReDim aReportParamFn(4)
            ReDim aReportParamF(4)

            'strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
            'strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value))
            
            aReportParamFn(0) = "Fondo"
            aReportParamFn(1) = "Lugar"
            aReportParamFn(2) = "NombreEmpresa"
            aReportParamFn(3) = "Usuario"
            aReportParamFn(4) = "Hora"
            
            aReportParamF(0) = Trim(cboFondo.Text)
            aReportParamF(1) = "San Isidro"
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
            aReportParamF(3) = gstrLogin
            aReportParamF(4) = Format(Time(), "hh:mm:ss")
                        
            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora
            'aReportParamS(2) = strFechaDesde
            'aReportParamS(3) = strFechaHasta
            'aReportParamS(4) = Codigo_Tipo_Bloqueo_Emision
        
    End Select

    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub

Private Sub cboEstado_Click()

    strCodEstado = ""
    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strCodEstado = Trim(arrEstado(cboEstado.ListIndex))
    
End Sub

Private Sub cboEstadoOperacion_Click()

    strCodestadoOperacion = Valor_Caracter
    If cboEstadoOperacion.ListIndex < 0 Then Exit Sub
    
    strCodestadoOperacion = Trim(arrEstadoOperacion(cboEstadoOperacion.ListIndex))
    
End Sub

Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset, adoTemporal As ADODB.Recordset
    
    strCodFondo = ""
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        
        '--------------------------------------------------------------------------LV------------
        Dim adoConsulta             As ADODB.Recordset

        Set adoConsulta = New ADODB.Recordset

        With adoComm
                .CommandText = "select FechaInicioEtapaPreOperativa from Fondo where DescripFondo='" & cboFondo.Text & "'"
                Set adoConsulta = .Execute

                If Not adoConsulta.EOF Then
                    dtpFechaDesde.Value = Trim(adoConsulta("FechaInicioEtapaPreOperativa"))
                End If
                adoConsulta.Close
        End With
        '----------------------------------------------------------------------------LV-----------
               
        If Not adoRegistro.EOF Then
            'dtpFechaDesde.Value = CVDate(adoRegistro("FechaCuota"))
            dtpFechaHasta.Value = CVDate(adoRegistro("FechaCuota")) 'dtpFechaDesde.Value
            strCodMoneda = Trim(adoRegistro("CodMoneda"))
            
            gdatFechaActual = adoRegistro("FechaCuota")
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
        
End Sub

Private Sub cboTipoBloqueo_Click()

    strCodTipoBloqueo = Valor_Caracter
    If cboTipoBloqueo.ListIndex < 0 Then Exit Sub
    
    strCodTipoBloqueo = Trim(arrTipoBloqueo(cboTipoBloqueo.ListIndex))
    
End Sub

Private Sub cboTipoDocumento_Click()

    strCodTipoDocumento = Valor_Caracter
    If cboTipoDocumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoDocumento = Trim(garrTipoDocumento(cboTipoDocumento.ListIndex))
    
End Sub

Private Sub cmdBusqueda_Click()
    
    Dim adoRegistro As ADODB.Recordset
    Dim intRegistro As Integer
    
    gstrFormulario = Me.Name
    frmBusquedaParticipeP.Show vbModal
    
   

    Set adoRegistro = New ADODB.Recordset

    adoComm.CommandText = "SELECT TipoIdentidad,NumIdentidad FROM ParticipeContrato WHERE CodParticipe= '" & gstrCodParticipe & "' "
    Set adoRegistro = adoComm.Execute
    '++REA 2015-05-21
    If Not adoRegistro.EOF Then
        intRegistro = ObtenerItemLista(garrTipoDocumento(), adoRegistro.Fields("TipoIdentidad"))
        If intRegistro >= 0 Then cboTipoDocumento.ListIndex = intRegistro
        
        txtNumDocumento.Text = Trim(adoRegistro.Fields("NumIdentidad"))
    End If
    '--REA 2015-05-21
    adoRegistro.Close: Set adoRegistro = Nothing
    Call Buscar
End Sub

Private Sub Form_Activate()

    Call CargarReportes
    
End Sub

Private Sub Form_Deactivate()

    ReDim garrTipoDocumento(0)
    Call OcultarReportes
    
End Sub

Private Sub Form_Load()
   
    Call InicializarValores
    Call CargarListas
    Call CargarReportes
    Call Buscar
    Call DarFormato
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
     
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    Dim elemento As Object
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
    For Each elemento In Me.Controls
    
        If TypeOf elemento Is TDBGrid Then
            Call FormatoGrilla(elemento)
        End If
    
    Next
            
End Sub
Private Sub CargarListas()
    
    Dim intRegistro As Integer
    
    '*** Fondos ***
    strSql = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSql, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    '*** Estado Certificado ***
    strSql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='ESTCER' ORDER BY DescripParametro"
    CargarControlLista strSql, cboEstado, arrEstado(), Sel_Todos
    
    intRegistro = ObtenerItemLista(arrEstado(), Estado_Certificado_Vigente)
    If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
    
    '*** Estado Bloqueo ***
    strSql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='ESTREG' ORDER BY DescripParametro"
    CargarControlLista strSql, cboEstadoOperacion, arrEstadoOperacion(), Valor_Caracter
            
    '*** Tipo Bloqueo ***
    strSql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPBLO' ORDER BY DescripParametro"
    CargarControlLista strSql, cboTipoBloqueo, arrTipoBloqueo(), Sel_Defecto
    
    '*** Tipo Documento Identidad ***
    strSql = "{ call up_ACSelDatos(11) }"
    CargarControlLista strSql, cboTipoDocumento, garrTipoDocumento(), Sel_Defecto
    
    If cboTipoDocumento.ListCount > 0 Then cboTipoDocumento.ListIndex = 0
                        
End Sub
Private Sub InicializarValores()

    strEstado = Reg_Defecto
    tabBloqueo.Tab = 0
    
    '*** Verificando Nivel de Acceso de Usuario ***
'    strNivAcceso = AccesoForm(gstrNomOpc, gstrNumInd)
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmCertificadoCliente = Nothing
    gstrCodParticipe = Valor_Caracter
    frmMainMdi.stbMdi.Panels(3).Text = "Acción..."
    
End Sub

Private Sub lblCantCuotas_Change()

    Call FormatoMillarEtiqueta(lblCantCuotas, Decimales_CantCuota)
    
End Sub

Private Sub tabBloqueo_Click(PreviousTab As Integer)

    Select Case tabBloqueo.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabBloqueo.Tab = 0
        
    End Select
    
End Sub

Private Sub tdgCertificado_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 2 Then
        Call DarFormatoValor(Value, Decimales_CantCuota)
    End If
    
End Sub

Private Sub tdgCertificado_SelChange(Cancel As Integer)

'    Dim dblCuotas   As Double, dblCuotasAcumulado   As Double
'    Dim intRegistro As Integer, intContador         As Integer
'
'    intContador = tdgCertificado.SelBookmarks.Count - 1
'
'    For intRegistro = 0 To intContador
'        tdgCertificado.Row = tdgCertificado.SelBookmarks(intRegistro) - 1
'        tdgCertificado.Refresh
'
'        dblCuotas = CDbl(tdgCertificado.Columns(2))
'        dblCuotasAcumulado = dblCuotasAcumulado + dblCuotas
'    Next
'
'    lblCuotas.Caption = CStr(dblCuotasAcumulado)
        
End Sub

'Private Sub tdgConsulta_DblClick()
'
'    On Error GoTo CtrlError         '/**/ HMC Habilitamos la rutina de Errores.
'
'    Dim intAccion   As Integer, lngNumError     As Long
'    Dim strSQL      As String
'
'    If tdgConsulta.Col = 3 Then
'        If tdgConsulta.Columns(3) = Valor_Indicador Then
'            tdgConsulta.Columns(3) = Valor_Caracter
'        Else
'            tdgConsulta.Columns(3) = Valor_Indicador
'            strSQL = "FechaCustodia='" & Convertyyyymmdd(gdatFechaActual) & "',"
'        End If
'        tdgConsulta.Update
'        Call ActualizaDatosCertificado(strSQL, tdgConsulta.Columns(0))
'        tdgConsulta.Refresh
'    End If
'
'    If tdgConsulta.Col = 4 Then
'        If tdgConsulta.Columns(4) = Valor_Indicador Then
'            tdgConsulta.Columns(4) = Valor_Caracter
'            strSQL = "FechaExtornoGarantia='" & Convertyyyymmdd(gdatFechaActual) & "',"
'        Else
'            tdgConsulta.Columns(4) = Valor_Indicador
'            strSQL = "FechaGarantia='" & Convertyyyymmdd(gdatFechaActual) & "',"
'        End If
'        tdgConsulta.Update
'        Call ActualizaDatosCertificado(strSQL, tdgConsulta.Columns(0))
'        tdgConsulta.Refresh
'    End If
'
'    If tdgConsulta.Col = 5 Then
'        If tdgConsulta.Columns(5) = Valor_Indicador Then
'            tdgConsulta.Columns(5) = Valor_Caracter
'        Else
'            tdgConsulta.Columns(5) = Valor_Indicador
'            strSQL = "FechaBloqueo='" & Convertyyyymmdd(gdatFechaActual) & "',"
'        End If
'        tdgConsulta.Update
'        Call ActualizaDatosCertificado(strSQL, tdgConsulta.Columns(0))
'        tdgConsulta.Refresh
'    End If
'
'    Call Buscar
'    Exit Sub
'
'CtrlError:
'    Me.MousePointer = vbDefault
'    intAccion = ControlErrores
'    Select Case intAccion
'        Case 0: Resume
'        Case 1: Resume Next
'        Case 2: Exit Sub
'        Case Else
'            lngNumError = Err.Number
'            Err.Raise Number:=lngNumError
'            Err.Clear
'    End Select
'
'End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 2 Then
        Call DarFormatoValor(Value, Decimales_CantCuota)
    End If
    
End Sub

Private Sub tdgConsulta_HeadClick(ByVal ColIndex As Integer)
    
    Dim strColNameTDB  As String
    Static numColindex As Integer
    Static strPrevColumTDB As String
    '** agregar para que no se raye la seleccion de registro con ordenamiento
    strColNameTDB = tdgConsulta.Columns(ColIndex).DataField
    
    If strColNameTDB = strPrevColumTDB Then
        If indSortAsc Then
            indSortAsc = False
            indSortDesc = True
        Else
            indSortAsc = True
            indSortDesc = False
        End If
    Else
        indSortAsc = True
        indSortDesc = False
    End If
    '***

    tdgConsulta.Splits(0).Columns(numColindex).HeadingStyle.ForegroundPicture = Null

    Call OrdenarDBGrid(ColIndex, adoConsulta, tdgConsulta)
    
    numColindex = ColIndex
    
    '****
    strPrevColumTDB = strColNameTDB
    '***
    
End Sub

Private Sub tdgConsulta_SelChange(Cancel As Integer)

    Dim dblCuotas   As Double, dblCuotasAcumulado   As Double
    Dim intRegistro As Integer, intContador         As Integer
        
    intContador = tdgConsulta.SelBookmarks.Count - 1

    For intRegistro = 0 To intContador
        tdgConsulta.Row = tdgConsulta.SelBookmarks(intRegistro) - 1
        tdgConsulta.Refresh
                
        dblCuotas = CDbl(tdgConsulta.Columns(2))
        dblCuotasAcumulado = dblCuotasAcumulado + dblCuotas
    Next
            
End Sub

Private Sub txtNumDocumento_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call ObtenerDatosParticipe
        Call Buscar
    End If
    
End Sub

Public Sub ObtenerCertificados()

'    Dim strSQL  As String
'
'    If strEstado = Reg_Adicion Then
'        strSQL = "SELECT NumCertificado,FechaSuscripcion,FechaOperacion,ValorCuota,CantCuotas FROM ParticipeCertificado " & _
'            "WHERE CodParticipe='" & gstrCodParticipe & "' AND CodFondo='" & strCodFondo & "' AND " & _
'            "CodAdministradora='" & gstrCodAdministradora & "' AND IndVigente='X' AND IndBloqueo='' " & _
'            "ORDER BY FechaSuscripcion"
'    Else
'        strSQL = "SELECT NumCertificado,FechaSuscripcion,FechaOperacion,ValorCuota,CantCuotas FROM ParticipeCertificado " & _
'            "WHERE CodParticipe='" & gstrCodParticipe & "' AND CodFondo='" & strCodFondo & "' AND " & _
'            "CodAdministradora='" & gstrCodAdministradora & "' AND IndVigente='X' AND IndBloqueo='X' AND " & _
'            "NumCertificado='" & tdgConsulta.Columns(1) & "' " & _
'            "ORDER BY FechaSuscripcion"
'    End If
'
'    With adoCertificado
'        .ConnectionString = gstrConnectConsulta
'        .RecordSource = strSQL
'        .Refresh
'    End With
'
'    tdgCertificado.Refresh
                
End Sub
Private Sub ObtenerDatosParticipe()

    Dim adoRegistro As ADODB.Recordset

    Set adoRegistro = New ADODB.Recordset
    adoRegistro.CursorLocation = adUseClient
    adoRegistro.CursorType = adOpenStatic

    adoComm.CommandText = "SELECT PC.CodParticipe,AP1.DescripParametro TipoIdentidad,PCD.NumIdentidad,DescripParticipe,FechaIngreso,PCD.TipoIdentidad CodIdentidad,PC.TipoMancomuno, AP2.DescripParametro DescripMancomuno " & _
    "FROM ParticipeContratoDetalle PCD JOIN ParticipeContrato PC " & _
    "ON(PCD.CodParticipe=PC.CodParticipe AND PCD.TipoIdentidad='" & strCodTipoDocumento & "' AND PCD.NumIdentidad='" & Trim(txtNumDocumento.Text) & "') " & _
    "JOIN AuxiliarParametro AP1 ON(AP1.CodParametro=PCD.TipoIdentidad AND CodTipoParametro='TIPIDE') " & _
    "JOIN AuxiliarParametro AP2 ON(AP2.CodParametro=PC.TipoMancomuno AND AP2.CodTipoParametro='TIPMAN')"
    adoRegistro.Open adoComm.CommandText, adoConn

    If Not adoRegistro.EOF Then
        If adoRegistro.RecordCount > 1 Then
            gstrFormulario = Me.Name
            frmBusquedaParticipeP.optCriterio(1).Value = vbChecked
            frmBusquedaParticipeP.txtNumDocumento = Trim(txtNumDocumento.Text)
            Call frmBusquedaParticipeP.Buscar
            frmBusquedaParticipeP.Show vbModal
        Else
            gstrCodParticipe = Trim(adoRegistro("CodParticipe"))
            lblDescripTipoParticipe.Caption = Trim(adoRegistro("DescripMancomuno"))
            lblDescripParticipe.Caption = Trim(adoRegistro("DescripParticipe"))
        End If
    End If
    adoRegistro.Close: Set adoRegistro = Nothing
    
End Sub

