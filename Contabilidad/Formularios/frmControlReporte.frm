VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmControlReporte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Reporte"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   11940
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   9360
      TabIndex        =   0
      Top             =   8040
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
      Left            =   1080
      TabIndex        =   1
      Top             =   8040
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Modificar"
      Tag0            =   "3"
      Visible0        =   0   'False
      ToolTipText0    =   "Modificar"
      UserControlWidth=   1200
   End
   Begin TabDlg.SSTab tabReporte 
      Height          =   7935
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   13996
      _Version        =   393216
      Style           =   1
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
      TabCaption(0)   =   "Reporte"
      TabPicture(0)   =   "frmControlReporte.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos"
      TabPicture(1)   =   "frmControlReporte.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(1)=   "fraDetalle"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Estructura"
      TabPicture(2)   =   "frmControlReporte.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmEstructura"
      Tab(2).ControlCount=   1
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -68640
         TabIndex        =   8
         Top             =   6360
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
         ToolTipText1    =   "Cancelar"
         UserControlWidth=   2700
      End
      Begin VB.Frame frmEstructura 
         Caption         =   "NOMBRE DEL REPORTE"
         Height          =   6915
         Left            =   -74580
         TabIndex        =   11
         Top             =   600
         Width           =   10965
         Begin VB.CheckBox chkDefault 
            Caption         =   "Establecer Por Defecto"
            Height          =   255
            Left            =   7560
            TabIndex        =   36
            Top             =   480
            Width           =   2295
         End
         Begin VB.CommandButton cmdAtras 
            Caption         =   "<="
            Height          =   375
            Left            =   360
            TabIndex        =   34
            ToolTipText     =   "Atras"
            Top             =   4680
            Width           =   375
         End
         Begin VB.ComboBox cboVistaProceso 
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   480
            Width           =   5120
         End
         Begin VB.ComboBox cboRubro 
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   960
            Width           =   5120
         End
         Begin VB.ComboBox cboSubRubro 
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1440
            Width           =   5120
         End
         Begin VB.CommandButton cmdActualizar 
            Caption         =   "A"
            Height          =   375
            Left            =   360
            TabIndex        =   22
            ToolTipText     =   "Actualizar"
            Top             =   5130
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
            TabIndex        =   21
            ToolTipText     =   "Quitar"
            Top             =   5970
            Width           =   375
         End
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
            TabIndex        =   20
            ToolTipText     =   "Agregar"
            Top             =   5550
            Width           =   375
         End
         Begin VB.Frame frVariable 
            Caption         =   "Definicion Variable"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   360
            TabIndex        =   12
            Top             =   1950
            Width           =   10065
            Begin VB.CheckBox chkFiltro 
               Caption         =   "Tiene Filtro"
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   360
               TabIndex        =   32
               Top             =   1500
               Width           =   1485
            End
            Begin VB.CommandButton cmdFiltro 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   8910
               TabIndex        =   30
               Top             =   1470
               Width           =   375
            End
            Begin VB.ComboBox cboVariable 
               Height          =   315
               Left            =   1890
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   540
               Width           =   3705
            End
            Begin VB.CommandButton cmdFormula 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   8910
               TabIndex        =   13
               Top             =   990
               Width           =   375
            End
            Begin VB.Label lblCodTipo 
               Height          =   255
               Left            =   9600
               TabIndex        =   33
               Top             =   600
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label lblFiltro 
               BackColor       =   &H8000000B&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1890
               TabIndex        =   31
               Top             =   1500
               Width           =   6930
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo"
               Height          =   195
               Index           =   6
               Left            =   6240
               TabIndex        =   19
               Top             =   630
               Width           =   315
            End
            Begin VB.Label lblTipo 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6840
               TabIndex        =   18
               Top             =   600
               Width           =   2505
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Index           =   7
               Left            =   360
               TabIndex        =   17
               Top             =   1050
               Width           =   360
            End
            Begin VB.Label lblValorVariable 
               BackColor       =   &H8000000B&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1890
               TabIndex        =   16
               Top             =   1020
               Width           =   6930
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nombre"
               Height          =   195
               Index           =   5
               Left            =   360
               TabIndex        =   15
               Top             =   570
               Width           =   555
            End
         End
         Begin TrueOleDBGrid60.TDBGrid tdgEstructura 
            Bindings        =   "frmControlReporte.frx":0054
            Height          =   2385
            Left            =   870
            OleObjectBlob   =   "frmControlReporte.frx":0071
            TabIndex        =   26
            Top             =   4260
            Width           =   9555
         End
         Begin VB.Label lblDefault 
            Height          =   255
            Left            =   7560
            TabIndex        =   35
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vista Usuario"
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   29
            Top             =   510
            Width           =   930
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rubro"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   28
            Top             =   990
            Width           =   435
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SubRubro"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   27
            Top             =   1470
            Width           =   720
         End
      End
      Begin VB.Frame fraDetalle 
         Caption         =   "Definición Reporte"
         Height          =   5055
         Left            =   -74520
         TabIndex        =   4
         Top             =   960
         Width           =   10935
         Begin VB.CheckBox chkIndPersonalizacion 
            Caption         =   "Indicador Personalizacion"
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
            Height          =   375
            Left            =   480
            TabIndex        =   9
            Top             =   1560
            Width           =   2895
         End
         Begin VB.Label txtDescripReporte 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2040
            TabIndex        =   10
            Top             =   1080
            Width           =   4305
         End
         Begin VB.Label lblCodReporte 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2040
            TabIndex        =   7
            Top             =   720
            Width           =   4335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo"
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   6
            Top             =   735
            Width           =   495
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion"
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   5
            Top             =   1080
            Width           =   840
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmControlReporte.frx":7137
         Height          =   6105
         Left            =   600
         OleObjectBlob   =   "frmControlReporte.frx":7151
         TabIndex        =   3
         Top             =   1200
         Width           =   10575
      End
   End
End
Attribute VB_Name = "frmControlReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrVistaProceso()   As String, arrRubro() As String, arrSubRubro() As String
Dim arrVariable() As String, arrSubRubroTmp() As String, arrVariableTmp() As String

Dim strEstado            As String, strIdVariable As String, strModo As String
Dim strCodReporte       As String

Dim adoRegistroAux          As ADODB.Recordset, adoConsulta            As ADODB.Recordset
Dim adoRegistroAlt          As ADODB.Recordset
Dim comprobar               As Boolean, indBorrado      As Boolean
Dim indSortAsc              As Boolean, indSortDesc     As Boolean

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

Private Sub InicializarValores()
    
    '*** Valores Iniciales ***
    comprobar = False
    indBorrado = False
    
    strEstado = Reg_Defecto
    tabReporte.Tab = 0
    tabReporte.TabEnabled(1) = False
    
    fraDetalle.Font = "Arial"
    fraDetalle.FontBold = True
    fraDetalle.ForeColor = &H800000
    
    cmdAtras.Visible = False
    cmdActualizar.Enabled = False
    
    lblDefault.Visible = False
    lblDefault.Font = "Arial"
    lblDefault.FontBold = True
    lblDefault.ForeColor = &H800000
    
    chkDefault.Visible = False
    chkDefault.Font = "Arial"
    chkDefault.FontBold = True
    chkDefault.ForeColor = &H800000
   
     With tabReporte
            .TabVisible(2) = False
     End With
     
     tdgConsulta.Columns(2).Alignment = dbgCenter
     cmdFiltro.Enabled = False
            
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub CargarListas()

    Dim strSQL  As String
    
        '*** Vistas de Proceso ***
    strSQL = "select CodVistaProceso AS CODIGO,DescripVistaProceso AS DESCRIP from VistaProceso " & _
             "where IndVigente='X' order by CodVistaProceso"
    CargarControlLista strSQL, cboVistaProceso, arrVistaProceso(), Sel_Defecto
    
    If cboVistaProceso.ListCount > 0 Then
        cboVistaProceso.ListIndex = 0
    End If
    
    strSQL = "select CodRubroEstructura AS CODIGO,DescripRubroEstructura AS DESCRIP from RubroEstructura " & _
                "order by DescripRubroEstructura"
    CargarControlLista strSQL, cboRubro, arrRubro(), Sel_Defecto
    
    If cboRubro.ListCount > 0 Then
        cboRubro.ListIndex = 0
    End If
    
End Sub

Private Sub CargarReportes()

    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Diario General"
    
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Diario General (ME)"
    
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Mayor General"
    
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Text = "Mayor General (ME)"
    
End Sub

Public Sub Buscar()

    Dim strSQL As String
    
    Set adoConsulta = New ADODB.Recordset
        
     
        
    strSQL = "SELECT CodReporte,DescripReporte,IndPersonalizado FROM ControlReporte " & _
                "ORDER BY CodReporte"
                        
    strEstado = Reg_Defecto
    
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    tdgConsulta.DataSource = adoConsulta
    
    If adoConsulta.RecordCount > 0 Then strEstado = Reg_Consulta
    
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

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
'        Case vNew
'            Call Adicionar
        Case vModify
            Call Modificar
'        Case vDelete
'            Call Eliminar
'        Case vSearch
'            Call Buscar
'        Case vReport
'            Call Imprimir
        Case vSave
            Call Grabar
        Case vCancel
            Call Cancelar
        Case vExit
            Call Salir
        
    End Select
    
End Sub

Public Sub Modificar()

    If strEstado = Reg_Consulta Then
    
        LlenarFormulario
        cmdOpcion.Visible = False
        
        With tabReporte
            .TabEnabled(0) = False
            .Tab = 1
            .TabEnabled(1) = True
        End With
        
    End If
    
End Sub

Public Sub Grabar()

    Dim objControlReporteEstructuraXML  As DOMDocument60
    Dim strMsgError                 As String
    Dim strControlReporteEstructuraXML As String
    Dim strPerzonalizado As String
    Dim adoRegistro As ADODB.Recordset
    Dim intResult As Long
    Dim borrar As Boolean
    
    
    If MsgBox(Mensaje_Edicion, vbQuestion + vbYesNo + vbDefaultButton2, gstrNombreEmpresa) = vbNo Then Exit Sub
    
    If strEstado = Reg_Adicion Or strEstado = Reg_Edicion Then
    
           If gstrCodVistaProceso <> Valor_Caracter Then
                Me.MousePointer = vbHourglass
                                                    
                With adoComm
                    
                    If adoRegistroAux.RecordCount = 0 Then
                    
                        borrar = True
                    Else
                        borrar = False
                    
                    End If
                    
                    
                    Call XMLADORecordset(objControlReporteEstructuraXML, "ControlReporteEstructura", "Estructura", adoRegistroAux, strMsgError)
                    strControlReporteEstructuraXML = objControlReporteEstructuraXML.xml
                                  
                                  
                    If chkIndPersonalizacion.Value = Checked Then
                        strPerzonalizado = "X"
                    Else
                        strPerzonalizado = ""
                    End If
                                  
                    '*** Cabecera ***
                    .CommandText = "{ call up_ACManControlReporteXML('" & _
                        strCodReporte & "','" & txtDescripReporte.Caption & "','','','','',''," & _
                        "'','','','','','','" & strPerzonalizado & "','','','','','','" & gstrCodVistaProceso & "','" & _
                        strControlReporteEstructuraXML & "','U') }"
                        
                    adoConn.Execute .CommandText
                    
                              
                End With
                
                Set adoRegistroAux = Nothing
                Set adoRegistroAlt = Nothing
                
                '*********PROCESO PARA AGREGAR DEFINICION DE VISTA A TABLA ControlReporteVista
                
                Set adoRegistro = New ADODB.Recordset
    
                With adoComm
                
                    If borrar = True Then
                    
                        .CommandText = "DELETE FROM ControlReporteVista " & _
                                            "WHERE CodVistaProceso='" & gstrCodVistaProceso & "'"
                                            
                        adoConn.Execute .CommandText
                        
                        .CommandText = "UPDATE ControlReporte SET IndPersonalizado='' WHERE CodReporte='" + strCodReporte + "'"

                        adoConn.Execute .CommandText
                        
                    Else
                    
                        .CommandText = "SELECT * FROM ControlReporteVista " & _
                                        "WHERE CodVistaProceso='" & gstrCodVistaProceso & "'"
        
                        Set adoRegistro = .Execute
        
                        If adoRegistro.EOF Then
                            
                            'POR DEFECTO TODAS LOS VISTAS PROCESO DEL REPORTE INGRESAN COMO POR DEFECTO
                            'LO CUAL PRODUCIRA UN ERROR SI SE INGRESA MAS DE UNA VISTA A UN REPORTE
                            'MEJORAR ESTE PUNTO......
                            .CommandText = "INSERT INTO ControlReporteVista " & _
                                            "VALUES('" & strCodReporte & "','" & gstrCodVistaProceso & _
                                            "','X')"
        
                             adoConn.Execute .CommandText
                                
                        End If
                    
                    End If
    
                End With
                
                '*****************************************************************************
                    
                Me.MousePointer = vbDefault
            
                MsgBox Mensaje_Adicion_Exitosa, vbExclamation
                
                frmMainMdi.stbMdi.Panels(3).Text = "Acción"
                
                cmdOpcion.Visible = True
                
                cboVistaProceso.Enabled = True
        
                tdgEstructura.Enabled = True
        
                With tabReporte
                .TabEnabled(0) = True
                .Tab = 0
                .TabVisible(2) = False
                .TabEnabled(1) = False
                End With
        
                LimpiarDatos
                Call Buscar
        
    
            Exit Sub
            
        Else
        
        MsgBox "Debe seleccionar una vista para poder guardar informacion", vbCritical
        
        End If
            
    End If
            
Ctrl_Error:
    '    adoComm.CommandText = "ROLLBACK TRAN ProcAsiento"
    '    adoConn.Execute adoComm.CommandText
        
        MsgBox Mensaje_Proceso_NoExitoso, vbCritical
        Me.MousePointer = vbDefault
        
        
End Sub

Public Sub Cancelar()

    adoRegistroAux.Close: Set adoRegistroAux = Nothing
    adoRegistroAlt.Close: Set adoRegistroAux = Nothing
    
    LimpiarDatos

    cmdOpcion.Visible = True
    
    cboVistaProceso.Enabled = True
    
    tdgEstructura.Enabled = True
    
    With tabReporte
        .TabEnabled(0) = True
        .Tab = 0
        .TabVisible(2) = False
        .TabEnabled(1) = False
    End With
    
    Call Buscar
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub LlenarFormulario()

    strCodReporte = tdgConsulta.Columns(0)
    
    Dim adoRegistro   As ADODB.Recordset
    Dim strSQL As String
        
    If strEstado = Reg_Consulta Then
        
        Dim intRegistro As Integer
        
        lblCodReporte.Caption = tdgConsulta.Columns(0)
        txtDescripReporte.Caption = tdgConsulta.Columns(1)
        
        If tdgConsulta.Columns(2) = "X" Then
            chkIndPersonalizacion.Value = Checked
        Else
            chkIndPersonalizacion.Value = Unchecked
        End If
        
        If chkIndPersonalizacion.Value = Checked Then
            tabReporte.TabVisible(2) = True
        Else
            tabReporte.TabVisible(2) = False
        End If
        
        frmEstructura.Caption = "Reporte: " + tdgConsulta.Columns(1)
        frmEstructura.ForeColor = &H800000
        frmEstructura.FontBold = True
        frmEstructura.Font = "Arial"
        
        chkFiltro.Font = "Arial"
        chkFiltro.ForeColor = &H800000
        chkFiltro.FontBold = True
        
        Set adoRegistro = New ADODB.Recordset
        
        With adoComm
                        .CommandText = "SELECT CodVistaProceso FROM ControlReporteVista " & _
                            "WHERE CodReporte='" & tdgConsulta.Columns(0) & "' AND PorDefecto='X'"
                        Set adoRegistro = .Execute
                        
                        If Not adoRegistro.EOF Then
                            intRegistro = ObtenerItemLista(arrVistaProceso(), adoRegistro("CodVistaProceso"))
                            If intRegistro >= 0 Then cboVistaProceso.ListIndex = intRegistro
                            
'                            lblDefault.Visible = True
'                            lblDefault.Caption = "VISTA DEFECTO"
                             lblDefault.Visible = False
                            
                        Else
                        
                            cboVistaProceso.ListIndex = 0
                            gstrCodVistaProceso = ""
                        
                        End If
                        
                        adoRegistro.Close: Set adoRegistro = Nothing
                        
        
        End With
        
        Call CargarDetalleGrilla
        
        Select Case strEstado

        Case Reg_Adicion


            strSQL = "select CodSubRubroEstructura AS CODIGO, DescripSubRubroEstructura AS DESCRIP from SubRubroEstructura " & _
                        "order by DescripSubRubroEstructura"
            CargarControlLista strSQL, cboSubRubro, arrSubRubro(), Sel_Defecto

            If cboSubRubro.ListCount > 0 Then
            
                cboSubRubro.ListIndex = 0
                
            End If

            strSQL = "select IdVariable AS CODIGO, DescripVariable AS DESCRIP from VariableUsuario " & _
                        " where TipoVariable!='02' and IndVigente='X' order by DescripVariable"
            CargarControlLista strSQL, cboVariable, arrVariable(), Sel_Defecto

            If cboVariable.ListCount > 0 Then
                cboVariable.ListIndex = 0
            End If
            
'        comprobar = True
        
        Case Reg_Edicion

            Dim strWhere As String

            Call CargaCamposToArray(arrSubRubroTmp, arrVariableTmp)

            strWhere = ObtenerCondicion(arrSubRubroTmp())


            strSQL = "select CodSubRubroEstructura AS CODIGO, DescripSubRubroEstructura AS DESCRIP from SubRubroEstructura " & _
                    "where CodSubRubroEstructura NOT IN (" + strWhere + ") order by DescripSubRubroEstructura "
            CargarControlLista strSQL, cboSubRubro, arrSubRubro(), Sel_Defecto

            If cboSubRubro.ListCount > 0 Then
                cboSubRubro.ListIndex = 0
            End If

            strWhere = ObtenerCondicion(arrVariableTmp())

            strSQL = "select IdVariable AS CODIGO, DescripVariable AS DESCRIP from VariableUsuario " & _
                        " where TipoVariable!='02' and IndVigente='X' and  IdVariable NOT IN (" + strWhere + ") order by DescripVariable"
            CargarControlLista strSQL, cboVariable, arrVariable(), Sel_Defecto

            If cboVariable.ListCount > 0 Then
                cboVariable.ListIndex = 0
            End If

'        comprobar = True
        
        End Select
        
    End If

End Sub

Private Sub CargarDetalleGrilla()
    
    Dim adoRegistro As ADODB.Recordset
    Dim adoField As ADODB.Field
    
    Dim strSQL As String
    
    '*********RECORDSET CLONADO
    Set adoRegistroAlt = New ADODB.Recordset
    
    
    Set adoRegistro = New ADODB.Recordset
        
    Call ConfiguraRecordsetAuxiliar
    
    
       
        strSQL = "SELECT CodReporte,CodVistaProceso,NumSecEstructura,CodVistaUsuario,CodRubroReporte, " + _
                "DescripRubroEstructura as DescripRubro," + _
                "CodSubRubroReporte,DescripSubRubroEstructura as DescripSubRubro," + _
                "CodVariableReporte,DescripParametro as TipoVariable," + _
                "TipoVariable as CodTipoVariable,CodVariableFormulaReporte," + _
                "ValorFiltroSubRubroReporte,NumOrdenEstructura FROM ControlReporteEstructura CRE " + _
                "JOIN RubroEstructura RE ON (CRE.CodRubroReporte=RE.CodRubroEstructura) " + _
                "JOIN SubRubroEstructura SRE ON (CRE.CodSubRubroReporte=SRE.CodSubRubroEstructura) " + _
                "JOIN VariableUsuario VU ON (CRE.CodVariableReporte=VU.IdVariable) " + _
                "JOIN AuxiliarParametro AP ON (AP.CodParametro=VU.TipoVariable) " + _
                "WHERE AP.CodTipoParametro='REPVAR' AND CRE.CodReporte='" + tdgConsulta.Columns(0) + _
                "' AND CodVistaProceso='" + gstrCodVistaProceso + "' " + _
                "ORDER BY NumSecEstructura"

        With adoComm
            .CommandText = strSQL
            Set adoRegistro = .Execute

                If Not adoRegistro.BOF Then
                
                    strEstado = Reg_Edicion
                    adoRegistro.MoveFirst
    
                    Do While Not adoRegistro.EOF
                        adoRegistroAux.AddNew
                        
                        For Each adoField In adoRegistroAux.Fields
                            adoRegistroAux.Fields(adoField.Name) = adoRegistro.Fields(adoField.Name)
                        Next
                        
                        adoRegistroAux.Update
                        adoRegistro.MoveNext
                    Loop

                     adoRegistro.Close: Set adoRegistro = Nothing

                    If Not adoRegistroAux.BOF Then adoRegistroAux.MoveFirst
                
                    '*******CAMBIO PARA CLONAR RECORDSET
                    Set adoRegistroAlt = adoRegistroAux.Clone
                    cmdQuitar.Enabled = True
                
                Else
                    '*******CAMBIO PARA CLONAR RECORDSET
                    Set adoRegistroAlt = adoRegistroAux.Clone
                    strEstado = Reg_Adicion
                    cmdQuitar.Enabled = False
                
                End If

        End With
    
    adoRegistroAux.UpdateBatch
    tdgEstructura.DataSource = adoRegistroAux
    tdgEstructura.Refresh
    
            
End Sub

'***FUNCION QUE AUN NO LA UTILIZO
Private Function ComprobarVigencia(ByVal strCodVistaUsuario As String) As Boolean

    Dim adoRegistro   As ADODB.Recordset
    
    With adoComm
                    .CommandText = "SELECT CodVistaUsuario FROM ControlReporteVista " & _
                        "WHERE CodReporte='" & tdgConsulta.Columns(0) & "' AND PorDefecto='X' AND CodVistaUsuario='" & _
                        strCodVistaUsuario & "'"
                    Set adoRegistro = .Execute
                    
                    If Not adoRegistro.EOF Then
                        ComprobarVigencia = True
                    Else
                        ComprobarVigencia = False
                    End If
                                  
                    adoRegistro.Close: Set adoRegistro = Nothing
                    
    End With


End Function

Private Sub ConfiguraRecordsetAuxiliar()

    Set adoRegistroAux = New ADODB.Recordset

    With adoRegistroAux
    
       .CursorLocation = adUseClient
       .Fields.Append "CodReporte", adVarChar, 35
       .Fields.Append "CodVistaProceso", adChar, 3
       .Fields.Append "NumSecEstructura", adInteger
       .Fields.Append "CodVistaUsuario", adChar, 3
       .Fields.Append "CodRubroReporte", adChar, 3
       .Fields.Append "DescripRubro", adVarChar, 200
       .Fields.Append "CodSubRubroReporte", adChar, 3
       .Fields.Append "DescripSubRubro", adVarChar, 200
       .Fields.Append "CodVariableReporte", adVarChar, 40
       .Fields.Append "CodTipoVariable", adChar, 2
       .Fields.Append "TipoVariable", adVarChar, 60
       .Fields.Append "CodVariableFormulaReporte", adVarChar, 10000
       .Fields.Append "ValorFiltroSubRubroReporte", adVarChar, 10000
       .Fields.Append "NumOrdenEstructura", adInteger
       .LockType = adLockBatchOptimistic
       
    End With
    
    adoRegistroAux.Open
    
End Sub

Private Function TodoOkFormula() As Boolean

    TodoOkFormula = False
    
    If cboVistaProceso.ListIndex <= 0 Then
    
        MsgBox "Debe seleccionar una Vista para definir una variable", vbInformation
    
    Exit Function
    
    End If
    
    If cboRubro.ListIndex <= 0 Then
    
        MsgBox "Debe seleccionar un Rubro para definir una variable", vbInformation
    
    Exit Function
    
    End If
    
    If cboSubRubro.ListIndex <= 0 Then
    
        MsgBox "Debe seleccionar un SubRubro para definir una variable", vbInformation
    
    Exit Function
    
    End If
    
    If cboVariable.ListIndex <= 0 Then
    
        MsgBox "Debe seleccionar una Variable para definir su valor", vbInformation
    
    Exit Function
    
    End If
    
    TodoOkFormula = True

End Function

Private Function TodoOkEstructura() As Boolean

    TodoOkEstructura = False
    
    If cboVistaProceso.ListIndex <= 0 Then
    
        MsgBox "Seleccione una Vista para continuar", vbInformation
    
    Exit Function
    
    End If
    
    If cboRubro.ListIndex <= 0 Then
    
        MsgBox "Seleccione un Rubro para continuar", vbInformation
    
    Exit Function
    
    End If
    
    If cboSubRubro.ListIndex <= 0 Then
    
        MsgBox "Seleccione un SubRubro para continuar", vbInformation
    
    Exit Function
    
    End If
    
    If cboVariable.ListIndex <= 0 Then
    
        MsgBox "Seleccione una Variable para continuar", vbInformation
    
    Exit Function
    
    End If
    
    If Len(Trim(lblValorVariable.Caption)) = 0 Then
    
        MsgBox "Asigne valor a la Variable para continuar", vbInformation
    
    Exit Function
    
    End If
    
    TodoOkEstructura = True

End Function

Private Sub LimpiarDatos()

    cboRubro.ListIndex = 0
    cboSubRubro.ListIndex = 0
    cboVariable.ListIndex = 0
    lblValorVariable.Caption = ""
    lblFiltro.Caption = ""
    lblTipo.Caption = ""
    chkFiltro.Value = Unchecked
    cmdFormula.Enabled = True
    
End Sub

Private Sub CargaCamposToArray(ByRef arrSubRubroTemp() As String, ByRef arrVariableTemp() As String, Optional ByVal strIndiceEditar As String)

    Dim intCont As Integer
    Dim indice As Integer
    Dim encontrado As Boolean
    Dim dblBookmark As Double
    Dim arrTemp As String
    
    encontrado = False

    intCont = 0

    ReDim arrSubRubroTemp(intCont)
    ReDim arrVariableTemp(intCont)


        If Not adoRegistroAlt.BOF Then
        
            adoRegistroAlt.MoveFirst

            If strIndiceEditar = "" Then

                Do Until adoRegistroAlt.EOF

                    ReDim Preserve arrSubRubroTemp(intCont)
                    ReDim Preserve arrVariableTemp(intCont)
                    
                    arrSubRubroTemp(intCont) = adoRegistroAlt("CodSubRubroReporte")
                    arrVariableTemp(intCont) = adoRegistroAlt("CodVariableReporte")
            
                    intCont = intCont + 1

                    adoRegistroAlt.MoveNext

                Loop

            Else

                dblBookmark = CDbl(strIndiceEditar)
                indice = CInt(strIndiceEditar) - 1

                'COMPRUEBO SI EL ELEMENTO QUE BUSCO ES EL PRIMERO DE LA LISTA
                If intCont = indice Then encontrado = True

                    Do Until adoRegistroAlt.EOF
    
                        If encontrado = False Then
    
                            ReDim Preserve arrSubRubroTemp(intCont)
                            ReDim Preserve arrVariableTemp(intCont)

    
                            arrSubRubroTemp(intCont) = adoRegistroAlt("CodSubRubroReporte")
                            arrVariableTemp(intCont) = adoRegistroAlt("CodVariableReporte")
    
                            intCont = intCont + 1
    
                            'CUANDO EL ELEMENTO QUE BUSCO NO ES EL PRIMERO DE LA LISTA
                            If intCont = indice Then encontrado = True
    
                            'adoRegistroAux.MoveNext
                            adoRegistroAlt.MoveNext
    
                        Else
    
                            'adoRegistroAux.MoveNext
                            adoRegistroAlt.MoveNext
    
                            encontrado = False
    
                        End If
    
                     Loop

            End If

            If strIndiceEditar = "" Then

                If adoRegistroAlt.EOF And adoRegistroAlt.RecordCount > 0 Then

                    adoRegistroAlt.MoveFirst
                    
                End If
            Else

                If adoRegistroAlt.EOF Then

                    adoRegistroAlt.MoveFirst
                    
                    adoRegistroAlt.Bookmark = dblBookmark

                End If
                
            End If

        Else

            arrSubRubroTemp(0) = ""
            arrVariableTemp(0) = ""

        End If
    
End Sub

Private Function ObtenerCondicion(ByRef arrTmp() As String) As String

    Dim i As Integer

    For i = 0 To UBound(arrTmp)

        arrTmp(i) = "'" + arrTmp(i) + "'"

    Next

    ObtenerCondicion = Join(arrTmp, ",")

End Function


Public Sub XMLADORecordset(ByRef objXML As DOMDocument60, ByVal strNomDocumento As String, ByVal strNomEntidad As String, ByVal adoRecordset As ADODB.Recordset, ByRef strMsgError As String, Optional strNomCampos As String = "", Optional strCampoCond As String, Optional strDatoCond As String, Optional optFilaInicio As Integer = 1)

    Dim objElem As MSXML2.IXMLDOMElement
    Dim objParent As MSXML2.IXMLDOMElement
    Dim lngPos As Long, lngParent As Long
    Dim i As Integer, j As Integer, aux As Integer
    Dim lblnSuccess As Boolean
    Dim NomCampos() As String, ArrayCols() As String
    Dim indCumpleCondicion As Boolean
    Dim adoField As ADODB.Field
    Dim n As Integer
    
    On Error GoTo ErrCreaXMLADORecordset
    
    NomCampos = Split(strNomCampos, ",")
    
    If objXML Is Nothing Then
    
        Set objXML = New MSXML2.DOMDocument60
        Set objXML.documentElement = objXML.createElement(strNomDocumento)
        
    End If
    
    Set objParent = objXML.documentElement
    
    n = 0
    
    'Recorriendo todas las filas de una rejilla
    If adoRecordset.RecordCount > 0 Then
                
        If UBound(NomCampos) = -1 Then
        
            For Each adoField In adoRecordset.Fields

                ReDim Preserve NomCampos(n)
                NomCampos(n) = adoField.Name
                n = n + 1
                
            Next
            
        End If
                
        adoRecordset.MoveFirst
        
        Do While Not adoRecordset.EOF
        
            indCumpleCondicion = True
            
            If strCampoCond <> "" Then
            
                If adoRecordset.Fields(strCampoCond).Value <> strDatoCond Then indCumpleCondicion = False
                
            End If
            
            If indCumpleCondicion Then
            
                Set objElem = objParent.appendChild(objXML.createElement(strNomEntidad))
                
                For j = 0 To UBound(NomCampos)
                
                    'añadiendo los atributos, solo para las columnas especificadas
                    If NomCampos(j) = "ValorFiltroSubRubroReporte" Then
                    
                        adoRecordset.Fields(NomCampos(j)).Value = Replace(adoRecordset.Fields(NomCampos(j)).Value, "'", "''")
                    
                    End If
                    
                    objElem.setAttribute NomCampos(j), "" & adoRecordset.Fields(NomCampos(j)).Value
                    
                Next
                
            End If
            
            adoRecordset.MoveNext
            
        Loop
        
    End If

    Set objElem = Nothing
    Set objParent = Nothing
    Exit Sub
    
ErrCreaXMLADORecordset:
    If strMsgError = "" Then strMsgError = "(XMLADORecordset) - " & err.Description

End Sub

Private Sub cboVariable_Click()
    
    Dim strSQL          As String
    Dim adoRegistro   As ADODB.Recordset
    strIdVariable = Valor_Caracter
 
    If cboVariable.ListIndex < 0 Then Exit Sub
    
    strIdVariable = Trim(arrVariable(cboVariable.ListIndex))
    cboVariable.ToolTipText = cboVariable.Text

    strSQL = "select CodParametro AS CODIGO,DescripParametro AS DESCRIP from VariableUsuario VUV, AuxiliarParametro AP " & _
    "where VUV.TipoVariable=AP.CodParametro AND AP.CodTipoParametro='REPVAR' " & _
    "AND VUV.IdVariable='" + strIdVariable + "'"
    
    Set adoRegistro = New ADODB.Recordset

    With adoComm
    
                .CommandText = strSQL
                Set adoRegistro = .Execute
                
                If Not adoRegistro.EOF Then
                
                   lblCodTipo.Caption = adoRegistro("CODIGO")
                   lblTipo.Caption = adoRegistro("DESCRIP")
                   
                Else
                
                   lblCodTipo.Caption = 0
                   lblTipo.Caption = ""
                   
                End If
                
                adoRegistro.Close: Set adoRegistro = Nothing
                
    End With
    
    lblValorVariable.Caption = ""
    lblFiltro.Caption = ""
    chkFiltro.Value = Unchecked
    
    If lblCodTipo.Caption = "99" Or lblCodTipo.Caption = "98" Then
    
        cmdFormula.Enabled = True
        chkFiltro.Enabled = False
        
    ElseIf lblCodTipo.Caption = "03" Then
    
        cmdFormula.Enabled = False
        lblValorVariable.Caption = 0
        chkFiltro.Enabled = False
    
    ElseIf lblCodTipo.Caption = "0" Then
        
        cmdFormula.Enabled = False
        chkFiltro.Enabled = False
        chkFiltro.Value = Unchecked
    
    Else
    
        cmdFormula.Enabled = True
        chkFiltro.Enabled = True
    
    End If

End Sub


Private Sub chkFiltro_Click()

    If chkFiltro.Value = Checked Then
    
        cmdFiltro.Enabled = True
        
    Else
    
        cmdFiltro.Enabled = False
        
        lblFiltro.Caption = ""
    
    End If

End Sub


Private Sub chkIndPersonalizacion_Click()

    If chkIndPersonalizacion.Value = Checked Then
        tabReporte.TabVisible(2) = True
    Else
        tabReporte.TabVisible(2) = False
    End If

End Sub



Private Sub cmdAgregar_Click()

    Dim adoRegistro As ADODB.Recordset
    Dim intSecuencial As Integer
    Dim strSQL As String
    
    intSecuencial = CInt(adoRegistroAux.RecordCount) + 1
    
    If TodoOkEstructura = True Then
    
            adoRegistroAux.AddNew
            adoRegistroAux.Fields("CodReporte") = strCodReporte
            adoRegistroAux.Fields("CodVistaProceso") = gstrCodVistaProceso
            adoRegistroAux.Fields("NumSecEstructura") = intSecuencial
            adoRegistroAux.Fields("CodVistaUsuario") = gstrCodVistaUsuario
            adoRegistroAux.Fields("CodRubroReporte") = arrRubro(cboRubro.ListIndex)
            adoRegistroAux.Fields("DescripRubro") = cboRubro.Text
            adoRegistroAux.Fields("CodSubRubroReporte") = arrSubRubro(cboSubRubro.ListIndex)
            adoRegistroAux.Fields("DescripSubRubro") = cboSubRubro.Text
            adoRegistroAux.Fields("CodVariableReporte") = arrVariable(cboVariable.ListIndex)
            adoRegistroAux.Fields("CodTipoVariable") = lblCodTipo.Caption
            adoRegistroAux.Fields("TipoVariable") = lblTipo.Caption
            adoRegistroAux.Fields("CodVariableFormulaReporte") = Trim(lblValorVariable.Caption) + Space(1)
            
            If chkFiltro.Value = Checked Then
            
                Dim strValorFiltro As String
                
                strValorFiltro = Trim(lblFiltro.Caption) + Space(1)
                
                adoRegistroAux.Fields("ValorFiltroSubRubroReporte") = strValorFiltro
            
            End If
            
            adoRegistroAux.Fields("NumOrdenEstructura") = intSecuencial
            
            tdgEstructura.DataSource = adoRegistroAux
            tdgEstructura.Refresh
             
            If adoRegistroAux.RecordCount > 0 Then
            
                cmdQuitar.Enabled = True
                
            End If
             
            Set adoRegistroAlt = adoRegistroAux.Clone

            Call CargaCamposToArray(arrSubRubroTmp, arrVariableTmp)
          
            
            Dim strWhere As String
        
            strWhere = ObtenerCondicion(arrSubRubroTmp())
        
        
            strSQL = "select CodSubRubroEstructura AS CODIGO, DescripSubRubroEstructura AS DESCRIP from SubRubroEstructura " & _
                "where CodSubRubroEstructura NOT IN (" + strWhere + ") order by DescripSubRubroEstructura "
            CargarControlLista strSQL, cboSubRubro, arrSubRubro(), Sel_Defecto
        
            If cboSubRubro.ListCount > 0 Then
            
                cboSubRubro.ListIndex = 0
                
            End If
        
            strWhere = ObtenerCondicion(arrVariableTmp())
        
            strSQL = "select IdVariable AS CODIGO, DescripVariable AS DESCRIP from VariableUsuario " & _
                    " where TipoVariable!='02' and  IdVariable NOT IN (" + strWhere + ") order by DescripVariable"
            CargarControlLista strSQL, cboVariable, arrVariable(), Sel_Defecto
        
            If cboVariable.ListCount > 0 Then
            
                cboVariable.ListIndex = 0
                
            End If
    
            LimpiarDatos
            
    End If

End Sub

Private Sub cmdActualizar_Click()
    
    If TodoOkEstructura = True Then
        
        adoRegistroAux.Fields("CodVistaUsuario") = gstrCodVistaUsuario
        adoRegistroAux.Fields("CodRubroReporte") = arrRubro(cboRubro.ListIndex)
        adoRegistroAux.Fields("DescripRubro") = cboRubro.Text
        adoRegistroAux.Fields("CodSubRubroReporte") = arrSubRubro(cboSubRubro.ListIndex)
        adoRegistroAux.Fields("DescripSubRubro") = cboSubRubro.Text
        adoRegistroAux.Fields("CodVariableReporte") = arrVariable(cboVariable.ListIndex)
        adoRegistroAux.Fields("CodTipoVariable") = lblCodTipo.Caption
        adoRegistroAux.Fields("TipoVariable") = lblTipo.Caption
        adoRegistroAux.Fields("CodVariableFormulaReporte") = Trim(lblValorVariable.Caption) + Space(1)
        
        
        If chkFiltro.Value = Checked Then
                
                Dim strValorFiltro As String
                
                strValorFiltro = Trim(lblFiltro.Caption) + Space(1)
                
                adoRegistroAux.Fields("ValorFiltroSubRubroReporte") = strValorFiltro
                
        End If
                
        Set adoRegistroAlt = adoRegistroAux.Clone
        
        Call cmdAtras_Click
        
    End If
            
                     
End Sub

Private Sub cmdQuitar_Click()

    Dim dblBook As String
    Dim Fila As String
    Dim strSQL As String
    Dim intIndice As Integer
    Dim dblBookPrimero As String
    Dim adoPrimero As ADODB.Recordset
    
    Set adoPrimero = New ADODB.Recordset
    
    If adoRegistroAux.RecordCount > 0 Then
              
        dblBook = adoRegistroAux.AbsolutePosition

    
        adoRegistroAux.Delete
'        indBorrado = True
            
        Set adoRegistroAlt = adoRegistroAux.Clone
        
        Call CargaCamposToArray(arrSubRubroTmp, arrVariableTmp)
          
            
            Dim strWhere As String
        
            strWhere = ObtenerCondicion(arrSubRubroTmp())
        
        
            strSQL = "select CodSubRubroEstructura AS CODIGO, DescripSubRubroEstructura AS DESCRIP from SubRubroEstructura " & _
                "where CodSubRubroEstructura NOT IN (" + strWhere + ") order by DescripSubRubroEstructura "
            CargarControlLista strSQL, cboSubRubro, arrSubRubro(), Sel_Defecto
        
            If cboSubRubro.ListCount > 0 Then
            
                cboSubRubro.ListIndex = 0
                
            End If
        
            strWhere = ObtenerCondicion(arrVariableTmp())
        
            strSQL = "select IdVariable AS CODIGO, DescripVariable AS DESCRIP from VariableUsuario " & _
                    " where TipoVariable!='02' and  IdVariable NOT IN (" + strWhere + ") order by DescripVariable"
            CargarControlLista strSQL, cboVariable, arrVariable(), Sel_Defecto
        
            If cboVariable.ListCount > 0 Then
                cboVariable.ListIndex = 0
            End If
            
            
        '**********SUBIR O BAJAR CURSOR SEGUN ELIMINE********
        If adoRegistroAux.RecordCount >= 1 And dblBook = 1 Then
        
            adoRegistroAux.MoveFirst
            tdgEstructura.MoveFirst

        ElseIf adoRegistroAux.RecordCount = 0 Then

            cmdQuitar.Enabled = False

        Else

            intIndice = CInt(dblBook) - 1
            dblBook = intIndice
            adoRegistroAux.AbsolutePosition = CDbl(dblBook)

        End If
    
        LimpiarDatos
    
    End If
    
    
End Sub

Private Sub cmdFiltro_Click()

    strModo = "FILTRO"
    
    If TodoOkFormula = False Then Exit Sub
    
    Dim strRespuesta As String
    
    strRespuesta = lblFiltro.Caption
    frmFormulas.mostrarForm strRespuesta, "01", adoRegistroAlt, strModo
    lblFiltro.Caption = strRespuesta
    
    If Trim(strRespuesta) = Valor_Caracter Then
        
        lblFiltro.Caption = ""
        chkFiltro.Value = Unchecked
    
    End If


End Sub

Private Sub cmdFormula_Click()

    strModo = "FORMULA"
    If TodoOkFormula = False Then Exit Sub
    
    Dim strRespuesta As String
    
    strRespuesta = lblValorVariable.Caption
    frmFormulas.mostrarForm strRespuesta, "01", adoRegistroAlt, strModo, lblCodTipo.Caption
    lblValorVariable.Caption = strRespuesta
    
    If Trim(strRespuesta) = Valor_Caracter Then
    
        gstrCodVistaUsuario = Valor_Caracter
        
    End If

End Sub


Private Sub cboVistaProceso_Click()
    
    Dim adoRegistro As ADODB.Recordset
    
    gstrCodVistaProceso = arrVistaProceso(cboVistaProceso.ListIndex)
    
    
    'LLENAR VARIABLE VISTA USUARIO
    If gstrCodVistaProceso <> Valor_Caracter Then
    
        With adoComm
            .CommandText = "SELECT CodVistaUsuario FROM VistaProcesoDetalle " & _
                                    "WHERE CodVistaProceso='" & gstrCodVistaProceso & "'"
            
            Set adoRegistro = .Execute
            
        
        End With
        
        Do Until adoRegistro.EOF
            gstrCodVistaUsuario = adoRegistro("CodVistaUsuario")
            adoRegistro.MoveNext
        Loop
    
    
    End If
    
'    If comprobar = False Then Exit Sub
'
'    If adoRegistroAux.EditMode = adEditAdd Or adoRegistroAux.EditMode = adEditInProgress _
'        Or indBorrado = True Then
'
'    If MsgBox("Si cambia de Vista perdera todo cambio que no aya salvado" + vbCrLf + "Desea Continuar?", vbYesNo _
'    + vbQuestion) = vbYes Then
'
'    End If
'
'    End If
    
        
End Sub

Private Sub tdgConsulta_DblClick()
    Call Modificar
End Sub

Private Sub tdgEstructura_DblClick()

    Dim dblBoomark As Integer
    Dim strSQL As String

    If adoRegistroAux.RecordCount > 0 Then
    
        dblBoomark = adoRegistroAux.AbsolutePosition
     
        Dim intRegistro As Integer
        
        intRegistro = ObtenerItemLista(arrRubro(), tdgEstructura.Columns(4))
        
        If intRegistro >= 0 Then cboRubro.ListIndex = intRegistro
        
        
        Call CargaCamposToArray(arrSubRubroTmp(), arrVariableTmp(), dblBoomark)
        
        Dim strWhere As String
        
        strWhere = ObtenerCondicion(arrSubRubroTmp())
        
        strSQL = "select CodSubRubroEstructura AS CODIGO, DescripSubRubroEstructura AS DESCRIP from SubRubroEstructura " & _
                "where CodSubRubroEstructura NOT IN (" + strWhere + ") order by DescripSubRubroEstructura "
                
        CargarControlLista strSQL, cboSubRubro, arrSubRubro(), Sel_Defecto
        
        intRegistro = ObtenerItemLista(arrSubRubro(), tdgEstructura.Columns(6))

        If intRegistro >= 0 Then cboSubRubro.ListIndex = intRegistro
        
        strWhere = ObtenerCondicion(arrVariableTmp())
        
        strSQL = "select IdVariable AS CODIGO, DescripVariable AS DESCRIP from VariableUsuario " & _
                    " where TipoVariable!='02' and IndVigente='X' and IdVariable NOT IN (" + strWhere + _
                    ") order by DescripVariable"
                    
        CargarControlLista strSQL, cboVariable, arrVariable(), Sel_Defecto
        
        intRegistro = ObtenerItemLista(arrVariable(), tdgEstructura.Columns(8))

        If intRegistro >= 0 Then cboVariable.ListIndex = intRegistro
               
        lblValorVariable.Caption = tdgEstructura.Columns(11)
        
    
        If Trim(CStr(tdgEstructura.Columns(12).Value)) = "" Then
        
            chkFiltro.Value = Unchecked
            
            lblFiltro.Caption = tdgEstructura.Columns(12)
        
        Else
        
            chkFiltro.Value = Checked
            
            Dim strFiltro As String
            
        '    strFiltro = LTrim$(Replace(CStr(tdgEstructura.Columns(11).Value), "AND", "", 1, 1))
            
            lblFiltro.Caption = tdgEstructura.Columns(12)
        
        
        End If
    
        cboVistaProceso.Enabled = False
        
        cmdAtras.Visible = True
        
        cmdActualizar.Enabled = True
        
        cmdAgregar.Enabled = False
        
        cmdQuitar.Enabled = False
        
        tdgEstructura.Enabled = False
    
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

Private Sub cmdAtras_Click()

    Dim strSQL As String

    Call CargaCamposToArray(arrSubRubroTmp, arrVariableTmp)
              
                
    Dim strWhere As String
            
    strWhere = ObtenerCondicion(arrSubRubroTmp())
            
            
    strSQL = "select CodSubRubroEstructura AS CODIGO, DescripSubRubroEstructura AS DESCRIP from SubRubroEstructura " & _
                    "where CodSubRubroEstructura NOT IN (" + strWhere + ") order by DescripSubRubroEstructura "
                    
    CargarControlLista strSQL, cboSubRubro, arrSubRubro(), Sel_Defecto
            
    If cboSubRubro.ListCount > 0 Then
        cboSubRubro.ListIndex = 0
    End If
            
    strWhere = ObtenerCondicion(arrVariableTmp())
            
    strSQL = "select IdVariable AS CODIGO, DescripVariable AS DESCRIP from VariableUsuario " & _
                        " where TipoVariable!='02' and IndVigente='X' and IdVariable NOT IN (" + strWhere + _
                        ") order by DescripVariable"
    
    CargarControlLista strSQL, cboVariable, arrVariable(), Sel_Defecto
            
    If cboVariable.ListCount > 0 Then
        cboVariable.ListIndex = 0
    End If
    
    cboVistaProceso.Enabled = True
    
    cmdAtras.Visible = False
    
    cmdActualizar.Enabled = False
    
    cmdAgregar.Enabled = True
    
    cmdQuitar.Enabled = True
    
    tdgEstructura.Enabled = True
    
    chkFiltro.Enabled = True
    
    LimpiarDatos

End Sub

'Private Sub RecargarGrillaVista()
'
'        Call CargarDetalleGrilla
'
'        Select Case strEstado
'
'        Case Reg_Adicion
'
'
'            strSql = "select CodSubRubroEstructura AS CODIGO, DescripSubRubroEstructura AS DESCRIP from SubRubroEstructura " & _
'                        "order by DescripSubRubroEstructura"
'            CargarControlLista strSql, cboSubRubro, arrSubRubro(), Sel_Defecto
'
'            If cboSubRubro.ListCount > 0 Then
'                cboSubRubro.ListIndex = 0
'            End If
'
'            strSql = "select IdVariable AS CODIGO, DescripVariable AS DESCRIP from VariableUsuario " & _
'                        " where TipoVariable!='02' and IndVigente='X' order by DescripVariable"
'            CargarControlLista strSql, cboVariable, arrVariable(), Sel_Defecto
'
'            If cboVariable.ListCount > 0 Then
'                cboVariable.ListIndex = 0
'            End If
'
'        Case Reg_Edicion
'
'            Dim strWhere As String
'
'            Call CargaCamposToArray(arrSubRubroTmp, arrVariableTmp)
'
'            strWhere = ObtenerCondicion(arrSubRubroTmp())
'
'
'            strSql = "select CodSubRubroEstructura AS CODIGO, DescripSubRubroEstructura AS DESCRIP from SubRubroEstructura " & _
'                    "where CodSubRubroEstructura NOT IN (" + strWhere + ") order by DescripSubRubroEstructura "
'            CargarControlLista strSql, cboSubRubro, arrSubRubro(), Sel_Defecto
'
'            If cboSubRubro.ListCount > 0 Then
'                cboSubRubro.ListIndex = 0
'            End If
'
'            strWhere = ObtenerCondicion(arrVariableTmp())
'
'            strSql = "select IdVariable AS CODIGO, DescripVariable AS DESCRIP from VariableUsuario " & _
'                        " where TipoVariable!='02' and IndVigente='X' and  IdVariable NOT IN (" + strWhere + ") order by DescripVariable"
'            CargarControlLista strSql, cboVariable, arrVariable(), Sel_Defecto
'
'            If cboVariable.ListCount > 0 Then
'                cboVariable.ListIndex = 0
'            End If
'
'
'        End Select
'
'End Sub







