VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmVistaUsuario 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vista Usuario"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   12345
   Begin VB.TextBox Invisible 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   345
      Left            =   13500
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   2550
      Width           =   495
   End
   Begin TabDlg.SSTab tabVista 
      Height          =   8715
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   15372
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
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
      TabCaption(0)   =   "Vista"
      TabPicture(0)   =   "frmVistaUsuario.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdOpcion"
      Tab(0).Control(1)=   "cmdSalir"
      Tab(0).Control(2)=   "tdgConsulta"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Datos"
      TabPicture(1)   =   "frmVistaUsuario.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frVista"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame frVista 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7605
         Left            =   150
         TabIndex        =   1
         Top             =   420
         Width           =   11655
         Begin VB.TextBox txtDescripcion 
            Height          =   315
            Left            =   2520
            TabIndex        =   10
            Top             =   780
            Width           =   5010
         End
         Begin VB.TextBox txtIdVista 
            Height          =   315
            Left            =   2520
            TabIndex        =   9
            Top             =   1260
            Width           =   5010
         End
         Begin VB.ComboBox cboTipoVista 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1710
            Width           =   2205
         End
         Begin VB.Frame frQuery 
            Caption         =   "Query"
            Height          =   2205
            Left            =   120
            TabIndex        =   6
            Top             =   2190
            Width           =   11355
            Begin RichTextLib.RichTextBox txtValorFormulaVista 
               Height          =   1665
               Left            =   150
               TabIndex        =   7
               Top             =   300
               Width           =   11025
               _ExtentX        =   19447
               _ExtentY        =   2937
               _Version        =   393217
               BackColor       =   14737632
               Enabled         =   -1  'True
               ScrollBars      =   3
               TextRTF         =   $"frmVistaUsuario.frx":0038
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.CommandButton cmdEditarQuery 
            Caption         =   "Editar Query"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   9630
            TabIndex        =   5
            Top             =   1680
            Width           =   1515
         End
         Begin TabDlg.SSTab tabCampos 
            Height          =   2955
            Left            =   120
            TabIndex        =   2
            Top             =   4500
            Width           =   11325
            _ExtentX        =   19976
            _ExtentY        =   5212
            _Version        =   393216
            Tabs            =   2
            TabHeight       =   520
            TabCaption(0)   =   "Asignar Tipo de Datos"
            TabPicture(0)   =   "frmVistaUsuario.frx":00C5
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "tdgCamposVista"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "chkFiltro"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Filtro"
            TabPicture(1)   =   "frmVistaUsuario.frx":00E1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "tdgCamposFiltro"
            Tab(1).Control(1)=   "cmdAtras"
            Tab(1).Control(2)=   "cmdActualizar"
            Tab(1).Control(3)=   "cmdQuitar"
            Tab(1).Control(4)=   "cmdAgregar"
            Tab(1).Control(5)=   "frCamposFiltro"
            Tab(1).ControlCount=   6
            Begin VB.Frame frCamposFiltro 
               Height          =   2355
               Left            =   -74820
               TabIndex        =   24
               Top             =   450
               Width           =   3855
               Begin VB.TextBox txtCampoFiltro 
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   26
                  Top             =   240
                  Width           =   2580
               End
               Begin VB.ComboBox cboIdVariableFiltro 
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   90
                  Style           =   2  'Dropdown List
                  TabIndex        =   25
                  Top             =   1050
                  Width           =   3585
               End
               Begin VB.Label lblIdVariableFiltro 
                  Height          =   285
                  Left            =   90
                  TabIndex        =   31
                  Top             =   1920
                  Width           =   3585
               End
               Begin VB.Label lblDescrip 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Id Variable"
                  Height          =   195
                  Index           =   6
                  Left            =   90
                  TabIndex        =   30
                  Top             =   1530
                  Width           =   750
               End
               Begin VB.Label lblDescrip 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Campo"
                  Height          =   195
                  Index           =   4
                  Left            =   120
                  TabIndex        =   28
                  Top             =   270
                  Width           =   495
               End
               Begin VB.Label lblDescrip 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Descripcion Variable"
                  Height          =   195
                  Index           =   5
                  Left            =   120
                  TabIndex        =   27
                  Top             =   720
                  Width           =   1455
               End
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
               Left            =   -70530
               Style           =   1  'Graphical
               TabIndex        =   23
               ToolTipText     =   "Agregar"
               Top             =   1770
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
               Left            =   -70530
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "Quitar"
               Top             =   2190
               Width           =   375
            End
            Begin VB.CommandButton cmdActualizar 
               Caption         =   "A"
               Height          =   375
               Left            =   -70530
               TabIndex        =   21
               ToolTipText     =   "Actualizar"
               Top             =   1350
               Width           =   375
            End
            Begin VB.CommandButton cmdAtras 
               Caption         =   "<="
               Height          =   375
               Left            =   -70530
               TabIndex        =   20
               ToolTipText     =   "Atras"
               Top             =   900
               Width           =   375
            End
            Begin VB.CheckBox chkFiltro 
               Caption         =   "Tiene Filtro?"
               Height          =   525
               Left            =   9930
               TabIndex        =   3
               Top             =   420
               Width           =   1275
            End
            Begin TrueOleDBGrid60.TDBGrid tdgCamposVista 
               Bindings        =   "frmVistaUsuario.frx":00FD
               Height          =   2355
               Left            =   180
               OleObjectBlob   =   "frmVistaUsuario.frx":011D
               TabIndex        =   4
               Top             =   420
               Width           =   9585
            End
            Begin TrueOleDBGrid60.TDBGrid tdgCamposFiltro 
               Bindings        =   "frmVistaUsuario.frx":417C
               Height          =   2355
               Left            =   -70020
               OleObjectBlob   =   "frmVistaUsuario.frx":419C
               TabIndex        =   19
               Top             =   450
               Width           =   6195
            End
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion"
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   15
            Top             =   780
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Id"
            Height          =   195
            Index           =   3
            Left            =   600
            TabIndex        =   14
            Top             =   1260
            Width           =   135
         End
         Begin VB.Label lblCodVistaUsuario 
            Caption         =   "Label1"
            Height          =   255
            Left            =   2520
            TabIndex        =   13
            Top             =   390
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo"
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   12
            Top             =   390
            Width           =   495
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            Height          =   195
            Index           =   2
            Left            =   600
            TabIndex        =   11
            Top             =   1710
            Width           =   315
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmVistaUsuario.frx":81F8
         Height          =   6225
         Left            =   -74640
         OleObjectBlob   =   "frmVistaUsuario.frx":8212
         TabIndex        =   16
         Top             =   720
         Width           =   11385
      End
      Begin TAMControls.ucBotonEdicion cmdSalir 
         Height          =   390
         Left            =   -64620
         TabIndex        =   17
         Top             =   7260
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   688
         Caption0        =   "&Salir"
         Tag0            =   "9"
         Visible0        =   0   'False
         ToolTipText0    =   "Salir"
         UserControlHeight=   390
         UserControlWidth=   1200
      End
      Begin TAMControls.ucBotonEdicion cmdOpcion 
         Height          =   390
         Left            =   -74610
         TabIndex        =   18
         Top             =   7260
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   688
         Buttons         =   3
         Caption0        =   "&Nuevo"
         Tag0            =   "0"
         Visible0        =   0   'False
         ToolTipText0    =   "Nuevo"
         Caption1        =   "&Modificar"
         Tag1            =   "1"
         Visible1        =   0   'False
         ToolTipText1    =   "Modificar"
         Caption2        =   "&Eliminar"
         Tag2            =   "4"
         Visible2        =   0   'False
         ToolTipText2    =   "Eliminar"
         UserControlHeight=   390
         UserControlWidth=   4200
      End
      Begin TAMControls.ucBotonEdicion cmdAccion 
         Height          =   390
         Left            =   8580
         TabIndex        =   29
         Top             =   8190
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   688
         Buttons         =   2
         Caption0        =   "&Guardar"
         Tag0            =   "2"
         Visible0        =   0   'False
         ToolTipText0    =   "Guardar"
         Caption1        =   "&Cancelar"
         Tag1            =   "8"
         Visible1        =   0   'False
         ToolTipText1    =   "Cancelar"
         UserControlHeight=   390
         UserControlWidth=   2700
      End
   End
End
Attribute VB_Name = "frmVistaUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim strEstado            As String
    Dim adoConsulta          As ADODB.Recordset
    Dim arrTipoVista()          As String, strCodTipoVista  As String
    Dim strStart As String, arrIdVariableFiltro() As String
    Dim adoRegistroCampos    As ADODB.Recordset, adoRegistroCamposFiltro As ADODB.Recordset
    Dim arrCamposGrilla() As String, arrValorIDCamposGrilla() As String
    
'**********************************BMM  19/02/2012*****************************

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
    strEstado = Reg_Defecto
    tabVista.Tab = 0
    tabVista.TabEnabled(1) = False

    
    lblCodVistaUsuario.FontBold = True
    
    frQuery.FontBold = True
    frQuery.ForeColor = &H800000
    
    chkFiltro.FontBold = True
    chkFiltro.ForeColor = &H800000
    lblIdVariableFiltro.FontBold = True
    
    cmdAtras.Visible = False
    cmdActualizar.Enabled = False

        
    tabCampos.Visible = False
    tabCampos.TabVisible(1) = False
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me

End Sub

Private Sub CargarListas()

  Dim strSQL As String
  
  strSQL = "SELECT CodParametro AS CODIGO, DescripParametro AS DESCRIP " & _
            "From AuxiliarParametro WHERE CodTipoParametro='TIPVIS'"
  CargarControlLista strSQL, cboTipoVista, arrTipoVista(), Sel_Defecto

  strSQL = "SELECT IdVariable AS CODIGO,IdVariable + ' ( ' + TipoDato + ' ) ' AS DESCRIP " & _
                            "FROM VariableUsuario " & _
                            "WHERE TipoVariable='02' AND IndVigente='X' " & _
                            "ORDER BY IdVariable"
                            
  CargarControlLista strSQL, cboIdVariableFiltro, arrIdVariableFiltro(), Sel_Defecto
    
End Sub

Private Sub CargarReportes()

'   frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Diario General"
    
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Diario General (ME)"
    
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Mayor General"
    
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Text = "Mayor General (ME)"
    
    
End Sub

Public Sub Buscar()

    Dim strSQL As String
    
    Set adoConsulta = New ADODB.Recordset
           
    strSQL = "SELECT CodVistaUsuario,DescripVista,IdVista,ValorFormulaVista,TipoVista,DescripParametro AS DescripTipo " & _
                "FROM VistaUsuario VU JOIN AuxiliarParametro AP ON (VU.TipoVista=AP.CodParametro " & _
                "AND AP.CodTipoParametro='TIPVIS') WHERE IndVigente='X'"
                        
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
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
            
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vQuery
            Call Modificar
        Case vDelete
            Call Eliminar
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

Public Sub Adicionar()
    If strEstado = Reg_Consulta Then
        strEstado = Reg_Adicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabVista
            .TabEnabled(0) = False
            .Tab = 1
            .TabEnabled(1) = True
        End With
    End If
    
End Sub

Public Sub Modificar()
    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabVista
            .TabEnabled(0) = False
            .Tab = 1
            .TabEnabled(1) = True
        End With
    End If
    
End Sub

Public Sub Grabar()

    Dim adoRegistro As ADODB.Recordset
    Dim objVistaUsuarioCampoXML  As DOMDocument60
    Dim strMsgError  As String
    Dim strVistaUsuarioCampoXML As String
    
    Set adoRegistro = New ADODB.Recordset

    If strEstado = Reg_Consulta Then Exit Sub
    
    If strEstado = Reg_Adicion Then
        
    If MsgBox(Mensaje_Adicion, vbQuestion + vbYesNo + vbDefaultButton2, gstrNombreEmpresa) = vbNo Then Exit Sub

        If TodoOK() Then
                
                Me.MousePointer = vbHourglass
                         
                On Error GoTo Ctrl_Error
                         
               '*** Guardar ***
                With adoComm
                
                    .CommandText = "SELECT * FROM VistaUsuario WHERE IdVista='" & _
                                    Trim(txtIdVista.Text) & "'"
                    
                    Set adoRegistro = .Execute
                    
                    If adoRegistro.EOF Then
                        
                        Call XMLADORecordset(objVistaUsuarioCampoXML, "VistaUsuarioCampo", "Estructura", adoRegistroCampos, strMsgError)
                         
                        If chkFiltro.Value = Checked Then
                            Call AdicionarCamposFiltroXML(objVistaUsuarioCampoXML, "Estructura", adoRegistroCamposFiltro, strMsgError)
                        End If
                        
                        strVistaUsuarioCampoXML = objVistaUsuarioCampoXML.xml
                        
                        .CommandText = "{ call up_ACManVistaUsuarioXML('" & _
                        lblCodVistaUsuario.Caption & "','" & txtDescripcion.Text & "','" & Trim(txtIdVista.Text) & _
                        "','" & strCodTipoVista & "','" & Replace(Trim(txtValorFormulaVista.Text), "'", "''") & "','X','" & _
                        strVistaUsuarioCampoXML & "','I') }"
    
                        adoConn.Execute .CommandText
                    
                    Else
                    
                        MsgBox "Ya Existe una Vista con este Id", vbCritical
                        txtIdVista.SetFocus
                        Me.MousePointer = vbDefault
                        Exit Sub
                        
                    End If
                
                    
                    
                End With
                

                Me.MousePointer = vbDefault

                MsgBox Mensaje_Adicion_Exitosa, vbExclamation

                frmMainMdi.stbMdi.Panels(3).Text = "Acción"

                cmdOpcion.Visible = True
                With tabVista
                    .TabEnabled(0) = True
                    .Tab = 0
                    .TabEnabled(1) = False
                End With

                Call Limpiar
                Call Buscar
                
                tdgCamposVista.DataSource = Nothing '
                tdgCamposFiltro.DataSource = Nothing '
 
        End If
    End If
    
    If strEstado = Reg_Edicion Then
    
    If MsgBox(Mensaje_Edicion, vbQuestion + vbYesNo + vbDefaultButton2, gstrNombreEmpresa) = vbNo Then Exit Sub

        If TodoOK() Then

                Me.MousePointer = vbHourglass
                
                
                On Error GoTo Ctrl_Error
                
                '*** Guardar ***
                With adoComm
                
                    Call XMLADORecordset(objVistaUsuarioCampoXML, "VistaUsuarioCampo", "Estructura", adoRegistroCampos, strMsgError)
                        
                        
                    If chkFiltro.Value = Checked Then
                        Call AdicionarCamposFiltroXML(objVistaUsuarioCampoXML, "Estructura", adoRegistroCamposFiltro, strMsgError)
                    End If
                        
                    strVistaUsuarioCampoXML = objVistaUsuarioCampoXML.xml
                        
                    .CommandText = "{ call up_ACManVistaUsuarioXML('" & _
                    lblCodVistaUsuario.Caption & "','" & txtDescripcion.Text & "','" & Trim(txtIdVista.Text) & _
                    "','" & strCodTipoVista & "','" & Trim(txtValorFormulaVista.Text) & "','X','" & _
                    strVistaUsuarioCampoXML & "','U') }"

                    adoConn.Execute .CommandText

                End With

                Me.MousePointer = vbDefault

                MsgBox Mensaje_Edicion_Exitosa, vbExclamation

                frmMainMdi.stbMdi.Panels(3).Text = "Acción"

                cmdOpcion.Visible = True
                With tabVista
                    .TabEnabled(0) = True
                    .Tab = 0
                    .TabEnabled(1) = False
                End With

                Call Limpiar
                Call Buscar
                
                tdgCamposVista.DataSource = Nothing '
                tdgCamposFiltro.DataSource = Nothing '
        End If
    
    End If


Exit Sub

Ctrl_Error:

'        adoComm.CommandText = "ROLLBACK TRAN ProcAsiento"
'        adoConn.Execute adoComm.CommandText
        
        MsgBox err.Description & vbCrLf & Mensaje_Proceso_NoExitoso, vbCritical
        Me.MousePointer = vbDefault
        

End Sub

Public Sub Eliminar()

    If MsgBox(Mensaje_Eliminacion, vbQuestion + vbYesNo + vbDefaultButton2, gstrNombreEmpresa) = vbNo Then Exit Sub
    
    If strEstado = Reg_Consulta Then
    
            Me.MousePointer = vbHourglass
                
                '*** Guardar ***
            With adoComm
                .CommandText = "{ call up_ACManVistaUsuarioXML('" & _
                tdgConsulta.Columns(0) & "','','','','','','','D') }"
                
                adoConn.Execute .CommandText
                
            End With
    
            Me.MousePointer = vbDefault
                        
            MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
                        
            Call Buscar
    
    End If

End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    cmdAccion.Button(0).Enabled = True
    
    With tabVista
        .TabEnabled(0) = True
        .Tab = 0
        .TabEnabled(1) = False
    End With
    
    Call Limpiar
    Call LimpiarFiltro
    Call Buscar
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub


Public Sub LlenarFormulario(ByVal strModo As String)

    Dim strCodVistaUsuario As String, strSQL As String
    
    Dim adoRegistro As ADODB.Recordset
    
    Dim intCont As Integer, intRegistro As Integer
    
'    txtValorFormulaVista.SelStart = 0
''    txtValorFormulaVista.SelLength = intTamaño
'    txtValorFormulaVista.SelColor = &H808000
    txtValorFormulaVista.Font.Bold = False
    
    Select Case strModo
    
    Case Reg_Adicion
        
        cboTipoVista.ListIndex = 0
        
        frVista.Caption = "Nueva Vista"
        frVista.ForeColor = &H800000
        frVista.FontBold = True
        frVista.Font = "Arial"
        
        
        lblCodVistaUsuario.Caption = NuevoCodigo()
        txtValorFormulaVista.Text = ""
        txtIdVista.Locked = False
        txtIdVista.BackColor = &H80000005
        
        txtValorFormulaVista.Locked = False
        txtValorFormulaVista.BackColor = &H80000005
        cmdAccion.Button(0).Enabled = False
        cmdEditarQuery.Caption = "Ok"
        cmdEditarQuery.Enabled = False
        cboIdVariableFiltro.ListIndex = 0
        chkFiltro.Value = Unchecked
        
        txtDescripcion.SetFocus
        
        Call ConfiguraRecordsetCampos
        
        Call CargarTipoDatos
        
        tdgCamposFiltro.DataSource = adoRegistroCamposFiltro
        
        Call BloquearEdicionGrilla
        
        tabCampos.Tab = 0
        tdgCamposVista.Caption = "Campos - Nueva Vista"
        tdgCamposFiltro.Caption = "Campos - Nueva Vista"
        
    Case Reg_Edicion
    
        cboTipoVista.ListIndex = 0
            
        strCodVistaUsuario = tdgConsulta.Columns(0)
        
        lblCodVistaUsuario.Caption = strCodVistaUsuario
        
        txtDescripcion.Text = tdgConsulta.Columns(1)
        
        frVista.Caption = "Vista: " + tdgConsulta.Columns(1)
        frVista.ForeColor = &H800000
        frVista.FontBold = True
        frVista.Font = "Arial"
        
        txtIdVista.Text = tdgConsulta.Columns(2)
        txtIdVista.Locked = True
        txtIdVista.BackColor = &H8000000F
        cmdEditarQuery.Caption = "Editar Query"
        cboIdVariableFiltro.ListIndex = 0
                
        Set adoRegistro = New ADODB.Recordset
            
            intRegistro = ObtenerItemLista(arrTipoVista, tdgConsulta.Columns(4))
            If cboTipoVista.ListCount > 0 Then cboTipoVista.ListIndex = intRegistro
            
            txtValorFormulaVista.Text = tdgConsulta.Columns(3)
            
            CambiarColorSQL
            
            txtValorFormulaVista.Locked = True
            txtValorFormulaVista.BackColor = &HE0E0E0
            
            cmdEditarQuery.Visible = True
            cmdEditarQuery.SetFocus
        
        Call ConfiguraRecordsetCampos
        
        strSQL = "SELECT CodVistaUsuario,SecCampo,NombreCampo,DescripParametro AS DescripTipo," & _
                     "TipoCampo , IdVariable FROM VistaUsuarioCampo VUC JOIN AuxiliarParametro AP " & _
                        "ON(VUC.TipoCampo=AP.CodParametro AND AP.CodTipoParametro='SUBREP') " & _
                        "WHERE CodVistaUsuario='" & strCodVistaUsuario & "' AND TipoCampo<>'03'"
        
        Call CargarGrillaCampos(strSQL, strCodVistaUsuario, True)
        
        Call CargarTipoDatos
        
        strSQL = "SELECT CodVistaUsuario,SecCampo,NombreCampo,DescripParametro AS DescripTipo," & _
                     "TipoCampo , IdVariable FROM VistaUsuarioCampo VUC JOIN AuxiliarParametro AP " & _
                        "ON(VUC.TipoCampo=AP.CodParametro AND AP.CodTipoParametro='SUBREP') " & _
                        "WHERE CodVistaUsuario='" & strCodVistaUsuario & "' AND TipoCampo='03'"
                        
        Call CargarGrillaCamposFiltro(strSQL)
        
        Call BloquearEdicionGrilla
        
        If adoRegistroCampos.RecordCount > 0 Then

            tabCampos.Visible = True
            tabCampos.Tab = 0
            tdgCamposVista.Caption = "Campos - " & Trim(txtDescripcion.Text)
        Else
            tabCampos.Visible = False

        End If
        
        If adoRegistroCamposFiltro.RecordCount > 0 Then
        
            tdgCamposFiltro.Caption = "Campos - " & Trim(txtDescripcion.Text)
            chkFiltro.Value = Checked
        Else
            chkFiltro.Value = Unchecked

        End If
               
    End Select
    
End Sub

Private Sub cboIdVariableFiltro_Click()

    lblIdVariableFiltro.Caption = arrIdVariableFiltro(cboIdVariableFiltro.ListIndex)

End Sub

Private Sub tdgConsulta_DblClick()
    Call Modificar
End Sub

Private Sub cboTipoVista_Click()

    strCodTipoVista = Valor_Caracter
    
    If cboTipoVista.ListIndex < 0 Then Exit Sub
    
    strCodTipoVista = arrTipoVista(cboTipoVista.ListIndex)
    
End Sub

Private Sub cmdEditarQuery_Click()
    
    If txtValorFormulaVista.Locked Then
        
        cmdEditarQuery.Caption = "Ok"
        txtValorFormulaVista.Locked = False
        txtValorFormulaVista.BackColor = &H80000005
        txtValorFormulaVista.SetFocus
        tabCampos.Visible = False
        cmdAccion.Button(0).Enabled = False
        
    Else
        
        If ValidarSQL(txtValorFormulaVista.Text) Then  'FALTA MODIFICAR ESTE FUNCION PARA QUE QUITA LOS CARACTERES INPUT/OUTPUT
        
            cmdEditarQuery.Caption = "Editar Query"
            txtValorFormulaVista.Locked = True
            txtValorFormulaVista.BackColor = &HE0E0E0
            cmdEditarQuery.SetFocus
            
            Call CargarGrillaCampos(txtValorFormulaVista.Text, lblCodVistaUsuario.Caption) '
            
            '
            
            If adoRegistroCampos.RecordCount > 0 Then

                tabCampos.Visible = True
                tabCampos.Tab = 0
                
                If strEstado = Reg_Edicion Then '
            
                    Call CargarVariablesAnteriores '
            
                End If '
                
            End If
            
            cmdAccion.Button(0).Enabled = True
        
        Else
        
            txtValorFormulaVista.SetFocus
            Exit Sub
            
        End If
    
    End If
    
    If Len(Trim(txtValorFormulaVista.Text)) > 0 Then
        cmdEditarQuery.Enabled = True
    Else
        cmdEditarQuery.Enabled = False
    End If
    
    
End Sub

Private Sub chkFiltro_Click()

    If chkFiltro.Value = Checked Then
    
        tabCampos.TabVisible(1) = True
    Else

        tabCampos.TabVisible(1) = False
    End If

End Sub

Private Sub cmdAgregar_Click()

    Dim intSecuencial As Integer
    
    If TodoOkFiltro() Then
    
        If ValidarCampoFiltro(Trim(txtCampoFiltro.Text)) Then
    
            intSecuencial = CInt(adoRegistroCamposFiltro.RecordCount) + 1
            
            adoRegistroCamposFiltro.AddNew
            adoRegistroCamposFiltro.Fields("CodVistaUsuario") = lblCodVistaUsuario.Caption
            adoRegistroCamposFiltro.Fields("SecCampo") = intSecuencial
            adoRegistroCamposFiltro.Fields("TipoCampo") = Tipo_Campo_Filtro
            adoRegistroCamposFiltro.Fields("DescripTipo") = "Filtro"
            adoRegistroCamposFiltro.Fields("NombreCampo") = Trim(txtCampoFiltro.Text)
            adoRegistroCamposFiltro.Fields("IdVariable") = Trim(lblIdVariableFiltro.Caption)
            
            adoRegistroCamposFiltro.Update
            
            Call LimpiarFiltro
        
        End If
    
    End If
    
End Sub

Private Sub cmdActualizar_Click()

    If TodoOkFiltro() Then
        
        If ValidarCampoFiltro(Trim(txtCampoFiltro.Text)) Then
        
            adoRegistroCamposFiltro.Fields("CodVistaUsuario") = lblCodVistaUsuario.Caption
            adoRegistroCamposFiltro.Fields("SecCampo") = intSecuencial ''''
            adoRegistroCamposFiltro.Fields("TipoCampo") = Tipo_Campo_Output
            adoRegistroCamposFiltro.Fields("DescripTipo") = "Filtro"
            adoRegistroCamposFiltro.Fields("NombreCampo") = Trim(txtCampoFiltro.Text)
            adoRegistroCamposFiltro.Fields("IdVariable") = Trim(lblIdVariableFiltro.Caption)
            
            Call cmdAtras_Click
        
        End If
        
    End If
    
End Sub

Private Sub cmdQuitar_Click()

    Dim dblBook As String
    Dim intIndice As Integer
    
    If adoRegistroCamposFiltro.RecordCount > 0 Then
              
        dblBook = adoRegistroCamposFiltro.AbsolutePosition

        adoRegistroCamposFiltro.Delete

        '**********SUBIR O BAJAR CURSOR SEGUN ELIMINE********
        If adoRegistroCamposFiltro.RecordCount >= 1 And dblBook = 1 Then
        
            adoRegistroCamposFiltro.MoveFirst
            tdgCamposFiltro.MoveFirst

        ElseIf adoRegistroCamposFiltro.RecordCount = 0 Then

            cmdQuitar.Enabled = False

        Else

            intIndice = CInt(dblBook) - 1
            dblBook = intIndice
            adoRegistroCamposFiltro.AbsolutePosition = CDbl(dblBook)

        End If
    
        Call LimpiarFiltro
    
    End If

End Sub

Private Sub cmdAtras_Click()

    cmdAtras.Visible = False
    
    cmdActualizar.Enabled = False
    
    cmdAgregar.Enabled = True
    
    cmdQuitar.Enabled = True
    
    tdgCamposFiltro.Enabled = True
    
    Call LimpiarFiltro
    
End Sub

Private Sub tdgCamposFiltro_DblClick()
        
        Dim intRegistro As Integer
        
        txtCampoFiltro.Text = Trim(tdgCamposFiltro.Columns(2))
        
        intRegistro = ObtenerItemLista(arrIdVariableFiltro, tdgCamposFiltro.Columns(5).Value)
        If intRegistro >= 0 Then cboIdVariableFiltro.ListIndex = intRegistro

        cmdAtras.Visible = True
        
        cmdActualizar.Enabled = True
        
        cmdAgregar.Enabled = False
        
        cmdQuitar.Enabled = False
        
        tdgCamposFiltro.Enabled = False
        
        cmdActualizar.SetFocus
End Sub

Private Function TodoOkFiltro() As Boolean
    
    TodoOkFiltro = False
    
    If Trim(txtCampoFiltro.Text) = Valor_Caracter Then
        MsgBox "El campo no puede estar vacio", vbCritical
        txtCampoFiltro.SetFocus
        Exit Function
    End If
    
    If cboIdVariableFiltro.ListIndex <= 0 Then
        MsgBox "Debe seleccionar un elemento", vbCritical
        cboIdVariableFiltro.SetFocus
        Exit Function
    End If
    
    TodoOkFiltro = True

End Function

Private Function TodoOK() As Boolean

    Dim adoRegistroAux As ADODB.Recordset
    
    Set adoRegistroAux = New ADODB.Recordset

    TodoOK = False

    If Trim(txtDescripcion.Text) = Valor_Caracter Then
        MsgBox "La Descripcion de la Vista no puede estar en blanco", vbCritical
        txtDescripcion.SetFocus
        Exit Function
    End If

    If Trim(txtIdVista.Text) = Valor_Caracter Then
        MsgBox "La Id la Vista no puede estar en blanco", vbCritical
        txtIdVista.SetFocus
        Exit Function
    End If
    
    If cboTipoVista.ListIndex <= 0 Then
        MsgBox "No ah seleccionado el tipo de Vista", vbCritical
        cboTipoVista.SetFocus
        Exit Function
    End If
    
    If Trim(txtValorFormulaVista.Text) = Valor_Caracter Then
        MsgBox "El query no puede estar vacio", vbCritical
        txtValorFormulaVista.SetFocus
        Exit Function
    End If
    
    
    If chkFiltro.Value = Checked Then

        If adoRegistroCamposFiltro.RecordCount = 0 Then
        
            MsgBox "No ah indicado ningun campo de filtro", vbCritical
            tabCampos.Tab = 1
            Exit Function
        
        End If

    End If
    
    If adoRegistroCampos.RecordCount > 0 Then
    
        Set adoRegistroAux = adoRegistroCampos.Clone
                
        adoRegistroAux.MoveFirst
        
        Do Until adoRegistroAux.EOF
            
            If Trim(adoRegistroAux.Fields("IdVariable").Value) = "" Then
            
                MsgBox "Debe asignar un IdVariable a todos los campos", vbCritical
                tabCampos.Tab = 0
                Exit Function
                
            End If
            
            adoRegistroAux.MoveNext
        Loop
        
        
        adoRegistroAux.Close: Set adoRegistroAux = Nothing
        
    End If
    

    '*** Si todo paso OK ***
    TodoOK = True

End Function

Private Sub Limpiar()
    lblCodVistaUsuario.Caption = Valor_Caracter
    txtDescripcion.Text = Valor_Caracter
    txtIdVista.Text = Valor_Caracter
    txtValorFormulaVista.Text = Valor_Caracter
    tabCampos.Visible = False
    chkFiltro.Value = Unchecked
End Sub

Private Sub LimpiarFiltro()

    txtCampoFiltro.Text = Valor_Caracter
    cboIdVariableFiltro.ListIndex = 0

End Sub
'
Public Function NuevoCodigo() As String

    Dim adoRegistro As ADODB.Recordset
    Dim strSQL As String
    Dim strCODIGO As String, intCodigo As Integer
  
    NuevoCodigo = Valor_Caracter

    Set adoRegistro = New ADODB.Recordset

    With adoComm

        strSQL = "SELECT MAX(CodVistaUsuario) AS CodVistaUsuario FROM VistaUsuario"

        .CommandText = strSQL

        Set adoRegistro = .Execute

        Do Until adoRegistro.EOF

            strCODIGO = adoRegistro("CodVistaUsuario")
            adoRegistro.MoveNext

        Loop

    End With

        intCodigo = CInt(strCODIGO) + 1

        strCODIGO = CStr(intCodigo)

        Select Case Len(strCODIGO)

            Case 1

            strCODIGO = "00" + strCODIGO

            Case 2

            strCODIGO = "0" + strCODIGO

        End Select

    adoRegistro.Close: Set adoRegistro = Nothing

    NuevoCodigo = strCODIGO

End Function

Private Sub CambiarColorSQL()
    
    Dim arrSQL() As String
    Dim i As Integer, intPos As Long, j As Integer
    Dim nStart As Integer
    Dim intTamaño As Integer
    Dim arrCaracteresEspeciales() As String
    
    nStart = 0

    ReDim arrSQL(2)

    arrSQL(0) = "SELECT"
    arrSQL(1) = "FROM"
    arrSQL(2) = "WHERE"

    '1º ASIGNANDO COLOR A TODO EL TEXTO

    intTamaño = Len(txtValorFormulaVista.Text)

    txtValorFormulaVista.SelStart = 0
    txtValorFormulaVista.SelLength = intTamaño
    txtValorFormulaVista.SelColor = &H808000
    txtValorFormulaVista.SelBold = False

    'PALABRAS AZUL
    For i = 0 To UBound(arrSQL)

        txtValorFormulaVista.SelStart = 0

        While nStart <> intTamaño - 1

        intPos = txtValorFormulaVista.Find(arrSQL(i), nStart)

        If intPos <> -1 Then
            txtValorFormulaVista.SelStart = intPos
            txtValorFormulaVista.SelLength = Len(arrSQL(i))
            txtValorFormulaVista.SelColor = ColorConstants.vbBlue
            txtValorFormulaVista.SelBold = True

            nStart = intPos + Len(arrSQL(i))
        Else

            nStart = intTamaño - 1

        End If

        Wend

        nStart = 0

    Next

    nStart = 0

    ReDim arrSQL(11)

    arrSQL(0) = "AND"
    arrSQL(1) = "OR"
    arrSQL(2) = "IN"
    arrSQL(3) = "BETWEEN"
    arrSQL(4) = "*"
    arrSQL(5) = "+"
    arrSQL(6) = "-"
    arrSQL(7) = "/"
    arrSQL(8) = "="
    arrSQL(9) = ">"
    arrSQL(10) = "<"
    arrSQL(11) = "!="

    'PALABRAS PLOMO
    For i = 0 To UBound(arrSQL)

        txtValorFormulaVista.SelStart = 0

        While nStart <> intTamaño - 1

        If Len(arrSQL(i)) = 1 Then

            intPos = txtValorFormulaVista.Find(arrSQL(i), nStart)
        Else
            intPos = txtValorFormulaVista.Find(arrSQL(i), nStart, , rtfWholeWord)
        End If

        If intPos <> -1 Then
            txtValorFormulaVista.SelStart = intPos
            txtValorFormulaVista.SelLength = Len(arrSQL(i))
            txtValorFormulaVista.SelColor = &H808080
            txtValorFormulaVista.SelBold = True

            nStart = intPos + Len(arrSQL(i))
        Else

            nStart = intTamaño - 1

        End If

        Wend

        nStart = 0

    Next
    
End Sub

Private Sub txtValorFormulaVista_Change()
       
    strStart = txtValorFormulaVista.SelStart
    
    Invisible.SetFocus
    txtValorFormulaVista.Visible = False
    
    CambiarColorSQL
    
    txtValorFormulaVista.Visible = True
    
    txtValorFormulaVista.SelStart = CInt(strStart)
    txtValorFormulaVista.SelBold = False
    txtValorFormulaVista.SetFocus
    
    If Len(Trim(txtValorFormulaVista.Text)) > 0 Then
        cmdEditarQuery.Enabled = True
    Else
        cmdEditarQuery.Enabled = False
    End If
    
    
    
End Sub

Private Function ValidarSQL(ByVal strSQL As String) As Boolean
    
    ValidarSQL = False
    
    Dim myPattern As String, myPatternIO As String
    Dim objRegex As RegExp, objRegex2 As RegExp
    Dim objMatch As Match, objMatch2 As Match
    Dim colMatches   As MatchCollection, colMatches2 As MatchCollection
    Dim strExec As String, strSQL2 As String, strSQL3 As String
    Dim arrFound() As String, arrFoundIO() As String, arrFoundTrueFieldIO() As String
    Dim intCont As Integer, i As Integer, intCont2 As Integer, j As Integer
    Dim intBeforeWhere As Long
    
    intCont = 0
        
    Set objRegExp = New RegExp
    Set objRegex2 = New RegExp
    
    myPattern = "@[a-zA-Z][\w]*"
    myPatternIO = "\[@([a-zA-Z][\w]*)\]"
    
    objRegExp.Pattern = myPattern
    objRegExp.Global = True
    
    objRegex2.Pattern = myPatternIO
    objRegex2.Global = True
    
    On Error GoTo Ctrl_E
           
    'CAPTURANDO EL INDICE DEL WHERE
    intBeforeWhere = InStr(1, strSQL, "FROM", vbBinaryCompare)
    
    'intBeforeWhere = InStr(1, strSQL, "WHERE", vbBinaryCompare) //CUANDO ERA POR EL WHERE
        
    If intBeforeWhere <> 0 Then
        
        'NUEVA CADENA DESPUES DEL SELECT PARA ASEGURARSE QUE SOLO REEMPLAZE EN LA PARTE DE LA CONDICION
        strSQL2 = Mid(strSQL, intBeforeWhere, (Len(strSQL) - intBeforeWhere) + 1)
            
        'CAPTURANDO EL SELECT PARA REEMPLAZAR LOS PARAMETROS INPUT/OUTPUT
        strSQL3 = Mid(strSQL, 1, intBeforeWhere - 1)
            
        Set colMatches = objRegExp.Execute(strSQL2)
        Set colMatches2 = objRegex2.Execute(strSQL3)
        
        If (colMatches.Count > 0) Then
            
            intCont = 0
            
            'CAPTURANDO VARIABLES @"
            For Each objMatch In colMatches
                
                ReDim Preserve arrFound(intCont)
                arrFound(intCont) = objMatch.Value
                
                intCont = intCont + 1
            Next
                        
            strExec = strSQL2
            
            'REEMPLAZANDO LAS VARIABLES @ POR '' PARA EJECUTAR EL QUERY Y VERIFICAR QUE SEA
            'UNA INSTRUCCION VALIDA
            For i = 0 To UBound(arrFound)
                
                strExec = Replace(strExec, arrFound(i), "''")
            
            Next
            
            'REEMPLAZANDO EL QUERY ORIGINAL POR EL LA NUEVA PARTE'
            strExec = Replace(strSQL, strSQL2, strExec)
            
        End If
        
        If (colMatches2.Count > 0) Then
            
            intCont2 = 0
            
            'CAPTURANDO VARIABLES INPUT/OUTPUT
            For Each objMatch2 In colMatches2
                ReDim Preserve arrFoundIO(intCont2)
                ReDim Preserve arrFoundTrueFieldIO(intCont2)
                arrFoundIO(intCont2) = objMatch2.Value
                arrFoundTrueFieldIO(intCont2) = CStr(objMatch2.SubMatches(0))
                intCont2 = intCont2 + 1
            Next
            
            'Reemplazando los campos I/O por lis true I/O
            For j = 0 To UBound(arrFoundIO)
                strExec = Replace(strExec, arrFoundIO(j), arrFoundTrueFieldIO(j))
            Next
            
        End If
    
    Else
    
        strExec = strSQL
        
    End If

    
    With adoComm
    
        .CommandText = strExec
        
        adoConn.Execute .CommandText
        
    End With


    ValidarSQL = True
    
Exit Function

Ctrl_E:
    
    MsgBox "Error de sintaxis en Query SQL " & vbCrLf & err.Description, vbCritical
    Exit Function

End Function

Private Sub ConfiguraRecordsetCampos()

    Set adoRegistroCampos = New ADODB.Recordset

    With adoRegistroCampos
       .CursorLocation = adUseClient
       .Fields.Append "CodVistaUsuario", adChar, 3
       .Fields.Append "SecCampo", adInteger
       .Fields.Append "NombreCampo", adVarChar, 500
       .Fields.Append "DescripTipo", adVarChar, 60
       .Fields.Append "TipoCampo", adVarChar, 2
       .Fields.Append "IdVariable", adVarChar, 200
'       .CursorType = adOpenStatic

       .LockType = adLockBatchOptimistic
    End With
    
    adoRegistroCampos.Open
    
    
    Set adoRegistroCamposFiltro = New ADODB.Recordset

    With adoRegistroCamposFiltro
       .CursorLocation = adUseClient
       .Fields.Append "CodVistaUsuario", adChar, 3
       .Fields.Append "SecCampo", adInteger
       .Fields.Append "NombreCampo", adVarChar, 500
       .Fields.Append "DescripTipo", adVarChar, 60
       .Fields.Append "TipoCampo", adVarChar, 2
       .Fields.Append "IdVariable", adVarChar, 200
'       .CursorType = adOpenStatic

       .LockType = adLockBatchOptimistic
    End With
    
    adoRegistroCamposFiltro.Open

End Sub

Private Sub CargarGrillaCampos(ByVal strSQL As String, ByVal strCodVistaUsuario As String, Optional ByVal firstLoad As Boolean = False)
    
    Dim adoRegistro As ADODB.Recordset
    Dim strExec As String
    Dim patternInput As String, patternOutput As String, patternInputOutput As String
    Dim strSQL2 As String
    Dim intSecuencial As Integer, i As Integer, indice As Integer
    Dim matchCollect As MatchCollection, matchFound As Match
    Dim objRegExp As RegExp
    Dim arrOutput() As String
    Dim intBeforeWhere As Long
    
    Dim strTMPInputOutput As String
        
    Set adoRegistro = New ADODB.Recordset
    
    indice = 0 '
    
    If firstLoad = True Then
        With adoComm
        
            .CommandText = strSQL
            
            Set adoRegistro = .Execute
            
            Do Until adoRegistro.EOF
            
                'ARRAYS PARA ALMACENAR LOS CAMPOS Y SUS ID DE VARIABLES ORIGINALES
                ReDim Preserve arrCamposGrilla(indice) '
                ReDim Preserve arrValorIDCamposGrilla(indice) '
                
                adoRegistroCampos.AddNew
                adoRegistroCampos.Fields("CodVistaUsuario") = adoRegistro.Fields("CodVistaUsuario")
                adoRegistroCampos.Fields("SecCampo") = adoRegistro.Fields("SecCampo")
                adoRegistroCampos.Fields("TipoCampo") = adoRegistro.Fields("TipoCampo")
                adoRegistroCampos.Fields("DescripTipo") = adoRegistro.Fields("DescripTipo")
                adoRegistroCampos.Fields("NombreCampo") = adoRegistro.Fields("NombreCampo")
                adoRegistroCampos.Fields("IdVariable") = adoRegistro.Fields("IdVariable")
                
                'ASIGNAR VALORES A LOS ARRAYS
                arrCamposGrilla(indice) = adoRegistro.Fields("NombreCampo")
                arrValorIDCamposGrilla(indice) = adoRegistro.Fields("IdVariable")
                
                adoRegistroCampos.Update
                
                indice = indice + 1 '
                adoRegistro.MoveNext
                
            Loop

        End With

        tdgCamposVista.DataSource = adoRegistroCampos
           
    Else
        intSecuencial = 1
        strExec = ""
        patternOutput = "[Ss][Ee][Ll][Ee][Cc][Tt]\s*([\s|\S]*)\s*[Ff][Rr][Oo][Mm]"
        patternInput = "@[a-zA-Z][\w]*"
        patternInputOutput = "\[@([a-zA-Z][\w]*)\]" 'cualquier campo que este entre corchetes y lleve un arroba al principio es un Input/Output
        
        'RECUPERO EL CAMPO OUTPUT DEL QUERY SQL
    
        Set objRegExp = New RegExp
        
        objRegExp.Pattern = patternOutput
        
        objRegExp.Global = True
        
        Set matchCollect = objRegExp.Execute(strSQL)
        
        If matchCollect.Count > 0 Then
    
            For Each matchFound In matchCollect
                arrOutput = Split(CStr(matchFound.SubMatches(0)), ",")
            Next
    
            If UBound(arrOutput) <> -1 Then
                
                'ELIMINA REGISTROS SI EL RECORDSET YA LOS TIENE
                
                If Not adoRegistroCampos.EOF Or adoRegistroCampos.RecordCount > 0 Then
                
                    adoRegistroCampos.MoveFirst
                    
                    Do Until adoRegistroCampos.EOF
                    
                        adoRegistroCampos.Delete
                        
                        adoRegistroCampos.MoveNext
                    
                    Loop
                
                End If
                
                
                Set objRegExp = New RegExp
        
                objRegExp.Pattern = patternInputOutput
                
                objRegExp.Global = True
                
                For i = 0 To UBound(arrOutput)
                    
                adoRegistroCampos.AddNew
                adoRegistroCampos.Fields("CodVistaUsuario") = strCodVistaUsuario
                adoRegistroCampos.Fields("SecCampo") = intSecuencial
                
                Set matchCollect = objRegExp.Execute(Trim(arrOutput(i)))
                
                If matchCollect.Count > 0 Then
                    adoRegistroCampos.Fields("TipoCampo") = Tipo_Campo_InputOutput
                    adoRegistroCampos.Fields("DescripTipo") = "Input/Output"
                    
                    For Each matchFound In matchCollect
                        strTMPInputOutput = Trim(CStr(matchFound.SubMatches(0)))
                    Next
                    
                    arrOutput(i) = strTMPInputOutput
                
                    arrOutput(i) = Replace(strTMPInputOutput, vbCrLf, "")
                    arrOutput(i) = Replace(strTMPInputOutput, vbTab, "")
                
                    adoRegistroCampos.Fields("NombreCampo") = strTMPInputOutput
                    
                Else
                    adoRegistroCampos.Fields("TipoCampo") = Tipo_Campo_Output
                    adoRegistroCampos.Fields("DescripTipo") = "Output"
                    
                    arrOutput(i) = Trim(arrOutput(i))
                
                    arrOutput(i) = Replace(arrOutput(i), vbCrLf, "")
                    arrOutput(i) = Replace(arrOutput(i), vbTab, "")
                
                    adoRegistroCampos.Fields("NombreCampo") = arrOutput(i)
                    
                End If
            
                adoRegistroCampos.Fields("IdVariable") = ""
                
                adoRegistroCampos.Update
                
                intSecuencial = intSecuencial + 1
    
                Next
                
            End If
            
        End If
        
        intSecuencial = 1
        
        Set objRegExp = New RegExp
        
        objRegExp.Pattern = patternInput
        objRegExp.Global = True
        
        'EXTRAIGO EL WHERE DE LA CONSULTA PARA EXTRAER LOS CAMPOS INPUT
        
        'CAPTURANDO EL INDICE DEL WHERE
        intBeforeWhere = InStr(1, strSQL, "FROM", vbBinaryCompare)
        
        'intBeforeWhere = InStr(1, strSQL, "WHERE", vbBinaryCompare) //CUANDO ERA POR EL WHERE
        
        If intBeforeWhere <> 0 Then
            'NUEVA CADENA DESPUES DEL SELECT PARA ASEGURARSE QUE SOLO REEMPLAZE EN LA PARTE DE LA CONDICION
            strSQL2 = Mid(strSQL, intBeforeWhere, (Len(strSQL) - intBeforeWhere) + 1)
            
            Set matchCollect = objRegExp.Execute(strSQL2)
            
            If matchCollect.Count > 0 Then
                For Each matchFound In matchCollect
                    adoRegistroCampos.AddNew
                    adoRegistroCampos.Fields("CodVistaUsuario") = strCodVistaUsuario
                    adoRegistroCampos.Fields("SecCampo") = intSecuencial
                    adoRegistroCampos.Fields("TipoCampo") = Tipo_Campo_Input
                    adoRegistroCampos.Fields("DescripTipo") = "Input"
                    adoRegistroCampos.Fields("NombreCampo") = matchFound.Value
                    adoRegistroCampos.Fields("IdVariable") = ""
                    
                    adoRegistroCampos.Update
                    
                    intSecuencial = intSecuencial + 1
                Next
            End If
        End If
        
        tdgCamposVista.DataSource = adoRegistroCampos
    End If
End Sub

Private Sub CargarGrillaCamposFiltro(ByVal strSQL As String)
    
    Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
    
        With adoComm
        
            .CommandText = strSQL
            
            Set adoRegistro = .Execute
            
            Do Until adoRegistro.EOF
                
                adoRegistroCamposFiltro.AddNew
                adoRegistroCamposFiltro.Fields("CodVistaUsuario") = adoRegistro.Fields("CodVistaUsuario")
                adoRegistroCamposFiltro.Fields("SecCampo") = adoRegistro.Fields("SecCampo")
                adoRegistroCamposFiltro.Fields("TipoCampo") = adoRegistro.Fields("TipoCampo")
                adoRegistroCamposFiltro.Fields("DescripTipo") = adoRegistro.Fields("DescripTipo")
                adoRegistroCamposFiltro.Fields("NombreCampo") = adoRegistro.Fields("NombreCampo")
                adoRegistroCamposFiltro.Fields("IdVariable") = adoRegistro.Fields("IdVariable")
                
                adoRegistroCamposFiltro.Update
        
                adoRegistro.MoveNext
                
            Loop

        End With

        tdgCamposFiltro.DataSource = adoRegistroCamposFiltro

End Sub

Public Sub CargarTipoDatos()

    Dim adoRegistro As ADODB.Recordset
    Dim Item As New ValueItem
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
    
        .CommandText = "SELECT IdVariable AS CODIGO,DescripVariable AS DESCRIP,TipoDato AS TIPDAT " & _
                            "FROM VariableUsuario " & _
                            "WHERE TipoVariable='02' AND IndVigente='X' " & _
                            "ORDER BY IdVariable"
        
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
        
            adoRegistro.MoveFirst
            
            While adoRegistro.EOF = False
                
                With tdgCamposVista.Columns("IdVariable").ValueItems
                    Item.Value = adoRegistro.Fields("CODIGO").Value
                    Item.DisplayValue = adoRegistro.Fields("CODIGO").Value & _
                                            " ( " & adoRegistro.Fields("TIPDAT").Value & " )"
                    .Add Item
                    .Translate = True
                    .Presentation = 2
                    
                End With
                    
                adoRegistro.MoveNext
            Wend

        End If
        adoRegistro.Close
    End With

End Sub

Public Function ValidarCampoFiltro(ByVal strCampo As String) As Boolean
    
    Dim adoRegistroAux As ADODB.Recordset
    
    Set adoRegistroAux = New ADODB.Recordset
    
    ValidarCampoFiltro = False
    
    Dim strSQL As String, strSQL2 As String
    
    Dim intIndexWhere As Integer
    
    strSQL = Trim(txtValorFormulaVista.Text)
    
    'CAPTURANDO EL INDICE DEL WHERE
    intIndexWhere = InStr(1, strSQL, "WHERE", vbBinaryCompare)

    
    On Error GoTo Ctrl_Error
    
    If intIndexWhere <> 0 Then
        
        'CADENA ANTES DEL WHERE PARA PROBAR SI EL CAMPO PASADO ES VALIDO
        strSQL2 = Mid(strSQL, 1, intIndexWhere - 1)
        strSQL2 = strSQL2 & " WHERE " & strCampo & "=''"
    
    Else

        strSQL2 = strSQL & " WHERE " & strCampo & "=''"

    End If
    
    With adoComm
    
        .CommandText = strSQL2
        
        adoConn.Execute .CommandText
    
    End With
    
    If Not adoRegistroCamposFiltro.EOF Then
    
        Set adoRegistroAux = adoRegistroCamposFiltro.Clone '
        
        adoRegistroAux.MoveFirst
        
        Do Until adoRegistroAux.EOF
        
            If (Trim(txtCampoFiltro.Text) = adoRegistroAux("NombreCampo").Value) Then
                    
                 GoTo Ctrl_Identidad
                    
            End If
            
            adoRegistroAux.MoveNext
        
        Loop
    
    End If
    
    ValidarCampoFiltro = True
    
Exit Function

Ctrl_Error:
    
MsgBox "(Filtro) " & err.Description, vbCritical
txtCampoFiltro.SetFocus

Exit Function

Ctrl_Identidad:

MsgBox "(Filtro) Ya existe este campo en la grilla", vbCritical
txtCampoFiltro.SetFocus

End Function

Public Sub BloquearEdicionGrilla()

    tdgCamposVista.Columns(0).Locked = True
    tdgCamposVista.Columns(1).Locked = True
    tdgCamposVista.Columns(2).Locked = True
    tdgCamposVista.Columns(3).Locked = True
    tdgCamposVista.Columns(4).Locked = True
    tdgCamposVista.Columns(5).DropDownList = True
    
    tdgCamposFiltro.Columns(0).Locked = True
    tdgCamposFiltro.Columns(1).Locked = True
    tdgCamposFiltro.Columns(2).Locked = True
    tdgCamposFiltro.Columns(3).Locked = True
    tdgCamposFiltro.Columns(4).Locked = True
    tdgCamposFiltro.Columns(5).Locked = True

End Sub

Public Sub CargarVariablesAnteriores()
    
    Dim i As Integer
    
    adoRegistroCampos.MoveFirst
    
    Do Until adoRegistroCampos.EOF
    
        For i = 0 To UBound(arrCamposGrilla)
    
            If adoRegistroCampos.Fields("NombreCampo") = arrCamposGrilla(i) Then
            
                adoRegistroCampos.Fields("IdVariable") = arrValorIDCamposGrilla(i)
            
            End If
    
        Next
    
        adoRegistroCampos.MoveNext
    
    Loop
    
End Sub

Public Sub AdicionarCamposFiltroXML(ByRef objXML As DOMDocument60, ByVal strNomEntidad As String, ByVal adoRecordset As ADODB.Recordset, ByRef strMsgError As String)

    Dim n As Integer, j As Integer
    Dim objParent As MSXML2.IXMLDOMElement
    Dim objElem As MSXML2.IXMLDOMElement
    Dim NomCampos() As String
    Dim adoField As ADODB.Field
    
    
    Set objParent = objXML.documentElement

    On Error GoTo ErrCreaXMLADORecordset
    
    n = 0
    
    'Recorriendo todas las filas de una rejilla
    If adoRecordset.RecordCount > 0 Then
                   
            For Each adoField In adoRecordset.Fields
                ReDim Preserve NomCampos(n)
                NomCampos(n) = adoField.Name
                n = n + 1
            Next
                
        adoRecordset.MoveFirst
        
        Do While Not adoRecordset.EOF
            
                Set objElem = objParent.appendChild(objXML.createElement(strNomEntidad))
                For j = 0 To UBound(NomCampos)
                    'añadiendo los atributos, solo para las columnas especificadas
                    objElem.setAttribute NomCampos(j), "" & adoRecordset.Fields(NomCampos(j)).Value
                Next
            
            
            adoRecordset.MoveNext
        Loop
    End If
    
    Set objElem = Nothing
    Set objParent = Nothing
    
    Exit Sub
ErrCreaXMLADORecordset:
    If strMsgError = "" Then strMsgError = "(XMLADORecordset) - " & err.Description

End Sub


