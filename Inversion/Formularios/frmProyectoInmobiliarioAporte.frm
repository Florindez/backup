VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmProyectoInmobiliarioConstitucion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Constitucion - Proyecto Inmobiliario"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11655
   Begin TabDlg.SSTab tabProyectoInmobiliario 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   14631
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Listado"
      TabPicture(0)   =   "frmProyectoInmobiliarioAporte.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCriterio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSalir"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdOpcion"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Constitucion"
      TabPicture(1)   =   "frmProyectoInmobiliarioAporte.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmDatosProyecto"
      Tab(1).Control(1)=   "frmAportes"
      Tab(1).Control(2)=   "cmdAccion"
      Tab(1).ControlCount=   3
      Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
         Height          =   735
         Left            =   480
         TabIndex        =   36
         Top             =   7080
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   1296
         Buttons         =   3
         Caption0        =   "&Ver Activos"
         Tag0            =   "0"
         Visible0        =   0   'False
         ToolTipText0    =   "Ver Activos"
         Caption1        =   "&Buscar"
         Tag1            =   "5"
         Visible1        =   0   'False
         ToolTipText1    =   "Buscar"
         Caption2        =   "&Imprimir"
         Tag2            =   "6"
         Visible2        =   0   'False
         ToolTipText2    =   "Imprimir"
         UserControlWidth=   4200
      End
      Begin TAMControls2.ucBotonEdicion2 cmdSalir 
         Height          =   735
         Left            =   9720
         TabIndex        =   35
         Top             =   7080
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1296
         Caption0        =   "&Salir"
         Tag0            =   "9"
         Visible0        =   0   'False
         ToolTipText0    =   "Salir"
         UserControlWidth=   1200
      End
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -67080
         TabIndex        =   34
         Top             =   7320
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
      Begin VB.Frame frmAportes 
         Caption         =   "Detalle Activos"
         Height          =   3855
         Left            =   -74880
         TabIndex        =   25
         Top             =   3360
         Width           =   11175
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "Quitar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3720
            Picture         =   "frmProyectoInmobiliarioAporte.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   2160
            Width           =   855
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3720
            Picture         =   "frmProyectoInmobiliarioAporte.frx":025A
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   1080
            Width           =   855
         End
         Begin VB.ListBox lstActivoDisponibles 
            Height          =   2595
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   3495
         End
         Begin VB.ComboBox cboTipoActivoProyectoInmobiliario 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   360
            Width           =   2145
         End
         Begin TrueOleDBGrid60.TDBGrid tdgProyectoActivos 
            Bindings        =   "frmProyectoInmobiliarioAporte.frx":046D
            Height          =   3375
            Left            =   4680
            OleObjectBlob   =   "frmProyectoInmobiliarioAporte.frx":0487
            TabIndex        =   30
            Top             =   240
            Width           =   6345
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Activo"
            Height          =   195
            Index           =   11
            Left            =   360
            TabIndex        =   28
            Top             =   360
            Width           =   1185
         End
      End
      Begin VB.Frame frmDatosProyecto 
         Caption         =   "Proyecto Inmobiliario"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   12
         Top             =   480
         Width           =   11175
         Begin VB.Label lblFondo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2640
            TabIndex        =   27
            Top             =   360
            Width           =   7545
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fondo"
            Height          =   195
            Index           =   10
            Left            =   960
            TabIndex        =   26
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label lblMontoProyectadoProyecto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   6480
            TabIndex        =   24
            Top             =   2280
            Width           =   2145
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Proyectado"
            Height          =   195
            Index           =   9
            Left            =   5040
            TabIndex        =   23
            Top             =   2280
            Width           =   1305
         End
         Begin VB.Label lblMontoTotalProyecto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2640
            TabIndex        =   22
            Top             =   2280
            Width           =   2145
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Total"
            Height          =   195
            Index           =   8
            Left            =   960
            TabIndex        =   21
            Top             =   2280
            Width           =   1545
         End
         Begin VB.Label lblDescripUnidadSuperficie 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2640
            TabIndex        =   20
            Top             =   1800
            Width           =   4545
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Unidad Superficie"
            Height          =   195
            Index           =   7
            Left            =   960
            TabIndex        =   19
            Top             =   1800
            Width           =   1545
         End
         Begin VB.Label lblDescripProyecto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2640
            TabIndex        =   18
            Top             =   1320
            Width           =   7545
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Descripcion"
            Height          =   195
            Index           =   6
            Left            =   960
            TabIndex        =   17
            Top             =   1320
            Width           =   1185
         End
         Begin VB.Label lblTipoProyecto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   5640
            TabIndex        =   16
            Top             =   840
            Width           =   4545
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo"
            Height          =   195
            Index           =   5
            Left            =   4800
            TabIndex        =   15
            Top             =   840
            Width           =   825
         End
         Begin VB.Label lblCodTitulo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2640
            TabIndex        =   14
            Top             =   840
            Width           =   1665
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Codigo Titulo"
            Height          =   195
            Index           =   4
            Left            =   960
            TabIndex        =   13
            Top             =   840
            Width           =   1185
         End
      End
      Begin VB.Frame fraCriterio 
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
         Height          =   1335
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   10695
         Begin VB.CheckBox chkFiltrarFechas 
            Caption         =   "Filtrar"
            Height          =   255
            Left            =   6840
            TabIndex        =   11
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox cboTipoProyectoInmobiliario 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   780
            Width           =   5145
         End
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   360
            Width           =   5145
         End
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   285
            Left            =   8880
            TabIndex        =   4
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   175505409
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaHasta 
            Height          =   285
            Left            =   8880
            TabIndex        =   5
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   175505409
            CurrentDate     =   38785
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo"
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   9
            Top             =   780
            Width           =   825
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fondo"
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   8
            Top             =   360
            Width           =   705
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   3
            Left            =   7920
            TabIndex        =   7
            Top             =   840
            Width           =   825
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Desde"
            Height          =   195
            Index           =   2
            Left            =   7920
            TabIndex        =   6
            Top             =   360
            Width           =   705
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmProyectoInmobiliarioAporte.frx":3EF2
         Height          =   4815
         Left            =   120
         OleObjectBlob   =   "frmProyectoInmobiliarioAporte.frx":3F0C
         TabIndex        =   10
         Top             =   2040
         Width           =   11145
      End
   End
End
Attribute VB_Name = "frmProyectoInmobiliarioConstitucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSQL As String, strEstado As String
Dim adoConsulta As ADODB.Recordset
Dim adoRegistroActivosProyectoInmoAux As ADODB.Recordset
Dim arrFondo() As String, arrTipoProyectoInmobiliario() As String, arrTipoActivoProyectoInmobiliario() As String
Dim arrActivo() As String
Dim strCodFondo As String, strTipoProyectoInmobiliario As String, strTipoActivoProyectoInmobiliario As String
Dim strItemsOriginales As String, strItemsEnLista As String

Private Sub Form_Load()
    Call InicializarValores
    Call CargarListas
    Call Buscar
    Call DarFormato
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
        
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
                
        Case vNew
            Call Adicionar
        Case vSearch
            Call Buscar
        Case vSave
            Call Grabar
        Case vCancel
            Call Cancelar
        Case vPrint
            Call SubImprimir
        Case vExit
            Call Salir
        
    End Select
    
End Sub

Private Sub InicializarValores()
    dtpFechaDesde.Value = gdatFechaActual
    dtpFechaHasta.Value = gdatFechaActual
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    
    tabProyectoInmobiliario.TabEnabled(0) = True
    tabProyectoInmobiliario.TabVisible(1) = False
    tabProyectoInmobiliario.Tab = 0
    
    strItemsOriginales = Valor_Caracter
    strItemsEnLista = Valor_Caracter
    
    Call chkFiltrarFechas_Click
    
End Sub

Private Sub CargarListas()

    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
       
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    strSQL = "SELECT CodFile CODIGO,DescripFile DESCRIP " & _
            "FROM InversionFile WHERE CodFile IN ('030','040') AND IndVigente='X'"
    CargarControlLista strSQL, cboTipoProyectoInmobiliario, arrTipoProyectoInmobiliario(), Sel_Todos
    
    If cboTipoProyectoInmobiliario.ListCount > 0 Then cboTipoProyectoInmobiliario.ListIndex = 0
    
    strSQL = "{ call up_IVListarTipoActivoProyectoInmobiliario }"
    CargarControlLista strSQL, cboTipoActivoProyectoInmobiliario, arrTipoActivoProyectoInmobiliario(), Valor_Caracter
    
    If cboTipoActivoProyectoInmobiliario.ListCount > 0 Then cboTipoActivoProyectoInmobiliario.ListIndex = 0
    
End Sub

Private Sub Buscar()
    
    Me.MousePointer = vbHourglass
    
    strSQL = "SELECT CodTitulo,CodAnalitica,FechaDefinicion," & _
        "DescripProyecto,TipoUnidadMedida,AP.ValorParametro DescripUnidadMedida," & _
        "FIP.CodFile,DescripFile DescripTipoProyecto," & _
        "CONVERT(VARCHAR(50),CONVERT(BIGINT,MontoTotalGeneral)) + ' ' + LTRIM(RTRIM(AP.ValorParametro)) MontoTotalGeneral," & _
        "CONVERT(VARCHAR(50),CONVERT(BIGINT,MontoTotalProyectado)) + ' ' + LTRIM(RTRIM(AP.ValorParametro)) MontoTotalProyectado " & _
        "FROM FondoInmobiliarioProyecto FIP " & _
        "JOIN AuxiliarParametro AP ON FIP.TipoUnidadMedida=AP.CodParametro AND AP.CodTipoParametro='UNDSUP' " & _
        "JOIN InversionFile INVF ON FIP.CodFile=INVF.CodFile " & _
        "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
        " FIP.CodFile LIKE '" & IIf(Trim(strTipoProyectoInmobiliario) = Valor_Caracter, "%", Trim(strTipoProyectoInmobiliario)) & "' AND FIP.IndVigente='X' "
    
    If chkFiltrarFechas.Value Then
        strSQL = strSQL & " AND CONVERT(DATE,FechaDefinicion)>='" & Convertyyyymmdd(dtpFechaDesde.Value) & "' AND CONVERT(DATE,FechaDefinicion)<='" & Convertyyyymmdd(dtpFechaHasta.Value) & "' "
    End If
    
    strSQL = strSQL & " ORDER BY FechaDefinicion"
    
    strEstado = Reg_Defecto
    
    Set adoConsulta = New ADODB.Recordset
    
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    tdgConsulta.DataSource = adoConsulta
    
    tdgConsulta.Refresh
    Call AutoAjustarGrillas
    Me.Refresh
    If adoConsulta.RecordCount > 0 Then strEstado = Reg_Consulta

    Me.MousePointer = vbDefault
    
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

Private Sub AutoAjustarGrillas()

    Dim i As Integer
    
    If Not adoConsulta Is Nothing Then
        If Not adoConsulta.EOF Then
            If adoConsulta.RecordCount > 0 Then
                For i = 0 To tdgConsulta.Columns.Count - 1
                    tdgConsulta.Columns(i).AutoSize
                Next
                tdgConsulta.Columns(0).AutoSize
                tdgConsulta.Columns(3).AutoSize
                tdgConsulta.Columns(6).AutoSize
            End If
        End If
    End If
    
    If Not adoRegistroActivosProyectoInmoAux Is Nothing Then
        If Not adoRegistroActivosProyectoInmoAux.EOF Then
            If adoRegistroActivosProyectoInmoAux.RecordCount > 0 Then
                For i = 0 To tdgProyectoActivos.Columns.Count - 1
                    tdgProyectoActivos.Columns(i).AutoSize
                Next
                tdgProyectoActivos.Columns(4).AutoSize
            End If
        End If
    End If

End Sub

Private Sub Adicionar()
    
    Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
    
    If tdgConsulta.SelBookmarks.Count <= 0 Then Exit Sub
    
    lblFondo.Caption = cboFondo.Text
    
    With adoComm
        
        .CommandText = "{ call up_IVObtenerDatosProyectoInmobiliario('" & strCodFondo & "','" & _
                 gstrCodAdministradora & "','" & Trim(tdgConsulta.Columns(0).Value) & "') }"
        Set adoRegistro = .Execute
    
        If Not adoRegistro.EOF Then
            While Not adoRegistro.EOF
                lblCodTitulo.Caption = Trim(adoRegistro.Fields("CodTitulo"))
                lblTipoProyecto.Caption = Trim(adoRegistro.Fields("DescripTipo"))
                lblDescripProyecto.Caption = Trim(adoRegistro.Fields("DescripProyecto"))
                lblDescripUnidadSuperficie.Caption = Trim(adoRegistro.Fields("DescripUnidadSuperficie"))
                lblMontoTotalProyecto.Caption = Trim(adoRegistro.Fields("MontoTotalTexto"))
                lblMontoProyectadoProyecto.Caption = Trim(adoRegistro.Fields("MontoProyectadoTexto"))
                adoRegistro.MoveNext
            Wend
        End If
    
        Call ConfiguraRecordsetAuxiliar
    
        .CommandText = "{ call up_IVObtenerActivosProyectoInmobiliario('" & strCodFondo & "','" & _
                 gstrCodAdministradora & "','" & Trim(tdgConsulta.Columns(0).Value) & "') }"
        
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            While Not adoRegistro.EOF
                adoRegistroActivosProyectoInmoAux.AddNew
                For Each adoField In adoRegistroActivosProyectoInmoAux.Fields
                    adoRegistroActivosProyectoInmoAux.Fields(adoField.Name) = adoRegistro.Fields(adoField.Name)
                    If adoField.Name = "CodFile" Or adoField.Name = "CodAnalitica" Then
                        strItemsOriginales = strItemsOriginales & adoRegistro.Fields(adoField.Name)
                    End If
                Next
                adoRegistroActivosProyectoInmoAux.Update
                adoRegistro.MoveNext
                strItemsOriginales = strItemsOriginales & "|"
            Wend
            strItemsEnLista = strItemsOriginales
            cmdQuitar.Enabled = True
        Else
            cmdQuitar.Enabled = False
        End If
        
        tdgProyectoActivos.DataSource = adoRegistroActivosProyectoInmoAux
        Call AutoAjustarGrillas
        tdgProyectoActivos.Refresh
    
    End With
    
    Call cboTipoActivoProyectoInmobiliario_Click
    
    tabProyectoInmobiliario.TabEnabled(0) = False
    tabProyectoInmobiliario.TabVisible(1) = True
    tabProyectoInmobiliario.Tab = 1
    
End Sub

Private Sub ConfiguraRecordsetAuxiliar()

    Set adoRegistroActivosProyectoInmoAux = New ADODB.Recordset

    With adoRegistroActivosProyectoInmoAux
       .CursorLocation = adUseClient
       .Fields.Append "TipoActivo", adChar, 2
       .Fields.Append "DescripTipoActivo", adVarChar, 8000
       .Fields.Append "CodFile", adChar, 3
       .Fields.Append "CodAnalitica", adChar, 8
       .Fields.Append "DescripActivo", adVarChar, 8000
'       .CursorType = adOpenStatic
       .LockType = adLockBatchOptimistic
    End With

    adoRegistroActivosProyectoInmoAux.Open
    
End Sub

Private Sub CalcularFiltro()
    
    Dim adoRegistroClone As ADODB.Recordset
    
    Set adoRegistroClone = adoRegistroActivosProyectoInmoAux.Clone
    
    If Not adoRegistroClone.EOF Then
        adoRegistroClone.MoveFirst
        strItemsEnLista = Valor_Caracter
        While Not adoRegistroClone.EOF
           For Each adoField In adoRegistroClone.Fields
               If adoField.Name = "CodFile" Or adoField.Name = "CodAnalitica" Then
                   strItemsEnLista = strItemsEnLista & adoRegistroClone.Fields(adoField.Name)
               End If
           Next
           adoRegistroClone.MoveNext
           strItemsEnLista = strItemsEnLista & "|"
        Wend
    End If
    
End Sub

Private Function TodoOK() As Boolean
    
    TodoOK = False
    
    If adoRegistroActivosProyectoInmoAux.RecordCount = 0 Then
        MsgBox "El proyecto inmobiliario no tiene ningun activo", vbCritical, Me.Caption
        Exit Function
    End If
    
    TodoOK = True
    
End Function

Private Sub Grabar()
        
    Dim objProyectoActivosXML  As DOMDocument60
    Dim strMsgError                 As String
    Dim strProyectoActivosXML As String
    Dim strCodTituloProyecto As String
    
    If TodoOK() Then
            
        Me.MousePointer = vbHourglass
        
        strCodTituloProyecto = lblCodTitulo.Caption
        
        Call XMLADORecordset(objProyectoActivosXML, "ProyectoActivos", "Activos", adoRegistroActivosProyectoInmoAux, strMsgError)
        strProyectoActivosXML = objProyectoActivosXML.xml
        
        With adoComm
            
            On Error GoTo Ctrl_Error
            
            .CommandText = "{ call up_IVAsigProyectoInmobiliarioActivo('" & strCodFondo & "','" & _
                 gstrCodAdministradora & "','" & strCodTituloProyecto & "','" & _
                 strProyectoActivosXML & "') }"
            
            adoConn.Execute .CommandText
             
        End With
        
        Set adoRegistroActivosProyectoInmoAux = Nothing
                
        Me.MousePointer = vbDefault
        
        MsgBox "Se guardaron los cambios exitosamente", vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        Call Cancelar
        
    End If
    
    Exit Sub
    
Ctrl_Error:
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Cancelar()
        
    cmdAgregar.Enabled = True
    cmdQuitar.Enabled = True
    If cboTipoActivoProyectoInmobiliario.ListCount > 0 Then cboTipoActivoProyectoInmobiliario.ListIndex = 0
    
    strItemsOriginales = Valor_Caracter
    strItemsEnLista = Valor_Caracter
    
    tabProyectoInmobiliario.TabEnabled(0) = True
    tabProyectoInmobiliario.TabVisible(1) = False
    tabProyectoInmobiliario.Tab = 0
    
End Sub

Private Sub Salir()
    Unload Me
End Sub

Private Sub Form_Resize()
    Call AutoAjustarGrillas
End Sub

Private Sub cboTipoActivoProyectoInmobiliario_Click()
    
    Dim adoRegistro As ADODB.Recordset, i As Integer, codigo As String
    
    strTipoActivoProyectoInmobiliario = Valor_Caracter
    If cboTipoActivoProyectoInmobiliario.ListIndex < 0 Then Exit Sub
    strTipoActivoProyectoInmobiliario = Trim(arrTipoActivoProyectoInmobiliario(cboTipoActivoProyectoInmobiliario.ListIndex))
    
    lstActivoDisponibles.Clear
    If lstActivoDisponibles.ListCount > 0 Then
        For i = 0 To lstActivoDisponibles.ListCount - 1
            lstActivoDisponibles.RemoveItem i
        Next
    End If
    
    If Not adoRegistroActivosProyectoInmoAux Is Nothing Then
        If adoRegistroActivosProyectoInmoAux.RecordCount <= 0 Then strItemsEnLista = Valor_Caracter
    End If
    
    strSQL = "{ call up_IVListarActivoDisponibleProyectoInmobiliario('" & strCodFondo & "','" & _
        gstrCodAdministradora & "','" & strTipoActivoProyectoInmobiliario & "','" & _
        strItemsOriginales & "','" & strItemsEnLista & "') }"
    
    CargarControlLista strSQL, lstActivoDisponibles, arrActivo(), Valor_Caracter
    
    If lstActivoDisponibles.ListCount > 0 Then cmdAgregar.Enabled = True Else cmdAgregar.Enabled = False
    
    Call AutoAjustarGrillas
    
End Sub

Private Sub chkFiltrarFechas_Click()
    If chkFiltrarFechas.Value Then
        dtpFechaDesde.Enabled = True
        dtpFechaHasta.Enabled = True
    Else
        dtpFechaDesde.Enabled = False
        dtpFechaHasta.Enabled = False
    End If
End Sub

Private Sub dtpFechaDesde_Change()
    If IsNull(dtpFechaDesde.Value) Then
        dtpFechaDesde.Value = gdatFechaActual
    Else
        If dtpFechaDesde.Value > dtpFechaHasta.Value Then
            dtpFechaDesde.Value = dtpFechaHasta.Value
        End If
    End If
End Sub

Private Sub dtpFechaHasta_Change()
    If IsNull(dtpFechaHasta.Value) Then
        dtpFechaHasta.Value = gdatFechaActual
    Else
        If dtpFechaHasta.Value < dtpFechaDesde.Value Then
            dtpFechaHasta.Value = dtpFechaDesde.Value
        End If
    End If
End Sub

Private Sub cboFondo_Click()
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
End Sub

Private Sub cboTipoProyectoInmobiliario_Click()
    strTipoProyectoInmobiliario = Valor_Caracter
    If cboTipoProyectoInmobiliario.ListIndex < 0 Then Exit Sub
    strTipoProyectoInmobiliario = Trim(arrTipoProyectoInmobiliario(cboTipoProyectoInmobiliario.ListIndex))
End Sub

Private Sub cmdAgregar_Click()

    Dim dblBookmark As Double
    Dim i As Integer, j As Integer, strTMP As String, arrTMP() As String
    Dim strCodigoFile As String, strCodigoAnalitica As String
    Dim arrNActivos() As String, indice As Integer
    i = lstActivoDisponibles.ListIndex
    
    If i <= -1 Then Exit Sub
    
    adoRegistroActivosProyectoInmoAux.AddNew
    adoRegistroActivosProyectoInmoAux.Fields("TipoActivo") = arrTipoActivoProyectoInmobiliario(cboTipoActivoProyectoInmobiliario.ListIndex)
    adoRegistroActivosProyectoInmoAux.Fields("DescripTipoActivo") = cboTipoActivoProyectoInmobiliario.Text
    
    strTMP = arrActivo(lstActivoDisponibles.ListIndex)
    arrTMP = Split(strTMP, "|")
    strCodigoFile = arrTMP(0)
    strCodigoAnalitica = arrTMP(1)
    
    adoRegistroActivosProyectoInmoAux.Fields("CodFile") = strCodigoFile
    adoRegistroActivosProyectoInmoAux.Fields("CodAnalitica") = strCodigoAnalitica
    adoRegistroActivosProyectoInmoAux.Fields("DescripActivo") = lstActivoDisponibles.Text
      
    lstActivoDisponibles.RemoveItem lstActivoDisponibles.ListIndex
     
    If lstActivoDisponibles.ListCount > 0 Then
        indice = 0
        ReDim Preserve arrNActivos(lstActivoDisponibles.ListCount - 1)
        For j = 0 To UBound(arrActivo)
            If j <> i Then
                arrNActivos(indice) = arrActivo(j)
                indice = indice + 1
            End If
        Next
        arrActivo = arrNActivos
    End If
    
    adoRegistroActivosProyectoInmoAux.Update
    dblBookmark = adoRegistroActivosProyectoInmoAux.Bookmark
      
    tdgProyectoActivos.DataSource = adoRegistroActivosProyectoInmoAux
    tdgProyectoActivos.Refresh
        
    adoRegistroActivosProyectoInmoAux.Bookmark = dblBookmark
    
    Call CalcularFiltro
     
    If lstActivoDisponibles.ListCount = 0 Then cmdAgregar.Enabled = False
    If adoRegistroActivosProyectoInmoAux.RecordCount > 0 Then cmdQuitar.Enabled = True
    
    Call AutoAjustarGrillas
    
End Sub

Private Sub cmdQuitar_Click()
    
    Dim dblBookmark As Double, n As Integer, strCODIGO As String
    
    If tdgProyectoActivos.SelBookmarks.Count > 0 Then
        
        If tdgProyectoActivos.Columns(0) = arrTipoActivoProyectoInmobiliario(cboTipoActivoProyectoInmobiliario.ListIndex) Then
            lstActivoDisponibles.AddItem adoRegistroActivosProyectoInmoAux.Fields("DescripActivo")
            n = UBound(arrActivo) + 1
            ReDim Preserve arrActivo(n)
            strCODIGO = adoRegistroActivosProyectoInmoAux.Fields("CodFile")
            strCODIGO = strCODIGO & "|"
            strCODIGO = strCODIGO & adoRegistroActivosProyectoInmoAux.Fields("CodAnalitica")
            arrActivo(n) = strCODIGO
        End If
        
        If lstActivoDisponibles.ListCount > 0 Then cmdAgregar.Enabled = True
        
        If adoRegistroActivosProyectoInmoAux.RecordCount > 0 Then
            dblBookmark = adoRegistroActivosProyectoInmoAux.Bookmark
            adoRegistroActivosProyectoInmoAux.Delete adAffectCurrent
            
            If adoRegistroActivosProyectoInmoAux.EOF Then
                adoRegistroActivosProyectoInmoAux.MovePrevious
                tdgProyectoActivos.MovePrevious
            End If
            
            adoRegistroActivosProyectoInmoAux.Update
            
            If adoRegistroActivosProyectoInmoAux.RecordCount > 0 And Not adoRegistroActivosProyectoInmoAux.BOF And Not adoRegistroActivosProyectoInmoAux.EOF And dblBookmark > 1 Then adoRegistroActivosProyectoInmoAux.Bookmark = dblBookmark - 1
    
            If adoRegistroActivosProyectoInmoAux.RecordCount > 0 And Not adoRegistroActivosProyectoInmoAux.BOF And Not adoRegistroActivosProyectoInmoAux.EOF Then adoRegistroActivosProyectoInmoAux.Bookmark = dblBookmark - 1
       
            tdgProyectoActivos.Refresh
            
            If adoRegistroActivosProyectoInmoAux.RecordCount = 0 Then
                cmdQuitar.Enabled = False
            End If
        End If
        
        Call CalcularFiltro
        
    End If
    
End Sub

Private Function RetornaFiltro() As String
    
'    Dim strResult As String
'    Dim i As Integer, j As Integer, arrOriginales() As String, arrActuales() As String
'
'    arrOriginales = Split(strItemsOriginales, "|")
'    arrActuales = Split(strItemsEnLista, "|")
'
'    strResult = strItemsOriginales
'
'    For i = 0 To UBound(arrOriginales)
'        For j = 0 To UBound(arrActuales)
'            If arrOriginales(i) = arrActuales(j) Then
'                arrActuales(j) = Valor_Caracter
'            End If
'        Next
'    Next
'
'    strResult = strResult & Join(arrActuales, "|")
'
'    RetornaFiltro = strResult
    
End Function

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

Private Sub SubImprimir()


    Dim strSeleccionRegistro    As String
    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strIndicador    As String, strFecDesde As String, strFecHasta   As String

    
   
            gstrNameRepo = "ConstitucionProInmobiliariaGrilla"
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(5)
            ReDim aReportParamFn(2)
            ReDim aReportParamF(2)
                        
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)


            If chkFiltrarFechas.Value Then
                strIndicador = "C"
                strFecDesde = Convertyyyymmdd(dtpFechaDesde.Value)
                strFecHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value))
            Else
                strIndicador = "S"
                strFecDesde = "20000101"
                strFecHasta = "20000101"
            End If
            
            If strTipoProyectoInmobiliario = Valor_Caracter Then
                strTipoProyectoInmobiliario = Valor_Comodin
            End If
                        
            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora
            aReportParamS(2) = strTipoProyectoInmobiliario
            aReportParamS(3) = strFecDesde
            aReportParamS(4) = strFecHasta
            aReportParamS(5) = strIndicador
            
           
       
    
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    


End Sub
