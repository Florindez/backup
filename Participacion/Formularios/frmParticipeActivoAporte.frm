VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{442CAE95-1D41-47B1-BE83-6995DA3CE254}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmParticipeActivoAporte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Participe - Activos"
   ClientHeight    =   7530
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   11685
   Begin TabDlg.SSTab tabParticipeActivo 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   12726
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
      TabCaption(0)   =   "Listado"
      TabPicture(0)   =   "frmParticipeActivoAporte.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdSalir"
      Tab(0).Control(1)=   "cmdOpcion"
      Tab(0).Control(2)=   "fraCriterio"
      Tab(0).Control(3)=   "tdgConsulta"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Datos"
      TabPicture(1)   =   "frmParticipeActivoAporte.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frmDatos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frmDetalleActivo"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdAccion"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   8280
         TabIndex        =   10
         Top             =   6240
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
      Begin TAMControls2.ucBotonEdicion2 cmdSalir 
         Height          =   735
         Left            =   -66240
         TabIndex        =   9
         Top             =   5280
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
         Left            =   -74640
         TabIndex        =   20
         Top             =   5280
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   1296
         Buttons         =   5
         Caption0        =   "&Nuevo"
         Tag0            =   "0"
         Visible0        =   0   'False
         ToolTipText0    =   "Nuevo"
         Caption1        =   "&Modificar"
         Tag1            =   "3"
         Visible1        =   0   'False
         ToolTipText1    =   "Modificar"
         Caption2        =   "&Buscar"
         Tag2            =   "5"
         Visible2        =   0   'False
         ToolTipText2    =   "Buscar"
         Caption3        =   "&Eliminar"
         Tag3            =   "4"
         Visible3        =   0   'False
         ToolTipText3    =   "Eliminar"
         Caption4        =   "&Imprimir"
         Tag4            =   "6"
         Visible4        =   0   'False
         ToolTipText4    =   "Imprimir"
         UserControlWidth=   7200
      End
      Begin VB.Frame frmDetalleActivo 
         Caption         =   "Detalle"
         Height          =   2475
         Left            =   270
         TabIndex        =   21
         Top             =   3720
         Width           =   10365
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   1350
            Width           =   3105
         End
         Begin VB.TextBox txtValorReferencial 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3330
            MaxLength       =   12
            TabIndex        =   37
            Text            =   " "
            Top             =   1860
            Width           =   1860
         End
         Begin VB.TextBox txtNumPartidaRegistral 
            Height          =   315
            Left            =   3360
            MaxLength       =   20
            TabIndex        =   30
            Top             =   360
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker dtpFechaUltimaTasacion 
            Height          =   285
            Left            =   3360
            TabIndex        =   28
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
            Format          =   118095873
            CurrentDate     =   38785
         End
         Begin TAMControls.TAMTextBox txtValorNominal 
            Height          =   315
            Left            =   7380
            TabIndex        =   29
            Top             =   660
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
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
            MaxLength       =   15
            Container       =   "frmParticipeActivoAporte.frx":0038
            Decimales       =   2
            Estilo          =   3
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   2000000000
         End
         Begin VB.Label lblSignoMoneda 
            Caption         =   "PEN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5550
            TabIndex        =   40
            Top             =   1920
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   12
            Left            =   1380
            TabIndex        =   39
            Top             =   1320
            Width           =   585
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Valor Referencial"
            Height          =   195
            Index           =   9
            Left            =   1140
            TabIndex        =   27
            Top             =   1800
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha Ultima Tasacion"
            Height          =   195
            Index           =   8
            Left            =   1080
            TabIndex        =   26
            Top             =   840
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Numero Partida"
            Height          =   195
            Index           =   7
            Left            =   1080
            TabIndex        =   25
            Top             =   360
            Width           =   1425
         End
      End
      Begin VB.Frame frmDatos 
         Caption         =   "Datos Basicos"
         Height          =   3075
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   10395
         Begin VB.TextBox txtDescripCuenta 
            Height          =   315
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   1410
            Width           =   5925
         End
         Begin VB.TextBox txtCodCuenta 
            Height          =   315
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   1410
            Width           =   1605
         End
         Begin VB.CommandButton cmdBusquedaCuenta 
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
            Left            =   8760
            TabIndex        =   35
            ToolTipText     =   "Búsqueda de Partícipe"
            Top             =   1380
            Width           =   375
         End
         Begin VB.ComboBox cboSubTipoActivo 
            Height          =   315
            Left            =   6720
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   990
            Width           =   2625
         End
         Begin VB.CommandButton cmdBusquedaParticipe 
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
            Left            =   8160
            TabIndex        =   24
            ToolTipText     =   "Búsqueda de Partícipe"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtUbicacionActivo 
            Height          =   315
            Left            =   2760
            MaxLength       =   200
            TabIndex        =   19
            Top             =   2490
            Width           =   5175
         End
         Begin VB.ComboBox cboTipoActivo 
            Height          =   315
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   990
            Width           =   2625
         End
         Begin VB.TextBox txtDescripActivo 
            Height          =   315
            Left            =   2760
            MaxLength       =   200
            TabIndex        =   15
            Top             =   1950
            Width           =   5175
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Cuenta"
            Height          =   285
            Index           =   11
            Left            =   1050
            TabIndex        =   34
            Top             =   1485
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            Caption         =   "SubTipo"
            Height          =   195
            Index           =   10
            Left            =   5520
            TabIndex        =   31
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label lblDescripParticipe 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2760
            TabIndex        =   23
            Top             =   450
            Width           =   5205
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Participe"
            Height          =   195
            Index           =   6
            Left            =   1080
            TabIndex        =   22
            Top             =   480
            Width           =   705
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Ubicacion"
            Height          =   195
            Index           =   4
            Left            =   1050
            TabIndex        =   18
            Top             =   2490
            Width           =   1185
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo"
            Height          =   195
            Index           =   5
            Left            =   1080
            TabIndex        =   16
            Top             =   960
            Width           =   1185
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Descripcion"
            Height          =   195
            Index           =   3
            Left            =   1020
            TabIndex        =   14
            Top             =   1950
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
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   9975
         Begin VB.CheckBox chkFiltrarParticipe 
            Caption         =   "Filtrar"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   855
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
            Index           =   1
            Left            =   6240
            TabIndex        =   2
            ToolTipText     =   "Búsqueda de Partícipe"
            Top             =   720
            Width           =   375
         End
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   285
            Left            =   8160
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
            Format          =   118095873
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaHasta 
            Height          =   285
            Left            =   8160
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
            Format          =   118095873
            CurrentDate     =   38785
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   2
            Left            =   7200
            TabIndex        =   12
            Top             =   840
            Width           =   705
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Desde"
            Height          =   195
            Index           =   1
            Left            =   7200
            TabIndex        =   11
            Top             =   360
            Width           =   705
         End
         Begin VB.Label lblDescripParticipeBusqueda 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   7
            Top             =   720
            Width           =   4725
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Participe"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   720
            Width           =   705
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmParticipeActivoAporte.frx":0054
         Height          =   3135
         Left            =   -74880
         OleObjectBlob   =   "frmParticipeActivoAporte.frx":006E
         TabIndex        =   8
         Top             =   2040
         Width           =   10185
      End
   End
End
Attribute VB_Name = "frmParticipeActivoAporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql                  As String, strEstado            As String
Dim strCodParticipeBusqueda As String, strCodParticipe      As String
Dim strTipoActivo           As String, strSubTipoActivo     As String
Dim arrTipoActivo()         As String, arrSubTipoActivo()   As String
Dim adoConsulta             As ADODB.Recordset
Dim indSortAsc              As Boolean, indSortDesc         As Boolean
Dim strCodCuenta            As String, strDescripCuenta
Dim arrMoneda()         As String, strCodMoneda         As String, strSignoMoneda   As String, strCodSignoMoneda As String

Private Sub cboMoneda_Click()

 
    strCodMoneda = Valor_Caracter: strSignoMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
    strSignoMoneda = ObtenerSignoMoneda(strCodMoneda)
    strCodSignoMoneda = ObtenerCodSignoMoneda(strCodMoneda)
    
    lblSignoMoneda.Caption = strCodSignoMoneda
           

End Sub

Private Sub cmdBusquedaCuenta_Click()

   Dim sSql As String
   
   Dim frmBus As frmBuscar
    
   Set frmBus = New frmBuscar
    
    With frmBus.TBuscarRegistro1
           
        .ADOConexion = adoConn
        .ADOConexion.CommandTimeout = 0
        'If Index <> 2 Then
        '    .iTipoGrilla = 1
        'Else
        '    .iTipoGrilla = 2
        .iTipoGrilla = 2
        
        
         frmBus.Caption = " Relación de Cuentas Contables"
         .sSql = "{ call up_PRObtenerCuentaActivo}"
         .OutputColumns = "1,2"
         .HiddenColumns = ""
        
        
        Screen.MousePointer = vbHourglass
                
        .BuscarTabla
        
        Screen.MousePointer = vbNormal
        frmBus.Show 1
        
        If .iParams.Count = 0 Then Exit Sub
        
        If .iParams(1).Valor <> "" Then
        
        
                    strCodCuenta = Trim(.iParams(1).Valor)
                    strDescripCuenta = Trim(.iParams(2).Valor)
                    
                    txtCodCuenta.Text = strCodCuenta
                    
                    txtDescripCuenta.Text = strCodCuenta & " - " & strDescripCuenta
                    
                    
                
        End If
        
        
    End With
    
    Set frmBus = Nothing
        


End Sub

Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call Buscar
    Call DarFormato
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
    
End Sub

Private Sub InicializarValores()
    dtpFechaDesde.Value = gdatFechaActual
    dtpFechaHasta.Value = gdatFechaActual
    dtpFechaUltimaTasacion.Value = gdatFechaActual
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    
    'cmdBusquedaParticipe.Enabled = False
    lblDescripParticipeBusqueda.Caption = Valor_Caracter
    strCodParticipeBusqueda = Valor_Caracter
    
    tabParticipeActivo.TabEnabled(0) = True
    tabParticipeActivo.TabVisible(1) = False
    tabParticipeActivo.Tab = 0
End Sub

Private Sub CargarListas()
    
    strSql = "SELECT CodFile CODIGO, DescripFile DESCRIP FROM InversionFile WHERE CodFile IN ('031','032')"
    CargarControlLista strSql, cboTipoActivo, arrTipoActivo(), Sel_Defecto
       
    If cboTipoActivo.ListCount > 0 Then cboTipoActivo.ListIndex = 0
    
    
        '*** Moneda ***
    strSql = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSql, cboMoneda, arrMoneda(), Sel_Defecto
    
End Sub

Private Sub Buscar()
    
    Me.MousePointer = vbHourglass
    
    strSql = "SELECT PAP.CodFile,CodAnalitica,FechaDefinicion," & _
            "IV.DescripFile DescripTipoActivo,DescripActivo,UbicacionActivo," & _
            "PAP.CodParticipe , DescripParticipe, ValorNominal " & _
            "FROM ParticipeActivoAporte PAP " & _
            "JOIN InversionFile IV ON PAP.CodFile=IV.CodFile " & _
            "JOIN ParticipeContrato P ON PAP.CodParticipe=P.CodParticipe " & _
            "LEFT JOIN InversionDetalleFile IDF ON " & _
            "PAP.CodDetalleFile = IDF.CodDetalleFile And PAP.CodFile = IDF.CodFile " & _
            "WHERE FechaDefinicion>='" & Convertyyyymmdd(dtpFechaDesde.Value) & "' AND " & _
            "FechaDefinicion<='" & Convertyyyymmdd(dtpFechaHasta.Value) & "' AND PAP.IndVigente='X' "
            
    If chkFiltrarParticipe.Value Then
        strSql = strSql & " AND PAP.CodParticipe='" & strCodParticipeBusqueda & "' "
    End If
    
    strSql = strSql & " ORDER BY FechaDefinicion"
    
    strEstado = Reg_Defecto
    
    Set adoConsulta = New ADODB.Recordset
    
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSql
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
        Case vSave
            Call Grabar
        Case vCancel
            Call Cancelar
        Case vExit
            Call Salir
        Case vPrint
            Call SubImprimir
        
    End Select
    
End Sub

Private Sub Adicionar()
    
    strEstado = Reg_Adicion
    
    tabParticipeActivo.TabEnabled(0) = False
    tabParticipeActivo.TabVisible(1) = True
    tabParticipeActivo.Tab = 1
    
    txtValorReferencial.Text = "0.00"
    
End Sub

Private Sub Modificar()
    
    Dim adoRegistro As ADODB.Recordset, intRegistro As Integer
    
    strEstado = Reg_Edicion
        
    Set adoRegistro = New ADODB.Recordset
    
    If tdgConsulta.SelBookmarks.Count <= 0 Then Exit Sub
    
    With adoComm
        
        .CommandText = "{ call up_IVObtenerDatosParticipeActivo('" & Trim(tdgConsulta.Columns(0).Value) & "','" & _
                 Trim(tdgConsulta.Columns(1).Value) & "') }"

        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            While Not adoRegistro.EOF
            
                strCodParticipe = Trim(adoRegistro.Fields("CodParticipe"))
                lblDescripParticipe.Caption = Trim(adoRegistro.Fields("DescripParticipe"))
                txtDescripActivo.Text = Trim(adoRegistro.Fields("DescripActivo"))
                intRegistro = ObtenerItemLista(arrTipoActivo(), Trim(adoRegistro.Fields("CodFile")))
                If intRegistro >= 0 Then cboTipoActivo.ListIndex = intRegistro
                txtDescripActivo.Text = Trim(adoRegistro.Fields("DescripActivo"))
                txtUbicacionActivo.Text = Trim(adoRegistro.Fields("UbicacionActivo"))
                txtNumPartidaRegistral.Text = Trim(adoRegistro.Fields("NumPartidaRegistral"))
                dtpFechaUltimaTasacion.Value = adoRegistro.Fields("FechaUltimaTasacion")
                txtValorReferencial.Text = Trim(adoRegistro.Fields("ValorNominal"))
                
                adoRegistro.MoveNext
            Wend
        End If
        
    End With
    
    tabParticipeActivo.TabEnabled(0) = False
    tabParticipeActivo.TabVisible(1) = True
    tabParticipeActivo.Tab = 1
    cboTipoActivo.Enabled = False
    
End Sub

Private Sub Eliminar()
    
    Dim strMensaje  As String
    Dim Accion As String
    
    Accion = "D"
    
    If tdgConsulta.SelBookmarks.Count <= 0 Then Exit Sub
    
    strMensaje = "Se procederá a eliminara el activo " & Trim(tdgConsulta.Columns(4).Value) & _
    " del participe " & Trim(tdgConsulta.Columns(7).Value) & _
    vbNewLine & vbNewLine & "¿ Seguro de continuar ?"
    
    If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
        With adoComm
        
            .CommandText = "{ call up_IVMantParticipeAporteActivo('" & tdgConsulta.Columns(0).Value & "','','" & _
                 tdgConsulta.Columns(1).Value & "','','','','','',0,'','" & Accion & "') }"
            
            adoConn.Execute .CommandText
            
            MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation
            Call Buscar
            
        End With
    End If
    
End Sub

Private Sub Grabar()
    
    Dim strCodDetalleFile As String, strCodAnalitica As String
    Dim adoRegistro As ADODB.Recordset
    Dim Accion As String
    
    If TodoOK() Then
        
        Me.MousePointer = vbHourglass
        
        If strEstado = Reg_Adicion Then
            Accion = "I"
        Else
            Accion = "U"
        End If
        
        With adoComm
            
            If Accion = "I" Then
                
                .CommandText = "{ call up_ACSelDatosParametro(21,'" & strTipoActivo & "') }"
                Set adoRegistro = .Execute
                
                If Not adoRegistro.EOF Then
                    strCodAnalitica = Format(CInt(adoRegistro("NumUltimo")) + 1, "00000000")
                End If
            
            Else
                
                strCodAnalitica = tdgConsulta.Columns(1).Value
                
            End If
            
            If Not cboSubTipoActivo.Visible Then
                strCodDetalleFile = Valor_Caracter
            Else
                strCodDetalleFile = strSubTipoActivo
            End If
            
            On Error GoTo Ctrl_Error
                        
            
            .CommandText = "{ call up_IVMantParticipeAporteActivo('" & strTipoActivo & "','" & _
                 strCodDetalleFile & "','" & strCodAnalitica & "','" & Trim(txtCodCuenta.Text) & "', '" & _
                 Convertyyyymmdd(gdatFechaActual) & "','" & strCodParticipe & "','" & _
                 Trim(txtDescripActivo.Text) & "','" & Trim(txtUbicacionActivo.Text) & "','" & _
                 Trim(txtNumPartidaRegistral.Text) & "','" & strCodMoneda & "', " & CDbl(txtValorReferencial.Text) & ",'" & _
                 Convertyyyymmdd(dtpFechaUltimaTasacion.Value) & "','" & Accion & "') }"
            
            
            
            
            adoConn.Execute .CommandText
             
        End With

        Me.MousePointer = vbDefault
        
        MsgBox "Se guardaron los cambios exitosamente", vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        Call Cancelar
        
    End If
    
Exit Sub
    
Ctrl_Error:
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    MsgBox err.Description
    Me.MousePointer = vbDefault
    

End Sub

Private Sub Cancelar()
    If cboTipoActivo.ListCount > 0 Then cboTipoActivo.ListIndex = 0
    
    lblDescripParticipe.Caption = Valor_Caracter
    strCodParticipe = Valor_Caracter
    txtDescripActivo.Text = Valor_Caracter
    txtUbicacionActivo.Text = Valor_Caracter
    txtNumPartidaRegistral.Text = Valor_Caracter
    dtpFechaUltimaTasacion.Value = gdatFechaActual
    txtValorReferencial.Text = Valor_Caracter
    
    lblDescrip(10).Visible = True
    cboSubTipoActivo.Visible = True
    
    Call Buscar
    
    tabParticipeActivo.TabEnabled(0) = True
    tabParticipeActivo.TabVisible(1) = False
    tabParticipeActivo.Tab = 0
    cboTipoActivo.Enabled = True
End Sub

Private Sub Salir()
    Unload Me
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
            End If
        End If
    End If
    
End Sub

Private Function TodoOK() As Boolean
    TodoOK = False
    
    If strCodParticipe = Valor_Caracter Then
        MsgBox "Debe seleccionar el participe propietario del activo", vbCritical, Me.Caption
        Exit Function
    End If
    
    If cboTipoActivo.ListIndex <= 0 Then
        MsgBox "Debe seleccionar el tipo de activo", vbCritical, Me.Caption
        Exit Function
    End If
    
    If Trim(txtDescripActivo.Text) = Valor_Caracter Then
        MsgBox "La descripcion del activo no puede estar vacio", vbCritical, Me.Caption
        Exit Function
    End If
    
    If Trim(txtUbicacionActivo.Text) = Valor_Caracter Then
        MsgBox "La ubicacion del activo no puede estar vacio", vbCritical, Me.Caption
        Exit Function
    End If
    
    If Trim(txtNumPartidaRegistral.Text) = Valor_Caracter Then
        MsgBox "El numero de partida del activo no puede estar vacio", vbCritical, Me.Caption
        Exit Function
    End If
    
    If Trim(txtValorReferencial.Text) = Valor_Caracter Then
        MsgBox "El valor referencial del activo no puede estar vacio", vbCritical, Me.Caption
        Exit Function
    End If
    
    If CDbl(txtValorReferencial.Text) <= 0 Then
        MsgBox "El valor referencial del activo no puede ser 0", vbCritical, Me.Caption
        Exit Function
    End If
    
    If cboSubTipoActivo.Visible Then
        If cboSubTipoActivo.ListIndex <= 0 Then
            MsgBox "Debe seleccionar el subtipo del activo", vbCritical, Me.Caption
            Exit Function
        End If
    End If
    
    TodoOK = True
End Function



Private Sub cmdBusquedaParticipe_Click()
    gstrFormulario = "frmParticipeActivoAporte2"
    frmBusquedaParticipeP.Show vbModal
    If gstrCodParticipe <> Valor_Caracter Then strCodParticipe = gstrCodParticipe
End Sub

Private Sub chkFiltrarParticipe_Click()
    
    If chkFiltrarParticipe.Value Then
        'cmdBusqueda.Enabled = True
    Else
        'cmdBusqueda.Enabled = False
        lblDescripParticipeBusqueda.Caption = Valor_Caracter
        strCodParticipeBusqueda = Valor_Caracter
    End If
    
End Sub

Private Sub cboTipoActivo_Click()
    strTipoActivo = Valor_Caracter
    If cboTipoActivo.ListIndex < 0 Then Exit Sub
    strTipoActivo = Trim(arrTipoActivo(cboTipoActivo.ListIndex))
    
    strSql = "SELECT CodDetalleFile CODIGO, DescripDetalleFile DESCRIP FROM InversionDetalleFile " & _
    "WHERE CodFile='" & strTipoActivo & "'"
    CargarControlLista strSql, cboSubTipoActivo, arrSubTipoActivo(), Sel_Defecto
       
    If cboSubTipoActivo.ListCount > 1 Then
        cboSubTipoActivo.ListIndex = 0
        lblDescrip(10).Visible = True
        cboSubTipoActivo.Visible = True
    Else
        lblDescrip(10).Visible = False
        cboSubTipoActivo.Visible = False
    End If
End Sub

Private Sub cboSubTipoActivo_Click()
    strSubTipoActivo = Valor_Caracter
    If cboSubTipoActivo.ListIndex < 0 Then Exit Sub
    strSubTipoActivo = Trim(arrSubTipoActivo(cboSubTipoActivo.ListIndex))
End Sub

Private Sub txtNumPartidaRegistral_KeyPress(KeyAscii As Integer)
    Call ValidaCajaTexto(KeyAscii, "N", txtNumPartidaRegistral, 0)
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

Private Sub SubImprimir()

    Dim strSeleccionRegistro    As String
    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
       
            gstrNameRepo = "ParticipeActivoAporteGrilla"
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(2)
            ReDim aReportParamFn(2)
            ReDim aReportParamF(2)
                        
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
                        
            aReportParamS(0) = Convertyyyymmdd(dtpFechaDesde.Value)
            aReportParamS(1) = Convertyyyymmdd(dtpFechaHasta.Value)
            
            If strCodParticipeBusqueda <> Valor_Caracter Then
                aReportParamS(2) = Trim(strCodParticipeBusqueda)
            Else
                aReportParamS(2) = "%"
            End If
    
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal

End Sub

