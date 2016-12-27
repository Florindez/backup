VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmProyectoInmobiliarioDefinicion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definicion - Proyecto Inmobiliario"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   12060
   Begin TabDlg.SSTab tabProyectoInmobiliario 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   14420
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Listado"
      TabPicture(0)   =   "frmFondoInmobiliarioDefinicion.frx":0000
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
      TabCaption(1)   =   "Datos"
      TabPicture(1)   =   "frmFondoInmobiliarioDefinicion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmProyectoDatos"
      Tab(1).Control(1)=   "frmProyectoPlanificacion"
      Tab(1).Control(2)=   "cmdAccion"
      Tab(1).ControlCount=   3
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -66480
         TabIndex        =   44
         Top             =   7080
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
      Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
         Height          =   735
         Left            =   240
         TabIndex        =   43
         Top             =   7200
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
      Begin TAMControls2.ucBotonEdicion2 cmdSalir 
         Height          =   735
         Left            =   10200
         TabIndex        =   42
         Top             =   7200
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1296
         Caption0        =   "&Salir"
         Tag0            =   "9"
         Visible0        =   0   'False
         ToolTipText0    =   "Salir"
         UserControlWidth=   1200
      End
      Begin VB.Frame frmProyectoPlanificacion 
         Caption         =   "Planificacion"
         Height          =   3375
         Left            =   -74760
         TabIndex        =   25
         Top             =   3360
         Width           =   11295
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
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
            Left            =   1920
            Picture         =   "frmFondoInmobiliarioDefinicion.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   2400
            Width           =   1200
         End
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "&Quitar"
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
            Left            =   9960
            Picture         =   "frmFondoInmobiliarioDefinicion.frx":059A
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   2520
            Width           =   1200
         End
         Begin VB.CommandButton cmdAgregarEditar 
            Caption         =   "&Agregar"
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
            Left            =   600
            Picture         =   "frmFondoInmobiliarioDefinicion.frx":07BC
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   2400
            Width           =   1200
         End
         Begin VB.ComboBox cboTipoUnidadInmobiliaria 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   480
            Width           =   2145
         End
         Begin TAMControls.TAMTextBox txtMedidaUnidadInmobiliaria 
            Height          =   315
            Left            =   2520
            TabIndex        =   27
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Container       =   "frmFondoInmobiliarioDefinicion.frx":0CDC
            Estilo          =   3
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   2000000000
         End
         Begin TAMControls.TAMTextBox txtCantidadUnidad 
            Height          =   315
            Left            =   2520
            TabIndex        =   32
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
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
            MaxLength       =   8
            Container       =   "frmFondoInmobiliarioDefinicion.frx":0CF8
            Estilo          =   3
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   999999
         End
         Begin TrueOleDBGrid60.TDBGrid tdgProyectoPlanificacion 
            Bindings        =   "frmFondoInmobiliarioDefinicion.frx":0D14
            Height          =   2055
            Left            =   4800
            OleObjectBlob   =   "frmFondoInmobiliarioDefinicion.frx":0D2E
            TabIndex        =   33
            Top             =   240
            Width           =   6345
         End
         Begin VB.Label lblTotalMedida 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2520
            TabIndex        =   35
            Top             =   1920
            Width           =   1665
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Total"
            Height          =   195
            Index           =   13
            Left            =   600
            TabIndex        =   34
            Top             =   1920
            Width           =   1545
         End
         Begin VB.Label lblSimboloMedida 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4080
            TabIndex        =   31
            Top             =   960
            Width           =   585
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Cantidad"
            Height          =   195
            Index           =   12
            Left            =   600
            TabIndex        =   30
            Top             =   1440
            Width           =   1785
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Unidad Inmobiliaria"
            Height          =   195
            Index           =   10
            Left            =   600
            TabIndex        =   29
            Top             =   480
            Width           =   1785
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Medida x Und."
            Height          =   195
            Index           =   11
            Left            =   600
            TabIndex        =   28
            Top             =   960
            Width           =   1785
         End
      End
      Begin VB.Frame frmProyectoDatos 
         Caption         =   "Datos"
         Height          =   2775
         Left            =   -74760
         TabIndex        =   11
         Top             =   480
         Width           =   11295
         Begin VB.ComboBox cboTipoUnidadMedidaSuperficie 
            Height          =   315
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1800
            Width           =   2025
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   315
            Left            =   2640
            MaxLength       =   200
            TabIndex        =   17
            Top             =   1320
            Width           =   5175
         End
         Begin VB.ComboBox cboTipoProyecto 
            Height          =   315
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   840
            Width           =   3345
         End
         Begin TAMControls.TAMTextBox txtMontoTotalGeneral 
            Height          =   315
            Left            =   2640
            TabIndex        =   22
            Top             =   2280
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
            Container       =   "frmFondoInmobiliarioDefinicion.frx":4DD3
            Estilo          =   3
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   2000000000
         End
         Begin TAMControls.TAMTextBox txtMontoTotalProyectado 
            Height          =   315
            Left            =   7080
            TabIndex        =   24
            Top             =   2280
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
            Container       =   "frmFondoInmobiliarioDefinicion.frx":4DEF
            Estilo          =   3
            ColorEnfoque    =   8454143
            Borde           =   1
            MaximoValor     =   2000000000
         End
         Begin VB.Label lblFondo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2640
            TabIndex        =   41
            Top             =   360
            Width           =   7785
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fondo"
            Height          =   195
            Index           =   14
            Left            =   720
            TabIndex        =   40
            Top             =   360
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Monto Proyectado"
            Height          =   195
            Index           =   9
            Left            =   5160
            TabIndex        =   23
            Top             =   2280
            Width           =   1785
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Monto Total "
            Height          =   195
            Index           =   8
            Left            =   720
            TabIndex        =   21
            Top             =   2280
            Width           =   1785
         End
         Begin VB.Label lblSimboloUnidadMedidaSuperficie 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4920
            TabIndex        =   20
            Top             =   1800
            Width           =   585
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Unidad Superficie"
            Height          =   195
            Index           =   7
            Left            =   720
            TabIndex        =   18
            Top             =   1800
            Width           =   1785
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Descripcion"
            Height          =   195
            Index           =   6
            Left            =   720
            TabIndex        =   16
            Top             =   1320
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo"
            Height          =   195
            Index           =   5
            Left            =   720
            TabIndex        =   14
            Top             =   840
            Width           =   1305
         End
         Begin VB.Label lblCodTitulo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   8760
            TabIndex        =   13
            Top             =   840
            Width           =   1665
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Codigo Titulo"
            Height          =   195
            Index           =   4
            Left            =   6840
            TabIndex        =   12
            Top             =   840
            Width           =   1305
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
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   11535
         Begin VB.CheckBox chkFiltrarFechas 
            Caption         =   "Filtrar"
            Height          =   255
            Left            =   7080
            TabIndex        =   39
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   360
            Width           =   5145
         End
         Begin VB.ComboBox cboTipoProyectoInmobiliario 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   780
            Width           =   5145
         End
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   285
            Left            =   9120
            TabIndex        =   5
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
            Left            =   9120
            TabIndex        =   6
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
            Caption         =   "Desde"
            Height          =   195
            Index           =   2
            Left            =   8160
            TabIndex        =   10
            Top             =   360
            Width           =   705
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   3
            Left            =   8160
            TabIndex        =   9
            Top             =   840
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
            Caption         =   "Tipo"
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   7
            Top             =   780
            Width           =   825
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmFondoInmobiliarioDefinicion.frx":4E0B
         Height          =   4815
         Left            =   120
         OleObjectBlob   =   "frmFondoInmobiliarioDefinicion.frx":4E25
         TabIndex        =   1
         Top             =   2160
         Width           =   11505
      End
   End
End
Attribute VB_Name = "frmProyectoInmobiliarioDefinicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSQL As String, strEstado As String, strTipoProyecto As String
Dim strTipoUnidadMedidaSuperficie As String, strTipoUnidadInmobiliaria As String
Dim strCodFondo As String, strTipoProyectoInmobiliario As String
Dim arrFondo() As String, arrTipoProyectoInmobiliario() As String, arrTipoProyecto() As String
Dim arrTipoUnidadMedidaSuperficie() As String, arrTipoUnidadInmobiliaria() As String
Dim adoConsulta As ADODB.Recordset, adoRegistroAux As ADODB.Recordset
Dim strTipoUnidadMedidaSuperficieEnUso As String, strDescripUnidadMedidaSuperficieEnUso As String

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

Private Sub InicializarValores()
    
    dtpFechaDesde.Value = gdatFechaActual
    dtpFechaHasta.Value = gdatFechaActual
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    
    cmdCancelar.Visible = False
    txtMontoTotalProyectado.Enabled = False
    tabProyectoInmobiliario.TabEnabled(0) = True
    tabProyectoInmobiliario.TabVisible(1) = False
    tabProyectoInmobiliario.Tab = 0
    
    strTipoUnidadMedidaSuperficieEnUso = Valor_Caracter
    strDescripUnidadMedidaSuperficieEnUso = Valor_Caracter
    
    Call chkFiltrarFechas_Click
    
End Sub

Private Sub CargarListas()
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
       
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    strSQL = "SELECT CodFile CODIGO,DescripFile DESCRIP " & _
            "FROM InversionFile WHERE CodFile IN ('030','040') AND IndVigente='X'"
    CargarControlLista strSQL, cboTipoProyectoInmobiliario, arrTipoProyectoInmobiliario(), Sel_Todos
    CargarControlLista strSQL, cboTipoProyecto, arrTipoProyecto(), Sel_Defecto
    
    If cboTipoProyectoInmobiliario.ListCount > 0 Then cboTipoProyectoInmobiliario.ListIndex = 0
    If cboTipoProyecto.ListCount > 0 Then cboTipoProyecto.ListIndex = 0
    
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP " & _
            "FROM AuxiliarParametro " & _
            "WHERE CodTipoParametro='UNDSUP' AND Estado='01' " & _
            "ORDER BY 1"
    
    CargarControlLista strSQL, cboTipoUnidadMedidaSuperficie, arrTipoUnidadMedidaSuperficie(), Sel_Defecto
    
    If cboTipoUnidadMedidaSuperficie.ListCount > 0 Then cboTipoUnidadMedidaSuperficie.ListIndex = 0
    
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP " & _
            "FROM AuxiliarParametro " & _
            "WHERE CodTipoParametro='UNDINM' AND Estado='01' " & _
            "ORDER BY 1"
    
    CargarControlLista strSQL, cboTipoUnidadInmobiliaria, arrTipoUnidadInmobiliaria(), Sel_Defecto
    
    If cboTipoUnidadInmobiliaria.ListCount > 0 Then cboTipoUnidadInmobiliaria.ListIndex = 0
    
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
        Case vPrint
            Call SubImprimir
        Case vExit
            Call Salir
        
    End Select
    
End Sub

Private Sub Adicionar()
        
    strEstado = Reg_Adicion
        
    Call ConfiguraRecordsetAuxiliar
    
    tdgProyectoPlanificacion.DataSource = adoRegistroAux
    
    lblFondo.Caption = cboFondo.Text
    
    cboTipoProyecto.Enabled = True
    txtMontoTotalProyectado.Text = "0"
    lblDescrip(4).Visible = False
    lblCodTitulo.Visible = False
    frmProyectoPlanificacion.Enabled = False
    tabProyectoInmobiliario.TabEnabled(0) = False
    tabProyectoInmobiliario.TabVisible(1) = True
    tabProyectoInmobiliario.Tab = 1
    
    If adoRegistroAux.RecordCount = 0 Then cmdQuitar.Enabled = False
    
End Sub

Private Sub Modificar()
    
    Dim adoRegistro As ADODB.Recordset, intRegistro As Integer
    
    strEstado = Reg_Edicion
        
    Set adoRegistro = New ADODB.Recordset
    
    cboTipoProyecto.Enabled = False
    frmProyectoPlanificacion.Enabled = False
    
    If tdgConsulta.SelBookmarks.Count <= 0 Then Exit Sub
    
    With adoComm
        
        .CommandText = "{ call up_IVObtenerDatosProyectoInmobiliario('" & strCodFondo & "','" & _
                 gstrCodAdministradora & "','" & Trim(tdgConsulta.Columns(0).Value) & "') }"
                 
        Set adoRegistro = .Execute
        
        lblFondo.Caption = cboFondo.Text
        
        If Not adoRegistro.EOF Then
            While Not adoRegistro.EOF
                lblCodTitulo.Caption = Trim(adoRegistro.Fields("CodTitulo"))
                intRegistro = ObtenerItemLista(arrTipoProyecto(), Trim(adoRegistro.Fields("CodFile")))
                If intRegistro >= 0 Then cboTipoProyecto.ListIndex = intRegistro
                txtDescripcion.Text = Trim(adoRegistro.Fields("DescripProyecto"))
                intRegistro = ObtenerItemLista(arrTipoUnidadMedidaSuperficie(), Trim(adoRegistro.Fields("TipoUnidadMedida")))
                If intRegistro >= 0 Then cboTipoUnidadMedidaSuperficie.ListIndex = intRegistro
                Call cboTipoUnidadMedidaSuperficie_Click
                txtMontoTotalGeneral.Text = Trim(adoRegistro.Fields("MontoTotalGeneral"))
                txtMontoTotalProyectado.Text = Trim(adoRegistro.Fields("MontoTotalProyectado"))
                strTipoUnidadMedidaSuperficieEnUso = Trim(adoRegistro.Fields("TipoUnidadMedida"))
                strDescripUnidadMedidaSuperficieEnUso = Trim(adoRegistro.Fields("DescripUnidadMedida"))
                adoRegistro.MoveNext
            Wend
        End If
            
        Call ConfiguraRecordsetAuxiliar
        
       .CommandText = "{ call up_IVObtenerPlanificacionProyectoInmobiliario('" & strCodFondo & "','" & _
                 gstrCodAdministradora & "','" & Trim(tdgConsulta.Columns(0).Value) & "') }"
        
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            While Not adoRegistro.EOF
                adoRegistroAux.AddNew
                For Each adoField In adoRegistroAux.Fields
                    adoRegistroAux.Fields(adoField.Name) = adoRegistro.Fields(adoField.Name)
                Next
                adoRegistroAux.Update
                adoRegistro.MoveNext
            Wend
        End If
        
        tdgProyectoPlanificacion.DataSource = adoRegistroAux
        
    End With
    
    Call txtMontoTotalGeneral_Change
    
    lblDescrip(4).Visible = True
    lblCodTitulo.Visible = True
    tabProyectoInmobiliario.TabEnabled(0) = False
    tabProyectoInmobiliario.TabVisible(1) = True
    tabProyectoInmobiliario.Tab = 1
    
End Sub

Private Sub Eliminar()
    Dim strMensaje  As String
    Dim Accion As String
    
    Accion = "D"
    
    If tdgConsulta.SelBookmarks.Count <= 0 Then Exit Sub
    
    strMensaje = "Se procederá a eliminara el proyecto inmobiliario " & Trim(tdgConsulta.Columns(3).Value) & _
    vbNewLine & vbNewLine & vbNewLine & "¿ Seguro de continuar ?"
    
    If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
        With adoComm
        
            .CommandText = "{ call up_IVMantProyectoInmboliario('" & strCodFondo & "','" & _
                gstrCodAdministradora & "','" & Trim(tdgConsulta.Columns(0).Value) & "','" & strTipoProyecto & "','" & _
                Valor_Caracter & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
                Valor_Caracter & "','" & strTipoUnidadMedidaSuperficie & "',0,0,'','" & Accion & "') }"
            
            adoConn.Execute .CommandText
            
            MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation
            Call Buscar
            
        End With
    End If
    
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

Private Sub Grabar()
    
    Dim strCodTituloProyecto As String, strCodAnaliticaProyecto As String
    Dim objProyectoDetallePlanificacionXML  As DOMDocument60
    Dim strMsgError                 As String
    Dim strProyectoDetallePlanificacionXML As String
    Dim adoRegistro As ADODB.Recordset
    Dim Accion As String
    
    Set adoRegistro = New ADODB.Recordset
    
    If strEstado = Reg_Adicion Then
        Accion = "I"
        strCodTituloProyecto = Valor_Caracter
    Else
        Accion = "U"
        strCodTituloProyecto = Trim(lblCodTitulo.Caption)
    End If
    
    If TodoOK() Then
        
        Me.MousePointer = vbHourglass
        
        Call XMLADORecordset(objProyectoDetallePlanificacionXML, "ProyectoInmobiliario", "Planificacion", adoRegistroAux, strMsgError)
        strProyectoDetallePlanificacionXML = objProyectoDetallePlanificacionXML.xml
        
        With adoComm
            
            On Error GoTo Ctrl_Error
            
            If Accion = "I" Then
                .CommandText = "{call up_ACSelDatosParametro(21,'" & strTipoProyecto & "') }"
                Set adoRegistro = .Execute
                    
                If Not adoRegistro.EOF Then
                    strCodAnaliticaProyecto = Format(CInt(adoRegistro("NumUltimo")) + 1, "00000000")
                End If
            Else
                strCodAnaliticaProyecto = Valor_Caracter
            End If
            
            .CommandText = "{ call up_IVMantProyectoInmboliario('" & strCodFondo & "','" & _
                 gstrCodAdministradora & "','" & strCodTituloProyecto & "','" & strTipoProyecto & "','" & _
                 strCodAnaliticaProyecto & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
                 Trim(txtDescripcion.Text) & "','" & strTipoUnidadMedidaSuperficie & "'," & _
                 CLng(txtMontoTotalGeneral.Text) & "," & CLng(txtMontoTotalProyectado.Text) & ",'" & _
                 strProyectoDetallePlanificacionXML & "','" & Accion & "') }"
            
            adoConn.Execute .CommandText
             
        End With
        
        Set adoRegistroAux = Nothing
                
        Me.MousePointer = vbDefault
        
        If Accion = "I" Then
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
        Else
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
        End If
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        Call Cancelar
        
    End If
    
    Exit Sub
    
Ctrl_Error:
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault

End Sub

Private Sub Cancelar()
    
    lblDescrip(4).Visible = True
    lblCodTitulo.Visible = True
    cboTipoUnidadInmobiliaria.Enabled = True
    txtMedidaUnidadInmobiliaria.Enabled = True
    cmdCancelar.Visible = False
    txtMontoTotalProyectado.Enabled = False
    lblCodTitulo.Caption = Valor_Caracter
    strTipoUnidadMedidaSuperficieEnUso = Valor_Caracter
    strDescripUnidadMedidaSuperficieEnUso = Valor_Caracter
    cboTipoProyecto.Enabled = True
    frmProyectoDatos.Enabled = True
    
    Call Limpiar
    Call LimpiarPlanificacion
    Call Buscar
    
    tabProyectoInmobiliario.TabEnabled(0) = True
    tabProyectoInmobiliario.TabVisible(1) = False
    tabProyectoInmobiliario.Tab = 0
    
End Sub

Private Sub Salir()
    Unload Me
End Sub

Private Sub AutoAjustarGrillas()

    Dim i As Integer

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

End Sub

Private Sub Form_Resize()
    Call AutoAjustarGrillas
End Sub

Private Sub ConfiguraRecordsetAuxiliar()

    Set adoRegistroAux = New ADODB.Recordset

    With adoRegistroAux
       .CursorLocation = adUseClient
       .Fields.Append "TipoUnidadInmobiliaria", adVarChar, 500
       .Fields.Append "DescripUnidadInmobiliaria", adVarChar, 500
       .Fields.Append "MedidaUnidadInmobiliaria", adDecimal, 19
       .Fields.Append "MedidaTexto", adVarChar, 500
       .Fields.Append "CantUnidadInmobiliaria", adDecimal, 19
       .Fields.Append "Total", adDecimal, 19
'       .CursorType = adOpenStatic
       .LockType = adLockBatchOptimistic
    End With

    adoRegistroAux.Open
    
End Sub

Private Sub Limpiar()
    If cboTipoProyecto.ListCount > 0 Then cboTipoProyecto.ListIndex = 0
    lblCodTitulo.Caption = Valor_Caracter
    txtDescripcion.Text = Valor_Caracter
    If cboTipoUnidadMedidaSuperficie.ListCount > 0 Then cboTipoUnidadMedidaSuperficie.ListIndex = 0
    lblSimboloUnidadMedidaSuperficie.Caption = Valor_Caracter
    txtMontoTotalGeneral.Text = Valor_Caracter
    txtMontoTotalProyectado.Text = Valor_Caracter
    If cboTipoUnidadInmobiliaria.ListCount > 0 Then cboTipoUnidadInmobiliaria.ListIndex = 0
    txtMedidaUnidadInmobiliaria.Text = Valor_Caracter
    lblSimboloMedida.Caption = Valor_Caracter
    txtCantidadUnidad.Text = Valor_Caracter
    lblTotalMedida.Caption = Valor_Caracter
End Sub

Private Sub LimpiarPlanificacion()
    If cboTipoUnidadInmobiliaria.ListCount > 0 Then cboTipoUnidadInmobiliaria.ListIndex = 0
    txtMedidaUnidadInmobiliaria.Text = Valor_Caracter
    txtCantidadUnidad.Text = Valor_Caracter
    lblTotalMedida.Caption = Valor_Caracter
End Sub

Private Function TotalProyectoPlanificacion() As Long
    Dim intTotal As Long
    Dim adoRegistroClone As ADODB.Recordset
    If adoRegistroAux.RecordCount > 0 Then
        Set adoRegistroClone = adoRegistroAux.Clone
        adoRegistroClone.MoveFirst
        While Not adoRegistroClone.EOF
            intTotal = intTotal + (CLng(adoRegistroClone.Fields("MedidaUnidadInmobiliaria")) * CLng(adoRegistroClone.Fields("CantUnidadInmobiliaria")))
            adoRegistroClone.MoveNext
        Wend
    Else
        intTotal = 0
    End If
    TotalProyectoPlanificacion = intTotal
End Function

Private Function TodoOK() As Boolean
    
    TodoOK = False
    
    If cboTipoProyecto.ListIndex = 0 Then
        MsgBox "Debe seleccionar el tipo de proyecto", vbCritical, Me.Caption
        If cboTipoProyecto.Enabled Then cboTipoProyecto.SetFocus
        Exit Function
    End If
    
    If Trim(txtDescripcion.Text) = Valor_Caracter Then
        MsgBox "El campo descripcion no puede estar en blanco", vbCritical, Me.Caption
        If txtDescripcion.Enabled Then txtDescripcion.SetFocus
        Exit Function
    End If
    
    If cboTipoUnidadMedidaSuperficie.ListIndex = 0 Then
        MsgBox "Debe seleccionar el tipo de unidad de superficie", vbCritical, Me.Caption
        If cboTipoUnidadMedidaSuperficie.Enabled Then cboTipoUnidadMedidaSuperficie.SetFocus
        Exit Function
    End If
    
    If Trim(txtMontoTotalGeneral.Text) = Valor_Caracter Or CLng(IIf(Trim(txtMontoTotalGeneral.Text) = Valor_Caracter, 0, txtMontoTotalGeneral.Text)) = 0 Then
        MsgBox "El campo Monto total no puede estar en blanco o ser 0", vbCritical, Me.Caption
        If txtMontoTotalGeneral.Enabled Then txtMontoTotalGeneral.SetFocus
        Exit Function
    End If
    
    If adoRegistroAux.RecordCount < 1 Then
        MsgBox "La planificacion del proyecto debe tener al menos un item en la lista", vbCritical, Me.Caption
        Exit Function
    End If
    
    If strTipoUnidadMedidaSuperficieEnUso <> Trim(arrTipoUnidadMedidaSuperficie(cboTipoUnidadMedidaSuperficie.ListIndex)) Then
        MsgBox "Se ingresaron unidades inmobiliarias medidas en " & _
        strDescripUnidadMedidaSuperficieEnUso & ",no se puede guardar el proyecto con la unidad seleccionada", vbCritical, Me.Caption
        If cboTipoUnidadMedidaSuperficie.Enabled Then cboTipoUnidadMedidaSuperficie.SetFocus
        Exit Function
    End If
    
    TodoOK = True
    
End Function

Private Function TodoOkPlanificacion() As Boolean
    TodoOkPlanificacion = False
    
    If strTipoUnidadMedidaSuperficieEnUso <> Valor_Caracter Then
        If cmdAgregarEditar.Caption = "&Agregar" Then
            If strTipoUnidadMedidaSuperficieEnUso <> Trim(arrTipoUnidadMedidaSuperficie(cboTipoUnidadMedidaSuperficie.ListIndex)) Then
                MsgBox "Se ingresaron unidades inmobiliarias medidas en " & _
                strDescripUnidadMedidaSuperficieEnUso & ",no se puede ingresar una unidad con una medida diferente", vbCritical, Me.Caption
                Exit Function
            End If
        End If
    End If
    
    If cboTipoUnidadInmobiliaria.ListIndex = 0 Then
        MsgBox "Debe seleccionar el tipo de unidad inmobiliaria", vbCritical, Me.Caption
        Exit Function
    End If
    
    If CLng(IIf(Trim(txtMedidaUnidadInmobiliaria.Text) = Valor_Caracter, 0, txtMedidaUnidadInmobiliaria.Text)) = 0 Then
        MsgBox "La medida de la unidad no puede ser 0 o estar vacia", vbCritical, Me.Caption
        Exit Function
    End If
    
    If CLng(IIf(Trim(txtCantidadUnidad.Text) = Valor_Caracter, 0, txtCantidadUnidad.Text)) = 0 Then
        MsgBox "La cantidad de unidades no puede ser 0 o estar vacia", vbCritical, Me.Caption
        Exit Function
    End If
    
    TodoOkPlanificacion = True
End Function

Private Function AgregaDetalle(ByVal strTipoUnidadInmobiliaria As String, ByVal intMedidaUnidad As Long, ByVal intCantidadUnidad As Long) As Boolean
    
    Dim intCantidad As Long
    Dim total As Long
    AgregaDetalle = False
        
    If adoRegistroAux.RecordCount > 0 Then
        adoRegistroAux.MoveFirst
        While Not adoRegistroAux.EOF
            If Trim(adoRegistroAux.Fields("TipoUnidadInmobiliaria")) = strTipoUnidadInmobiliaria And _
            CLng(adoRegistroAux.Fields("MedidaUnidadInmobiliaria")) = intMedidaUnidad Then
                adoRegistroAux.Fields("TipoUnidadInmobiliaria") = strTipoUnidadInmobiliaria
                adoRegistroAux.Fields("MedidaUnidadInmobiliaria") = intMedidaUnidad
                intCantidad = CLng(adoRegistroAux.Fields("CantUnidadInmobiliaria"))
                adoRegistroAux.Fields("CantUnidadInmobiliaria") = intCantidad + intCantidadUnidad
                total = (intCantidad + intCantidadUnidad) * intMedidaUnidad
                adoRegistroAux.Fields("Total") = total
                AgregaDetalle = True
                Exit Function
            End If
            adoRegistroAux.MoveNext
        Wend
    End If
    
End Function

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

Private Sub cboTipoProyecto_Click()
    strTipoProyecto = Valor_Caracter
    If cboTipoProyecto.ListIndex < 0 Then Exit Sub
    strTipoProyecto = Trim(arrTipoProyecto(cboTipoProyecto.ListIndex))
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

Private Sub tdgProyectoPlanificacion_DblClick()
    
    Dim strTMP As String
    Dim intRegistro As Integer
    
    If tdgProyectoPlanificacion.SelBookmarks.Count > 0 Then
        intRegistro = ObtenerItemLista(arrTipoUnidadInmobiliaria(), adoRegistroAux.Fields("TipoUnidadInmobiliaria"))
        If intRegistro >= 0 Then cboTipoUnidadInmobiliaria.ListIndex = intRegistro
        txtMedidaUnidadInmobiliaria.Text = adoRegistroAux.Fields("MedidaUnidadInmobiliaria")
        txtCantidadUnidad.Text = adoRegistroAux.Fields("CantUnidadInmobiliaria")
        lblTotalMedida.Caption = adoRegistroAux.Fields("Total")
        strTMP = CStr(adoRegistroAux.Fields("MedidaTexto"))
        lblSimboloMedida.Caption = Mid(strTMP, InStr(strTMP, " ") + 1, Len(strTMP))
        cmdAgregarEditar.Caption = "&Editar"
        cboTipoUnidadInmobiliaria.Enabled = False
        txtMedidaUnidadInmobiliaria.Enabled = False
        frmProyectoDatos.Enabled = False
        cmdCancelar.Visible = True
    End If
    
End Sub

Private Sub cboTipoUnidadMedidaSuperficie_Click()
    
    Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
    
    strTipoUnidadMedidaSuperficie = Valor_Caracter
    lblSimboloUnidadMedidaSuperficie.Caption = Valor_Caracter
    lblSimboloMedida.Caption = Valor_Caracter
    If cboTipoUnidadMedidaSuperficie.ListIndex < 0 Then Exit Sub
    strTipoUnidadMedidaSuperficie = Trim(arrTipoUnidadMedidaSuperficie(cboTipoUnidadMedidaSuperficie.ListIndex))
    
    With adoComm
        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro " & _
            "WHERE CodTipoParametro='UNDSUP' AND CodParametro='" & strTipoUnidadMedidaSuperficie & "' AND Estado='01'"
        Set adoRegistro = .Execute
    End With
    
    If Not adoRegistro.EOF Then
        While Not adoRegistro.EOF
            lblSimboloUnidadMedidaSuperficie.Caption = Trim(adoRegistro("ValorParametro").Value)
            lblSimboloMedida.Caption = Trim(adoRegistro("ValorParametro").Value)
            adoRegistro.MoveNext
        Wend
    End If
    
    If cboTipoUnidadMedidaSuperficie.ListIndex <> 0 Then
        If CLng(IIf(Trim(txtMontoTotalGeneral.Text) = Valor_Caracter, 0, txtMontoTotalGeneral.Text)) <> 0 And _
        CLng(IIf(Trim(txtMontoTotalGeneral.Text) = Valor_Caracter, 0, txtMontoTotalGeneral.Text)) >= CLng(IIf(Trim(txtMontoTotalProyectado.Text) = Valor_Caracter, 0, txtMontoTotalProyectado.Text)) Then
            frmProyectoPlanificacion.Enabled = True
        Else
            frmProyectoPlanificacion.Enabled = False
        End If
    Else
        frmProyectoPlanificacion.Enabled = False
    End If
    
End Sub

Private Sub cboTipoUnidadInmobiliaria_Click()
    strTipoUnidadInmobiliaria = Valor_Caracter
    If cboTipoUnidadInmobiliaria.ListIndex < 0 Then Exit Sub
    strTipoUnidadInmobiliaria = Trim(arrTipoUnidadInmobiliaria(cboTipoUnidadInmobiliaria.ListIndex))
End Sub

Private Sub txtMontoTotalGeneral_Change()
    Dim tmp As Variant
    tmp = txtMontoTotalGeneral.Text
    If IIf(Trim(tmp) = Valor_Caracter, 0, CVar(tmp)) > CLng(txtMontoTotalGeneral.MaximoValor) Then
        MsgBox "Ha sobrepasado el valor maximo (" & txtMontoTotalGeneral.MaximoValor & ")", vbCritical, Me.Caption
        txtMontoTotalGeneral.Text = txtMontoTotalGeneral.MaximoValor
    End If
    
    If cboTipoUnidadMedidaSuperficie.ListIndex <> 0 Then
        If CLng(IIf(Trim(txtMontoTotalGeneral.Text) = Valor_Caracter, 0, txtMontoTotalGeneral.Text)) <> 0 And _
        CLng(IIf(Trim(txtMontoTotalGeneral.Text) = Valor_Caracter, 0, txtMontoTotalGeneral.Text)) >= CLng(IIf(Trim(txtMontoTotalProyectado.Text) = Valor_Caracter, 0, txtMontoTotalProyectado.Text)) Then
            frmProyectoPlanificacion.Enabled = True
        Else
            frmProyectoPlanificacion.Enabled = False
        End If
    Else
        frmProyectoPlanificacion.Enabled = False
    End If
End Sub

Private Sub txtCantidadUnidad_Change()
    Dim medida As Long, cantidad As Long
    Dim tmp As Variant
    tmp = txtCantidadUnidad.Text
    If IIf(Trim(tmp) = Valor_Caracter, 0, CVar(tmp)) > CLng(txtCantidadUnidad.MaximoValor) Then
        MsgBox "Ha sobrepasado el valor maximo (" & txtCantidadUnidad.MaximoValor & ")", vbCritical, Me.Caption
        txtCantidadUnidad.Text = txtCantidadUnidad.MaximoValor
    End If
        
    If CLng(IIf(Trim(txtMedidaUnidadInmobiliaria.Text) = Valor_Caracter, 0, txtMedidaUnidadInmobiliaria.Text)) = 0 Then
        medida = 0
    Else
        medida = CLng(IIf(Trim(txtMedidaUnidadInmobiliaria.Text) = Valor_Caracter, 0, txtMedidaUnidadInmobiliaria.Text))
    End If
    
    If CLng(IIf(Trim(txtCantidadUnidad.Text) = Valor_Caracter, 0, txtCantidadUnidad.Text)) = 0 Then
        cantidad = 0
    Else
        cantidad = CLng(IIf(Trim(txtCantidadUnidad.Text) = Valor_Caracter, 0, txtCantidadUnidad.Text))
    End If
    
    
    lblTotalMedida.Caption = medida * cantidad
End Sub

Private Sub txtMedidaUnidadInmobiliaria_Change()
    Dim medida As Long, cantidad As Long
    Dim tmp As Variant
    tmp = txtMedidaUnidadInmobiliaria.Text
    If IIf(Trim(tmp) = Valor_Caracter, 0, CVar(tmp)) > CLng(txtMedidaUnidadInmobiliaria.MaximoValor) Then
        MsgBox "Ha sobrepasado el valor maximo (" & txtMedidaUnidadInmobiliaria.MaximoValor & ")", vbCritical, Me.Caption
        txtMedidaUnidadInmobiliaria.Text = txtMedidaUnidadInmobiliaria.MaximoValor
    End If
    
    If CLng(IIf(Trim(txtMedidaUnidadInmobiliaria.Text) = Valor_Caracter, 0, txtMedidaUnidadInmobiliaria.Text)) = 0 Then
        medida = 0
    Else
        medida = CLng(IIf(Trim(txtMedidaUnidadInmobiliaria.Text) = Valor_Caracter, 0, txtMedidaUnidadInmobiliaria.Text))
    End If
    
    If CLng(IIf(Trim(txtCantidadUnidad.Text) = Valor_Caracter, 0, txtCantidadUnidad.Text)) = 0 Then
        cantidad = 0
    Else
        cantidad = CLng(IIf(Trim(txtCantidadUnidad.Text) = Valor_Caracter, 0, txtCantidadUnidad.Text))
    End If
    
    lblTotalMedida.Caption = medida * cantidad
End Sub

Private Sub cmdAgregarEditar_Click()
    Dim dblBookmark As Double
    Dim intTotalIngresado As Long, intTotalActual As Long, intTotalGeneral As Long
    
    If TodoOkPlanificacion() Then
        intTotalIngresado = CLng(txtMedidaUnidadInmobiliaria.Text) * CLng(txtCantidadUnidad.Text)
        intTotalActual = TotalProyectoPlanificacion() + intTotalIngresado
        intTotalGeneral = CLng(txtMontoTotalGeneral.Text)
        
        If intTotalActual <= intTotalGeneral Then
            If cmdAgregarEditar.Caption = "&Agregar" Then
                
                If strTipoUnidadMedidaSuperficieEnUso = Valor_Caracter Then
                    strTipoUnidadMedidaSuperficieEnUso = Trim(arrTipoUnidadMedidaSuperficie(cboTipoUnidadMedidaSuperficie.ListIndex))
                    strDescripUnidadMedidaSuperficieEnUso = Trim(cboTipoUnidadMedidaSuperficie.Text)
                End If
                            
                If Not AgregaDetalle(arrTipoUnidadInmobiliaria(cboTipoUnidadInmobiliaria.ListIndex), CLng(txtMedidaUnidadInmobiliaria.Text), CLng(Trim(txtCantidadUnidad.Text))) Then
                    adoRegistroAux.AddNew
                    adoRegistroAux.Fields("TipoUnidadInmobiliaria") = arrTipoUnidadInmobiliaria(cboTipoUnidadInmobiliaria.ListIndex)
                    adoRegistroAux.Fields("DescripUnidadInmobiliaria") = cboTipoUnidadInmobiliaria.Text
                    adoRegistroAux.Fields("MedidaUnidadInmobiliaria") = CLng(txtMedidaUnidadInmobiliaria.Text)
                    adoRegistroAux.Fields("MedidaTexto") = Trim(txtMedidaUnidadInmobiliaria.Text) & " " & Trim(lblSimboloMedida.Caption)
                    adoRegistroAux.Fields("CantUnidadInmobiliaria") = CLng(Trim(txtCantidadUnidad.Text))
                    adoRegistroAux.Fields("Total") = CLng(txtMedidaUnidadInmobiliaria.Text) * CLng(txtCantidadUnidad.Text)
                End If
                            
            Else
                
                adoRegistroAux.Fields("TipoUnidadInmobiliaria") = arrTipoUnidadInmobiliaria(cboTipoUnidadInmobiliaria.ListIndex)
                adoRegistroAux.Fields("DescripUnidadInmobiliaria") = cboTipoUnidadInmobiliaria.Text
                adoRegistroAux.Fields("MedidaUnidadInmobiliaria") = CLng(txtMedidaUnidadInmobiliaria.Text)
                adoRegistroAux.Fields("MedidaTexto") = Trim(txtMedidaUnidadInmobiliaria.Text) & " " & Trim(lblSimboloMedida.Caption)
                adoRegistroAux.Fields("CantUnidadInmobiliaria") = CLng(Trim(txtCantidadUnidad.Text))
                adoRegistroAux.Fields("Total") = CLng(txtMedidaUnidadInmobiliaria.Text) * CLng(txtCantidadUnidad.Text)
                               
                cmdAgregarEditar.Caption = "&Agregar"
                cboTipoUnidadInmobiliaria.Enabled = True
                txtMedidaUnidadInmobiliaria.Enabled = True
                cmdCancelar.Visible = False
                lblSimboloMedida.Caption = lblSimboloUnidadMedidaSuperficie.Caption
                
                frmProyectoDatos.Enabled = True
    
            End If
            cmdQuitar.Enabled = True
            adoRegistroAux.Update
            dblBookmark = adoRegistroAux.Bookmark
                
            tdgProyectoPlanificacion.DataSource = adoRegistroAux
            tdgProyectoPlanificacion.Refresh
                
            adoRegistroAux.Bookmark = dblBookmark
            txtMontoTotalProyectado.Text = CStr(TotalProyectoPlanificacion())
            Call LimpiarPlanificacion
        Else
            MsgBox "El total de los montos ingresados sobrepasa el monto total de planificacion", vbCritical, Me.Caption
        End If
        
    End If
    
End Sub

Private Sub cmdQuitar_Click()
    
    Dim dblBookmark As Double
    
    If tdgProyectoPlanificacion.SelBookmarks.Count <= 0 Then Exit Sub
    
    If adoRegistroAux.RecordCount > 0 Then
            
        dblBookmark = adoRegistroAux.Bookmark
        adoRegistroAux.Delete adAffectCurrent
        
        If adoRegistroAux.EOF Then
            adoRegistroAux.MovePrevious
            tdgProyectoPlanificacion.MovePrevious
        End If
        
        adoRegistroAux.Update
        
        If adoRegistroAux.RecordCount > 0 And Not adoRegistroAux.BOF And Not adoRegistroAux.EOF And dblBookmark > 1 Then adoRegistroAux.Bookmark = dblBookmark - 1

        If adoRegistroAux.RecordCount > 0 And Not adoRegistroAux.BOF And Not adoRegistroAux.EOF Then adoRegistroAux.Bookmark = dblBookmark - 1
   
        tdgProyectoPlanificacion.Refresh
        
        If adoRegistroAux.RecordCount = 0 Then
            cmdQuitar.Enabled = False
            strTipoUnidadMedidaSuperficieEnUso = Valor_Caracter
            strDescripUnidadMedidaSuperficieEnUso = Valor_Caracter
        End If
        
        txtMontoTotalProyectado.Text = CStr(TotalProyectoPlanificacion())
        
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiarPlanificacion
    cmdAgregarEditar.Caption = "&Agregar"
    cboTipoUnidadInmobiliaria.Enabled = True
    txtMedidaUnidadInmobiliaria.Enabled = True
    cmdCancelar.Visible = False
    lblSimboloMedida.Caption = lblSimboloUnidadMedidaSuperficie.Caption
    frmProyectoDatos.Enabled = True
End Sub

Private Sub SubImprimir()


    Dim strSeleccionRegistro    As String
    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strIndicador    As String, strFecDesde As String, strFecHasta   As String

    
   
            gstrNameRepo = "DefinicionProInmobiliariaGrilla"
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
