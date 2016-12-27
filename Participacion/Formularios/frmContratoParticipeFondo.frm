VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmContratoParticipeFondo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contrato Participe por Fondo"
   ClientHeight    =   9840
   ClientLeft      =   1140
   ClientTop       =   960
   ClientWidth     =   14295
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9840
   ScaleWidth      =   14295
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   11280
      TabIndex        =   10
      Top             =   9000
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
      TabIndex        =   11
      Top             =   9000
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   1296
      Buttons         =   4
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      Visible0        =   0   'False
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Consultar"
      Tag1            =   "3"
      Visible1        =   0   'False
      ToolTipText1    =   "Consultar"
      Caption2        =   "&Buscar"
      Tag2            =   "5"
      Visible2        =   0   'False
      ToolTipText2    =   "Buscar"
      Caption3        =   "&Imprimir"
      Tag3            =   "6"
      Visible3        =   0   'False
      ToolTipText3    =   "Imprimir"
      UserControlWidth=   5700
   End
   Begin TabDlg.SSTab tabContrato 
      Height          =   8805
      Left            =   -15
      TabIndex        =   6
      Top             =   60
      Width           =   14280
      _ExtentX        =   25188
      _ExtentY        =   15531
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Participe"
      TabPicture(0)   =   "frmContratoParticipeFondo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraContrato(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmContratoParticipeFondo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(1)=   "fraResumen"
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -65160
         TabIndex        =   5
         Top             =   7920
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
      Begin VB.Frame fraContrato 
         Caption         =   "Criterios de Búsqueda"
         Height          =   1575
         Index           =   0
         Left            =   360
         TabIndex        =   17
         Top             =   840
         Width           =   12600
         Begin VB.TextBox txtNumDocumentoBusq 
            Height          =   285
            Left            =   3105
            MaxLength       =   15
            TabIndex        =   23
            Top             =   900
            Width           =   2940
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   285
            Left            =   9105
            MaxLength       =   50
            TabIndex        =   22
            Top             =   495
            Width           =   2940
         End
         Begin VB.OptionButton optParticipe 
            Caption         =   "Descripción"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   6720
            TabIndex        =   21
            Top             =   510
            Width           =   1905
         End
         Begin VB.OptionButton optParticipe 
            Caption         =   "Código Partícipe"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   20
            Top             =   495
            Width           =   1905
         End
         Begin VB.TextBox txtCodParticipeBusqueda 
            Height          =   285
            Left            =   3105
            MaxLength       =   20
            TabIndex        =   19
            Top             =   480
            Width           =   2940
         End
         Begin VB.OptionButton optParticipe 
            Caption         =   "Num. Documento"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   18
            Top             =   915
            Width           =   1890
         End
         Begin VB.Label lblContador 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9105
            TabIndex        =   24
            Top             =   870
            Width           =   2940
         End
      End
      Begin VB.Frame fraResumen 
         ForeColor       =   &H00800000&
         Height          =   7245
         Left            =   -74640
         TabIndex        =   7
         Top             =   630
         Width           =   13455
         Begin VB.TextBox txtCodParticipe 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2160
            MaxLength       =   20
            TabIndex        =   39
            Top             =   870
            Width           =   2055
         End
         Begin VB.ComboBox cboEstado 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9750
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   2250
            Width           =   3045
         End
         Begin VB.ComboBox cboPromotor 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   7155
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   2700
            Width           =   5625
         End
         Begin VB.CommandButton cmdBusqueda 
            Caption         =   "..."
            Height          =   315
            Index           =   0
            Left            =   4260
            TabIndex        =   29
            ToolTipText     =   "Búsqueda de Cliente"
            Top             =   870
            Width           =   315
         End
         Begin VB.ComboBox cboFondo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2175
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   3210
            Width           =   10605
         End
         Begin VB.CommandButton cmdActualizar 
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   330
            TabIndex        =   16
            Top             =   4680
            Width           =   375
         End
         Begin VB.TextBox txtNumDocumento 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2145
            MaxLength       =   20
            TabIndex        =   14
            Top             =   2250
            Width           =   3645
         End
         Begin TrueOleDBGrid60.TDBGrid tdgMovimiento 
            Bindings        =   "frmContratoParticipeFondo.frx":0038
            Height          =   2055
            Left            =   960
            OleObjectBlob   =   "frmContratoParticipeFondo.frx":0054
            TabIndex        =   13
            Top             =   4680
            Width           =   11805
         End
         Begin VB.TextBox txtDescripComentario 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   705
            Left            =   2160
            TabIndex        =   2
            Top             =   3690
            Width           =   10605
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
            Left            =   330
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Quitar detalle"
            Top             =   5520
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
            Left            =   330
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Agregar detalle"
            Top             =   5100
            Width           =   375
         End
         Begin VB.TextBox txtHoraContrato 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4950
            Locked          =   -1  'True
            TabIndex        =   1
            Text            =   "00:00"
            Top             =   2700
            Width           =   855
         End
         Begin MSComCtl2.DTPicker dtpFechaContrato 
            Height          =   315
            Left            =   2160
            TabIndex        =   0
            Top             =   2700
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   175570945
            CurrentDate     =   38068
         End
         Begin VB.Label lblDescripParticipe 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4620
            TabIndex        =   40
            Top             =   900
            Width           =   8145
         End
         Begin VB.Label lblNumDocumento 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "07883712"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5820
            TabIndex        =   38
            Top             =   1770
            Width           =   3585
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo/Num Doc. ID."
            ForeColor       =   &H00800000&
            Height          =   270
            Index           =   4
            Left            =   360
            TabIndex        =   37
            Top             =   1800
            Width           =   1665
         End
         Begin VB.Label lblTipoDocumento 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DNI - DOCUMENTO NACIONAL DE IDENTIDAD"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2160
            TabIndex        =   36
            Top             =   1770
            Width           =   3585
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Estado"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   3
            Left            =   8730
            TabIndex        =   35
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Promotor Creac."
            ForeColor       =   &H00800000&
            Height          =   405
            Index           =   5
            Left            =   6210
            TabIndex        =   33
            Top             =   2730
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Cliente Titular"
            ForeColor       =   &H00800000&
            Height          =   270
            Index           =   2
            Left            =   360
            TabIndex        =   31
            Top             =   1365
            Width           =   1335
         End
         Begin VB.Label lblDescripTitular 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2175
            TabIndex        =   30
            Top             =   1320
            Width           =   10605
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Participe"
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   1
            Left            =   360
            TabIndex        =   28
            Top             =   900
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fondo"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   360
            TabIndex        =   26
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Nro. Contrato"
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   9
            Left            =   360
            TabIndex        =   15
            Top             =   2280
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Comentario"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   8
            Left            =   360
            TabIndex        =   12
            Top             =   3720
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Hora Creac."
            ForeColor       =   &H00800000&
            Height          =   165
            Index           =   7
            Left            =   3840
            TabIndex        =   9
            Top             =   2730
            Width           =   1065
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha Creac."
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   360
            TabIndex        =   8
            Top             =   2745
            Width           =   1275
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmContratoParticipeFondo.frx":5F12
         Height          =   5205
         Left            =   360
         OleObjectBlob   =   "frmContratoParticipeFondo.frx":5F2C
         TabIndex        =   27
         Top             =   2730
         Width           =   12615
      End
   End
End
Attribute VB_Name = "frmContratoParticipeFondo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()              As String, arrPromotor()         As String
Dim arrEstado()             As String

Dim strCodFondo             As String, strCodPromotor       As String
Dim strCodEstado            As String, strEstado            As String

Dim intRegistro             As Integer, strTipoDocumento    As String
Dim strNumDocumento         As String, adoRegistroAux       As ADODB.Recordset
Dim adoConsulta             As ADODB.Recordset

Public Sub Buscar()

    Dim strSql          As String
    Dim strFechaDesde   As String, strFechaHasta As String
    Dim datFecha        As Date
    
    Set adoConsulta = New ADODB.Recordset
    
    strEstado = Reg_Defecto
    
    If Trim(txtCodParticipeBusqueda.Text) <> "" Or Trim(txtNumDocumentoBusq.Text) <> "" Or Trim(txtDescripcion.Text) <> "" Then
        
        strSql = "SELECT CodParticipe,Tabla1.DescripParametro TipoIdentidad,NumIdentidad,DescripParticipe,FechaIngreso,Tabla2.DescripParametro TipoMancomuno, TipoIdentidad CodTipoIdentidad, ClaseParticipe "
        strSql = strSql & "FROM ParticipeContrato JOIN AuxiliarParametro Tabla1 ON(Tabla1.CodParametro=ParticipeContrato.TipoIdentidad AND Tabla1.CodTipoParametro='TIPIDE') "
        strSql = strSql & "JOIN AuxiliarParametro Tabla2 ON(Tabla2.CodParametro=ParticipeContrato.TipoMancomuno AND Tabla2.CodTipoParametro='TIPMAN') "
        
        If optParticipe(0).Value Then
            strSql = strSql & "WHERE CodParticipe='" & Trim(txtCodParticipeBusqueda.Text) & "'"
        ElseIf optParticipe(1).Value Then
            strSql = strSql & "WHERE NumIdentidad='" & Trim(txtNumDocumentoBusq.Text) & "'"
        Else
            strSql = strSql & "WHERE DescripParticipe LIKE '%" & Trim(txtDescripcion.Text) & "%'"
        End If
                                
        tdgConsulta.Columns(0).Caption = "Código"
        tdgConsulta.Columns(0).DataField = "CodParticipe"
        
        tdgConsulta.Columns(1).Caption = "Descripción"
        tdgConsulta.Columns(1).DataField = "DescripParticipe"
        
        tdgConsulta.Columns(2).Caption = "Tipo Mancomuno"
        tdgConsulta.Columns(2).DataField = "TipoMancomuno"
        
        tdgConsulta.Columns(3).Caption = "Tipo Ident."
        tdgConsulta.Columns(3).DataField = "TipoIdentidad"
        
        tdgConsulta.Columns(4).Caption = "Número"
        tdgConsulta.Columns(4).DataField = "NumIdentidad"
        
        If Not tdgConsulta.Columns(5).Visible Then tdgConsulta.Columns(5).Visible = True
            
        tdgConsulta.Columns(5).Caption = "Ingreso"
        tdgConsulta.Columns(5).DataField = "FechaIngreso"
                                
        tdgConsulta.Columns(6).Caption = "CodTipoIdentidad"
        tdgConsulta.Columns(6).DataField = "CodTipoIdentidad"
        
        tdgConsulta.Columns(7).Caption = "ClaseParticipe"
        tdgConsulta.Columns(7).DataField = "ClaseParticipe"
        
    End If
                
    If strSql = Valor_Caracter Then Exit Sub
    
    Me.MousePointer = vbHourglass
   
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSql
    End With
    
    tdgConsulta.DataSource = adoConsulta
    
    If adoConsulta.RecordCount > 0 Then strEstado = Reg_Consulta
                        
    Me.MousePointer = vbDefault


End Sub


'Private Sub TotalizarMovimientos()
'
'    Dim strFechaDesde   As String, strFechaHasta As String
'    Dim datFecha        As Date
'    Dim strSQL          As String
'    Dim intRegistro     As Integer
'
'
'    Dim dblMontoDebe        As Double, dblMontoHaber        As Double
'    Dim dblAcumuladoDebeMN  As Double, dblAcumuladoDebeME   As Double
'    Dim dblAcumuladoHaberMN As Double, dblAcumuladoHaberME  As Double
'    Dim intContador         As Integer
'
'    'intContador = adoRegistroAux.RecordCount - 1
'
'    lblTotalDebeME.Caption = "0"
'    lblTotalHaberME.Caption = "0"
'    lblTotalDebeMN.Caption = "0"
'    lblTotalHaberMN.Caption = "0"
'
'    dblAcumuladoDebeMN = 0
'    dblAcumuladoHaberMN = 0
'    dblAcumuladoDebeME = 0
'    dblAcumuladoHaberME = 0
'
'    If Not adoRegistroAux.EOF And Not adoRegistroAux.BOF Then
'        adoRegistroAux.MoveFirst
'    End If
'
'    While Not adoRegistroAux.EOF
'        dblMontoDebe = CDbl(adoRegistroAux.Fields("MontoDebe"))
'        dblMontoHaber = CDbl(adoRegistroAux.Fields("MontoHaber"))
'
'        If adoRegistroAux.Fields("CodMonedaMovimiento") = Codigo_Moneda_Local Then
'            dblAcumuladoDebeMN = dblAcumuladoDebeMN + dblMontoDebe
'            dblAcumuladoHaberMN = dblAcumuladoHaberMN + dblMontoHaber
'        Else
'            dblAcumuladoDebeME = dblAcumuladoDebeME + dblMontoDebe
'            dblAcumuladoHaberME = dblAcumuladoHaberME + dblMontoHaber
'        End If
'
'        adoRegistroAux.MoveNext
'    Wend
'
'    lblTotalDebeME.Caption = CStr(dblAcumuladoDebeME)
'    lblTotalHaberME.Caption = CStr(dblAcumuladoHaberME)
'    lblTotalDebeMN.Caption = CStr(dblAcumuladoDebeMN)
'    lblTotalHaberMN.Caption = CStr(dblAcumuladoHaberMN)
'
'
'End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabContrato
        .TabEnabled(0) = True
        .Tab = 0
    End With
    Call Buscar
    
End Sub

Private Sub CargarReportes()

   frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Diario General"
    
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Diario General (ME)"
    
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Mayor General"
    
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Text = "Mayor General (ME)"
    
    
End Sub
Private Sub Deshabilita()

End Sub



Private Sub Habilita()

End Sub


Public Sub Imprimir()

    'Call SubImprimir(1)
    
End Sub



'Private Sub LimpiarDatos()
'
'    txtCodCuenta.Text = Valor_Caracter
'    txtDescripCuenta.Text = Valor_Caracter
'    txtDescripFileAnalitica.Text = Valor_Caracter
'    txtDescripAuxiliar.Text = Valor_Caracter
'
'    strTipoAuxiliar = Valor_Caracter
'    strCodAuxiliar = Valor_Caracter
'    strCodCuenta = Valor_Caracter
'
'    txtCodFile.Text = Valor_Caracter
'    txtCodAnalitica.Text = Valor_Caracter
'    txtMontoMovimiento.Text = "0"
'    txtDescripMovimiento.Text = Valor_Caracter
'
'End Sub

Public Sub Salir()

    Unload Me
    
End Sub

'Public Sub SubImprimir(index As Integer)
'
'    Dim adoRegistro             As ADODB.Recordset
'    Dim strSeleccionRegistro    As String
'    Dim frmReporte              As frmVisorReporte
'    Dim aReportParamS(), aReportParamF(), aReportParamFn()
'
''    Select Case index
''
''
''       Case 1
''
'''            If tabAsiento.Tab = 1 And strEstado = Reg_Edicion Then
'''                '*** Comprobante seleccionado ***
'''                strSeleccionRegistro = "{AsientoContable.NumAsiento} = '" & Trim(lblNumAsiento.Caption) & "'"
'''                strSeleccionRegistro = strSeleccionRegistro & " AND {AsientoContableDetalle.NumAsiento} = '" & Trim(lblNumAsiento.Caption) & "'"
'''                strSeleccionRegistro = strSeleccionRegistro & " AND {AsientoContableDetalle.FechaMovimiento} = '" & Convertyyyymmdd(dtpFechaAsiento.Value) & "'"
'''                gstrSelFrml = strSeleccionRegistro
'''            Else
'''                '*** Lista de comprobantes por rango de fecha ***
'''                strSeleccionRegistro = "{AsientoContable.FechaAsiento} IN 'Fch1' TO 'Fch2'"
'''                gstrSelFrml = strSeleccionRegistro
'''                'frmRangoFecha.Show vbModal
'''                frmFiltroReporte.Show vbModal
'''            End If
''
''
''            If gstrSelFrml <> "0" Then
''                Set frmReporte = New frmVisorReporte
''
''                ReDim aReportParamS(8)
''                ReDim aReportParamFn(5)
''                ReDim aReportParamF(5)
''
''                aReportParamFn(0) = "Usuario"
''                aReportParamFn(1) = "FechaDesde"
''                aReportParamFn(2) = "FechaHasta"
''                aReportParamFn(3) = "Hora"
''                aReportParamFn(4) = "Fondo"
''                aReportParamFn(5) = "NombreEmpresa"
''
''                aReportParamF(0) = gstrLogin
''                aReportParamF(3) = Format(Time(), "hh:mm:ss")
''                aReportParamF(4) = Trim(cboFondo.Text)
''
'''                If tabAsiento.Tab = 1 And strEstado = Reg_Edicion Then
'''                    aReportParamF(1) = dtpFechaAsiento
'''                    aReportParamF(2) = dtpFechaAsiento
'''                Else
'''                    aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
'''                    aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
'''                End If
''                aReportParamF(5) = gstrNombreEmpresa & Space(1)
''
''                aReportParamS(0) = strCodFondo
''                aReportParamS(1) = gstrCodAdministradora
''                aReportParamS(5) = gstrCodClaseTipoCambioFondo
''                aReportParamS(6) = gstrValorTipoCambioCierre
''
'''                If tabAsiento.Tab = 1 And strEstado = Reg_Edicion Then
'''                    aReportParamS(2) = Convertyyyymmdd(dtpFechaAsiento.Value)
'''                    aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, dtpFechaAsiento.Value))
'''                    aReportParamS(4) = Codigo_Moneda_Local
'''                    aReportParamS(7) = Codigo_Listar_Individual
'''                    aReportParamS(8) = Trim(lblNumAsiento.Caption)
'''                Else
''                    aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
''                    aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
''                    aReportParamS(4) = Codigo_Moneda_Local
''                    aReportParamS(7) = Codigo_Listar_Todos
''                    aReportParamS(8) = "0000000000"
'''                End If
''
''                If gstrCodMonedaReporte = "01" Then
'''                    If chkSimulacion.Value Then
'''                        gstrNameRepo = "SLibroDiario"
'''                    Else
'''                        gstrNameRepo = "LibroDiario"
'''                    End If
''                Else
'''                    If chkSimulacion.Value Then
'''                        gstrNameRepo = "SLibroDiario"
'''                    Else
''                        gstrNameRepo = "LibroDiarioME"
'''                    End If
''                End If
''
''
''            End If
''
''
''        Case 2
''
''            '*** Lista de comprobantes por rango de fecha ***
''            strSeleccionRegistro = "{AsientoContable.FechaAsiento} IN 'Fch1' TO 'Fch2'"
''            gstrSelFrml = strSeleccionRegistro
''            'frmRangoFecha.Show vbModal
''            frmFiltroReporte.Show vbModal
''
''            If gstrSelFrml <> "0" Then
''                Set adoRegistro = New ADODB.Recordset
''
''                '*** Se Realizó Cierre anteriormente ? ***
''                adoComm.CommandText = "{ call up_GNValidaCierreRealizado('" & _
''                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
''                    Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)) & "','" & _
''                    Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)))) & "') }"
''                Set adoRegistro = adoComm.Execute
''
''                If adoRegistro.EOF Then
''                    MsgBox "El Cierre del Día " & Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10) & " No fué realizado.", vbCritical, Me.Caption
''
''                    adoRegistro.Close: Set adoRegistro = Nothing
''                    Exit Sub
''                End If
''                adoRegistro.Close: Set adoRegistro = Nothing
''
''                Set frmReporte = New frmVisorReporte
''
''                Dim strCuenta As String
''
''                ReDim aReportParamS(8)
''                ReDim aReportParamFn(5)
''                ReDim aReportParamF(5)
''
''                aReportParamFn(0) = "Usuario"
''                aReportParamFn(1) = "FechaDesde"
''                aReportParamFn(2) = "FechaHasta"
''                aReportParamFn(3) = "Hora"
''                aReportParamFn(4) = "Fondo"
''                aReportParamFn(5) = "NombreEmpresa"
''
''                aReportParamF(0) = gstrLogin
''                aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
''                aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
''                aReportParamF(3) = Format(Time(), "hh:mm:ss")
''                aReportParamF(4) = Trim(cboFondo.Text)
''                aReportParamF(5) = gstrNombreEmpresa & Space(1)
''
''                aReportParamS(0) = strCodFondo
''                aReportParamS(1) = gstrCodAdministradora
''                aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
''                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
''                aReportParamS(4) = gstrCodMonedaReporte 'strCodMoneda
''                aReportParamS(5) = gstrCodClaseTipoCambioFondo 'Codigo_Listar_Todos
''                aReportParamS(6) = gstrValorTipoCambioCierre   '"0000000000"
''
''                If gstrCodCuenta = "0000000000" Then
''                    aReportParamS(7) = Codigo_Listar_Todos
''                Else
''                    aReportParamS(7) = Codigo_Listar_Individual
''                End If
''                aReportParamS(8) = gstrCodCuenta '"0000000000"
''
''                If gstrCodMonedaReporte = Codigo_Moneda_Local Then
''                    gstrNameRepo = "LibroMayor"
''                Else
''                    gstrNameRepo = "HistLibroMayor2"
''                End If
''
''
''            End If
''    End Select
'
'    If gstrSelFrml = "0" Then Exit Sub
'
'    gstrSelFrml = Valor_Caracter
'    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"
'
'    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())
'
'    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
'    frmReporte.Show vbModal
'
'    Set frmReporte = Nothing
'
'    Screen.MousePointer = vbNormal
'
'End Sub



Private Function ValidaCuadreContable() As Boolean

    Dim curMontoDebe        As Currency, curMontoHaber      As Currency
    Dim curMontoContable    As Currency
    
    ValidaCuadreContable = False
    
    adoRegistroAux.MoveFirst
    
    Do While Not adoRegistroAux.EOF
        If adoRegistroAux.Fields("IndDebeHaber") = Codigo_Tipo_Naturaleza_Debe Then
            curMontoDebe = CCur(adoRegistroAux.Fields("MontoContable"))
        Else
            curMontoHaber = CCur(adoRegistroAux.Fields("MontoContable"))
        End If
        curMontoContable = curMontoContable + CCur(adoRegistroAux.Fields("MontoContable"))
                
        adoRegistroAux.MoveNext
    Loop
    
    If curMontoContable <> 0 Then Exit Function
    
    ValidaCuadreContable = True
    
End Function


Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = adoRegistro("FechaCuota")
            gdblTipoCambio = adoRegistro("ValorTipoCambio")
            gstrCodMoneda = adoRegistro("CodMoneda")
            'dtpFechaDesde.Value = gdatFechaActual
            'dtpFechaHasta.Value = dtpFechaDesde.Value
            'txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, Codigo_Moneda_Local, gstrCodMoneda))
            'If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, gdatFechaActual), Codigo_Moneda_Local, gstrCodMoneda))
            'gdblTipoCambio = CDbl(txtTipoCambio.Text)
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        
            'Call Buscar
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub




Private Sub cboEstado_Click()

    strCodEstado = Valor_Caracter
    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strCodEstado = Trim(arrEstado(cboEstado.ListIndex))
    
   
End Sub


'Private Sub cboTipoAuxiliar_Click()
'
'
'    strTipoAuxiliar = ""
'
'    If cboTipoAuxiliar.ListIndex < 0 Then Exit Sub
'
'    strTipoAuxiliar = Trim(arrTipoAuxiliar(cboTipoAuxiliar.ListIndex))
'    'If strTipoAuxiliar = Valor_Caracter Then strTipoAuxiliar = "00"
'
'    strCodAuxiliar = ""
'    strDescripAuxiliar = ""
'    txtDescripAuxiliar.Text = ""
'
'End Sub


Private Sub cboPromotor_Click()

    strCodPromotor = Valor_Caracter
    If cboPromotor.ListIndex < 0 Then Exit Sub
    
    strCodPromotor = Trim(arrPromotor(cboPromotor.ListIndex))
    
   
End Sub

Private Sub cmdActualizar_Click()


'    [CodParticipe] [char](20) NOT NULL,
'    [CodFondo] CodigoCorto NOT NULL,
'    [CodAdministradora] CodigoCorto NOT NULL,
'    [NumContrato] [char](15) NOT NULL,
'    [FechaContrato] [datetime] NOT NULL,
'    [CodPromotor] varchar(8) NOT NULL,
'    [CodSucursal] varchar(6) NOT NULL,
'    [CodAgencia] varchar(3) NOT NULL,
'    [EstadoContrato] Codigo NOT NULL,
'    [DescripComentario] varchar(200) NOT NULL,
  If tdgMovimiento.SelBookmarks.Count >= 1 Then
    adoRegistroAux.Fields("CodParticipe") = txtCodParticipe.Text
    adoRegistroAux.Fields("CodFondo") = strCodFondo
    adoRegistroAux.Fields("CodAdministradora") = gstrCodAdministradora
    adoRegistroAux.Fields("DescripFondo") = cboFondo.List(cboFondo.ListIndex)
    adoRegistroAux.Fields("NumContrato") = Trim(txtNumDocumento.Text)
    adoRegistroAux.Fields("FechaContrato") = dtpFechaContrato.Value
    adoRegistroAux.Fields("CodPromotor") = strCodPromotor
    adoRegistroAux.Fields("CodSucursal") = "" 'strCodSucursal
    adoRegistroAux.Fields("CodAgencia") = ""  'strCodAgencia
    adoRegistroAux.Fields("EstadoContrato") = strCodEstado
    adoRegistroAux.Fields("DescripComentario") = Trim(txtDescripComentario.Text)
End If


End Sub

Private Sub cmdAgregar_Click()

    Dim adoRegistro As ADODB.Recordset
    Dim intSecuencial As Integer
    Dim dblBookmark As Double
    
    'If strEstado = Reg_Consulta Then Exit Sub
    
    If strEstado = Reg_Adicion Or strEstado = Reg_Edicion Then
        If TodoOK() Then
           
            adoRegistroAux.AddNew
            
            adoRegistroAux.Fields("CodParticipe") = txtCodParticipe.Text
            adoRegistroAux.Fields("CodFondo") = strCodFondo
            adoRegistroAux.Fields("CodAdministradora") = gstrCodAdministradora
            adoRegistroAux.Fields("DescripFondo") = cboFondo.List(cboFondo.ListIndex)
            adoRegistroAux.Fields("NumContrato") = Trim(txtNumDocumento.Text)
            adoRegistroAux.Fields("FechaContrato") = dtpFechaContrato.Value
            adoRegistroAux.Fields("CodPromotor") = strCodPromotor
            adoRegistroAux.Fields("CodSucursal") = "" 'strCodSucursal
            adoRegistroAux.Fields("CodAgencia") = "" 'strCodAgencia
            adoRegistroAux.Fields("EstadoContrato") = strCodEstado
            adoRegistroAux.Fields("DescripComentario") = Trim(txtDescripComentario.Text)
            
            adoRegistroAux.Update
            
            dblBookmark = adoRegistroAux.Bookmark
            
            tdgMovimiento.Refresh
            
            'Call NumerarRegistros
            
            adoRegistroAux.Bookmark = dblBookmark
                                    
            cmdQuitar.Enabled = True
        
        End If
    End If
    
End Sub
Private Sub NumerarRegistros()

    Dim n As Long
    
    n = 1
    
    If Not adoRegistroAux.EOF And Not adoRegistroAux.BOF Then
        adoRegistroAux.MoveFirst
    End If
    
    While Not adoRegistroAux.EOF
        adoRegistroAux.Fields("SecMovimiento") = n
        adoRegistroAux.Update
        n = n + 1
        adoRegistroAux.MoveNext
    Wend


End Sub

Private Sub cmdBusqueda_Click(index As Integer)

    Dim sSql As String
    Dim frmBus As frmBuscar

    Set frmBus = New frmBuscar

    With frmBus.TBuscarRegistro1

        .ADOConexion = adoConn
        .ADOConexion.CommandTimeout = 0

        .iTipoGrilla = 2
        
        frmBus.Caption = " Relación de Participes"
        .sSql = "{ call up_ACSelDatos(30) }"
        
        .OutputColumns = "1,2,3,4,5,6,7,8"
        .HiddenColumns = "1,2,5,6,7,8"
        
        .BuscarTabla
        
        frmBus.Show 1
       
        If .iParams.Count = 0 Then Exit Sub
        
        If .iParams(1).Valor <> "" Then
            txtCodParticipe.Text = Trim(.iParams(1).Valor)
            lblTipoDocumento.Caption = Trim(.iParams(2).Valor)
            lblNumDocumento.Caption = Trim(.iParams(3).Valor)
            lblDescripParticipe.Caption = Trim(.iParams(4).Valor)
            lblDescripTitular.Caption = Trim(.iParams(8).Valor)
            strTipoDocumento = Trim(.iParams(6).Valor)
        Else
            txtCodParticipe.Text = Valor_Caracter
            lblTipoDocumento.Caption = Valor_Caracter
            lblNumDocumento.Caption = Valor_Caracter
            lblDescripParticipe.Caption = Valor_Caracter
            lblDescripTitular.Caption = Valor_Caracter
            strTipoDocumento = Valor_Caracter
        End If


    End With

    Set frmBus = Nothing


End Sub







'Private Sub cmdBusqueda_Click(index As Integer)
'
'   Dim sSql As String
'
'
'    Dim frmBus As frmBuscar
'
'    Set frmBus = New frmBuscar
'
'    With frmBus.TBuscarRegistro1
'
'        .ADOConexion = adoConn
'        .ADOConexion.CommandTimeout = 0
'        'If Index <> 2 Then
'        '    .iTipoGrilla = 1
'        'Else
'        '    .iTipoGrilla = 2
'        .iTipoGrilla = 2
'
'        Select Case index
'
'            Case 0
'
'
'                frmBus.Caption = " Relación de Cuentas Contables"
'                .sSql = "SELECT CodCuenta,DescripCuenta,TipoFile,IndAuxiliar,TipoAuxiliar FROM PlanContable "
'                .sSql = .sSql & " WHERE IndMovimiento='" & Valor_Indicador & "' AND CodAdministradora='" & gstrCodAdministradora & "' ORDER BY CodCuenta"
'                .OutputColumns = "1,2,3,4,5"
'                .HiddenColumns = "3,4,5"
'
'            Case 1
'
'                frmBus.Caption = " Relación de File Analiticas"
'                .sSql = "{ call up_CNSelFileAnalitico('" & strCodFondo & "','" & gstrCodAdministradora & "','" & strCodCuenta & "','" & strTipoFile & "') }"
'                .OutputColumns = "1,2,3,4,5"
'                .HiddenColumns = ""
'
'
'            Case 2
'
''                If cboTipoAuxiliar.ListIndex = -1 Then
''                    MsgBox "Seleccione primero el Tipo de Auxiliar!", vbInformation + vbOKOnly, Me.Caption
''                    Exit Sub
''                End If
'
'                frmBus.Caption = " Relación de Auxiliares Contables"
'                .sSql = "{ call up_CNSelAuxiliarContable('" & strTipoAuxiliar & "') }"
'                .OutputColumns = "1,2,3"
'                .HiddenColumns = "3"
'
'        End Select
'
'        Screen.MousePointer = vbHourglass
'
'        .BuscarTabla
'
'        Screen.MousePointer = vbNormal
'        frmBus.Show 1
'
'        If .iParams.Count = 0 Then Exit Sub
'
'        If .iParams(1).Valor <> "" Then
'
'
'            Select Case index
'
'                Case 0
'
''                    strCodCustodio = .iParams(1).Valor  '.sCodigo
''                    txtDescripCustodio.Text = .iParams(2).Valor '.sDescripcion
'                    strTipoFile = Trim(.iParams(3).Valor)
'                    strIndAuxiliar = Trim(.iParams(4).Valor)
'                    strTipoAuxiliar = Trim(.iParams(5).Valor)
'
'                    strCodCuenta = Trim(.iParams(1).Valor)
'                    strDescripCuenta = Trim(.iParams(2).Valor)
'
'                    txtCodCuenta.Text = strCodCuenta
'
'                    txtDescripCuenta.Text = strCodCuenta & " - " & strDescripCuenta
'
'                    txtDescripFileAnalitica.Text = ""
'                    strCodFile = ""
'                    strCodAnalitica = ""
'                    strDescripFileAnalitica = ""
'
'                    'txtCodFile.Text = ""
'                    'txtCodAnalitica.Text = ""
'
'                    If strIndAuxiliar = Valor_Indicador Then
'                        cmdBusqueda(2).Enabled = True
''                        If strTipoAuxiliar = "00" Then 'Todos
''                            cboTipoAuxiliar.ListIndex = -1
''                            cboTipoAuxiliar.Locked = False
''                        Else
''                            intRegistro = ObtenerItemLista(arrTipoAuxiliar(), strTipoAuxiliar)
''                            If intRegistro >= 0 Then cboTipoAuxiliar.ListIndex = intRegistro
''                            cboTipoAuxiliar.Locked = True
''                        End If
'                    Else
'                        cmdBusqueda(2).Enabled = False
'                        txtDescripAuxiliar.Text = ""
'                        strTipoAuxiliar = ""
'                        strCodAuxiliar = ""
'                    End If
'
'                    If strTipoFile = Valor_Caracter Then
'                        cmdBusqueda(1).Enabled = False
'                    Else
'                        cmdBusqueda(1).Enabled = True
''                        intRegistro = ObtenerItemLista(arrTipoFile(), strTipoFile)
''                        If intRegistro >= 0 Then cboTipoFile.ListIndex = intRegistro
'                    End If
'
'                Case 1
'
'                    strCodFile = Trim(.iParams(1).Valor)
'                    strCodAnalitica = Trim(.iParams(2).Valor)
'                    strDescripFileAnalitica = Trim(.iParams(3).Valor)
'                    strCodMoneda = Trim(.iParams(4).Valor)
'
'                    txtCodFile.Text = strCodFile
'                    txtCodAnalitica.Text = strCodAnalitica
'
'                    If strTipoFile = Valor_File_Generico Then
'                        txtCodAnalitica.Enabled = True
'                        txtDescripFileAnalitica.Text = "Analítica Genérica"
'                    Else
'                        txtDescripFileAnalitica.Text = strCodFile & "-" & strCodAnalitica & " - " & strDescripFileAnalitica
'                        txtCodAnalitica.Enabled = True 'False
'                    End If
'
'                    cboMonedaMovimiento.ListIndex = -1
'                    intRegistro = ObtenerItemLista(arrMonedaMovimiento(), strCodMoneda)
'                    If intRegistro >= 0 Then cboMonedaMovimiento.ListIndex = intRegistro
'
'                Case 2
'
'                     strCodAuxiliar = Trim(.iParams(1).Valor)
'                     strDescripAuxiliar = Trim(.iParams(2).Valor)
'
'                     txtDescripAuxiliar.Text = strDescripAuxiliar
'
'            End Select
'
'        End If
'
'
'    End With
'
'    Set frmBus = Nothing
'
'End Sub

Private Sub cmdQuitar_Click()

    Dim dblBookmark As Double

    
    If adoRegistroAux.RecordCount > 0 Then
    
        dblBookmark = adoRegistroAux.Bookmark
    
        adoRegistroAux.Delete adAffectCurrent
        
        If adoRegistroAux.EOF Then
            adoRegistroAux.MovePrevious
            tdgMovimiento.MovePrevious
        End If
            
        adoRegistroAux.Update
        
        If adoRegistroAux.RecordCount = 0 Then cmdQuitar.Enabled = False

        If adoRegistroAux.RecordCount > 0 And Not adoRegistroAux.BOF And Not adoRegistroAux.EOF And dblBookmark > 1 Then adoRegistroAux.Bookmark = dblBookmark - 1
        
        'Call NumerarRegistros
        
        'If adoRegistroAux.RecordCount > 0 And Not adoRegistroAux.BOF And Not adoRegistroAux.EOF Then adoRegistroAux.Bookmark = dblBookmark - 1
        
        'Call TotalizarMovimientos
        
        'If adoRegistroAux.RecordCount > 0 And Not adoRegistroAux.BOF And Not adoRegistroAux.EOF Then adoRegistroAux.Bookmark = dblBookmark - 1
   
        'tdgMovimiento.Refresh
    
    End If
    
    
End Sub

Private Sub Form_Activate()

'    Call CargarReportes
    
End Sub

Private Sub Form_Deactivate()

    Call OcultarReportes
    
End Sub


Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
'    Call CargarReportes
    Call Buscar
    Call DarFormato
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
        
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
            
End Sub
Private Sub CargarListas()

    Dim strSql  As String
    
    '*** Fondos ***
    strSql = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSql, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    strSql = "{ call up_ACSelDatos(6) }"
    CargarControlLista strSql, cboPromotor, arrPromotor(), Valor_Caracter
    
    '*** Estados ***
    strSql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='INDREG' AND CodParametro<>'03' ORDER BY DescripParametro"
    CargarControlLista strSql, cboEstado, arrEstado(), Valor_Caracter
    

        
End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabContrato.Tab = 0
    tabContrato.TabEnabled(1) = False
        
    Call ConfiguraRecordsetAuxiliar
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 15
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 50
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 18
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 10
    
'    tdgMovimiento.Columns(0).Width = tdgMovimiento.Width * 0.01 * 6
'    tdgMovimiento.Columns(1).Width = tdgMovimiento.Width * 0.01 * 12
'    tdgMovimiento.Columns(2).Width = tdgMovimiento.Width * 0.01 * 12
'    tdgMovimiento.Columns(3).Width = tdgMovimiento.Width * 0.01 * 20
'    tdgMovimiento.Columns(4).Width = tdgMovimiento.Width * 0.01 * 20
            
    'MsgBox gstrTempPath & "Layout.grx"
            
    tdgMovimiento.LayoutFileName = gstrTempPath & "Layout.grx"
    tdgMovimiento.Layouts.Add "TestLayout"
            
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    Set frmContratoParticipeFondo = Nothing
    
End Sub

Private Sub ExtornarComprobante(strpCodFondo As String, strpNumAsiento As String)

'    Dim strSQL As String, adoresultaux1 As New Recordset, adoresultaux4 As New Recordset
'    Dim WComCon As RCabasicon, WMovCon() As RDetasicon
'    Dim Sec_com As Long, Res As Integer, s_OldCom$, NewNroCom$, nSec%
'    Dim s_DscNewCom$
'
'    On Error GoTo Ctrl_Error
'
'    '* (2) Extorno Contable
'    '** Leer comprobante origen
'    With adoComm
'        .CommandText = "SELECT * FROM FMCOMPRO WHERE COD_FOND='" & s_ParCodFon$ & "' AND NRO_COMP='" & s_ParNroCom$ & "' "
'        Set adoresultaux4 = .Execute
'
'        If Not adoresultaux4.EOF Then
'                .CommandText = "SELECT FLG_CIER FROM FMCUOTAS WHERE COD_FOND='" & strCodFond & "' AND "
'                .CommandText = .CommandText & "FCH_CUOT='" & adoresultaux4("FCH_COMP") & "'"
'                Set adoresultaux1 = .Execute
'
'                If Not adoresultaux1.EOF Then
'                    If adoresultaux1("FLG_CIER") = "X" Then
'                        MsgBox "El comprobante contable no puede ser extornado por haber sido mayorizado en un cierre contable.", vbCritical, Me.Caption
'                        adoresultaux1.Close: Set adoresultaux1 = Nothing
'                        gblnRollBack = True
'                        Exit Sub
'                    End If
'                End If
'                adoresultaux1.Close: Set adoresultaux1 = Nothing
'        End If
'
'        Sec_com = 1
'        .CommandText = "SELECT NRO_ULTI_SOLI FROM FMPARAME WHERE COD_FOND='" & s_ParCodFon$ & "' AND COD_PARA='COM'"
'        Set adoresultaux1 = .Execute
'
'        If Not IsNull(adoresultaux1("NRO_ULTI_SOLI")) Then Sec_com = CLng(adoresultaux1("NRO_ULTI_SOLI")) + 1
'        adoresultaux1.Close: Set adoresultaux1 = Nothing
'
'    End With
'
'    WComCon.COD_FOND = adoresultaux4("COD_FOND")
'    WComCon.COD_MONC = adoresultaux4("COD_MONC")
'    WComCon.CNT_MOVI = adoresultaux4("CNT_MOVI")
'    WComCon.COD_MONE = adoresultaux4("COD_MONE")
'    WComCon.FLG_AUTO = adoresultaux4("FLG_AUTO")
'    WComCon.FLG_CONT = adoresultaux4("FLG_CONT")
'    WComCon.GEN_COMP = adoresultaux4("GEN_COMP")
'    WComCon.HOR_COMP = adoresultaux4("HOR_COMP")
'    WComCon.NRO_DOCU = ""
'    WComCon.NRO_OPER = adoresultaux4("NRO_OPER")
'    WComCon.PER_DIGI = gstrLogin
'    WComCon.PER_REVI = ""
'    WComCon.STA_COMP = ""
'    WComCon.SUB_SIST = "C"
'    WComCon.TIP_CAMB = adoresultaux4("TIP_CAMB")
'    WComCon.TIP_COMP = ""
'    WComCon.TIP_DOCU = ""
'    '** Variable
'    WComCon.VAL_COMP = adoresultaux4("VAL_COMP")
'    WComCon.NRO_COMP = Format(Sec_com, "00000000")
'    WComCon.FCH_COMP = FmtFec(gstrFechaAct, "win", "yyyymmdd", Res)
'    WComCon.FCH_CONT = FmtFec(gstrFechaAct, "win", "yyyymmdd", Res)
'    WComCon.MES_CONT = Mid$(WComCon.FCH_COMP, 5, 2)
'    WComCon.prd_cont = Mid$(WComCon.FCH_COMP, 1, 4)
'    WComCon.DSL_COMP = "Ext Comprobante (" & adoresultaux4("NRO_COMP") & ">>" & adoresultaux4("FCH_COMP") & ")"
'    WComCon.GLO_COMP = WComCon.DSL_COMP
'    s_OldCom = WComCon.NRO_COMP
'    NewNroCom = Format(Sec_com, "00000000")
'    WComCon.NRO_COMP = NewNroCom
'    s_DscNewCom$ = "(Ext)" & adoresultaux4("DSL_COMP")
'    adoresultaux4.Close
'
'    adoresultaux4.CursorLocation = adUseClient
'    adoresultaux4.CursorType = adOpenStatic
'
'    strSQL = "SELECT * FROM FMMOVCON WHERE COD_FOND='" & s_ParCodFon$ & "' AND NRO_COMP='" & s_ParNroCom$ & "' "
'    adoComm.CommandText = strSQL
'    adoresultaux4.Open adoComm.CommandText, adoConn, , , adCmdText
'    'Set adoresultaux4 = adoComm.Execute
'
'    If adoresultaux4.EOF Then
'        MsgBox "El Sistema no puede encontrar el comprobante contable.", vbExclamation
'        Exit Sub
'    Else
'        adoresultaux4.MoveLast
'        ReDim WMovCon(adoresultaux4.RecordCount)
'        adoresultaux4.MoveFirst
'    End If
'    nSec = 0
'    Do While Not adoresultaux4.EOF
'       nSec = nSec + 1
'       LIniRDetAsiCon WMovCon(nSec)
'       WMovCon(nSec).SEC_MOVI = adoresultaux4("SEC_MOVI")
'       WMovCon(nSec).COD_FOND = adoresultaux4("COD_FOND")
'       WMovCon(nSec).COD_MONE = adoresultaux4("COD_MONE")
'       WMovCon(nSec).FLG_PROC = "X"
'       WMovCon(nSec).STA_MOVI = "X"
'       WMovCon(nSec).TIP_GENR = "P"
'       WMovCon(nSec).CTA_AMAR = ""
'       WMovCon(nSec).CTA_AUTO = ""
'       WMovCon(nSec).CTA_ORIG = ""
'       WMovCon(nSec).COD_FILE = adoresultaux4("COD_FILE")
'       WMovCon(nSec).COD_ANAL = adoresultaux4("COD_ANAL")
'       WMovCon(nSec).FCH_MOVI = FmtFec(gstrFechaAct, "win", "yyyymmdd", Res)
'       WMovCon(nSec).prd_cont = Mid$(WMovCon(nSec).FCH_MOVI, 1, 4)
'       WMovCon(nSec).MES_COMP = Mid$(WMovCon(nSec).FCH_MOVI, 5, 2)
'       WMovCon(nSec).NRO_COMP = WComCon.NRO_COMP
'       WMovCon(nSec).FLG_DEHA = IIf(adoresultaux4("FLG_DEHA") = "D", "H", "D")
'       WMovCon(nSec).COD_CTA = adoresultaux4("COD_CTA")
'       WMovCon(nSec).DSC_MOVI = "Extorno(" & adoresultaux4("DSC_MOVI") & ")"
'       WMovCon(nSec).COD_FILE = adoresultaux4("COD_FILE")
'       WMovCon(nSec).COD_ANAL = adoresultaux4("COD_ANAL")
'       WMovCon(nSec).VAL_MOVN = (adoresultaux4("VAL_MOVN") * -1)
'       WMovCon(nSec).VAL_MOVX = (adoresultaux4("VAL_MOVX") * -1)
'       WMovCon(nSec).VAL_CONT = (adoresultaux4("VAL_CONT") * -1)
'       adoresultaux4.MoveNext
'    Loop
'
'    WComCon.CNT_MOVI = nSec
'    LGraAsi WComCon, WMovCon() 'Grabar el asiento
'    Call UpdNewNro(s_ParCodFon$, "COM", NewNroCom)
'
''    strsql = "UPDATE fmCompro SET "
''    strsql = strsql & " DSL_COMP='" & s_DscNewCom$ & "',"
''    strsql = strsql & " GLO_COMP='" & s_DscNewCom$ & "'"
''    strsql = strsql & " WHERE COD_FOND='" & s_ParCodFon$ & "'"
''    strsql = strsql & " AND NRO_COMP='" & s_ParNroCom$ & "'"
''    adoConn.Execute strsql
'    Exit Sub
'
'Ctrl_Error:
'    gblnRollBack = True
'    MsgBox "Error " & Err.Number & " => " & Err.Description, vbCritical
'    Exit Sub
    
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
        Case vReport
            Call Imprimir
        Case vPrint
            Call SubImprimir
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
        strEstado = Reg_Edicion
        Call Deshabilita
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabContrato
            .TabEnabled(0) = False
            .Tab = 1
        End With
    End If
    
End Sub

Public Sub Eliminar()
    
    Dim strFechaGrabar As String
    Dim strNumAsiento As String
    
    If strEstado <> Reg_Edicion Then
        If strEstado <> Reg_Consulta Then Exit Sub
    End If

    If MsgBox("Desea Anular el comprobante contable " & tdgConsulta.Columns(0).Value & " ?", vbYesNo + vbQuestion + vbDefaultButton2) <> vbYes Then
        Exit Sub
    End If

    On Error GoTo Ctrl_Error

    Me.MousePointer = vbHourglass
                                        
    With adoComm
        
        .CommandType = adCmdText
        
        strFechaGrabar = Convertyyyymmdd(gdatFechaActual) & Space(1) & Format(Time, "hh:ss")
    
        strNumAsiento = tdgConsulta.Columns("NumAsiento").Value
        
        '*** Cabecera ***
        .CommandText = "{ call up_ACProcAsientoContableAnulacion('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "') }"
        adoConn.Execute .CommandText
       
'ASIENTO DE EXTORNO -- POR EL MOMENTO DESHABILITADO
'        .CommandText = "{ call up_ACProcAsientoContableExtorno('" & _
'            strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
'            strFechaGrabar & "') }"
'        adoConn.Execute .CommandText
       
        Me.MousePointer = vbDefault
       
        Call Buscar
        
        MsgBox Mensaje_Proceso_Exitoso, vbExclamation
            
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"

                                                                       
    End With
    
    Exit Sub
    
Ctrl_Error:
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
    
    
End Sub

Public Sub Grabar()
    
    Dim objParticipeContratoFondoXML    As DOMDocument60
    Dim strMsgError                     As String
    
    Dim strParticipeContratoFondoXML    As String
    Dim strCodParticipe As String
    
    
    If strEstado = Reg_Adicion Or strEstado = Reg_Edicion Then
        If TodoOK() Then
            Dim intCantRegistros    As Integer, intRegistro         As Integer
            Dim adoRegistro         As ADODB.Recordset
            Dim strNumAsiento       As String, strFechaGrabar       As String
            
            Me.MousePointer = vbHourglass
                                                
            With adoComm
                
                strFechaGrabar = Convertyyyymmdd(dtpFechaContrato.Value) & Space(1) & Format(Time, "hh:ss")
                strCodParticipe = Trim(txtCodParticipe.Text)
                
                On Error GoTo Errorhandler
                
                Call XMLADORecordset(objParticipeContratoFondoXML, "ParticipeContratoFondo", "Fondo", adoRegistroAux, strMsgError)
                strParticipeContratoFondoXML = objParticipeContratoFondoXML.xml 'CrearXMLDetalle(objTipoCambioReemplazoXML)
                
                '*** Cabecera ***
                .CommandText = "{ call up_PRManParticipeContratoFondoXML('" & _
                    strCodParticipe & "','" & _
                    strParticipeContratoFondoXML & "','" & _
                    IIf(strEstado = Reg_Adicion, "I", "U") & "') }"
                adoConn.Execute .CommandText
                
                                                                                
            End With
            
            Set adoRegistroAux = Nothing
                
            Me.MousePointer = vbDefault
        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabContrato
                .TabEnabled(0) = True
                .Tab = 0
            End With
            
            Call Buscar
        End If
    End If
    Exit Sub
    
Errorhandler:
'    adoComm.CommandText = "ROLLBACK TRAN ProcAsiento"
'    adoConn.Execute adoComm.CommandText
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
            
End Sub

Public Sub Adicionar()

    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar asiento..."
    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    
    With tabContrato
        .TabEnabled(0) = False
        .Tab = 1
    End With
    
    Call Habilita
    
End Sub

Private Sub LlenarFormulario(strModo As String)
        
    Dim intRegistro As Integer
    Dim strSql As String
    
    Select Case strModo
        Case Reg_Adicion
            
            txtCodParticipe.Text = Valor_Caracter
            lblDescripParticipe.Caption = Valor_Caracter
            lblDescripTitular.Caption = Valor_Caracter
            lblTipoDocumento.Caption = Valor_Caracter
            lblNumDocumento.Caption = Valor_Caracter
            
            txtNumDocumento.Text = Valor_Caracter
            
            dtpFechaContrato.Value = gdatFechaActual
            txtHoraContrato.Text = Format(Time, "hh:mm")

            cboEstado.ListIndex = -1
            cboFondo.ListIndex = -1
            cboPromotor.ListIndex = -1
            
            txtDescripComentario.Text = Valor_Caracter
                       
            Call CargarDetalleGrilla
            
                        
        Case Reg_Edicion
            
            Dim adoRecordset As New ADODB.Recordset
                                    
            strSql = "{ call up_PRLstParticipeContratoFondo ('" & tdgConsulta.Columns("CodParticipe") & "')}"
                     
            adoComm.CommandText = strSql
                     
            Set adoRecordset = adoComm.Execute

            If Not adoRecordset.EOF Then
                
                txtCodParticipe.Text = adoRecordset.Fields("CodParticipe")
                lblDescripParticipe.Caption = adoRecordset.Fields("DescripParticipe")
                lblDescripTitular.Caption = adoRecordset.Fields("DescripCliente")
                lblTipoDocumento.Caption = adoRecordset.Fields("DescripTipoDocumento")
                lblNumDocumento.Caption = Trim(adoRecordset.Fields("NumDocumento"))
                
                Call CargarDetalleGrilla
            
                adoRegistroAux.MoveFirst
            
            Else
                MsgBox "El Sistema no puede encontrar el comprobante contable para consultar!", vbExclamation
                Exit Sub
            End If
                                                            
        End Select
               
    
End Sub

Private Sub CargarDetalleGrilla()
    
    Dim adoRegistro As ADODB.Recordset
    Dim adoField As ADODB.Field
    
    Dim strSql As String
    
    Set adoRegistro = New ADODB.Recordset
        
    Call ConfiguraRecordsetAuxiliar
    
    If strEstado = Reg_Edicion Then
        
        strSql = "{ call up_PRLstParticipeContratoFondo ('" & tdgConsulta.Columns("CodParticipe") & "')}"
        
        
        With adoRegistro
        'With adoMovimiento
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
                    'adoMovimiento.MoveNext
                Loop
                adoRegistroAux.MoveFirst
            End If
            
        End With
    
    End If
    
    tdgMovimiento.DataSource = adoRegistroAux
    
    If adoRegistroAux.RecordCount > 0 Then
        strEstado = Reg_Edicion
        Call CargarDetalleParticipeContratoFondo
    End If
            
End Sub
Sub CargarDetalleParticipeContratoFondo()

    dtpFechaContrato.Value = adoRegistroAux.Fields("FechaContrato")
    'txtHoraContrato.Text = adoRegistroAux.Fields("HoraContrato")
    txtNumDocumento.Text = adoRegistroAux.Fields("NumContrato")
    
    txtDescripComentario.Text = adoRegistroAux.Fields("DescripComentario")
    
    intRegistro = ObtenerItemLista(arrFondo(), adoRegistroAux.Fields("CodFondo"))
    If intRegistro >= 0 Then cboFondo.ListIndex = intRegistro
    
    intRegistro = ObtenerItemLista(arrEstado(), adoRegistroAux.Fields("EstadoContrato"))
    If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
    
    intRegistro = ObtenerItemLista(arrPromotor(), adoRegistroAux.Fields("CodPromotor"))
    If intRegistro >= 0 Then cboPromotor.ListIndex = intRegistro

End Sub

Private Function TodoOK() As Boolean

    TodoOK = False
    
    If Trim(txtCodParticipe.Text) = Valor_Caracter Then
        MsgBox "Participe no Ingresado", vbCritical, gstrNombreEmpresa
        txtCodParticipe.SetFocus
        Exit Function
    End If
            
    If Trim(txtNumDocumento.Text) = Valor_Caracter Then
        MsgBox "Número de Contrato no Ingresado", vbCritical, gstrNombreEmpresa
        txtNumDocumento.SetFocus
        Exit Function
    End If
    
    If cboEstado.ListIndex < 0 Then
        MsgBox "No ha ingresado el Estado del Contrato", vbCritical, gstrNombreEmpresa
        cboEstado.SetFocus
        Exit Function
    End If
    
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function
  

 

Private Sub tabContrato_Click(PreviousTab As Integer)

    Select Case tabContrato.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabContrato.Tab = 0
        
    End Select
End Sub
 

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 4 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
    If ColIndex = 5 Then
        Call DarFormatoValor(Value, Decimales_TipoCambio)
    End If
    
End Sub


Private Sub tdgConsulta_HeadClick(ByVal ColIndex As Integer)

    Static numColindex As Integer

    tdgConsulta.Splits(0).Columns(numColindex).HeadingStyle.ForegroundPicture = Null

    Call OrdenarDBGrid(ColIndex, adoConsulta, tdgConsulta)
    
    numColindex = ColIndex

End Sub


Private Sub ConfiguraRecordsetAuxiliar()

    Set adoRegistroAux = New ADODB.Recordset

    With adoRegistroAux
       .CursorLocation = adUseClient
       .Fields.Append "CodParticipe", adVarChar, 20
       .Fields.Append "CodFondo", adVarChar, 3
       .Fields.Append "CodAdministradora", adVarChar, 3
       .Fields.Append "DescripFondo", adVarChar, 50
       .Fields.Append "NumContrato", adVarChar, 15
       .Fields.Append "FechaContrato", adDate, 10
       .Fields.Append "CodPromotor", adVarChar, 8
       .Fields.Append "CodSucursal", adVarChar, 6
       .Fields.Append "CodAgencia", adVarChar, 3
       .Fields.Append "EstadoContrato", adVarChar, 2
       .Fields.Append "DescripComentario", adVarChar, 100
       .LockType = adLockBatchOptimistic
    End With
    
    adoRegistroAux.Open

End Sub

Private Sub tdgMovimiento_DblClick()
    Call CargarDetalleParticipeContratoFondo
End Sub



Private Sub SubImprimir()


    Dim strSeleccionRegistro    As String
    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    
   
            gstrNameRepo = "ContratoParticipeGrilla"
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(1)
            ReDim aReportParamFn(2)
            ReDim aReportParamF(2)
                        
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
            
            If optParticipe(0).Value Then
                
                aReportParamS(0) = "CP"
                aReportParamS(1) = Trim(txtCodParticipeBusqueda.Text)
                
                
            ElseIf optParticipe(2).Value Then
                aReportParamS(0) = "DP"
                aReportParamS(1) = Trim(txtDescripcion.Text)
               
            Else
                aReportParamS(0) = "NI"
                aReportParamS(1) = Trim(txtNumDocumentoBusq.Text)
         
            End If
       
    
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    


End Sub

