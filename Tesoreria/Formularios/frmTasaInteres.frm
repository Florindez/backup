VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmTasaInteres 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de Tasas de Interés"
   ClientHeight    =   5640
   ClientLeft      =   1500
   ClientTop       =   2445
   ClientWidth     =   8010
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
   ScaleHeight     =   5640
   ScaleWidth      =   8010
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   6240
      TabIndex        =   2
      Top             =   4800
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
      Left            =   480
      TabIndex        =   1
      Top             =   4800
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   1296
      Buttons         =   3
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      Visible0        =   0   'False
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      Visible1        =   0   'False
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Eliminar"
      Tag2            =   "4"
      Visible2        =   0   'False
      ToolTipText2    =   "Eliminar"
      UserControlWidth=   4200
   End
   Begin TabDlg.SSTab tabTasaInteres 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8070
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmTasaInteres.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmTasaInteres.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraDefinicion"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAccion"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   4560
         TabIndex        =   13
         Top             =   3720
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
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmTasaInteres.frx":0038
         Height          =   3255
         Left            =   -74760
         OleObjectBlob   =   "frmTasaInteres.frx":0052
         TabIndex        =   26
         Top             =   600
         Width           =   7215
      End
      Begin VB.Frame fraDefinicion 
         Caption         =   "Definición"
         Height          =   3135
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   7215
         Begin VB.TextBox txtTasaDiaria 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   25
            Top             =   2640
            Width           =   2295
         End
         Begin VB.TextBox txtTasaInteres 
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
            Height          =   285
            Left            =   1560
            TabIndex        =   9
            Top             =   2280
            Width           =   2295
         End
         Begin VB.ComboBox cboTipoTasa 
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1095
            Width           =   2295
         End
         Begin VB.ComboBox cboBaseAnual 
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1890
            Width           =   2295
         End
         Begin VB.ComboBox cboMoneda 
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1500
            Width           =   2295
         End
         Begin VB.ComboBox cboPeriodo 
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   315
            Width           =   2295
         End
         Begin VB.ComboBox cboCapitalizacion 
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   705
            Width           =   2295
         End
         Begin VB.TextBox txtDescripTasa 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4200
            MaxLength       =   40
            TabIndex        =   12
            Top             =   2640
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker dtpFechaHasta 
            Height          =   315
            Left            =   5520
            TabIndex        =   11
            Top             =   1069
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
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
            Format          =   203161601
            CurrentDate     =   38068
         End
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   315
            Left            =   5520
            TabIndex        =   10
            Top             =   677
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
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
            Format          =   203161601
            CurrentDate     =   38068
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Tasa"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   10
            Left            =   240
            TabIndex        =   24
            Top             =   1110
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vigencia"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   9
            Left            =   4200
            TabIndex        =   23
            Top             =   315
            Width           =   2655
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Base Anual"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   8
            Left            =   240
            TabIndex        =   22
            Top             =   1890
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Moneda"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   7
            Left            =   240
            TabIndex        =   21
            Top             =   1500
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Capitalización"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   240
            TabIndex        =   20
            Top             =   735
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Descripción"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   4200
            TabIndex        =   19
            Top             =   2235
            Width           =   2655
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Periodo Tasa"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   240
            TabIndex        =   18
            Top             =   345
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Desde"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   3
            Left            =   4200
            TabIndex        =   17
            Top             =   702
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Hasta"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   4
            Left            =   4200
            TabIndex        =   16
            Top             =   1089
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tasa Interés"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   5
            Left            =   240
            TabIndex        =   15
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tasa Diaria"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   14
            Top             =   2730
            Width           =   990
         End
      End
   End
End
Attribute VB_Name = "frmTasaInteres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrPeriodo()    As String, arrCapitalizacion()  As String
Dim arrCodMoneda()  As String, arrBaseAnual()       As String
Dim arrTipoTasa()   As String

Dim strCodPeriodo   As String, strCodCapitalizacion As String
Dim strCodMoneda    As String, strCodBaseAnual      As String
Dim strCodTipoTasa  As String, strValorPeriodo      As String
Dim strEstado       As String, strCodTasa           As String
Dim strSignoMoneda  As String, strSQL               As String

Dim adoConsulta     As ADODB.Recordset
Dim indSortAsc      As Boolean, indSortDesc         As Boolean

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
Public Sub Abrir()

End Sub

Public Sub Anterior()

End Sub

Public Sub Ayuda()

End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabTasaInteres
        .TabEnabled(0) = True
        .Tab = 0
    End With
    Call Buscar
    
End Sub



Public Sub Exportar()

End Sub

Public Sub Importar()

End Sub


Public Sub Imprimir()

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

Public Sub SubImprimir(index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String

    If tabTasaInteres.Tab = 1 Then Exit Sub
    
    Select Case index
        Case 1
            gstrNameRepo = "TasaInteresBancaria"
                        
            Set frmReporte = New frmVisorReporte
            
            ReDim aReportParamS(0)
            ReDim aReportParamFn(2)
            ReDim aReportParamF(2)
                        
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            
                        
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
                                    
    End Select

    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub

Public Sub Ultimo()

End Sub




Private Sub cboBaseAnual_Click()

    strCodBaseAnual = ""
    If cboBaseAnual.ListIndex < 0 Then Exit Sub
    
    strCodBaseAnual = Trim(arrBaseAnual(cboBaseAnual.ListIndex))
    
    If CDbl(txtTasaInteres.Text) > 0 Then Call CalcularFactorDiario
    
End Sub


Private Sub cboCapitalizacion_Click()

    strCodCapitalizacion = ""
    If cboCapitalizacion.ListIndex < 0 Then Exit Sub
    
    strCodCapitalizacion = Trim(arrCapitalizacion(cboCapitalizacion.ListIndex))
    txtDescripTasa.Text = Trim(txtTasaInteres.Text) & "%" & Space(1) & strSignoMoneda & Space(1) & Trim(cboPeriodo.Text) & "/" & Trim(cboCapitalizacion.Text)
    
End Sub


Private Sub cboMoneda_Click()

    strCodMoneda = Valor_Caracter: strSignoMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrCodMoneda(cboMoneda.ListIndex))
    strSignoMoneda = ObtenerSignoMoneda(strCodMoneda)
    txtDescripTasa.Text = Trim(txtTasaInteres.Text) & "%" & Space(1) & strSignoMoneda & Space(1) & Trim(cboPeriodo.Text) & "/" & Trim(cboCapitalizacion.Text)
    
End Sub


Private Sub cboPeriodo_Click()

    strCodPeriodo = "": strValorPeriodo = ""
    If cboPeriodo.ListIndex < 0 Then Exit Sub
    
    strCodPeriodo = Left(arrPeriodo(cboPeriodo.ListIndex), 2)
    strValorPeriodo = Trim(Right(arrPeriodo(cboPeriodo.ListIndex), 10))
    
    If CDbl(txtTasaInteres.Text) > 0 Then Call CalcularFactorDiario
    txtDescripTasa.Text = Trim(txtTasaInteres.Text) & "%" & Space(1) & strSignoMoneda & Space(1) & Trim(cboPeriodo.Text) & "/" & Trim(cboCapitalizacion.Text)
    
End Sub


Private Sub cboTipoTasa_Click()

    strCodTipoTasa = ""
    If cboTipoTasa.ListIndex < 0 Then Exit Sub
    
    strCodTipoTasa = Trim(arrTipoTasa(cboTipoTasa.ListIndex))
    
End Sub


Private Sub Form_Activate()

    Call CargarReportes
    
End Sub

Private Sub Form_Deactivate()

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


Private Sub CargarListas()
                  
    '*** Tipo de Periodo ***
    strSQL = "SELECT (CodParametro + ValorParametro) CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' AND NumValidacion=2 ORDER BY DescripParametro"
    CargarControlLista strSQL, cboPeriodo, arrPeriodo(), Sel_Defecto
    
    '*** Tipo de Capitalización ***
    strSQL = "{ call up_ACSelDatos(17) }"
    CargarControlLista strSQL, cboCapitalizacion, arrCapitalizacion(), Sel_Defecto
            
    '*** Tipo de Tasa ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='NATTAS' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoTasa, arrTipoTasa(), Sel_Defecto
    
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrCodMoneda(), Sel_Defecto
    
    '*** Base Anual ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='BASANU' AND NumValidacion=1 ORDER BY DescripParametro"
    CargarControlLista strSQL, cboBaseAnual, arrBaseAnual(), Sel_Defecto
    
    '*** Abono Intereses ***
    strSQL = ""
    'CargarControlLista strSQL, cboAbono, arrAbono(), Sel_Defecto
        
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Listado de Tasas"
    
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
    Call FormatoEtiqueta(lblDescrip(9), vbCenter)
    Call FormatoEtiqueta(lblDescrip(0), vbCenter)
    
End Sub

Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabTasaInteres.Tab = 0
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 62
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub
Public Sub Eliminar()

    Dim strSQL As String
    Dim intRes As Integer
   
    

   If strEstado = "EDICION" Or strEstado = "CONSULTA" Then
        If MsgBox("Desea Eliminar Registro?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                
            strCodTasa = tdgConsulta.Columns(0).Value
                
            adoComm.CommandText = "UPDATE TasaInteresBancariaDetalle SET IndVigente='' WHERE CodTasa='" & strCodTasa & "'"
            adoConn.Execute adoComm.CommandText
            
            
            MsgBox "La Operación ha sido Exitosa.", vbExclamation, "Observación"

            Call Buscar
        End If
    End If
                             
End Sub

Public Sub Grabar()
    
    Dim adoRecord As ADODB.Recordset
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    If strEstado = Reg_Adicion Then
        Set adoRecord = New ADODB.Recordset
            
        adoComm.CommandText = "{ call up_ACSelDatos(19) }"
        Set adoRecord = adoComm.Execute
        
        If Not adoRecord.EOF Then
            If IsNull(adoRecord("NumTasa")) Then
                strCodTasa = "001"
            Else
                strCodTasa = Format(adoRecord("NumTasa") + 1, "000")
            End If
        Else
            strCodTasa = "001"
        End If
        adoRecord.Close: Set adoRecord = Nothing
            
        If TodoOk() Then
            Me.MousePointer = vbHourglass
            '*** Guardar Tasa Interés ***
            With adoComm
                .CommandText = "{ call up_TSManTasaBancaria('"
                .CommandText = .CommandText & strCodTasa & "','"
                .CommandText = .CommandText & Trim(txtDescripTasa.Text) & "','"
                .CommandText = .CommandText & strCodBaseAnual & "','"
                .CommandText = .CommandText & strCodTipoTasa & "','"
                .CommandText = .CommandText & strCodMoneda & "','"
                .CommandText = .CommandText & strCodPeriodo & "','"
                .CommandText = .CommandText & strCodCapitalizacion & "','"
                .CommandText = .CommandText & "01','"
                .CommandText = .CommandText & "X','"
                .CommandText = .CommandText & Convertyyyymmdd(dtpFechaDesde.Value) & "','"
                .CommandText = .CommandText & Convertyyyymmdd(dtpFechaHasta.Value) & "',"
                .CommandText = .CommandText & CDec(txtTasaInteres.Text) & ","
                .CommandText = .CommandText & CDec(txtTasaDiaria.Text) & ",'"
                .CommandText = .CommandText & "X','"
                .CommandText = .CommandText & "I') }"
                
                adoConn.Execute .CommandText
                
            End With
                                                                                    
            Me.MousePointer = vbDefault
        
            MsgBox "Adición ha sido realizada exitosamente.", vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabTasaInteres
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
    End If

    If strEstado = Reg_Edicion Then
        If TodoOk() Then
        
            Me.MousePointer = vbHourglass
            '*** Guardar Tasa Interés ***
            With adoComm
                .CommandText = "{ call up_TSManTasaBancaria('"
                .CommandText = .CommandText & strCodTasa & "','"
                .CommandText = .CommandText & Trim(txtDescripTasa.Text) & "','"
                .CommandText = .CommandText & strCodBaseAnual & "','"
                .CommandText = .CommandText & strCodTipoTasa & "','"
                .CommandText = .CommandText & strCodMoneda & "','"
                .CommandText = .CommandText & strCodPeriodo & "','"
                .CommandText = .CommandText & strCodCapitalizacion & "','"
                .CommandText = .CommandText & "01','"
                .CommandText = .CommandText & "X','"
                .CommandText = .CommandText & Convertyyyymmdd(dtpFechaDesde.Value) & "','"
                .CommandText = .CommandText & Convertyyyymmdd(dtpFechaHasta.Value) & "',"
                .CommandText = .CommandText & CDec(txtTasaInteres.Text) & ","
                .CommandText = .CommandText & CDec(txtTasaDiaria.Text) & ",'"
                .CommandText = .CommandText & "X','"
                .CommandText = .CommandText & "U') }"
                
                adoConn.Execute .CommandText
                
            End With

            Me.MousePointer = vbDefault
            
            MsgBox "Actualización realizada exitosamente.", vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabTasaInteres
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
    End If
            
End Sub


Private Function TodoOk() As Boolean
        
    TodoOk = False
    
    If Trim(txtDescripTasa.Text) = "" Then
        MsgBox "Descripción no Ingresada", vbCritical, gstrNombreEmpresa
        txtDescripTasa.SetFocus
        Exit Function
    End If
    
    If cboPeriodo.ListIndex = 0 Then
        MsgBox "Seleccione el Periodo", vbCritical, gstrNombreEmpresa
        cboPeriodo.SetFocus
        Exit Function
    End If
    
    If cboCapitalizacion.ListIndex = 0 Then
        MsgBox "Seleccione la Capitalización", vbCritical, gstrNombreEmpresa
        cboCapitalizacion.SetFocus
        Exit Function
    End If
    
    If cboTipoTasa.ListIndex = 0 Then
        MsgBox "Seleccione el tipo de tasa", vbCritical, gstrNombreEmpresa
        cboTipoTasa.SetFocus
        Exit Function
    End If
    
    If cboMoneda.ListIndex = 0 Then
        MsgBox "Seleccione la moneda", vbCritical, gstrNombreEmpresa
        cboMoneda.SetFocus
        Exit Function
    End If
    
    If cboBaseAnual.ListIndex = 0 Then
        MsgBox "Seleccione la base anual", vbCritical, gstrNombreEmpresa
        cboBaseAnual.SetFocus
        Exit Function
    End If
    
    If CDec(txtTasaInteres.Text) = 0 Then
        MsgBox "Tasa de Interés no ingresada", vbCritical, gstrNombreEmpresa
        txtTasaInteres.SetFocus
        Exit Function
    End If

    '*** Si todo paso OK ***
    TodoOk = True
  
End Function

Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabTasaInteres
            .TabEnabled(0) = False
            .Tab = 1
        End With
        Call Habilita
    End If
    
End Sub



Public Sub Adicionar()

    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar tasa interés..."
    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabTasaInteres
        .TabEnabled(0) = False
        .Tab = 1
    End With
    Call Habilita
    
End Sub

Public Sub Buscar()

    Dim strSQL As String
    
    strSQL = "{ call up_ACSelDatos(18) }"
                            
    Set adoConsulta = New ADODB.Recordset
                            
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

Private Sub Deshabilita()

    fraDefinicion.Enabled = False
    
End Sub

Private Sub Habilita()

    fraDefinicion.Enabled = True
    
End Sub



Private Sub CalcularFactorDiario()

    Dim dblTasa         As Double
    Dim dblFactorDiario As Double
    Dim intDias         As Integer

    If strValorPeriodo <> "" Then
        dblTasa = CDbl(txtTasaInteres.Text)
        intDias = CInt(strValorPeriodo)
    
        If strCodPeriodo = Codigo_Frecuencia_Anual Then
            If strCodBaseAnual = Codigo_Base_Actual_Actual Or strCodBaseAnual = Codigo_Base_Actual_365 Or strCodBaseAnual = Codigo_Base_30_365 Then intDias = 365
        End If
    
        dblFactorDiario = ((1 + dblTasa / 100) ^ (1 / intDias) - 1) * 100
        txtTasaDiaria.Text = CStr(dblFactorDiario)
    End If
    
End Sub









Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmTasaInteres = Nothing
    
End Sub

Private Sub tabTasaInteres_Click(PreviousTab As Integer)

    Select Case tabTasaInteres.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabTasaInteres.Tab = 0
                                
    End Select
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRecord As ADODB.Recordset
    Dim strSQL As String
    
    Select Case strModo
        Case Reg_Adicion
            txtDescripTasa.Text = ""
            txtTasaInteres.Text = "0"
            txtTasaDiaria.Text = "0"
            
            cboPeriodo.ListIndex = -1
            If cboPeriodo.ListCount > 0 Then cboPeriodo.ListIndex = 0
            
            cboCapitalizacion.ListIndex = -1
            If cboCapitalizacion.ListCount > 0 Then cboCapitalizacion.ListIndex = 0
            
            cboTipoTasa.ListIndex = -1
            If cboTipoTasa.ListCount > 0 Then cboTipoTasa.ListIndex = 0
            
            cboMoneda.ListIndex = -1
            If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = 0
            
            cboBaseAnual.ListIndex = -1
            If cboBaseAnual.ListCount > 0 Then cboBaseAnual.ListIndex = 0
            
            dtpFechaDesde.Value = gdatFechaActual
            dtpFechaHasta.Enabled = False
            dtpFechaHasta.Value = dtpFechaDesde.Value
                        
            cboPeriodo.SetFocus
                        
        Case Reg_Edicion
            Dim intRegistro As Integer
            Dim adoTemporal As ADODB.Recordset
            
            Set adoRecord = New ADODB.Recordset
            
            strCodTasa = Trim(tdgConsulta.Columns(0))
            
            adoComm.CommandText = "{ call up_ACSelDatosParametro(20,'" & strCodTasa & "') }"
            Set adoRecord = adoComm.Execute
            
            If Not adoRecord.EOF Then
                txtDescripTasa.Text = adoRecord("DescripTasa")
                txtTasaInteres.Text = CDbl(adoRecord("ValorTasa"))
                txtTasaDiaria.Text = CDbl(adoRecord("FactorDiario"))
                
                Set adoTemporal = New ADODB.Recordset
                
                adoComm.CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' AND CodParametro='" & adoRecord("PeriodoTasa") & "'"
                Set adoTemporal = adoComm.Execute
                
                If Not adoTemporal.EOF Then
                    intRegistro = ObtenerItemLista(arrPeriodo(), adoRecord("PeriodoTasa") & adoTemporal("ValorParametro"))
                    If intRegistro > 0 Then cboPeriodo.ListIndex = intRegistro
                End If
                adoTemporal.Close: Set adoTemporal = Nothing
                                                
                intRegistro = ObtenerItemLista(arrCapitalizacion(), adoRecord("PeriodoCapitalizacion"))
                If intRegistro > 0 Then cboCapitalizacion.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrTipoTasa(), adoRecord("TipoTasa"))
                If intRegistro > 0 Then cboTipoTasa.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrCodMoneda(), adoRecord("CodMoneda"))
                If intRegistro > 0 Then cboMoneda.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrBaseAnual(), adoRecord("BaseAnual"))
                If intRegistro > 0 Then cboBaseAnual.ListIndex = intRegistro
                                
                dtpFechaDesde.Value = adoRecord("FechaInicio")
                dtpFechaHasta.Enabled = True
                If adoRecord("FechaInicio") = Valor_Fecha Then
                    dtpFechaHasta.Value = dtpFechaDesde.Value
                Else
                    dtpFechaHasta.Value = adoRecord("FechaFinal")
                End If
            End If
            adoRecord.Close: Set adoRecord = Nothing
    
    End Select
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 2 Then
        Call DarFormatoValor(Value, Decimales_Tasa)
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

Private Sub txtTasaDiaria_Change()

    Call FormatoCajaTexto(txtTasaDiaria, Decimales_TasaDiaria)
    
End Sub

Private Sub txtTasaDiaria_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTasaDiaria, Decimales_TasaDiaria)
    
End Sub



Private Sub txtTasaInteres_Change()

    Call FormatoCajaTexto(txtTasaInteres, Decimales_Tasa)
    Call CalcularFactorDiario
    txtDescripTasa.Text = Trim(txtTasaInteres.Text) & "%" & Space(1) & strSignoMoneda & Space(1) & Trim(cboPeriodo.Text) & "/" & Trim(cboCapitalizacion.Text)
    
End Sub


Private Sub txtTasaInteres_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTasaInteres, Decimales_Tasa)
    
End Sub


Private Sub txtTasaInteres_LostFocus()

    Call CalcularFactorDiario
    
End Sub


