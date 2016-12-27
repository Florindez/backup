VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmFondoComision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comisiones SAFI"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   8250
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   6360
      TabIndex        =   3
      Top             =   6000
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   6000
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   1296
      Buttons         =   3
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Buscar"
      Tag2            =   "5"
      ToolTipText2    =   "Buscar"
      UserControlWidth=   4200
   End
   Begin TabDlg.SSTab tabComision 
      Height          =   5655
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8388608
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
      TabPicture(0)   =   "frmFondoComision.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraComision(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datos"
      TabPicture(1)   =   "frmFondoComision.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(1)=   "adoRango"
      Tab(1).Control(2)=   "fraComision(1)"
      Tab(1).ControlCount=   3
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -70440
         TabIndex        =   12
         Top             =   4800
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   1296
         Buttons         =   2
         Caption0        =   "&Guardar"
         Tag0            =   "2"
         ToolTipText0    =   "Guardar"
         Caption1        =   "&Cancelar"
         Tag1            =   "8"
         ToolTipText1    =   "Cancelar"
         UserControlWidth=   2700
      End
      Begin MSAdodcLib.Adodc adoRango 
         Height          =   330
         Left            =   -73200
         Top             =   4920
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
         Caption         =   "adoRango"
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
      Begin VB.Frame fraComision 
         Caption         =   "Definición Comisión"
         Height          =   4260
         Index           =   1
         Left            =   -74760
         TabIndex        =   16
         Top             =   480
         Width           =   7215
         Begin VB.ComboBox cboTipoComision 
            Height          =   315
            Left            =   2040
            TabIndex        =   28
            Text            =   "cboTipoComision"
            Top             =   840
            Width           =   4815
         End
         Begin VB.ComboBox cboEstado 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   3740
            Width           =   1575
         End
         Begin VB.ComboBox cboCreditoFiscal 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   2520
            Width           =   4820
         End
         Begin MSComCtl2.DTPicker dtpFechaPatrimonio 
            Height          =   285
            Left            =   5400
            TabIndex        =   11
            Top             =   3360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   175898625
            CurrentDate     =   39526
         End
         Begin VB.CheckBox chkRango 
            Caption         =   "Aplicar sobre Patrimonio Neto de Precierre de la fecha:"
            Height          =   255
            Left            =   2040
            TabIndex        =   9
            Top             =   3015
            Width           =   4935
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2040
            TabIndex        =   10
            Top             =   3375
            Width           =   1575
         End
         Begin VB.ComboBox cboValorComision 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2100
            Width           =   4820
         End
         Begin VB.ComboBox cboVariable 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1680
            Width           =   4820
         End
         Begin VB.ComboBox cboClaseComision 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1245
            Width           =   4820
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
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
            Height          =   195
            Index           =   10
            Left            =   360
            TabIndex        =   26
            Top             =   3760
            Width           =   600
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Crédito Fiscal"
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
            Height          =   435
            Index           =   9
            Left            =   360
            TabIndex        =   25
            Top             =   2520
            Width           =   1245
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
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
            Height          =   195
            Index           =   8
            Left            =   4200
            TabIndex        =   24
            Top             =   3405
            Width           =   540
         End
         Begin VB.Label lblDescripFondo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2040
            TabIndex        =   4
            Top             =   360
            Width           =   4820
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   7
            Left            =   360
            TabIndex        =   23
            Top             =   375
            Width           =   540
         End
         Begin VB.Label lblDescrip 
            Caption         =   "%"
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
            Index           =   6
            Left            =   3720
            TabIndex        =   22
            Top             =   3375
            Width           =   255
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
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
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   21
            Top             =   3405
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Valor"
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
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   20
            Top             =   2115
            Width           =   885
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Variable"
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
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   19
            Top             =   1695
            Width           =   705
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
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
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   18
            Top             =   1260
            Width           =   390
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión"
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
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   17
            Top             =   795
            Width           =   765
         End
      End
      Begin VB.Frame fraComision 
         Caption         =   "Criterios de Búsqueda"
         Height          =   1065
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   7185
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   480
            Width           =   3975
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fondo"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   480
            TabIndex        =   15
            Top             =   495
            Width           =   615
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmFondoComision.frx":0038
         Height          =   3495
         Left            =   -74760
         OleObjectBlob   =   "frmFondoComision.frx":0052
         TabIndex        =   1
         Top             =   1680
         Width           =   7215
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta1 
         Bindings        =   "frmFondoComision.frx":47FA
         Height          =   3495
         Left            =   240
         OleObjectBlob   =   "frmFondoComision.frx":4814
         TabIndex        =   29
         Top             =   1680
         Width           =   7215
      End
   End
End
Attribute VB_Name = "frmFondoComision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()          As String, arrTipoComision()            As String
Dim arrClaseComision()  As String, arrValorComision()           As String
Dim arrVariable()       As String, arrOperacion()               As String
Dim arrCreditoFiscal()  As String, arrEstado()                  As String

Dim strCodFondo         As String, strCodTipoComision           As String
Dim strCodClaseComision As String, strCodValorComision          As String
Dim strCodVariable      As String, strIndRango                  As String
Dim strCodOperacion     As String, strIndOperacion              As String
Dim strCodDetalleFile   As String, strCodCreditoFiscal          As String
Dim strEstado           As String, strCodEstado                 As String
Dim adoConsulta         As ADODB.Recordset
Dim indSortAsc          As Boolean, indSortDesc                 As Boolean

Private Sub InicializarComision()

    txtValor.Text = "0": dtpFechaPatrimonio.Value = gdatFechaActual
    
End Sub

Private Sub cboClaseComision_Click()

    strCodClaseComision = Valor_Caracter
    If cboClaseComision.ListIndex < 0 Then Exit Sub
    
    strCodClaseComision = Trim(arrClaseComision(cboClaseComision.ListIndex))
    
    If strCodClaseComision <> Codigo_Tipo_Comision_Fija Then
        Dim strSQL As String
                
        strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPCOV' AND ValorParametro='" & strCodTipoComision & "' ORDER BY DescripParametro"
        CargarControlLista strSQL, cboVariable, arrVariable(), Sel_Defecto
        
        If cboVariable.ListCount > 0 Then cboVariable.ListIndex = 0
        cboVariable.Enabled = True
    Else
        cboVariable.Enabled = False
    End If
    
End Sub


Private Sub cboCreditoFiscal_Click()

    strCodCreditoFiscal = Valor_Caracter
    If cboCreditoFiscal.ListIndex < 0 Then Exit Sub
    
    strCodCreditoFiscal = arrCreditoFiscal(cboCreditoFiscal.ListIndex)
    
End Sub


Private Sub cboEstado_Click()

    strCodEstado = Valor_Caracter
    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strCodEstado = arrEstado(cboEstado.ListIndex)
    
End Sub


Private Sub cboFondo_Click()

    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Call Buscar
    
End Sub

Private Sub cboTipoComision_Click()
    
    Dim strSQL  As String
    
    strCodTipoComision = Valor_Caracter: strCodDetalleFile = Valor_Caracter
    If cboTipoComision.ListIndex < 0 Then Exit Sub
    
    strCodTipoComision = Left(Trim(arrTipoComision(cboTipoComision.ListIndex)), 2)
    strCodDetalleFile = Right(Trim(arrTipoComision(cboTipoComision.ListIndex)), 3)
    
    If strCodTipoComision = "03" Then
        strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPCOR' ORDER BY DescripParametro"
        CargarControlLista strSQL, cboVariable, arrVariable(), Sel_Defecto
    End If
    
End Sub

Private Sub cboValorComision_Click()

    strCodValorComision = Valor_Caracter
    If cboValorComision.ListIndex < 0 Then Exit Sub
    
    strCodValorComision = Trim(arrValorComision(cboValorComision.ListIndex))
    
End Sub


Private Sub cboVariable_Click()

    strCodVariable = Valor_Caracter
    If cboVariable.ListIndex < 0 Then Exit Sub
    
    strCodVariable = Trim(arrVariable(cboVariable.ListIndex))
    
End Sub



Private Sub chkRango_Click()

    If chkRango.Value Then
        strIndRango = Valor_Indicador
'        Call ColorControlHabilitado(dtpFechaPatrimonio)
        dtpFechaPatrimonio.Enabled = True
    Else
        strIndRango = Valor_Caracter
'        Call ColorControlDeshabilitado(dtpFechaPatrimonio)
        dtpFechaPatrimonio.Enabled = False
    End If
            
End Sub

Private Sub Form_Activate()

    Call CargarReportes
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Listado de Comisiones del Fondo"
    
End Sub
Private Sub Form_Deactivate()

    Call OcultarReportes
    
End Sub


Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call DarFormato
    Call Buscar
    
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
        
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
 

End Sub


Private Sub DarFormato()

    Dim intCont As Integer
    Dim elemento As Object
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
    For intCont = 0 To (fraComision.Count - 1)
        Call FormatoMarco(fraComision(intCont))
    Next
    
    For Each elemento In Me.Controls
    
        If TypeOf elemento Is TDBGrid Then
            Call FormatoGrilla(elemento)
        End If
    
    Next
            
End Sub
Private Sub CargarListas()
        
    Dim strSQL  As String
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    '*** Tipo de Comisión SAFI ***
    strSQL = "SELECT (CodComision + CodDetalleFile) CODIGO,DescripComision DESCRIP FROM ComisionEmpresa " & _
        "WHERE CodTipoComision='" & Codigo_Comision_Empresa_Safi & "' AND Estado='" & Estado_Activo & "' ORDER BY DescripComision"
    CargarControlLista strSQL, cboTipoComision, arrTipoComision(), Sel_Defecto
    
    '*** Tipo de Tasa de Interés ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPTAS' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboClaseComision, arrClaseComision(), Valor_Caracter
    
    '*** Tipo de Valor de Comisión ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='VALCOM' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboValorComision, arrValorComision(), Valor_Caracter
    
    '*** Tipo Crédito Fiscal ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='CREFIS' AND ValorParametro='CC' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboCreditoFiscal, arrCreditoFiscal(), Sel_Defecto
    
    '*** Estados ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='ESTREG' AND CodParametro<>'03' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Valor_Caracter
            
End Sub
Private Sub InicializarValores()
                    
    '*** Valores Iniciales ***
    strIndOperacion = Valor_Caracter
    strIndRango = Valor_Caracter
    tabComision.Tab = 0
    tabComision.TabEnabled(1) = False
       
       
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta1.Columns(1).Width = tdgConsulta1.Width * 0.01 * 40
    tdgConsulta1.Columns(2).Width = tdgConsulta1.Width * 0.01 * 20
    tdgConsulta1.Columns(3).Width = tdgConsulta1.Width * 0.01 * 20
    tdgConsulta1.Columns(4).Width = tdgConsulta1.Width * 0.01 * 10
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmFondoComision = Nothing
    
End Sub

Private Sub tabComision_Click(PreviousTab As Integer)

    Select Case tabComision.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabComision.Tab = 0
        
    End Select
    
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
        Case vSave
            Call Grabar
        Case vCancel
            Call Cancelar
        Case vExit
            Call Salir
        
    End Select
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabComision
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .Tab = 0
    End With
    strEstado = Reg_Consulta
    
End Sub

Public Sub Grabar()

    Dim adoRegistro         As ADODB.Recordset
    Dim intAccion           As Integer, lngNumError         As Long
    Dim strFechaGrabar      As String, strCodAnalitica      As String
    Dim strIndEstado        As String
    Dim datFechaFinPeriodo  As Date
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    'On Error GoTo CtrlError
    
    If strEstado = Reg_Adicion Then
        If TodoOK() Then
            Me.MousePointer = vbHourglass
            
            Set adoRegistro = New ADODB.Recordset
            If chkRango.Value Then
                strFechaGrabar = Convertyyyymmdd(dtpFechaPatrimonio.Value)
            Else
                strFechaGrabar = Valor_Fecha
            End If
            
            strIndEstado = Valor_Indicador
            If strCodEstado = Estado_Inactivo Then strIndEstado = Valor_Caracter
            
            '*** Guardar ***
            With adoComm
                '*** Obtener el número de la analítica ***
                .CommandText = "{call up_ACSelDatosParametro(21,'098') }"
                Set adoRegistro = .Execute
                
                If Not adoRegistro.EOF Then
                    strCodAnalitica = Format(CInt(adoRegistro("NumUltimo")) + 1, "00000000")
                End If
                adoRegistro.Close: Set adoRegistro = Nothing
                
                .CommandText = "{ call up_GNManFondoComision('" & strCodFondo & "','" & _
                    gstrCodAdministradora & "','" & strCodTipoComision & "','" & _
                    strCodClaseComision & "','" & strCodValorComision & "','" & _
                    strCodVariable & "','" & strCodOperacion & "',"
                    
                    If strCodValorComision = Codigo_Tipo_Costo_Porcentaje Then
                        .CommandText = .CommandText & CDec(txtValor.Text) & ",0,'"
                    Else
                        .CommandText = .CommandText & "0," & CDec(txtValor.Text) & ",'"
                    End If
                    .CommandText = .CommandText & strIndRango & "','" & Valor_Indicador & "','" & _
                        strIndEstado & "','" & strFechaGrabar & "','" & strCodAnalitica & "','" & _
                        "098','" & strCodDetalleFile & "','" & strCodCreditoFiscal & "','I') }"
                adoConn.Execute .CommandText
                
                '*** Actualizar el número de analítica **
                .CommandText = "UPDATE InversionFile SET NumUltimo = NumUltimo + 1 " & _
                    "WHERE CodFile='098'"
                adoConn.Execute .CommandText
                
                datFechaFinPeriodo = Convertddmmyyyy(Format(Year(gdatFechaActual), "0000") & "1231")
                '*** Generar Periodo de Corte y Pago a la Administradora ***
                frmMainMdi.stbMdi.Panels(3).Text = "Generando Periodo Contable..."
                Call GenerarPeriodoComision(gstrTipoAdministradora, gstrCodAdministradora, strCodFondo, gdatFechaActual, datFechaFinPeriodo, strCodTipoComision, strCodAnalitica, frmMainMdi.stbMdi)
            End With
                                                                                                                
            Me.MousePointer = vbDefault
                        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabComision
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
    End If
    
    If strEstado = Reg_Edicion Then
        If TodoOK() Then
            Me.MousePointer = vbHourglass
            
            If chkRango.Value Then
                strFechaGrabar = Convertyyyymmdd(dtpFechaPatrimonio.Value)
            Else
                strFechaGrabar = Valor_Fecha
            End If
            
            strIndEstado = Valor_Indicador
            If strCodEstado = Estado_Inactivo Then strIndEstado = Valor_Caracter
            
            '*** Actualizar ***
            With adoComm
                .CommandText = "{ call up_GNManFondoComision('" & strCodFondo & "','" & _
                gstrCodAdministradora & "','" & strCodTipoComision & "','" & _
                strCodClaseComision & "','" & strCodValorComision & "','" & _
                strCodVariable & "','" & strCodOperacion & "',"
                
                If strCodValorComision = Codigo_Tipo_Costo_Porcentaje Then
                    .CommandText = .CommandText & CDec(txtValor.Text) & ",0,'"
                Else
                    .CommandText = .CommandText & "0," & CDec(txtValor.Text) & ",'"
                End If
                .CommandText = .CommandText & strIndRango & "','" & Valor_Indicador & "','" & _
                        strIndEstado & "','" & strFechaGrabar & "','" & strCodAnalitica & "','" & _
                        "098','" & strCodDetalleFile & "','" & strCodCreditoFiscal & "','U') }"
                adoConn.Execute .CommandText
            End With

            Me.MousePointer = vbDefault
                        
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabComision
                .TabEnabled(0) = True
                .TabEnabled(1) = False
                .Tab = 0
            End With
            Call Buscar
        End If
    End If
    Exit Sub
                
CtrlError:
    Me.MousePointer = vbDefault
    intAccion = ControlErrores
    Select Case intAccion
        Case 0: Resume
        Case 1: Resume Next
        Case 2: Exit Sub
        Case Else
            lngNumError = err.Number
            err.Raise Number:=lngNumError
            err.Clear
    End Select
                
End Sub

Private Function TodoOK() As Boolean
        
    TodoOK = False
    
    If Trim(strCodFondo) = Valor_Caracter Then
        MsgBox "Debe Seleccionar el Fondo.", vbCritical
        cboFondo.SetFocus
        Exit Function
    End If
            
    If Trim(strCodTipoComision) = Valor_Caracter Then
        MsgBox "Debe seleccionar el tipo de Comisión del Fondo.", vbCritical
        cboTipoComision.SetFocus
        Exit Function
    End If
    
    If Trim(strCodCreditoFiscal) = Valor_Caracter Then
        MsgBox "Debe seleccionar el tipo de Crédito Fiscal.", vbCritical
        cboCreditoFiscal.SetFocus
        Exit Function
    End If
                
    If CDec(txtValor.Text) < 0 Then
        MsgBox "El Valor de la Comisión no puede ser menor que 0", vbCritical
        txtValor.SetFocus
        Exit Function
    End If
        
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function
Private Sub LlenarFormulario(strModo As String)

    Dim adoRecord   As ADODB.Recordset
    Dim intRegistro As Integer
    
    Select Case strModo
        Case Reg_Adicion
            lblDescripFondo.Caption = Trim(cboFondo.Text)
                        
            cboTipoComision.ListIndex = -1
            If cboTipoComision.ListCount > 0 Then cboTipoComision.ListIndex = 0
            
            intRegistro = ObtenerItemLista(arrClaseComision(), Codigo_Tipo_Comision_Fija)
            If intRegistro >= 0 Then cboClaseComision.ListIndex = intRegistro
            
            cboVariable.ListIndex = -1
            If cboVariable.ListCount > 0 Then cboVariable.ListIndex = 0
            
            intRegistro = ObtenerItemLista(arrValorComision(), Codigo_Tipo_Costo_Porcentaje)
            If intRegistro >= 0 Then cboValorComision.ListIndex = intRegistro
            
            cboCreditoFiscal.ListIndex = -1
            If cboCreditoFiscal.ListCount > 0 Then cboCreditoFiscal.ListIndex = 0
                                    
            chkRango.Value = vbChecked
            chkRango.Value = vbUnchecked
            txtValor.Text = "0"
            dtpFechaPatrimonio.Value = gdatFechaActual
            
            intRegistro = ObtenerItemLista(arrEstado(), Estado_Activo)
            If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
            
            cboTipoComision.SetFocus
                        
        Case Reg_Edicion
            Set adoRecord = New ADODB.Recordset
                                    
            adoComm.CommandText = "SELECT * FROM FondoComision " & _
                "WHERE CodComision='" & tdgConsulta1.Columns(0) & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                
            Set adoRecord = adoComm.Execute
            
            If Not adoRecord.EOF Then
                lblDescripFondo.Caption = Trim(cboFondo.Text)
                intRegistro = ObtenerItemLista(arrTipoComision(), adoRecord("CodComision") & adoRecord("CodDetalleFile"))
                If intRegistro >= 0 Then cboTipoComision.ListIndex = intRegistro
                                
                intRegistro = ObtenerItemLista(arrClaseComision(), adoRecord("CodTipoComision"))
                If intRegistro >= 0 Then cboClaseComision.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrVariable(), adoRecord("CodVariable"))
                If intRegistro >= 0 Then cboVariable.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrValorComision(), adoRecord("CodValorcomision"))
                If intRegistro >= 0 Then cboValorComision.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrCreditoFiscal(), adoRecord("CodCreditoFiscal"))
                If intRegistro >= 0 Then cboCreditoFiscal.ListIndex = intRegistro
                
                If strCodValorComision = Codigo_Tipo_Costo_Porcentaje Then
                    txtValor.Text = CStr(adoRecord("PorcenComision"))
                Else
                    txtValor.Text = CStr(adoRecord("MontoComision"))
                End If
                
                dtpFechaPatrimonio.Value = adoRecord("FechaBase")
                chkRango.Value = vbChecked
                chkRango.Value = vbUnchecked
                
                If Trim(adoRecord("IndRango")) = Valor_Indicador Then
                    chkRango.Value = vbChecked
                Else
                    chkRango.Value = vbUnchecked
                End If
                
                If Trim(adoRecord("IndVigente")) = Valor_Indicador Then
                    intRegistro = ObtenerItemLista(arrEstado(), Estado_Activo)
                Else
                    intRegistro = ObtenerItemLista(arrEstado(), Estado_Inactivo)
                End If
                If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
                
            End If
            adoRecord.Close: Set adoRecord = Nothing
    
    End Select
    
End Sub
Public Sub Imprimir()

    Call SubImprimir(1)
    
End Sub


Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()

    If tabComision.Tab = 1 Then Exit Sub

    Select Case Index
        Case 1
            gstrNameRepo = "FondoComision"
                        
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(2)
            ReDim aReportParamFn(3)
            ReDim aReportParamF(3)

            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            aReportParamFn(3) = "Fondo"
            
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
            aReportParamF(3) = Trim(cboFondo.Text)
                        
            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora
            aReportParamS(2) = Codigo_Comision_Empresa_Safi
            
    End Select

    gstrSelFrml = ""
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub


Public Sub Buscar()
        
    Dim strSQL  As String
    Set adoConsulta = New ADODB.Recordset
    
    strSQL = "SELECT FC.CodComision,DescripComision,CodValorComision,CodVariable,PorcenComision," & _
        "MontoComision , IndRango, IndVigente,FechaBase=CASE Convert(char(8),FechaBase,112) WHEN '19000101' THEN NULL ELSE FechaBase END " & _
        "FROM FondoComision FC JOIN ComisionEmpresa COMEMP " & _
        "ON(COMEMP.CodComision=FC.CodComision AND COMEMP.CodTipoComision='" & Codigo_Comision_Empresa_Safi & "') " & _
        "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'" & _
        "ORDER BY DescripComision"
                        
    strEstado = Reg_Defecto
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
        
    tdgConsulta1.DataSource = adoConsulta
    
    If adoConsulta.RecordCount > 0 Then strEstado = Reg_Consulta

End Sub

Public Sub Eliminar()

End Sub
Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabComision
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .Tab = 1
        End With
    End If
        
End Sub
Public Sub Adicionar()
                
    If strCodFondo = Valor_Caracter Then
        MsgBox "No existen fondos definidos...", vbCritical, Me.Caption
        Exit Sub
    End If
    
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Comisión del Fondo..."
                
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabComision
        .TabEnabled(0) = False
        .Tab = 1
    End With
      
End Sub


Private Sub tdgConsulta1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 2 Then
        Call DarFormatoValor(Value, Decimales_Tasa)
    End If
    
End Sub

Private Sub txtNumDias_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N")
    
End Sub

Private Sub txtValor_Change()

    Call FormatoCajaTexto(txtValor, Decimales_Tasa)
    
End Sub


Private Sub txtValor_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtValor, Decimales_Tasa)
    
End Sub

Private Sub tdgConsulta1_HeadClick(ByVal ColIndex As Integer)
    
    Dim strColNameTDB  As String
    Static numColindex As Integer
    Static strPrevColumTDB As String
    '** agregar para que no se raye la seleccion de registro con ordenamiento
    strColNameTDB = tdgConsulta1.Columns(ColIndex).DataField
    
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

    tdgConsulta1.Splits(0).Columns(numColindex).HeadingStyle.ForegroundPicture = Null

    Call OrdenarDBGrid(ColIndex, adoConsulta, tdgConsulta1)
    
    numColindex = ColIndex
    
    '****
    strPrevColumTDB = strColNameTDB
    '***
    
End Sub
