VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{442CAE95-1D41-47B1-BE83-6995DA3CE254}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmCambioHorarioFondo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Horas de Atención por Fondo"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   Icon            =   "frmCambioHorarioFondo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   10005
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   8400
      TabIndex        =   3
      Top             =   5280
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
      Left            =   480
      TabIndex        =   4
      Top             =   5280
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   1296
      Buttons         =   5
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Eliminar"
      Tag2            =   "4"
      ToolTipText2    =   "Eliminar"
      Caption3        =   "&Buscar"
      Tag3            =   "5"
      ToolTipText3    =   "Buscar"
      Caption4        =   "&Imprimir"
      Tag4            =   "6"
      ToolTipText4    =   "Imprimir"
      UserControlWidth=   7200
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Agregar detalle"
      Top             =   7020
      Width           =   375
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Quitar detalle"
      Top             =   7620
      Width           =   375
   End
   Begin TabDlg.SSTab tabCatalogo 
      Height          =   5025
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   8864
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
      TabPicture(0)   =   "frmCambioHorarioFondo.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraBusqueda"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmCambioHorarioFondo.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDetalle"
      Tab(1).Control(1)=   "cmdAccion"
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -69480
         TabIndex        =   1
         Top             =   4080
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
      Begin VB.Frame fraBusqueda 
         Caption         =   "Criterios de Búsqueda"
         Height          =   975
         Left            =   390
         TabIndex        =   15
         Top             =   780
         Width           =   7515
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   360
            Width           =   5505
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
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   390
            Width           =   735
         End
      End
      Begin VB.Frame fraDetalle 
         Height          =   3405
         Left            =   -74670
         TabIndex        =   2
         Top             =   630
         Width           =   8055
         Begin VB.ComboBox cboAgencia 
            Height          =   315
            Left            =   2190
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   750
            Width           =   5505
         End
         Begin VB.ComboBox cboSucursal 
            Height          =   315
            Left            =   2190
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   330
            Width           =   5505
         End
         Begin MSComCtl2.DTPicker dtpHoraInicio 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   4
            EndProperty
            Height          =   285
            Left            =   2205
            TabIndex        =   7
            Top             =   1245
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   146669571
            UpDown          =   -1  'True
            CurrentDate     =   38831
         End
         Begin MSComCtl2.DTPicker dtpHoraTermino 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   4
            EndProperty
            Height          =   285
            Left            =   2205
            TabIndex        =   8
            Top             =   1605
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   146669571
            UpDown          =   -1  'True
            CurrentDate     =   38831
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Hora de Inicio"
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
            Left            =   210
            TabIndex        =   12
            Top             =   1260
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Hora de Término"
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
            Left            =   210
            TabIndex        =   11
            Top             =   1635
            Width           =   1425
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Sucursal"
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
            Left            =   240
            TabIndex        =   10
            Top             =   390
            Width           =   750
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Agencia"
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
            Left            =   240
            TabIndex        =   9
            Top             =   750
            Width           =   705
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmCambioHorarioFondo.frx":0044
         Height          =   2595
         Left            =   540
         OleObjectBlob   =   "frmCambioHorarioFondo.frx":005E
         TabIndex        =   18
         Top             =   2040
         Width           =   8895
      End
   End
End
Attribute VB_Name = "frmCambioHorarioFondo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()      As String, arrSucursal()    As String
Dim arrAgencia()    As String

Dim strCodFondo     As String, strCodSucursal   As String
Dim strCodAgencia   As String

Dim strSQL          As String
'Dim adoRegistroAux  As ADODB.Recordset
Dim adoRegistro As ADODB.Recordset
Dim intRegistro     As String
'Dim vntTmp          As Variant
Dim strEstado       As String
Dim indSortAsc      As Boolean, indSortDesc     As Boolean

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
        Case vPrint
            Call SubImprimir
        
    End Select
    
End Sub

Public Sub Adicionar()

    If cboFondo.ListIndex < 0 Then Exit Sub
    
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Datos de Catálogo..."
                    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabCatalogo
        .TabEnabled(0) = False
        .Tab = 1
    End With
    
End Sub

Public Sub Buscar()
   
   
    Set adoRegistro = New ADODB.Recordset
    
'    Set adoRegistroAux = New ADODB.Recordset
'
'
'    With adoRegistroAux
'       .CursorLocation = adUseClient
'       .Fields.Append "CodSucursal", adChar, 3
'       .Fields.Append "DescripSucursal", adVarChar, 25
'       .Fields.Append "CodAgencia", adChar, 6
'       .Fields.Append "DescripAgencia", adVarChar, 50
'       .Fields.Append "HoraInicio", adChar, 5
'       .Fields.Append "HoraTermino", adChar, 5
'       .CursorType = adOpenStatic
'       .LockType = adLockBatchOptimistic
'    End With
'
'    adoRegistroAux.Open

    
    strSQL = "{ call up_ACSelDatosParametro(49,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"

    
    With adoRegistro
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
        '.ActiveConnection = Nothing
        
'        If .RecordCount > 0 Then
'            .MoveFirst
'            Do While Not .EOF
'                adoRegistroAux.AddNew Array("CodSucursal", "DescripSucursal", "CodAgencia", "DescripAgencia", "HoraInicio", "HoraTermino"), Array(.Fields("CodSucursal").Value, .Fields("DescripSucursal").Value, .Fields("CodAgencia").Value, .Fields("DescripAgencia").Value, .Fields("HoraInicio").Value, .Fields("HoraTermino").Value)
'                adoRegistro.MoveNext
'            Loop
'            adoRegistroAux.MoveFirst
'        End If
        
    End With
    
    tdgConsulta.DataSource = adoRegistro
    
    'tdgConsulta.Refresh
    'tdgConsulta.ReBind
    
    If adoRegistro.RecordCount > 0 Then strEstado = Reg_Consulta
    
'    Me.Refresh
'
'    DoEvents
        
End Sub

Private Sub CargarListas()

    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    strSQL = "SELECT CodSucursal CODIGO, DescripSucursal DESCRIP FROM SucursalBancaria"
    CargarControlLista strSQL, cboSucursal, arrSucursal(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
End Sub
Public Sub Cancelar()

    fraDetalle.Enabled = True
    cmdOpcion.Visible = True
    With tabCatalogo
        .TabEnabled(0) = True
        .Tab = 0
    End With
    Call Buscar
    
    'strEstado = Reg_Consulta
    
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

Public Sub Eliminar()


'    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
'        strEstado = Reg_Eliminacion
'        LlenarFormulario strEstado
'        cmdOpcion.Visible = False
'        With tabCatalogo
'            .TabEnabled(0) = False
'            .Tab = 1
'        End With
'    End If
    
    
End Sub


Public Sub Grabar()
    
    Dim intAccion   As Integer, lngNumError     As Integer
    
    On Error GoTo CtrlError
    

    If TodoOK() Then
    
        If strEstado = Reg_Adicion Then
        
            '*** Guardar ***
            With adoComm
                .CommandText = "INSERT INTO FondoHorarioAtencion " & _
                    "(CodFondo, CodAdministradora, CodSucursal, CodAgencia, HoraInicio, HoraTermino) VALUES ('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & strCodSucursal & "'," & _
                    "'" & strCodAgencia & "','" & Format(dtpHoraInicio.Value, "hh:mm") & "','" & Format(dtpHoraTermino.Value, "hh:mm") & "')"
                    
                adoConn.Execute .CommandText
            End With
                                                                                                                        
            Me.MousePointer = vbDefault
                        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabCatalogo
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        
        End If
    
    
        If strEstado = Reg_Edicion Then
        
            '*** Guardar ***
            With adoComm
                .CommandText = "UPDATE FondoHorarioAtencion SET " & _
                    " HoraInicio = '" & Format(dtpHoraInicio.Value, "hh:mm") & "'," & _
                    " HoraTermino = '" & Format(dtpHoraTermino.Value, "hh:mm") & "' WHERE " & _
                    " CodFondo = '" & strCodFondo & "' AND " & _
                    " CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
                    " CodSucursal = '" & strCodSucursal & "' AND " & _
                    " CodAgencia = '" & strCodAgencia & "'"
                
                adoConn.Execute .CommandText
            End With
    
            Me.MousePointer = vbDefault
                        
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabCatalogo
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
            
        End If


        If strEstado = Reg_Eliminacion Then
            If MsgBox(Mensaje_Eliminacion, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                '*** Cambiar de Estado ***
                adoComm.CommandText = "DELETE FondoHorarioAtencion WHERE " & _
                        " CodFondo = '" & strCodFondo & "' AND " & _
                        " CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
                        " CodSucursal = '" & strCodSucursal & "' AND " & _
                        " CodAgencia = '" & strCodAgencia & "'"
                adoConn.Execute adoComm.CommandText
                
                fraDetalle.Enabled = True
                cmdOpcion.Visible = True
                With tabCatalogo
                    .TabEnabled(0) = True
                    .Tab = 0
                End With
                Call Buscar

            End If
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


Public Sub Imprimir()
    
    Call SubImprimir '(1)
    
End Sub

Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabCatalogo
            .TabEnabled(0) = False
            .Tab = 1
        End With
        
    End If
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Private Sub cboAgencia_Click()


    strCodAgencia = Valor_Caracter
    If cboAgencia.ListIndex < 0 Then Exit Sub
    
    strCodAgencia = Trim(arrAgencia(cboAgencia.ListIndex))
    
    
End Sub

Private Sub cboFondo_Click()

    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Call Buscar

End Sub

Private Sub cboSucursal_Click()

    strCodSucursal = Valor_Caracter
    If cboSucursal.ListIndex < 0 Then Exit Sub
    
    strCodSucursal = Trim(arrSucursal(cboSucursal.ListIndex))
    
    strSQL = "SELECT CodAgencia CODIGO, DescripAgencia DESCRIP FROM AgenciaBancaria WHERE CodSucursal = '" & strCodSucursal & "'"
    CargarControlLista strSQL, cboAgencia, arrAgencia(), Valor_Caracter

End Sub

Private Sub cmdAgregar_Click()

    
'    If TodoOK() Then
'        'If Not tdgDinamica.EOF Then
'            If adoRegistroAux.Supports(adAddNew) Then
'                adoRegistroAux.AddNew Array("CodSucursal", "DescripSucursal", "CodAgencia", "DescripAgencia", "HoraInicio", "HoraTermino"), Array(strCodSucursal, cboSucursal.List(cboSucursal.ListIndex), strCodAgencia, cboAgencia.List(cboAgencia.ListIndex), Format(dtpHoraInicio.Value, "hh:mm"), Format(dtpHoraTermino.Value, "hh:mm"))
'            End If
'        'End If
'    End If
'
'
'    tdgDinamica.Refresh
'
   
    
End Sub
Private Function TodoOK() As Boolean

    TodoOK = False
    
    If cboSucursal.ListIndex = -1 Then
        MsgBox "Debe seleccionar una sucursal válida!", vbOKOnly + vbExclamation, Me.Caption
        Exit Function
    End If

    If cboAgencia.ListIndex = -1 Then
        MsgBox "Debe seleccionar una agencia válida!", vbOKOnly + vbExclamation, Me.Caption
        Exit Function
    End If

    TodoOK = True


End Function
        
Private Sub LlenarFormulario(strModo As String)
    
'    With adoRegistroAux
'        If .RecordCount > 0 Then
'            If .EOF Or .BOF Then
'               .MoveFirst
'            End If
'
'            'CodSucursal,CodAgencia,HoraInicio,HoraTermino
'
'            intRegistro = ObtenerItemLista(arrSucursal(), .Fields("CodSucursal").Value)
'            If intRegistro >= 0 Then cboSucursal.ListIndex = intRegistro
'
'            dtpHoraInicio.Value = .Fields("HoraInicio").Value
'            dtpHoraTermino.Value = .Fields("HoraTermino").Value
'
'        End If
'
'    End With


    Select Case strModo
        Case Reg_Adicion
            cboSucursal.ListIndex = -1
            dtpHoraInicio.Value = "00:00"
            dtpHoraTermino.Value = "00:00"
        
        Case Reg_Edicion, Reg_Eliminacion
            intRegistro = ObtenerItemLista(arrSucursal(), tdgConsulta.Columns("CodSucursal").Value)
            If intRegistro >= 0 Then cboSucursal.ListIndex = intRegistro

            intRegistro = ObtenerItemLista(arrAgencia(), tdgConsulta.Columns("CodAgencia").Value)
            If intRegistro >= 0 Then cboAgencia.ListIndex = intRegistro


            dtpHoraInicio.Value = tdgConsulta.Columns("HoraInicio").Value
            dtpHoraTermino.Value = tdgConsulta.Columns("HoraTermino").Value
            
            If strModo = Reg_Eliminacion Then
                fraDetalle.Enabled = False
            End If
               
    End Select

End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabCatalogo.Tab = 0
    
    tabCatalogo.TabEnabled(1) = False

    
    
    '*** Ancho por defecto de las columnas de la grilla ***
'    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 14
'    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 60
'    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 10
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub SubImprimir()

    Dim strSeleccionRegistro    As String
    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()

    If tabCatalogo.Tab = 1 Then Exit Sub
       
        gstrNameRepo = "CambioHorarioFondoGrilla"
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
                    
        aReportParamS(0) = Trim(strCodFondo)
        aReportParamS(1) = Trim(gstrCodAdministradora)

    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal

End Sub

'Private Sub CargarReportes()
'
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Vista Activa"
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
'
'End Sub

Private Sub Form_Activate()

    'Call CargarReportes
    
End Sub

'Private Sub Form_Deactivate()
'
'    Call OcultarReportes
'
'End Sub

Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
'    Call CargarReportes
    Call DarFormato
    
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

'    Call OcultarReportes
    Set frmCambioHorarioFondo = Nothing
    
End Sub

Private Sub fraTipoCambio_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub tabCatalogo_Click(PreviousTab As Integer)

    Select Case tabCatalogo.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabCatalogo.Tab = 0
        
    End Select

End Sub

Private Sub tdgConsulta_DblClick()

    If strEstado <> Reg_Consulta Or tdgConsulta.Bookmark = 0 Then Exit Sub
    Call Accion(vModify)

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

    Call OrdenarDBGrid(ColIndex, adoRegistro, tdgConsulta)
    
    numColindex = ColIndex
    
    '****
    strPrevColumTDB = strColNameTDB
    '***
    
End Sub
