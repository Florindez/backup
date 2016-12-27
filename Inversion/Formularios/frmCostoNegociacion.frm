VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmCostoNegociacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Costos de Negociación"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   11070
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   480
      TabIndex        =   35
      Top             =   6240
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   1296
      Buttons         =   4
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
      UserControlWidth=   5700
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   9240
      TabIndex        =   33
      Top             =   6240
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TabDlg.SSTab tabCostos 
      Height          =   6075
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   10716
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
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
      TabPicture(0)   =   "frmCostoNegociacion.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCriterio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmCostoNegociacion.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraCostos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAccion"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   7320
         TabIndex        =   34
         Top             =   5190
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
      Begin VB.Frame fraCostos 
         Caption         =   "Comisiones, Retribuciones y Contribuciones"
         Height          =   4575
         Left            =   300
         TabIndex        =   9
         Top             =   540
         Width           =   10365
         Begin VB.ComboBox cboAgente 
            Height          =   315
            Left            =   2670
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1890
            Width           =   5415
         End
         Begin VB.CheckBox chkIndicadorAgente 
            Caption         =   "Indicador Agente"
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
            Height          =   255
            Left            =   390
            TabIndex        =   31
            Top             =   1920
            Width           =   2475
         End
         Begin VB.CheckBox chkIndicadorAfectoImpuesto 
            Caption         =   "Afecto Impuesto"
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
            Height          =   255
            Left            =   360
            TabIndex        =   30
            Top             =   3960
            Width           =   2475
         End
         Begin VB.TextBox txtValorAlterno 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6000
            TabIndex        =   16
            Top             =   3420
            Width           =   2080
         End
         Begin VB.TextBox txtCantDias 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6000
            TabIndex        =   15
            Top             =   2970
            Width           =   2080
         End
         Begin VB.TextBox txtValorCosto 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   14
            Top             =   2970
            Width           =   2080
         End
         Begin VB.ComboBox cboSigno 
            Height          =   315
            Left            =   6000
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2520
            Width           =   2080
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   3390
            Width           =   2080
         End
         Begin VB.ComboBox cboTipoCosto 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2520
            Width           =   2080
         End
         Begin VB.ComboBox cboCosto 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1320
            Width           =   2080
         End
         Begin VB.Label lblTipoOperacion 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6000
            TabIndex        =   29
            Top             =   480
            Width           =   2085
         End
         Begin VB.Label lblMecanismoNegociacion 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1800
            TabIndex        =   28
            Top             =   840
            Width           =   2085
         End
         Begin VB.Label lblTipoValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1800
            TabIndex        =   27
            Top             =   480
            Width           =   2080
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor Alternativo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   4320
            TabIndex        =   26
            Top             =   3435
            Width           =   1155
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cant. Días"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   4320
            TabIndex        =   25
            Top             =   2985
            Width           =   765
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Aplicar"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   4320
            TabIndex        =   24
            Top             =   2535
            Width           =   480
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   360
            TabIndex        =   23
            Top             =   3405
            Width           =   585
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor Costo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   360
            TabIndex        =   22
            Top             =   2985
            Width           =   810
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Valor Costo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   360
            TabIndex        =   21
            Top             =   2535
            Width           =   1170
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Costo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   360
            TabIndex        =   20
            Top             =   1340
            Width           =   405
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Mecanismo Neg."
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   19
            Top             =   860
            Width           =   1200
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Operación"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   4320
            TabIndex        =   18
            Top             =   500
            Width           =   1320
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Valor"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   17
            Top             =   500
            Width           =   945
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmCostoNegociacion.frx":0038
         Height          =   2895
         Left            =   -74700
         OleObjectBlob   =   "frmCostoNegociacion.frx":0052
         TabIndex        =   8
         Top             =   2550
         Width           =   10185
      End
      Begin VB.Frame fraCriterio 
         Caption         =   "Criterios de Búsqueda"
         Height          =   1905
         Left            =   -74700
         TabIndex        =   1
         Top             =   420
         Width           =   10185
         Begin VB.ComboBox cboTipoValor 
            Height          =   315
            Left            =   2895
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   450
            Width           =   5985
         End
         Begin VB.ComboBox cboTipoOperacion 
            Height          =   315
            Left            =   2895
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1275
            Width           =   5985
         End
         Begin VB.ComboBox cboMecanismoNegociacion 
            Height          =   315
            Left            =   2895
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   855
            Width           =   5985
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Valor"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   7
            Top             =   510
            Width           =   945
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Operación"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   6
            Top             =   1335
            Width           =   1320
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Mecanismo Negociación"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   5
            Top             =   900
            Width           =   1755
         End
      End
   End
End
Attribute VB_Name = "frmCostoNegociacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Ficha de Mantenimiento de Costos de Negociación"
Option Explicit

Dim arrMecanismoNegociacion()   As String, arrTipoOperacion()   As String
Dim arrTipoValor()              As String, arrMoneda()          As String
Dim arrTipoCosto()              As String, arrCosto()           As String
Dim arrSigno()                  As String, arrAgente()          As String

Dim strCodMecanismoNegociacion  As String, strCodTipoOperacion  As String
Dim strCodTipoValor             As String, strCodValor          As String
Dim strCodMoneda                As String, strCodSigno          As String
Dim strCodTipoCosto             As String, strCodCosto          As String
Dim strEstado                   As String, strSQL               As String
Dim strCodAgente                As String, strIndAgente         As String
Dim adoConsulta                 As ADODB.Recordset
Dim indSortAsc                  As Boolean, indSortDesc         As Boolean

Public Sub Adicionar()

    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Costo..."
                
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabCostos
        .TabEnabled(0) = False
        .TabEnabled(1) = True
        .Tab = 1
    End With
    
End Sub

Private Sub LlenarFormulario(strModo As String)
    
    Dim intRegistro As Integer
    
    Select Case strModo
        Case Reg_Adicion
        

            lblTipoValor.Caption = Trim(cboTipoValor.Text)
            lblMecanismoNegociacion.Caption = Trim(cboMecanismoNegociacion.Text)
            lblTipoOperacion.Caption = Trim(cboTipoOperacion.Text)
            
            cboCosto.ListIndex = -1
            If cboCosto.ListCount > 0 Then cboCosto.ListIndex = 0
            cboCosto.Enabled = True

            intRegistro = ObtenerItemLista(arrTipoCosto(), Codigo_Tipo_Costo_Porcentaje)
            If intRegistro >= 0 Then cboTipoCosto.ListIndex = intRegistro
            
            cboMoneda.ListIndex = -1
            If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = 0

            cboAgente.ListIndex = -1
            If cboAgente.ListCount > 0 Then cboAgente.ListIndex = 0
            cboAgente.Enabled = False

            cboSigno.ListIndex = -1
            If cboSigno.ListCount > 0 Then cboSigno.ListIndex = 0

            txtValorCosto.Text = "0"
            txtCantDias.Text = "0"
            txtValorAlterno.Text = "0"
            
            chkIndicadorAfectoImpuesto.Value = vbUnchecked
            chkIndicadorAgente.Enabled = True
            
            
            cboCosto.SetFocus
                        
        Case Reg_Edicion
            lblTipoValor.Caption = Trim(cboTipoValor.Text)
            lblMecanismoNegociacion.Caption = Trim(cboMecanismoNegociacion.Text)
            lblTipoOperacion.Caption = Trim(cboTipoOperacion.Text)
            
            If cboCosto.ListCount > 0 Then cboCosto.ListIndex = 0
            intRegistro = ObtenerItemLista(arrCosto(), tdgConsulta.Columns("CodCosto").Value)
            If intRegistro >= 0 Then cboCosto.ListIndex = intRegistro
            cboCosto.Enabled = False
            
            If cboAgente.ListCount > 0 Then cboAgente.ListIndex = 0
            
            If tdgConsulta.Columns("IndAgente").Value = Valor_Indicador Then
                intRegistro = ObtenerItemLista(arrAgente(), tdgConsulta.Columns("CodAgente").Value)
                If intRegistro >= 0 Then cboAgente.ListIndex = intRegistro
                chkIndicadorAgente.Value = vbChecked
                cboAgente.Enabled = False
                chkIndicadorAgente.Enabled = False
            Else
                chkIndicadorAgente.Value = vbUnchecked
                cboAgente.Enabled = True
                chkIndicadorAgente.Enabled = True
            End If
            
            If cboTipoCosto.ListCount > 0 Then cboTipoCosto.ListIndex = 0
            intRegistro = ObtenerItemLista(arrTipoCosto(), tdgConsulta.Columns("TipoCosto").Value)
            If intRegistro >= 0 Then cboTipoCosto.ListIndex = intRegistro
            
            txtValorCosto.Text = CStr(tdgConsulta.Columns("ValorCosto").Value)
            
            If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = 0
            intRegistro = ObtenerItemLista(arrMoneda(), tdgConsulta.Columns("CodMoneda").Value)
            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
            
            If cboSigno.ListCount > 0 Then cboSigno.ListIndex = 0
            intRegistro = ObtenerItemLista(arrSigno(), tdgConsulta.Columns("SignoRestriccion").Value)
            If intRegistro >= 0 Then cboSigno.ListIndex = intRegistro
            
            txtCantDias.Text = CStr(tdgConsulta.Columns("CantDias").Value)
            txtValorAlterno.Text = CStr(tdgConsulta.Columns("ValorAlterno").Value)
            
            If tdgConsulta.Columns("IndAfectoImpuesto").Value = Valor_Indicador Then
                chkIndicadorAfectoImpuesto.Value = vbChecked
            Else
                chkIndicadorAfectoImpuesto.Value = vbUnchecked
            End If
    
    End Select
    
End Sub
Public Sub Buscar()

    Set adoConsulta = New ADODB.Recordset

    strSQL = "SELECT AP.DescripParametro DescripCosto,AP0.DescripParametro DescripTipoCosto," & _
        "ValorCosto,MON.DescripMoneda,CodCosto,TipoCosto,TipoOperacion,TipoValor," & _
        "TipoPlazo,CodAgente, DescripPersona AS DescripAgente, SignoRestriccion,CantDias,ValorAlterno,CN.CodMoneda, IndAfectoImpuesto, IndAgente " & _
        "FROM CostoNegociacion CN JOIN AuxiliarParametro AP ON(AP.CodParametro=CN.CodCosto AND AP.CodTipoParametro='TIPCOM') " & _
        "JOIN AuxiliarParametro AP0 ON(AP0.CodParametro=CN.TipoCosto AND AP0.CodTipoParametro='VALCOM') " & _
        "LEFT JOIN Moneda MON ON(MON.CodMoneda=CN.CodMoneda) " & _
        "LEFT JOIN InstitucionPersona IP ON(IP.CodPersona=CN.CodAgente AND IP.TipoPersona = '" & Codigo_Tipo_Persona_Agente & "') " & _
        "WHERE TipoValor='" & strCodTipoValor & "' AND TipoOperacion='" & strCodMecanismoNegociacion & "' AND " & _
        "TipoPlazo='" & strCodTipoOperacion & "'"
        
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

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabCostos
        .TabEnabled(0) = True
        .Tab = 0
        .TabEnabled(1) = False
    End With
    Call Buscar
    
End Sub

Private Sub CargarListas()

    '*** Tipo de Valores ***
    strSQL = "SELECT (CodParametro + ValorParametro) CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPVAL' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoValor, arrTipoValor(), Valor_Caracter
    
    If cboTipoValor.ListCount > 0 Then cboTipoValor.ListIndex = 0
    
    '*** Tipo de Plazo de Operación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='PLZOPE' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoOperacion, arrTipoOperacion(), Valor_Caracter
    
    If cboTipoOperacion.ListCount > 0 Then cboTipoOperacion.ListIndex = 0

    '*** Tipo de Costo ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPCOM' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboCosto, arrCosto(), Sel_Defecto
    
    If cboCosto.ListCount > 0 Then cboCosto.ListIndex = 0
    
    '*** Agentes ***
    strSQL = "SELECT CodPersona CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona = '" & Codigo_Tipo_Persona_Agente & "' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboAgente, arrAgente(), Sel_Defecto
    
    If cboAgente.ListCount > 0 Then cboAgente.ListIndex = 0
    
    '*** Tipo de Valor de Comisión ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='VALCOM' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoCosto, arrTipoCosto(), Valor_Caracter
            
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Sel_Todos
    
    If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = 0
    
    '*** Signos aplicables a restricción ***
    strSQL = "SELECT CodParametro CODIGO,ValorParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='SIGAPL'"
    CargarControlLista strSQL, cboSigno, arrSigno(), Sel_Defecto
    
    If cboSigno.ListCount > 0 Then cboSigno.ListIndex = 0
        
End Sub

Public Sub Eliminar()

    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        Dim strMensaje  As String
        
        strMensaje = "Se procederá a eliminar el Costo " & tdgConsulta.Columns(0) & vbNewLine & vbNewLine & vbNewLine & "¿ Seguro de continuar ?"
        
        If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                
            '*** Anular Costo ***
            adoComm.CommandText = "DELETE CostoNegociacion " & _
                "WHERE TipoValor='" & strCodTipoValor & "' AND TipoOperacion='" & strCodMecanismoNegociacion & "' AND " & _
                "TipoPlazo='" & strCodTipoOperacion & "' AND CodCosto='" & Trim(tdgConsulta.Columns(6).Value) & "'"
            adoConn.Execute adoComm.CommandText
            
            MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation, Me.Caption
            
            tabCostos.TabEnabled(0) = True
            tabCostos.Tab = 0
            Call Buscar
            
            Exit Sub
        End If
    End If
    
End Sub

Public Sub Grabar()
            
    Dim strIndAfectoImpuesto As String
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    'On Error GoTo Ctrl_Error
    
    strIndAfectoImpuesto = Valor_Caracter
    
    If strEstado = Reg_Adicion Or strEstado = Reg_Edicion Then
        If TodoOK() Then
        
            If chkIndicadorAfectoImpuesto.Value = vbChecked Then
                strIndAfectoImpuesto = Valor_Indicador
            Else
                strIndAfectoImpuesto = Valor_Caracter
            End If
        
            If chkIndicadorAgente.Value = vbChecked Then
                strIndAgente = Valor_Indicador
            Else
                strIndAgente = Valor_Caracter
            End If
        
            With adoComm
                .CommandText = "{call up_TEManCostoNegociacion ('" & _
                    strCodCosto & "','" & strCodTipoCosto & "','" & _
                    strCodMecanismoNegociacion & "','" & strCodMoneda & "','" & _
                    strCodTipoValor & "','" & strCodTipoOperacion & "','" & _
                    strCodAgente & "'," & CDec(txtValorCosto.Text) & ",'" & strCodSigno & "'," & _
                    CInt(txtCantDias.Text) & "," & CDec(txtValorAlterno.Text) & ",'" & strIndAfectoImpuesto & "','" & _
                    strIndAgente & "','" & IIf(strEstado = Reg_Adicion, "I", "U") & "')}"
                adoConn.Execute .CommandText
            End With
            
            Me.MousePointer = vbDefault
        
            If strEstado = Reg_Adicion Then
                MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            Else
                MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            End If
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabCostos
                .TabEnabled(0) = True
                .Tab = 0
            End With
            
            Call cboTipoOperacion_Click
        End If
    End If
    
    Exit Sub
    
Ctrl_Error:
    Select Case err.Number
        Case -2147217873: MsgBox Mensaje_Registro_Duplicado, vbCritical
        Case Else: MsgBox Mensaje_Error_Inesperado, vbCritical
    End Select
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
    
End Sub

Public Sub Imprimir()
    
End Sub

Public Sub Modificar()

    If strEstado = Reg_Defecto Then Exit Sub
    
    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabCostos
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .Tab = 1
        End With
        'Call Habilita
    End If
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim strSeleccionRegistro    As String

    If tabCostos.Tab = 1 Then Exit Sub
    
    gstrNameRepo = "CostoNegociacion"
    
    Select Case Index
        Case 1
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
                        
            aReportParamS(0) = strCodTipoValor
            aReportParamS(1) = strCodMecanismoNegociacion
            aReportParamS(2) = strCodTipoOperacion
        Case 2
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
                        
            aReportParamS(0) = Valor_Caracter
            aReportParamS(1) = Valor_Caracter
            aReportParamS(2) = Valor_Caracter
            
    End Select
    
    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub

Private Sub cboAgente_Click()

    strCodAgente = Valor_Caracter
    If cboAgente.ListIndex < 0 Then Exit Sub
        
    strCodAgente = Trim(arrAgente(cboAgente.ListIndex))

End Sub

Private Sub cboCosto_Click()

    strCodCosto = Valor_Caracter
    If cboCosto.ListIndex < 0 Then Exit Sub
        
    strCodCosto = Trim(arrCosto(cboCosto.ListIndex))
    
End Sub


Private Sub cboMecanismoNegociacion_Click()

    strCodMecanismoNegociacion = Valor_Caracter
    If cboMecanismoNegociacion.ListIndex < 0 Then Exit Sub
    
    strCodMecanismoNegociacion = Trim(arrMecanismoNegociacion(cboMecanismoNegociacion.ListIndex))
    
    Call Buscar
    
End Sub


Private Sub cboMoneda_Click()
    
    strCodMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
        
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
End Sub


Private Sub cboSigno_Click()

    strCodSigno = Valor_Caracter
    If cboSigno.ListIndex < 0 Then Exit Sub
        
    strCodSigno = Trim(arrSigno(cboSigno.ListIndex))
    
End Sub

Private Sub cboTipoCosto_Click()

    strCodTipoCosto = Valor_Caracter
    If cboTipoCosto.ListIndex < 0 Then Exit Sub
        
    strCodTipoCosto = Trim(arrTipoCosto(cboTipoCosto.ListIndex))
    
    txtValorCosto_Change
    txtValorAlterno_Change
    
End Sub


Private Sub cboTipoOperacion_Click()

    strCodTipoOperacion = Valor_Caracter
    If cboTipoOperacion.ListIndex < 0 Then Exit Sub
    
    strCodTipoOperacion = Trim(arrTipoOperacion(cboTipoOperacion.ListIndex))
    
    Call Buscar
    
End Sub


Private Sub cboTipoValor_Click()

    Dim adoRecCONS As ADODB.Recordset
    
    strCodTipoValor = Valor_Caracter: strCodValor = Valor_Caracter
    If cboTipoValor.ListIndex < 0 Then Exit Sub
    
    strCodTipoValor = Left(Trim(arrTipoValor(cboTipoValor.ListIndex)), 2)
    strCodValor = Trim(Mid(arrTipoValor(cboTipoValor.ListIndex), 3, 10))
            
    '*** Concepto de Costo ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPCCO' AND ValorParametro='" & strCodValor & "'"
    CargarControlLista strSQL, cboMecanismoNegociacion, arrMecanismoNegociacion(), Valor_Caracter
    
    If cboMecanismoNegociacion.ListCount > 0 Then cboMecanismoNegociacion.ListIndex = 0

    Call Buscar
    
End Sub

Private Sub chkIndicadorAgente_Click()

    If chkIndicadorAgente.Value = vbChecked Then
        cboAgente.Enabled = True
    Else
        cboAgente.Enabled = False
        cboAgente.ListIndex = 0
    End If
    
End Sub

Private Sub Form_Activate()

    frmMainMdi.stbMdi.Panels(3).Text = Me.Caption
   
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
           
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
     
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
Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Vista Activa"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Listado"
    
End Sub
Private Sub InicializarValores()
    
    strEstado = Reg_Defecto
    tabCostos.Tab = 0
    tabCostos.TabEnabled(1) = False
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
                    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set frmCostoNegociacion = Nothing
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub

Private Function TodoOK() As Boolean

    TodoOK = False
    
    If cboTipoValor.ListIndex < 0 Then
        MsgBox "¡ Seleccione el Tipo de Valor !", vbCritical, Me.Caption
        cboTipoValor.SetFocus
        Exit Function
    ElseIf cboMecanismoNegociacion.ListIndex < 0 Then
        MsgBox "¡ Seleccione el Concepto del Costo !", vbCritical, Me.Caption
        cboMecanismoNegociacion.SetFocus
        Exit Function
    ElseIf cboTipoOperacion.ListIndex < 0 Then
        MsgBox "¡ Seleccione el Plazo del Concepto del Costo !", vbCritical, Me.Caption
        cboTipoOperacion.SetFocus
        Exit Function
    End If
        
    '*** Si todo paso OK ***
    TodoOK = True

End Function

Private Sub tabCostos_Click(PreviousTab As Integer)

    Select Case tabCostos.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabCostos.Tab = 0
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

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 2 Then
        Call DarFormatoValor(Value, Decimales_Tasa)
    End If
    
End Sub


Private Sub txtCantDias_Change()

    Call FormatoCajaTexto(txtCantDias, 0)
    
End Sub

Private Sub txtCantDias_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N", txtCantDias, 0)
    
End Sub


Private Sub txtValorAlterno_Change()

    If strCodTipoCosto = Codigo_Tipo_Costo_Porcentaje Then
        Call FormatoCajaTexto(txtValorAlterno, Decimales_Tasa)
    Else
        Call FormatoCajaTexto(txtValorAlterno, Decimales_Monto)
    End If
    
End Sub

Private Sub txtValorAlterno_KeyPress(KeyAscii As Integer)

    If strCodTipoCosto = Codigo_Tipo_Costo_Porcentaje Then
        Call ValidaCajaTexto(KeyAscii, "M", txtValorAlterno, Decimales_Tasa)
    Else
        Call ValidaCajaTexto(KeyAscii, "M", txtValorAlterno, Decimales_Monto)
    End If
    
End Sub

Private Sub txtValorCosto_Change()

    If strCodTipoCosto = Codigo_Tipo_Costo_Porcentaje Then
        Call FormatoCajaTexto(txtValorCosto, Decimales_Tasa)
    Else
        Call FormatoCajaTexto(txtValorCosto, Decimales_Monto)
    End If
    
End Sub

Private Sub txtValorCosto_KeyPress(KeyAscii As Integer)

    If strCodTipoCosto = Codigo_Tipo_Costo_Porcentaje Then
        Call ValidaCajaTexto(KeyAscii, "M", txtValorCosto, Decimales_Tasa)
    Else
        Call ValidaCajaTexto(KeyAscii, "M", txtValorCosto, Decimales_Monto)
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
