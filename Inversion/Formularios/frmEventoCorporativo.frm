VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmEventoCorporativo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eventos Corporativos"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   10485
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   8640
      TabIndex        =   38
      Top             =   7320
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TabDlg.SSTab tabEvento 
      Height          =   7245
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   12779
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
      TabPicture(0)   =   "frmEventoCorporativo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCriterio"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdgConsulta"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datos Principales"
      TabPicture(1)   =   "frmEventoCorporativo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(1)=   "fraDatosRegistro"
      Tab(1).Control(2)=   "fraDatosBasicos"
      Tab(1).ControlCount=   3
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -68280
         TabIndex        =   37
         Top             =   6280
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
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmEventoCorporativo.frx":0038
         Height          =   4305
         Left            =   300
         OleObjectBlob   =   "frmEventoCorporativo.frx":0052
         TabIndex        =   28
         Top             =   2280
         Width           =   9645
      End
      Begin VB.Frame fraDatosRegistro 
         Height          =   3945
         Left            =   -74760
         TabIndex        =   11
         Top             =   2190
         Width           =   9825
         Begin VB.ComboBox cboTituloReferencia 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   330
            Width           =   7035
         End
         Begin VB.CheckBox chkGeneraConfirmacionAutomatica 
            Caption         =   "Genera Confirmación Automática de Evento Corporativo"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   330
            TabIndex        =   30
            Top             =   2850
            Width           =   5235
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2220
            TabIndex        =   25
            Top             =   2040
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker dtpFechaOperacion 
            Height          =   315
            Left            =   2220
            TabIndex        =   21
            Top             =   750
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   176095233
            CurrentDate     =   38790
         End
         Begin VB.TextBox txtConcepto 
            Height          =   285
            Left            =   2220
            TabIndex        =   12
            Top             =   2400
            Width           =   5415
         End
         Begin MSComCtl2.DTPicker dtpFechaJunta 
            Height          =   285
            Left            =   2220
            TabIndex        =   22
            Top             =   1110
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            Format          =   176095233
            CurrentDate     =   38790
         End
         Begin MSComCtl2.DTPicker dtpFechaCorte 
            Height          =   285
            Left            =   6090
            TabIndex        =   23
            Top             =   1110
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            Format          =   176095233
            CurrentDate     =   38790
         End
         Begin MSComCtl2.DTPicker dtpFechaEntrega 
            Height          =   285
            Left            =   6090
            TabIndex        =   24
            Top             =   1470
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            Format          =   176095233
            CurrentDate     =   38790
         End
         Begin MSComCtl2.DTPicker dtpFechaCalculo 
            Height          =   285
            Left            =   2220
            TabIndex        =   31
            Top             =   1470
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            Format          =   176095233
            CurrentDate     =   38790
         End
         Begin MSComCtl2.DTPicker dtpFechaVencimiento 
            Height          =   315
            Left            =   6090
            TabIndex        =   35
            Top             =   750
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   176095233
            CurrentDate     =   38790
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha Vencimiento"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   4170
            TabIndex        =   36
            Top             =   780
            Width           =   1815
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Valor Referencia"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   14
            Left            =   330
            TabIndex        =   33
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha Cálculo"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   13
            Left            =   330
            TabIndex        =   32
            Top             =   1470
            Width           =   1605
         End
         Begin VB.Label lblDescrip 
            Caption         =   "%"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   12
            Left            =   3825
            TabIndex        =   29
            Top             =   2070
            Width           =   1335
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   3
            X1              =   330
            X2              =   7430
            Y1              =   1875
            Y2              =   1875
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha Operación"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   4
            Left            =   330
            TabIndex        =   20
            Top             =   750
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha Junta"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   5
            Left            =   330
            TabIndex        =   19
            Top             =   1110
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha Corte"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   6
            Left            =   4170
            TabIndex        =   18
            Top             =   1125
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha Entrega"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   7
            Left            =   4170
            TabIndex        =   17
            Top             =   1485
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Derechos"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   8
            Left            =   330
            TabIndex        =   16
            Top             =   2055
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Concepto"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   10
            Left            =   330
            TabIndex        =   15
            Top             =   2415
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Moneda"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   5160
            TabIndex        =   14
            Top             =   2070
            Width           =   645
         End
         Begin VB.Label lblMoneda 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nuevos Soles"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   5820
            TabIndex        =   13
            Top             =   2040
            Width           =   1815
         End
      End
      Begin VB.Frame fraDatosBasicos 
         Height          =   1515
         Left            =   -74760
         TabIndex        =   6
         Top             =   480
         Width           =   9825
         Begin VB.ComboBox cboTitulo 
            Height          =   315
            Left            =   2190
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   420
            Width           =   7005
         End
         Begin VB.ComboBox cboEvento 
            Height          =   315
            Left            =   2190
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   840
            Width           =   3855
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Título"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   10
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Evento"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   9
            Top             =   870
            Width           =   735
         End
      End
      Begin VB.Frame fraCriterio 
         Caption         =   "Criterios de Búsqueda"
         Height          =   1695
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   9705
         Begin VB.ComboBox cboEstado 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1140
            Width           =   2535
         End
         Begin VB.ComboBox cboTituloCriterio 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   360
            Width           =   5295
         End
         Begin VB.ComboBox cboEventoCriterio 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   750
            Width           =   3615
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Estado"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   11
            Left            =   600
            TabIndex        =   26
            Top             =   1160
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Título"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   5
            Top             =   380
            Width           =   735
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Evento"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   4
            Top             =   770
            Width           =   855
         End
      End
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   360
      TabIndex        =   39
      Top             =   7320
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
End
Attribute VB_Name = "frmEventoCorporativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Ficha de Mantenimiento de Eventos Corporativos"
Option Explicit

Dim arrTitulo()             As String, arrTituloReferencia()    As String
Dim arrEvento()             As String
Dim arrTituloCriterio()     As String, arrEventoCriterio()      As String
Dim arrEstado()             As String
Dim strCodTitulo            As String, strCodEvento             As String
Dim strCodTituloCriterio    As String, strCodEventoCriterio     As String
Dim strCodEstado            As String
Dim strEstado               As String, strSQL                   As String

Dim strCodFile              As String, strCodAnalitica          As String
Dim strCodMoneda            As String

Dim strCodFileReferencia    As String, strCodAnaliticaReferencia      As String
Dim strCodMonedaReferencia  As String, strCodTituloReferencia   As String

Dim lngNumAcuerdo           As Long
Dim blnOrdLinea             As Boolean
Dim strCodSignoMoneda       As String

Dim adoConsulta             As ADODB.Recordset
Dim indSortAsc              As Boolean, indSortDesc             As Boolean

Public Sub Imprimir()

End Sub

Public Sub Modificar()

    If strEstado = Reg_Defecto Then Exit Sub
    
    If strCodEstado <> Estado_Acuerdo_Ingresado Then Exit Sub
    
    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabEvento
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .Tab = 1
        End With
        
    End If
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub



Private Sub cboEstado_Click()

    strCodEstado = Valor_Caracter
    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strCodEstado = Trim(arrEstado(cboEstado.ListIndex))
    
End Sub


Private Sub cboEvento_Click()

    strCodEvento = Valor_Caracter
    If cboEvento.ListIndex < 0 Then Exit Sub
    
    strCodEvento = Trim(arrEvento(cboEvento.ListIndex))
    
    If strCodEvento = Codigo_Evento_Dividendo Then
        lblDescrip(12).Caption = strCodSignoMoneda & " por acción"
        lblDescrip(14).Enabled = False
        lblDescrip(15).Enabled = False
        cboTituloReferencia.Enabled = False
        cboTituloReferencia.ListIndex = cboTitulo.ListIndex
        dtpFechaVencimiento.Visible = False
    End If
    
    If strCodEvento = Codigo_Evento_Liberacion Then
        lblDescrip(12).Caption = "%"
        lblDescrip(14).Enabled = False
        lblDescrip(15).Enabled = False
        cboTituloReferencia.Enabled = False
        cboTituloReferencia.ListIndex = cboTitulo.ListIndex
        dtpFechaVencimiento.Visible = False
    End If
    
    If strCodEvento = Codigo_Evento_Nominal Then
        lblDescrip(12).Caption = ""
        lblDescrip(14).Enabled = False
        lblDescrip(15).Enabled = False
        cboTituloReferencia.Enabled = False
        cboTituloReferencia.ListIndex = cboTitulo.ListIndex
        dtpFechaVencimiento.Visible = False
    End If
    
    If strCodEvento = Codigo_Evento_Preferente Then
        lblDescrip(12).Caption = "%"
        lblDescrip(14).Enabled = True
        lblDescrip(15).Enabled = True
        cboTituloReferencia.Enabled = True
        dtpFechaVencimiento.Visible = True
    End If
    
    
End Sub


Private Sub cboEventoCriterio_Click()

    strCodEventoCriterio = Valor_Caracter
    If cboEventoCriterio.ListIndex < 0 Then Exit Sub
    
    strCodEventoCriterio = Trim(arrEventoCriterio(cboEventoCriterio.ListIndex))
    
End Sub


Private Sub cboTitulo_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodTitulo = Valor_Caracter
    If cboTitulo.ListIndex < 0 Then Exit Sub
    
    strCodTitulo = Trim(arrTitulo(cboTitulo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        .CommandText = "SELECT M.CodSigno, M.DescripMoneda, II.CodFile,II.CodAnalitica,II.CodMoneda,II.CodMoneda1 FROM InstrumentoInversion II" & _
            " JOIN Moneda M on (M.CodMoneda = II.CodMoneda)" & _
            " WHERE CodTitulo='" & strCodTitulo & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodFile = Trim(adoRegistro("CodFile"))
            strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
            strCodMoneda = Trim(adoRegistro("CodMoneda"))
            lblMoneda.Caption = Trim(adoRegistro("DescripMoneda"))
            strCodSignoMoneda = Trim(adoRegistro("CodSigno"))
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
            
End Sub


Private Sub cboTituloCriterio_Click()

    strCodTituloCriterio = Valor_Caracter
    If cboTituloCriterio.ListIndex < 0 Then Exit Sub
    
    strCodTituloCriterio = Trim(arrTituloCriterio(cboTituloCriterio.ListIndex))
    
End Sub


Private Sub cboTituloReferencia_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodTituloReferencia = Valor_Caracter
    If cboTituloReferencia.ListIndex < 0 Then Exit Sub
    
    strCodTituloReferencia = Trim(arrTituloReferencia(cboTituloReferencia.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        .CommandText = "SELECT M.CodSigno, M.DescripMoneda, II.CodFile,II.CodAnalitica,II.CodMoneda,II.CodMoneda1 FROM InstrumentoInversion II" & _
            " JOIN Moneda M on (M.CodMoneda = II.CodMoneda)" & _
            " WHERE CodTitulo='" & strCodTituloReferencia & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodFileReferencia = Trim(adoRegistro("CodFile"))
            strCodAnaliticaReferencia = Trim(adoRegistro("CodAnalitica"))
            strCodMonedaReferencia = Trim(adoRegistro("CodMoneda"))
            lblMoneda.Caption = Trim(adoRegistro("DescripMoneda"))
            strCodSignoMoneda = Trim(adoRegistro("CodSigno"))
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With

End Sub

Private Sub Form_Activate()

    frmMainMdi.stbMdi.Panels(3).Text = Me.Caption
    Call CargarReportes
    
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
        Case vPrint
            Call SubImprimir
        Case vCancel
            Call Cancelar
        Case vExit
            Call Salir
        
    End Select
    
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabEvento
        .TabEnabled(0) = True
        .Tab = 0
        .TabEnabled(1) = False
    End With
    Call Buscar
    
End Sub
Public Sub Grabar()

    Dim adoRegistro         As ADODB.Recordset
    Dim strFechaOperacion   As String, strFechaJunta        As String
    Dim strFechaCorte       As String, strFechaEntrega      As String
    Dim strMensaje          As String, strIndGenerado       As String
    Dim intRegistro         As Integer, strFechaCalculo     As String
    Dim strFechaVencimiento As String
    
    Dim strIndConfirmacionAutomatica              As String

    If strEstado = Reg_Consulta Then Exit Sub

    strFechaOperacion = Convertyyyymmdd(dtpFechaOperacion.Value)
    strFechaJunta = Convertyyyymmdd(dtpFechaJunta.Value)
    strFechaCorte = Convertyyyymmdd(dtpFechaCorte.Value)
    strFechaEntrega = Convertyyyymmdd(dtpFechaEntrega.Value)
    strFechaCalculo = Convertyyyymmdd(dtpFechaCalculo.Value)
    strFechaVencimiento = Convertyyyymmdd(dtpFechaVencimiento.Value)

    If chkGeneraConfirmacionAutomatica.Value = vbChecked Then
        strIndConfirmacionAutomatica = Valor_Indicador
    Else
        strIndConfirmacionAutomatica = Valor_Caracter
    End If
    
    If strEstado = Reg_Adicion Then
        If TodoOK() Then

            Me.MousePointer = vbHourglass

            Set adoRegistro = New ADODB.Recordset
            '*** Guardar Acuerdo (Evento) ***
            With adoComm
                .CommandText = "SELECT MAX(NumAcuerdo) NumAcuerdo FROM EventoCorporativoAcuerdo WHERE CodAdministradora='" & gstrCodAdministradora & "'"
                Set adoRegistro = .Execute
                
                If Not adoRegistro.EOF Then
                    If IsNull(adoRegistro("NumAcuerdo")) Then
                        lngNumAcuerdo = 1
                    Else
                        lngNumAcuerdo = CInt(adoRegistro("NumAcuerdo")) + 1
                    End If
                Else
                    lngNumAcuerdo = 1
                End If
                adoRegistro.Close: Set adoRegistro = Nothing

                .CommandText = "BEGIN TRAN ProcAcuerdo"
                adoConn.Execute .CommandText

                On Error GoTo Ctrl_Error

                .CommandText = "{ call up_IVManEventoCorporativoAcuerdo('" & gstrCodAdministradora & "','" & _
                    strCodTitulo & "'," & lngNumAcuerdo & ",'" & strFechaOperacion & "','" & _
                    strCodFile & "','" & strCodAnalitica & "','" & strCodTituloReferencia & "','" & _
                    strCodFileReferencia & "','" & strCodAnaliticaReferencia & "','" & _
                    strFechaJunta & "','" & strFechaCorte & "','" & _
                    strFechaCalculo & "','" & strFechaEntrega & "','" & strFechaVencimiento & "',"
                    
                If strCodEvento = Codigo_Evento_Liberacion Or strCodEvento = Codigo_Evento_Preferente Or strCodEvento = Codigo_Evento_Nominal Then
                    .CommandText = .CommandText & CDec(txtValor.Text) & ",0,0,'"
                ElseIf strCodEvento = Codigo_Evento_Dividendo Then
                    .CommandText = .CommandText & "0," & CDec(txtValor.Text) & ",0,'"
'                Else '*** VN ***
'                    .CommandText = .CommandText & "0,0," & CDec(txtValor.Text) & ",'"
                End If
                .CommandText = .CommandText & strCodMoneda & "','" & strCodEvento & "','" & _
                    Trim(txtConcepto.Text) & "','" & Estado_Acuerdo_Ingresado & "','" & strIndConfirmacionAutomatica & "','" & gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','I') }"
                adoConn.Execute .CommandText
                                                
'        If blnOrdLinea Then
'            GeneraOrdenEvento lngNroAcue, strCodTitulo, strCodAnalitica, strCodEvento, CDbl(txtValor.Text), strFchCorte, blnOrdLinea
'            MsgBox "El Acuerdo se ingresó con fecha menor a la del día... Se GENERO la orden en línea para ser confirmada hoy ", vbExclamation, Me.Caption
'        End If
                                                                                
                .CommandText = "COMMIT TRAN ProcAcuerdo"
                adoConn.Execute .CommandText

            End With

            Me.MousePointer = vbDefault

            MsgBox Mensaje_Adicion_Exitosa, vbExclamation

            frmMainMdi.stbMdi.Panels(3).Text = "Acción"

            cmdOpcion.Visible = True
            With tabEvento
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
    End If
    
    If strEstado = Reg_Edicion Then
        If TodoOK() Then

            Me.MousePointer = vbHourglass

            Set adoRegistro = New ADODB.Recordset
            '*** Guardar Acuerdo (Evento) ***
            With adoComm
                lngNumAcuerdo = CLng(tdgConsulta.Columns(0).Value)
                
                .CommandText = "BEGIN TRAN ProcAcuerdo"
                adoConn.Execute .CommandText

                On Error GoTo Ctrl_Error

                .CommandText = "{ call up_IVManEventoCorporativoAcuerdo('" & gstrCodAdministradora & "','" & _
                    strCodTitulo & "'," & lngNumAcuerdo & ",'" & strFechaOperacion & "','" & _
                    strCodFile & "','" & strCodAnalitica & "','" & strCodTituloReferencia & "','" & _
                    strCodFileReferencia & "','" & strCodAnaliticaReferencia & "','" & _
                    strFechaJunta & "','" & strFechaCorte & "','" & _
                    strFechaCalculo & "','" & strFechaEntrega & "','" & strFechaVencimiento & "',"
                    
                If strCodEvento = Codigo_Evento_Liberacion Or strCodEvento = Codigo_Evento_Preferente Or strCodEvento = Codigo_Evento_Nominal Then
                    .CommandText = .CommandText & CDec(txtValor.Text) & ",0,0,'"
                ElseIf strCodEvento = Codigo_Evento_Dividendo Then
                    .CommandText = .CommandText & "0," & CDec(txtValor.Text) & ",0,'"
'                Else '*** VN ***
'                    .CommandText = .CommandText & "0,0," & CDec(txtValor.Text) & ",'"
                End If
                .CommandText = .CommandText & strCodMoneda & "','" & strCodEvento & "','" & _
                    Trim(txtConcepto.Text) & "','" & Estado_Acuerdo_Ingresado & "','" & strIndConfirmacionAutomatica & "','" & gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','U') }"
                adoConn.Execute .CommandText
                                                
'        If blnOrdLinea Then
'            GeneraOrdenEvento lngNroAcue, strCodTitulo, strCodAnalitica, strCodEvento, CDbl(txtValor.Text), strFchCorte, blnOrdLinea
'            MsgBox "El Acuerdo se ingresó con fecha menor a la del día... Se GENERO la orden en línea para ser confirmada hoy ", vbExclamation, Me.Caption
'        End If
                                                                                
                .CommandText = "COMMIT TRAN ProcAcuerdo"
                adoConn.Execute .CommandText

            End With

            Me.MousePointer = vbDefault

            MsgBox Mensaje_Edicion_Exitosa, vbExclamation

            frmMainMdi.stbMdi.Panels(3).Text = "Acción"

            cmdOpcion.Visible = True
            With tabEvento
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
    End If
    Exit Sub

Ctrl_Error:
    adoComm.CommandText = "ROLLBACK TRAN ProcAcuerdo"
    adoConn.Execute adoComm.CommandText

    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
        
End Sub

Private Function TodoOK() As Boolean
        
    TodoOK = False
          
    If cboTitulo.ListIndex <= 0 Then
        MsgBox "Debe seleccionar el Título.", vbCritical, Me.Caption
        If cboTitulo.Enabled Then cboTitulo.SetFocus
        Exit Function
    End If
    
    If cboEvento.ListIndex <= 0 Then
        MsgBox "Debe seleccionar el Evento.", vbCritical, Me.Caption
        If cboEvento.Enabled Then cboEvento.SetFocus
        Exit Function
    End If
                          
'    If dtpFechaCorte.Value < dtpFechaOperacion.Value Then
'        MsgBox "La Fecha de Corte no puede ser menor a la Fecha de Operación.", vbCritical, Me.Caption
'        If dtpFechaCorte.Enabled Then dtpFechaCorte.SetFocus
'        Exit Function
'    End If
    
    If dtpFechaCorte.Value > dtpFechaEntrega.Value Then
        MsgBox "La Fecha de Corte no puede ser mayor a la Fecha de Entrega.", vbCritical, Me.Caption
        If dtpFechaCorte.Enabled Then dtpFechaCorte.SetFocus
        Exit Function
    End If
    
    If CDec(txtValor.Text) = 0 Then
        MsgBox "Debe indicar el valor del evento.", vbCritical, Me.Caption
        If txtValor.Enabled Then txtValor.SetFocus
        Exit Function
    End If
    
    If Trim(txtConcepto.Text) = Valor_Caracter Then
        MsgBox "Debe indicar el concepto.", vbCritical, Me.Caption
        If txtConcepto.Enabled Then txtConcepto.SetFocus
        Exit Function
    End If
            
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function
Public Sub Eliminar()

    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        If tdgConsulta.Columns(11).Value <> Estado_Acuerdo_Ingresado Then Exit Sub
        If MsgBox("Se procederá a eliminar el acuerdo número" & Space(1) & CStr(tdgConsulta.Columns(0).Value) & _
            Space(1) & "(" & Trim(tdgConsulta.Columns(1).Value) & ")" & vbNewLine & vbNewLine & vbNewLine & _
            "¿ Seguro de continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
    
            '*** Anular Acuerdo ***
            adoComm.CommandText = "UPDATE EventoCorporativoAcuerdo SET EstadoEvento='" & Estado_Acuerdo_Anulado & "' WHERE CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodTitulo='" & tdgConsulta.Columns(10) & "' AND NumAcuerdo='" & tdgConsulta.Columns(0) & "'"
            adoConn.Execute adoComm.CommandText
                                    
            tabEvento.TabEnabled(0) = True
            tabEvento.Tab = 0
            Call Buscar
            
            Exit Sub
        End If
    End If
    
End Sub
Public Sub Adicionar()

    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Evento..."
                
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabEvento
        .TabEnabled(0) = False
        .TabEnabled(1) = True
        .Tab = 1
    End With
    'Call Habilita
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRecord   As ADODB.Recordset
    Dim strSQL      As String
    Dim intRegistro As Integer
    
    Select Case strModo
        Case Reg_Adicion
            cboTitulo.ListIndex = -1
            If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0
            cboTitulo.Enabled = True
                                    
            cboEvento.ListIndex = -1
            If cboEvento.ListCount > 0 Then cboEvento.ListIndex = 0
            cboEvento.Enabled = True
                
            dtpFechaOperacion.Value = gdatFechaActual
            dtpFechaVencimiento.Value = gdatFechaActual
            dtpFechaJunta.Value = dtpFechaOperacion.Value
            dtpFechaCorte.Value = dtpFechaOperacion.Value 'DateAdd("d", 1, dtpFechaOperacion.Value)
            dtpFechaEntrega.Value = dtpFechaCorte.Value 'DateAdd("d", 1, dtpFechaCorte.Value)
            dtpFechaCalculo.Value = dtpFechaCorte.Value
                        
            txtValor.Text = "0"
            txtConcepto.Text = Valor_Caracter
            lblMoneda.Caption = Valor_Caracter
                        
            cboTitulo.SetFocus
                        
        Case Reg_Edicion
            Set adoRecord = New ADODB.Recordset
            
            adoComm.CommandText = "SELECT * FROM EventoCorporativoAcuerdo " & _
                "WHERE NumAcuerdo=" & CInt(tdgConsulta.Columns("NumAcuerdo").Value) & " AND CodTitulo='" & _
                Trim(tdgConsulta.Columns("CodTitulo").Value) & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            Set adoRecord = adoComm.Execute
            
            If Not adoRecord.EOF Then
                If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0
                intRegistro = ObtenerItemLista(arrTitulo(), tdgConsulta.Columns("CodTitulo").Value)
                If intRegistro >= 0 Then cboTitulo.ListIndex = intRegistro
                cboTitulo.Enabled = False
                
                If cboTituloReferencia.ListCount > 0 Then cboTituloReferencia.ListIndex = 0
                intRegistro = ObtenerItemLista(arrTituloReferencia(), tdgConsulta.Columns("CodTituloReferencia").Value)
                If intRegistro >= 0 Then cboTituloReferencia.ListIndex = intRegistro
                cboTituloReferencia.Enabled = False
                
                cboEvento.ListIndex = -1
                If cboEvento.ListCount > 0 Then cboEvento.ListIndex = 0
                intRegistro = ObtenerItemLista(arrEvento(), strCodEventoCriterio)
                If intRegistro >= 0 Then cboEvento.ListIndex = intRegistro
                cboEvento.Enabled = False
                
                If adoRecord("IndGeneraConfirmacionAutomatica") = Valor_Indicador Then
                    chkGeneraConfirmacionAutomatica.Value = vbChecked
                Else
                    chkGeneraConfirmacionAutomatica.Value = vbUnchecked
                End If
                
                dtpFechaVencimiento.Value = adoRecord("FechaVencimiento")
                dtpFechaOperacion.Value = adoRecord("FechaOperacion")
                dtpFechaJunta.Value = adoRecord("FechaJunta")
                dtpFechaCorte.Value = adoRecord("FechaCorte")
                dtpFechaEntrega.Value = adoRecord("FechaEntrega")
                dtpFechaCalculo.Value = adoRecord("FechaCalculo")
                            
                txtValor.Text = tdgConsulta.Columns("ValorEvento").Value
                txtConcepto.Text = Trim(adoRecord("DescripAcuerdo"))
                'lblMoneda.Caption = Valor_Caracter
                            
'                cboTitulo.SetFocus
            End If
    End Select
    
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

Public Sub Buscar()

    Dim strSQL  As String
    
    Set adoConsulta = New ADODB.Recordset
    
    strSQL = "SELECT EC.CodAdministradora,EC.CodTitulo,II.Nemotecnico,NumAcuerdo,FechaOperacion,EC.CodFile,EC.CodAnalitica,EC.CodTituloReferencia,III.Nemotecnico AS NemotecnicoReferencia,EC.CodFileReferencia,EC.CodAnaliticaReferencia, FechaCorte,EC.CodMoneda,EstadoEvento," & _
        "CASE PorcenAccionesLiberadas WHEN 0 THEN (CASE PorcenDividendoEfectivo WHEN 0 THEN EC.ValorNominal ELSE PorcenDividendoEfectivo END) " & _
        "ELSE PorcenAccionesLiberadas END ValorEvento " & _
        "FROM EventoCorporativoAcuerdo EC " & _
        "JOIN InstrumentoInversion II ON(II.CodTitulo=EC.CodTitulo) " & _
        "JOIN InstrumentoInversion III ON(III.CodTitulo=EC.CodTituloReferencia) WHERE "
        
    If cboTituloCriterio.ListIndex > 0 Then
        strSQL = strSQL & "EC.CodTitulo='" & strCodTituloCriterio & "' AND TipoAcuerdo='" & strCodEventoCriterio & "' AND " & _
        "EC.CodAdministradora='" & gstrCodAdministradora & "' AND "
    Else
        strSQL = strSQL & "TipoAcuerdo='" & strCodEventoCriterio & "' AND " & _
        "EC.CodAdministradora='" & gstrCodAdministradora & "' AND "
    End If
    strSQL = strSQL & "EstadoEvento='" & strCodEstado & "'"
    
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

    Me.MousePointer = vbDefault
    
End Sub
Private Sub CargarReportes()

    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Vista Activa"
    
End Sub
Private Sub CargarListas()
        
    Dim intRegistro As Integer
    
    '*** Títulos ***
    strSQL = "SELECT CodTitulo CODIGO,(Nemotecnico + space(1) + DescripTitulo) DESCRIP FROM InstrumentoInversion " & _
        "WHERE CodFile='004' AND IndVigente='X' ORDER BY DESCRIP"
        
    CargarControlLista strSQL, cboTituloCriterio, arrTituloCriterio(), Sel_Todos
    CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
    CargarControlLista strSQL, cboTituloReferencia, arrTituloReferencia(), Sel_Defecto
    
    If cboTituloCriterio.ListCount > 0 Then cboTituloCriterio.ListIndex = 0
    
    '*** Tipo de Evento ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPEVE' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEventoCriterio, arrEventoCriterio(), Sel_Defecto
    CargarControlLista strSQL, cboEvento, arrEvento(), Sel_Defecto
        
    If cboEventoCriterio.ListCount > 0 Then cboEventoCriterio.ListIndex = 0
    
    '*** Tipo de Evento ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='ESTACU' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Sel_Defecto
    
    intRegistro = ObtenerItemLista(arrEstado(), Estado_Acuerdo_Ingresado)
    If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
        
End Sub
Private Sub InicializarValores()
    
    strEstado = Reg_Defecto
    tabEvento.Tab = 0
    tabEvento.TabEnabled(1) = False
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 8
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 20
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 15
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 15
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
                    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set frmEventoCorporativo = Nothing
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub



Private Sub dtpFechaCorte_LostFocus()

'    If Not EsDiaUtil(dtpFechaCorte.Value) Then
'        If MsgBox("La Fecha de corte es un día no útil, esta seguro que la fecha es correcta ? ", vbYesNo + vbDefaultButton2) = vbNo Then
'            If dtpFechaCorte.Enabled Then dtpFechaCorte.SetFocus
'        End If
'    End If
    
    Dim adoRegistro As ADODB.Recordset
    
    If dtpFechaCorte.Value < dtpFechaOperacion.Value Then
        With adoComm
            Set adoRegistro = New ADODB.Recordset

            .CommandText = "SELECT CodFondo FROM EventoCorporativoOrden " & _
                "WHERE CodFile='004' AND CodAnalitica='" & strCodAnalitica & "' AND NumAcuerdo=" & lngNumAcuerdo & _
                " AND CodAdministradora='" & gstrCodAdministradora & "'"
                
            Set adoRegistro = .Execute

            If adoRegistro.EOF Then
                blnOrdLinea = True
            Else
                blnOrdLinea = False
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
        End With
    End If
    
End Sub


Private Sub dtpFechaEntrega_LostFocus()

'    If CVDate(dtpFechaEntrega.Value) < CVDate(dtpFechaCorte.Value) Then
'       MsgBox "Fecha de Entrega debe se mayor que la de Corte y Junta Porfavor..", vbExclamation
'       If dtpFechaEntrega.Enabled Then dtpFechaEntrega.SetFocus
'    End If
'
End Sub


Private Sub dtpFechaJunta_LostFocus()

'    If CVDate(dtpFechaJunta.Value) >= CVDate(dtpFechaCorte.Value) Then
'        MsgBox "Fecha de Junta debe ser menor que la de Corte", vbCritical, Me.Caption
'        If dtpFechaJunta.Enabled Then dtpFechaJunta.SetFocus
'    End If
'
'    If CVDate(dtpFechaJunta.Value) > CVDate(dtpFechaEntrega.Value) Then
'       MsgBox "Fecha de Junta debe ser menor que la de Entrega", vbCritical, Me.Caption
'       If dtpFechaJunta.Enabled Then dtpFechaJunta.SetFocus
'    End If
'
'    If CVDate(dtpFechaJunta.Value) > CVDate(dtpFechaOperacion.Value) Then
'       MsgBox "Fecha de Junta debe ser menor o igual que la del Sistema", vbCritical, Me.Caption
'       If dtpFechaJunta.Enabled Then dtpFechaJunta.SetFocus
'    End If
    
End Sub


Private Sub tabEvento_Click(PreviousTab As Integer)

    Select Case tabEvento.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabEvento.Tab = 0
    End Select
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 4 Then
        Call DarFormatoValor(Value, Decimales_TasaDiaria)
    End If
    
End Sub


Private Sub txtValor_Change()

    If strCodEvento = Codigo_Evento_Dividendo Then Call FormatoCajaTexto(txtValor, Decimales_TasaDiaria)
    If strCodEvento = Codigo_Evento_Liberacion Then Call FormatoCajaTexto(txtValor, Decimales_TasaDiaria)
    If strCodEvento = Codigo_Evento_Nominal Then Call FormatoCajaTexto(txtValor, Decimales_Monto)
    
End Sub


Private Sub txtValor_KeyPress(KeyAscii As Integer)

    If strCodEvento = Codigo_Evento_Dividendo Then Call ValidaCajaTexto(KeyAscii, "M", txtValor, Decimales_TasaDiaria)
    If strCodEvento = Codigo_Evento_Liberacion Then Call ValidaCajaTexto(KeyAscii, "M", txtValor, Decimales_TasaDiaria)
    If strCodEvento = Codigo_Evento_Nominal Then Call ValidaCajaTexto(KeyAscii, "M", txtValor, Decimales_Monto)
    
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
    
   
            gstrNameRepo = "EventoCorporativoGrilla"
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(3)
            ReDim aReportParamFn(2)
            ReDim aReportParamF(2)
                        
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
            
            If cboTituloCriterio.ListIndex = 0 Then
                strCodTituloCriterio = Valor_Comodin
            End If
                        
            aReportParamS(0) = gstrCodAdministradora
            aReportParamS(1) = strCodTituloCriterio
            aReportParamS(2) = strCodEventoCriterio
            aReportParamS(3) = strCodEstado
       
    
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    


End Sub

