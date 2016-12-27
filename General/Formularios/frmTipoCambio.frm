VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{442CAE95-1D41-47B1-BE83-6995DA3CE254}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmTipoCambio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Cambio"
   ClientHeight    =   6000
   ClientLeft      =   1380
   ClientTop       =   900
   ClientWidth     =   6645
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
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6000
   ScaleWidth      =   6645
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   4920
      TabIndex        =   7
      Top             =   5160
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
      TabIndex        =   6
      Top             =   5160
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1296
      Buttons         =   2
      Caption0        =   "&Modificar"
      Tag0            =   "3"
      ToolTipText0    =   "Modificar"
      Caption1        =   "&Buscar"
      Tag1            =   "5"
      ToolTipText1    =   "Buscar"
      UserControlWidth=   2700
   End
   Begin TabDlg.SSTab tabTipoCambio 
      Height          =   5055
      Left            =   0
      TabIndex        =   10
      Top             =   60
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmTipoCambio.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraTipoCambio(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdgConsulta"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmTipoCambio.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraTipoCambio(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAccion"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   3360
         TabIndex        =   9
         Top             =   4200
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
      Begin VB.Frame fraTipoCambio 
         Caption         =   "Tipo de Cambio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   6120
         Begin MSComCtl2.DTPicker dtpFechaTipoCambio 
            Height          =   285
            Left            =   2160
            TabIndex        =   8
            Top             =   2400
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   146669569
            CurrentDate     =   38806
         End
         Begin TAMControls.TAMTextBox txtValorCompra 
            Height          =   315
            Left            =   2160
            TabIndex        =   31
            Top             =   2790
            Width           =   2025
            _ExtentX        =   3572
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
            Container       =   "frmTipoCambio.frx":0038
            Text            =   "0.000000000000"
            Decimales       =   12
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   12
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtValorVenta 
            Height          =   315
            Left            =   2160
            TabIndex        =   32
            Top             =   3150
            Width           =   2025
            _ExtentX        =   3572
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
            Container       =   "frmTipoCambio.frx":0054
            Text            =   "0.000000000000"
            Decimales       =   12
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   12
            MaximoValor     =   999999999
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Venta"
            Height          =   195
            Index           =   12
            Left            =   360
            TabIndex        =   30
            Top             =   3180
            Width           =   1005
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            X1              =   360
            X2              =   5700
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Label lblMonedaCambio 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2160
            TabIndex        =   29
            Top             =   1240
            Width           =   3495
         End
         Begin VB.Label lblMoneda 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2160
            TabIndex        =   28
            Top             =   800
            Width           =   3495
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda"
            Height          =   195
            Index           =   11
            Left            =   360
            TabIndex        =   27
            Top             =   820
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda Cambio"
            Height          =   195
            Index           =   10
            Left            =   360
            TabIndex        =   26
            Top             =   1260
            Width           =   1365
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Compra"
            Height          =   195
            Index           =   7
            Left            =   360
            TabIndex        =   23
            Top             =   2820
            Width           =   1140
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            Height          =   195
            Index           =   6
            Left            =   360
            TabIndex        =   22
            Top             =   2400
            Width           =   540
         End
         Begin VB.Label lblPeriodo 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   4320
            TabIndex        =   21
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblMes 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2160
            TabIndex        =   20
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblTipo 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2160
            TabIndex        =   19
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label lblDescrip 
            BackStyle       =   0  'Transparent
            Caption         =   "Año"
            Height          =   255
            Index           =   5
            Left            =   3840
            TabIndex        =   18
            Top             =   1695
            Width           =   375
         End
         Begin VB.Label lblDescrip 
            BackStyle       =   0  'Transparent
            Caption         =   "Mes"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   17
            Top             =   1695
            Width           =   375
         End
         Begin VB.Label lblDescrip 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   16
            Top             =   380
            Width           =   375
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmTipoCambio.frx":0070
         Height          =   1875
         Left            =   -74760
         OleObjectBlob   =   "frmTipoCambio.frx":008A
         TabIndex        =   5
         Top             =   2655
         Width           =   6120
      End
      Begin VB.Frame fraTipoCambio 
         Caption         =   "Criterios de Búsqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Index           =   0
         Left            =   -74760
         TabIndex        =   11
         Top             =   480
         Width           =   6120
         Begin VB.ComboBox cboMonedaCambio 
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
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1080
            Width           =   3735
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
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   720
            Width           =   3735
         End
         Begin VB.ComboBox cboClaseTC 
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
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   3735
         End
         Begin VB.ComboBox cboMes 
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
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1485
            Width           =   1455
         End
         Begin VB.ComboBox cboAnio 
            Appearance      =   0  'Flat
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
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1485
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda Cambio"
            Height          =   195
            Index           =   9
            Left            =   360
            TabIndex        =   25
            Top             =   1080
            Width           =   1365
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda"
            Height          =   195
            Index           =   8
            Left            =   360
            TabIndex        =   24
            Top             =   720
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   14
            Top             =   375
            Width           =   390
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mes"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   13
            Top             =   1500
            Width           =   360
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año"
            Height          =   195
            Index           =   0
            Left            =   3720
            TabIndex        =   12
            Top             =   1500
            Width           =   345
         End
      End
   End
End
Attribute VB_Name = "frmTipoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Mantenimiento de Tasas VAC"
Option Explicit

Dim strClaseTC          As String, arrClaseTC()         As String
Dim strAnio             As String, arrAnio()            As String
Dim strMes              As String, arrMes()             As String
Dim strCodMoneda        As String, arrMoneda()          As String
Dim strCodMonedaCambio  As String, arrMonedaCambio()    As String
Dim strDiaInicial       As String, strDiaFinal          As String
Dim strFechaDesde       As String, strFechaHasta        As String
Dim strEstado           As String, strSQL               As String
Dim adoConsulta         As ADODB.Recordset
Dim indSortAsc          As Boolean, indSortDesc         As Boolean

Public Sub Adicionar()

End Sub

Public Sub Buscar()

    Dim intmes As Integer, intAnio As Integer
    Dim intTemporal As Integer
    Dim datFechaInicioMes As Date, datFechaFinMes As Date
    Dim datFechaTemporal As Date
    
    Set adoConsulta = New ADODB.Recordset
    
    intmes = CInt(strMes)
    intAnio = CInt(strAnio)
    
    If strCodMoneda = Valor_Caracter Then Exit Sub
    
    If intmes = 1 Then
        intTemporal = UltimoDiaMes(12, intAnio - 1)
        datFechaTemporal = Convertddmmyyyy(Format(intAnio - 1, "0000") & Format(12, "00") & Format(intTemporal, "00"))
    Else
        intTemporal = UltimoDiaMes(intmes - 1, intAnio)
        datFechaTemporal = Convertddmmyyyy(Format(intAnio, "0000") & Format(intmes - 1, "00") & Format(intTemporal, "00"))
    End If
    datFechaInicioMes = DateAdd("d", 1, datFechaTemporal)
    strDiaInicial = CStr(datFechaInicioMes)
    
    intTemporal = UltimoDiaMes(intmes, intAnio)
    datFechaTemporal = Convertddmmyyyy(Format(intAnio, "0000") & Format(intmes, "00") & Format(intTemporal, "00"))
    datFechaFinMes = DateAdd("d", 1, datFechaTemporal)
    strDiaFinal = CStr(datFechaTemporal)
            
    strSQL = "{ call up_GNGenFechasTipoCambio('" & strClaseTC & "','" & strCodMoneda & "','" & Convertyyyymmdd(datFechaInicioMes) & "','" & Convertyyyymmdd(datFechaFinMes) & "' ) }"
    adoConn.Execute strSQL
    
    strSQL = "SELECT FechaTipoCambio,ValorTipoCambioCompra,ValorTipoCambioVenta " & _
        "FROM TipoCambioFondoTemporal " & _
        "ORDER BY FechaTipoCambio"
        
    strEstado = Reg_Consulta
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
    With tabTipoCambio
        .TabEnabled(0) = True
        .Tab = 0
        .TabEnabled(1) = False
    End With
    strEstado = Reg_Consulta
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Vista Activa"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Por Rango de Fecha"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    
End Sub



Public Sub Eliminar()

End Sub


Public Sub Grabar()

    Dim adoRegistro As ADODB.Recordset
    Dim intRegistro As Integer, intAccion           As Integer
    Dim intNumError As Integer
    Dim strFecha    As String, strFechaSiguiente    As String
        
    Set adoRegistro = New ADODB.Recordset
    
    On Error GoTo CtrlError
    
    If TodoOK() Then
        With adoComm
            '*** Actualizar valor de cambio en tabla temporal ***
            .CommandText = "UPDATE TipoCambioFondoTemporal SET " & _
                "ValorTipoCambioCompra=" & CDec(txtValorCompra.Text) & "," & _
                "ValorTipoCambioVenta=" & CDec(txtValorVenta.Text) & " " & _
                "WHERE (FechaTipoCambio>='" & strFechaDesde & "'AND FechaTipoCambio<'" & strFechaHasta & "') AND " & _
                "CodMoneda='" & strCodMoneda & "'"
            adoConn.Execute .CommandText
                
            '*** Registro del valor de cambio en tabla definitiva ***
            .CommandText = "SELECT * FROM TipoCambioFondoTemporal"
            Set adoRegistro = .Execute
        
            Do While Not adoRegistro.EOF
                strFecha = Convertyyyymmdd(adoRegistro("FechaTipoCambio"))
                strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, adoRegistro("FechaTipoCambio")))
            
                .CommandText = "UPDATE TipoCambioFondo SET " & _
                    "ValorTipoCambioCompra=" & CDbl(adoRegistro("ValorTipoCambioCompra")) & "," & _
                    "ValorTipoCambioVenta=" & CDbl(adoRegistro("ValorTipoCambioVenta")) & " " & _
                    "WHERE (FechaTipoCambio>='" & strFecha & "'AND FechaTipoCambio<'" & strFechaSiguiente & "') AND CodTipoCambio='" & strClaseTC & "' AND " & _
                    "CodMoneda='" & strCodMoneda & "'"
                adoConn.Execute .CommandText, intRegistro
                
                If intRegistro = 0 And (CDbl(adoRegistro("ValorTipoCambioCompra")) > 0 Or CDbl(adoRegistro("ValorTipoCambioVenta")) > 0) Then
                    .CommandText = "INSERT INTO TipoCambioFondo VALUES ('" & strClaseTC & "','" & Convertyyyymmdd(adoRegistro("FechaTipoCambio")) & "','" & _
                        strCodMoneda & "','" & strCodMonedaCambio & "'," & _
                        CDbl(adoRegistro("ValorTipoCambioCompra")) & "," & _
                        CDbl(adoRegistro("ValorTipoCambioVenta")) & ")"
                    adoConn.Execute .CommandText
                End If
                
                adoRegistro.MoveNext
            Loop
            adoRegistro.Close: Set adoRegistro = Nothing
            
            '*** Actualizar en FondoValorCuota ***
            strFecha = Convertyyyymmdd(dtpFechaTipoCambio.Value)
            strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, dtpFechaTipoCambio.Value))
    
            .CommandText = "UPDATE FondoValorCuota SET ValorTipoCambio=" & CDec(txtValorCompra.Text) & " " & _
                "WHERE (FechaCuota >= '" & strFecha & "' AND FechaCuota < '" & strFechaSiguiente & "')"
            adoConn.Execute .CommandText
        End With
        
        MsgBox Mensaje_Edicion_Exitosa, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
                
        cmdOpcion.Visible = True
        With tabTipoCambio
            .TabEnabled(0) = True
            .Tab = 0
        End With
        Call Buscar
    End If
    Exit Sub
    
CtrlError:
    intAccion = ControlErrores
    Select Case intAccion
        Case 0: Resume
        Case 1: Resume Next
        Case Else
            intNumError = err.Number
            err.Raise Number:=intNumError
            err.Clear
    End Select
        
End Sub


Private Function TodoOK() As Boolean
        
    Dim adoRegistro         As ADODB.Recordset
    Dim strFecha            As String, strFechaSiguiente        As String
    
    TodoOK = False
    
    Set adoRegistro = New ADODB.Recordset
    strFecha = Convertyyyymmdd(dtpFechaTipoCambio.Value)
    strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, dtpFechaTipoCambio.Value))
                
    '*** Se Realizó Cierre anteriormente ? ***
    adoComm.CommandText = "SELECT IndCierre FROM FondoValorCuota " & _
        "WHERE (FechaCuota >='" & strFecha & "' AND FechaCuota < '" & strFechaSiguiente & "') AND IndCierre = 'X'"
    Set adoRegistro = adoComm.Execute

    If Not adoRegistro.EOF Then
        If adoRegistro("IndCierre") = Valor_Indicador Then
            'MsgBox "El Cierre Diario del Día " & CStr(dtpFechaTipoCambio.Value) & " ya fué realizado antes. No se puede actualizar el Tipo de Cambio.", vbCritical, Me.Caption

            'adoRegistro.Close: Set adoRegistro = Nothing
            'Exit Function
        End If
    End If
    adoRegistro.Close: Set adoRegistro = Nothing
    
    If CDbl(txtValorCompra.Text) = 0 Then
        MsgBox "Debe ingresar el valor de cambio compra.", vbCritical, Me.Caption
        
        txtValorCompra.SetFocus
        Exit Function
    End If

    If CDbl(txtValorVenta.Text) = 0 Then
        MsgBox "Debe ingresar el valor de cambio venta.", vbCritical, Me.Caption
        
        txtValorVenta.SetFocus
        Exit Function
    End If
        
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function
Public Sub Imprimir()
    
    Call SubImprimir(1)
    
End Sub

Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabTipoCambio
            .TabEnabled(0) = False
            .Tab = 1
            .TabEnabled(1) = True
        End With
        
    End If
    
End Sub


Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro   As ADODB.Recordset
    Dim intRegistro As Integer
    
    Select Case strModo
        Case Reg_Edicion
            Set adoRegistro = New ADODB.Recordset
            
            adoComm.CommandText = "SELECT MAX(FechaTipoCambio) FechaTipoCambio FROM TipoCambioFondoTemporal"
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                If dtpFechaTipoCambio.MinDate > adoRegistro("FechaTipoCambio") Then dtpFechaTipoCambio.MinDate = adoRegistro("FechaTipoCambio")
                dtpFechaTipoCambio.MaxDate = adoRegistro("FechaTipoCambio")
            End If
            adoRegistro.Close
            
            adoComm.CommandText = "SELECT MIN(FechaTipoCambio) FechaTipoCambio FROM TipoCambioFondoTemporal"
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                If dtpFechaTipoCambio.MaxDate < adoRegistro("FechaTipoCambio") Then dtpFechaTipoCambio.MaxDate = adoRegistro("FechaTipoCambio")
                dtpFechaTipoCambio.MinDate = adoRegistro("FechaTipoCambio")
            End If
            adoRegistro.Close
                        
            strFechaDesde = Convertyyyymmdd(tdgConsulta.Columns(0))
            strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, tdgConsulta.Columns(0)))
            
            adoComm.CommandText = "SELECT FechaTipoCambio,ValorTipoCambioCompra,ValorTipoCambioVenta FROM TipoCambioFondoTemporal " & _
                "WHERE FechaTipoCambio >='" & strFechaDesde & "' AND FechaTipoCambio <'" & strFechaHasta & "'"
                
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                lblTipo.Caption = Trim(cboClaseTC.Text)
                lblMoneda.Caption = Trim(cboMoneda.Text)
                lblMonedaCambio.Caption = Trim(cboMonedaCambio.Text)
                lblMes.Caption = Trim(cboMes.Text)
                lblPeriodo.Caption = Trim(cboAnio.Text)
                
                dtpFechaTipoCambio.Value = adoRegistro("FechaTipoCambio")
                txtValorCompra.Text = CStr(adoRegistro("ValorTipoCambioCompra"))
                txtValorVenta.Text = CStr(adoRegistro("ValorTipoCambioVenta"))
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
    
    End Select
    
End Sub


Public Sub Salir()

    Unload Me
    
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



Public Sub SubImprimir(Index As Integer)

    Dim strSeleccionRegistro    As String
    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()

    Set frmReporte = New frmVisorReporte
    
    ReDim aReportParamS(2)
    ReDim aReportParamFn(5)
    ReDim aReportParamF(5)

    aReportParamFn(0) = "Usuario"
    aReportParamFn(1) = "FechaDel"
    aReportParamFn(2) = "FechaAl"
    aReportParamFn(3) = "Hora"
    aReportParamFn(4) = "Tipo"
    aReportParamFn(5) = "NombreEmpresa"

    aReportParamF(0) = gstrLogin
    aReportParamF(3) = Format(Time, "hh:mm:ss")
    aReportParamF(4) = strClaseTC
    aReportParamF(5) = gstrNombreEmpresa & Space(1)
    
    aReportParamS(0) = strClaseTC
    
    Select Case Index
        Case 1
            If cboClaseTC.ListIndex < 0 Then
                MsgBox "Seleccione Tipo.", vbCritical
                Exit Sub
            End If

            If cboMes.ListIndex < 0 Then
                MsgBox "Seleccione Mes.", vbCritical
                Exit Sub
            End If

            If cboAnio.ListIndex < 0 Then
                MsgBox "Seleccione Año.", vbCritical
                Exit Sub
            End If
            aReportParamF(1) = strDiaInicial
            aReportParamF(2) = strDiaFinal
            
            aReportParamS(1) = Convertyyyymmdd(CDate(strDiaInicial))
            aReportParamS(2) = Convertyyyymmdd(DateAdd("d", 1, CDate(strDiaFinal)))

        Case 2
            If cboClaseTC.ListIndex < 0 Then
                MsgBox "Seleccione Tipo.", vbCritical
                Exit Sub
            End If

            strSeleccionRegistro = "{TipoCambioFondo.FechaTipoCambio} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal
            If gstrSelFrml <> "0" Then
                Me.Refresh
                aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
                aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                
                aReportParamS(1) = Convertyyyymmdd(CDate(aReportParamF(1)))
                aReportParamS(2) = Convertyyyymmdd(DateAdd("d", 1, CDate(aReportParamF(2))))
            Else
                Exit Sub
            End If
    End Select

    gstrNameRepo = "TipoDeCambio"
    gstrSelFrml = ""
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub



Private Sub cboAnio_Click()

    strAnio = "0000"
    If cboAnio.ListIndex < 0 Then Exit Sub
    
    strAnio = Trim(arrAnio(cboAnio.ListIndex))
    Me.Refresh
    
    If strMes = "00" Then Exit Sub
    
    Call Buscar
    
End Sub


Private Sub cboClaseTC_Click()

    strClaseTC = Valor_Caracter
    If cboClaseTC.ListIndex < 0 Then Exit Sub
    
    strClaseTC = Trim(arrClaseTC(cboClaseTC.ListIndex))
    
    cboMes_Click
    
End Sub


Private Sub cboMes_Click()

    strMes = "00"
    If cboMes.ListIndex < 0 Then Exit Sub
        
    strMes = arrMes(cboMes.ListIndex)
    Me.Refresh
    
    If strAnio = "0000" Then Exit Sub
    
    Call Buscar
        
End Sub


Private Sub cboMoneda_Click()

    Dim intRegistro As Integer
    
    strCodMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Left(Trim(arrMoneda(cboMoneda.ListIndex)), 2)
    
    intRegistro = ObtenerItemLista(arrMonedaCambio(), Right(Trim(arrMoneda(cboMoneda.ListIndex)), 2))
    If intRegistro >= 0 Then cboMonedaCambio.ListIndex = intRegistro
    
    cboMes_Click
    
End Sub


Private Sub cboMonedaCambio_Click()

    strCodMonedaCambio = Valor_Caracter
    If cboMonedaCambio.ListIndex < 0 Then Exit Sub
    
    strCodMonedaCambio = Trim(arrMonedaCambio(cboMonedaCambio.ListIndex))
            
End Sub


Private Sub dtpFechaTipoCambio_Change()

    Dim adoRegistro     As ADODB.Recordset
    
    strFechaDesde = Convertyyyymmdd(dtpFechaTipoCambio.Value)
    strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaTipoCambio.Value))
    
    Set adoRegistro = New ADODB.Recordset
    
    adoComm.CommandText = "SELECT ValorTipoCambioCompra,ValorTipoCambioVenta FROM TipoCambioFondoTemporal " & _
        "WHERE FechaTipoCambio >='" & strFechaDesde & "' AND FechaTipoCambio <'" & strFechaHasta & "'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        txtValorCompra.Text = CStr(adoRegistro("ValorTipoCambioCompra"))
        txtValorVenta.Text = CStr(adoRegistro("ValorTipoCambioVenta"))
    End If
    adoRegistro.Close: Set adoRegistro = Nothing
            
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

Private Sub DarFormato()

    Dim intCont As Integer
    Dim elemento As Object
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
    For intCont = 0 To (fraTipoCambio.Count - 1)
        Call FormatoMarco(fraTipoCambio(intCont))
    Next
    
    For Each elemento In Me.Controls
    
        If TypeOf elemento Is TDBGrid Then
            Call FormatoGrilla(elemento)
        End If
    
    Next
            
End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    txtValorCompra.Text = "0"
    txtValorVenta.Text = "0"
    tabTipoCambio.Tab = 0
    tabTipoCambio.TabEnabled(1) = False
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 28
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 32
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub
Private Sub CargarListas()

    Dim intRegistro As Integer
        
    '*** Clase de tipo de cambio ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPCAM' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboClaseTC, arrClaseTC(), Valor_Caracter
    
    intRegistro = ObtenerItemLista(arrClaseTC(), Codigo_TipoCambio_Conasev)
    If intRegistro >= 0 Then cboClaseTC.ListIndex = intRegistro
    
    '*** Moneda de Cambio ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMonedaCambio, arrMonedaCambio(), Valor_Caracter
    
'    If cboMonedaCambio.ListCount > 0 Then cboMonedaCambio.ListIndex = 0

    '*** Moneda ***
    strSQL = "SELECT (CodMoneda + CodMonedaCambio) CODIGO,DescripMoneda DESCRIP FROM Moneda WHERE CodSigno<>'' AND CodMoneda<>'" & Codigo_Moneda_Local & "' AND Estado='01' ORDER BY DescripMoneda"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Valor_Caracter
    
    If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = 0
    
    '*** Meses ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='DSCMES' ORDER BY CodParametro"
    CargarControlLista strSQL, cboMes, arrMes(), Valor_Caracter
    
    '*** Años ***
    strSQL = "SELECT ValorParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='RNGANI' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboAnio, arrAnio(), Valor_Caracter
                        
    If gstrPeriodoActual = Valor_Caracter Then
        intRegistro = ObtenerItemLista(arrAnio(), Format(Year(Date), "0000"))
        If intRegistro >= 0 Then cboAnio.ListIndex = intRegistro
        
        intRegistro = ObtenerItemLista(arrMes(), Format(Month(Date), "00"))
        If intRegistro >= 0 Then cboMes.ListIndex = intRegistro
    Else
        intRegistro = ObtenerItemLista(arrAnio(), gstrPeriodoActual)
        If intRegistro >= 0 Then cboAnio.ListIndex = intRegistro
        
        intRegistro = ObtenerItemLista(arrMes(), gstrMesActual)
        If intRegistro >= 0 Then cboMes.ListIndex = intRegistro
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmTipoCambio = Nothing
        
End Sub

Private Sub tabTipoCambio_Click(PreviousTab As Integer)

    Select Case tabTipoCambio.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabTipoCambio.Tab = 0
        
    End Select
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 1 Then
        Call DarFormatoValor(Value, Decimales_TipoCambio)
    End If
    
    If ColIndex = 2 Then
        Call DarFormatoValor(Value, Decimales_TipoCambio)
    End If
    
End Sub

Private Sub txtValorCompra_Change()

    'Call FormatoCajaTexto(txtValorCompra, Decimales_TipoCambio)
            
End Sub

Private Sub txtValorCompra_KeyPress(KeyAscii As Integer)

    'Call ValidaCajaTexto(KeyAscii, "M", txtValorCompra, Decimales_TipoCambio)
    
End Sub

Private Sub txtValorVenta_Change()

    'Call FormatoCajaTexto(txtValorVenta, Decimales_TipoCambio)
    
End Sub

Private Sub txtValorVenta_KeyPress(KeyAscii As Integer)

    'Call ValidaCajaTexto(KeyAscii, "M", txtValorVenta, Decimales_TipoCambio)
    
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
