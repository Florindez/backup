VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmFeriados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Feriados"
   ClientHeight    =   6555
   ClientLeft      =   1380
   ClientTop       =   900
   ClientWidth     =   8745
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
   ScaleHeight     =   6555
   ScaleWidth      =   8745
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   735
      Left            =   6120
      Picture         =   "frmFeriados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5640
      Width           =   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   7470
      TabIndex        =   26
      Top             =   5640
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
      Left            =   120
      TabIndex        =   25
      Top             =   5640
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   1296
      Buttons         =   4
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
      UserControlWidth=   5700
   End
   Begin MSAdodcLib.Adodc adoConsulta2 
      Height          =   330
      Left            =   4920
      Top             =   6030
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabFeriados 
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmFeriados.frx":0671
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTipoCambio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmFeriados.frx":068D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdOperacion"
      Tab(1).Control(1)=   "fraDetalle"
      Tab(1).Control(2)=   "ucBotonEdicion1"
      Tab(1).Control(3)=   "cmdOperacion2"
      Tab(1).ControlCount=   4
      Begin TAMControls2.ucBotonEdicion2 cmdOperacion 
         Height          =   735
         Left            =   -70440
         TabIndex        =   27
         Top             =   4560
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
      Begin VB.Frame fraTipoCambio 
         Caption         =   "Criterios de Búsqueda"
         Height          =   1695
         Left            =   720
         TabIndex        =   11
         Top             =   600
         Width           =   7080
         Begin VB.ComboBox cboPais 
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
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1140
            Width           =   3735
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
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   645
            Width           =   1455
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
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   660
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "País"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   24
            Top             =   1200
            Width           =   405
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año"
            Height          =   195
            Index           =   0
            Left            =   2880
            TabIndex        =   15
            Top             =   660
            Width           =   345
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mes"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   14
            Top             =   660
            Width           =   360
         End
      End
      Begin VB.Frame fraDetalle 
         Height          =   3855
         Left            =   -74640
         TabIndex        =   1
         Top             =   600
         Width           =   7050
         Begin VB.ComboBox CboPais2 
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
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   2760
            Width           =   3495
         End
         Begin VB.TextBox txtMotivo 
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   20
            Top             =   2160
            Width           =   2445
         End
         Begin MSComCtl2.DTPicker dtpFechaFeriado 
            Height          =   285
            Left            =   1320
            TabIndex        =   2
            Top             =   1440
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
            Format          =   178323457
            CurrentDate     =   38806
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "País"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   22
            Top             =   2820
            Width           =   405
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año"
            Height          =   195
            Index           =   5
            Left            =   2880
            TabIndex        =   8
            Top             =   495
            Width           =   345
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
            Left            =   1200
            TabIndex        =   7
            Top             =   480
            Width           =   1335
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
            Left            =   3360
            TabIndex        =   6
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            Height          =   195
            Index           =   6
            Left            =   360
            TabIndex        =   5
            Top             =   1440
            Width           =   540
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo"
            Height          =   195
            Index           =   7
            Left            =   360
            TabIndex        =   4
            Top             =   2160
            Width           =   585
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            X1              =   360
            X2              =   5000
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mes"
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   3
            Top             =   495
            Width           =   360
         End
      End
      Begin TAMControls.ucBotonEdicion cmdAccion 
         Height          =   390
         Left            =   -72480
         TabIndex        =   9
         Top             =   4035
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   688
         Buttons         =   2
         Caption0        =   "&Guardar"
         Tag0            =   "2"
         Visible0        =   0   'False
         ToolTipText0    =   "Guardar"
         Caption1        =   "&Cancelar"
         Tag1            =   "8"
         Visible1        =   0   'False
         ToolTipText1    =   "Cancelar"
         UserControlHeight=   390
         UserControlWidth=   2700
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmFeriados.frx":06A9
         Height          =   2715
         Left            =   690
         OleObjectBlob   =   "frmFeriados.frx":06C3
         TabIndex        =   10
         Top             =   2430
         Width           =   7110
      End
      Begin TAMControls.ucBotonEdicion ucBotonEdicion1 
         Height          =   390
         Left            =   -66000
         TabIndex        =   16
         Top             =   3480
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   688
         Buttons         =   2
         Caption0        =   "&Guardar"
         Tag0            =   "2"
         Visible0        =   0   'False
         ToolTipText0    =   "Guardar"
         Caption1        =   "&Cancelar"
         Tag1            =   "8"
         Visible1        =   0   'False
         ToolTipText1    =   "Cancelar"
         UserControlHeight=   390
         UserControlWidth=   2700
      End
      Begin TAMControls.ucBotonEdicion cmdOperacion2 
         Height          =   390
         Left            =   -70500
         TabIndex        =   19
         Top             =   4740
         Visible         =   0   'False
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   688
         Buttons         =   2
         Caption0        =   "&Guardar"
         Tag0            =   "2"
         Visible0        =   0   'False
         ToolTipText0    =   "Guardar"
         Caption1        =   "&Cancelar"
         Tag1            =   "8"
         Visible1        =   0   'False
         ToolTipText1    =   "Cancelar"
         UserControlHeight=   390
         UserControlWidth=   2700
      End
   End
   Begin TAMControls.ucBotonEdicion cmdOpcion2 
      Height          =   390
      Left            =   450
      TabIndex        =   17
      Top             =   5880
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   688
      Buttons         =   4
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
      UserControlHeight=   390
      UserControlWidth=   5700
   End
   Begin TAMControls.ucBotonEdicion cmdSalir2 
      Height          =   390
      Left            =   6960
      TabIndex        =   18
      Top             =   5880
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   688
      Caption0        =   "&Salir"
      Tag0            =   "9"
      Visible0        =   0   'False
      ToolTipText0    =   "Salir"
      UserControlHeight=   390
      UserControlWidth=   1200
   End
End
Attribute VB_Name = "frmFeriados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Mantenimiento de Tasas VAC"
Option Explicit

Dim strClaseTC          As String, arrClaseTC()         As String
Dim strAnio             As String, arrAnio()            As String
Dim strMes              As String, arrMes()             As String
Dim strPais             As String, arrPais()            As String
Dim strPais2            As String, arrPais2()           As String
Dim strCodMoneda        As String, arrMoneda()          As String
Dim strCodMonedaCambio  As String, arrMonedaCambio()    As String
Dim strDiaInicial       As String, strDiaFinal          As String
Dim strFechaDesde       As String, strFechaHasta        As String
Dim strEstado           As String, strSQL               As String
Dim strFechaFeriado      As String, strMotivo           As String
Dim adoConsulta As ADODB.Recordset

Public Sub Adicionar()

    frmMainMdi.stbMdi.Panels(3).Text = "Ingresar Feriados..."
                
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabFeriados
        .TabEnabled(0) = False
        .TabEnabled(1) = True
        .Tab = 1
    End With

End Sub

Public Sub Buscar()

    Dim intmes As Integer, intAnio As Integer
    Dim intTemporal As Integer
    Dim datFechaInicioMes As Date, datFechaFinMes As Date
    Dim datFechaTemporal As Date
    
    intmes = CInt(strMes)
    intAnio = CInt(strAnio)
      
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
            
    strSQL = "{ call up_GNGenFeriados('" & strPais & "','" & Convertyyyymmdd(datFechaInicioMes) & "','" & Convertyyyymmdd(datFechaFinMes) & "' ) }"
    adoConn.Execute strSQL
    
    Set adoConsulta = New ADODB.Recordset
    
    With adoConsulta
'        .ConnectionString = gstrConnectConsulta
'        .RecordSource = strSQL
'        .Refresh


        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    tdgConsulta.DataSource = adoConsulta
    tdgConsulta.Refresh
    
    If adoConsulta.RecordCount > 0 Then strEstado = Reg_Consulta
            
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabFeriados
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .Tab = 0
    End With
    strEstado = Reg_Consulta
    
End Sub

Public Sub Eliminar()

  Dim adoRegistro            As ADODB.Recordset
  Dim strOperacion           As String, strFechaOperacionM       As String
  Dim strDatoConsultaCli     As String, strDatoCli               As String
  Dim strDatoConsultaVal     As String, strDatoConsultaValor     As String
  Dim strDatoTipo            As String, strDatoConsultFecGar     As String
  Dim strDatoConsultaTip     As String, strDatoConsultaCantidad  As String
  Dim strTipo                As String, strDatoConsultaMoneda    As String
  Dim strDatoConsultaMonto   As String, strDatosConsultaClase    As String
  Dim rst As Boolean, strMensaje                                 As String

    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
    
    strFechaFeriado = Convertyyyymmdd(tdgConsulta.Columns(0))
      
        strMensaje = "Se procederá a eliminar el dia feriado " & tdgConsulta.Columns(0) & vbNewLine & vbNewLine & vbNewLine & "¿ Seguro de continuar ?"
        
        If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                     
            Set adoRegistro = New ADODB.Recordset
         
            With adoComm
               
           .CommandText = "{ call up_GNEliFeriados('" & strFechaFeriado & "') }"
            Set adoRegistro = .Execute
                
            End With
        
             Me.MousePointer = vbDefault
             
             MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation
            
             Call Buscar
         Else
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
        End If
             
    End If
  
End Sub

Public Sub Grabar()

    Dim adoRegistro As ADODB.Recordset
    Dim intRegistro As Integer, intAccion           As Integer
    Dim intNumError As Integer, lngNumError             As Long
    Dim strFecha    As String, strFechaSiguiente    As String
    Dim mensaje As String
        
    If strEstado = Reg_Consulta Then Exit Sub
    
    On Error GoTo CtrlError
    
    strFecha = Convertyyyymmdd(dtpFechaFeriado.Value)
    
    If strEstado = Reg_Adicion Then
    
        mensaje = Mensaje_Adicion
    
        If TodoOK() Then
               
            If MsgBox(mensaje, vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) <> vbYes Then Exit Sub
            
            Set adoRegistro = New ADODB.Recordset
              '*** Guardar Feriados ***
            With adoComm
               
                .CommandText = "{ call up_GNManFeriados( '" & strPais & "', '" & strFecha & "', '" & Trim(txtMotivo.Text) & "' , '" & gstrLogin & "', 'I','','','') }"
                Set adoRegistro = .Execute
                
            End With
            
            Me.MousePointer = vbDefault
        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            
            
            With tabFeriados
                .TabEnabled(0) = True
                .TabEnabled(1) = False
                .Tab = 0
            End With
                  
            Call Buscar
            Call Limpiar
            Me.tdgConsulta.AllowRowSelect = True
            Me.tdgConsulta.SetFocus
            
      End If
        
    ElseIf strEstado = Reg_Edicion Then
    
        If TodoOK() Then
                  
         mensaje = Mensaje_Edicion
         
         If MsgBox(mensaje, vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) <> vbYes Then Exit Sub
         
            With adoComm
                         
                .CommandText = "{ call up_GNManFeriados( '" & strPais2 & "', '" & strFechaFeriado & "', '" & strMotivo & "' , '" & gstrLogin & "', 'U','" & strPais2 & "','" & strFecha & "','" & Trim(txtMotivo.Text) & "') }"
                Set adoRegistro = .Execute

            End With
    
             Me.MousePointer = vbDefault
             
             MsgBox Mensaje_Modificar, vbExclamation
             
             Call Buscar
             
            tabFeriados.TabEnabled(0) = True
            tabFeriados.TabEnabled(1) = False
            tabFeriados.Tab = 0
            cmdOpcion.Visible = True
                          
         'Exit Sub
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
'    adoComm.CommandText = "ROLLBACK TRAN ProcOrden"
'    adoConn.Execute adoComm.CommandText
           
End Sub

Private Function TodoOK() As Boolean
        
    Dim adoRegistro         As ADODB.Recordset
    Dim strFecha            As String, strFechaSiguiente        As String
    
    TodoOK = False
   
    If Trim(txtMotivo.Text) = Valor_Caracter Then
        MsgBox "Debe ingresar el motivo del feriado.", vbCritical, Me.Caption
        
        txtMotivo.SetFocus
        Exit Function
    End If
        
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function

Public Sub Imprimir()
    
    Call SubImprimir(1)
    
End Sub

Public Sub Modificar()

    If tdgConsulta.SelBookmarks.Count < 1 Then Exit Sub
    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabFeriados
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .Tab = 1
        End With
        
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro   As ADODB.Recordset
    Dim intRegistro As Integer
    
    Select Case strModo
    
        Case Reg_Adicion
                  
            dtpFechaFeriado.Value = gdatFechaActual
            lblMes.Caption = Trim(cboMes.Text)
            lblPeriodo.Caption = Trim(cboAnio.Text)
             
        Case Reg_Edicion
        
            lblMes.Caption = Trim(cboMes.Text)
            lblPeriodo.Caption = Trim(cboAnio.Text)
            
            strFechaFeriado = Convertyyyymmdd(tdgConsulta.Columns(0))
            strMotivo = Trim(tdgConsulta.Columns(1))
          
            Set adoRegistro = New ADODB.Recordset
            
            With adoComm
            
            .CommandText = "{ call up_GNSelFeriados ('" & strPais2 & "', '" & strFechaFeriado & "', '" & strMotivo & "') }"
                       
             Set adoRegistro = .Execute
              
              If Not adoRegistro.EOF Then
            
               Me.dtpFechaFeriado.Value = Trim(adoRegistro("FechaFeriado"))
               Me.txtMotivo.Text = Trim(adoRegistro("Motivo"))
                                   
              End If
              
             cmdOperacion.Visible = True
             Set cmdOperacion.FormularioActivo = Me
            
            adoRegistro.Close: Set adoRegistro = Nothing
                    
        End With
          
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

 Select Case Index
     Case 1
        
            If cboMes.ListIndex < 0 Then
                MsgBox "Seleccione Mes.", vbCritical
                Exit Sub
            End If

            If cboAnio.ListIndex < 0 Then
                MsgBox "Seleccione Año.", vbCritical
               Exit Sub
            End If
        
            gstrNameRepo = "Feriados"
        
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
            aReportParamF(1) = strDiaInicial
            aReportParamF(2) = strDiaFinal
            aReportParamF(3) = Format(Time, "hh:mm:ss")
            aReportParamF(4) = strClaseTC
            aReportParamF(5) = gstrNombreEmpresa & Space(1)
  
            
            aReportParamS(0) = Convertyyyymmdd(CDate(strDiaInicial))
            aReportParamS(1) = Convertyyyymmdd(DateAdd("d", 1, CDate(strDiaFinal)))
            aReportParamS(2) = strPais
            
        End Select

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
    
End Sub

Private Sub cboMes_Click()

    strMes = "00"
    If cboMes.ListIndex < 0 Then Exit Sub
        
    strMes = arrMes(cboMes.ListIndex)
    Me.Refresh
    
    If strAnio = "0000" Then Exit Sub
    
    Call Buscar
        
End Sub

'Private Sub dtpFechaTipoCambio_Change()
'
'    Dim adoRegistro     As ADODB.Recordset
'
'    strFechaDesde = Convertyyyymmdd(dtpFechaTipoCambio.Value)
'    strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaTipoCambio.Value))
'
'    Set adoRegistro = New ADODB.Recordset
'
'    adoComm.CommandText = "SELECT ValorTipoCambioCompra,ValorTipoCambioVenta FROM TipoCambioFondoTemporal " & _
'        "WHERE FechaTipoCambio >='" & strFechaDesde & "' AND FechaTipoCambio <'" & strFechaHasta & "'"
'    Set adoRegistro = adoComm.Execute
'
'    If Not adoRegistro.EOF Then
'        txtValorCompra.Text = CStr(adoRegistro("ValorTipoCambioCompra"))
'        txtValorVenta.Text = CStr(adoRegistro("ValorTipoCambioVenta"))
'    End If
'    adoRegistro.Close: Set adoRegistro = Nothing
'
'End Sub

Private Sub cboPais_Click()
    
    If cboPais.ListIndex < 0 Then Exit Sub
        
    strPais = arrPais(cboPais.ListIndex)
    Me.Refresh
           
End Sub

Private Sub cboPais2_Click()
    
    If CboPais2.ListIndex < 0 Then Exit Sub
        
    strPais2 = arrPais2(CboPais2.ListIndex)
    Me.Refresh
        
End Sub

Private Sub cmdImprimir_Click()
    Call Imprimir
End Sub

Public Sub Limpiar()
 
    Me.txtMotivo.Text = ""
      
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
    
'    For intCont = 0 To (fraTipoCambio.Count - 1)
'        Call FormatoMarco(fraTipoCambio(intCont))
'    Next
            
End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabFeriados.Tab = 0
    tabFeriados.TabEnabled(1) = False
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 28
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 32
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdOperacion.FormularioActivo = Me
    
End Sub
Private Sub CargarListas()

    Dim intRegistro As Integer
            
    '*** Meses ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='DSCMES' ORDER BY CodParametro"
    CargarControlLista strSQL, cboMes, arrMes(), Valor_Caracter
    
    '*** Años ***
    strSQL = "SELECT ValorParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='RNGANI' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboAnio, arrAnio(), Valor_Caracter
    
    '*** Países ***
    strSQL = "SELECT CodPais CODIGO,DescripPais DESCRIP FROM Pais"
    CargarControlLista strSQL, cboPais, arrPais(), Valor_Caracter
    CargarControlLista strSQL, CboPais2, arrPais2(), Valor_Caracter
    If cboPais.ListCount > 0 Then cboPais.ListIndex = 0
    If CboPais2.ListCount > 0 Then CboPais2.ListIndex = 0
    
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

'Private Sub tabTipoCambio_Click(PreviousTab As Integer)

 '   Select Case tabTipoCambio.Tab
        'Case 1
  '          If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vQuery)
   '         If strEstado = Reg_Defecto Then tabTipoCambio.Tab = 0
        
    'End Select
    
'End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 1 Then
        Call DarFormatoValor(Value, Decimales_TipoCambio)
    End If
    
    If ColIndex = 2 Then
        Call DarFormatoValor(Value, Decimales_TipoCambio)
    End If
    
End Sub

'Private Sub txtValorCompra_Change()
'
'    Call FormatoCajaTexto(txtValorCompra, Decimales_TipoCambio)
'
'End Sub
'
'
'
'Private Sub txtValorCompra_KeyPress(KeyAscii As Integer)
'
'    Call ValidaCajaTexto(KeyAscii, "M", txtValorCompra, Decimales_TipoCambio)
'
'End Sub


'Private Sub txtValorVenta_Change()
'
'    Call FormatoCajaTexto(txtValorVenta, Decimales_TipoCambio)
'
'End Sub
'
'
'Private Sub txtValorVenta_KeyPress(KeyAscii As Integer)
'
'    Call ValidaCajaTexto(KeyAscii, "M", txtValorVenta, Decimales_TipoCambio)
'
'End Sub

Private Sub tdgConsulta_HeadClick(ByVal ColIndex As Integer)
    Static numColindex As Integer

    tdgConsulta.Splits(0).Columns(numColindex).HeadingStyle.ForegroundPicture = Null

    Call OrdenarDBGrid(ColIndex, adoConsulta, tdgConsulta)
    
    numColindex = ColIndex
End Sub
