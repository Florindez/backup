VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{442CAE95-1D41-47B1-BE83-6995DA3CE254}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmRelacionados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personas Relacionadas a la Administradora"
   ClientHeight    =   7545
   ClientLeft      =   -15
   ClientTop       =   720
   ClientWidth     =   7455
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
   ScaleHeight     =   7545
   ScaleWidth      =   7455
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   5760
      TabIndex        =   12
      Top             =   6720
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
      TabIndex        =   1
      Top             =   6720
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
      Caption2        =   "&Eliminar"
      Tag2            =   "4"
      ToolTipText2    =   "Eliminar"
      UserControlWidth=   4200
   End
   Begin TabDlg.SSTab tabPersona 
      Height          =   6555
      Left            =   135
      TabIndex        =   13
      Top             =   120
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   11562
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
      TabPicture(0)   =   "frmRelacionados.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmRelacionados.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(1)=   "adoRelacion"
      Tab(1).Control(2)=   "fraDatos"
      Tab(1).ControlCount=   3
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -70920
         TabIndex        =   11
         Top             =   5640
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
      Begin MSAdodcLib.Adodc adoRelacion 
         Height          =   330
         Left            =   -74040
         Top             =   5760
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
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmRelacionados.frx":0038
         Height          =   5415
         Left            =   240
         OleObjectBlob   =   "frmRelacionados.frx":0052
         TabIndex        =   0
         Top             =   600
         Width           =   6705
      End
      Begin VB.Frame fraDatos 
         Caption         =   "Datos"
         ForeColor       =   &H00800000&
         Height          =   5175
         Left            =   -74760
         TabIndex        =   14
         Top             =   360
         Width           =   6700
         Begin TrueOleDBGrid60.TDBGrid tdgRelacion 
            Bindings        =   "frmRelacionados.frx":29FE
            Height          =   1215
            Left            =   2670
            OleObjectBlob   =   "frmRelacionados.frx":2A18
            TabIndex        =   25
            Top             =   3000
            Width           =   3255
         End
         Begin VB.TextBox txtRazonSocial 
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
            Left            =   2670
            MaxLength       =   78
            TabIndex        =   8
            Top             =   2586
            Width           =   3260
         End
         Begin VB.TextBox txtNombres 
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
            Left            =   2670
            MaxLength       =   25
            TabIndex        =   7
            Top             =   2215
            Width           =   3260
         End
         Begin VB.TextBox txtApellidoMaterno 
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
            Left            =   2670
            MaxLength       =   25
            TabIndex        =   6
            Top             =   1844
            Width           =   3260
         End
         Begin VB.TextBox txtApellidoPaterno 
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
            Left            =   2670
            MaxLength       =   25
            TabIndex        =   5
            Top             =   1473
            Width           =   3260
         End
         Begin VB.ComboBox cboTipoDocumento 
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
            Left            =   2670
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   731
            Width           =   3260
         End
         Begin VB.TextBox txtNumIdentidad 
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
            Left            =   2670
            MaxLength       =   15
            TabIndex        =   4
            Top             =   1102
            Width           =   3260
         End
         Begin VB.ComboBox cboClasePersona 
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
            Left            =   2670
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   360
            Width           =   3260
         End
         Begin MSComCtl2.DTPicker dtpFechaFin 
            Height          =   285
            Left            =   2670
            TabIndex        =   10
            Top             =   4635
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
            CurrentDate     =   38068
         End
         Begin MSComCtl2.DTPicker dtpFechaInicio 
            Height          =   285
            Left            =   2670
            TabIndex        =   9
            Top             =   4305
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
            CurrentDate     =   38068
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Razón Social"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   7
            Left            =   600
            TabIndex        =   24
            Top             =   2606
            Width           =   1290
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Nombres"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   5
            Left            =   600
            TabIndex        =   23
            Top             =   2235
            Width           =   930
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Apellido Materno"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   4
            Left            =   600
            TabIndex        =   22
            Top             =   1864
            Width           =   1710
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Apellido Paterno"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   3
            Left            =   600
            TabIndex        =   21
            Top             =   1493
            Width           =   1545
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Documento"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   2
            Left            =   600
            TabIndex        =   20
            Top             =   751
            Width           =   1500
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Num. Documento"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   1
            Left            =   600
            TabIndex        =   19
            Top             =   1122
            Width           =   1500
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Clase Persona"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   6
            Left            =   600
            TabIndex        =   18
            Top             =   380
            Width           =   1500
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha de Inicio"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   8
            Left            =   600
            TabIndex        =   17
            Top             =   4320
            Width           =   1575
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha Final"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   9
            Left            =   600
            TabIndex        =   16
            Top             =   4650
            Width           =   1575
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Relación"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   15
            Top             =   2970
            Width           =   1200
         End
      End
   End
End
Attribute VB_Name = "frmRelacionados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrClasePersona()           As String, arrRelacion()                As String
Dim arrTipoRelacion()           As String, arrTipoDocumento()           As String
Dim arrCargoRelacion()          As String

Dim strCodClasePersona          As String, strCodRelacion               As String
Dim strCodTipoRelacion          As String, strCodTipoDocumento          As String
Dim strCodCargoRelacion         As String, strCodInstitucion            As String
Dim strEstado                   As String, strSQL                       As String
Dim adoConsulta                 As ADODB.Recordset
Dim indSortAsc                  As Boolean, indSortDesc                 As Boolean

Public Sub Buscar()

    Dim strSQL As String
    Set adoConsulta = New ADODB.Recordset
            
    strSQL = "SELECT CodPersona,DescripPersona,TipoIdentidad,NumIdentidad " & _
        "FROM InstitucionPersona " & _
        "WHERE TipoPersona='" & Codigo_Tipo_Persona_Relacionado & "' AND CodSucursal='999' " & _
        "ORDER BY DescripPersona"
                        
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

Private Sub BuscarRelacion()
        
    strSQL = "SELECT CodParametro,DescripParametro " & _
        "FROM AuxiliarParametro WHERE CodTipoParametro='TIPREL' " & _
        "ORDER BY CodParametro"
        
    With adoRelacion
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With
        
    tdgRelacion.Refresh
            
End Sub
Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabPersona
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .Tab = 0
    End With
    Call Buscar
    
End Sub

Private Sub Deshabilita()

    
End Sub

Public Sub Eliminar()

    If strEstado = Reg_Edicion Or strEstado = Reg_Consulta Then
        frmMainMdi.stbMdi.Panels(3).Text = "Eliminar Persona..."

        If MsgBox(Mensaje_Eliminacion, vbQuestion + vbYesNo + vbDefaultButton2, gstrNombreEmpresa) = vbYes Then
            frmMainMdi.stbMdi.Panels(3).Text = "Eliminando persona..."
            With adoComm
                .CommandText = "DELETE InstitucionPersona WHERE CodPersona='" & Trim(tdgConsulta.Columns(0)) & "' AND " & _
                    "TipoPersona='" & Codigo_Tipo_Persona_Relacionado & "'"
                adoConn.Execute .CommandText
                
                .CommandText = "DELETE ParticipeRelacion WHERE CodPersona='" & Trim(tdgConsulta.Columns(0)) & "' AND " & _
                    "TipoPersona='" & Codigo_Tipo_Persona_Relacionado & "'"
                adoConn.Execute .CommandText
            End With

            frmMainMdi.stbMdi.Panels(3).Text = "Persona eliminada..."
            
            MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation, "Observación"

            tabPersona.Tab = 0
            Call Buscar
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        Else
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        End If
    End If
  
End Sub


Public Sub Grabar()

    Dim intRegistro As Integer, intContador As Integer
    Dim intAccion   As Integer, lngNumError As Long
                
    If strEstado = Reg_Consulta Then Exit Sub
    
    On Error GoTo CtrlError
    
    If strEstado = Reg_Adicion Then
        If TodoOK() Then
            Me.MousePointer = vbHourglass
                                    
            '*** Guardar Relacionado ***
            With adoComm
                .CommandText = "{ call up_GNManInstitucionPersona('" & _
                    strCodInstitucion & "','" & Codigo_Tipo_Persona_Relacionado & "','" & _
                    strCodClasePersona & "','" & strCodTipoDocumento & "','" & _
                    Trim(txtNumIdentidad.Text) & "','" & Trim(txtApellidoPaterno.Text) & "','" & _
                    Trim(txtApellidoMaterno.Text) & "','" & Trim(txtNombres.Text) & "','" & _
                    Trim(txtRazonSocial.Text) & "','"
                If strCodClasePersona = Codigo_Persona_Juridica Then
                    .CommandText = .CommandText & Trim(txtRazonSocial.Text) & "','"
                Else
                    .CommandText = .CommandText & Trim(txtApellidoPaterno.Text) & Space(1) & Trim(txtApellidoMaterno.Text) & Space(1) & Trim(txtNombres.Text) & "','"
                End If
                
                .CommandText = .CommandText & "','','','','','','','','','','','','','','" & _
                    "','',0,0,0,0,'','','" & _
                    Convertyyyymmdd(dtpFechaInicio.Value) & "','" & Convertyyyymmdd(dtpFechaFin.Value) & "','" & _
                    "999','99999999','',0,0,'','','','X','','','','','','','','','','" & _
                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','I') }"
                adoConn.Execute .CommandText
                
                '*** Detalle Relación ***
                intContador = tdgRelacion.SelBookmarks.Count - 1

                For intRegistro = 0 To intContador
                    
                    adoRelacion.Recordset.MoveFirst
                    adoRelacion.Recordset.Move (tdgRelacion.SelBookmarks(intRegistro) - 1)
                        
                    .CommandText = "INSERT INTO ParticipeRelacion VALUES ('" & _
                        strCodInstitucion & "','" & Codigo_Tipo_Persona_Relacionado & "','" & _
                        tdgRelacion.Columns(0).Value & "')"
                    adoConn.Execute .CommandText
                                                            
                Next

            End With

            Me.MousePointer = vbDefault
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"

            cmdOpcion.Visible = True
            With tabPersona
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
    End If
    
    If strEstado = Reg_Edicion Then
        If TodoOK() Then
            If MsgBox(Mensaje_Edicion, vbQuestion + vbYesNo, gstrNombreEmpresa) = vbNo Then Exit Sub
            
            Dim intValorRetorno     As Integer
            
            Me.MousePointer = vbHourglass
                                                                        
            '*** Actualizar Relacionado ***
            With adoComm
                .CommandText = "{ call up_GNManInstitucionPersona('" & _
                    Trim(tdgConsulta.Columns(0)) & "','" & Codigo_Tipo_Persona_Relacionado & "','" & _
                    strCodClasePersona & "','" & strCodTipoDocumento & "','" & _
                    Trim(txtNumIdentidad.Text) & "','" & Trim(txtApellidoPaterno.Text) & "','" & _
                    Trim(txtApellidoMaterno.Text) & "','" & Trim(txtNombres.Text) & "','" & _
                    Trim(txtRazonSocial.Text) & "','"
                If strCodClasePersona = Codigo_Persona_Juridica Then
                    .CommandText = .CommandText & Trim(txtRazonSocial.Text) & "','"
                Else
                    .CommandText = .CommandText & Trim(txtApellidoPaterno.Text) & Space(1) & Trim(txtApellidoMaterno.Text) & Space(1) & Trim(txtNombres.Text) & "','"
                End If
                
                .CommandText = .CommandText & "','','','','','','','','','','','','','','" & _
                    "','',0,0,0,0,'','','" & _
                    Convertyyyymmdd(dtpFechaInicio.Value) & "','" & Convertyyyymmdd(dtpFechaFin.Value) & "','" & _
                    "999','99999999','',0,0,'','','','X','','','','','','','','','','" & _
                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','U') }"
                adoConn.Execute .CommandText
                
                
                '*** Detalle Relación ***
                .CommandText = "DELETE ParticipeRelacion " & _
                        "WHERE CodPersona='" & Trim(tdgConsulta.Columns(0)) & "' AND " & _
                        "TipoPersona='" & Codigo_Tipo_Persona_Relacionado & "'"
                    adoConn.Execute .CommandText
                    
                intContador = tdgRelacion.SelBookmarks.Count - 1
                                
                For intRegistro = 0 To intContador
                                    
                    adoRelacion.Recordset.MoveFirst
                    adoRelacion.Recordset.Move (tdgRelacion.SelBookmarks(intRegistro) - 1)
                                            
                    .CommandText = "INSERT INTO ParticipeRelacion VALUES ('" & _
                        Trim(tdgConsulta.Columns(0)) & "','" & Codigo_Tipo_Persona_Relacionado & "','" & _
                        tdgRelacion.Columns(0).Value & "')"
                    adoConn.Execute .CommandText

                Next
                                
            End With

            Me.MousePointer = vbDefault
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"

            cmdOpcion.Visible = True
            With tabPersona
                .TabEnabled(0) = True
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
    
    If cboTipoDocumento.ListIndex = 0 Then
        MsgBox "Seleccione el Tipo de Documento.", vbCritical
        cboTipoDocumento.SetFocus
        Exit Function
    End If
    
    If Trim(txtNumIdentidad.Text) = Valor_Caracter Then
        MsgBox "El Campo Número de Documento no es Válido!.", vbCritical
        txtNumIdentidad.SetFocus
        Exit Function
    End If
    
    If strCodClasePersona = Codigo_Persona_Natural Then
        If Trim(txtApellidoPaterno.Text) = "" Then
            MsgBox "El Campo Apellido Paterno no es Válido!.", vbCritical
            txtApellidoPaterno.SetFocus
            Exit Function
        End If
        
        If Trim(txtApellidoMaterno.Text) = Valor_Caracter Then
            If MsgBox("El Campo Apellido Materno no es Válido!." & vbNewLine & vbNewLine & _
                "Seguro de Continuar ?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbNo Then
                                
                txtApellidoMaterno.SetFocus
                Exit Function
            End If
        End If
        
        If Trim(txtNombres.Text) = Valor_Caracter Then
            MsgBox "El Campo Nombres no es Válido!.", vbCritical
            txtNombres.SetFocus
            Exit Function
        End If
    Else
        If Trim(txtRazonSocial.Text) = Valor_Caracter Then
            MsgBox "El Campo Razón Social no es Válido!.", vbCritical
            txtRazonSocial.SetFocus
            Exit Function
        End If
    End If
    
'    If strCodTipoRelacion = Valor_Caracter Then
'        MsgBox "Seleccione el Tipo de Relación.", vbCritical
'        cboTipoRelacion.SetFocus
'        Exit Function
'    End If
    
'    If cboCargoRelacion.ListCount > 1 Then
'        If strCodCargoRelacion = Valor_Caracter Then
'            MsgBox "Seleccione el Cargo de la Relación.", vbCritical
'            cboCargoRelacion.SetFocus
'            Exit Function
'        End If
'    End If
                                                                        
    '*** Si todo pasó OK ***
    TodoOK = True

End Function
Private Sub Habilita()

    
End Sub


Public Sub Imprimir()

    'Call SubImprimir(1)
    
End Sub


Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Lista de Relacionados"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    
End Sub
Public Sub Modificar()

    If strEstado = Reg_Consulta Then
    
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabPersona
            .TabEnabled(0) = False
            .Tab = 1
        End With
        
    End If
    
End Sub





Public Sub Salir()

    Unload Me
    
End Sub



Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String

    If tabPersona.Tab = 1 Then Exit Sub
    
    Select Case Index
        Case 1
            gstrNameRepo = "ParticipeRelacion"
                        
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
                        
            aReportParamS(0) = Codigo_Tipo_Persona_Relacionado
            aReportParamS(1) = strCodRelacion
            
    End Select

    gstrSelFrml = ""
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub
Private Sub cboClasePersona_Click()

    Dim strSQL As String
    
    strCodClasePersona = Valor_Caracter
    If cboClasePersona.ListIndex < 0 Then Exit Sub
    
    strCodClasePersona = Trim(arrClasePersona(cboClasePersona.ListIndex))
    
    '*** Tipo Documento Identidad ***
    strSQL = "{ call up_ACSelDatosParametro(4,'" & strCodClasePersona & "') }"
    CargarControlLista strSQL, cboTipoDocumento, arrTipoDocumento(), Sel_Defecto
    
    If cboTipoDocumento.ListCount > 0 Then cboTipoDocumento.ListIndex = 0
    
    If strCodClasePersona = Codigo_Persona_Natural Then
        txtApellidoPaterno.Enabled = True
        txtApellidoMaterno.Enabled = True
        txtNombres.Enabled = True
        Call ColorControlHabilitado(txtApellidoPaterno)
        Call ColorControlHabilitado(txtApellidoMaterno)
        Call ColorControlHabilitado(txtNombres)
        txtRazonSocial.Enabled = False
        Call ColorControlDeshabilitado(txtRazonSocial)
    Else
        txtApellidoPaterno.Enabled = False
        txtApellidoMaterno.Enabled = False
        txtNombres.Enabled = False
        Call ColorControlDeshabilitado(txtApellidoPaterno)
        Call ColorControlDeshabilitado(txtApellidoMaterno)
        Call ColorControlDeshabilitado(txtNombres)
        txtRazonSocial.Enabled = True
        Call ColorControlHabilitado(txtRazonSocial)
    End If
            
End Sub



Private Sub cboTipoDocumento_Click()

    strCodTipoDocumento = Valor_Caracter
    If cboTipoDocumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoDocumento = Trim(arrTipoDocumento(cboTipoDocumento.ListIndex))
    txtNumIdentidad.Text = Valor_Caracter
    txtNumIdentidad.MaxLength = ObtenerNumMaximoDocumentoIdentidad(strCodTipoDocumento)
            
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
    
    For Each elemento In Me.Controls
    
        If TypeOf elemento Is TDBGrid Then
            Call FormatoGrilla(elemento)
        End If
    
    Next
            
End Sub
Private Sub CargarListas()

    Dim strSQL  As String
    
    '*** Clase Persona ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='CLSPER' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboClasePersona, arrClasePersona(), Valor_Caracter
                        
End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabPersona.Tab = 0
    tabPersona.TabEnabled(1) = False
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 14
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 70
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmRelacionados = Nothing
    
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




Public Sub Adicionar()

    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Persona Relacionada..."
    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabPersona
        .TabEnabled(0) = False
        .TabEnabled(1) = True
        .Tab = 1
    End With
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro     As ADODB.Recordset
    Dim intRegistro     As Integer
    
    Select Case strModo
        Case Reg_Adicion
            Set adoRegistro = New ADODB.Recordset
            
            adoComm.CommandText = "SELECT COUNT(*) SecuencialInstitucion FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Relacionado & "'"
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                strCodInstitucion = Format(adoRegistro("SecuencialInstitucion"), "00000000")
            Else
                strCodInstitucion = "00000001"
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
            
            cboClasePersona.ListIndex = -1
            intRegistro = ObtenerItemLista(arrClasePersona(), Codigo_Persona_Natural)
            If intRegistro >= 0 Then cboClasePersona.ListIndex = intRegistro
            
            cboTipoDocumento.ListIndex = -1
            If cboTipoDocumento.ListCount > 0 Then cboTipoDocumento.ListIndex = 0
                        
            txtNumIdentidad.Text = Valor_Caracter
            txtApellidoPaterno.Text = Valor_Caracter
            txtApellidoMaterno.Text = Valor_Caracter
            txtNombres.Text = Valor_Caracter
            txtRazonSocial.Text = Valor_Caracter
            
            Call BuscarRelacion
                                    
            cboTipoDocumento.SetFocus
                        
        Case Reg_Edicion
            Dim strCodPersona   As String

            Set adoRegistro = New ADODB.Recordset

            strCodPersona = Trim(tdgConsulta.Columns(0))
            
            adoComm.CommandText = "SELECT IP.CodPersona,DescripPersona,TipoIdentidad,NumIdentidad,CodTipoRelacion," & _
                "ApellidoPaterno,ApellidoMaterno,Nombres,RazonSocial,ClasePersona,FechaInicioRelacion,FechaFinalRelacion " & _
                "FROM InstitucionPersona IP LEFT JOIN ParticipeRelacion PR ON(PR.CodPersona=IP.CodPersona AND PR.TipoPersona=IP.TipoPersona) " & _
                "WHERE IP.TipoPersona='" & Codigo_Tipo_Persona_Relacionado & "' AND " & _
                "IP.CodPersona='" & strCodPersona & "'"
            Set adoRegistro = adoComm.Execute

            If Not adoRegistro.EOF Then
                intRegistro = ObtenerItemLista(arrClasePersona(), adoRegistro("ClasePersona"))
                If intRegistro >= 0 Then cboClasePersona.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrTipoDocumento(), adoRegistro("TipoIdentidad"))
                If intRegistro >= 0 Then cboTipoDocumento.ListIndex = intRegistro
                
                txtNumIdentidad.Text = Trim(adoRegistro("NumIdentidad"))
                txtApellidoPaterno.Text = Trim(adoRegistro("ApellidoPaterno"))
                txtApellidoMaterno.Text = Trim(adoRegistro("ApellidoMaterno"))
                txtNombres.Text = Trim(adoRegistro("Nombres"))
                txtRazonSocial.Text = Trim(adoRegistro("RazonSocial"))
                
                dtpFechaInicio.Value = CVDate(adoRegistro("FechaInicioRelacion"))
                dtpFechaFin.Value = CVDate(adoRegistro("FechaFinalRelacion"))
                
                Call BuscarRelacion
                
                Do While Not adoRegistro.EOF
                    adoRelacion.Recordset.MoveFirst
                    Do While Not adoRelacion.Recordset.EOF
                        If adoRegistro("CodTipoRelacion") = adoRelacion.Recordset.Fields("CodParametro").Value Then
                            tdgRelacion.SelBookmarks.Add adoRelacion.Recordset.Bookmark
                        End If
                                            
                        adoRelacion.Recordset.MoveNext
                    Loop
                    adoRegistro.MoveNext
                Loop
                
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
    End Select
    
End Sub

Private Sub tabPersona_Click(PreviousTab As Integer)

    Select Case tabPersona.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabPersona.Tab = 0
        
    End Select
    
End Sub

Private Sub txtApellidoMaterno_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub


Private Sub txtNombres_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtRazonSocial_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
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
