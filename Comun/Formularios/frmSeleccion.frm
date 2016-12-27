VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmSeleccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Entidad de Trabajo"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
      Bindings        =   "frmSeleccion.frx":0000
      Height          =   2295
      Left            =   240
      OleObjectBlob   =   "frmSeleccion.frx":001A
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   375
      Left            =   240
      Top             =   2520
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSeleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL              As String, strCodFondo          As String
Dim strEstado           As String

Public Sub Buscar()

    If gstrTipoAdministradoraContable = Codigo_Tipo_Fondo_Administradora Then
        strSQL = "SELECT CodAdministradora Codigo,DescripAdministradora Descripcion FROM Administradora " & _
            "WHERE CodTipoAdministradora='" & Codigo_Tipo_Fondo_Administradora & "' AND " & _
            "Estado='" & Estado_Activo & "' ORDER BY DescripAdministradora"
    Else
        strSQL = "SELECT CodFondo Codigo,DescripFondo Descripcion FROM Fondo " & _
            "WHERE CodAdministradora='" & gstrCodAdministradora & "' AND " & _
            "Estado='" & Estado_Activo & "' ORDER BY DescripFondo"
    End If
    
    strEstado = Reg_Defecto
    With adoConsulta
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With

    tdgConsulta.Refresh

    If adoConsulta.Recordset.RecordCount > 0 Then strEstado = Reg_Consulta
    
End Sub


Private Sub InicializarValores()

    strCodFondo = "000"
    
End Sub

Private Sub Form_Load()

    Call Buscar
    
    CentrarForm Me
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set frmSeleccion = Nothing
    
End Sub


Private Sub tdgConsulta_DblClick()

    If strEstado = Valor_Caracter Then Exit Sub
    
    If gstrTipoAdministradoraContable = Codigo_Tipo_Fondo_Administradora Then
        Dim adoRegistro     As ADODB.Recordset
        
        gstrCodAdministradoraContable = tdgConsulta.Columns(0).Value
        frmMainMdi.txtEmpresa.Text = Trim(tdgConsulta.Columns(1).Value)
        
        With adoComm
            '*** Actualizar Administradora por defecto ***
            .CommandText = "UPDATE Administradora SET IndDefecto='' WHERE IndDefecto='" & Valor_Indicador & "'"
            adoConn.Execute .CommandText
            
            .CommandText = "UPDATE Administradora SET IndDefecto='" & Valor_Indicador & "' " & _
                "WHERE CodTipoAdministradora='" & Codigo_Tipo_Fondo_Administradora & "' AND " & _
                "CodAdministradora='" & tdgConsulta.Columns(0).Value & "'"
            adoConn.Execute .CommandText
            
            Set adoRegistro = New ADODB.Recordset
            '*** Actualizar Fecha de trabajo ***
            .CommandText = "SELECT FechaContable,ValorTipoCambio,CodMoneda FROM AdministradoraCalendario " & _
                "WHERE CodAdministradora='" & gstrCodAdministradoraContable & "' AND CodFondo='" & strCodFondo & "' AND " & _
                "IndAbierto='" & Valor_Indicador & "'"
                Set adoRegistro = .Execute
                
                If Not adoRegistro.EOF Then
                    gdatFechaActual = adoRegistro("FechaContable"): gdblTipoCambio = CDbl(adoRegistro("ValorTipoCambio"))
                    gstrFechaActual = Convertyyyymmdd(adoRegistro("FechaContable"))
                    gstrCodMoneda = adoRegistro("CodMoneda")
                        
                    frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
                End If
                adoRegistro.Close: Set adoRegistro = Nothing
        End With
    Else
        gstrCodFondoContable = tdgConsulta.Columns(0).Value
        frmMainMdi.txtEmpresa.Text = Trim(tdgConsulta.Columns(1).Value)
    End If
    
    Unload Me
    
End Sub


