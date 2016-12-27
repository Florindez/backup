VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{442CAE95-1D41-47B1-BE83-6995DA3CE254}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmBusquedaCuenta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas Contables"
   ClientHeight    =   3480
   ClientLeft      =   1995
   ClientTop       =   2265
   ClientWidth     =   6255
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
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3480
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TAMControls2.ucBotonEdicion2 cmdAccion 
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   1296
      Buttons         =   3
      Caption0        =   "&Buscar"
      Tag0            =   "5"
      Visible0        =   0   'False
      ToolTipText0    =   "Buscar"
      Caption1        =   "&Seleccionar"
      Tag1            =   "3"
      Visible1        =   0   'False
      ToolTipText1    =   "Seleccionar"
      Caption2        =   "&Cancelar"
      Tag2            =   "9"
      Visible2        =   0   'False
      ToolTipText2    =   "Cancelar"
      UserControlWidth=   4200
   End
   Begin TrueOleDBGrid60.TDBGrid tdgCuenta 
      Bindings        =   "frmBusquedaCuenta.frx":0000
      Height          =   1815
      Left            =   165
      OleObjectBlob   =   "frmBusquedaCuenta.frx":0018
      TabIndex        =   3
      Top             =   720
      Width           =   5895
   End
   Begin VB.TextBox txtCodCuenta 
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
      Left            =   1590
      MaxLength       =   10
      TabIndex        =   0
      Top             =   210
      Width           =   1545
   End
   Begin MSAdodcLib.Adodc adoCuenta 
      Height          =   330
      Left            =   4800
      Top             =   2760
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
      Caption         =   "adoConsulta"
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
   Begin VB.Label lblDescrip 
      Caption         =   "Cuenta"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmBusquedaCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strEstado As String

Private Sub Form_Load()

    Call InicializarValores
    
    CentrarForm Me
    
End Sub

Private Sub InicializarValores()

    strEstado = Reg_Defecto
    txtCodCuenta.Text = Valor_Caracter
    
    Set cmdAccion.FormularioActivo = Me
    
End Sub
Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vModify
            Call Modificar
        Case vSearch
            Call Buscar
        Case vExit
            Call Salir
        
    End Select
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub Buscar()
        
    Dim strSQL As String
                                                                                    
    Me.MousePointer = vbHourglass
                
    If Trim(txtCodCuenta.Text) <> "" Then
        strSQL = "SELECT CodCuenta,DescripCuenta FROM PlanContable " & _
            "WHERE CodCuenta LIKE '" & Trim(txtCodCuenta.Text) & "%' AND " & _
            "IndMovimiento='X' AND CodAdministradora='" & gstrCodAdministradora & "'"
                
        With adoCuenta
            .ConnectionString = gstrConnectConsulta
            .RecordSource = strSQL
            .Refresh
        End With
        
        tdgCuenta.Refresh
        
        If adoCuenta.Recordset.RecordCount > 0 Then strEstado = Reg_Consulta
        
    End If
    
    Me.MousePointer = vbDefault
                                    
End Sub

Public Sub Modificar()

    Dim intRegistro As Integer
    
    If strEstado = Reg_Consulta Then
        Select Case gstrFormulario
            Case "frmAsientoContable"
                frmAsientoContable.txtCodCuenta = Trim(tdgCuenta.Columns(0))
                frmAsientoContable.txtDescripMovimiento = Trim(tdgCuenta.Columns(1))
            Case "frmFiltroReporte"
                frmFiltroReporte.txtCodCuenta = Trim(tdgCuenta.Columns(0))
            Case "frmMovimientoAnalitica"
                frmMovimientoAnalitica.txtCodCuenta = Trim(tdgCuenta.Columns(0))
            Case "frmDinamicaContable"
                'frmDinamicaContable.lblCuenta = Trim(tdgCuenta.Columns(0))
                'frmDinamicaContable.txtDescripParametro = Trim(tdgCuenta.Columns(1))
                
        End Select
        
        Call Salir
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set frmBusquedaCuenta = Nothing
    
End Sub
