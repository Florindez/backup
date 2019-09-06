VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmFiltroReporte3 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFiltroReporte3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2910
      Picture         =   "frmFiltroReporte3.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1620
      Width           =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1380
      Picture         =   "frmFiltroReporte3.frx":056E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1620
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1455
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5325
      Begin VB.TextBox txtDestino 
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "C:\"
         Top             =   960
         Width           =   3855
      End
      Begin VB.CommandButton cmd_Listar 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   315
         Left            =   4860
         TabIndex        =   7
         Top             =   945
         Width           =   315
      End
      Begin MSComCtl2.DTPicker dtpFechaDesde 
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   210
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   202178561
         CurrentDate     =   38068
      End
      Begin MSComCtl2.DTPicker dtpFechaHasta 
         Height          =   315
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   202178561
         CurrentDate     =   38068
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destino"
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
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label lblHasta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
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
         Left            =   180
         TabIndex        =   2
         Top             =   660
         Width           =   765
      End
      Begin VB.Label lblDesde 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
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
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmFiltroReporte3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim strNombreArchivo    As String
Dim RptPath             As String



Private Sub cmdAceptar_Click()
    Dim n_dias, res As Integer, gn As Integer
    Dim NomFile As String
    Dim i As Integer
    
    'Separador = Trim(Me.txtSeparador.Text)
    
    On Error GoTo CmdGenerarArchivos_Exit
   
    Call GenerarArchivoFact(dtpFechaDesde.Value, dtpFechaHasta.Value)
    
    MsgBox "Archivo(s) generado(s) correctamente.", 48
    
    Unload Me
    
    Exit Sub
    
CmdGenerarArchivos_Exit:
    Me.MousePointer = 0
    MsgBox Error(err)
    MsgBox "Error de generación de archivos.", 48
    Exit Sub

    'GeneraArchivoRegulatorio_Participes (Separador)
    'GeneraArchivoRegulatorio_SaldosContables (Separador)
    'GeneraArchivoRegulatorio_ValorCuotayCuentasPrincipales (Separador)

    'MsgBox "Se genero el Archivo"

End Sub

Private Sub GenerarArchivoFact(FechaDesde As Date, FechaHasta As Date)
    Dim adoAux As ADODB.Recordset
    Dim res As Integer
    Dim strAnio As String
    Dim strMes As String
    Dim strDia As String
    Dim strName2 As String
    Dim s_fecha As String
    
    '----------------------------------------------------
    'DR 04/05/99 NOMENCLATURA DEL ARCHIVO A GENERAR
    '----------------------------------------------------
    RptPath = Trim(txtDestino.Text)
    
    adoComm.CommandText = "select NumRucFondo from Fondo where CodFondo = '" & gstrCodFondoContable & "'"
    Set adoAux = adoComm.Execute
    
    strAnio = Year(FechaDesde)
    strMes = Format(Month(FechaDesde), "00")
    strDia = Format(Day(FechaDesde), "00")
    
    s_fecha = strAnio + strMes + strDia
    
    strAnio = Year(FechaHasta)
    strMes = Format(Month(FechaHasta), "00")
    strDia = Format(Day(FechaHasta), "00")

    
    strName2 = "-" & strAnio & strMes & strDia
    
    'ID_File = Mid$(Trim$(File), 1, 2)
    '----------------------------------------------------------
    'GENERANDO EL NOMBRE DEL ARCHIVO
    '----------------------------------------------------------
    strNombreArchivo = RptPath + "FACT" + Trim$(adoAux("NumRucFondo")) + "-" + s_fecha + strName2 + ".csv"
    
    Call GeneraArchivoFacturacion(FechaDesde, FechaHasta)
    
End Sub

Private Sub GeneraArchivoFacturacion(ByVal FechaDesde As Date, ByVal FechaHasta As Date)
 
    Dim numCampos As Integer
    Dim n As Integer
    Dim strRegistro As String
    Dim strFechaRegistro As String
    Dim numNumRegistroControl As Long
    Dim numCantRegistros As Long
    Dim numOut As Integer
    
    Me.MousePointer = vbHourglass
    
    strFechaRegistro = Convertyyyymmdd(FechaDesde)
    
    numCantRegistros = 0
    
    
    Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
    
    numCampos = 25
    strRegistro = ""
    
    With adoComm
          .CommandText = "{call up_RepRegistroVentaFact "
        
        .CommandText = .CommandText & "('" & gstrCodFondoContable & "','" & gstrCodAdministradora & "','" & Convertyyyymmdd(dtpFechaDesde.Value) & "','" & Convertyyyymmdd(dtpFechaHasta.Value) & "')}"
        
        Set adoRegistro = .Execute
        
        numOut = FreeFile
        Open strNombreArchivo For Binary Access Read Write As numOut


        Do While Not adoRegistro.EOF
            
            For n = 0 To numCampos - 1
                strRegistro = strRegistro & Trim$(adoRegistro.Fields(n).Value) & ","
            Next n
                
            strRegistro = strRegistro + Chr(13) + Chr(10)
                
            'Print #1, strRegistro
            Put numOut, , strRegistro
        
            strRegistro = ""
            
            numCantRegistros = numCantRegistros + 1
        
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
'    'ACTUALIZAR REGISTRO DE CONTROL
'    strRegistro = ""
'    strRegistro = strRegistro + "C"
'    strRegistro = strRegistro + Format(Trim(strCodAdministradoraCNS), "@@@@@@")
'    strRegistro = strRegistro + strCodRegistro
'    strRegistro = strRegistro + Right(strFechaRegistro, 2) + Mid(strFechaRegistro, 5, 2) + Left(strFechaRegistro, 4)
'    strRegistro = strRegistro + Format(numNumRegistroControl, "0000")
'    strRegistro = strRegistro + Format(numCantRegistros, "00000")
'
    Put numOut, 1, strRegistro
    
    Close #1
    'Libro 8.2 vacio:
'    If strCodRegistro = "RC" Then
'        strNombreArchivo = Left(strNombreArchivo, Len(strNombreArchivo) - 15) & "80200001011.TXT"
'        numOut = FreeFile
'        Open strNombreArchivo For Binary Access Read Write As numOut
'
'    '    Call ActualizaRegistroControl(strCodRegistro, gstrCodAdministradora, strFechaRegistro, numNumRegistroControl, numCantRegistros)
'        Put numOut, 1, strRegistro
'
'        Close #1
'    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    dtpFechaDesde.Value = gdatFechaActual
    dtpFechaHasta.Value = gdatFechaActual
End Sub
