VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGeneraArchivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Archivos Regulatorios"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   8970
   Begin VB.CommandButton cmdSalr 
      Caption         =   "Salir"
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
      Left            =   6900
      Picture         =   "frmGeneraArchivo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Generar Archivos"
      Top             =   1320
      Width           =   1200
   End
   Begin VB.CommandButton CmdGenerarArchivos 
      Caption         =   "Generar"
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
      Left            =   5550
      Picture         =   "frmGeneraArchivo.frx":0582
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Generar Archivos"
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Frame fraCriterio 
      Caption         =   "Criterios para la Generación"
      Height          =   2205
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   8865
      Begin VB.CommandButton cmd_Listar 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   315
         Left            =   7770
         TabIndex        =   9
         Top             =   780
         Width           =   315
      End
      Begin VB.TextBox txtDestino 
         Height          =   315
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "C:\"
         Top             =   780
         Width           =   6165
      End
      Begin MSComCtl2.DTPicker dtpFechaDesde 
         Height          =   315
         Left            =   1530
         TabIndex        =   4
         Top             =   360
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         Format          =   113967105
         CurrentDate     =   38779
      End
      Begin MSComCtl2.DTPicker dtpFechaHasta 
         Height          =   315
         Left            =   6540
         TabIndex        =   5
         Top             =   360
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         Format          =   113967105
         CurrentDate     =   38779
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   5490
         TabIndex        =   7
         Top             =   420
         Width           =   825
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   390
         TabIndex        =   6
         Top             =   420
         Width           =   900
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Destino"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   420
         TabIndex        =   2
         Top             =   780
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmGeneraArchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' -- Variables para la conexión y el recordset
Private cn      As New ADODB.Connection
Dim adoRegistro As ADODB.Recordset

Private Sub cmd_Listar_Click()

   frm_ListaDir.Show 1
   If gs_FormName <> "" And gs_FormName <> "c:\Scotia" + CStr(Day(DateTime.Now())) + "" + CStr(Month(DateTime.Now())) + "" + CStr(Year(DateTime.Now())) + ".txt" Then
      txtDestino.Text = gs_FormName + "\Scotia" + CStr(Day(DateTime.Now())) + "" + CStr(Month(DateTime.Now())) + "" + CStr(Year(DateTime.Now())) + ".txt"
   End If
   If gs_FormName = "c:\Scotia" + CStr(Day(DateTime.Now())) + "" + CStr(Month(DateTime.Now())) + "" + CStr(Year(DateTime.Now())) + ".txt" Then
      txtDestino.Text = gs_FormName
   End If

End Sub
 

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub CmdGenerarArchivos_Click()

    Set adoRegistro = New ADODB.Recordset
            
    adoComm.CommandText = "{ call up_PRGeneraArchivoScotia('" + Convertyyyymmdd(dtpFechaDesde.Value) + "','" + Convertyyyymmdd(dtpFechaHasta.Value) + "') }"
    Set adoRegistro = adoComm.Execute
  
    Call Exportar_Recordset(adoRegistro, txtDestino.Text, vbTab)
  
End Sub

Private Sub cmdSalr_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    CentrarForm Me
    'InicializarValores
    txtDestino.Text = "c:\Scotia" + CStr(Day(DateTime.Now())) + "" + CStr(Month(DateTime.Now())) + "" + CStr(Year(DateTime.Now())) + ".txt"

    dtpFechaDesde.Value = gdatFechaActual
    dtpFechaHasta.Value = gdatFechaActual
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
End Sub


Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vSearch
            'Call Buscar
        Case vReport
'            Call Imprimir
        Case vExit
            Call Salir
        
    End Select
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

  
' --------------------------------------------------------------------------------
' \\ -- Función que exporta el recordset a un archivo de texto
' --------------------------------------------------------------------------------
Public Function Exportar_Recordset( _
    rs As Object, _
    Optional sFileName As String, _
    Optional sDelimiter As String = " ", _
    Optional bPrintField As Boolean = False) As Boolean
  
    Dim iFreeFile   As Integer
    Dim iField      As Long
    Dim i           As Long
    Dim obj_Field   As ADODB.Field
  
    On Error GoTo error_handler:
      
    Screen.MousePointer = vbHourglass
    ' -- Otener número de archivo disponible
    iFreeFile = FreeFile
    ' -- Crear el archivo
    Open sFileName For Output As #iFreeFile
  
    With rs
        iField = .Fields.Count - 1
        On Error Resume Next
        ' -- Primer registro
        .MoveFirst
        On Error GoTo error_handler
        ' -- Recorremos campo por campo y los registros de cada uno
        Do While Not .EOF
            For i = 0 To iField
                  
                ' -- Asigna el objeto Field
                Set obj_Field = .Fields(i)
                ' -- Verificar que el campo no es de ipo bunario o  un tipo no válido para grabar en el archivo
                If isValidField(obj_Field) Then
                    If i < iField Then
                        If bPrintField Then
                            ' -- Escribir el campo y el valor
                            'Print #iFreeFile, obj_Field.Name & ":" & obj_Field.Value & sDelimiter;
                            Print #iFreeFile, obj_Field.Name & ":" & obj_Field.Value;
                        Else
                            ' -- Guardar solo el valor sin el campo
                            'Print #iFreeFile, obj_Field.Value & sDelimiter;
                            Print #iFreeFile, obj_Field.Value;
                        End If
                    Else
                        If bPrintField Then
                            ' -- Escribir el nombre del campo y el valor de la última columna ( Sin delimitador y sin punto y coma para añadir nueva línea )
                            Print #iFreeFile, obj_Field.Name & ": " & obj_Field.Value
                        Else
                            ' -- Guardar solo el valor sin el campo
                            Print #iFreeFile, obj_Field.Value
                        End If
                    End If
                End If
            Next
            ' -- Mover el cursor al siguiente registro
            .MoveNext
        Loop
    End With
      
    ' -- Cerrar el recordset
    adoRegistro.Close
    Exportar_Recordset = True
    Screen.MousePointer = vbDefault
    MsgBox "El Archivo fue generado correctamente en la ruta : " + sFileName, vbOKOnly
    Close #iFreeFile
    Exit Function
error_handler:
 On Error Resume Next
 Close #iFreeFile
    adoRegistro.Close: Set adoRegistro = Nothing
 Screen.MousePointer = vbDefault
End Function
  
' ----------------------------------------------------------------------------------------------
' -- Si el campo es nulo ( binario, o tipo desconocido etc..) devuelve False para no añadir el dato
' ----------------------------------------------------------------------------------------------
Private Function isValidField(obj_Field As ADODB.Field) As Boolean
      
    With obj_Field
        On Error GoTo error_handler
        Select Case obj_Field.Type
            Case adBinary, adIDispatch, adIUnknown, adUserDefined
                isValidField = False
            ' -- Campo válido
            Case Else
                isValidField = True
        End Select
    End With
Exit Function
error_handler:
End Function
 
' ---------------------------------------------------------------------------------
' \\ -- Cerrar la base de datos y el recordset al finalizar
' ---------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    If Not adoRegistro Is Nothing Then
        If adoRegistro.State = adStateOpen Then adoRegistro.Close
        Set adoRegistro = Nothing
    End If
    
End Sub

