VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCalendario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendario"
   ClientHeight    =   8445
   ClientLeft      =   1080
   ClientTop       =   1470
   ClientWidth     =   14880
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
   Icon            =   "frmCalendario.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8445
   ScaleWidth      =   14880
   Begin MSComCtl2.MonthView mthCalendario 
      Height          =   7500
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   13470
      _ExtentX        =   23760
      _ExtentY        =   13229
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthColumns    =   4
      MonthRows       =   3
      MonthBackColor  =   -2147483629
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   146669569
      TitleBackColor  =   -2147483645
      TitleForeColor  =   0
      CurrentDate     =   39086
      MinDate         =   2
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   12480
      TabIndex        =   0
      Top             =   7920
      Width           =   1215
   End
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Adicionar()

End Sub

Public Sub Buscar()

End Sub

Private Sub CargarDiasNoLaborables(datFecha As Date)

    Dim adoConsulta     As ADODB.Recordset
    Dim lngPeriodo      As Long
    Dim datFechaFeriado As Date
    
    lngPeriodo = Year(datFecha)
    
    Set adoConsulta = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT FechaFeriado FROM CalendarioNoLaborable WHERE YEAR(FechaFeriado)=" & lngPeriodo
        Set adoConsulta = .Execute
        
        Do While Not adoConsulta.EOF
            datFechaFeriado = CVDate(adoConsulta("FechaFeriado"))
            
            mthCalendario_DateClick (datFechaFeriado)
            
            adoConsulta.MoveNext
        Loop
        adoConsulta.Close: Set adoConsulta = Nothing
    End With
        
End Sub

Public Sub Eliminar()

End Sub

Public Sub Grabar()

End Sub

Public Sub Imprimir()

End Sub

Private Sub InicializarValores()

    mthCalendario.Value = gdatFechaActual
    
End Sub

Public Sub Modificar()

End Sub

Public Sub Salir()

    Unload Me

End Sub


Private Sub cmdSalir_Click()

    Unload Me
    
End Sub


Private Sub Form_Load()

    Call InicializarValores
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
            
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set frmCalendario = Nothing
    
End Sub






Private Sub mthCalendario_DateClick(ByVal DateClicked As Date)

    If Weekday(DateClicked) = vbSaturday Then Exit Sub
    If Weekday(DateClicked) = vbSunday Then Exit Sub
    
    Dim strFechaGrabar  As String, strFechaSiguiente    As String
    
    mthCalendario.DayBold(DateClicked) = Not mthCalendario.DayBold(DateClicked)
    strFechaGrabar = Convertyyyymmdd(DateClicked)
    strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, DateClicked))
    
    With adoComm
        If mthCalendario.DayBold(DateClicked) Then
            .CommandText = "INSERT INTO CalendarioNoLaborable VALUES ('" & strFechaGrabar & "')"
        Else
            .CommandText = "DELETE CalendarioNoLaborable WHERE FechaFeriado >='" & strFechaGrabar & "' AND FechaFeriado <'" & strFechaSiguiente & "'"
        End If
        adoConn.Execute .CommandText
    End With
    
End Sub

Private Sub mthCalendario_GetDayBold(ByVal StartDate As Date, ByVal Count As Integer, State() As Boolean)

    Dim intContador As Integer
    Dim adoConsulta     As ADODB.Recordset
    Dim lngPeriodo      As Long
    Dim datFechaFeriado As Date
    Dim datFechaTemporal As Date
    
    lngPeriodo = Year(gdatFechaActual)
    
    Set adoConsulta = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT FechaFeriado FROM CalendarioNoLaborable WHERE YEAR(FechaFeriado)=" & lngPeriodo
        Set adoConsulta = .Execute
        
        Do While Not adoConsulta.EOF
            datFechaFeriado = CVDate(adoConsulta("FechaFeriado"))
                        
            intContador = Weekday(datFechaFeriado)
            datFechaTemporal = StartDate
            While Weekday(datFechaTemporal) <> intContador
                datFechaTemporal = DateAdd("d", 1, datFechaTemporal)
            Wend
            
            While intContador < Count
                If datFechaTemporal >= datFechaFeriado Then
                    State(intContador - mthCalendario.StartOfWeek) = True
                    intContador = Count
                Else
                    intContador = intContador + 7
                    datFechaTemporal = DateAdd("d", 7, datFechaTemporal)
                End If
            Wend
            
            adoConsulta.MoveNext
        Loop
        adoConsulta.Close: Set adoConsulta = Nothing
    End With
    
    intContador = vbSaturday
    While intContador < Count
        State(intContador - mthCalendario.StartOfWeek) = True
        intContador = intContador + 7
    Wend
    
    intContador = vbSunday
    While intContador < Count
        State(intContador - mthCalendario.StartOfWeek) = True
        intContador = intContador + 7
    Wend
        
End Sub

