VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmTablaDesarrollo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla de Desarrollo"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   13575
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabCuponera 
      Height          =   6765
      Left            =   0
      TabIndex        =   2
      Top             =   30
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   11933
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
      TabCaption(0)   =   "Cuponera"
      TabPicture(0)   =   "frmTablaDesarrollo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCuponera"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos"
      TabPicture(1)   =   "frmTablaDesarrollo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDatos"
      Tab(1).Control(1)=   "cmdAceptar"
      Tab(1).Control(2)=   "cmdCancelar"
      Tab(1).ControlCount=   3
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
         Left            =   -66240
         TabIndex        =   7
         Top             =   5280
         Width           =   1200
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
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
         Left            =   -67680
         TabIndex        =   6
         Top             =   5280
         Width           =   1200
      End
      Begin VB.Frame fraDatos 
         Caption         =   "Actualización Valores Cuponera"
         Height          =   4335
         Left            =   -74760
         TabIndex        =   5
         Top             =   480
         Width           =   9975
         Begin VB.TextBox txtAmortizacion 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7440
            TabIndex        =   19
            Top             =   1640
            Width           =   1460
         End
         Begin VB.TextBox txtTasa 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7440
            TabIndex        =   18
            Top             =   960
            Width           =   1460
         End
         Begin VB.TextBox txtNumDias 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   15
            Top             =   1640
            Width           =   1460
         End
         Begin MSComCtl2.DTPicker dtpFechaPago 
            Height          =   285
            Left            =   2280
            TabIndex        =   17
            Top             =   3000
            Width           =   1460
            _ExtentX        =   2566
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
            Format          =   188350465
            CurrentDate     =   38768
         End
         Begin VB.Label lblFechaCorte 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2280
            TabIndex        =   16
            Top             =   2320
            Width           =   1460
         End
         Begin VB.Label lblNumCupon 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2280
            TabIndex        =   14
            Top             =   960
            Width           =   1460
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Amortización (%)"
            Height          =   195
            Index           =   5
            Left            =   5280
            TabIndex        =   13
            Top             =   1660
            Width           =   1155
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tasa (%)"
            Height          =   195
            Index           =   4
            Left            =   5280
            TabIndex        =   12
            Top             =   980
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Pago"
            Height          =   195
            Index           =   3
            Left            =   720
            TabIndex        =   11
            Top             =   3020
            Width           =   375
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Corte"
            Height          =   195
            Index           =   2
            Left            =   720
            TabIndex        =   10
            Top             =   2340
            Width           =   375
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Días"
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   9
            Top             =   1660
            Width           =   345
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Num. Cupón"
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   8
            Top             =   980
            Width           =   885
         End
      End
      Begin VB.Frame fraCuponera 
         Caption         =   "Tabla de Desarrollo"
         Height          =   5805
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   12675
         Begin TrueOleDBGrid60.TDBGrid tdgCuponera 
            Bindings        =   "frmTablaDesarrollo.frx":0038
            Height          =   5085
            Left            =   360
            OleObjectBlob   =   "frmTablaDesarrollo.frx":0052
            TabIndex        =   4
            Top             =   360
            Width           =   11955
         End
         Begin MSAdodcLib.Adodc adoCuponera 
            Height          =   330
            Left            =   7050
            Top             =   6000
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
      End
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
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
      Left            =   1200
      Picture         =   "frmTablaDesarrollo.frx":954A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6960
      Width           =   1200
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
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
      Left            =   11160
      Picture         =   "frmTablaDesarrollo.frx":961A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   1200
   End
End
Attribute VB_Name = "frmTablaDesarrollo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strCodTitulop        As String
Public strEstadop           As String
Dim strSQL                  As String

Private Sub Habilita()

    txtNumDias.Enabled = False
    dtpFechaPago.Enabled = False
    txtTasa.Enabled = False
    txtAmortizacion.Enabled = False
    
    Select Case tdgCuponera.Col
        Case 2
            txtNumDias.Enabled = True
        Case 4
            dtpFechaPago.Enabled = True
        Case 5
            txtTasa.Enabled = True
        Case 10
            txtAmortizacion.Enabled = True
    End Select
    
End Sub

Private Sub LlenarFormulario()

    lblNumCupon.Caption = tdgCuponera.Columns(0).Value
    txtNumDias.Text = CStr(tdgCuponera.Columns(2).Value)
    lblFechaCorte.Caption = tdgCuponera.Columns(3).Value
    dtpFechaPago.Value = tdgCuponera.Columns(4).Value
    txtTasa.Text = CStr(CDbl(tdgCuponera.Columns(5).Value))
    txtAmortizacion.Text = CStr(CDbl(tdgCuponera.Columns(10).Value))
        
End Sub

Private Sub cmdAceptar_Click()

    tdgCuponera.AllowUpdate = True
    If txtNumDias.Enabled Then
        tdgCuponera.Columns(2).Value = CLng(txtNumDias.Text)
        tdgCuponera.Update
        tdgCuponera.Columns(3).Value = CVDate(lblFechaCorte.Caption)
        tdgCuponera.Update
        tdgCuponera.Columns(4).Value = dtpFechaPago.Value
        tdgCuponera.Update
        
        '*** Actualizar fechas de cupón siguiente ***
        If tdgCuponera.Row < (adoCuponera.Recordset.RecordCount - 1) Then
            tdgCuponera.Row = tdgCuponera.Row + 1
            tdgCuponera.Columns(1).Value = DateAdd("d", 1, CVDate(lblFechaCorte.Caption))
            tdgCuponera.Update
            If tdgCuponera.Row = (adoCuponera.Recordset.RecordCount - 1) Then
                tdgCuponera.Columns(2).Value = CLng(DateDiff("d", tdgCuponera.Columns(1).Value, tdgCuponera.Columns(3).Value)) + 1
            Else
                tdgCuponera.Columns(3).Value = DateAdd("d", CLng(tdgCuponera.Columns(2).Value) - 1, tdgCuponera.Columns(1).Value)
            End If
            tdgCuponera.Update
        End If
    ElseIf dtpFechaPago.Enabled Then
        tdgCuponera.Columns(4).Value = dtpFechaPago.Value
        tdgCuponera.Update
    ElseIf txtTasa.Enabled Then
        tdgCuponera.Columns(5).Value = CDbl(txtTasa.Text)
        tdgCuponera.Update
    ElseIf txtAmortizacion.Enabled Then
        tdgCuponera.Columns(10).Value = CDbl(txtAmortizacion.Text)
        tdgCuponera.Update
    End If
    tdgCuponera.AllowUpdate = False
    
    If txtTasa.Enabled Then
        Call frmTitulosInversion.GenerarFactoresMontos(lblNumCupon.Caption, CDbl(txtTasa.Text))
    Else
        Call frmTitulosInversion.GenerarFactoresMontos("000", 0)
    End If
    
    cmdModificar.Visible = True
    With tabCuponera
        .TabEnabled(0) = True
        .Tab = 0
    End With
    Call BuscarCronograma
    
End Sub

Private Sub cmdCancelar_Click()

    cmdModificar.Visible = True
    With tabCuponera
        .TabEnabled(0) = True
        .Tab = 0
    End With
    Call BuscarCronograma
    
End Sub

Private Sub cmdModificar_Click()

    If strEstadop = Reg_Adicion Then
        Call LlenarFormulario
        cmdModificar.Visible = False
        With tabCuponera
            .TabEnabled(0) = False
            .Tab = 1
        End With
        Call Habilita
    Else
        If Trim(tdgCuponera.Columns(15).Value) = Valor_Indicador Or Trim(tdgCuponera.Columns(14).Value) = Valor_Caracter Then
            Call LlenarFormulario
            cmdModificar.Visible = False
            With tabCuponera
                .TabEnabled(0) = False
                .Tab = 1
            End With
            Call Habilita
        End If
    End If
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
        
    Call BuscarCronograma
    CentrarForm Me
    
End Sub

Private Sub BuscarCronograma()
    
    strSQL = "SELECT * FROM InstrumentoInversionCalendarioTmp " & _
        "WHERE CodTitulo='" & strCodTitulop & "' ORDER BY NumCupon"
    
    With adoCuponera
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With

    tdgCuponera.Refresh
    tabCuponera.Tab = 0
    
    If frmTitulosInversion.chkAmortiza.Value Then
        tdgCuponera.Columns(10).Visible = True
    Else
        tdgCuponera.Columns(10).Visible = False
    End If
    
    If frmTitulosInversion.chkAjuste.Value Then
        tdgCuponera.Columns(18).Visible = True
        tdgCuponera.Columns(19).Visible = True
    Else
        tdgCuponera.Columns(18).Visible = False
        tdgCuponera.Columns(19).Visible = False
    End If
            
End Sub

Private Sub tabCuponera_Click(PreviousTab As Integer)

    Select Case tabCuponera.Tab
        Case 1
            If PreviousTab = 0 Then
                If strEstadop = Reg_Adicion Then
                    Call cmdModificar_Click
                Else
                    If Trim(tdgCuponera.Columns(15).Value) = Valor_Indicador Or Trim(tdgCuponera.Columns(14).Value) = Valor_Caracter Then
                        Call cmdModificar_Click
                    Else
                        tabCuponera.Tab = 0
                    End If
                End If
            End If
    End Select
    
End Sub

Private Sub tdgCuponera_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 5 Then
        Call DarFormatoValor(Value, Decimales_TasaDiaria)
    End If
    
    If ColIndex = 6 Then
        Call DarFormatoValor(Value, Decimales_TasaDiaria)
    End If
    
    If ColIndex = 7 Then
        Call DarFormatoValor(Value, Decimales_TasaDiaria)
    End If
    
    If ColIndex = 10 Then
        Call DarFormatoValor(Value, Decimales_TasaDiaria)
    End If
    
End Sub

Private Sub txtAmortizacion_Change()

    Call FormatoCajaTexto(txtAmortizacion, Decimales_TasaDiaria)
    
End Sub

Private Sub txtAmortizacion_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtAmortizacion, Decimales_TasaDiaria)
    
End Sub

Private Sub txtNumDias_Change()

    If IsNumeric(txtNumDias.Text) Then
        lblFechaCorte.Caption = CStr(DateAdd("d", CLng(txtNumDias.Text) - 1, tdgCuponera.Columns(1).Value))
        dtpFechaPago.Value = CVDate(lblFechaCorte.Caption)
    End If
    
End Sub

Private Sub txtNumDias_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N")
    
End Sub

Private Sub txtTasa_Change()

    Call FormatoCajaTexto(txtTasa, Decimales_TasaDiaria)
    
End Sub

Private Sub txtTasa_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTasa, Decimales_TasaDiaria)
    
End Sub

