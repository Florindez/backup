VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBackupRestore 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BackUp / Restaurar"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   18375
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
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
      Left            =   16290
      Picture         =   "frmBackupRestore.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8640
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   18165
      _ExtentX        =   32041
      _ExtentY        =   14843
      _Version        =   393216
      Style           =   1
      Tabs            =   1
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
      TabCaption(0)   =   "BackUp"
      TabPicture(0)   =   "frmBackupRestore.frx":0582
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dgvBackups"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdRestore"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmBackup"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame frmBackup 
         Height          =   1935
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   14655
         Begin VB.TextBox txtNombreArchivo 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   12
            Top             =   360
            Width           =   9255
         End
         Begin VB.CheckBox chkEditarNombre 
            Caption         =   "Archivo"
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
            Left            =   240
            MaskColor       =   &H00800000&
            TabIndex        =   11
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdExplorar 
            Caption         =   "..."
            Height          =   375
            Left            =   10920
            TabIndex        =   10
            Top             =   360
            Width           =   375
         End
         Begin VB.OptionButton optPrecierre 
            Caption         =   "Pre Cierre"
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
            Left            =   1680
            TabIndex        =   9
            Top             =   1440
            Width           =   1335
         End
         Begin VB.OptionButton optAvance 
            Caption         =   "Avance"
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
            Left            =   240
            TabIndex        =   8
            Top             =   1440
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   285
            Left            =   1560
            TabIndex        =   7
            Top             =   1080
            Width           =   12855
         End
         Begin VB.CheckBox chkRuta 
            Caption         =   "Ruta"
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
            Height          =   255
            Left            =   240
            MaskColor       =   &H00800000&
            TabIndex        =   6
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtRuta 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   5
            Top             =   720
            Width           =   9255
         End
         Begin VB.CommandButton cmdBackup 
            Caption         =   "Ejecutar Backup"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   12720
            TabIndex        =   4
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblDescripcion 
            Caption         =   "Descripcion"
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
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "Restaurar"
         Enabled         =   0   'False
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
         Left            =   16710
         Picture         =   "frmBackupRestore.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   7560
         Width           =   1200
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dgvBackups 
         Height          =   4815
         Left            =   120
         OleObjectBlob   =   "frmBackupRestore.frx":0AE7
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2640
         Width           =   17805
      End
   End
End
Attribute VB_Name = "frmBackupRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim adoRegistroAux      As ADODB.Recordset
Dim adoRegistroBackup   As ADODB.Recordset
Dim numSecuencial       As Integer
Dim dir                 As String

Private Sub chkEditarNombre_Click()
    If chkEditarNombre.Value = vbChecked Then
        txtNombreArchivo.Enabled = True
    Else
        txtNombreArchivo.Enabled = False
    End If
End Sub

Private Sub chkRuta_Click()
    If chkRuta.Value = vbChecked Then
        txtRuta.Enabled = True
    Else
        txtRuta.Enabled = False
    End If
End Sub

Private Sub cmdBackup_Click()
    Dim strSQL  As String
    Dim correlativo As Integer
    
    'Realización del backup
    strSQL = "BACKUP DATABASE [" & gstrDataBase & "] TO  DISK = N'" & dir & txtNombreArchivo.Text & ".bak" & "' WITH NOFORMAT, INIT,  NAME = N'fondos-Completa Base de datos Copia de seguridad', SKIP, NOREWIND, NOUNLOAD,  STATS = 10"
    adoConn.Execute strSQL
    
    'obtencion del correlativo
    adoComm.CommandText = "Select max(CodBackup) as Correlativo from Audit.dbo.TablaBackup"
    Set adoRegistroAux = adoComm.Execute
    If adoRegistroAux("Correlativo") > 0 Then
        correlativo = adoRegistroAux("Correlativo") + 1
    Else
        correlativo = 1
    End If
    
    'Registro del backup
    strSQL = "Insert into Audit.dbo.TablaBackup values (" & correlativo & ",'" & txtRuta.Text & "','" & txtNombreArchivo.Text & ".bak" & "','" & txtDescripcion.Text & _
            "','" & gstrFechaActual & "',CONVERT(CHAR(19),'" & DateTime.Now & "',113),'"
    
    If optAvance.Value = True Then
        strSQL = strSQL & "AVANCE'"
    Else
        strSQL = strSQL & "PRECIERRE'"
    End If
    
    strSQL = strSQL & "," & numSecuencial & ")"
    
    adoConn.Execute strSQL
    
    MsgBox "El BackUp de la base de datos se realizó correctamente" & vbCrLf & "El nombre de archivo es " & txtNombreArchivo.Text, vbInformation, Me.Caption
    
    Call CargarGrilla
    txtDescripcion.Text = ""
    Call GeneraNombrePorDefecto
    
End Sub

Private Sub cmdRestore_Click()
    
    Dim strMensaje As String
    Dim strSQL  As String
    Dim BackUp As String
    
    strMensaje = "Para restaurar el Backup seleccionado se procederá a interrumpir la conexión " & _
                    "de todos los Usuarios a la Base de Datos del Sistema, incluyendo la conexión actual. ¿Está seguro de continuar?"
    
    If MsgBox(strMensaje, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
        GoTo cmdRestore_fin
    End If
    
    Set adoConn = New ADODB.Connection
    
    If adoConn.State = 1 Then
        adoConn.Close:  Set adoConn = Nothing
    End If
    
    gstrConnectConsulta = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & gstrLogin & ";Password=" & gstrPassword & ";" & _
                          "Data Source=" & gstrServer & ";" & _
                          "Initial Catalog=master;"
    
    '*** SQLOLEDB - Base de Datos ***
    gstrConnect = "User ID=" & gstrLogin & ";Password=" & gstrPassword & ";" & _
                        "Data Source=" & gstrServer & ";" & _
                        "Initial Catalog=master;" & _
                        "Application Name=" & App.Title & ";" & _
                        "Auto Translate=False"
    
    frmMainMdi.stbMdi.Panels(3).Text = "Desconectándose de la Base de datos..."
    
    With adoConn
        .Provider = "SQLOLEDB"
        .ConnectionString = gstrConnect
        .CommandTimeout = 0
        .ConnectionTimeout = 0
        .Open
    End With
    
    With adoComm
        .CommandTimeout = 0

        .CommandText = "Use master exec KillConexiones " & gstrDataBase
        .Execute
    End With
    
    BackUp = dgvBackups.Columns.Item(1).Value
    
    frmMainMdi.stbMdi.Panels(3).Text = "Procesando Restauración de la Base de datos..."
    
    'Realización del Restore
    With adoComm
        .CommandTimeout = 0
        .CommandText = "RESTORE DATABASE [" & gstrDataBase & "] FROM  DISK = N'" & dgvBackups.Columns.ColumnByFieldName("RutaBackup").Value & dgvBackups.Columns.ColumnByFieldName("NombreArchivo").Value & "' WITH FILE = 1, NOUNLOAD,  REPLACE,  STATS = 10"
        adoComm.Execute
    End With
    
    
    MsgBox "Restauración de la Base de Datos completada. Por favor, reinicie Spectrum Fondos para que hagan efecto los cambios.", vbInformation

    frmMainMdi.stbMdi.Panels(3).Text = "BackUp Restaurado..."
       
cmdRestore_fin:
   Me.MousePointer = vbDefault
   frmMainMdi.stbMdi.Panels(3).Text = "Acción..."
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CargarGrilla()
    
    strSQL = "SELECT  CodBackup, RutaBackup, NombreArchivo, DescripBackup, FechaSistema, FechaBackup, TipoBackup from Audit.dbo.TablaBackup where CodBackup > 0 ORDER BY CodBackup desc"
    
    With dgvBackups
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = gstrConnectConsulta
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.Active = False
            .Dataset.ADODataset.CommandText = strSQL
            .Dataset.DisableControls
            .Dataset.Active = True
            .KeyField = "CodBackup"
    End With

End Sub


Private Sub Form_Load()
    
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    ConfGrid dgvBackups, True, False, False, False
    Call CargarGrilla
    'Obtencion de valores por defecto
    txtRuta.Text = gstrBackupPath
    Call GeneraNombrePorDefecto
    
End Sub

Private Sub optBackup_Click()
    If optBackup.Value = True Then
        optRestore.Value = False
        optAvance.Enabled = True
        optPrecierre.Enabled = True
        frmRestaurar.Enabled = False
        frmBackup.Enabled = True
    End If
    Call GeneraNombrePorDefecto

End Sub

Private Sub optAvance_Click()
   If optAvance.Value = True Then
        optPrecierre.Value = False
    End If
    Call GeneraNombrePorDefecto

End Sub

Private Sub optPrecierre_Click()
     If optPrecierre.Value = True Then
        optAvance.Value = False
    End If
    Call GeneraNombrePorDefecto

End Sub

Private Sub optRestore_Click()
    If optRestore.Value = True Then
        optBackup.Value = False
        optAvance.Enabled = False
        optPrecierre.Enabled = False
        frmRestaurar.Enabled = True
        frmBackup.Enabled = False
    End If
    Call GeneraNombrePorDefecto
    
End Sub

Private Sub GeneraNombrePorDefecto()
    Dim prefix      As String
    Dim database    As String
    Dim backupdate  As String
    '++REA 2015-06-02
    Dim prefix3 As String
    Dim prefix2 As String
    Dim RegCount As Integer
    Dim RegPosition As Double
    '--REA 2015-06-02
    
    If optAvance.Value = True Then
        txtDescripcion.Text = "Avance "
    Else
        txtDescripcion.Text = "Pre-Cierre "
    End If

    'prefix = Mid(gstrBackupPath, (InStr(gstrBackupPath, "_") + 1))
    'prefix = Mid(prefix, 1, Len(prefix) - 1)
    'prefix = "BK"
    prefix2 = "PYME"
    prefix3 = "_"
    prefix = "Backup" '_PYME1_20141126_PYME2_20141110_2
    'database = gstrDataBase
    'backupdate = gstrFechaActual
    backupdate = ""
    
    'dir = "F:" + Mid(gstrBackupPath, 12, Len(gstrBackupPath))
    dir = gstrBackupPath 'Mid(gstrBackupPath, 12, Len(gstrBackupPath))
    
    '++REA 2015-06-02
    adoComm.CommandText = "SELECT CodFondo, FechaCuota FROM fondos.dbo.FondoValorCuota where IndAbierto = 'X' order by CodFondo"
    Set adoRegistroAux = adoComm.Execute
    Do While Not adoRegistroAux.EOF
        backupdate = backupdate & prefix2 & CInt(adoRegistroAux("CodFondo")) & prefix3 & Convertyyyymmdd(adoRegistroAux("FechaCuota")) & prefix3
        adoRegistroAux.MoveNext
    Loop
    
    backupdate = Mid(backupdate, 1, Len(backupdate) - 1)
    '--REA 2015-06-02
    
    txtNombreArchivo.Text = prefix & "_" & backupdate
    
    Call ObtenerSecuencial(prefix & "_" & backupdate)
    
    'If numSecuencial > 1 Then
        txtNombreArchivo.Text = txtNombreArchivo.Text & "_" & numSecuencial
    'End If
    
End Sub


Private Sub ObtenerSecuencial(ByVal FileName As String)
    adoComm.CommandText = "SELECT MAX(NumSecuencial) as Secuencial from Audit.dbo.TablaBackup Where NombreArchivo Like '" & txtNombreArchivo.Text & "%'"
    Set adoRegistroAux = adoComm.Execute
    If adoRegistroAux("Secuencial") > 0 Then
        numSecuencial = adoRegistroAux("Secuencial") + 1
    Else
        numSecuencial = 1
    End If
End Sub

Private Sub txtRuta_Change()
    dir = txtRuta.Text
End Sub
