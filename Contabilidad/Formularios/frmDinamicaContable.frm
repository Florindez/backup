VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmDinamicaContable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dinámica Contable"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   13965
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   480
      TabIndex        =   28
      Top             =   8340
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   1296
      Buttons         =   3
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
      UserControlWidth=   4200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   12030
      TabIndex        =   27
      Top             =   8340
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      Visible0        =   0   'False
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TabDlg.SSTab tabDinamica 
      Height          =   8295
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   14631
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
      TabCaption(0)   =   "Dinámica"
      TabPicture(0)   =   "frmDinamicaContable.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTipoCambio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmDinamicaContable.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDetalle"
      Tab(1).Control(1)=   "cmdAccion"
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -64500
         TabIndex        =   26
         Top             =   7440
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
         Height          =   1035
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   12885
         Begin VB.ComboBox cboTipoOperacionBus 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   450
            Width           =   5985
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Operacion"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   29
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame fraDetalle 
         Caption         =   "Definición Dinámica"
         Height          =   6795
         Left            =   -74640
         TabIndex        =   4
         Top             =   630
         Width           =   13155
         Begin VB.CommandButton cmdAccionDinamica 
            Caption         =   "..."
            Height          =   315
            Index           =   4
            Left            =   6540
            TabIndex        =   49
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   3240
            Width           =   315
         End
         Begin VB.CommandButton cmdAccionDinamica 
            Caption         =   "..."
            Height          =   315
            Index           =   8
            Left            =   12450
            TabIndex        =   47
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   3960
            Width           =   315
         End
         Begin VB.CommandButton cmdAccionDinamica 
            Caption         =   "..."
            Height          =   315
            Index           =   7
            Left            =   6540
            TabIndex        =   46
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   3960
            Width           =   315
         End
         Begin VB.CheckBox chkIndicadorCtaBancaria 
            Caption         =   "Es Cuenta Bancaria?"
            Height          =   195
            Left            =   360
            TabIndex        =   41
            Top             =   2220
            Width           =   2535
         End
         Begin VB.CommandButton cmdAccionDinamica 
            Caption         =   "..."
            Height          =   315
            Index           =   3
            Left            =   12450
            TabIndex        =   40
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   2880
            Width           =   315
         End
         Begin VB.CommandButton cmdAccionDinamica 
            Caption         =   "..."
            Height          =   315
            Index           =   2
            Left            =   6540
            TabIndex        =   38
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   2880
            Width           =   315
         End
         Begin VB.ComboBox cboVistaDinamica 
            Height          =   315
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   480
            Width           =   6285
         End
         Begin VB.CommandButton cmdAgregar 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   575
            Left            =   480
            Picture         =   "frmDinamicaContable.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Agregar detalle"
            Top             =   5460
            Width           =   495
         End
         Begin VB.CommandButton cmdQuitar 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   550
            Left            =   480
            Picture         =   "frmDinamicaContable.frx":02E5
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Quitar detalle"
            Top             =   6060
            Width           =   495
         End
         Begin VB.CommandButton cmdActualizar 
            Height          =   575
            Left            =   480
            Picture         =   "frmDinamicaContable.frx":0537
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Actualizar Detalle"
            Top             =   4860
            Width           =   495
         End
         Begin VB.CommandButton cmdAtras 
            Height          =   575
            Left            =   -30
            Picture         =   "frmDinamicaContable.frx":07F2
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Actualizar Detalle"
            Top             =   4860
            Width           =   495
         End
         Begin VB.TextBox txtDescripcionDinamica 
            Height          =   315
            Left            =   2460
            MaxLength       =   150
            TabIndex        =   18
            Top             =   840
            Width           =   9945
         End
         Begin VB.CommandButton cmdAccionDinamica 
            Caption         =   "..."
            Height          =   315
            Index           =   6
            Left            =   12450
            TabIndex        =   17
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   3600
            Width           =   315
         End
         Begin VB.CommandButton cmdAccionDinamica 
            Caption         =   "..."
            Height          =   315
            Index           =   0
            Left            =   12450
            TabIndex        =   15
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   1590
            Width           =   315
         End
         Begin VB.CommandButton cmdAccionDinamica 
            Caption         =   "..."
            Height          =   315
            Index           =   9
            Left            =   12450
            TabIndex        =   14
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   4320
            Width           =   315
         End
         Begin VB.CommandButton cmdAccionDinamica 
            Caption         =   "..."
            Height          =   315
            Index           =   5
            Left            =   12450
            TabIndex        =   13
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   3240
            Width           =   315
         End
         Begin VB.CommandButton cmdAccionDinamica 
            Caption         =   "..."
            Height          =   315
            Index           =   1
            Left            =   12450
            TabIndex        =   11
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   2490
            Width           =   315
         End
         Begin TrueOleDBGrid60.TDBGrid tdgDinamica 
            Height          =   1665
            Left            =   1200
            OleObjectBlob   =   "frmDinamicaContable.frx":0C77
            TabIndex        =   9
            Top             =   4860
            Width           =   11565
         End
         Begin VB.ComboBox cboTipoOperacion 
            Height          =   315
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1200
            Width           =   6285
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda"
            Height          =   195
            Index           =   13
            Left            =   360
            TabIndex        =   50
            Top             =   3300
            Width           =   585
         End
         Begin VB.Label lblCodMoneda 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   315
            Left            =   2460
            TabIndex        =   48
            Top             =   3240
            Width           =   4095
         End
         Begin VB.Label lblCodContraparte 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   315
            Left            =   8340
            TabIndex        =   45
            Top             =   3960
            Width           =   4095
         End
         Begin VB.Label lblTipoContraparte 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   315
            Left            =   2460
            TabIndex        =   44
            Top             =   3960
            Width           =   4095
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cod.Contraparte"
            Height          =   195
            Index           =   12
            Left            =   7080
            TabIndex        =   43
            Top             =   3990
            Width           =   1155
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Contraparte"
            Height          =   195
            Index           =   11
            Left            =   360
            TabIndex        =   42
            Top             =   4050
            Width           =   1185
         End
         Begin VB.Label lblCodAnalitica 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   315
            Left            =   8340
            TabIndex        =   39
            Top             =   2880
            Width           =   4095
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cod.Analitica"
            Height          =   195
            Index           =   10
            Left            =   7080
            TabIndex        =   37
            Top             =   2910
            Width           =   930
         End
         Begin VB.Label lblCodFile 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   315
            Left            =   2460
            TabIndex        =   36
            Top             =   2880
            Width           =   4095
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cod.File"
            Height          =   195
            Index           =   9
            Left            =   360
            TabIndex        =   35
            Top             =   2940
            Width           =   570
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción Movimiento"
            Height          =   195
            Index           =   8
            Left            =   360
            TabIndex        =   34
            Top             =   4380
            Width           =   1695
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Movimiento"
            Height          =   195
            Index           =   7
            Left            =   360
            TabIndex        =   33
            Top             =   3660
            Width           =   1545
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cargo/Abono"
            Height          =   195
            Index           =   6
            Left            =   7080
            TabIndex        =   32
            Top             =   3270
            Width           =   960
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta"
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   31
            Top             =   2580
            Width           =   510
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Condicion Dinamica"
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   30
            Top             =   1680
            Width           =   1650
         End
         Begin VB.Label lblDescripDetalle 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   315
            Left            =   2460
            TabIndex        =   25
            Top             =   4320
            Width           =   9975
         End
         Begin VB.Label lblCondicionDinamica 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   315
            Left            =   2460
            TabIndex        =   19
            Top             =   1590
            Width           =   9975
         End
         Begin VB.Label lblMontoMovimiento 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   315
            Left            =   2460
            TabIndex        =   16
            Top             =   3600
            Width           =   9975
         End
         Begin VB.Line Line1 
            X1              =   300
            X2              =   12780
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Label lblCargoAbono 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   315
            Left            =   8340
            TabIndex        =   12
            Top             =   3240
            Width           =   4095
         End
         Begin VB.Label lblCuenta 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   315
            Left            =   2460
            TabIndex        =   10
            Top             =   2520
            Width           =   9975
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Operación"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   8
            Top             =   1305
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   7
            Top             =   930
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vista Dinámica"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   6
            Top             =   555
            Width           =   1050
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmDinamicaContable.frx":5EEB
         Height          =   5895
         Left            =   360
         OleObjectBlob   =   "frmDinamicaContable.frx":5F05
         TabIndex        =   1
         Top             =   1800
         Width           =   12885
      End
   End
End
Attribute VB_Name = "frmDinamicaContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim arrTipoOperacionBus() As String
Dim arrTipoOperacion()   As String, arrVistaDinamica()          As String
Dim strCodVistaDinamica         As String
Dim strCodTipoOperacion  As String, strCodTipoOperacionBus As String
Dim strSQL               As String
Dim strEstado                   As String
Dim adoConsulta                 As ADODB.Recordset
Dim adoRegistroAux          As ADODB.Recordset
Dim strDinamicaContableDetalleXML As String

Private Sub cboTipoOperacion_Click()
    strCodTipoOperacion = Valor_Caracter
    If cboTipoOperacion.ListIndex < 0 Then Exit Sub
    strCodTipoOperacion = Trim(arrTipoOperacion(cboTipoOperacion.ListIndex))
End Sub

Private Sub cboTipoOperacionBus_Click()
    strCodTipoOperacionBus = Valor_Caracter
    If cboTipoOperacionBus.ListIndex < 0 Then Exit Sub
    strCodTipoOperacionBus = Trim(arrTipoOperacionBus(cboTipoOperacionBus.ListIndex))
    Call Buscar
End Sub

Private Sub cboVistaDinamica_Click()
    
    strCodVistaDinamica = Valor_Caracter
    If cboVistaDinamica.ListIndex < 0 Then Exit Sub
    
    strCodVistaDinamica = Trim(arrVistaDinamica(cboVistaDinamica.ListIndex))
    
End Sub



Private Sub cmdActualizar_Click()
    If strEstado = Reg_Consulta Then Exit Sub
    
    If strEstado = Reg_Adicion Or strEstado = Reg_Edicion Then
        
        adoRegistroAux.Fields("CodCuenta") = lblCuenta.Caption
        adoRegistroAux.Fields("CodFile") = lblCodFile.Caption
        adoRegistroAux.Fields("CodAnalitica") = lblCodAnalitica.Caption
        adoRegistroAux.Fields("CodMoneda") = lblCodMoneda.Caption
        adoRegistroAux.Fields("IndDebeHaber") = lblCargoAbono.Caption
        adoRegistroAux.Fields("MontoMovimiento") = lblMontoMovimiento.Caption
        adoRegistroAux.Fields("TipoContraparte") = lblTipoContraparte.Caption
        adoRegistroAux.Fields("CodContraparte") = lblCodContraparte.Caption
        adoRegistroAux.Fields("DescripMovimiento") = lblDescripDetalle.Caption
        
    End If
    
End Sub


Private Sub cmdAccionDinamica_Click(Index As Integer)
    
    
    If cboVistaDinamica.ListCount > 0 Then
        If cboVistaDinamica.ListIndex = 0 Then
            MsgBox "Debe seleccionar la Vista Dinamica", vbCritical
            cboVistaDinamica.SetFocus
            Exit Sub
        End If
    End If
    
    gstrTextoAdministradorFormula = Valor_Caracter
    
    CargarAdministradorFormulas
    
    Select Case Index
    
        Case 0
        
            If Trim(lblCondicionDinamica.Caption) <> Valor_Caracter Then
                gstrTextoAdministradorFormula = Trim(lblCondicionDinamica.Caption)
                frmAdministradorFormulas.AdministradorFormulas1.TextoFormula = Trim(lblCondicionDinamica.Caption)
            End If
            
            frmAdministradorFormulas.Show vbModal
            
            lblCondicionDinamica.Caption = gstrTextoAdministradorFormula
    
        Case 1
    
            
            If Trim(lblCuenta.Caption) <> Valor_Caracter Then
                gstrTextoAdministradorFormula = Trim(lblCuenta.Caption)
                frmAdministradorFormulas.AdministradorFormulas1.TextoFormula = Trim(lblCuenta.Caption)
            End If
            
            frmAdministradorFormulas.Show vbModal
            
            lblCuenta.Caption = gstrTextoAdministradorFormula
    
        Case 2  'CodFile
        
            If Trim(lblCodFile.Caption) <> Valor_Caracter Then
                gstrTextoAdministradorFormula = Trim(lblCodFile.Caption)
                frmAdministradorFormulas.AdministradorFormulas1.TextoFormula = Trim(lblCodFile.Caption)
            End If
            
            frmAdministradorFormulas.Show vbModal
            
            lblCodFile.Caption = gstrTextoAdministradorFormula
        
        Case 3  'CodAnalitica
        
            If Trim(lblCodAnalitica.Caption) <> Valor_Caracter Then
                gstrTextoAdministradorFormula = Trim(lblCodAnalitica.Caption)
                frmAdministradorFormulas.AdministradorFormulas1.TextoFormula = Trim(lblCodAnalitica.Caption)
            End If
            
            frmAdministradorFormulas.Show vbModal
            
            lblCodAnalitica.Caption = gstrTextoAdministradorFormula
        
        Case 4  'CodMoneda
        
            If Trim(lblCodMoneda.Caption) <> Valor_Caracter Then
                gstrTextoAdministradorFormula = Trim(lblCodMoneda.Caption)
                frmAdministradorFormulas.AdministradorFormulas1.TextoFormula = Trim(lblCodMoneda.Caption)
            End If
            
            frmAdministradorFormulas.Show vbModal
            
            lblCodMoneda.Caption = gstrTextoAdministradorFormula
        
        
        Case 5  'Cargo/Abono
            
            If Trim(lblCargoAbono.Caption) <> Valor_Caracter Then
                gstrTextoAdministradorFormula = Trim(lblCargoAbono.Caption)
                frmAdministradorFormulas.AdministradorFormulas1.TextoFormula = Trim(lblCargoAbono.Caption)
            End If
            
            frmAdministradorFormulas.Show vbModal
            
            lblCargoAbono.Caption = gstrTextoAdministradorFormula
        
        Case 6  'MontoMovimiento
        
            If Trim(lblMontoMovimiento.Caption) <> Valor_Caracter Then
                gstrTextoAdministradorFormula = Trim(lblMontoMovimiento.Caption)
                frmAdministradorFormulas.AdministradorFormulas1.TextoFormula = Trim(lblMontoMovimiento.Caption)
            End If
            
            frmAdministradorFormulas.Show vbModal
            
            lblMontoMovimiento.Caption = gstrTextoAdministradorFormula
        
        Case 7  'TipoContraparte
        
            If Trim(lblTipoContraparte.Caption) <> Valor_Caracter Then
                gstrTextoAdministradorFormula = Trim(lblTipoContraparte.Caption)
                frmAdministradorFormulas.AdministradorFormulas1.TextoFormula = Trim(lblTipoContraparte.Caption)
            End If
            
            frmAdministradorFormulas.Show vbModal
            
            lblTipoContraparte.Caption = gstrTextoAdministradorFormula
        
        
        Case 8  'CodContraparte
            
            If Trim(lblCodContraparte.Caption) <> Valor_Caracter Then
                gstrTextoAdministradorFormula = Trim(lblCodContraparte.Caption)
                frmAdministradorFormulas.AdministradorFormulas1.TextoFormula = Trim(lblCodContraparte.Caption)
            End If
            
            frmAdministradorFormulas.Show vbModal
            
            lblCodContraparte.Caption = gstrTextoAdministradorFormula
        
        Case 9  'Descripcion
    
            If Trim(lblDescripDetalle.Caption) <> Valor_Caracter Then
                gstrTextoAdministradorFormula = Trim(lblDescripDetalle.Caption)
                frmAdministradorFormulas.AdministradorFormulas1.TextoFormula = Trim(lblDescripDetalle.Caption)
            End If
            
            frmAdministradorFormulas.Show vbModal
            
            lblDescripDetalle.Caption = gstrTextoAdministradorFormula
    
    End Select


End Sub







Private Sub cmdAgregar_Click()

    Dim adoRegistro As ADODB.Recordset
    Dim intSecuencial As Integer
    Dim dblBookmark As Double
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    If strEstado = Reg_Adicion Or strEstado = Reg_Edicion Then
        If TodoOkDetalleDinamica() Then
           
            adoRegistroAux.AddNew
            adoRegistroAux.Fields("CodCuenta") = Trim(lblCuenta.Caption)
            adoRegistroAux.Fields("CodFile") = Trim(lblCodFile.Caption)
            adoRegistroAux.Fields("CodAnalitica") = Trim(lblCodAnalitica.Caption)
            adoRegistroAux.Fields("IndDebeHaber") = Trim(lblCargoAbono.Caption)
            adoRegistroAux.Fields("CodMoneda") = Trim(lblCodMoneda.Caption)
            adoRegistroAux.Fields("MontoMovimiento") = Trim(lblMontoMovimiento.Caption)
            adoRegistroAux.Fields("TipoContraparte") = Trim(lblTipoContraparte.Caption)
            adoRegistroAux.Fields("CodContraparte") = Trim(lblCodContraparte.Caption)
            adoRegistroAux.Fields("DescripMovimiento") = Trim(lblDescripDetalle.Caption)
            
            adoRegistroAux.Update
            
            dblBookmark = adoRegistroAux.Bookmark
            
            tdgDinamica.Refresh
            
            adoRegistroAux.Bookmark = dblBookmark
            
            cmdQuitar.Enabled = True
            
            Call LimpiarDatos
        
        End If
    End If
End Sub

Private Sub LimpiarDatos()
    
    lblCuenta.Caption = Valor_Caracter
    lblCargoAbono.Caption = Valor_Caracter
    lblMontoMovimiento.Caption = Valor_Caracter
    lblDescripDetalle.Caption = Valor_Caracter
    
End Sub

Private Function TodoOkDetalleDinamica()
    
    TodoOkDetalleDinamica = False
    
'    If Trim(lblCuenta.Caption) = Valor_Caracter Then
'        MsgBox "Cuenta no ingresada", vbCritical, gstrNombreEmpresa
'        If cmdAdmiCuenta.Enabled Then cmdAdmiCuenta.SetFocus
'        Exit Function
'    End If
                  
'    If Trim(lblCargoAbono.Caption) = Valor_Caracter Then
'        MsgBox "Cargo - Abono no ingresada", vbCritical, gstrNombreEmpresa
'        If cmdAdmiCargoAbono.Enabled Then cmdAdmiCargoAbono.SetFocus
'        Exit Function
'    End If
    
'    If Trim(lblMontoMovimiento.Caption) = Valor_Caracter Then
'        MsgBox "Monto Movimiento no ingresada", vbCritical, gstrNombreEmpresa
'        If cmdAdmiMontoMovimiento.Enabled Then cmdAdmiMontoMovimiento.SetFocus
'        Exit Function
'    End If
    
'    If Trim(lblDescripDetalle.Caption) = Valor_Caracter Then
'        MsgBox "Descripcion no ingresada", vbCritical, gstrNombreEmpresa
'        If cmdAdmiDescripcion.Enabled Then cmdAdmiDescripcion.SetFocus
'        Exit Function
'    End If
        
    '*** Si todo pasó OK ***
    TodoOkDetalleDinamica = True
    
End Function

Private Sub cmdAtras_Click()
    cmdAtras.Enabled = False
    
    cmdActualizar.Enabled = False
    
    cmdAgregar.Enabled = True
    
    cmdQuitar.Enabled = True
End Sub

Private Sub cmdQuitar_Click()
    Dim dblBookmark As Double

    If adoRegistroAux.RecordCount > 0 Then
    
        dblBookmark = adoRegistroAux.Bookmark
    
        adoRegistroAux.Delete adAffectCurrent
        
        If adoRegistroAux.EOF Then
            adoRegistroAux.MovePrevious
            tdgDinamica.MovePrevious
        End If
            
        adoRegistroAux.Update
        
        If adoRegistroAux.RecordCount = 0 Then cmdQuitar.Enabled = False

        If adoRegistroAux.RecordCount > 0 And Not adoRegistroAux.BOF And Not adoRegistroAux.EOF And dblBookmark > 1 Then adoRegistroAux.Bookmark = dblBookmark - 1
        
        If adoRegistroAux.RecordCount > 0 And Not adoRegistroAux.BOF And Not adoRegistroAux.EOF Then adoRegistroAux.Bookmark = dblBookmark - 1
        
        If adoRegistroAux.RecordCount > 0 And Not adoRegistroAux.BOF And Not adoRegistroAux.EOF Then adoRegistroAux.Bookmark = dblBookmark - 1
   
        tdgDinamica.Refresh
    
    End If
End Sub





Private Sub Form_Activate()

    'Call CargarReportes
    
End Sub

Private Sub Form_Deactivate()

    'Call OcultarReportes
    
End Sub


Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call CargarReportes
    Call Buscar
    Call DarFormato
    
    CentrarForm Me
    
End Sub


Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
'    lblDescrip(9).ForeColor = &H800000
'    lblDescrip(9).Font = "Arial"
'    lblDescrip(9).FontBold = True
'    lblDescrip(10).ForeColor = &H800000
'    lblDescrip(10).Font = "Arial"
'    lblDescrip(10).FontBold = True
            
End Sub

Public Sub CargarAdministradorFormulas()
    Me.MousePointer = vbHourglass
    frmAdministradorFormulas.AdministradorFormulas1.CargarVariables gstrConnectNET, "up_ACLstVariablesVistaProceso", strCodVistaDinamica, Tipo_Campo_Output, Valor_Caracter
    frmAdministradorFormulas.AdministradorFormulas1.CargarFunciones gstrConnectNET, "up_CNLstVistaUsuarioFuncion", strCodVistaDinamica
    frmAdministradorFormulas.AdministradorFormulas1.CargarOperadoresConCadena "+|-"
    Me.MousePointer = vbDefault
End Sub

Public Sub Buscar()
    
    Dim strSQL As String

    strSQL = "SELECT CodAdministradora,CodVistaProceso,DescripDinamica,TipoOperacion,DescripParametro DescripTipoOperacion,CondicionDinamica " & _
            "FROM DinamicaContable2 DC JOIN AuxiliarParametro AP ON DC.TipoOperacion=AP.CodParametro AND CodTipoParametro='OPECAJ' " & _
            "WHERE CodAdministradora='" & gstrCodAdministradora & "' "
    
    If strCodTipoOperacionBus <> Valor_Caracter Then
        strSQL = strSQL & " AND TipoOperacion='" & strCodTipoOperacionBus & "'"
    End If
    
    Set adoConsulta = New ADODB.Recordset

    strEstado = Reg_Defecto

    With adoConsulta
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
Private Sub CargarReportes()
'
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Vista Activa"
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    
End Sub

Private Sub CargarListas()

    Dim strSQL As String, intRegistro As Integer
        
    '*** Tipo de Operación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='OPECAJ' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoOperacion, arrTipoOperacion(), Sel_Defecto
    
    If cboTipoOperacion.ListCount > 0 Then cboTipoOperacion.ListIndex = 0

    CargarControlLista strSQL, cboTipoOperacionBus, arrTipoOperacionBus(), Sel_Todos
    
    If cboTipoOperacionBus.ListCount > 0 Then cboTipoOperacionBus.ListIndex = 0
    
    '*** Vistas Dinamica ***
    strSQL = "SELECT CodVistaProceso CODIGO,DescripVistaProceso DESCRIP FROM VistaProceso " & _
                "WHERE IndVigente='X'"
    CargarControlLista strSQL, cboVistaDinamica, arrVistaDinamica(), Sel_Defecto
    
    If cboVistaDinamica.ListCount > 0 Then cboVistaDinamica.ListIndex = 0
    
            
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vModify
            Call Modificar
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
    
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Dinámica Contable..."
                
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabDinamica
        .TabEnabled(0) = False
        .TabEnabled(1) = True
        .Tab = 1
    End With
    
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabDinamica
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .Tab = 0
    End With
    Call Buscar
    
End Sub

Public Sub Eliminar()

End Sub

Public Sub Grabar()

    Dim strMsgError                     As String
    Dim objDinamicaContableDetalleXML   As DOMDocument60
    Dim strDinamicaContableDetalleXML   As String
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    If strEstado = Reg_Adicion Or strEstado = Reg_Edicion Then
        If TodoOK() Then
            
            If MsgBox("¿Desea guardar los cambios realizados?", vbQuestion + vbYesNo + vbDefaultButton2, gstrNombreEmpresa) = vbNo Then Exit Sub
            
            Me.MousePointer = vbHourglass
                                                                                                                         
            With adoComm
                
                On Error GoTo Ctrl_Error
                
                Call XMLADORecordset(objDinamicaContableDetalleXML, "DinamicaContable", "Detalle", adoRegistroAux, strMsgError)
                strDinamicaContableDetalleXML = objDinamicaContableDetalleXML.xml
                
                '*** Cabecera ***
                .CommandText = "{ call up_ACManDinamicaContableXML('" & _
                    gstrCodAdministradora & "','" & strCodVistaDinamica & "','" & Trim(txtDescripcionDinamica.Text) & "','" & _
                    strCodTipoOperacion & "','" & _
                    Trim(lblCondicionDinamica.Caption) & "','" & strDinamicaContableDetalleXML & "','" & _
                    IIf(strEstado = Reg_Adicion, "I", "U") & "') }"
                adoConn.Execute .CommandText
                
                                                                                
            End With
            
            Set adoRegistroAux = Nothing
                                                                                                                         
            Me.MousePointer = vbDefault
                        
            If strEstado = Reg_Adicion Then
                MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            End If
            
            If strEstado = Reg_Edicion Then
                MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            End If
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabDinamica
                .TabEnabled(0) = True
                .Tab = 0
            End With
            
            Call Buscar
            
        End If
    End If
    
    Exit Sub
    
Ctrl_Error:
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
    
End Sub

Private Function TodoOK() As Boolean
            
    TodoOK = False
    
    If cboVistaDinamica.ListIndex = 0 Then
        MsgBox "No ah seleccionado la Vista Dinamica", vbCritical
        If cboVistaDinamica.Enabled Then cboVistaDinamica.SetFocus
        Exit Function
    End If
        
    If Trim(txtDescripcionDinamica.Text) = Valor_Caracter Then
        MsgBox "La descripcion esta vacia", vbCritical
        If txtDescripcionDinamica.Enabled Then txtDescripcionDinamica.SetFocus
        Exit Function
    End If
    
    If cboTipoOperacion.ListIndex = 0 Then
        MsgBox "No ah seleccionado la Vista Dinamica", vbCritical
        If cboTipoOperacion.Enabled Then cboTipoOperacion.SetFocus
        Exit Function
    End If
    
    If Trim(lblCondicionDinamica.Caption) = Valor_Caracter Then
        MsgBox "La condicion de la dinamica esta vacia", vbCritical
        'If cmdAdmiCondicion.Enabled Then cmdAdmiCondicion.SetFocus
        Exit Function
    End If

    If adoRegistroAux.RecordCount = 0 Then
        MsgBox "No existen registros en el detalle", vbCritical
        Exit Function
    End If
    
'    '*** Si todo paso OK ***
    TodoOK = True
  
End Function
Public Sub Imprimir()
    
    'Call SubImprimir(1)
    
End Sub

Public Sub Modificar()
    
    If tdgConsulta.SelBookmarks.Count = 0 Then Exit Sub
    
    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabDinamica
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .Tab = 1
        End With
        
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)
    
    Dim intRegistro As Integer
    
    Select Case strModo
        Case Reg_Adicion
            
            If cboVistaDinamica.ListCount > 0 Then cboVistaDinamica.ListIndex = 0
            txtDescripcionDinamica.Text = Valor_Caracter
            If cboTipoOperacion.ListCount > 0 Then cboTipoOperacion.ListIndex = 0
            
            Call CargarDetalleGrilla
            Call LimpiarDatos
            
            
            'cmdAdmiCondicion.Enabled = True
            cboTipoOperacion.Enabled = True
            cboVistaDinamica.Enabled = True
            
            cmdAtras.Enabled = False
            cmdActualizar.Enabled = False
            cmdQuitar.Enabled = False
        
        Case Reg_Edicion
            
            intRegistro = ObtenerItemLista(arrVistaDinamica(), Trim(tdgConsulta.Columns(1)))
            If intRegistro >= 0 Then cboVistaDinamica.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrTipoOperacion(), Trim(tdgConsulta.Columns(3)))
            If intRegistro >= 0 Then cboTipoOperacion.ListIndex = intRegistro
            
            txtDescripcionDinamica.Text = tdgConsulta.Columns(2)
            
            lblCondicionDinamica.Caption = tdgConsulta.Columns(5)
            
            'cmdAdmiCondicion.Enabled = False
            cboTipoOperacion.Enabled = False
            cboVistaDinamica.Enabled = False
            
            Call CargarDetalleGrilla
            
            cmdAtras.Enabled = False
            cmdActualizar.Enabled = False
            cmdQuitar.Enabled = False
    End Select
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Private Sub InicializarValores()

    strEstado = Reg_Defecto
    tabDinamica.Tab = 0
    
    ConfiguraRecordsetAuxiliar
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
    tabDinamica.TabEnabled(0) = True
    tabDinamica.TabEnabled(1) = False
'
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    Set frmDinamicaContable = Nothing
    
End Sub

Private Sub tabDinamica_Click(PreviousTab As Integer)

'    Select Case tabDinamica.Tab
'        Case 1
'            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vQuery)
'            If strEstado = Reg_Defecto Then tabDinamica.Tab = 0
'
'    End Select
    
End Sub

Private Sub ConfiguraRecordsetAuxiliar()

    Set adoRegistroAux = New ADODB.Recordset

    With adoRegistroAux
       .CursorLocation = adUseClient
       .Fields.Append "CodCuenta", adVarChar, 8000
       .Fields.Append "CodFile", adVarChar, 8000
       .Fields.Append "CodAnalitica", adVarChar, 8000
       .Fields.Append "IndDebeHaber", adVarChar, 8000
       .Fields.Append "CodMoneda", adVarChar, 8000
       .Fields.Append "MontoMovimiento", adVarChar, 8000
       .Fields.Append "TipoContraparte", adVarChar, 8000
       .Fields.Append "CodContraparte", adVarChar, 8000
       .Fields.Append "DescripMovimiento", adVarChar, 8000
'       .CursorType = adOpenStatic
       .LockType = adLockBatchOptimistic
    End With
    
    adoRegistroAux.Open

End Sub

Private Sub CargarDetalleGrilla()
    
    Dim adoRegistro As ADODB.Recordset
    Dim adoField As ADODB.Field
    
    Dim strSQL As String

    Call ConfiguraRecordsetAuxiliar
    
    If strEstado = Reg_Edicion Then
    
        Set adoRegistro = New ADODB.Recordset
    
        strSQL = "{ call up_ACListarDinamicaContableDetalle('" & gstrCodAdministradora & "','" & Trim(tdgConsulta.Columns("TipoOperacion")) & "') }"
           
        With adoRegistro
            .ActiveConnection = gstrConnectConsulta
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open strSQL
        
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    adoRegistroAux.AddNew
                    For Each adoField In adoRegistroAux.Fields
                        adoRegistroAux.Fields(adoField.Name) = adoRegistro.Fields(adoField.Name)
                    Next
                    adoRegistroAux.Update
                    adoRegistro.MoveNext
                    'adoMovimiento.MoveNext
                Loop
                adoRegistroAux.MoveFirst
            End If
        End With
        
    End If
    
    tdgDinamica.DataSource = adoRegistroAux
            
End Sub

Private Sub tdgDinamica_DblClick()
    
    lblCuenta.Caption = adoRegistroAux.Fields("CodCuenta")
    lblCodFile.Caption = adoRegistroAux.Fields("CodFile")
    lblCodAnalitica.Caption = adoRegistroAux.Fields("CodAnalitica")
    lblCargoAbono.Caption = adoRegistroAux.Fields("IndDebeHaber")
    
    lblCodMoneda.Caption = adoRegistroAux.Fields("CodMoneda")
    lblMontoMovimiento.Caption = adoRegistroAux.Fields("MontoMovimiento")
    
    lblTipoContraparte.Caption = adoRegistroAux.Fields("TipoContraparte")
    lblCodContraparte.Caption = adoRegistroAux.Fields("CodContraparte")
    
    lblDescripDetalle.Caption = adoRegistroAux.Fields("DescripMovimiento")
    

    cmdAgregar.Enabled = False
    cmdQuitar.Enabled = False
    cmdAtras.Enabled = True
    cmdActualizar.Enabled = True
    
End Sub
