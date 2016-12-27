VERSION 5.00
Object = "{BE4F3AC8-AEC9-101A-947B-00DD010F7B46}#1.0#0"; "MSOUTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{442CAE95-1D41-47B1-BE83-6995DA3CE254}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmPlanCuenta 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan de Cuentas"
   ClientHeight    =   10275
   ClientLeft      =   1875
   ClientTop       =   1455
   ClientWidth     =   12525
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
   ScaleHeight     =   10275
   ScaleWidth      =   12525
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   10440
      TabIndex        =   2
      Top             =   9480
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
      Left            =   720
      TabIndex        =   1
      Top             =   9480
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
   Begin TabDlg.SSTab tabPlanCuenta 
      Height          =   9315
      Left            =   30
      TabIndex        =   15
      Top             =   60
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   16431
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Cuentas"
      TabPicture(0)   =   "frmPlanCuenta.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDescrip(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cboGrupoContable"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraCuentas"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Datos"
      TabPicture(1)   =   "frmPlanCuenta.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(1)=   "fraDatos"
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -66120
         TabIndex        =   13
         Top             =   8400
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
      Begin VB.Frame fraDatos 
         Height          =   7725
         Left            =   -74640
         TabIndex        =   18
         Top             =   630
         Width           =   11685
         Begin VB.ComboBox cboCuentaTraslacionPerdida 
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
            Left            =   1935
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   6480
            Width           =   8085
         End
         Begin VB.ComboBox cboCuentaTraslacionGanancia 
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
            Left            =   1935
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   6870
            Width           =   8100
         End
         Begin VB.CheckBox chkPartidaMonetaria 
            Caption         =   "Partida Monetaria"
            ForeColor       =   &H00800000&
            Height          =   345
            Left            =   360
            TabIndex        =   36
            Top             =   4320
            Width           =   2595
         End
         Begin VB.ComboBox cboTipoFile 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1935
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   3390
            Width           =   3615
         End
         Begin VB.ComboBox cboTipoAuxiliar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   3900
            Width           =   3165
         End
         Begin VB.CheckBox chkAuxiliar 
            Caption         =   "Indicador Auxiliar"
            ForeColor       =   &H00800000&
            Height          =   345
            Left            =   360
            TabIndex        =   31
            Top             =   3870
            Width           =   2595
         End
         Begin VB.ComboBox cboRubroEEFF 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1935
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1125
            Width           =   5295
         End
         Begin VB.ComboBox cboTipoEEFF 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1935
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   750
            Width           =   5295
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
            Height          =   315
            Left            =   1935
            TabIndex        =   6
            Top             =   1515
            Width           =   2775
         End
         Begin VB.TextBox txtDescripCuenta 
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
            Left            =   1935
            MaxLength       =   1000
            TabIndex        =   7
            Top             =   1875
            Width           =   9300
         End
         Begin VB.ComboBox cboMovimiento 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1935
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   2250
            Width           =   3615
         End
         Begin VB.ComboBox cboMoneda 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1935
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   2610
            Width           =   3615
         End
         Begin VB.ComboBox cboNaturaleza 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1935
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   3000
            Width           =   3615
         End
         Begin VB.ComboBox cboCuentaGanancia 
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
            Left            =   1935
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   5640
            Width           =   8100
         End
         Begin VB.ComboBox cboCuentaPerdida 
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
            Left            =   1935
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   5280
            Width           =   8085
         End
         Begin VB.ComboBox cboTipoCuenta 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1935
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   360
            Width           =   5295
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Grupo Actual"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   17
            Left            =   5550
            TabIndex        =   43
            Top             =   1530
            Width           =   1185
         End
         Begin VB.Label lblGrupoActual 
            Alignment       =   1  'Right Justify
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
            Left            =   6750
            TabIndex        =   42
            Top             =   1500
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Ajuste por Traslación :"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   16
            Left            =   360
            TabIndex        =   41
            Top             =   6090
            Width           =   2175
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Pérdida"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   15
            Left            =   360
            TabIndex        =   40
            Top             =   6510
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Ganancia"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   14
            Left            =   360
            TabIndex        =   39
            Top             =   6870
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo File"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   13
            Left            =   360
            TabIndex        =   35
            Top             =   3420
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Auxiliar"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   12
            Left            =   3990
            TabIndex        =   33
            Top             =   3930
            Width           =   1635
         End
         Begin VB.Label lblNivelCuenta 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1"
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
            Left            =   4680
            TabIndex        =   17
            Top             =   1515
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo de EEFF"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   360
            TabIndex        =   19
            Top             =   765
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Rubro de EEFF"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   360
            TabIndex        =   29
            Top             =   1155
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Código"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   3
            Left            =   360
            TabIndex        =   28
            Top             =   1530
            Width           =   735
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Descripción"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   4
            Left            =   360
            TabIndex        =   27
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Movimiento"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   5
            Left            =   360
            TabIndex        =   26
            Top             =   2310
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Moneda"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   6
            Left            =   360
            TabIndex        =   25
            Top             =   2655
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Naturaleza"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   7
            Left            =   360
            TabIndex        =   24
            Top             =   3045
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Ganancia"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   10
            Left            =   360
            TabIndex        =   23
            Top             =   5640
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Pérdida"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   9
            Left            =   360
            TabIndex        =   22
            Top             =   5280
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Ajuste por Tipo de Cambio :"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   8
            Left            =   360
            TabIndex        =   21
            Top             =   4890
            Width           =   2595
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo de Cuenta"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   11
            Left            =   360
            TabIndex        =   20
            Top             =   375
            Width           =   1455
         End
      End
      Begin VB.Frame fraCuentas 
         Caption         =   "Cuentas Contables"
         Height          =   7455
         Left            =   360
         TabIndex        =   14
         Top             =   1170
         Width           =   11625
         Begin MSOutl.Outline otlCuenta 
            Height          =   7050
            Left            =   120
            TabIndex        =   30
            Top             =   270
            Width           =   11415
            _Version        =   65536
            _ExtentX        =   20135
            _ExtentY        =   12435
            _StockProps     =   77
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Style           =   5
            PicturePlus     =   "frmPlanCuenta.frx":0038
            PictureMinus    =   "frmPlanCuenta.frx":0196
            PictureLeaf     =   "frmPlanCuenta.frx":02F4
            PictureOpen     =   "frmPlanCuenta.frx":03EE
            PictureClosed   =   "frmPlanCuenta.frx":054C
         End
      End
      Begin VB.ComboBox cboGrupoContable 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   630
         Width           =   4695
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Grupo"
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   16
         Top             =   675
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmPlanCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrGrupoContable()              As String, arrTipoEEFF()                    As String
Dim arrRubroEEFF()                  As String, arrMovimiento()                  As String
Dim arrMoneda()                     As String, arrNaturaleza()                  As String
Dim arrCuentaPerdida()              As String, arrCuentaGanancia()              As String
Dim arrCuentaTraslacionPerdida()    As String, arrCuentaTraslacionGanancia()    As String
Dim arrTipoCuenta()                 As String, arrCuentas()                     As String
Dim arrTipoAuxiliar()               As String, arrTipoFile()                    As String

Dim strCodGrupoContable             As String, strCodTipoEEFF                   As String
Dim strCodRubroEEFF                 As String, strCodMovimiento                 As String
Dim strCodMoneda                    As String, strCodNaturaleza                 As String
Dim strCodCuentaPerdida             As String, strCodCuentaGanancia             As String
Dim strCodCuentaTraslacionPerdida   As String, strCodCuentaTraslacionGanancia   As String
Dim strCodTipoCuenta                As String, strCodGrupoContableSel           As String
Dim strEstado                       As String, strTipoAuxiliar                  As String
Dim strIndicadorAuxiliar            As String, strTipoFile                      As String
Dim NumVersionPlanContable          As Integer, strIndicadorPartidaMonetaria    As String

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

    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar cuentas..."
    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabPlanCuenta
        .TabEnabled(0) = False
        .TabEnabled(1) = True
        .Tab = 1
    End With
    Call Habilita
    
End Sub

Private Sub LlenarFormulario(strModo As String)
     Dim strSQL  As String

    NumVersionPlanContable = ObtenerPlanContableVersion()
        
    Select Case strModo
    
        Case Reg_Adicion
        
            cboTipoCuenta.ListIndex = -1
            If cboTipoCuenta.ListCount > 0 Then cboTipoCuenta.ListIndex = 0
        
            cboTipoEEFF.ListIndex = -1
            If cboTipoEEFF.ListCount > 0 Then cboTipoEEFF.ListIndex = 0
            
            cboRubroEEFF.ListIndex = -1
            If cboRubroEEFF.ListCount > 0 Then cboRubroEEFF.ListIndex = 0
            
            txtCodCuenta.Text = ""
            txtDescripCuenta.Text = ""
            
            cboMovimiento.ListIndex = -1
            If cboMovimiento.ListCount > 0 Then cboMovimiento.ListIndex = 0
            
            cboMoneda.ListIndex = -1
            If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = 0
            
            cboNaturaleza.ListIndex = -1
            If cboNaturaleza.ListCount > 0 Then cboNaturaleza.ListIndex = 0
            
            cboCuentaPerdida.ListIndex = -1
            If cboCuentaPerdida.ListCount > 0 Then cboCuentaPerdida.ListIndex = 0
            
            cboCuentaGanancia.ListIndex = 0
            If cboCuentaGanancia.ListCount > 0 Then cboCuentaGanancia.ListIndex = 0
                        
            cboCuentaTraslacionPerdida.ListIndex = 0
            If cboCuentaTraslacionPerdida.ListCount > 0 Then cboCuentaTraslacionPerdida.ListIndex = 0
            
            cboCuentaTraslacionGanancia.ListIndex = -1
            If cboCuentaTraslacionGanancia.ListCount > 0 Then cboCuentaTraslacionGanancia.ListIndex = 0
            
            chkAuxiliar.Value = vbUnchecked
            
            Call chkAuxiliar_Click
            
            chkPartidaMonetaria.Value = vbUnchecked
            
            Call chkPartidaMonetaria_Click
            
            cboTipoFile.ListIndex = -1
            'If cboTipoFile.ListCount > 0 Then cboTipoFile.ListIndex = 0
                        
            cboTipoCuenta.SetFocus
                        
        Case Reg_Edicion
            Dim adoRegistro As ADODB.Recordset
            Dim intRegistro As Integer
            
            Set adoRegistro = New ADODB.Recordset
            
            If otlCuenta.ListIndex <= 0 Then Exit Sub

            adoComm.CommandText = "SELECT CodCuenta,CodGrupoCuenta," & _
                "CodDivisionaria,NivelCuenta,CodTipoEEFF,CodRubroEEFF,CodMoneda,NaturalezaCuenta,IndMovimiento,IndMonetaria," & _
                "CuentaAmarre,CuentaAutomatica,CuentaDestino,CuentaPerdidaCambio,CuentaGananciaCambio," & _
                "CuentaPerdidaTraslacion,CuentaGananciaTraslacion,DescripCuenta," & _
                "TipoFile,IndAuxiliar,TipoAuxiliar " & _
                "FROM PlanContable WHERE CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodCuenta='" & arrCuentas(otlCuenta.ListIndex - 1) & "' AND NumVersion=" & NumVersionPlanContable
                        
            Set adoRegistro = adoComm.Execute

            If Not adoRegistro.EOF Then
                cboTipoCuenta.ListIndex = -1
                If Len(Trim(adoRegistro("CodCuenta"))) = 2 Then
                    intRegistro = ObtenerItemLista(arrTipoCuenta(), Codigo_Tipo_Cuenta_Grupo)
                Else
                    intRegistro = ObtenerItemLista(arrTipoCuenta(), Codigo_Tipo_Cuenta_Cuenta)
                End If
                If intRegistro >= 0 Then cboTipoCuenta.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrTipoEEFF(), adoRegistro("CodTipoEEFF"))
                If intRegistro >= 0 Then cboTipoEEFF.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrRubroEEFF(), adoRegistro("CodRubroEEFF"))
                If intRegistro >= 0 Then cboRubroEEFF.ListIndex = intRegistro
                                
                txtCodCuenta.Text = Trim(adoRegistro("CodCuenta"))
                txtDescripCuenta.Text = Trim(adoRegistro("DescripCuenta"))
                
                lblGrupoActual.Caption = adoRegistro("CodGrupoCuenta")

                If cboMovimiento.Enabled Then
                    If Trim(adoRegistro("IndMovimiento")) = Valor_Caracter Then
                        intRegistro = ObtenerItemLista(arrMovimiento(), Codigo_Respuesta_No)
                    Else
                        intRegistro = ObtenerItemLista(arrMovimiento(), Codigo_Respuesta_Si)
                    End If
                    If intRegistro >= 0 Then cboMovimiento.ListIndex = intRegistro
                    
                    If Trim(adoRegistro("CodMoneda")) = "00" Then
                        cboMoneda.ListIndex = 0
                    Else
                        intRegistro = ObtenerItemLista(arrMoneda(), adoRegistro("CodMoneda"))
                        If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
                    End If
                    
                    If Trim(adoRegistro("NaturalezaCuenta")) = "T" Then
                        cboNaturaleza.ListIndex = 0
                    Else
                        intRegistro = ObtenerItemLista(arrNaturaleza(), adoRegistro("NaturalezaCuenta"))
                        If intRegistro >= 0 Then cboNaturaleza.ListIndex = intRegistro
                    End If
                    
                    intRegistro = ObtenerItemLista(arrCuentaPerdida(), adoRegistro("CuentaPerdidaCambio"))
                    If intRegistro >= 0 Then cboCuentaPerdida.ListIndex = intRegistro

                    intRegistro = ObtenerItemLista(arrCuentaGanancia(), adoRegistro("CuentaGananciaCambio"))
                    If intRegistro >= 0 Then cboCuentaGanancia.ListIndex = intRegistro
                    
                    intRegistro = ObtenerItemLista(arrCuentaTraslacionPerdida(), adoRegistro("CuentaPerdidaTraslacion"))
                    If intRegistro >= 0 Then cboCuentaTraslacionPerdida.ListIndex = intRegistro
                    
                    intRegistro = ObtenerItemLista(arrCuentaTraslacionGanancia(), adoRegistro("CuentaGananciaTraslacion"))
                    If intRegistro >= 0 Then cboCuentaTraslacionGanancia.ListIndex = intRegistro
                    
                    If adoRegistro("IndAuxiliar") = Valor_Indicador Then
                        chkAuxiliar.Value = vbChecked
                    Else
                        chkAuxiliar.Value = vbUnchecked
                    End If
                    
                    Call chkAuxiliar_Click
                   
                    If adoRegistro("IndMonetaria") = Valor_Indicador Then
                        chkPartidaMonetaria.Value = vbChecked
                    Else
                        chkPartidaMonetaria.Value = vbUnchecked
                    End If
                    
                    Call chkPartidaMonetaria_Click
                   
                    If Trim(adoRegistro("TipoAuxiliar")) = "00" Then 'TODOS
                        cboTipoAuxiliar.ListIndex = 0
                    Else
                        intRegistro = ObtenerItemLista(arrTipoAuxiliar(), adoRegistro("TipoAuxiliar"))
                        If intRegistro >= 0 Then cboTipoAuxiliar.ListIndex = intRegistro
                    End If
                    
                    
                    intRegistro = ObtenerItemLista(arrTipoFile(), adoRegistro("TipoFile"))
                    If intRegistro >= 0 Then cboTipoFile.ListIndex = intRegistro
                   
                End If
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
    
    End Select
    
End Sub
Public Sub Buscar()

    Dim adoRegistro As ADODB.Recordset
    Dim intContador As Integer, intRegistros As Integer

    Set adoRegistro = New ADODB.Recordset
    adoRegistro.CursorLocation = adUseClient
    adoRegistro.CursorType = adOpenStatic
    
    strEstado = Reg_Defecto
    intContador = 1
    With adoComm
        
        .CommandText = "{ call up_CNObtenerPlanContable ('" & _
                            gstrCodAdministradora & "',"
            
        If cboGrupoContable.ListIndex > 0 Then
            .CommandText = .CommandText & "'" & strCodGrupoContable & "' ) }"
        Else
            .CommandText = .CommandText & "'%') }"
        End If
        
        adoRegistro.Open .CommandText, adoConn
      
        otlCuenta.Clear
        otlCuenta.List(0) = "[Plan de Cuentas Contables]"
        intRegistros = adoRegistro.RecordCount
        ReDim arrCuentas(intRegistros)
        adoRegistro.MoveFirst
        
        Do While Not adoRegistro.EOF
            arrCuentas(intContador - 1) = adoRegistro("CodCuenta")
            
            otlCuenta.AddItem adoRegistro("CodCuenta") + " " + adoRegistro("DescripCuenta")
            otlCuenta.indent(intContador) = adoRegistro("NivelCuenta")
                                      
            otlCuenta.ItemData(intContador) = adoRegistro("CodGrupoCuenta")
                                        
            intContador = intContador + 1
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
    End With

    For intContador = 1 To intRegistros
        If otlCuenta.HasSubItems(intContador) = False Then
            otlCuenta.PictureType(intContador) = 2
        End If
    Next
    
    If intRegistros > 0 Then strEstado = Reg_Consulta

End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabPlanCuenta
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .Tab = 0
    End With
    Call Buscar
    
End Sub

Public Sub Eliminar()

    If otlCuenta.ListIndex < 0 Then Exit Sub
    
    If strEstado = Reg_Consulta Then
        Dim adoConsulta     As ADODB.Recordset
        Dim intNivelCuenta
        
        If MsgBox(Mensaje_Eliminacion, vbQuestion + vbYesNo + vbDefaultButton2, gstrNombreEmpresa) = vbNo Then Exit Sub
        
        intNivelCuenta = Len(Trim(arrCuentas(otlCuenta.ListIndex - 1))) - 1
        Set adoConsulta = New ADODB.Recordset
        With adoComm
            '*** Verificar si existen cuentas dependientes ***
            .CommandText = "SELECT COUNT(CodCuenta) NumRegistros FROM PlanContable WHERE CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodCuenta LIKE '" & Trim(arrCuentas(otlCuenta.ListIndex - 1)) & "%' AND NivelCuenta >" & intNivelCuenta & " AND NumVersion=" & NumVersionPlanContable
                
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                If CInt(adoConsulta("NumRegistros")) > 0 Then
                    MsgBox "Cuenta no se puede eliminar. Existen cuentas dependientes.", vbCritical
                    
                    adoConsulta.Close: Set adoConsulta = Nothing
                    Exit Sub
                End If
            End If
            adoConsulta.Close
            
            If intNivelCuenta = 1 Then
                '*** Eliminar Grupo Contable ***
'                .CommandText = "DELETE PlanContableGrupo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                    "CodGrupoCuenta='" & Trim(arrCuentas(otlCuenta.ListIndex - 1)) & "'"
                .CommandText = "{ call up_CNManPlanContableGrupo('" & _
                gstrCodAdministradora & "'," & NumVersionPlanContable & ",'" & Trim(arrCuentas(otlCuenta.ListIndex - 1)) & "'," & _
                "'','','','D') }"
                        
                adoConn.Execute .CommandText
            End If
            
            '*** Eliminar Cuenta ***
'            .CommandText = "DELETE PlanContable WHERE CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                "CodCuenta='" & Trim(arrCuentas(otlCuenta.ListIndex - 1)) & "'"
            
            .CommandText = "{ call up_CNManPlanContable('" & _
                gstrCodAdministradora & "'," & NumVersionPlanContable & ",'" & Trim(arrCuentas(otlCuenta.ListIndex - 1)) & "'," & _
                "'',''," & _
                "0,'','','','','','','','','','','','','','','','','','','','D') }"
            
            adoConn.Execute .CommandText
                        
        End With
        
        MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation
        Call Buscar
    End If

End Sub

Public Sub Grabar()

    Dim intCantRegistros    As Integer, intRegistro        As Integer
    Dim strIndMovimiento    As String
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    strIndMovimiento = Valor_Caracter
    If strCodMovimiento = Codigo_Respuesta_Si Then strIndMovimiento = Valor_Indicador
    
    If strEstado = Reg_Adicion Then
        If TodoOK() Then
            Me.MousePointer = vbHourglass
                                                
            With adoComm
                If strCodTipoCuenta = Codigo_Tipo_Cuenta_Grupo Then
                
                    strCodGrupoContableSel = Trim(txtCodCuenta.Text)
                
                    '*** Grupo Contable ***
                    .CommandText = "{ call up_CNManPlanContableGrupo('" & _
                        gstrCodAdministradora & "'," & NumVersionPlanContable & ",'" & strCodGrupoContableSel & "','" & _
                        Trim(txtDescripCuenta.Text) & "','" & strCodTipoEEFF & "','" & _
                        strCodRubroEEFF & "','I') }"
                        
                    adoConn.Execute .CommandText
                    

                    .CommandText = "{ call up_CNManPlanContable('" & _
                        gstrCodAdministradora & "'," & NumVersionPlanContable & ",'" & strCodGrupoContableSel & "','" & _
                        strCodGrupoContableSel & "','" & Mid(strCodGrupoContableSel, 2, 1) & "'," & _
                        CInt(lblNivelCuenta.Caption) & ",'" & Trim(txtDescripCuenta.Text) & "','" & _
                        strTipoFile & "','" & strIndicadorAuxiliar & "','" & strTipoAuxiliar & "','" & _
                        strCodTipoEEFF & "','" & strCodRubroEEFF & "','" & strCodMoneda & "','" & _
                        strCodNaturaleza & "','" & strIndMovimiento & "','" & strIndicadorPartidaMonetaria & "','','','','" & _
                        strCodCuentaPerdida & "','" & strCodCuentaGanancia & "','" & strCodCuentaTraslacionPerdida & "','" & strCodCuentaTraslacionGanancia & "','','','I') }"

                Else
                    
                    .CommandText = "{ call up_CNManPlanContable('" & _
                        gstrCodAdministradora & "'," & NumVersionPlanContable & ",'" & Trim(txtCodCuenta.Text) & "','" & _
                        strCodGrupoContableSel & "','" & Mid(txtCodCuenta.Text, Len(strCodGrupoContableSel) + 1, 2) & "'," & _
                        CInt(lblNivelCuenta.Caption) & ",'" & Trim(txtDescripCuenta.Text) & "','" & _
                        strTipoFile & "','" & strIndicadorAuxiliar & "','" & strTipoAuxiliar & "','" & _
                        strCodTipoEEFF & "','" & strCodRubroEEFF & "','" & strCodMoneda & "','" & _
                        strCodNaturaleza & "','" & strIndMovimiento & "','" & strIndicadorPartidaMonetaria & "','','','','" & _
                        strCodCuentaPerdida & "','" & strCodCuentaGanancia & "','" & strCodCuentaTraslacionPerdida & "','" & strCodCuentaTraslacionGanancia & "','','','I') }"
                
                End If
                adoConn.Execute .CommandText
                
            End With
                
            Me.MousePointer = vbDefault
        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabPlanCuenta
                .TabEnabled(0) = True
                .Tab = 0
            End With
            
            Call Buscar
        End If
    End If
    
    If strEstado = Reg_Edicion Then
        If TodoOK() Then
            Me.MousePointer = vbHourglass
                                                
            With adoComm
                If strCodTipoCuenta = Codigo_Tipo_Cuenta_Grupo Then
                    '*** Grupo Contable ***
                    .CommandText = "{ call up_CNManPlanContableGrupo('" & _
                        gstrCodAdministradora & "'," & NumVersionPlanContable & ",'" & Trim(txtCodCuenta.Text) & "','" & _
                        Trim(txtDescripCuenta.Text) & "','" & strCodTipoEEFF & "','" & _
                        strCodRubroEEFF & "','U') }"
                        
                    adoConn.Execute .CommandText

                    .CommandText = "{ call up_CNManPlanContable('" & _
                        gstrCodAdministradora & "'," & NumVersionPlanContable & ",'" & Trim(txtCodCuenta.Text) & "','" & _
                        Trim(txtCodCuenta.Text) & "','" & Mid(txtCodCuenta.Text, 2, 1) & "'," & _
                        CInt(lblNivelCuenta.Caption) & ",'" & Trim(txtDescripCuenta.Text) & "','" & _
                        strTipoFile & "','" & strIndicadorAuxiliar & "','" & strTipoAuxiliar & "','" & _
                        strCodTipoEEFF & "','" & strCodRubroEEFF & "','" & strCodMoneda & "','" & _
                        strCodNaturaleza & "','" & strIndMovimiento & "','" & strIndicadorPartidaMonetaria & "','','','','" & _
                        strCodCuentaPerdida & "','" & strCodCuentaGanancia & "','" & strCodCuentaTraslacionPerdida & "','" & strCodCuentaTraslacionGanancia & "','','','U') }"
                
                
                Else
                
                    .CommandText = "{ call up_CNManPlanContable('" & _
                        gstrCodAdministradora & "'," & NumVersionPlanContable & ",'" & Trim(txtCodCuenta.Text) & "','" & _
                        strCodGrupoContableSel & "','" & Mid(txtCodCuenta.Text, Len(strCodGrupoContableSel) + 1, 2) & "'," & _
                        CInt(lblNivelCuenta.Caption) & ",'" & Trim(txtDescripCuenta.Text) & "','" & _
                        strTipoFile & "','" & strIndicadorAuxiliar & "','" & strTipoAuxiliar & "','" & _
                        strCodTipoEEFF & "','" & strCodRubroEEFF & "','" & strCodMoneda & "','" & _
                        strCodNaturaleza & "','" & strIndMovimiento & "','" & strIndicadorPartidaMonetaria & "','','','','" & _
                        strCodCuentaPerdida & "','" & strCodCuentaGanancia & "','" & strCodCuentaTraslacionPerdida & "','" & strCodCuentaTraslacionGanancia & "','','','U') }"
                
                
                End If
                adoConn.Execute .CommandText
                
            End With
                
            Me.MousePointer = vbDefault
        
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabPlanCuenta
                .TabEnabled(0) = True
                .Tab = 0
            End With
            
            Call Buscar
        End If
    End If
        
    Exit Sub
        
End Sub


Private Function TodoOK() As Boolean

    TodoOK = False
        
    If strCodTipoCuenta = Valor_Caracter Then
        MsgBox "Seleccione el Tipo de Cuenta", vbCritical, gstrNombreEmpresa
        cboTipoCuenta.SetFocus
        Exit Function
    End If
    
    If strCodTipoEEFF = Valor_Caracter Then
        MsgBox "Seleccione el Tipo de EEFF", vbCritical, gstrNombreEmpresa
        cboTipoEEFF.SetFocus
        Exit Function
    End If
      
    If strCodRubroEEFF = Valor_Caracter Then
        MsgBox "Seleccione el Rubro de EEFF", vbCritical, gstrNombreEmpresa
        cboRubroEEFF.SetFocus
        Exit Function
    End If
    
    If Trim(txtCodCuenta.Text) = Valor_Caracter Then
        MsgBox "Registre el código de cuenta", vbCritical, gstrNombreEmpresa
        txtCodCuenta.SetFocus
        Exit Function
    End If
    
    If Trim(txtDescripCuenta.Text) = Valor_Caracter Then
        MsgBox "Registre el la descripción del código de cuenta", vbCritical, gstrNombreEmpresa
        txtDescripCuenta.SetFocus
        Exit Function
    End If
    
    If strCodTipoCuenta = Codigo_Tipo_Cuenta_Cuenta Then
        If cboCuentaPerdida.ListIndex < 0 And Not (Trim(txtCodCuenta.Text) Like "[6-7]*") Then
            MsgBox "Seleccione la cuenta de pérdida por diferencia en cambio", vbCritical, gstrNombreEmpresa
            cboCuentaPerdida.SetFocus
            Exit Function
        End If
        
        If cboCuentaGanancia.ListIndex < 0 And Not (Trim(txtCodCuenta.Text) Like "[6-7]*") Then
            MsgBox "Seleccione la cuenta de ganancia por diferencia en cambio", vbCritical, gstrNombreEmpresa
            cboCuentaGanancia.SetFocus
            Exit Function
        End If
        
        'Valida que el grupo de la cuenta a ingresar sea el correcto
        If ObtenerGrupoCuentaContable(gstrCodAdministradora, Trim(txtCodCuenta.Text), NumVersionPlanContable) <> strCodGrupoContableSel Then
            MsgBox "El grupo de la cuenta a ingresar es incorrecto!", vbCritical, gstrNombreEmpresa
            cboCuentaGanancia.SetFocus
            Exit Function
        End If
        
    End If
    
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function
Public Sub Imprimir()

End Sub

Public Sub SubImprimir(Index As Integer)

    Dim strSeleccionRegistro    As String
    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()

    Select Case Index
        Case 1
            gstrNameRepo = "PlanContable"
            Set frmReporte = New frmVisorReporte
            
            ReDim aReportParamS(2)
            ReDim aReportParamFn(2)
            ReDim aReportParamF(2)
            
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
                
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
                
            If strCodGrupoContable <> Valor_Caracter Then
                '*** Grupo Contable seleccionado ***
                aReportParamS(0) = strCodGrupoContable
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Codigo_Listar_Individual
                'aReportParamS(3) = NumVersionPlanContable
            Else
                '*** Lista de comprobantes por rango de fecha ***
                aReportParamS(0) = strCodGrupoContable
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Codigo_Listar_Todos
                'aReportParamS(3) = NumVersionPlanContable
            End If
    End Select
        
    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub
Public Sub Modificar()
    
     Dim strSQL  As String
    
    If otlCuenta.ListIndex < 0 Then Exit Sub
    
    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        Call Deshabilita
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabPlanCuenta
            .TabEnabled(0) = False
            .Tab = 1
        End With
        
        
        cboCuentaPerdida.Enabled = True
        If cboCuentaPerdida.ListIndex = -1 Then

             '*** Cuenta Pérdida ***
            strSQL = "SELECT CodCuenta CODIGO,(CodCuenta + space(1) + DescripCuenta) DESCRIP FROM PlanContable WHERE IndMovimiento='X' AND CodCuenta LIKE '67611%' AND CodAdministradora='" & gstrCodAdministradora & "' AND NumVersion=" & NumVersionPlanContable
            CargarControlLista strSQL, cboCuentaPerdida, arrCuentaPerdida(), Sel_Defecto

            cboCuentaPerdida.ListIndex = 0

        End If


        cboCuentaGanancia.Enabled = True
        If cboCuentaGanancia.ListIndex = -1 Then

            '*** Cuenta Ganancia ***
            strSQL = "SELECT CodCuenta CODIGO,(CodCuenta + space(1) + DescripCuenta) DESCRIP FROM PlanContable WHERE IndMovimiento='X' AND CodCuenta LIKE '77611%' AND CodAdministradora='" & gstrCodAdministradora & "' AND NumVersion=" & NumVersionPlanContable
            CargarControlLista strSQL, cboCuentaGanancia, arrCuentaGanancia(), Sel_Defecto

            cboCuentaGanancia.ListIndex = 0

        End If

    End If
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub




Private Sub cboCuentaGanancia_Click()

    strCodCuentaGanancia = ""
    If cboCuentaGanancia.ListIndex < 0 Then Exit Sub
    
    strCodCuentaGanancia = Trim(arrCuentaGanancia(cboCuentaGanancia.ListIndex))
    
End Sub


Private Sub cboCuentaPerdida_Click()

    strCodCuentaPerdida = ""
    If cboCuentaPerdida.ListIndex < 0 Then Exit Sub
    
    strCodCuentaPerdida = Trim(arrCuentaPerdida(cboCuentaPerdida.ListIndex))
     
End Sub


Private Sub cboCuentaTraslacionGanancia_Click()

    strCodCuentaTraslacionGanancia = ""
    If cboCuentaTraslacionGanancia.ListIndex < 0 Then Exit Sub
    
    strCodCuentaTraslacionGanancia = Trim(arrCuentaTraslacionGanancia(cboCuentaTraslacionGanancia.ListIndex))


End Sub

Private Sub cboCuentaTraslacionPerdida_Click()

    strCodCuentaTraslacionPerdida = ""
    If cboCuentaTraslacionPerdida.ListIndex < 0 Then Exit Sub
    
    strCodCuentaTraslacionPerdida = Trim(arrCuentaTraslacionPerdida(cboCuentaTraslacionPerdida.ListIndex))


End Sub

Private Sub cboGrupoContable_Click()

    strCodGrupoContable = Valor_Caracter
    If cboGrupoContable.ListIndex < 0 Then Exit Sub
    
    strCodGrupoContable = Trim(arrGrupoContable(cboGrupoContable.ListIndex))
    
    Call Buscar
    
End Sub

Private Sub cboMoneda_Click()

    strCodMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
    If strCodMoneda = Valor_Caracter Then strCodMoneda = "00"
    
End Sub

Private Sub cboMovimiento_Click()

    strCodMovimiento = Valor_Caracter
    If cboMovimiento.ListIndex < 0 Then Exit Sub
    
    strCodMovimiento = Trim(arrMovimiento(cboMovimiento.ListIndex))
    
End Sub

Private Sub cboNaturaleza_Click()

    strCodNaturaleza = Valor_Caracter
    If cboNaturaleza.ListIndex < 0 Then Exit Sub
    
    strCodNaturaleza = Trim(arrNaturaleza(cboNaturaleza.ListIndex))
    
    If strCodNaturaleza = Valor_Caracter Then strCodNaturaleza = "T"
    
End Sub

Private Sub cboRubroEEFF_Click()

    strCodRubroEEFF = ""
    If cboRubroEEFF.ListIndex < 0 Then Exit Sub
    
    strCodRubroEEFF = Trim(arrRubroEEFF(cboRubroEEFF.ListIndex))
    
End Sub

Private Sub cboTipoAuxiliar_Click()

    strTipoAuxiliar = ""
    
    If cboTipoAuxiliar.ListIndex < 0 Then Exit Sub
    
    strTipoAuxiliar = Trim(arrTipoAuxiliar(cboTipoAuxiliar.ListIndex))
    
    If strTipoAuxiliar = Valor_Caracter Then strTipoAuxiliar = "00"

End Sub

Private Sub cboTipoCuenta_Click()

    strCodTipoCuenta = ""
    If cboTipoCuenta.ListIndex < 0 Then Exit Sub
    
    strCodTipoCuenta = Trim(arrTipoCuenta(cboTipoCuenta.ListIndex))
    
    If strCodTipoCuenta = Codigo_Tipo_Cuenta_Cuenta Then
        lblGrupoActual.Caption = strCodGrupoContableSel
        txtCodCuenta.MaxLength = 10
        Call Habilita
    Else
        lblGrupoActual.Caption = Valor_Caracter
        txtCodCuenta.MaxLength = 2
        Call Deshabilita
    End If
    
End Sub

Private Sub Habilita()
    
'    cboMovimiento.Enabled = True
'    cboMoneda.Enabled = True
'    cboNaturaleza.Enabled = True
'    cboCuentaPerdida.Enabled = True
'    cboCuentaGanancia.Enabled = True
'    txtCodCuenta.Enabled = True
    
    lblDescrip(17).Visible = True
    lblGrupoActual.Visible = True
    cboMovimiento.Enabled = True
    cboMoneda.Enabled = True
    cboNaturaleza.Enabled = True
    cboCuentaPerdida.Enabled = True
    cboCuentaGanancia.Enabled = True
    
    If cboCuentaPerdida.ListCount >= 0 Then cboCuentaPerdida.ListIndex = 0
    If cboCuentaGanancia.ListCount >= 0 Then cboCuentaGanancia.ListIndex = 0
    
    
End Sub

Private Sub Deshabilita()
    
'    txtCodCuenta.Enabled = False
    
    lblDescrip(17).Visible = False
    lblGrupoActual.Visible = False
    cboMovimiento.Enabled = False
    cboMoneda.Enabled = False
    cboNaturaleza.Enabled = False
    cboCuentaPerdida.Enabled = False
    cboCuentaGanancia.Enabled = False
    
End Sub

Private Sub cboTipoEEFF_Click()

    Dim strSQL As String
    
    strCodTipoEEFF = ""
    If cboTipoEEFF.ListIndex < 0 Then Exit Sub
    
    strCodTipoEEFF = Trim(arrTipoEEFF(cboTipoEEFF.ListIndex))
    
    '*** Rubro EEFF ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='RUEEFF' AND ValorParametro='" & strCodTipoEEFF & "' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboRubroEEFF, arrRubroEEFF(), ""
    
End Sub


Private Sub cboTipoFile_Click()

    strTipoFile = ""
    
    If cboTipoFile.ListIndex < 0 Then Exit Sub
    
    strTipoFile = Trim(arrTipoFile(cboTipoFile.ListIndex))

End Sub

Private Sub chkAuxiliar_Click()

    If chkAuxiliar.Value = vbChecked Then
        lblDescrip(12).Visible = True
        cboTipoAuxiliar.Visible = True
        strIndicadorAuxiliar = Valor_Indicador
    Else
        lblDescrip(12).Visible = False
        cboTipoAuxiliar.Visible = False
        cboTipoAuxiliar.ListIndex = -1
        strIndicadorAuxiliar = Valor_Caracter
    End If

End Sub

Private Sub chkPartidaMonetaria_Click()

    If chkPartidaMonetaria.Value = vbChecked Then
        cboCuentaPerdida.Enabled = True
        cboCuentaGanancia.Enabled = True
        strIndicadorPartidaMonetaria = Valor_Indicador
    Else
        cboCuentaPerdida.ListIndex = -1
        cboCuentaPerdida.Enabled = False
        cboCuentaGanancia.ListIndex = -1
        cboCuentaGanancia.Enabled = False
        strIndicadorPartidaMonetaria = Valor_Caracter
    End If


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
    'Call Buscar
    Call DarFormato
    
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
    
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
            
End Sub
Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Plan de Cuentas"
    
End Sub

Private Sub CargarListas()

    Dim strSQL  As String
    
    NumVersionPlanContable = ObtenerPlanContableVersion()
    
    If NumVersionPlanContable = -1 Then
        MsgBox "No existe version de Plan Contable vigente", vbCritical
        Exit Sub
    End If
    '*** Grupo de Cuentas ***
    strSQL = "{ call up_ACSelDatosParametro(30,'" & gstrCodAdministradora & "') }"
    CargarControlLista strSQL, cboGrupoContable, arrGrupoContable(), Sel_Todos
    
    If cboGrupoContable.ListCount > 0 Then cboGrupoContable.ListIndex = 0
    
    '*** Tipo Cuenta ***
    'strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='TIPPLN' ORDER BY DescripParametro"
    CargarControlListaAuxiliarParametro "TIPPLN", cboTipoCuenta, arrTipoCuenta(), ""
    
    '*** Tipo EEFF ***
    'strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='TIEEFF' ORDER BY DescripParametro"
    CargarControlListaAuxiliarParametro "TIEEFF", cboTipoEEFF, arrTipoEEFF(), Sel_Defecto

    '*** Movimiento ***
    'strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='RESPSN' ORDER BY DescripParametro"
    CargarControlListaAuxiliarParametro "RESPSN", cboMovimiento, arrMovimiento(), ""
    
    '*** Tipo de Auxiliar ***
    CargarControlListaAuxiliarParametro "TIPAUX", cboTipoAuxiliar, arrTipoAuxiliar(), Sel_Todos
 
    '*** Tipo de File ***
    CargarControlListaAuxiliarParametro "TIPFIL", cboTipoFile, arrTipoFile(), ""
    
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Sel_Todos 'ojo
    
    '*** Naturaleza ***
    'strSQL = "SELECT ValorParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='NATCTA' ORDER BY DescripParametro"
    CargarControlListaAuxiliarParametro "NATCTA", cboNaturaleza, arrNaturaleza(), Sel_Todos
    
    '*** Cuenta Pérdida ***
    strSQL = "SELECT CodCuenta CODIGO,(CodCuenta + space(1) + DescripCuenta) DESCRIP FROM PlanContable WHERE IndMovimiento='X' AND CodCuenta LIKE '67611%' AND CodAdministradora='" & gstrCodAdministradora & "' AND NumVersion=" & NumVersionPlanContable
    CargarControlLista strSQL, cboCuentaPerdida, arrCuentaPerdida(), Sel_Defecto
    
    cboCuentaPerdida.ListIndex = 1

    '*** Cuenta Ganancia ***
    strSQL = "SELECT CodCuenta CODIGO,(CodCuenta + space(1) + DescripCuenta) DESCRIP FROM PlanContable WHERE IndMovimiento='X' AND CodCuenta LIKE '77611%' AND CodAdministradora='" & gstrCodAdministradora & "' AND NumVersion=" & NumVersionPlanContable
    CargarControlLista strSQL, cboCuentaGanancia, arrCuentaGanancia(), Sel_Defecto
        
    cboCuentaGanancia.ListIndex = 0
        
    '*** Cuenta Pérdida Traslacion ***
    strSQL = "SELECT CodCuenta CODIGO,(CodCuenta + space(1) + DescripCuenta) DESCRIP FROM PlanContable WHERE IndMovimiento='X' AND CodCuenta LIKE '67612%' AND CodAdministradora='" & gstrCodAdministradora & "' AND NumVersion=" & NumVersionPlanContable
    CargarControlLista strSQL, cboCuentaTraslacionPerdida, arrCuentaTraslacionPerdida(), Sel_Defecto

    '*** Cuenta Ganancia Traslacion ***
    strSQL = "SELECT CodCuenta CODIGO,(CodCuenta + space(1) + DescripCuenta) DESCRIP FROM PlanContable WHERE IndMovimiento='X' AND CodCuenta LIKE '77612%' AND CodAdministradora='" & gstrCodAdministradora & "' AND NumVersion=" & NumVersionPlanContable
    CargarControlLista strSQL, cboCuentaTraslacionGanancia, arrCuentaTraslacionGanancia(), Sel_Defecto


End Sub
Private Sub InicializarValores()
    
    strEstado = Reg_Defecto
    tabPlanCuenta.Tab = 0
    tabPlanCuenta.TabEnabled(1) = False
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    Set frmPlanCuenta = Nothing
    
End Sub

Private Sub otlCuenta_Click()

    If otlCuenta.Expand(otlCuenta.ListIndex) = True Then
        otlCuenta.Expand(otlCuenta.ListIndex) = False
        If otlCuenta.HasSubItems(otlCuenta.ListIndex) = False Then
            otlCuenta.PictureType(otlCuenta.ListIndex) = 2
        Else
            otlCuenta.PictureType(otlCuenta.ListIndex) = 0
        End If
    Else
        otlCuenta.Expand(otlCuenta.ListIndex) = True
        If otlCuenta.HasSubItems(otlCuenta.ListIndex) Then
            otlCuenta.PictureType(otlCuenta.ListIndex) = 1
        End If
        strCodGrupoContableSel = Format(otlCuenta.ItemData(otlCuenta.ListIndex), "00")
    
    End If
    

End Sub

Private Sub otlCuenta_DblClick()

    Call Modificar
    
End Sub

Private Sub scrol_SpinDown()

    If Str(Val(lblNivelCuenta) - 1) > 0 Then
        lblNivelCuenta = CStr(Val(lblNivelCuenta) - 1)
    End If
    
End Sub

Private Sub scrol_SpinUp()

    If Str(Val(lblNivelCuenta) + 1) < 100 Then
        lblNivelCuenta = CStr(Val(lblNivelCuenta) + 1)
    End If
    
End Sub

Private Sub txtCodCuenta_Change()

    If Len(Trim(txtCodCuenta.Text)) > 1 Then
        lblNivelCuenta.Caption = CStr(Len(Trim(txtCodCuenta.Text)) - 1)
    End If
    
End Sub


