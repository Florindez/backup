VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmQuiebreDeposito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quiebre de Depósitos"
   ClientHeight    =   6660
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   11640
   Begin VB.Frame fraListaDepositos 
      Caption         =   "Lista de Depositos"
      Height          =   6405
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   8280
         TabIndex        =   34
         Top             =   5520
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
      Begin VB.ComboBox cboTitulo 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1020
         Width           =   8895
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos de Operación"
         Height          =   3675
         Left            =   360
         TabIndex        =   3
         Top             =   1650
         Width           =   10635
         Begin VB.ComboBox cboBaseAnual 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2610
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   2085
            Width           =   2235
         End
         Begin VB.ComboBox cboTipoTasa 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2610
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1650
            Width           =   2235
         End
         Begin VB.TextBox txtFileAnalitica 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   7650
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "003-00000008"
            Top             =   540
            Width           =   2235
         End
         Begin TAMControls.TAMTextBox txtMontoInversion 
            Height          =   315
            Left            =   7650
            TabIndex        =   7
            Top             =   1650
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmQuiebreDeposito.frx":0000
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtValorTipoCambio 
            Height          =   315
            Left            =   7650
            TabIndex        =   8
            Top             =   960
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmQuiebreDeposito.frx":001C
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
            MaximoValor     =   999999999
         End
         Begin MSComCtl2.DTPicker dtpFechaProceso 
            Height          =   315
            Left            =   2610
            TabIndex        =   11
            Top             =   960
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   175570945
            CurrentDate     =   38068
         End
         Begin TAMControls.TAMTextBox txtMontoOperacion 
            Height          =   315
            Left            =   7650
            TabIndex        =   12
            Top             =   2970
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmQuiebreDeposito.frx":0038
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   999999999
         End
         Begin MSComCtl2.DTPicker dtpFechaApertura 
            Height          =   315
            Left            =   2610
            TabIndex        =   18
            Top             =   540
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   175570945
            CurrentDate     =   38068
         End
         Begin TAMControls.TAMTextBox txtTasaInteresDeposito 
            Height          =   315
            Left            =   2610
            TabIndex        =   20
            Top             =   2520
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmQuiebreDeposito.frx":0054
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtTasaInteresQuiebre 
            Height          =   315
            Left            =   2610
            TabIndex        =   22
            Top             =   2970
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmQuiebreDeposito.frx":0070
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtMontoInteresDevengado 
            Height          =   315
            Left            =   7650
            TabIndex        =   24
            Top             =   2070
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmQuiebreDeposito.frx":008C
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtMontoInteresCastigado 
            Height          =   315
            Left            =   7650
            TabIndex        =   27
            Top             =   2520
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmQuiebreDeposito.frx":00A8
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   999999999
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Base Anual"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   360
            TabIndex        =   33
            Top             =   2115
            Width           =   810
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Tasa"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   360
            TabIndex        =   32
            Top             =   1680
            Width           =   720
         End
         Begin VB.Line Line3 
            X1              =   270
            X2              =   9660
            Y1              =   1410
            Y2              =   1410
         End
         Begin VB.Line Line2 
            X1              =   5220
            X2              =   9870
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Line1 
            X1              =   5160
            X2              =   9870
            Y1              =   2430
            Y2              =   2430
         End
         Begin VB.Label lblMoneda 
            Caption         =   "USD"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   3
            Left            =   9960
            TabIndex        =   29
            Top             =   2580
            Width           =   435
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Intereses Castigados"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   11
            Left            =   5220
            TabIndex        =   28
            Top             =   2580
            Width           =   2355
         End
         Begin VB.Label lblMoneda 
            Caption         =   "USD"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   9960
            TabIndex        =   26
            Top             =   2100
            Width           =   435
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Intereses Devengados"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   10
            Left            =   5220
            TabIndex        =   25
            Top             =   2100
            Width           =   2175
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tasa Interes Quiebre (%)"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   9
            Left            =   360
            TabIndex        =   23
            Top             =   3000
            Width           =   2235
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tasa Interes Anual (%)"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   8
            Left            =   360
            TabIndex        =   21
            Top             =   2550
            Width           =   2205
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha Apertura"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   7
            Left            =   360
            TabIndex        =   19
            Top             =   570
            Width           =   1935
         End
         Begin VB.Label lblDescrip 
            Caption         =   "File Analitica"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   5
            Left            =   5220
            TabIndex        =   17
            Top             =   570
            Width           =   1935
         End
         Begin VB.Label lblMoneda 
            Caption         =   "USD"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   9960
            TabIndex        =   15
            Top             =   3030
            Width           =   435
         End
         Begin VB.Label lblMoneda 
            Caption         =   "USD"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   9960
            TabIndex        =   14
            Top             =   1680
            Width           =   435
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Monto Operación"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   4
            Left            =   5220
            TabIndex        =   13
            Top             =   3030
            Width           =   2385
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo de Cambio"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   1
            Left            =   5220
            TabIndex        =   6
            Top             =   1020
            Width           =   1695
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Inversión"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   6
            Left            =   5220
            TabIndex        =   5
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha Quiebre"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   360
            TabIndex        =   4
            Top             =   1020
            Width           =   1935
         End
      End
      Begin VB.ComboBox cboFondo 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   570
         Width           =   8895
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Deposito"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fondo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmQuiebreDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Ficha de Quiebre de Depósitos"
Option Explicit

Dim arrFondo()          As String
Dim arrTitulo()         As String
Dim arrTipoTasa()       As String
Dim arrBaseAnual()      As String

Dim strFechaSiguiente As String, strFechaActual As String
Dim strCodFondo As String
Dim strCodMoneda As String, strCodTitulo As String
Dim strCodFile As String, strCodAnalitica As String
Dim strAccion As String
Dim aDirReg() As Variant
Dim strSQL As String
Dim intBaseCalculo  As Integer, strCodBaseAnual    As String
Dim strCodTipoTasa  As String, dblMontoInteresCastigado As Double
Dim intDiasPlazo    As Integer, intNumDias30 As Integer

Private Sub cboBaseAnual_Click()

    strCodBaseAnual = Valor_Caracter
    If cboBaseAnual.ListIndex < 0 Then Exit Sub
    
    strCodBaseAnual = Trim(arrBaseAnual(cboBaseAnual.ListIndex))
    
    '*** Base de Cálculo ***
    intBaseCalculo = 365
    Select Case strCodBaseAnual
        Case Codigo_Base_30_360: intBaseCalculo = 360
        Case Codigo_Base_Actual_365: intBaseCalculo = 365
        Case Codigo_Base_Actual_360: intBaseCalculo = 360
        Case Codigo_Base_30_365: intBaseCalculo = 365
        Case Codigo_Base_Actual_Actual
            Dim adoRegistro     As ADODB.Recordset
            
            Set adoRegistro = New ADODB.Recordset
        
            adoComm.CommandText = "SELECT dbo.uf_ACValidaEsBisiesto(" & CInt(Right(dtpFechaProceso.Value, 4)) & ") AS 'EsBisiesto'"
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                If adoRegistro("EsBisiesto") = 0 Then intBaseCalculo = 366
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
    End Select
    
End Sub



Private Sub cboTipoTasa_Click()

    strCodTipoTasa = Valor_Caracter
    If cboTipoTasa.ListIndex < 0 Then Exit Sub
    
    strCodTipoTasa = Trim(arrTipoTasa(cboTipoTasa.ListIndex))
    
End Sub

Private Sub cboTitulo_Click()

    Dim adoConsulta     As ADODB.Recordset
    Dim adoRegistro     As ADODB.Recordset
    Dim intRegistro     As Integer

    strCodTitulo = Valor_Caracter: strCodAnalitica = Valor_Caracter
    
    If cboTitulo.ListIndex < 0 Then Exit Sub

    strCodTitulo = Trim(arrTitulo(cboTitulo.ListIndex))

    With adoComm
        Set adoConsulta = New ADODB.Recordset

        .CommandText = "SELECT CodFile,CodAnalitica,CodMoneda, FechaEmision, TasaInteres, CodTipoTasa, BaseAnual FROM InstrumentoInversion WHERE CodTitulo='" & strCodTitulo & "'"
        Set adoConsulta = .Execute

        If Not adoConsulta.EOF Then
            strCodFile = Trim(adoConsulta("CodFile"))
            strCodAnalitica = Trim(adoConsulta("CodAnalitica"))
            strCodMoneda = Trim(adoConsulta("CodMoneda"))
            dtpFechaApertura.Value = Trim(adoConsulta("FechaEmision"))
            
            intDiasPlazo = DateDiff("d", dtpFechaApertura.Value, dtpFechaProceso.Value)
            
            intNumDias30 = Dias360(dtpFechaApertura.Value, dtpFechaProceso.Value, True)

            intRegistro = ObtenerItemLista(arrTipoTasa(), adoConsulta("CodTipoTasa"))
            If intRegistro >= 0 Then cboTipoTasa.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrBaseAnual(), adoConsulta("BaseAnual"))
            If intRegistro >= 0 Then cboBaseAnual.ListIndex = intRegistro
            
            txtFileAnalitica.Text = strCodFile + "-" + strCodAnalitica
            
            lblMoneda(0).Caption = ObtenerCodSignoMoneda(strCodMoneda) 'ObtenerDescripcionMoneda(strCodMoneda)
            lblMoneda(1).Caption = lblMoneda(0).Caption 'ObtenerCodSignoMoneda(strCodMoneda) 'ObtenerDescripcionMoneda(strCodMoneda)
            lblMoneda(2).Caption = lblMoneda(0).Caption 'ObtenerCodSignoMoneda(strCodMoneda) 'ObtenerDescripcionMoneda(strCodMoneda)
            lblMoneda(3).Caption = lblMoneda(0).Caption 'ObtenerCodSignoMoneda(strCodMoneda) 'ObtenerDescripcionMoneda(strCodMoneda)
                                
            Set adoRegistro = New ADODB.Recordset
        
            adoComm.CommandText = "SELECT dbo.uf_CNObtenerSaldoInversionCosto('" & strCodFondo & "','" & gstrCodAdministradora & "','" & strCodFile & "','" & strCodAnalitica & "','" & strCodMoneda & "','" & strFechaActual & "','" & strCodMoneda & "') AS 'SaldoInversionCosto'"
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                txtMontoInversion.Text = CStr(adoRegistro("SaldoInversionCosto"))
            End If
            adoRegistro.Close
                                
            adoComm.CommandText = "SELECT dbo.uf_CNObtenerSaldoInversionValorRazonable('" & strCodFondo & "','" & gstrCodAdministradora & "','" & strCodFile & "','" & strCodAnalitica & "','" & strCodMoneda & "','" & strFechaActual & "','" & strCodMoneda & "') AS 'SaldoInversionValorRazonable'"
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                txtMontoInteresDevengado.Text = CStr(adoRegistro("SaldoInversionValorRazonable"))
            End If
            adoRegistro.Close ': Set adoRegistro = Nothing
                                
            txtMontoOperacion.Text = CStr(txtMontoInversion.Value + txtMontoInteresDevengado.Value)
            
            txtTasaInteresDeposito.Text = adoConsulta("TasaInteres")
            txtTasaInteresQuiebre.Text = adoConsulta("TasaInteres")
                                
        End If
        adoConsulta.Close: Set adoConsulta = Nothing

    End With

End Sub

Private Sub cboFondo_Click()

        
    Dim adoRegistro As ADODB.Recordset, strCodFile As String
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            dtpFechaProceso.Value = adoRegistro("FechaCuota")
            'dtpFechaPago.Value = DateAdd("d", gintDiasPagoRescate, dtpFechaProceso.Value)
            strCodMoneda = adoRegistro("CodMoneda")
            
            txtValorTipoCambio.Text = ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaProceso.Value, strCodMoneda, Codigo_Moneda_Local)
            
            If txtValorTipoCambio.Value = 0 Then txtValorTipoCambio.Value = ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaProceso.Value), strCodMoneda, Codigo_Moneda_Local)
            
            gdatFechaActual = adoRegistro("FechaCuota")
            strFechaActual = Convertyyyymmdd(gdatFechaActual)
            strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, CVDate(adoRegistro("FechaCuota"))))
            
            gstrPeriodoActual = CStr(Year(gdatFechaActual))
            gstrMesActual = Format(Month(gdatFechaActual), "00")
            
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        
            strCodFile = "006" 'DPZ
        
            strSQL = "{call up_IVLstOperacionesVigentes ('" & strCodFondo & "','" & gstrCodAdministradora & "','" & strCodFile & "','" & gstrFechaActual & "') }"
                                
            CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
        
        
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
        
               
   
End Sub

Private Sub Form_Activate()

    frmMainMdi.stbMdi.Panels(3).Text = Me.Caption
    
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
    
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
End Sub
Private Sub CargarListas()

    Dim intRegistro As Integer

    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
   
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
        
    '*** Tipo Tasa ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='NATTAS' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoTasa, arrTipoTasa(), ""
        
    '*** Base de Cálculo ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='BASANU' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboBaseAnual, arrBaseAnual(), Valor_Caracter
        
End Sub
Private Sub CargarReportes()

    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Registro de Quiebre de Depositos"
    
End Sub
Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
        
End Sub

Private Sub InicializarValores()

    dtpFechaProceso.Value = gdatFechaActual
    
    'Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set frmQuiebreDeposito = Nothing
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub


Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
                
        Case vSave
            Call Grabar
        Case vCancel
            Call Cancelar
        
    End Select
    
End Sub
Private Sub Cancelar()

    Unload Me

End Sub

Private Sub Grabar()

    Dim adoRegistro             As ADODB.Recordset, adoConsulta     As ADODB.Recordset
    Dim dblPrecioCierre         As Double, dblPrecioPromedio        As Double
    Dim dblTirCierre            As Double, dblFactorDiario          As Double
    Dim dblTasaInteres          As Double, dblFactorDiarioCupon     As Double
    Dim dblPrecioUnitario       As Double, dblValorPromedioKardex   As Double
    Dim dblInteresCorridoPromedio As Double, dblTirOperacionKardex  As Double
    Dim dblTirPromedioKardex    As Double, dblTirNetaKardex         As Double
    Dim curSaldoInversion       As Currency, curSaldoInteresCorrido As Currency
    Dim curSaldoFluctuacion     As Currency, curValorAnterior       As Currency
    Dim curValorActual          As Currency, curMontoRenta          As Currency
    Dim curMontoContable        As Currency, curMontoMovimientoMN   As Currency
    Dim curMontoMovimientoME    As Currency, curSaldoValorizar      As Currency
    Dim curCantMovimiento       As Currency, curKarValProm          As Currency
    Dim curValorMovimiento      As Currency, curSaldoInicialKardex  As Currency
    Dim curSaldoFinalKardex     As Currency, curValorSaldoKardex    As Currency
    Dim curValComi              As Currency, curVacCorrido          As Currency
    Dim curSaldoAmortizacion    As Currency
    Dim intCantRegistros        As Integer, intContador             As Integer
    Dim intRegistro             As Integer, intBaseCalculo          As Integer
    Dim intDiasPlazo            As Integer, intDiasDeRenta          As Integer
    Dim strNumAsiento           As String, strDescripAsiento        As String
    Dim strNumOperacion         As String, strNumKardex             As String
    Dim strNumCaja              As String, strFechaPago             As String
    Dim strCodTitulo            As String, strCodEmisor             As String
    Dim strDescripMovimiento    As String, strIndDebeHaber          As String
    Dim strCodCuenta            As String, strFiles                 As String
    Dim strModalidadInteres     As String
    Dim strCodTasa              As String, strIndCuponCero          As String
    Dim strCodDetalleFile       As String
    Dim strCodSubDetalleFile    As String, strFechaGrabar           As String
    Dim strSQLOperacion         As String, strSQLKardex             As String
    Dim strSQLOrdenCaja         As String, strSQLOrdenCajaDetalle   As String
    Dim strIndUltimoMovimiento  As String, strTipoMovimientoKardex  As String
    Dim blnVenceTitulo          As Boolean, blnVenceCupon           As Boolean
    Dim dblTipoCambioCierre     As Double
    Dim strCodModulo            As String, strMensaje   As String
    
    Dim dblValorTipoCambio          As Double, strTipoDocumento             As String
    Dim strNumDocumento             As String
    Dim strTipoPersonaContraparte   As String, strCodPersonaContraparte     As String
    Dim strIndContracuenta          As String, strCodContracuenta           As String
    Dim strCodFileContracuenta       As String, strCodAnaliticaContracuenta  As String
    
    '*** Verificación de Vencimiento de Valores de Depósito ***
    frmMainMdi.stbMdi.Panels(3).Text = "Quiebre de Depósito..."
       
    Set adoRegistro = New ADODB.Recordset
    
    strMensaje = "Para proceder al Registro de la Orden Confirme lo siquiente : " & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
        "Fecha de Operación" & Space(4) & ">" & Space(2) & CStr(dtpFechaProceso.Value) & Chr(vbKeyReturn) & _
        "Fecha de Liquidación" & Space(3) & ">" & Space(2) & CStr(dtpFechaProceso.Value) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
        "Monto Total" & Space(17) & ">" & Space(2) & Trim(lblMoneda(0).Caption) & Space(1) & CStr(txtMontoInversion.Text) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
        Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
        "¿ Seguro de continuar ?"

    If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
       Me.Refresh: Exit Sub
    End If
    
    Me.MousePointer = vbHourglass

    With adoComm
        '*** Obtener datos del último movimiento del kardex ***
        .CommandText = "SELECT IK.CodTitulo,II.CodMoneda,IK.CodAnalitica,IK.CodFile,SaldoFinal,CodDetalleFile,CodSubDetalleFile," & _
            "ValorNominal,CodTipoTasa,BaseAnual,CodTasa,TasaInteres,DiasPlazo,IndCuponCero,FechaEmision,FechaVencimiento," & _
            "II.CodEmisor,IK.PrecioUnitario,IK.MontoMovimiento,IK.SaldoInteresCorrido,IK.MontoSaldo,IK.ValorPromedioInteresCorrido " & _
            "FROM InversionKardex IK JOIN InstrumentoInversion II " & _
            "ON(II.CodTitulo=IK.CodTitulo) " & _
            "WHERE IK.CodFile = '" & strCodFile & "' AND IK.CodAnalitica = '" & strCodAnalitica & "' AND FechaOperacion <='" & strFechaActual & "' AND SaldoFinal > 0 AND " & _
            "IK.CodAdministradora='" & gstrCodAdministradora & "' AND IK.CodFondo='" & strCodFondo & "' AND " & _
            "IK.NumKardex IN (SELECT MAX(IKM.NumKardex) FROM InversionKardex IKM " & _
                              "WHERE IKM.CodAdministradora = IK.CodAdministradora " & _
                              "AND IKM.CodFondo = IK.CodFondo " & _
                              "AND IKM.CodFile = IK.CodFile " & _
                              "AND IKM.CodTitulo = IK.CodTitulo " & _
                              "AND IKM.FechaOperacion < '" & strFechaSiguiente & "')"
        
        Set adoRegistro = .Execute
    
   
        If Not adoRegistro.EOF Then
            
            '*** Obtener Secuenciales ***
            strNumAsiento = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumComprobante)
            strNumOperacion = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOperacion)
            strNumKardex = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumKardex)
            strNumCaja = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOrdenCaja)
            
             '*** Si vence el título o el cupón ***
            strCodTitulo = Trim(adoRegistro("CodTitulo"))
            strCodFile = Trim(adoRegistro("CodFile"))
            strCodDetalleFile = Trim(adoRegistro("CodDetalleFile"))
            strCodSubDetalleFile = Trim(adoRegistro("CodSubDetalleFile"))
            strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
            strCodEmisor = Trim(adoRegistro("CodEmisor"))
            strModalidadInteres = Trim(adoRegistro("CodDetalleFile"))
            strCodTasa = Trim(adoRegistro("CodTasa"))
            dblTasaInteres = CDbl(adoRegistro("TasaInteres"))
            intDiasPlazo = CInt(adoRegistro("DiasPlazo"))
            strIndCuponCero = Trim(adoRegistro("IndCuponCero"))
            curSaldoValorizar = CCur(adoRegistro("SaldoFinal"))
            curKarValProm = CDbl(adoRegistro("ValorPromedioInteresCorrido"))
            intDiasDeRenta = DateDiff("d", CVDate(adoRegistro("FechaEmision")), gdatFechaActual) + 1
        
            Set adoConsulta = New ADODB.Recordset
                    
            '*** Verificar Dinamica Contable ***
            .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
                "WHERE TipoOperacion='" & Codigo_Dinamica_Vencimiento & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"

            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                If CInt(adoConsulta("NumRegistros")) > 0 Then
                    intCantRegistros = CInt(adoConsulta("NumRegistros"))
                Else
                    MsgBox "NO EXISTE Dinámica Contable para la valorización", vbCritical
                    adoConsulta.Close: Set adoConsulta = Nothing
                    Exit Sub
                End If
            End If
            adoConsulta.Close
            
            strFechaPago = strFechaActual
        
            '*** Obtener las cuentas de inversión ***
            Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, Trim(adoRegistro("CodMoneda")))
            
            '*** Obtener tipo de cambio ***
            'dblTipoCambioCierre = ObtenerValorTCCierre(adoRegistro("CodMoneda"))
            dblTipoCambioCierre = ObtenerValorTipoCambio(adoRegistro("CodMoneda"), Codigo_Moneda_Local, strFechaActual, strFechaActual, gstrCodClaseTipoCambioOperacionFondo, Codigo_Valor_TipoCambioCompra)

            
            '*** Obtener Saldo de Inversión ***
            If Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                "WHERE (FechaSaldo >='" & strFechaActual & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaInversion & "' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoInversion = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Interés Corrido ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                "WHERE (FechaSaldo >='" & strFechaActual & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaInteresCorrido & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoInteresCorrido = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Obtener Saldo de Provisión ***
            If adoRegistro("CodMoneda") = Codigo_Moneda_Local Then
                .CommandText = "SELECT SaldoFinalContable Saldo "
            Else
                .CommandText = "SELECT SaldoFinalME Saldo "
            End If
            .CommandText = .CommandText & "FROM PartidaContableSaldos " & _
                "WHERE (FechaSaldo >='" & strFechaActual & "' AND FechaSaldo <'" & strFechaSiguiente & "') AND " & _
                "CodCuenta='" & strCtaProvInteres & " ' AND CodAnalitica='" & adoRegistro("CodAnalitica") & "' AND " & _
                "CodFile='" & strCodFile & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodMoneda = '" & Trim(adoRegistro("CodMoneda")) & "' AND " & _
                "CodMonedaContable = '" & Codigo_Moneda_Local & "'"
            
            Set adoConsulta = .Execute
            
            If Not adoConsulta.EOF Then
                curSaldoFluctuacion = CDbl(adoConsulta("Saldo"))
            End If
            adoConsulta.Close
            
            '*** Calculos ***
            curCtaXCobrar = CCur(txtMontoOperacion.Value) 'Round(curSaldoInversion + curSaldoInteresCorrido + curSaldoFluctuacion, 2)
            curCtaInversion = curSaldoInversion
            curCtaCosto = curSaldoInversion
            curCtaInteresCorrido = curSaldoInteresCorrido
            curCtaInteresCastigado = CCur(txtMontoInteresCastigado.Value)
            curCtaProvInteres = curSaldoFluctuacion
            curCtaIngresoOperacional = curSaldoInversion + curSaldoInteresCorrido + curSaldoFluctuacion - curCtaProvInteres
            
            curCantMovimiento = CCur(adoRegistro("SaldoFinal"))
            dblPrecioUnitario = curKarValProm
            curValorMovimiento = dblPrecioUnitario * curCantMovimiento * CCur(adoRegistro("ValorNominal")) * -1
            curSaldoInicialKardex = CCur(adoRegistro("SaldoFinal"))
            curSaldoFinalKardex = 0
            curValorSaldoKardex = 0
            dblValorPromedioKardex = 0
            dblInteresCorridoPromedio = 0
            curValComi = 0
            curVacCorrido = 0
            dblTirOperacionKardex = 0
            dblTirPromedioKardex = 0
            dblTirNetaKardex = 0
            curSaldoAmortizacion = 0
            
            '************************
            '*** Armar sentencias ***
            '************************
            strDescripAsiento = "Quiebre" & Space(1) & "(" & strCodFile & "-" & strCodAnalitica & ")"
            '*** Operación ***
            strSQLOperacion = "{ call up_IVAdicInversionOperacion('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strNumOperacion & "','" & strFechaActual & "','" & strCodTitulo & "','" & Left(strFechaActual, 4) & "','" & _
                Mid(strFechaActual, 5, 2) & "','','" & Estado_Activo & "','" & strCodAnalitica & "','" & _
                strCodFile & "','" & strCodAnalitica & "','" & strCodDetalleFile & "','" & strCodSubDetalleFile & "','" & _
                Codigo_Caja_Vencimiento & "','','','" & strDescripAsiento & "','" & strCodEmisor & "','" & _
                "','','" & strFechaActual & "','" & strFechaActual & "','" & _
                strFechaActual & "','" & adoRegistro("CodMoneda") & "','" & adoRegistro("CodMoneda") & "','" & adoRegistro("CodMoneda") & "'," & CDec(adoRegistro("SaldoFinal")) & "," & CDec(gdblTipoCambio) & "," & CDec(gdblTipoCambio) & "," & _
                CDec(adoRegistro("ValorNominal")) & "," & CDec(adoRegistro("PrecioUnitario")) & "," & CDec(adoRegistro("MontoMovimiento")) & "," & CDec(adoRegistro("MontoMovimiento")) & "," & CDec(adoRegistro("SaldoInteresCorrido")) & "," & _
                "0,0,0,0,0,0,0,0,0," & CDec(curCtaXCobrar) & "," & CDec(curCtaXCobrar) & ",0,0,0,0,0,0,0,0,0," & _
                "0,0,0,0,0,0,'X','" & strNumAsiento & "','','','" & _
                "','','','','',0,'','','','',''," & CDec(dblTasaInteres) & "," & _
                "0,0,'','','','" & gstrLogin & "') }"
                                            
            strIndUltimoMovimiento = "X"
            strTipoMovimientoKardex = "S"
            '*** Kardex ***
            strSQLKardex = "{ call up_IVAdicInversionKardex('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strCodTitulo & "','" & strNumKardex & "','" & strFechaActual & "','" & Left(strFechaActual, 4) & "','" & _
                Mid(strFechaActual, 5, 2) & "','','" & strNumOperacion & "','" & strCodEmisor & "','','O','" & _
                strFechaActual & "','" & strTipoMovimientoKardex & "','O'," & curCantMovimiento & ",'" & adoRegistro("CodMoneda") & "'," & _
                dblPrecioUnitario & "," & dblPrecioUnitario & "," & dblPrecioUnitario & "," & curValorMovimiento & "," & curValComi & "," & curSaldoInicialKardex & "," & _
                curSaldoFinalKardex & "," & curValorSaldoKardex & ",'" & strDescripAsiento & "'," & dblValorPromedioKardex & ",'" & _
                strIndUltimoMovimiento & "','" & strCodFile & "','" & strCodAnalitica & "'," & dblInteresCorridoPromedio & "," & _
                curSaldoInteresCorrido & "," & dblTirOperacionKardex & "," & dblTirOperacionKardex & "," & dblTirPromedioKardex & "," & dblTirPromedioKardex & "," & curVacCorrido & "," & _
                dblTirNetaKardex & "," & curSaldoAmortizacion & ") }"

            '*** Orden de Cobro/Pago ***
            strSQLOrdenCaja = "{ call up_ACAdicMovimientoFondo('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strNumCaja & "','" & strFechaActual & "','" & Trim(frmMainMdi.Tag) & "','" & strNumOperacion & "','" & strFechaPago & "','" & _
                strNumAsiento & "','','','','E','" & strCtaXCobrar & "'," & CDec(curCtaXCobrar) & ",'" & _
                strCodFile & "','" & strCodAnalitica & "','" & adoRegistro("CodMoneda") & "','" & _
                strDescripAsiento & "','" & Codigo_Caja_Vencimiento & "','','" & Estado_Caja_NoConfirmado & "','','','','','','','','','','0','" & adoRegistro("CodEmisor") & "','02','" & gstrLogin & "') }"
            
            '*** Orden de Cobro/Pago Detalle ***
            strSQLOrdenCajaDetalle = "{ call up_ACAdicMovimientoFondoDetalle('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strNumCaja & "','" & strFechaActual & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripAsiento & "','" & _
                "H','" & strCtaXCobrar & "'," & CDec(curCtaXCobrar) * -1 & ",'" & _
                strCodFile & "','" & strCodAnalitica & "','','','" & adoRegistro("CodMoneda") & "','') }"
                            
            '*** Monto Orden ***
            If curCtaXCobrar > 0 Then
                                                                                        
'                    On Error GoTo Ctrl_Error
                
'                    .CommandText = "BEGIN TRANSACTION ProcAsiento"
'                    adoConn.Execute .CommandText
                                                        
                '*** Actualizar indicador de último movimiento en Kardex ***
                .CommandText = "UPDATE InversionKardex SET IndUltimoMovimiento='' " & _
                    "WHERE CodAnalitica='" & strCodAnalitica & "' AND CodFile='" & strCodFile & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "IndUltimoMovimiento='X'"
    
                adoConn.Execute .CommandText
    
                '*** Inserta movimiento en el kardex ***
                adoConn.Execute strSQLKardex
                
                '*** Operación ***
                adoConn.Execute strSQLOperacion
                
                '*** Contabilizar ***
                strFechaGrabar = strFechaActual & Space(1) & Format(Time, "hh:mm")
                
                '*** Cabecera ***
                .CommandText = "{ call up_ACAdicAsientoContable('"
                .CommandText = .CommandText & strCodFondo & "','" & gstrCodAdministradora & "','" & strNumAsiento & "','" & _
                    strFechaGrabar & "','" & _
                    Left(strFechaActual, 4) & "','" & Mid(strFechaActual, 5, 2) & "','" & _
                    "','" & _
                    strDescripAsiento & "','" & Trim(adoRegistro("CodMoneda")) & "','" & _
                    Codigo_Moneda_Local & "','" & _
                    "','" & _
                    "'," & _
                    CDec(curCtaXCobrar) & ",'" & Estado_Activo & "'," & _
                    intCantRegistros & ",'" & _
                    strFechaActual & Space(1) & Format(Time, "hh:ss") & "','" & _
                    strCodModulo & "','" & _
                    "'," & _
                    dblTipoCambioCierre & ",'" & _
                    "','" & _
                    "','" & _
                    strDescripAsiento & "','" & _
                    "','" & _
                    "X','') }"
                adoConn.Execute .CommandText
                
                '*** Detalle ***
                .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
                    "WHERE TipoOperacion='" & Codigo_Dinamica_Vencimiento & "' AND CodFile='" & strCodFile & "' AND (CodDetalleFile='" & _
                    strCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                    IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"
                Set adoConsulta = .Execute
        
                Do While Not adoConsulta.EOF
                
                    Select Case Trim(adoConsulta("TipoCuentaInversion"))
                        Case Codigo_CtaInversion
                            curMontoMovimientoMN = curCtaInversion
                            
                        Case Codigo_CtaProvInteres
                            curMontoMovimientoMN = curCtaProvInteres
                            
                        Case Codigo_CtaInteres
                            curMontoMovimientoMN = curCtaInteres
                            
                        Case Codigo_CtaInteresCastigado
                            curMontoMovimientoMN = curCtaInteresCastigado
                            
                        Case Codigo_CtaCosto
                            curMontoMovimientoMN = curCtaCosto
                            
                        Case Codigo_CtaIngresoOperacional
                            curMontoMovimientoMN = curCtaIngresoOperacional
                            
                        Case Codigo_CtaInteresVencido
                            curMontoMovimientoMN = curCtaInteresVencido
                            
                        Case Codigo_CtaVacCorrido
                            curMontoMovimientoMN = curCtaVacCorrido
                            
                        Case Codigo_CtaXPagar
                            curMontoMovimientoMN = curCtaXPagar
                            
                        Case Codigo_CtaXCobrar
                            curMontoMovimientoMN = curCtaXCobrar
                            
                        Case Codigo_CtaInteresCorrido
                            curMontoMovimientoMN = curCtaInteresCorrido
                            
                        Case Codigo_CtaProvReajusteK
                            curMontoMovimientoMN = curCtaProvReajusteK
                            
                        Case Codigo_CtaReajusteK
                            curMontoMovimientoMN = curCtaReajusteK
                            
                        Case Codigo_CtaProvFlucMercado
                            curMontoMovimientoMN = curCtaProvFlucMercado
                            
                        Case Codigo_CtaFlucMercado
                            curMontoMovimientoMN = curCtaFlucMercado
                            
                        Case Codigo_CtaProvInteresVac
                            curMontoMovimientoMN = curCtaProvInteresVac
                            
                        Case Codigo_CtaInteresVac
                            curMontoMovimientoMN = curCtaInteresVac
                            
                        Case Codigo_CtaIntCorridoK
                            curMontoMovimientoMN = curCtaIntCorridoK
                            
                        Case Codigo_CtaProvFlucK
                            curMontoMovimientoMN = curCtaProvFlucK
                            
                        Case Codigo_CtaFlucK
                            curMontoMovimientoMN = curCtaFlucK
                            
                        Case Codigo_CtaInversionTransito
                            curMontoMovimientoMN = curCtaInversionTransito
                            
                    End Select
                    
                    
                    strIndDebeHaber = Trim(adoConsulta("IndDebeHaber"))
                    If strIndDebeHaber = "H" Then
                        curMontoMovimientoMN = curMontoMovimientoMN * -1
                        If curMontoMovimientoMN > 0 Then strIndDebeHaber = "D"
                    ElseIf strIndDebeHaber = "D" Then
                        If curMontoMovimientoMN < 0 Then strIndDebeHaber = "H"
                    End If
                    
                    If strIndDebeHaber = "T" Then
                        If curMontoMovimientoMN > 0 Then
                            strIndDebeHaber = "D"
                        Else
                            strIndDebeHaber = "H"
                        End If
                    End If
                    
                    strDescripMovimiento = Trim(adoConsulta("DescripDinamica"))
                    curMontoMovimientoME = 0
                    curMontoContable = curMontoMovimientoMN
        
                    dblValorTipoCambio = 1
                    
                    If Trim(adoRegistro("CodMoneda")) <> Codigo_Moneda_Local Then
                        curMontoContable = Round(curMontoMovimientoMN * dblTipoCambioCierre, 2)
                        curMontoMovimientoME = curMontoMovimientoMN
                        curMontoMovimientoMN = 0
                        dblValorTipoCambio = dblTipoCambioCierre
                    End If
                                
                    strTipoDocumento = ""
                    strNumDocumento = ""
                    strTipoPersonaContraparte = Codigo_Tipo_Persona_Emisor
                    strCodPersonaContraparte = strCodEmisor
                    strIndContracuenta = ""
                    strCodContracuenta = ""
                    strCodFileContracuenta = ""
                    strCodAnaliticaContracuenta = ""
                    strIndUltimoMovimiento = ""
                                
                    If curMontoContable <> 0 Then
                                
                        '*** Movimiento ***
                        .CommandText = "{ call up_ACAdicAsientoContableDetalle('"
                        .CommandText = .CommandText & strNumAsiento & "','" & strCodFondo & "','" & _
                            gstrCodAdministradora & "'," & _
                            CInt(adoConsulta("NumSecuencial")) & ",'" & _
                            strFechaGrabar & "','" & _
                            Left(strFechaActual, 4) & "','" & _
                            Mid(strFechaActual, 5, 2) & "','" & _
                            strDescripMovimiento & "','" & _
                            strIndDebeHaber & "','" & _
                            Trim(adoConsulta("CodCuenta")) & "','" & _
                            Trim(adoRegistro("CodMoneda")) & "'," & _
                            CDec(curMontoMovimientoMN) & "," & _
                            CDec(curMontoMovimientoME) & "," & _
                            CDec(curMontoContable) & "," & _
                            dblValorTipoCambio & ",'" & _
                            Trim(adoRegistro("CodFile")) & "','" & _
                            Trim(adoRegistro("CodAnalitica")) & "','" & _
                            strTipoDocumento & "','" & _
                            strNumDocumento & "','" & _
                            strTipoPersonaContraparte & "','" & _
                            strCodPersonaContraparte & "','" & _
                            strIndContracuenta & "','" & _
                            strCodContracuenta & "','" & _
                            strCodFileContracuenta & "','" & _
                            strCodAnaliticaContracuenta & "','" & _
                            strIndUltimoMovimiento & "') }"
    
                        adoConn.Execute .CommandText
                    
                    End If
                                   
                    adoConsulta.MoveNext
                Loop
                adoConsulta.Close: Set adoConsulta = Nothing
                                
                '-- Verifica y ajusta posibles descuadres
                .CommandText = "{ call up_ACProcAsientoContableAjuste('" & _
                        strCodFondo & "','" & _
                        gstrCodAdministradora & "','" & _
                        strNumAsiento & "') }"
                adoConn.Execute .CommandText
                                
                '*** Orden de Cobro ***
                adoConn.Execute strSQLOrdenCaja
                adoConn.Execute strSQLOrdenCajaDetalle
    
                '*** Actualizar Secuenciales **
                .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    Valor_NumComprobante & "','" & strNumAsiento & "') }"
                adoConn.Execute .CommandText
                
                .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    Valor_NumOperacion & "','" & strNumOperacion & "') }"
                adoConn.Execute .CommandText
                
                .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    Valor_NumKardex & "','" & strNumKardex & "') }"
                adoConn.Execute .CommandText
                
                .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    Valor_NumOrdenCaja & "','" & strNumCaja & "') }"
                adoConn.Execute .CommandText
                                                                        
'                    .CommandText = "COMMIT TRANSACTION ProcAsiento"
'                    adoConn.Execute .CommandText
        
            End If
            
        End If
        
        adoRegistro.Close: Set adoRegistro = Nothing
                
    End With
    
    
    Me.MousePointer = vbDefault

    MsgBox Mensaje_Adicion_Exitosa, vbExclamation
    
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
    Exit Sub
       
Ctrl_Error:
'    adoComm.CommandText = "ROLLBACK TRANSACTION ProcAsiento"
'    adoConn.Execute adoComm.CommandText
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
    
End Sub

Private Sub txtTasaInteresQuiebre_Change()

    Call CalcularQuiebre
    
End Sub

Private Sub CalcularQuiebre()

    Dim dblMontoInversionQuiebre  As Double
                                
    dblMontoInversionQuiebre = CStr(ValorVencimiento(txtMontoInversion.Value, txtTasaInteresQuiebre.Value, intBaseCalculo, intDiasPlazo, intNumDias30, strCodTipoTasa, strCodBaseAnual)) 'ACR
    
    txtMontoOperacion.Text = dblMontoInversionQuiebre
                        
    dblMontoInteresCastigado = CDbl(txtMontoInversion.Value) + CDbl(txtMontoInteresDevengado.Value) - CDbl(txtMontoOperacion.Value)
            
    txtMontoInteresCastigado.Text = CStr(dblMontoInteresCastigado)

End Sub
