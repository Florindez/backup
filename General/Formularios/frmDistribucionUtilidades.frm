VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmDistribucionUtilidades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distribución de Utilidades"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   13995
   Begin VB.CommandButton cmdCargar 
      Caption         =   "&Cargar"
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
      Left            =   2070
      Picture         =   "frmDistribucionUtilidades.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8580
      Width           =   1200
   End
   Begin VB.Frame fraCarga 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   13935
      Begin VB.Frame fraParametros 
         Caption         =   "Parametros de Distribución"
         Height          =   765
         Left            =   360
         TabIndex        =   10
         Top             =   2250
         Width           =   13185
         Begin VB.CommandButton cmdCalcularValorCuota 
            Caption         =   "Calcular Valor Cuota"
            Height          =   375
            Left            =   9660
            TabIndex        =   13
            Top             =   240
            Width           =   1995
         End
         Begin TAMControls.TAMTextBox txtValorCuotaActualizado 
            Height          =   315
            Left            =   5850
            TabIndex        =   11
            Top             =   300
            Width           =   1785
            _ExtentX        =   3149
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
            Container       =   "frmDistribucionUtilidades.frx":054B
            Text            =   "0.00000000"
            Decimales       =   8
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   8
            MaximoValor     =   999999999
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor Cuota Actualizado"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   3900
            TabIndex        =   12
            Top             =   330
            Width           =   1695
         End
      End
      Begin VB.Frame frmCarga 
         Caption         =   "Datos para Distribución"
         Height          =   1725
         Left            =   330
         TabIndex        =   1
         Top             =   390
         Width           =   13215
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   600
            Width           =   10005
         End
         Begin MSComCtl2.DTPicker dtpFechaRegistro 
            Height          =   345
            Left            =   1650
            TabIndex        =   3
            Top             =   1050
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   175898625
            CurrentDate     =   38790
         End
         Begin TAMControls.TAMTextBox txtValorCuota 
            Height          =   315
            Left            =   9870
            TabIndex        =   9
            Top             =   1080
            Width           =   1785
            _ExtentX        =   3149
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
            Container       =   "frmDistribucionUtilidades.frx":0567
            Text            =   "0.00000000"
            Decimales       =   8
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   8
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtValorTipoCambio 
            Height          =   315
            Left            =   5070
            TabIndex        =   15
            Top             =   1080
            Width           =   1785
            _ExtentX        =   3149
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
            Container       =   "frmDistribucionUtilidades.frx":0583
            Text            =   "0.00000000"
            Decimales       =   8
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   8
            MaximoValor     =   999999999
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   3990
            TabIndex        =   14
            Top             =   1110
            Width           =   885
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor Cuota Actual"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   8310
            TabIndex        =   8
            Top             =   1140
            Width           =   1320
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fondo"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   5
            Top             =   660
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   4
            Top             =   1110
            Width           =   450
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Height          =   4215
         Left            =   330
         OleObjectBlob   =   "frmDistribucionUtilidades.frx":059F
         TabIndex        =   6
         Top             =   3240
         Width           =   13215
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsultaCO 
         Bindings        =   "frmDistribucionUtilidades.frx":6FC7
         Height          =   645
         Left            =   450
         OleObjectBlob   =   "frmDistribucionUtilidades.frx":6FE1
         TabIndex        =   7
         Top             =   8790
         Width           =   13215
      End
      Begin TAMControls.TAMTextBox txtValorReparto 
         Height          =   315
         Left            =   11370
         TabIndex        =   18
         Top             =   7530
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   16777215
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
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmDistribucionUtilidades.frx":C25B
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
      Begin TAMControls.TAMTextBox txtValorReinversion 
         Height          =   315
         Left            =   11370
         TabIndex        =   19
         Top             =   7860
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   16777215
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
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmDistribucionUtilidades.frx":C277
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
      Begin TAMControls.TAMTextBox txtRetiran 
         Height          =   315
         Left            =   7770
         TabIndex        =   22
         Top             =   7530
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   16777215
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
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmDistribucionUtilidades.frx":C293
         Text            =   "0"
         Estilo          =   4
         CambiarConFoco  =   -1  'True
         ColorEnfoque    =   8454143
         AceptaNegativos =   -1  'True
         Apariencia      =   1
         Borde           =   1
         MaximoValor     =   999999999
      End
      Begin TAMControls.TAMTextBox txtReinvierten 
         Height          =   315
         Left            =   7770
         TabIndex        =   23
         Top             =   7890
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   16777215
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
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmDistribucionUtilidades.frx":C2AF
         Text            =   "0"
         Estilo          =   4
         CambiarConFoco  =   -1  'True
         ColorEnfoque    =   8454143
         AceptaNegativos =   -1  'True
         Apariencia      =   1
         Borde           =   1
         MaximoValor     =   999999999
      End
      Begin VB.Label Label1 
         Caption         =   "Valor Reinversión"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   9810
         TabIndex        =   21
         Top             =   7920
         Width           =   1425
      End
      Begin VB.Label lblValorRepato 
         Caption         =   "Valor Reparto"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   9810
         TabIndex        =   20
         Top             =   7560
         Width           =   1605
      End
      Begin VB.Label lblCantReinvierten 
         Caption         =   "Cant. Reinversión"
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   6180
         TabIndex        =   17
         Top             =   7950
         Width           =   1515
      End
      Begin VB.Label lblCantRetiran 
         Caption         =   "Cant. Reparto"
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   6180
         TabIndex        =   16
         Top             =   7560
         Width           =   1605
      End
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   10770
      TabIndex        =   25
      Top             =   8580
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      Visible0        =   0   'False
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
End
Attribute VB_Name = "frmDistribucionUtilidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Ficha de Mantenimiento de Precios/Tir"
Option Explicit

Dim strEstado                   As String, strSQL               As String
Dim blnSelec                    As Boolean
Dim dblPrecio                   As Double, dblPreProm           As Double
Dim dblTir                      As Double
Dim arrFondo()                  As String  'para el cbofondo.
Dim strCodFondo                 As String
Dim adoRegistro                 As ADODB.Recordset
Dim adoRegistroAux              As ADODB.Recordset
Dim adoClone                    As ADODB.Recordset
Dim adoClone1                   As ADODB.Recordset
Dim conexion                    As ADODB.Connection
Dim tstObservaciones            As New TrueOleDBGrid60.Style
Dim adoField                    As ADODB.Field
Dim adoConsulta                 As ADODB.Recordset
Dim strCodMoneda                As String
Dim dblValorCuotaNominal        As Double
Dim strFechaCorte               As String

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim strSeleccionRegistro    As String

    'If tabPrecio.Tab = 1 Then Exit Sub

    gstrNameRepo = "InstrumentoPrecioTir"
    Select Case Index
        Case 1
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(4)
            ReDim aReportParamFn(2)
            ReDim aReportParamF(2)

            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"

            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)

            aReportParamS(0) = "" 'strCodFile
            aReportParamS(1) = "" 'strCodClaseInstrumento
            aReportParamS(2) = Valor_Caracter
            aReportParamS(3) = Valor_Caracter
            aReportParamS(4) = 1
        Case 2
            strSeleccionRegistro = "{InstrumentoPrecioTir.FechaCotizacion} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal

            If gstrSelFrml <> "0" Then
                Set frmReporte = New frmVisorReporte

                ReDim aReportParamS(4)
                ReDim aReportParamFn(4)
                ReDim aReportParamF(4)

                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "FechaDesde"
                aReportParamFn(2) = "FechaHasta"
                aReportParamFn(3) = "Hora"
                aReportParamFn(4) = "NombreEmpresa"

                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
                aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                aReportParamF(3) = Format(Time(), "hh:mm:ss")
                aReportParamF(4) = gstrNombreEmpresa & Space(1)

                aReportParamS(0) = "" 'strCodFile
                aReportParamS(1) = "" 'strCodClaseInstrumento
                aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
                aReportParamS(4) = 2
            End If
            
            Case 3
            'intNemotecnicoInd = 1
            strSeleccionRegistro = "{InstrumentoPrecioTir.FechaCotizacion} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal

            If gstrSelFrml <> "0" Then
            
            '/* Para validar al cerrar el Rango de Fechas */
            If Mid(gstrSelFrml, 44, 4) = "Fch1" Then
                Exit Sub
            End If
            
            'If intNemotecnicoInd = 1 Then
           '    strNemotecnicoVal = InputBox("Ingrese el Nemotecnico al final, si desea visualizar mas de uno escriba la palabra 'TODOS' ", Me.Caption, UCase("Todos"))
           ' End If
            
                Set frmReporte = New frmVisorReporte

                ReDim aReportParamS(2)
                ReDim aReportParamFn(4)
                ReDim aReportParamF(4)

                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "FechaDesde"
                aReportParamFn(2) = "FechaHasta"
                aReportParamFn(3) = "Hora"
                aReportParamFn(4) = "NombreEmpresa"
                'aReportParamFn(5) = "Fondo"

                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
                aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                aReportParamF(3) = Format(Time(), "hh:mm:ss")
                aReportParamF(4) = gstrNombreEmpresa & Space(1)
                'aReportParamF(4) = gstrNombreEmpresa & Space(1)

'                aReportParamS(0) = "001"   'Mid(strNemotecnicoVal, 1, 3)
'                aReportParamS(1) = gstrCodAdministradora 'Mid(strNemotecnicoVal, 5, 3) 'ponemos la administradora x defecto
                aReportParamS(0) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(1) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
                aReportParamS(2) = "" 'Mid(UCase(strNemotecnicoVal), 1, Len(strNemotecnicoVal)) 'Mid(strNemotecnicoVal, 9, Len(strNemotecnicoVal))
                gstrNameRepo = "InstrumentoPrecioTirDet"
                Else
                    Exit Sub '/*  para validar al dar clic a cancelar en el frmRangoFechas   */
            End If
            
'/* 12:47 p.m. 03/09/2008*/
'/* Se copiaron estas lineas para llamar al nuevo reporte de Grafico de Precio de Mercado */

            Case 4
            'intNemotecnicoInd = 1
            strSeleccionRegistro = "{InstrumentoPrecioTir.FechaCotizacion} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal

            If gstrSelFrml <> "0" Then
            
            '/* Para validar al cerrar el Rango de Fechas */
            If Mid(gstrSelFrml, 44, 4) = "Fch1" Then
                Exit Sub
            End If
            
'            If Len(Trim(arrFondo(cboFondo.ListIndex))) = 0 Then
'                MsgBox "Para mostrar el Reporte tiene que Seleccionar un Fondo", vbExclamation
'                Exit Sub
'            End If
            '''''
            
                Set frmReporte = New frmVisorReporte

                ReDim aReportParamS(4)
                ReDim aReportParamFn(4)
                ReDim aReportParamF(4)

                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "FechaDesde"
                aReportParamFn(2) = "FechaHasta"
                aReportParamFn(3) = "Hora"
                aReportParamFn(4) = "NombreEmpresa"
                'aReportParamFn(5) = "Fondo"

                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
                aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                aReportParamF(3) = Format(Time(), "hh:mm:ss")
                aReportParamF(4) = gstrNombreEmpresa & Space(1)
                'aReportParamF(4) = gstrNombreEmpresa & Space(1)

                aReportParamS(0) = "001" 'Trim(arrFondo(cboFondo.ListIndex)) 'Mid(strNemotecnicoVal, 1, 3)
                aReportParamS(1) = gstrCodAdministradora 'Mid(strNemotecnicoVal, 5, 3) 'ponemos la administradora x defecto
                aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
                aReportParamS(4) = "" 'Mid(UCase(strNemotecnicoVal), 1, Len(strNemotecnicoVal)) 'Mid(strNemotecnicoVal, 9, Len(strNemotecnicoVal))
                gstrNameRepo = "InstrumentoPrecioTirDetGraf"
                Else
                    Exit Sub '/* para validar al dar clic a cancelar en el frmRangoFechas  */
            End If
            
            Case 5
            'intNemotecnicoInd = 1
            strSeleccionRegistro = "{InstrumentoPrecioTir.FechaCotizacion} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal

            If gstrSelFrml <> "0" Then
            
            '/* Para validar al cerrar el Rango de Fechas */
            If Mid(gstrSelFrml, 44, 4) = "Fch1" Then
                Exit Sub
            End If
            
'            If Len(Trim(arrFondo(cboFondo.ListIndex))) = 0 Then
'                MsgBox "Para mostrar el Reporte tiene que Seleccionar un Fondo", vbExclamation
'                Exit Sub
'            End If
            '''''
            
                Set frmReporte = New frmVisorReporte

                ReDim aReportParamS(1)
                ReDim aReportParamFn(5)
                ReDim aReportParamF(5)

                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "FechaDesde"
                aReportParamFn(2) = "FechaHasta"
                aReportParamFn(3) = "Hora"
                aReportParamFn(4) = "NombreEmpresa"
                aReportParamFn(5) = "Fondo"

                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
                aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                aReportParamF(3) = Format(Time(), "hh:mm:ss")
                aReportParamF(4) = gstrNombreEmpresa & Space(1)
                aReportParamF(5) = cboFondo.List(cboFondo.ListIndex)

                aReportParamS(0) = Trim(arrFondo(cboFondo.ListIndex)) 'Trim(arrFondo(cboFondo.ListIndex)) 'Mid(strNemotecnicoVal, 1, 3)
                aReportParamS(1) = gstrCodAdministradora 'Mid(strNemotecnicoVal, 5, 3) 'ponemos la administradora x defecto
                'aReportParamS(3) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                'aReportParamS(4) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))

                
                
                gstrNameRepo = "ParticipeDistribucionHist"
                Else
                    Exit Sub '/* para validar al dar clic a cancelar en el frmRangoFechas  */
            End If
'/* */
            
            
    End Select

    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub

Public Sub Adicionar()

End Sub

Public Sub Cancelar()

'    cmdOpcion.Visible = True
'    With tabPrecio
'        .TabEnabled(0) = True
'        .Tab = 0
'    End With
'    Call Buscar
    
End Sub

Public Sub Eliminar()

End Sub

Public Sub Grabar()

'    Dim strFechaInicio  As String, strFechaFin  As String
'    Dim intRegistro     As Integer
'
'    If strEstado = Reg_Consulta Then Exit Sub
'
'    If strEstado = Reg_Edicion Then
'        'If TodoOK() Then
'            strFechaInicio = Convertyyyymmdd(dtpFechaRegistro.Value)
'            strFechaFin = Convertyyyymmdd(DateAdd("d", 1, dtpFechaRegistro.Value))
'
'            With adoComm
'                .CommandText = "UPDATE InstrumentoPrecioTir SET " & _
'                    "PrecioCierre=" & CDec(txtPrecioCierre.Text) & "," & _
'                    "TirCierre=" & CDec(txtTirCierre.Text) & "," & _
'                    "UsuarioEdicion='" & gstrLogin & "' " & _
'                    "WHERE CodTitulo='" & Trim(tdgConsulta.Columns(2).Value) & "' AND " & _
'                    "(FechaCotizacion>='" & strFechaInicio & "' AND FechaCotizacion<'" & strFechaFin & "')"
'                adoConn.Execute .CommandText, intRegistro
'
'                If intRegistro = 0 Then
'                    .CommandText = "UPDATE InstrumentoPrecioTir SET " & _
'                        "IndUltimoPrecio=''," & _
'                        "UsuarioEdicion='" & gstrLogin & "' " & _
'                        "WHERE CodTitulo='" & Trim(tdgConsulta.Columns(2).Value) & "' AND " & _
'                        "FechaCotizacion = (SELECT MAX(FechaCotizacion) FROM InstrumentoPrecioTir " & _
'                        "                   WHERE CodTitulo='" & Trim(tdgConsulta.Columns(2).Value) & "' AND " & _
'                        "                   FechaCotizacion < '" & strFechaInicio & "')"
'                    adoConn.Execute .CommandText
'
'                    .CommandText = "INSERT INTO InstrumentoPrecioTir " & _
'                     "(CodTitulo, FechaCotizacion, Nemotecnico," & _
'                     "CodFile, CodDetalleFile, CodAnalitica," & _
'                     "PrecioCierre, TirCierre, PrecioPromedio," & _
'                     "IndUltimoPrecio, UsuarioEdicion) " & _
'                     " VALUES ('" & _
'                     Trim(tdgConsulta.Columns(2).Value) & "','" & strFechaInicio & "','" & _
'                     Trim(lblNemotecnico.Caption) & "','" & strCodFile & "','" & _
'                     Trim(tdgConsulta.Columns(7).Value) & "','" & Trim(tdgConsulta.Columns(6).Value) & "'," & _
'                     CDec(txtPrecioCierre.Text) & "," & CDec(txtTirCierre.Text) & "," & _
'                     CDec(txtPrecioCierre.Text) & ",'X','" & gstrLogin & "')"
'                    adoConn.Execute .CommandText
'                End If
'
'            End With
'
'            Me.MousePointer = vbDefault
'
'            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
'
'            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
'
'            cmdOpcion.Visible = True
'            With tabPrecio
'                .TabEnabled(0) = True
'                .Tab = 0
'            End With
'
'            Call Buscar
'        'End If
'    End If

End Sub

Public Sub Imprimir()

End Sub

Public Sub Salir()

    Unload Me
    
End Sub
'
'Private Sub cboClaseInstrumento_Click()
'
'    strCodClaseInstrumento = Valor_Caracter
'    If cboClaseInstrumento.ListIndex < 0 Then Exit Sub
'
'    strCodClaseInstrumento = Trim(arrClaseInstrumento(cboClaseInstrumento.ListIndex))
'
'    Call Buscar
'
'End Sub


'Private Sub cboTipoInstrumento_Click()
'
'    strCodFile = Valor_Caracter
'    If cboTipoInstrumento.ListIndex < 0 Then Exit Sub
'
'    strCodFile = Trim(arrTipoInstrumento(cboTipoInstrumento.ListIndex))
'
'    '*** Clase de Instrumento ***
'    strSQL = "SELECT CodDetalleFile CODIGO,DescripDetalleFile DESCRIP FROM InversionDetalleFile WHERE CodFile='" & strCodFile & "' and CodDetalleEstructura<>'' ORDER BY DescripDetalleFile"
'    CargarControlLista strSQL, cboClaseInstrumento, arrClaseInstrumento(), Sel_Defecto
'
'    If cboClaseInstrumento.ListCount > 0 Then cboClaseInstrumento.ListIndex = 0
'
'    Call Buscar
'
'End Sub

Private Sub cboFondo_Click()

    Dim adoConsulta As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    'Set adoConsulta = New ADODB.Recordset
    
    With adoComm
        '{*** Fecha Vigente, Moneda ***
        .CommandText = "{ call up_ACSelDatosParametro (23,'" & strCodFondo & "','" & gstrCodAdministradora & "','000') }"
        Set adoConsulta = .Execute
        
        If Not adoConsulta.EOF Then
            gdatFechaActual = CVDate(adoConsulta("FechaCuota"))
            gstrFechaActual = Convertyyyymmdd(gdatFechaActual)
            strFechaCorte = Convertyyyymmdd(DateAdd("d", -1, gdatFechaActual))
            dtpFechaRegistro.Value = gdatFechaActual
                          
            strCodMoneda = adoConsulta("CodMoneda")
                          
            txtValorCuota.Text = CStr(adoConsulta("ValorCuotaInicialReal"))
            dblValorCuotaNominal = adoConsulta("ValorCuotaNominal")
            txtValorCuotaActualizado.Text = CStr(dblValorCuotaNominal)
            
            txtValorTipoCambio.Text = CStr(2.586) 'CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaRegistro.Value, Codigo_Moneda_Local, strCodMoneda))
            
            If CDbl(txtValorTipoCambio.Text) = 0 Then txtValorTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaRegistro.Value), Codigo_Moneda_Local, strCodMoneda))
            
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            
            gstrPeriodoActual = Format(Year(gdatFechaActual), "0000")
            gstrMesActual = Format(Month(gdatFechaActual), "00")
            
            'ACTUALIZA PARAMETROS GLOBALES POR FONDO
            If Not CargarParametrosGlobales() Then Exit Sub

        End If
        adoConsulta.Close: Set adoConsulta = Nothing
    End With
  Call Buscar


End Sub

'Private Sub cmdBuscar_Click()
'
'    gs_FormName = ""
'    frmFileExplorer.Show vbModal
'
'    If Trim(gs_FormName) <> "" Then txtArchivo.Text = gs_FormName
'
'
'End Sub

Private Sub cmdCancelar_Click()

    Call Cancelar

End Sub

'Private Sub cmdCargar_Click()
'
'    Call CargarPrecios_xInterfaz
'
'End Sub

'Private Sub CargarPrecios_xInterfaz()
'
'Dim objExcel As Excel.Application
'Dim xLibro As Excel.Workbook
'Dim Col As Integer, fila As Integer
'Dim precio As Double
'Dim fechaCarga As String
'Dim strNemotecnico, strCodTitulo, strMsgError As String
'Dim blnOpenExcel As Boolean
'
'Dim intColNemonicoElex, intColPrecioElex, intColPrecioAntElex, intColFechaAntElex  As Integer
'Dim lngFilaIniElex  As Long
'
'Dim intColNemonicoBloom, intColPrecioBloom As Integer
'Dim lngFilaIniBloom  As Long
'
'blnOpenExcel = False
'
'On Error GoTo CtrlError
'
'If MsgBox("Desea realizar la carga de precios de mercado de los instrumentos ?.", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
'   Me.Refresh: Exit Sub
'End If
'
'If Dir(Trim(txtArchivo.Text)) = "" Then
'    MsgBox "No se encontró el archivo con los precios de carga. Revise si está indicado correctamente. ", vbCritical
'    Exit Sub
'End If
'
'Screen.MousePointer = vbHourglass
'
'Set objExcel = New Excel.Application
'Set xLibro = objExcel.Workbooks.Open(Trim(txtArchivo.Text))   ' ("c:\precio.xls")
'objExcel.Visible = False
'blnOpenExcel = True
'
''-----------------------------------
''Valores pre-establecidos
''-----------------------------------
'
'' Seteos Elex
'intColNemonicoElex = 2
'intColPrecioElex = 7
'intColPrecioAntElex = 3
'intColFechaAntElex = 4
'lngFilaIniElex = 3
'
'' Seteos Bloomberg
'intColNemonicoBloom = 20
'intColPrecioBloom = 21
'lngFilaIniBloom = 2
'
''-----------------------------------
''Carga de precios de valores locales
''-----------------------------------
'
'With xLibro
'    With .Sheets(1)
'
'        fila = lngFilaIniElex
'        strNemotecnico = Trim(.Cells(fila, intColNemonicoElex))
'
'        'For fila = lngFilaInicial To 300
'        Do While Trim(strNemotecnico) <> ""
'
'            'Obtener el precio de la acción.
'            If Trim(.Cells(fila, intColPrecioElex)) = "------" Then
'                'Si no se encuentra el último valor se tomará el valor anterior.
'                If Trim(.Cells(fila, intColPrecioAntElex)) = "------" Then
'                    GoTo siga
'                Else
'                    precio = CDbl(.Cells(fila, intColPrecioAntElex))
'                    If Trim(.Cells(fila, intColFechaAntElex)) <> "-----" Then
'                    fechaCarga = Convertyyyymmdd(Trim(.Cells(fila, intColFechaAntElex)))
'
'                    End If
'                End If
'            Else
'                precio = CDbl(.Cells(fila, intColPrecioElex))
'                fechaCarga = gstrFechaActual
'            End If
'
'            With adoComm
'                .CommandText = "{ call up_IVActPrecioValores ('" & strNemotecnico & "'," & precio & ",'" & _
'                                  fechaCarga & "','" & gstrLogin & "' ) }"
'                adoConn.Execute .CommandText
'
'            End With
'
'siga:
'
'            fila = fila + 1
'            strNemotecnico = Trim(.Cells(fila, intColNemonicoElex))
'
'        Loop
'
'    End With
'
'End With
'
''---------------------------------------------------------------
''Ahora con el mismo archivo se cargan los precios del extranjero
''---------------------------------------------------------------
'
'With xLibro
'    With .Sheets(1)
'
'        fila = lngFilaIniBloom
'        strNemotecnico = Trim(.Cells(fila, intColNemonicoBloom))
'
'
'        Do While Trim(strNemotecnico) <> ""
'
'            'Obtener el precio de la acción.
'            If IsNumeric(.Cells(fila, intColPrecioBloom)) = True Then
'               precio = CDbl((.Cells(fila, intColPrecioBloom)))
'            Else
'                GoTo siga2
'            End If
'
'            With adoComm
'                .CommandText = "{ call up_IVActPrecioValores ('" & strNemotecnico & "'," & precio & ",'" & _
'                                  gstrFechaActual & "','" & gstrLogin & "' ) }"

'Private Sub cmdCargaPreliminar_Click()
'
'    On Error GoTo CtrlError
'
'    'Archivo Excel de Precios
'    Dim strPathFile As String
'
'    strPathFile = Trim(txtArchivo.Text)
'    Dim rango, Hoja As String
'    Dim corr As Long
'
'    rango = ""
'    Hoja = "Operaciones"
'
'    Set adoRegistro = New ADODB.Recordset
'
'    'Manejo de Excel
'
'    Set conexion = New ADODB.Connection
'
'    conexion.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
'                  "Data Source=" & strPathFile & _
'                  ";Extended Properties=Excel 12.0;"
'
'    If rango <> ":" Then
'       Hoja = Hoja & "$" & rango
'    End If
'
'    strSQL = "SELECT FechaOperacion, FechaLiquidacion, TipoOperacion, " & _
'              "Nemotecnico, Moneda, Cantidad, Precio, Subtotal, ComisionSAB, ComisionBVL, " & _
'              "ComisionConasev, ComisionCavali, FondoLiquidacion, FondoGarantia, IGV, " & _
'              "TotalOperacion, Broker, 'X' AS IndRegistroOK FROM [" & Hoja & "] WHERE Nemotecnico <> ''"
'
''        If rs.EOF = True Then
''            MsgBox "El archivo tiene inconsistencias; no se puede cargar al sistema", vbExclamation
''            rs.Close: conexion.Close
''            Exit Sub
''        End If
''
''        rs.Close
'
'    ' Mostramos los datos en el datagrid
'    Dim i As Integer: i = 0
'    corr = 0
'
'    Call ConfiguraRecordsetAuxiliar
'
'    With adoRegistro
'    'With adoMovimiento
'        .ActiveConnection = conexion
'        .CursorLocation = adUseClient
'        .CursorType = adOpenStatic
'        .LockType = adLockBatchOptimistic
'        .Open strSQL
'
'        If .RecordCount > 0 Then
'            .MoveFirst
'            Do While Not .EOF
'                adoRegistroAux.AddNew
'                For Each adoField In adoRegistroAux.Fields
'                    adoRegistroAux.Fields(adoField.Name) = adoRegistro.Fields(adoField.Name)
'                Next
'                adoRegistroAux.Update
'                adoRegistro.MoveNext
'                'adoMovimiento.MoveNext
'            Loop
'            adoRegistroAux.MoveFirst
'        End If
'
'    End With
'
'    Set adoClone = adoRegistroAux.Clone
'
'    tdgConsulta.DataSource = adoRegistroAux
'
'    tdgConsulta.Refresh
'
'    Me.MousePointer = vbDefault
'
'    cmdCargar.Enabled = False
'
'    If adoRegistroAux.RecordCount = 0 Then
'        MsgBox "No existen registros para cargar!", vbExclamation ''& cmdCommand.CommandText
'        cmdValidarCarga.Enabled = False
'        strEstado = Reg_Defecto 'estado inicial: no hay carga
'        Exit Sub
'    Else
'        cmdValidarCarga.Enabled = True
'        strEstado = Reg_Consulta 'hay carga por validar
'    End If
'
'    MsgBox Mensaje_Carga_Exitosa, vbExclamation
'
'    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
'
'    Exit Sub
'
'CtrlError:
'    Me.MousePointer = vbDefault
'
'    MsgBox "Error al Leer el Archivo de Operaciones!", vbCritical, Me.Caption
'
'
'End Sub

Private Sub cmdCalcularValorCuota_Click()

    Dim adoRegistro                     As New ADODB.Recordset
    Dim dblValorCapital                 As Double
    Dim dblValorUtilidadRepartida       As Double
    Dim dblValorUtilidadReinvertida     As Double
    Dim dblValorUtilidadNoDistribuida   As Double
    Dim dblValorTotalResultados         As Double
    Dim dblValorCuotaActualizado        As Double
    Dim dblValorNuevoCapital            As Double
    Dim dblValorPatrimonioInicial       As Double
    Dim dblBookmark                     As Double
    
    
    If adoRegistroAux.RecordCount = 0 Then
        MsgBox "No existen registros para validar!", vbExclamation ''& cmdCommand.CommandText
        Exit Sub
    End If
    
    dblValorCapital = 0
    dblValorUtilidadRepartida = 0
    dblValorUtilidadReinvertida = 0
    dblValorUtilidadNoDistribuida = 0
    dblValorTotalResultados = 0
    dblValorPatrimonioInicial = 0
    
'    adoComm.CommandText = "SELECT SUM(SaldoInicialContable + SaldoParcialContable) AS ValorResultados FROM PartidaContableSaldos " & _
'                          "WHERE " & _
'                          "CodFondo = '" & strCodFondo & "' AND " & _
'                          "CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
'                          "FechaSaldo = '" & gstrFechaActual & "' AND " & _
'                          "CodCuenta LIKE '59%' AND " & _
'                          "CodMonedaContable = '" & Codigo_Moneda_Local & "'"

    adoComm.CommandText = "SELECT dbo.uf_ACObtenerResultadosEjercicio('" & strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaCorte & "') AS 'ValorResultados'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        dblValorTotalResultados = adoRegistro.Fields("ValorResultados") * -1
    End If
    adoRegistro.Close: Set adoRegistro = Nothing
    
    dblBookmark = adoRegistroAux.Bookmark

    adoRegistroAux.MoveFirst
    
    Do While Not adoRegistroAux.EOF

        dblValorCapital = dblValorCapital + Round((adoRegistroAux.Fields("CantCuotas") * adoRegistroAux.Fields("ValorCuota")), 2)
        
        If Trim(adoRegistroAux.Fields("IndReinvierte")) = Valor_Indicador Then
            dblValorUtilidadReinvertida = dblValorUtilidadReinvertida + adoRegistroAux.Fields("ValorUtilidad")
        End If
        
        If Trim(adoRegistroAux.Fields("IndReinvierte")) = Valor_Caracter Then
            dblValorUtilidadRepartida = dblValorUtilidadRepartida + adoRegistroAux.Fields("ValorUtilidad")
        End If

        adoRegistroAux.MoveNext
    Loop
            
    adoRegistroAux.Bookmark = dblBookmark
        
'    If strCodMoneda <> Codigo_Moneda_Local Then
'        dblValorCapital = Round(dblValorCapital * CDbl(txtValorTipoCambio.Text), 2)
'        dblValorUtilidadRepartida = Round(dblValorUtilidadRepartida * CDbl(txtValorTipoCambio.Text), 2)
'        dblValorUtilidadReinvertida = Round(dblValorUtilidadReinvertida * CDbl(txtValorTipoCambio.Text), 2)
'    End If
        
    'SE ASUME QUE SE DISTRIBUYE TODA LA UTILIDAD
    dblValorUtilidadNoDistribuida = 0 'dblValorTotalResultados - (dblValorUtilidadRepartida + dblValorUtilidadReinvertida)
        
    dblValorNuevoCapital = dblValorCapital - dblValorUtilidadRepartida + dblValorUtilidadReinvertida
        
    dblValorCuotaActualizado = Round(txtValorCuota.Value - ((dblValorUtilidadRepartida + dblValorUtilidadReinvertida) / dblValorTotalResultados) * (txtValorCuota.Value - dblValorCuotaNominal), 3)
    
    'Round(((dblValorNuevoCapital + dblValorUtilidadNoDistribuida) / dblValorPatrimonioInicial) * dblValorCuotaNominal, 5)

    txtValorCuotaActualizado.Text = CStr(dblValorCuotaActualizado)

End Sub

'                adoConn.Execute .CommandText
'
'            End With
'
'
'siga2:
'
'            fila = fila + 1
'            strNemotecnico = Trim(.Cells(fila, intColNemonicoBloom))
'
'        Loop
'
'    End With
'End With
'
''Cerrando el archivo excel
'xLibro.Close True
'Set xLibro = Nothing
'Set objExcel = Nothing
'
'Screen.MousePointer = vbNormal
'MsgBox "Finalizó exitosamente la carga de precios de mercado.", vbExclamation
'
'Call Buscar
'
'Exit Sub
'
'CtrlError:
'    If blnOpenExcel = True Then
'        xLibro.Close True
'        Set xLibro = Nothing
'        Set objExcel = Nothing
'    End If
'
'    Me.MousePointer = vbDefault
'    strMsgError = "Error " & Str(err.Number) & vbNewLine
'    strMsgError = strMsgError & err.Description
'    MsgBox strMsgError, vbCritical, "Error"
'
'
'End Sub
Private Sub cmdCargar_Click()
        
    Dim strFechaCarga As String

    Dim objSolicitudDistribucionUtilidadXML    As DOMDocument60
    Dim strSolicitudDistribucionUtilidadXML    As String
    Dim strMsgError                     As String

    If Not TodoOK() Then Exit Sub
    
    If MsgBox("Desea Proceder con la Carga de Operaciones del dia " & dtpFechaRegistro.Value & " ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    strFechaCarga = Convertyyyymmdd(dtpFechaRegistro.Value)

    Call XMLADORecordset(objSolicitudDistribucionUtilidadXML, "ParticipeSolicitud", "Solicitud", adoRegistroAux, strMsgError)
    strSolicitudDistribucionUtilidadXML = objSolicitudDistribucionUtilidadXML.xml

    'txtValorCuotaActualizado.Text = CStr(100#)

    With adoComm
        .CommandText = "{ call up_GNGenSolicitudDistribucionUtilidad ('" & strCodFondo & "','" & gstrCodAdministradora & "'," & _
                        "'" & strFechaCarga & "'," & CDbl(txtValorCuota.Value) & "," & CDbl(txtValorCuotaActualizado.Value) & ",'" & strSolicitudDistribucionUtilidadXML & "') }"
        adoConn.Execute .CommandText
    End With

    cmdCargar.Enabled = False
    
    MsgBox Mensaje_Carga_Exitosa, vbExclamation
    
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
    Exit Sub

CtrlError:
    Me.MousePointer = vbDefault

    MsgBox "Error al Leer El Archivo, Verifique que la estructura sea la correcta. "

End Sub

Private Function TodoOK()

    TodoOK = False

    If adoRegistroAux.RecordCount = 0 Then
        MsgBox "No existen registros para cargar!", vbExclamation ''& cmdCommand.CommandText
        Exit Function
    End If

    TodoOK = True


End Function




Private Sub Command1_Click()

    With tdgConsulta.PrintInfo
        ' Set the page header
        .PageHeaderFont.Italic = True
        .PageHeader = "Composers table"
        
        ' Column headers will be on every page
        .RepeatColumnHeaders = True
        
        ' Display page numbers (centered)
        .PageFooter = "\tPage: \p"
        ' Invoke Print Preview
        .PrintPreview
    End With

End Sub

'Private Sub cmdValidarCarga_Click()
'
'    Dim strCriteria As String
'    Dim adoField As ADODB.Field
'    Dim adoRegistro As New ADODB.Recordset
'    Dim dblBookmark As Double
'
'    If adoRegistroAux.RecordCount = 0 Then
'        MsgBox "No existen registros para validar!", vbExclamation ''& cmdCommand.CommandText
'        Exit Sub
'    End If
'
'    Dim strTipoValidacion As String
'
'    dblBookmark = adoRegistroAux.Bookmark
'
'    adoRegistroAux.MoveFirst
'
'    Do While Not adoRegistroAux.EOF
'
'        For Each adoField In adoRegistroAux.Fields
'
'            'TipoOperacion,Nemotecnico,Moneda,Broker
'            If adoField.Name = "TipoOperacion" Or adoField.Name = "Nemotecnico" Or adoField.Name = "Moneda" Or adoField.Name = "Broker" Then
'
'                If adoField.Name = "TipoOperacion" Then 'TipoOperacion
'                    strTipoValidacion = "01"
'                ElseIf adoField.Name = "Nemotecnico" Then  'Nemotecnico
'                    strTipoValidacion = "02"
'                ElseIf adoField.Name = "Moneda" Then 'Moneda
'                    strTipoValidacion = "03"
'                ElseIf adoField.Name = "Broker" Then  'Broker
'                    strTipoValidacion = "04"
'                End If
'
'                adoComm.CommandText = "SELECT dbo.uf_IVValidaDatoCargaOperacion('" & strTipoValidacion & "','" & adoRegistroAux.Fields(adoField.Name).Value & "') AS 'ValidaDato'"
'                Set adoRegistro = adoComm.Execute
'
'                If Not adoRegistro.EOF Then
'                    If Not adoRegistro("ValidaDato") Then
'                        adoRegistroAux("IndRegistroOK").Value = Valor_Caracter
'                        adoRegistroAux.Fields(adoField.Name).Value = "¿?" & adoRegistroAux.Fields(adoField.Name).Value
'                    End If
'                End If
'                adoRegistro.Close: Set adoRegistro = Nothing
'            End If
'
'        Next
'
'        adoRegistroAux.MoveNext
'    Loop
'
'    adoRegistroAux.Bookmark = dblBookmark
'
'    tdgConsulta.FetchRowStyle = True
'
'    tdgConsulta.Refresh
'
'    cmdCargar.Enabled = True
'    cmdValidarCarga.Enabled = False
'
'    MsgBox "Validación Finalizada!", vbExclamation, Me.Caption
'
'End Sub

Private Sub Form_Activate()

    frmMainMdi.stbMdi.Panels(3).Text = Me.Caption
    Call CargarReportes
   
End Sub

Private Sub Form_Deactivate()

    Call OcultarReportes
    
End Sub

Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call CargarReportes
    Call Buscar
    Call DarFormato
    
    CentrarForm Me
      
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)

End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
            
End Sub
Private Sub CargarReportes()

'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Vista Activa"
'   'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
'   'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Por Rango de Fechas"
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Text = "Por Rango de Fechas - Detallado"
'
'    '/* 12:37 p.m. 03/09/2008                                   */
'    '/* Se agrego estas lineas para llamar al nuevo reporte     */
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Text = "Por Rango de Fechas - Gráfico"
'    '/**/
'
     frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo5").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo5").Text = "Distribución de Utilidades - Histórico"
End Sub
Private Sub CargarListas()
    
    Dim strSQL  As String
    
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
        
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    'adoComm.CommandText = " SELECT ValorCuotaInicialReal FROM FondoSerieValorCuota WHERE IndAbierto='X' AND CodFondo= '" & gstrCodAdministradora & "'"
      '  Set adoRegistro = adoComm.Execute
        'Label1.Caption = adoRegistro("ValorCuotaInicialReal")
        

End Sub
Private Sub InicializarValores()
    
    strEstado = Reg_Defecto
    'tabPrecio.Tab = 0

    dtpFechaRegistro.Value = gdatFechaActual
    
    Set tstObservaciones = tdgConsulta.Styles.Add("Observaciones")
    'tstObservaciones.BackColor = vbRed
    'tstObservaciones.Font.B = True
    
     ' Before modifying the grid's properties, make sure the grid is
 
    ' Create an additional split.
    Dim S As TrueOleDBGrid60.Split
    Set S = tdgConsulta.Splits.Add(0)
 
    ' Hide all columns in the leftmost split, Splits(0), except for columns 0 and 1.
    Dim c As TrueOleDBGrid60.Column
    Dim Cols As TrueOleDBGrid60.Columns
    Set Cols = tdgConsulta.Splits(0).Columns
    For Each c In Cols
        c.Visible = False
    Next c
    Cols(0).Visible = True
    Cols(1).Visible = True

    
    
    
 
    ' Configure Splits(0) to display exactly two columns, and disable resizing.
    With tdgConsulta.Splits(0)
        .SizeMode = dbgNumberOfColumns
        .Size = 2
        .AllowSizing = False
    End With
 
    ' Usually, if you fix columns 0 and 1 from scrolling  in a split, you will
    ' want to make them invisible in other splits.
    Set Cols = tdgConsulta.Splits(1).Columns
    Cols(0).Visible = False
    Cols(1).Visible = False
 
    ' Turn off the record selectors in Split 1.
    tdgConsulta.Splits(1).RecordSelectors = False
    
    '++REA 20-04-2015 Ocultar Columna reinvertir utilidades
    'tdgConsulta.Columns(11).Visible = False
    '--REA 20-04-2015
    
    Set cmdSalir.FormularioActivo = Me
    'Set cmdAccion.FormularioActivo = Me
    'Set cmdOpcion.FormularioActivo = Me
                
End Sub
Private Sub ConfiguraRecordsetAuxiliar()

    Set adoRegistroAux = New ADODB.Recordset

    With adoRegistroAux
       .CursorLocation = adUseClient
       .Fields.Append "CodParticipe", adChar, 20
       .Fields.Append "DescripParticipe", adVarChar, 200
       .Fields.Append "NumCertificado", adVarChar, 10
       .Fields.Append "CodMoneda", adChar, 2
       .Fields.Append "CodSigno", adChar, 3
       .Fields.Append "FechaIngreso", adDate, 10
       .Fields.Append "CantCuotas", adDecimal, 19
       .Fields.Append "ValorCuota", adDecimal, 19
       .Fields.Append "ValorUtilidadBruta", adDecimal, 19
       .Fields.Append "TasaImptoRenta", adDecimal, 19
       .Fields.Append "ValorImptoRenta", adDecimal, 19
       .Fields.Append "ValorUtilidadNeta", adDecimal, 19
       .Fields.Append "IndReinvierte", adChar, 1
       .Fields.Append "ClaseParticipe", adChar, 2
       .LockType = adLockBatchOptimistic
    End With

    With adoRegistroAux.Fields.Item("CantCuotas")
        .Precision = 19
        .NumericScale = 8
    End With

    With adoRegistroAux.Fields.Item("ValorCuota")
        .Precision = 19
        .NumericScale = 8
    End With

    With adoRegistroAux.Fields.Item("ValorUtilidadBruta")
        .Precision = 19
        .NumericScale = 2
    End With
    
    With adoRegistroAux.Fields.Item("ValorUtilidadNeta")
        .Precision = 19
        .NumericScale = 2
    End With
    
    With adoRegistroAux.Fields.Item("TasaImptoRenta")
        .Precision = 19
        .NumericScale = 8
    End With
    
    With adoRegistroAux.Fields.Item("ValorImptoRenta")
        .Precision = 19
        .NumericScale = 2
    End With
    
    adoRegistroAux.Open

End Sub
Public Sub Buscar()
            
    Dim strSQL As String
    Dim intTotalReinvierten As Long
    Dim intTotalRetiran As Long
    Dim dblTotalValorReinvierten As Double
    Dim dblTotalValorRetiran As Double
    
    Set adoRegistro = New ADODB.Recordset
    
    Call ConfiguraRecordsetAuxiliar
    
    intTotalReinvierten = 0
    intTotalRetiran = 0
    dblTotalValorReinvierten = 0
    dblTotalValorRetiran = 0
    
    strEstado = Reg_Defecto

    strSQL = "{ call up_GNLstParticipes('" & strCodFondo & "','" & gstrCodAdministradora & "','" & Convertyyyymmdd(dtpFechaRegistro.Value) & "'," & txtValorCuota.Value & "," & txtValorCuotaActualizado.Value & ")}"


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
                    If adoField.Name = "IndReinvierte" Then
                        If adoRegistro.Fields(adoField.Name) = "X" Then
                            intTotalReinvierten = intTotalReinvierten + 1
                            dblTotalValorReinvierten = dblTotalValorReinvierten + adoRegistro.Fields("ValorUtilidadNeta")
                        Else
                            intTotalRetiran = intTotalRetiran + 1
                            dblTotalValorRetiran = dblTotalValorRetiran + adoRegistro.Fields("ValorUtilidadNeta")
                        End If
                    End If
                    
                Next
                adoRegistroAux.Update
                adoRegistro.MoveNext
                'adoMovimiento.MoveNext
            Loop
            adoRegistroAux.MoveFirst
        End If
        
    End With
    
    tdgConsulta.DataSource = adoRegistroAux
    
    txtReinvierten.Text = CStr(intTotalReinvierten)
    txtRetiran.Text = CStr(intTotalRetiran)
    
    txtValorReinversion.Text = CStr(dblTotalValorReinvierten)
    txtValorReparto.Text = CStr(dblTotalValorRetiran)
    
    tdgConsulta.Refresh
    
    If adoRegistroAux.RecordCount > 0 Then strEstado = Reg_Consulta
    
End Sub

Public Sub Accion(nAccion As ButtonAction)

    Select Case nAccion

        Case vNew
            Call Adicionar
        Case vQuery
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

Public Sub Modificar()

'    If strEstado = Reg_Consulta Then
'        strEstado = Reg_Edicion
'        LlenarFormulario strEstado
'        cmdOpcion.Visible = False
'        With tabPrecio
'            .TabEnabled(0) = False
'            .Tab = 1
'        End With
'    End If

End Sub

'Private Sub LlenarFormulario(strModo As String)
'
'    Dim adoRegistro     As ADODB.Recordset
'
'    Select Case strModo
'        Case Reg_Edicion
'            lblDescripInstrumento.Caption = Trim(cboTipoInstrumento.Text)
'            lblFechaRegistro.Caption = CStr(dtpFechaRegistro.Value)
'            lblNemotecnico.Caption = CStr(tdgConsulta.Columns(1))
'            lblPrecioAnterior.Caption = CStr(tdgConsulta.Columns(3))
'            lblTirAnterior.Caption = CStr(tdgConsulta.Columns(4))
'            If Trim(tdgConsulta.Columns(0).Value) = Valor_Caracter Then
'                txtPrecioCierre.Text = "0"
'                txtTirCierre.Text = "0"
'            Else
'                If CVDate(tdgConsulta.Columns(0).Value) < dtpFechaRegistro.Value Then
'                    txtPrecioCierre.Text = "0"
'                    txtTirCierre.Text = "0"
'                Else
'                    txtPrecioCierre.Text = CStr(tdgConsulta.Columns(3))
'                    txtTirCierre.Text = CStr(tdgConsulta.Columns(4))
'                End If
'            End If
'
'            Set adoRegistro = New ADODB.Recordset
'
'            adoComm.CommandText = "SELECT IndPrecio,IndTir FROM InversionFile WHERE CodFile='" & strCodFile & "'"
'            Set adoRegistro = adoComm.Execute
'
'            If Not adoRegistro.EOF Then
'                txtPrecioCierre.Enabled = True
'                If Trim(adoRegistro("IndPrecio")) = Valor_Caracter Then txtPrecioCierre.Enabled = False
'                txtTirCierre.Enabled = True
'                If Trim(adoRegistro("IndTir")) = Valor_Caracter Then txtTirCierre.Enabled = False
'            End If
'            adoRegistro.Close: Set adoRegistro = Nothing
'
'    End Select
'
'End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set frmDistribucionUtilidades = Nothing
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub

Private Sub tdgConsulta_AfterColUpdate(ByVal ColIndex As Integer)
    
    Dim intTotalReinvierten As Long
    Dim intTotalRetiran As Long
    Dim dblTotalValorReinvierten As Double
    Dim dblTotalValorRetiran As Double
    
    Dim intPosicion As Long
    
    intTotalReinvierten = 0
    intTotalRetiran = 0
    dblTotalValorReinvierten = 0
    dblTotalValorRetiran = 0

    intPosicion = adoRegistroAux.Bookmark
    adoRegistroAux.MoveFirst
    
    Do While Not adoRegistroAux.EOF
        If adoRegistroAux.Fields("IndReinvierte") = Valor_Indicador Then
            intTotalReinvierten = intTotalReinvierten + 1
            dblTotalValorReinvierten = dblTotalValorReinvierten + adoRegistroAux.Fields("ValorUtilidadNeta")
        Else
            intTotalRetiran = intTotalRetiran + 1
            dblTotalValorRetiran = dblTotalValorRetiran + adoRegistroAux.Fields("ValorUtilidadNeta")
        End If
        
        adoRegistroAux.MoveNext
    Loop
    
    adoRegistroAux.AbsolutePosition = intPosicion
    txtReinvierten.Text = CStr(intTotalReinvierten)
    txtRetiran.Text = CStr(intTotalRetiran)
    
    txtValorReinversion.Text = CStr(dblTotalValorReinvierten)
    txtValorReparto.Text = CStr(dblTotalValorRetiran)
    
    

End Sub


Private Sub tdgConsulta_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid60.StyleDisp)

'    Dim strTipoValidacion As String
'    Dim adoRegistro As New ADODB.Recordset
'
'    adoClone.Bookmark = Bookmark
'
'    Select Case Col
'        Case 2, 3, 4, 9 'TipoOperacion,Nemotecnico,Moneda,Broker
'
'            If Col = 2 Then 'TipoOperacion
'                strTipoValidacion = "01"
'            ElseIf Col = 3 Then 'Nemotecnico
'                strTipoValidacion = "02"
'            ElseIf Col = 4 Then 'Moneda
'                strTipoValidacion = "03"
'            ElseIf Col = 9 Then 'Broker
'                strTipoValidacion = "04"
'            End If
'
'            adoComm.CommandText = "SELECT dbo.uf_IVValidaDatoCargaOperacion('" & strTipoValidacion & "','" & tdgConsulta.Columns(Col).CellText(Bookmark) & "') AS 'ValidaDato'"
'            Set adoRegistro = adoComm.Execute
'
'            If Not adoRegistro.EOF Then
'                If Not adoRegistro("ValidaDato") Then
'                    'tdgConsulta.Columns("IndRegistroOK").CellText(Bookmark) = Valor_Caracter
'                    adoClone("IndRegistroOK").Value = Valor_Caracter
'                    CellStyle.Font.B = True
'                End If
'            End If
'            adoRegistro.Close: Set adoRegistro = Nothing
'
'    End Select


End Sub

Private Sub tdgConsulta_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
'
'    adoClone.Bookmark = Bookmark
'
'    If Trim(adoClone("IndRegistroOK").Value) = Valor_Caracter Then
'        RowStyle.ForeColor = vbWhite
'        RowStyle.BackColor = vbRed
'    End If

End Sub


Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If tdgConsulta.Columns(ColIndex).DataField = "CantCuotas" Then
        Call DarFormatoValor(Value, Decimales_CantCuota)
    End If
    
    If tdgConsulta.Columns(ColIndex).DataField = "ValorCuota" Then
        Call DarFormatoValor(Value, Decimales_ValorCuota)
    End If
    
    If tdgConsulta.Columns(ColIndex).DataField = "ValorUtilidadBruta" Or _
       tdgConsulta.Columns(ColIndex).DataField = "ValorImptoRenta" Or _
       tdgConsulta.Columns(ColIndex).DataField = "ValorUtilidadNeta" Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
    If tdgConsulta.Columns(ColIndex).DataField = "TasaImptoRenta" Then
        Call DarFormatoValor(Value, Decimales_Tasa)
    End If
    
    
End Sub


Private Sub tdgConsulta_HeadClick(ByVal ColIndex As Integer)

    Static numColindex As Integer

    tdgConsulta.Splits(1).Columns(numColindex).HeadingStyle.ForegroundPicture = Null

    Call OrdenarDBGrid(ColIndex, adoRegistroAux, tdgConsulta)
    
    numColindex = ColIndex

End Sub


Private Sub tdgConsulta_UnboundColumnFetch(Bookmark As Variant, ByVal Col As Integer, Value As Variant)

'    adoClone.Bookmark = Bookmark
'
'    If Col = 7 Then
'        Value = adoClone("ComisionSAB") + adoClone("ComisionBVL") + _
'                adoClone("ComisionConasev") + adoClone("ComisionCavali") + _
'                adoClone("FondoLiquidacion") + adoClone("FondoGarantia") + adoClone("IGV")
'    End If
    

End Sub

