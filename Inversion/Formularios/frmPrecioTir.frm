VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmPrecioTir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Precio / Tir de Mercado"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   10785
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   600
      TabIndex        =   34
      Top             =   6120
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1296
      Buttons         =   2
      Caption0        =   "&Modificar"
      Tag0            =   "3"
      ToolTipText0    =   "&Modificar"
      Caption1        =   "&Buscar"
      Tag1            =   "5"
      ToolTipText1    =   "Buscar"
      UserControlWidth=   2700
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   9000
      TabIndex        =   33
      Top             =   6120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TabDlg.SSTab tabPrecio 
      Height          =   5955
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   10504
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmPrecioTir.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraPrecioTir"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdgConsulta"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmPrecioTir.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraDatos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdAccion2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Carga de Precios"
      TabPicture(2)   =   "frmPrecioTir.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Text1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "frmCarga"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion2 
         Height          =   735
         Left            =   -68280
         TabIndex        =   32
         Top             =   5040
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
      Begin VB.Frame frmCarga 
         Caption         =   " Cargar desde ... "
         Height          =   2295
         Left            =   360
         TabIndex        =   26
         Top             =   1350
         Width           =   9855
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8850
            TabIndex        =   31
            Top             =   630
            Width           =   315
         End
         Begin VB.TextBox txtArchivo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "C:\Precio.xls"
            Top             =   630
            Width           =   7665
         End
         Begin VB.CommandButton cmdCargar 
            Caption         =   "Cargar &precios"
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
            Left            =   5910
            Picture         =   "frmPrecioTir.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   1320
            Width           =   1545
         End
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
            Left            =   7710
            Picture         =   "frmPrecioTir.frx":059F
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1320
            Width           =   1545
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Archivo "
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   360
            TabIndex        =   30
            Top             =   660
            Width           =   585
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   360
         TabIndex        =   25
         Text            =   "Precios de Mercado. Fuentes:  Elex / Bloomberg"
         Top             =   720
         Width           =   9855
      End
      Begin VB.Frame fraDatos 
         Caption         =   "Datos"
         Height          =   4335
         Left            =   -74670
         TabIndex        =   10
         Top             =   600
         Width           =   9885
         Begin VB.TextBox txtTirCierre 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6360
            TabIndex        =   19
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox txtPrecioCierre 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2520
            TabIndex        =   18
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label lblTirAnterior 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6360
            TabIndex        =   24
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tir Anterior"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   4680
            TabIndex        =   23
            Top             =   1700
            Width           =   765
         End
         Begin VB.Label lblDescripInstrumento 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2520
            TabIndex        =   22
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Instrumento"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   720
            TabIndex        =   21
            Top             =   740
            Width           =   825
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tir de Cierre"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   4680
            TabIndex        =   20
            Top             =   2175
            Width           =   855
         End
         Begin VB.Label lblPrecioAnterior 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2520
            TabIndex        =   17
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label lblNemotecnico 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2520
            TabIndex        =   16
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label lblFechaRegistro 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6360
            TabIndex        =   15
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio de Cierre"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   720
            TabIndex        =   14
            Top             =   2180
            Width           =   1125
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio Anterior"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   720
            TabIndex        =   13
            Top             =   1700
            Width           =   1035
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nemotécnico"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   720
            TabIndex        =   12
            Top             =   1220
            Width           =   945
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   4680
            TabIndex        =   11
            Top             =   740
            Width           =   450
         End
      End
      Begin TAMControls.ucBotonEdicion cmdAccion 
         Height          =   390
         Left            =   -68100
         TabIndex        =   9
         Top             =   4320
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   688
         Buttons         =   2
         Caption0        =   "&Guardar"
         Tag0            =   "2"
         ToolTipText0    =   "Guardar"
         Caption1        =   "&Cancelar"
         Tag1            =   "8"
         ToolTipText1    =   "Cancelar"
         UserControlHeight=   390
         UserControlWidth=   2700
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmPrecioTir.frx":0B01
         Height          =   3165
         Left            =   -74670
         OleObjectBlob   =   "frmPrecioTir.frx":0B1B
         TabIndex        =   8
         Top             =   2100
         Width           =   9885
      End
      Begin VB.Frame fraPrecioTir 
         Caption         =   "Criterios de Búsqueda"
         Height          =   1440
         Left            =   -74670
         TabIndex        =   1
         Top             =   600
         Width           =   9885
         Begin VB.ComboBox cboTipoInstrumento 
            Height          =   315
            Left            =   1755
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   450
            Width           =   3615
         End
         Begin VB.ComboBox cboClaseInstrumento 
            Height          =   315
            Left            =   1755
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   870
            Width           =   3615
         End
         Begin MSComCtl2.DTPicker dtpFechaRegistro 
            Height          =   345
            Left            =   6690
            TabIndex        =   2
            Top             =   450
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
            Format          =   293404673
            CurrentDate     =   38790
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Instrumento"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   570
            TabIndex        =   7
            Top             =   525
            Width           =   825
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Clase"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   570
            TabIndex        =   6
            Top             =   930
            Width           =   390
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   5790
            TabIndex        =   5
            Top             =   450
            Width           =   450
         End
      End
   End
End
Attribute VB_Name = "frmPrecioTir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Ficha de Mantenimiento de Precios/Tir"
Option Explicit

Dim arrClaseInstrumento()       As String, arrTipoInstrumento() As String
Dim strCodClaseInstrumento      As String, strCodFile           As String
Dim strEstado                   As String, strSQL               As String
Dim blnSelec                    As Boolean
Dim dblPrecio                   As Double, dblPreProm           As Double
Dim dblTir                      As Double
Dim arrFondo()       As String  'para el cbofondo.
Dim adoConsulta                 As ADODB.Recordset
Dim indSortAsc                  As Boolean, indSortDesc         As Boolean

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim strSeleccionRegistro    As String

    If tabPrecio.Tab = 1 Then Exit Sub

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

            aReportParamS(0) = strCodFile
            aReportParamS(1) = strCodClaseInstrumento
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

                aReportParamS(0) = strCodFile
                aReportParamS(1) = strCodClaseInstrumento
                aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
                aReportParamS(4) = 2
            End If
            
            Case 3
            intNemotecnicoInd = 1
            strSeleccionRegistro = "{InstrumentoPrecioTir.FechaCotizacion} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal

            If gstrSelFrml <> "0" Then
            
            '/* Para validar al cerrar el Rango de Fechas */
            If Mid(gstrSelFrml, 44, 4) = "Fch1" Then
                Exit Sub
            End If
            
            If intNemotecnicoInd = 1 Then
               strNemotecnicoVal = InputBox("Ingrese el Nemotecnico al final, si desea visualizar mas de uno escriba la palabra 'TODOS' ", Me.Caption, UCase("Todos"))
            End If
            
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
                aReportParamS(2) = Mid(UCase(strNemotecnicoVal), 1, Len(strNemotecnicoVal)) 'Mid(strNemotecnicoVal, 9, Len(strNemotecnicoVal))
                gstrNameRepo = "InstrumentoPrecioTirDet"
                Else
                    Exit Sub '/*  para validar al dar clic a cancelar en el frmRangoFechas   */
            End If
            
'/* 12:47 p.m. 03/09/2008*/
'/* Se copiaron estas lineas para llamar al nuevo reporte de Grafico de Precio de Mercado */

            Case 4
            intNemotecnicoInd = 1
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
                aReportParamS(4) = Mid(UCase(strNemotecnicoVal), 1, Len(strNemotecnicoVal)) 'Mid(strNemotecnicoVal, 9, Len(strNemotecnicoVal))
                gstrNameRepo = "InstrumentoPrecioTirDetGraf"
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

    cmdOpcion.Visible = True
    With tabPrecio
        .TabEnabled(0) = True
        .Tab = 0
    End With
    Call Buscar
    
End Sub

Public Sub Eliminar()

End Sub

Public Sub Grabar()

    Dim strFechaInicio  As String, strFechaFin  As String
    Dim intRegistro     As Integer
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    If strEstado = Reg_Edicion Then
        'If TodoOK() Then
            strFechaInicio = Convertyyyymmdd(dtpFechaRegistro.Value)
            strFechaFin = Convertyyyymmdd(DateAdd("d", 1, dtpFechaRegistro.Value))
            
            With adoComm
                .CommandText = "UPDATE InstrumentoPrecioTir SET " & _
                    "PrecioCierre=" & CDec(txtPrecioCierre.Text) & "," & _
                    "TirCierre=" & CDec(txtTirCierre.Text) & "," & _
                    "UsuarioEdicion='" & gstrLogin & "' " & _
                    "WHERE CodTitulo='" & Trim(tdgConsulta.Columns(2).Value) & "' AND " & _
                    "(FechaCotizacion>='" & strFechaInicio & "' AND FechaCotizacion<'" & strFechaFin & "')"
                adoConn.Execute .CommandText, intRegistro
                
                If intRegistro = 0 Then
                    .CommandText = "UPDATE InstrumentoPrecioTir SET " & _
                        "IndUltimoPrecio=''," & _
                        "UsuarioEdicion='" & gstrLogin & "' " & _
                        "WHERE CodTitulo='" & Trim(tdgConsulta.Columns(2).Value) & "' AND " & _
                        "FechaCotizacion = (SELECT MAX(FechaCotizacion) FROM InstrumentoPrecioTir " & _
                        "                   WHERE CodTitulo='" & Trim(tdgConsulta.Columns(2).Value) & "' AND " & _
                        "                   FechaCotizacion < '" & strFechaInicio & "')"
                    adoConn.Execute .CommandText
                
                    .CommandText = "INSERT INTO InstrumentoPrecioTir " & _
                     "(CodTitulo, FechaCotizacion, Nemotecnico," & _
                     "CodFile, CodDetalleFile, CodAnalitica," & _
                     "PrecioCierre, TirCierre, PrecioPromedio," & _
                     "IndUltimoPrecio, UsuarioEdicion) " & _
                     " VALUES ('" & _
                     Trim(tdgConsulta.Columns(2).Value) & "','" & strFechaInicio & "','" & _
                     Trim(lblNemotecnico.Caption) & "','" & strCodFile & "','" & _
                     Trim(tdgConsulta.Columns(7).Value) & "','" & Trim(tdgConsulta.Columns(6).Value) & "'," & _
                     CDec(txtPrecioCierre.Text) & "," & CDec(txtTirCierre.Text) & "," & _
                     CDec(txtPrecioCierre.Text) & ",'X','" & gstrLogin & "')"
                    adoConn.Execute .CommandText
                End If
            
            End With
            
            Me.MousePointer = vbDefault
        
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabPrecio
                .TabEnabled(0) = True
                .Tab = 0
            End With
            
            Call Buscar
        'End If
    End If

End Sub

Public Sub Imprimir()

End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Private Sub cboClaseInstrumento_Click()

    strCodClaseInstrumento = Valor_Caracter
    If cboClaseInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodClaseInstrumento = Trim(arrClaseInstrumento(cboClaseInstrumento.ListIndex))
    
    Call Buscar
    
End Sub


Private Sub cboTipoInstrumento_Click()
                        
    strCodFile = Valor_Caracter
    If cboTipoInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodFile = Trim(arrTipoInstrumento(cboTipoInstrumento.ListIndex))
                
    '*** Clase de Instrumento ***
    strSQL = "SELECT CodDetalleFile CODIGO,DescripDetalleFile DESCRIP FROM InversionDetalleFile WHERE CodFile='" & strCodFile & "' and CodDetalleEstructura<>'' ORDER BY DescripDetalleFile"
    CargarControlLista strSQL, cboClaseInstrumento, arrClaseInstrumento(), Sel_Defecto
         
    If cboClaseInstrumento.ListCount > 0 Then cboClaseInstrumento.ListIndex = 0
                        
    Call Buscar

    If adoConsulta.RecordCount >= 1 Then
        cmdOpcion.Visible = True
    End If
        
End Sub

Private Sub cmdBuscar_Click()

    gs_FormName = ""
    frmFileExplorer.Show vbModal
    
    If Trim(gs_FormName) <> "" Then txtArchivo.Text = gs_FormName


End Sub

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

    'On Error GoTo CtrlError

        'Archivo Excel de Precios
        Dim strPathFile As String
        
        strPathFile = Trim(txtArchivo.Text)
        Dim rango, Hoja As String
        Dim corr As Long
        rango = ""
        Hoja = "Precios"
        
        'Manejo de Excel
        Dim conexion As ADODB.Connection, rs As ADODB.Recordset

        Set conexion = New ADODB.Connection

        conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                      "Data Source=" & strPathFile & _
                      ";Extended Properties=""Excel 8.0;HDR=Yes;"""

        ' Nuevo recordset
        Set rs = New ADODB.Recordset

        With rs
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
        End With

        If rango <> ":" Then
           Hoja = Hoja & "$" & rango
        End If

        rs.Open "SELECT Ticker, Date, Price FROM [" & Hoja & "] WHERE Ticker IS NULL OR Ticker = ''", conexion, , , adCmdText

'        If rs.EOF = True Then
'            MsgBox "El archivo tiene inconsistencias; no se puede cargar al sistema", vbExclamation
'            rs.Close: conexion.Close
'            Exit Sub
'        End If
'
'        rs.Close

        ' Mostramos los datos en el datagrid
        Dim i As Integer: i = 0
        corr = 0
        
        If rs.EOF = False Then
            rs.MoveFirst
            Do While Not rs.EOF
                With adoComm
                    .CommandText = "{ call up_IVActPrecioValores ('" & rs.Fields("Ticker") & "'," & rs.Fields("Price") & ",'" & _
                                      Convertyyyymmdd(rs.Fields("Date")) & "') }"
                    adoConn.Execute .CommandText
                End With
                
                rs.MoveNext
            Loop
        Else
            MsgBox "No records were returned using the query " ''& cmdCommand.CommandText
        End If

        rs.Close
        conexion.Close

        Me.MousePointer = vbDefault
    
        MsgBox Mensaje_Adicion_Exitosa, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        cmdOpcion.Visible = True
        With tabPrecio
            .TabEnabled(0) = True
            .Tab = 0
        End With
        Call Buscar
        
        Exit Sub

CtrlError:
    Me.MousePointer = vbDefault

    MsgBox "Error al Leer El Archivo, Verifique que la estructura sea la correcta. "

End Sub

Private Sub Command1_Click()

End Sub

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
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
         
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
        
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    Dim elemento As Object
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
    For Each elemento In Me.Controls
    
        If TypeOf elemento Is TDBGrid Then
            Call FormatoGrilla(elemento)
        End If
    
    Next
            
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Vista Activa"
   'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
   'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Por Rango de Fechas"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Text = "Por Rango de Fechas - Detallado"
    
    '/* 12:37 p.m. 03/09/2008                                   */
    '/* Se agrego estas lineas para llamar al nuevo reporte     */
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Text = "Por Rango de Fechas - Grafico"
    '/**/
    
End Sub
Private Sub CargarListas()
    
    Dim strSQL  As String
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT CodFile CODIGO,DescripFile DESCRIP FROM InversionFile WHERE IndInstrumento='X' AND (IndPrecio='X' OR IndTir='X') AND IndVigente='X' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumento, arrTipoInstrumento(), Sel_Defecto
    
    If cboTipoInstrumento.ListCount > 0 Then cboTipoInstrumento.ListIndex = 0
    
    'Agregado para el funcionamiento del combo cbofondos
    '*** Fondos ***
    'strSQL = "{ call up_ACSelDatos(8) }"
    'CargarControlLista strSQL, cboFondo, arrFondo(), Sel_Defecto
    
    'If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
                        

End Sub
Private Sub InicializarValores()
    
    strEstado = Reg_Defecto
    tabPrecio.Tab = 0

    dtpFechaRegistro.Value = gdatFechaActual
    
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion2.FormularioActivo = Me
    Set cmdOpcion.FormularioActivo = Me
                
End Sub
Public Sub Buscar()

    Set adoConsulta = New ADODB.Recordset
                                
    strSQL = "{call up_IVLstPrecioTitulo ('" & Convertyyyymmdd(dtpFechaRegistro.Value) & "','" & strCodFile & "','" & strCodClaseInstrumento & "') }"
                                
    strEstado = Reg_Defecto
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With

    tdgConsulta.DataSource = adoConsulta

    If adoConsulta.RecordCount > 0 Then strEstado = Reg_Consulta
        
End Sub

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

Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabPrecio
            .TabEnabled(0) = False
            .Tab = 1
        End With
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro     As ADODB.Recordset
    
    Select Case strModo
        Case Reg_Edicion
            lblDescripInstrumento.Caption = Trim(cboTipoInstrumento.Text)
            lblFechaRegistro.Caption = CStr(dtpFechaRegistro.Value)
            lblNemotecnico.Caption = CStr(tdgConsulta.Columns(1))
            lblPrecioAnterior.Caption = CStr(tdgConsulta.Columns(3))
            lblTirAnterior.Caption = CStr(tdgConsulta.Columns(4))
            If Trim(tdgConsulta.Columns(0).Value) = Valor_Caracter Then
                txtPrecioCierre.Text = "0"
                txtTirCierre.Text = "0"
            Else
                If CVDate(tdgConsulta.Columns(0).Value) < dtpFechaRegistro.Value Then
                    txtPrecioCierre.Text = "0"
                    txtTirCierre.Text = "0"
                Else
                    txtPrecioCierre.Text = CStr(tdgConsulta.Columns(3))
                    txtTirCierre.Text = CStr(tdgConsulta.Columns(4))
                End If
            End If
            
            Set adoRegistro = New ADODB.Recordset
    
            adoComm.CommandText = "SELECT IndPrecio,IndTir FROM InversionFile WHERE CodFile='" & strCodFile & "'"
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                txtPrecioCierre.Enabled = True
                If Trim(adoRegistro("IndPrecio")) = Valor_Caracter Then txtPrecioCierre.Enabled = False
                txtTirCierre.Enabled = True
                If Trim(adoRegistro("IndTir")) = Valor_Caracter Then txtTirCierre.Enabled = False
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
            
    End Select
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set frmPrecioTir = Nothing
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub

Private Sub lblPrecioAnterior_Change()

    Call FormatoMillarEtiqueta(lblPrecioAnterior, Decimales_Precio)
    
End Sub

Private Sub lblTirAnterior_Change()

    Call FormatoMillarEtiqueta(lblTirAnterior, Decimales_Precio)
    
End Sub

Private Sub tabPrecio_Click(PreviousTab As Integer)

    Select Case tabPrecio.Tab
        Case 1
            cmdOpcion.Visible = False
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabPrecio.Tab = 0
        Case 2
            cmdOpcion.Visible = False
    
    End Select
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 3 Then
        Call DarFormatoValor(Value, Decimales_Precio)
    End If
    
    If ColIndex = 4 Then
        Call DarFormatoValor(Value, Decimales_Tasa)
    End If
    
End Sub

Private Sub txtPrecioCierre_Change()

    Call FormatoCajaTexto(txtPrecioCierre, Decimales_Precio)
    
    
End Sub

Private Sub txtPrecioCierre_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtPrecioCierre, Decimales_Precio)
        
End Sub

Private Sub txtPrecioCierre_LostFocus()

    txtTirCierre.Text = "0"
    
End Sub

Private Sub txtTirCierre_Change()

    Call FormatoCajaTexto(txtTirCierre, Decimales_Tasa)
    
End Sub

Private Sub txtTirCierre_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTirCierre, Decimales_Tasa)
    
End Sub

Private Sub txtTirCierre_LostFocus()

    txtPrecioCierre.Text = "0"
    
End Sub

Private Sub tdgConsulta_HeadClick(ByVal ColIndex As Integer)
    
    Dim strColNameTDB  As String
    Static numColindex As Integer
    Static strPrevColumTDB As String
    '** agregar para que no se raye la seleccion de registro con ordenamiento
    strColNameTDB = tdgConsulta.Columns(ColIndex).DataField
    
    If strColNameTDB = strPrevColumTDB Then
        If indSortAsc Then
            indSortAsc = False
            indSortDesc = True
        Else
            indSortAsc = True
            indSortDesc = False
        End If
    Else
        indSortAsc = True
        indSortDesc = False
    End If
    '***

    tdgConsulta.Splits(0).Columns(numColindex).HeadingStyle.ForegroundPicture = Null

    Call OrdenarDBGrid(ColIndex, adoConsulta, tdgConsulta)
    
    numColindex = ColIndex
    
    '****
    strPrevColumTDB = strColNameTDB
    '***
    
End Sub
