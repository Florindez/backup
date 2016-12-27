VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Begin VB.Form frmVisorReporte 
   Caption         =   "Reporte"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   ForeColor       =   &H00000000&
   Icon            =   "frmVisorReporte.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer crvVisor 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      lastProp        =   600
      _cx             =   16960
      _cy             =   12091
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "frmVisorReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents ReportSection As CRAXDRT.Section
Attribute ReportSection.VB_VarHelpID = -1

Private aParamS(), aParamF(), aParamFn()
Public strReportPath As String
Public strIndLogo    As String

Private Sub Form_Activate()

   MousePointer = vbDefault
   
End Sub

Private Sub Form_Initialize()
    
    strIndLogo = "X"
   
End Sub

Private Sub Form_Load()
      
    Dim crxAplicacion As CRAXDRT.Application
    Dim crxReporte As CRAXDRT.Report
        
    Dim crxTables As CRAXDRT.DatabaseTables 'nuevo
    Dim crxTable As CRAXDRT.DatabaseTable 'nuevo
    Dim crxSubreportObject As CRAXDRT.SubreportObject 'nuevo
    Dim crxSubReport As CRAXDRT.Report 'nuevo
    Dim crxSections As CRAXDRT.Sections 'nuevo
    Dim crxSection As CRAXDRT.Section 'nuevo
    Dim intCantSections As Integer
    Dim intCantReportObjects As Integer
    Dim bmp As StdPicture
    
    Dim intAccion   As Integer
    Dim lngNumError As Long
   
    Screen.MousePointer = vbHourglass
      
    'On Error GoTo CtrlError
    
   '*** Poner el objeto de Informe a un archivo RPT ***
   Set crxAplicacion = New CRAXDRT.Application
   Set crxReporte = crxAplicacion.OpenReport(strReportPath)
           
   '*** Conexión a la BD ***
   '"Connect Timeout"
   'crxReporte.Database.Tables(1).ConnectionProperties("Connect Timeout") = 10000
   'crxReporte.Database.Tables(1).Add "Connect Timeout", "15"
   crxReporte.Database.Tables(1).SetLogOnInfo gstrServer, gstrDataBase, gstrLogin, gstrPassword

    'JAFR 08/02/2016 --Refrescar el nombre de la base de datos
    'This removes the schema from the Database Table's Location property.
    Set crxTables = crxReporte.Database.Tables
    For Each crxTable In crxTables
        With crxTable
            .Location = .Location
        End With
    Next

    Set crxSections = crxReporte.Sections

    For intCantSections = 1 To crxSections.Count
        Set crxSection = crxSections(intCantSections)

        For intCantReportObjects = 1 To crxSection.ReportObjects.Count

            If crxSection.ReportObjects(intCantReportObjects).Kind = crSubreportObject Then
                Set crxSubreportObject = crxSection.ReportObjects(intCantReportObjects)

                'Open the subreport, and treat like any other report
                Set crxSubReport = crxSubreportObject.OpenSubreport
                Set crxTables = crxSubReport.Database.Tables

                For Each crxTable In crxTables
                    With crxTable
                        .SetLogOnInfo gstrServer, _
                        gstrDataBase, gstrLogin, gstrPassword
                        .Location = .Location
                    End With
                Next

            End If

        Next intCantReportObjects

    Next intCantSections
    
    'Insertar el Logo del reporte en la esquina superior izquierda del Page Header
    If Dir(gstrImagePath & "LogoReporte.jpg") <> Valor_Caracter And strIndLogo = Valor_Indicador Then
        Call crxReporte.Sections(2).AddPictureObject(gstrImagePath & "LogoReporte.jpg", 0, 0)
    End If
    
    'FIN JAFR
   
   '*** Asegurar que el informe accede a la Base de datos y refresca los datos ***
   crxReporte.DiscardSavedData
    
        
   Dim n As Integer, m As Integer
   
    If UBound(aParamF) > 0 Then
        For n = 0 To UBound(aParamF)
            For m = 1 To crxReporte.FormulaFields.Count
                If UCase(crxReporte.FormulaFields.Item(m).FormulaFieldName) = UCase(aParamFn(n)) Then
                    If crxReporte.FormulaFields.Item(m).ValueType = crNumberField Then
                        crxReporte.FormulaFields.Item(m).Text = CCur(aParamF(n))
                    Else
                        crxReporte.FormulaFields.Item(m).Text = Chr(34) & aParamF(n) & Chr(34)
                    End If
                    Exit For
                End If
            Next
        Next
    End If
       
    If UBound(aParamS) > 0 Then
        For n = 0 To UBound(aParamS)
          If crxReporte.ParameterFields(n + 1).ValueType = crDateTimeField Then
              crxReporte.ParameterFields(n + 1).AddCurrentValue CDate(Mid(aParamS(n), 1, 4) & "/" & Mid(aParamS(n), 5, 2) & "/" & Mid(aParamS(n), 7, 2)) 'CDate(aParamS(n))
          Else
              crxReporte.ParameterFields(n + 1).AddCurrentValue aParamS(n)
          End If
        Next
    End If
   'CrystalReport1.ParameterFields(i) = "BooleanParam;False;True"
   'CrystalReport1.ParameterFields(i) = "CurrencyParam;" & CurrVar & ";True"
   'CrystalReport1.ParameterFields(i) = "DateParam;Date(" & Text2.Text & ");True"
   'CrystalReport1.ParameterFields(i) = "StringParam;" & Text6.Text & ";True"
   
   'crxReporte.ParameterFields.GetItemByName("@Fecha").AddCurrentValue CDate("11/29/2002 12:00:00 AM")
   'crxReporte.ParameterFields.GetItemByName("@Linea").AddCurrentValue "101404"
   'crxReporte.ParameterFields.GetItemByName("@Prestatario").AddCurrentValue "%"
  
    If gstrSelFrml <> Valor_Caracter Then
        crxReporte.RecordSelectionFormula = gstrSelFrml
    End If
    
   crvVisor.ReportSource = crxReporte
   crvVisor.ViewReport
   
   '*** Ventana de vista previa al 100 % ***
   crvVisor.Zoom 100
   crvVisor.DisplayGroupTree = False
   crvVisor.EnableExportButton = True
   crvVisor.EnableRefreshButton = False
   crvVisor.EnableAnimationCtrl = False
   crvVisor.EnablePrintButton = True
   Screen.MousePointer = vbDefault
    Exit Sub
    
CtrlError:
    Screen.MousePointer = vbDefault
    intAccion = ControlErrores
    Select Case intAccion
        Case 0: Resume
        Case 1: Resume Next
        Case 2: Exit Sub
        Case Else
            lngNumError = err.Number
            err.Raise Number:=lngNumError
            err.Clear
    End Select
     
End Sub

Private Sub Form_Resize()

    With crvVisor
        .Top = 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With

End Sub

Public Sub SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())
        
    ReDim aParam(UBound(aReportParamS))
    ReDim aParamFn(UBound(aReportParamFn))
    ReDim aParamF(UBound(aReportParamF))

    aParamS() = aReportParamS()
    aParamFn() = aReportParamFn()
    aParamF() = aReportParamF()

End Sub

