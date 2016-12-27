Attribute VB_Name = "NegocioParticipacion"
Option Explicit

Function ValidarNumFolio(strpNumFolio As String, strpTipoSolicitud As String, strpCodFondo As String, ctrlpFormulario As Form) As Boolean

    Dim adoRegistro As ADODB.Recordset
    Dim adoRegistro2 As ADODB.Recordset
    Dim maxNumFolio As String
        
    Set adoRegistro = New ADODB.Recordset
    Set adoRegistro2 = New ADODB.Recordset
    ValidarNumFolio = False
    
    With adoComm
        .CommandText = "SELECT NumFolio FROM ParticipeSolicitud WHERE NumFolio='" & strpNumFolio & "' AND TipoSolicitud='" & strpTipoSolicitud & "' AND " & _
            "CodFondo='" & strpCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            adoComm.CommandText = "SELECT (MAX(NumFolio)+1) NumFolio" & _
            " FROM ParticipeSolicitud WHERE TipoSolicitud='" & strpTipoSolicitud & "' AND CodFondo='" & _
            strpCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            Set adoRegistro2 = adoComm.Execute
            
            If Not adoRegistro2.EOF Then
                maxNumFolio = CStr(adoRegistro2("NumFolio"))
            End If
            adoRegistro2.Close: Set adoRegistro2 = Nothing
            
            MsgBox "El Número de Papeleta ya existe. Está disponible el Nº de papeleta: " & maxNumFolio, vbCritical, gstrNombreEmpresa
            ctrlpFormulario.txtNumPapeleta.SetFocus
            ctrlpFormulario.txtNumPapeleta.SelStart = 0
            ctrlpFormulario.txtNumPapeleta.SelLength = Len(ctrlpFormulario.txtNumPapeleta.Text)
            Exit Function
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Si No existe ***
    ValidarNumFolio = True
    
End Function
