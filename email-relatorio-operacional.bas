Sub EnviarEmailRelatorioSemanal()

    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim Assinatura As String
    Dim CorpoEmail As String
    
    On Error Resume Next
    Set OutlookApp = GetObject(Class:="Outlook.Application")
    
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    With OutlookMail
        .To = "exemplo@xxx.com; exemplo@xxx.com; exemplo@xxx.com; exemplo@xxx.com"
        .Cc = ""
        .Subject = "Relatório semanal de fechamento operacional - CO -  [W/ano]"
        
        .Display
        
        Assinatura = .HTMLBody
        
        CorpoEmail = "<div style='font-family:Calibri; font-size:11pt;'>" & _
            "Boa tarde, <br><br>" & _
            "Segue o relatório semanal referente ao fechamento operacional (W): <br>" & _
        "</div>"
        
        .HTMLBody = CorpoEmail & Assinatura
        
    End With
    
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing

End Sub