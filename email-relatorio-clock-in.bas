Sub EnviarEmailRelatorioClockIn()

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
        .To = "exemplo@xxx.com; exemplo@xxx.com"
        .Cc = "exemplo@xxx.com"
        .Subject = "RES: Relatório de Clock-in - CO"
        
        .Display
        
        Assinatura = .HTMLBody
        
        CorpoEmail = "<div style='font-family:Calibri; font-size:11pt;'>" & _
                "Bom dia, <br><br>" & _
                "Segue o relatório de clock-in atualizado atualizado: <br>" & _
            "</div>"
        
        .HTMLBody = CorpoEmail & Assinatura
        
    End With
    
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing

End Sub