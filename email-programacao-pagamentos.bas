Sub EnviarEmailCobrancaVOClaro()

    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim Assinatura As String
    Dim CorpoEmail As String
    Dim CaminhoArquivo As String
    
    If ThisWorkbook.Path = "" Then
        MsgBox "Salve a planilha antes de enviar o e-mail.", vbExclamation
        Exit Sub
    End If
    
    ThisWorkbook.Save
    
    On Error Resume Next
    Set OutlookApp = GetObject(Class:="Outlook.Application")
    
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    CaminhoArquivo = ThisWorkbook.FullName
    
    With OutlookMail
        .To = "exemplo@xxx.com.br; exemplo@xxx.com.br"
        .Cc = "exemplo@xxx.com.br; exemplo@xxx.com.br"
        .Subject = "Programação de Pagamentos - CO - Wxxx"
        
        .Display
        
        Assinatura = .HTMLBody
        
        CorpoEmail = "<div style='font-family:Calibri; font-size:11pt;'>" & _
                "Bom dia, <br><br>" & _
                "Segue a programação de pagamentos: <br><br>" & _
            "</div>"
        
        .HTMLBody = CorpoEmail & Assinatura
        
        .Attachments.Add CaminhoArquivo
        
    End With
    
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing

End Sub