Sub EnviarEmailCobrancaVOClaro()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim assinatura As String
    Dim CorpoEmail As String
    On Error Resume Next
    Set OutlookApp = GetObject(Class:="Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    Set OutlookMail = OutlookApp.CreateItem(0)
    With OutlookMail
        .To = "alisson.rocha.manpower@h-partners.com; cassio.santos@huawei.com; diego.cardoso@huawei.com; mariana.rocha.souza@h-partners.com; misael.andrade@huawei.com"
        .CC = "joao.moreira@premcell.com.br; luan.pereira@premcell.com.br; kelly.martins@premcell.com.br"
        .Subject = "Cobrança de VO - Claro - CO"
        .Display
        assinatura = .HTMLBody
        CorpoEmail = "<div style='font-family: Calibri; font-size: 11pt;'>" & _
            "Boa tarde, <br><br>" & _
            "Poderiam verificar as VOs abaixo por gentileza? Esses casos estão com pendência de aprovação: <br>" & _
        "</div>"
        .HTMLBody = CorpoEmail & assinatura
    End With
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub