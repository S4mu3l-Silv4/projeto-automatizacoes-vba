Sub EnviarEmailCobrancaVOVivo()
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
        .To = "bryan.ristoff@h-partners.com; daniela.queiroz.freitas@huawei.com; Fatima.Ibrahim@huawei.com; gabriel.de.sobrinho@huawei.com; joao.brigagao.randstad@h-partners.com; luan.maciel@h-partners.com; misael.andrade@huawei.com"
        .CC = "joao.moreira@premcell.com.br; luan.pereira@premcell.com.br; kelly.martins@premcell.com.br"
        .Subject = "Cobrança de VO - Vivo - CO"
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