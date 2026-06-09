Sub gerarEmail()

    On Error GoTo erro

    Dim OutApp As Object
    Dim OutMail As Object

    Dim wsBase As Worksheet
    Dim wsDuplicados As Worksheet

    Dim ultimaLinha As Long
    Dim linhaDuplicado As Long

    Dim signature As String
    Dim nomeSalvar As String
    Dim caminhoArquivo As String


    Set wsBase = ThisWorkbook.Sheets("Mascara")

    ultimaLinha = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row


    Application.DisplayAlerts = False

    On Error Resume Next
    Sheets("TEMP_DUPLICADOS").Delete
    On Error GoTo 0

    Application.DisplayAlerts = True


    Set wsDuplicados = Sheets.Add

    wsDuplicados.Name = "TEMP_DUPLICADOS"


    wsBase.Range("A6:D6").Copy wsDuplicados.Range("A1")


    If ultimaLinha < 7 Then

        MsgBox "Nenhum registro encontrado na máscara.", vbInformation

        Application.DisplayAlerts = False
        wsDuplicados.Delete
        Application.DisplayAlerts = True

        Exit Sub

    End If

    wsBase.Range("A7:D" & ultimaLinha).Copy _
    wsDuplicados.Range("A2")

    wsDuplicados.Columns("A:D").AutoFit


    nomeSalvar = "Registros_Duplicados_" & Format(Now, "ddmmyyyy_hhmmss") & ".xlsx"

    caminhoArquivo = Environ("TEMP") & "\" & nomeSalvar

    wsDuplicados.Copy

    ActiveWorkbook.SaveAs caminhoArquivo

    ActiveWorkbook.Close False


    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    OutMail.Display

    signature = OutMail.HTMLBody

    With OutMail

        .To = "usuario@empresa.com.br"
        .CC = "gestor@empresa.com.br"

        .Subject = "Notificação de Chaves Duplicadas Identificadas - Fundo2"

        ' CORPO EMAIL (inclua o assunto a ser tratado abaixo)
        .HTMLBody = _
        "<BODY style='font-size:11pt;font-family:Calibri'>" & _
        "Prezados,<br><br>" & _
        "Bom dia,<br><br>" & _
        "Constatamos arquivos pendentes de verificação na esteira de validação automática.<br><br>" & _
        "Segue em anexo o arquivo contendo os registros duplicados identificados pelo sistema.<br><br>" & _
        "Atenciosamente,<br><br>" & _
        signature

        ' ANEXO
        .Attachments.Add caminhoArquivo

        .Display
        '.Send

    End With

    Application.DisplayAlerts = False
    wsDuplicados.Delete
    Application.DisplayAlerts = True

    Set OutMail = Nothing
    Set OutApp = Nothing

    MsgBox "Email criado com sucesso!", vbInformation

    Exit Sub

erro:

    MsgBox "Erro: " & Err.Description, vbCritical

End Sub