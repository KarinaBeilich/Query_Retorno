Option Explicit

Private Const ABA_DADOS As String = "TESTE"
Private Const CONTA_OUTLOOK As String = "usuario@empresa.com"
Private Const PASTA_EMAIL As String = "PastaTeste"

Private Const FILTRO_ASSUNTO As String = _
"Confirmação de recebimento de arquivo retorno"

Private Const DIAS_ATRAS As Long = 4
Private Const PREFIXO_ARQ As String = "RET"

Sub AtualizarArquivosEmail()

    Dim ws As Worksheet

    Dim olApp As Object
    Dim olNs As Object
    Dim pasta As Object
    Dim itens As Object
    Dim item As Object

    Dim html As Object
    Dim tables As Object
    Dim tbl As Object
    Dim rw As Object

    Dim linhaDestino As Long
    Dim filtro As String
    Dim nomeArquivo As String
    Dim nomeCedente As String

    Dim dictArquivos As Object

    On Error GoTo TrataErro

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set ws = ThisWorkbook.Sheets(ABA_DADOS)

    ws.Range("B3:B" & ws.Rows.Count).ClearContents
    ws.Range("D3:F" & ws.Rows.Count).ClearContents

    linhaDestino = 3

    Set dictArquivos = CreateObject("Scripting.Dictionary")
    dictArquivos.CompareMode = vbTextCompare

    Set olApp = CreateObject("Outlook.Application")
    Set olNs = olApp.GetNamespace("MAPI")

    On Error Resume Next
    Set pasta = olNs.Folders(CONTA_OUTLOOK).Folders(PASTA_EMAIL)
    On Error GoTo TrataErro

    If pasta Is Nothing Then
        On Error Resume Next
        Set pasta = olNs.GetDefaultFolder(6).Parent.Folders(PASTA_EMAIL)
        On Error GoTo TrataErro
    End If

    If pasta Is Nothing Then
        MsgBox "Não foi possível encontrar a pasta '" & PASTA_EMAIL & "' na raiz do Outlook deste computador." & vbCrLf & _
               "Certifique-se de que o nome está exatamente igual." & vbCrLf & _
               "Nome procurado: " & PASTA_EMAIL, vbCritical, "Pasta Não Encontrada"
        GoTo Finalizar
    End If

    filtro = "[ReceivedTime] >= '" & _
             Format(Date - DIAS_ATRAS, "mm/dd/yyyy 00:00") & "'"

    Set itens = pasta.Items.Restrict(filtro)
    itens.Sort "[ReceivedTime]", True

    For Each item In itens
        If TypeName(item) = "MailItem" Then
            If item.ReceivedTime >= DateAdd("d", -DIAS_ATRAS, Now) Then
                If InStr(1, item.Subject, FILTRO_ASSUNTO, vbTextCompare) > 0 Then

                    Set html = CreateObject("htmlfile")
                    html.Body.innerHTML = item.HTMLBody

                    Set tables = html.getElementsByTagName("table")

                    For Each tbl In tables
                        For Each rw In tbl.Rows

                            If rw.Cells.Length >= 2 Then

                                nomeArquivo = Trim(rw.Cells(0).innerText)
                                nomeCedente = Trim(rw.Cells(1).innerText)

                                If InStr(1, UCase(nomeArquivo), PREFIXO_ARQ, vbTextCompare) > 0 Then

                                    If Not dictArquivos.Exists(nomeArquivo) Then
                                        dictArquivos.Add nomeArquivo, True

                                        ws.Cells(linhaDestino, "B").Value = nomeArquivo
                                        ws.Cells(linhaDestino, "D").Value = nomeCedente 
                                        ws.Cells(linhaDestino, "E").Value = item.Subject
                                        ws.Cells(linhaDestino, "F").Value = ExtrairDataArquivo(nomeArquivo)

                                        ws.Cells(linhaDestino, "F").NumberFormat = "dd/mm/yyyy"

                                        linhaDestino = linhaDestino + 1
                                    End If
                                End If
                            End If
                        Next rw
                    Next tbl

                End If
            End If
        End If
    Next item

    ws.Columns("B:F").AutoFit

    MsgBox dictArquivos.Count & _
           " arquivo(s) único(s) carregado(s) do Outlook.", vbInformation

Finalizar:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

TrataErro:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Erro: " & Err.Description, vbCritical

End Sub

Private Function ExtrairDataArquivo(nomeArquivo As String) As Variant
    Dim partes() As String
    Dim dataTxt As String

    On Error GoTo ErroData

    nomeArquivo = Replace(nomeArquivo, ".RET", "", , , vbTextCompare)
    partes = Split(nomeArquivo, "-")

    If UBound(partes) >= 4 Then
        dataTxt = partes(4)
        If Len(dataTxt) = 8 And IsNumeric(dataTxt) Then
            ExtrairDataArquivo = DateSerial( _
                CLng(Left(dataTxt, 4)), _
                CLng(Mid(dataTxt, 5, 2)), _
                CLng(Right(dataTxt, 2)))
            Exit Function
        End If
    End If

ErroData:
    ExtrairDataArquivo = ""
End Function