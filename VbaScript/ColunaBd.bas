Option Explicit

Sub AtualizarColunaA()

    Dim wsBase As Worksheet
    Dim wsTeste As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long

    Dim texto As String
    Dim bloco As String
    Dim dict As Object

    Dim dataFiltro As String
    Dim dataHoje As Date

    Application.ScreenUpdating = False

    Set wsBase = ThisWorkbook.Sheets("Fundo2")
    Set wsTeste = ThisWorkbook.Sheets("TESTE")
    Set dict = CreateObject("Scripting.Dictionary")

    wsTeste.Range("A3:A" & wsTeste.Rows.Count).ClearContents

    ultimaLinha = wsBase.Cells(wsBase.Rows.Count, "B").End(xlUp).Row

    dataHoje = Date

    If Weekday(dataHoje, vbMonday) = 1 Then
        dataFiltro = Format(dataHoje - 2, "yyyymmdd") 
    Else
        dataFiltro = Format(dataHoje - 1, "yyyymmdd") 
    End If

    Dim linhaDestino As Long
    linhaDestino = 3

    For i = 2 To ultimaLinha

        texto = wsBase.Cells(i, "B").Value

        If Len(texto) > 0 Then

            If InStr(texto, "RET-FUNDO2") > 0 Then

                bloco = Mid(texto, InStr(texto, "RET-FUNDO2"))

                If InStr(bloco, ".RET") > 0 Then
                    bloco = Left(bloco, InStr(bloco, ".RET") - 1)
                End If

                If InStr(bloco, dataFiltro) > 0 Then

                    If Not dict.Exists(bloco) Then
                        dict.Add bloco, True
                        wsTeste.Cells(linhaDestino, "A").Value = bloco
                        linhaDestino = linhaDestino + 1
                    End If

                End If

            End If

        End If

    Next i

    Application.ScreenUpdating = True

    MsgBox dict.Count & " arquivo(s) carregado(s) para " & dataFiltro, vbInformation

End Sub