Sub CarregarDuplicados()

    Dim wsFundo1 As Worksheet
    Dim wsFundo2 As Worksheet
    Dim wsMascara As Worksheet

    Dim UltimaLinhaFundo1 As Long
    Dim UltimaLinhaFundo2 As Long
    Dim linhaDestino As Long
    Dim i As Long

    Set wsFundo1 = Sheets("Base Fundo1")
    Set wsFundo2 = Sheets("Base Fundo2")
    Set wsMascara = Sheets("Mascara")

    wsMascara.Range("A7:D10000").ClearContents

    linhaDestino = 7

    ' --- PROCESSAMENTO FUNDO 1 ---
    UltimaLinhaFundo1 = wsFundo1.Cells(wsFundo1.Rows.Count, "B").End(xlUp).Row

    For i = 2 To UltimaLinhaFundo1

        If UCase(Trim(wsFundo1.Cells(i, 6).Value)) = "DUPLICADO" Then

            ' Nome do Arquivo -> coluna B
            wsMascara.Cells(linhaDestino, 1).Value = wsFundo1.Cells(i, 2).Value

            ' cccCodigo -> coluna D
            wsMascara.Cells(linhaDestino, 2).Value = wsFundo1.Cells(i, 4).Value

            ' dataImportacao -> coluna E
            wsMascara.Cells(linhaDestino, 3).Value = wsFundo1.Cells(i, 5).Value

            ' Origem
            wsMascara.Cells(linhaDestino, 4).Value = "Fundo1"

            linhaDestino = linhaDestino + 1

        End If

    Next i

    ' --- PROCESSAMENTO FUNDO 2 ---
    UltimaLinhaFundo2 = wsFundo2.Cells(wsFundo2.Rows.Count, "B").End(xlUp).Row

    For i = 2 To UltimaLinhaFundo2

        If UCase(Trim(wsFundo2.Cells(i, 6).Value)) = "DUPLICADO" Then

            ' Nome do Arquivo -> coluna B
            wsMascara.Cells(linhaDestino, 1).Value = wsFundo2.Cells(i, 2).Value

            ' cccCodigo -> coluna D
            wsMascara.Cells(linhaDestino, 2).Value = wsFundo2.Cells(i, 4).Value

            ' dataImportacao -> coluna E
            wsMascara.Cells(linhaDestino, 3).Value = wsFundo2.Cells(i, 5).Value

            ' Origem
            wsMascara.Cells(linhaDestino, 4).Value = "Fundo2"

            linhaDestino = linhaDestino + 1

        End If

    Next i

    MsgBox "Duplicados carregados com sucesso!", vbInformation

End Sub