Sub CriarHyperlinks()
    Dim ws As Worksheet
    Dim caminhoArquivo As String
    Dim i As Long
    Dim ultimaLinha As Long

    ' Defina a planilha
    Set ws = ThisWorkbook.Sheets("BOT DE PESQUISA")

    ' Encontre a última linha preenchida na coluna A a partir da linha 2
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Iterar sobre cada célula na coluna A a partir da linha 2 até a última linha preenchida
    For i = 2 To ultimaLinha
        caminhoArquivo = ws.Cells(i, 1).Value ' Caminho do arquivo na coluna A
        If caminhoArquivo <> "" Then
            ' Crie o hyperlink na coluna B
            ws.Hyperlinks.Add Anchor:=ws.Cells(i, 2), Address:=caminhoArquivo, TextToDisplay:="Abrir Arquivo"
        End If
    Next i

    
End Sub

