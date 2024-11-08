Sub DeletarConteudoMIC()
    Dim ws As Worksheet
    Dim i As Long
    
    ' Defina a aba que você quer trabalhar
    Set ws = ThisWorkbook.Sheets("Diario Mic")
    
    ' Deleta o conteúdo das células da coluna B da linha 2 até a linha 7000
    ws.Range("A2:A1000").ClearContents
    
    ' Deleta o conteúdo das células da coluna C da linha 2 até a linha 7000
    ws.Range("F2:F1000").ClearContents
    
    ' Deleta o conteúdo das células da coluna D da linha 2 até a linha 7000
    ws.Range("L2:L1000").ClearContents
End Sub

