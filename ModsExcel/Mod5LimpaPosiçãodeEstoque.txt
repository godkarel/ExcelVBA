Sub DeletarConteudo()
    Dim ws As Worksheet
    Dim i As Long
    
    ' Defina a aba que você quer trabalhar
    Set ws = ThisWorkbook.Sheets("Equipamentos")
    
    ' Deleta o conteúdo das células da coluna B da linha 2 até a linha 7000
    ws.Range("B2:B15000").ClearContents
    
    ' Deleta o conteúdo das células da coluna C da linha 2 até a linha 7000
    ws.Range("C2:C15000").ClearContents
    
    ' Deleta o conteúdo das células da coluna D da linha 2 até a linha 7000
    ws.Range("D2:D15000").ClearContents
    
    ' Deleta o conteúdo das células da coluna E da linha 2 até a linha 7000
    ws.Range("E2:E15000").ClearContents
    
    ' Deleta o conteúdo das células da coluna F da linha 2 até a linha 7000
    ws.Range("F2:F15000").ClearContents
    
    ' Deleta o conteúdo das células da coluna G da linha 2 até a linha 7000
    ws.Range("G2:G15000").ClearContents
End Sub

