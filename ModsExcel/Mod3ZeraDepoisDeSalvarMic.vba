Sub ZerarDepoisDeSalvarMIC()
    Dim ws As Worksheet
    Dim i As Long
    
    ' Defina a aba que você quer trabalhar
    Set ws = ThisWorkbook.Sheets("Diario Mic")
    
    ' Deleta o conteúdo das células da coluna B da linha 2 até a linha 7000
    ws.Range("A2:A2000").ClearContents
    
    ' Deleta o conteúdo das células da coluna C da linha 2 até a linha 7000
    ws.Range("F2:F2000").ClearContents
    
    ' Deleta o conteúdo das células da coluna D da linha 2 até a linha 7000
    ws.Range("L2:L2000").ClearContents

    ThisWorkbook.Save
End Sub


