Sub DeletaAntesUTPProcurados()
    Dim ws As Worksheet
    Dim i As Long
    
    ' Define a planilha "CONTROLEUTP"
    Set ws = ThisWorkbook.Sheets("CONTROLEUTP")
    
    ' Loop de 100 a 2 em ordem decrescente para evitar problemas com a exclusão de linhas
    For i = 100 To 2 Step -1
        ' Exclui a linha i
        ws.Rows(i).Delete
    Next i
End Sub

