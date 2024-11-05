Sub TransformarColunaMEmInteiros()
    Dim wsDestino As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim valor As Variant
    
    ' Definir a aba de destino como "Carregamento"
    Set wsDestino = ThisWorkbook.Sheets("Carregamento")
    
    ' Encontrar a última linha com dados na coluna M
    ultimaLinha = wsDestino.Cells(wsDestino.Rows.Count, "M").End(xlUp).Row
    
    ' Loop para percorrer a coluna M (a partir da linha 2) e converter em números inteiros
    For i = 2 To ultimaLinha
        valor = wsDestino.Cells(i, "M").Value
        
        ' Verificar se o valor é numérico antes de converter
        If IsNumeric(valor) Then
            wsDestino.Cells(i, "M").Value = Int(valor) ' Converter para número inteiro
        End If
    Next i
End Sub
