Sub ExcluirLinhasComDataMaisProximaDoAtual()
    Dim wsDestino As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim dataMaisProxima As Date
    Dim dataAtual As Date
    Dim intervaloDatas As Range
    Dim celula As Range
    Dim diferencaMinima As Long
    Dim diferencaDias As Long
    
    ' Definir a aba de destino como "Carregamento"
    Set wsDestino = ThisWorkbook.Sheets("Carregamento")
    
    ' Encontrar a última linha com dados na coluna I
    ultimaLinha = wsDestino.Cells(wsDestino.Rows.Count, "I").End(xlUp).Row
    
    ' Definir o intervalo de datas da coluna I (a partir da linha 2, excluindo o cabeçalho)
    Set intervaloDatas = wsDestino.Range("I2:I" & ultimaLinha)
    
    ' Definir a data atual
    dataAtual = Date
    
    ' Inicializar a data mais próxima e a diferença mínima com valores iniciais
    diferencaMinima = 999999 ' Um valor grande inicial
    
    ' Loop para encontrar a data mais próxima do dia atual
    For Each celula In intervaloDatas
        If IsDate(celula.Value) Then
            ' Converter o valor da célula em uma data
            diferencaDias = Abs(CLng(CDate(celula.Value) - dataAtual))
            ' Verificar se essa diferença é menor que a mínima encontrada até agora
            If diferencaDias < diferencaMinima Then
                diferencaMinima = diferencaDias
                dataMaisProxima = CDate(celula.Value)
            End If
        End If
    Next celula
    
    ' Agora que temos a data mais próxima, excluir todas as linhas que não correspondem a essa data
    For i = ultimaLinha To 2 Step -1
        If wsDestino.Cells(i, "I").Value <> Format(dataMaisProxima, "yyyy-mm-dd") Then
            wsDestino.Rows(i).Delete
        End If
    Next i
    
    MsgBox "Linhas com datas diferentes da mais próxima (" & Format(dataMaisProxima, "yyyy-mm-dd") & ") foram excluídas.", vbInformation
End Sub
