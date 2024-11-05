Sub CarregarDadosArquivoBTP()
    Dim wbOrigem As Workbook
    Dim wsDestino As Worksheet
    Dim wsOrigem As Worksheet
    Dim arquivoSelecionado As String
    Dim ultimaLinha As Long
    Dim i As Long
    Dim colunaMaxima As Long
    
    ' Definir a aba de destino como "Carregamento"
    Set wsDestino = ThisWorkbook.Sheets("Carregamento")
    
    ' Abrir a caixa de diálogo para seleção do arquivo XLSX
    arquivoSelecionado = Application.GetOpenFilename("Arquivos Excel (*.xlsx), *.xlsx", , "Selecione o Arquivo XLSX")
    
    ' Verificar se um arquivo foi selecionado
    If arquivoSelecionado = "Falso" Then
        MsgBox "Nenhum arquivo foi selecionado.", vbExclamation
        Exit Sub
    End If
    
    ' Abrir o arquivo selecionado
    Set wbOrigem = Workbooks.Open(arquivoSelecionado)
    
    ' Definir a primeira planilha do arquivo como origem
    Set wsOrigem = wbOrigem.Sheets(1)
    
    ' Limpar os dados anteriores da aba "Carregamento" nas colunas B até M (exceto o cabeçalho, ou seja, a linha 1)
    wsDestino.Range("B2:M" & wsDestino.Rows.Count).ClearContents
    
    ' Encontrar a última linha com dados na planilha de origem
    ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, 1).End(xlUp).Row
    
    ' Limitar o loop até a coluna M (13ª coluna)
    colunaMaxima = 13 ' Coluna M é a 13ª coluna
    
    ' Loop para copiar da coluna A até a coluna M do arquivo de origem e colar da coluna B até a coluna N na aba "Carregamento"
    For i = 1 To colunaMaxima
        wsOrigem.Range(wsOrigem.Cells(2, i), wsOrigem.Cells(ultimaLinha, i)).Copy wsDestino.Cells(2, i + 1) ' Coluna A para B, B para C, até M para N
    Next i
    
    ' Fechar o arquivo de origem
    wbOrigem.Close False
    
    MsgBox "Dados carregados com sucesso para a aba 'Carregamento'", vbInformation
End Sub
