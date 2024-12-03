Sub CriarTabelaDinamica()
    Dim wsCarregamento As Worksheet
    Dim ultimaLinha As Long
    Dim ultimaColuna As Long
    Dim tblDestino As Worksheet
    Dim sourceRange As String
    
    ' Definir a aba "Carregamento"
    Set wsCarregamento = ThisWorkbook.Sheets("Carregamento")
    
    ' Criar uma nova planilha para a Tabela Dinâmica
    Set tblDestino = ThisWorkbook.Sheets.Add
    tblDestino.Name = "Dinamica" ' Nome da planilha onde a tabela dinâmica será criada
    
    ' Determinar a última linha e a última coluna da aba "Carregamento"
    ultimaLinha = wsCarregamento.Cells(wsCarregamento.Rows.Count, "B").End(xlUp).Row
    ultimaColuna = wsCarregamento.Cells(1, wsCarregamento.Columns.Count).End(xlToLeft).Column
    
    ' Construir o intervalo de origem da Tabela Dinâmica usando a notação de células A1
    sourceRange = wsCarregamento.Name & "!" & wsCarregamento.Cells(1, 2).Address(False, False) & ":" & wsCarregamento.Cells(ultimaLinha, ultimaColuna).Address(False, False)
    
    ' Verificar se o intervalo está correto
    MsgBox "Intervalo de origem: " & sourceRange
    
    ' Criar o cache da Tabela Dinâmica e configurar a Tabela Dinâmica
    ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=sourceRange, _
        Version:=8).CreatePivotTable _
        TableDestination:=tblDestino.Range("A3").Address(ReferenceStyle:=xlR1C1), _
        TableName:="Dinamica", _
        DefaultVersion:=8
    
    ' Configurar as opções da Tabela Dinâmica
    With tblDestino.PivotTables("Dinamica")
        .ColumnGrand = True
        .RowGrand = True
        .HasAutoFormat = True
        .RowAxisLayout xlCompactRow
    End With
    
    ' Configurar os campos da Tabela Dinâmica
    With tblDestino.PivotTables("Dinamica2").PivotFields("1TEC")
        .Orientation = xlRowField
        .Position = 1
    End With
    With tblDestino.PivotTables("Dinamica2").PivotFields("Descrição")
        .Orientation = xlRowField
        .Position = 2
    End With
    tblDestino.PivotTables("Dinamica2").AddDataField tblDestino.PivotTables("Dinamica2").PivotFields("Quantidade lançada"), "Soma de Quantidade lançada", xlSum
    
    MsgBox "Tabela dinâmica criada com sucesso!", vbInformation
End Sub


