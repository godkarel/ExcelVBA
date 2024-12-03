Sub ExcluirDevolucoes()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim deleteRows() As Long
    Dim deleteCount As Long
    
    ' Definir a aba onde estão os dados pelo nome
    Set ws = ThisWorkbook.Sheets("CONTROLEUTP")
    
    ' Encontrar a última linha com dados na coluna C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ReDim deleteRows(1 To lastRow) ' Array para armazenar linhas a serem excluídas
    deleteCount = 0 ' Contador de linhas a serem excluídas
    
    ' Percorrer as linhas da coluna C para verificar cada técnico
    For i = 2 To lastRow ' Assumindo que a primeira linha contém cabeçalhos e começa na linha 2
        ' Verificar se a célula na coluna E contém "CABO UTP"
        If ws.Cells(i, "E").Value = "CABO UTP" Then
            Dim technician As String
            Dim devRow As Long
            
            ' Obter o nome do técnico da coluna C
            technician = ws.Cells(i, "C").Value
            
            ' Procurar a próxima linha com o mesmo nome do técnico na coluna C
            devRow = ws.Range("C:C").Find(technician, After:=ws.Cells(i, "C")).Row
            
            ' Verificar se a célula correspondente na coluna F contém "DEV CABO UTP"
            If ws.Cells(devRow, "E").Value = "DEV CABO UTP" Then
                ' Marcar as linhas para exclusão
                deleteCount = deleteCount + 1
                deleteRows(deleteCount) = i
                deleteCount = deleteCount + 1
                deleteRows(deleteCount) = devRow
            End If
        End If
    Next i
    
    ' Excluir as linhas marcadas para exclusão
    If deleteCount > 0 Then
        For i = deleteCount To 1 Step -1
            ws.Rows(deleteRows(i)).Delete
        Next i
        
    Else
        
    End If
End Sub

Sub DeletarLinhasPorValor()
    Dim ws As Worksheet
    Dim rng As Range, cell As Range
    Dim deleteRange As Range
    
    ' Define a planilha "CONTROLEUTP"
    Set ws = ThisWorkbook.Sheets("CONTROLEUTP")
    
    ' Define o intervalo na coluna E
    Set rng = ws.Range("E1:E" & ws.Cells(ws.Rows.Count, "E").End(xlUp).Row)
    
    ' Loop pelas células do intervalo
    For Each cell In rng
        ' Verifica se o valor na célula é "DEV CABO UTP"
        If cell.Value = "DEV CABO UTP" Then
            ' Adiciona a célula ao intervalo de exclusão
            If deleteRange Is Nothing Then
                Set deleteRange = cell.EntireRow
            Else
                Set deleteRange = Union(deleteRange, cell.EntireRow)
            End If
        End If
    Next cell
    
    ' Exclui as linhas no intervalo de exclusão
    If Not deleteRange Is Nothing Then
        deleteRange.Delete
    End If
End Sub


