Sub ExcluirLinhasBaseadoNaListaDeExclusao()
    Dim wsCarregamento As Worksheet
    Dim wsExclusoes As Worksheet
    Dim ultimaLinhaCarregamento As Long
    Dim ultimaLinhaExclusoes As Long
    Dim i As Long, j As Long
    Dim valorExclusao As String

    ' Definir as planilhas
    Set wsCarregamento = ThisWorkbook.Sheets("Carregamento")
    Set wsExclusoes = ThisWorkbook.Sheets("Lista de Exclusoes")
    
    ' Encontrar a última linha usada na aba "Carregamento" e "Lista de Exclusoes"
    ultimaLinhaCarregamento = wsCarregamento.Cells(wsCarregamento.Rows.Count, "K").End(xlUp).Row
    ultimaLinhaExclusoes = wsExclusoes.Cells(wsExclusoes.Rows.Count, "B").End(xlUp).Row
    
    ' Loop por cada item na "Lista de Exclusoes" (a partir de B2)
    For i = 2 To ultimaLinhaExclusoes
        valorExclusao = wsExclusoes.Cells(i, 2).Value ' Valor da célula B2, B3, etc.
        
        ' Verifica se o valor de exclusão não está vazio
        If valorExclusao <> "" Then
            ' Percorre a aba "Carregamento" de baixo para cima, excluindo as linhas que correspondem ao valor na coluna K
            For j = ultimaLinhaCarregamento To 2 Step -1
                If wsCarregamento.Cells(j, 11).Value = valorExclusao Then ' Coluna K (11ª coluna)
                    wsCarregamento.Rows(j).Delete
                End If
            Next j
        End If
        
        ' Atualiza a última linha da aba "Carregamento" após as exclusões
        ultimaLinhaCarregamento = wsCarregamento.Cells(wsCarregamento.Rows.Count, "K").End(xlUp).Row
    Next i

End Sub

