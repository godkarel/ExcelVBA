Sub FiltrarCaboUTP()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    
    ' Referenciar a aba CONTROLEUTP
    Set ws = ThisWorkbook.Sheets("CONTROLEUTP")
    
    ' Desativar atualização da tela
    Application.ScreenUpdating = False
    
    ' Encontrar a última linha na coluna E
    ultimaLinha = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    
    ' Filtrar a coluna E para mostrar apenas linhas que contêm "CABO UTP" ou "DEV CABO UTP"
    ws.Range("E1:E" & ultimaLinha).AutoFilter Field:=1, Criteria1:="<>CABO UTP", Operator:=xlAnd, Criteria2:="<>DEV CABO UTP"
    
    ' Excluir as linhas visíveis (ou seja, aquelas que não contêm "CABO UTP" ou "DEV CABO UTP")
    For i = ultimaLinha To 2 Step -1
        If Not ws.Rows(i).Hidden Then
            ws.Rows(i).Delete
        End If
    Next i
    
    ' Remover o filtro
    ws.AutoFilterMode = False
    
    ' Ativar atualização da tela
    Application.ScreenUpdating = True
End Sub

