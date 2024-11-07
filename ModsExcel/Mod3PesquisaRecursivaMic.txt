Sub Loop_NomeMic()

    Dim LINHADIARIO As Long
    Dim LINHANS As Long
    Dim ws14 As Worksheet
    Dim ws13 As Worksheet
    Dim ultimaLinha As Long
    Dim lastRow As Long
    Dim countFilledRows As Long
    Dim nextEmptyRow As Long
    Dim j As Long
    
    ' Definindo as variáveis de planilhas
    Set ws14 = ThisWorkbook.Sheets("Diario Mic")
    Set ws13 = ThisWorkbook.Sheets("MICELANEAS")
    
    nextEmptyRow = 2
    
    Do While ws14.Cells(nextEmptyRow, 1).Value <> ""
        nextEmptyRow = nextEmptyRow + 1
    Loop
    
    lastRow = ws13.Cells(ws13.Rows.Count, 3).End(xlUp).Row - 4
    
    countFilledRows = Application.WorksheetFunction.CountA(ws13.Range(ws13.Cells(5, 3), ws13.Cells(lastRow, 3)))
    
    ' Inicializa a linha de início para LINHANS
    LINHANS = 5
    
    ' Encontra a última linha na coluna A da Planilha14
    ultimaLinha = lastRow + nextEmptyRow

    ' Desativa a atualização da tela para melhorar a performance
    Application.ScreenUpdating = False

    ' Laço para percorrer as linhas de 2 até a última linha na Planilha14
    For LINHADIARIO = nextEmptyRow To ultimaLinha
    
         ' Verificar se a célula A está vazia
        If ws13.Cells(LINHANS, 4) = "" Then
            ws13.Cells(2, 3) = ""
        End If

        ' Verifica se a célula na coluna A está vazia
        If ws14.Cells(LINHADIARIO, 1) = "" Then

            ' Copia o valor da Planilha13, coluna B para Planilha14, coluna C
            ws14.Cells(LINHADIARIO, 1) = ws13.Cells(LINHANS, 4)

            ' Copia o valor da Planilha13, coluna D para Planilha14, coluna A
            ws14.Cells(LINHADIARIO, 6) = ws13.Cells(LINHANS, 5)
            
            ws14.Cells(LINHADIARIO, 12) = ws13.Cells(2, 4)
            ' Verifica se a célula na coluna C não está vazia
            If ws13.Cells(2, 4) <> "" Then
                ' Copia o valor da Planilha13, célula D1 para Planilha14, coluna L
                ws14.Cells(LINHADIARIO, 12) = ws13.Cells(2, 4)
            End If

            ' Incrementa LINHANS
            LINHANS = LINHANS + 1

        End If

    Next LINHADIARIO

    ' Chama a sub-rotina ZeraDiarioPraNovoAtendimentoMIC
    ZeraDiarioPraNovoAtendimentoMIC

    ' Reativa a atualização da tela
    Application.ScreenUpdating = True

End Sub



