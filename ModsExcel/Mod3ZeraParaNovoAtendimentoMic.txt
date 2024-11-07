Sub ZeraDiarioPraNovoAtendimentoMIC()
    ' Zera o Diario para inserir o atendimento de um novo tecnico
    Dim LINHADIARIO As Long
    Dim ws As Worksheet
    Dim linhas As Long

    ' Define a planilha em que estamos trabalhando
    Set ws = Planilha13
    linhas = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    ' Loop para limpar as entradas do Diario
    For LINHADIARIO = 5 To linhas
        If ws.Cells(LINHADIARIO, 2) <> "" Then
            ws.Cells(LINHADIARIO, 2).ClearContents
            ws.Cells(LINHADIARIO, 4).ClearContents
            ws.Cells(LINHADIARIO, 5).ClearContents
            ws.Cells(LINHADIARIO, 3).ClearContents
        End If
    Next LINHADIARIO

    ' Aplica o código para busca
    AplicaCodigoParaBusca

    ' Zera o nome do tecnico em várias planilhas
    Planilha3.Cells(1, 3).Value = ""
    Planilha9.Cells(1, 3).Value = ""
    Planilha13.Cells(2, 3).Value = ""
End Sub


Sub AplicaCodigoParaBusca()
    ' Define a planilha em que estamos trabalhando
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("MICELANEAS")

    ' Aplica a fórmula em B5
    ws.Range("B5").FormulaR1C1 = "=IF(RC[1]<>"""",VLOOKUP(RC[1],'Biblioteca de Mic'!R1C[-1]:R149C[1],3,0),"""")"

    ' Preenche a fórmula até a linha 147
    ws.Range("B5:B147").FillDown

    ' Aplica a fórmula em D5
    ws.Range("D5").FormulaR1C1 = "=IF(RC[-1]<>"""",VLOOKUP(RC[-1],'Biblioteca de Mic'!R1C1:R149C3,2,0),"""")"
    
    ws.Range("D5:D147").FillDown

    ' Seleciona a célula C5
    ws.Range("C5").Select
End Sub

