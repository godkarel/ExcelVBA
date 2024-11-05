Sub ZeraTudoDepoisdeSalvarMic()
    ' MACRO PARA ZERAR TODOS OS VALORES APÓS EXECUTAR O SALVAMENTO PARA DEIXAR A PLANILHA ZERADA.
    
    Dim LINHADIARIO As Long
    Dim LINHANS As Long
    Dim ws As Worksheet
    
    ' Define a planilha alvo
    Set ws = ThisWorkbook.Sheets("Planilha14")
    
    LINHANS = 4
    
    ' Loop pelas linhas da planilha
    For LINHADIARIO = 2 To 1000
        If ws.Cells(LINHADIARIO, 1) <> "" Then
            ' Zera os valores das células específicas
            ws.Cells(LINHADIARIO, 3).Resize(1, 3).ClearContents
            ws.Cells(LINHADIARIO, 6).ClearContents
            ws.Cells(LINHADIARIO, 12).ClearContents
            LINHANS = LINHANS + 1
        End If
    Next LINHADIARIO
End Sub

