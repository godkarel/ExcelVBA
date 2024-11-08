Sub ZeraDiarioPraNovoAtendimentoDINR()
    ' Zera o Diario para inserir o atendimento de um novo tecnico
    Dim LINHADIARIO As Long
    Dim ws As Worksheet
    Dim linhas As Long

    ' Define a planilha em que estamos trabalhando
    Set ws = Planilha19
    linhas = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    ' Loop para limpar as entradas do Diario
    For LINHADIARIO = 4 To linhas
        If ws.Cells(LINHADIARIO, 2) <> "" Then
            ws.Cells(LINHADIARIO, 2).ClearContents
            ws.Cells(LINHADIARIO, 3).ClearContents
        End If
    Next LINHADIARIO

    ' Zera o nome do tecnico
    ws.Cells(1, 3).Value = ""
    ws.Select
End Sub

