Sub ZeraDiarioPraNovoAtendimento()
    ' Zera o Diario para inserir o atendimento de um novo tecnico
    Dim LINHADIARIO As Long
    Dim ws As Worksheet
    Dim linhas As Long

    ' Define a planilha em que estamos trabalhando
    Set ws = Planilha3
    linhas = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    ' Loop para limpar as entradas do Diario
    For LINHADIARIO = 4 To linhas
        If ws.Cells(LINHADIARIO, 2) <> "" Then
            ws.Cells(LINHADIARIO, 2).ClearContents
        End If
    Next LINHADIARIO

    ' Zera o nome do tecnico em várias planilhas
    Planilha3.Cells(1, 3).Value = ""
    Planilha9.Cells(1, 3).Value = ""
    Planilha13.Cells(2, 3).Value = ""
End Sub

