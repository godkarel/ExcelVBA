Sub ZerarParaNovoKitSeiLa()
    ' Zera o Diario para inserir o atendimento de um novo tecnico
    Dim LINHADIARIO As Long
    Dim ws As Worksheet
    Dim linhas As Long

    ' Define a planilha em que estamos trabalhando
    Set ws = Planilha28
    linhas = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    ' Loop para limpar as entradas do Diario
    For LINHADIARIO = 4 To linhas
        If ws.Cells(LINHADIARIO, 3) <> "" Then
            ws.Cells(LINHADIARIO, 3).ClearContents
        End If
    Next LINHADIARIO
    SelecionaDiarioMic
End Sub


