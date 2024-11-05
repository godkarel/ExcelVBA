Sub DeletarAbaDinamica()
    Dim ws As Worksheet
    Dim abaNome As String
    abaNome = "Dinamica"
    
    ' Verificar se a aba existe
    On Error Resume Next ' Ignorar erros se a aba não existir
    Set ws = ThisWorkbook.Sheets(abaNome)
    On Error GoTo 0 ' Retornar ao tratamento de erros padrão

    ' Se a aba existir, excluí-la
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False ' Desativar alertas de confirmação
        ws.Delete
        Application.DisplayAlerts = True ' Reativar alertas de confirmação
    End If
End Sub
