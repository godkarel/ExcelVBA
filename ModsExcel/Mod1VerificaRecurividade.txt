Sub VerificadorRecursividade()
    Dim valorColuna10 As String
    Dim valorColuna11 As String
    
    valorColuna10 = Planilha3.Cells(1, 10).Value
    valorColuna11 = Planilha3.Cells(1, 11).Value
    
    If valorColuna10 = "Verdadeiro" Then
        Loop_RE
    End If
    
    If valorColuna11 = "Verdadeiro" Then
        Loop_Nome
    End If
End Sub

