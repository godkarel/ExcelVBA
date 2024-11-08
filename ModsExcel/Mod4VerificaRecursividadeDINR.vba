Sub VerificadorRecursividadeDINR()
    Dim valorColuna10 As String
    Dim valorColuna11 As String
    
    valorColuna10 = Planilha19.Cells(1, 10).Value
    valorColuna11 = Planilha19.Cells(1, 11).Value
    
    If valorColuna10 = "Verdadeiro" Then
        Loop_NomeDINR
    End If
    
    If valorColuna11 = "Verdadeiro" Then
        Loop_NomeDINR
    End If
End Sub

