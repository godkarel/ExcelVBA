Sub VerificadorRecursividadeDEV()

If Planilha9.Cells(1, 10) <> "" Then

    If Planilha9.Cells(1, 10) = "Verdadeiro" Then
        
        Loop_REDEV
        
    End If
    
End If

If Planilha9.Cells(1, 11) <> "" Then

    If Planilha9.Cells(1, 11) = "Verdadeiro" Then
        
        Loop_NomeDEV
    
    End If
    
End If



End Sub
