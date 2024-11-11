Sub AplicaOqueAchouNaLinhaCerta(Texto)

    Dim LINHADIARIO As Long

    For LINHADIARIO = 5 To 100
        
        Application.ScreenUpdating = False

        If Planilha13.Cells(LINHADIARIO, 3) = "" Then
        
            If Texto <> "" Then
    
                Planilha13.Cells(LINHADIARIO, 3) = Texto
        
                Texto = ""
                
                UserForm6.Hide
                ' Sai do loop depois de adicionar o texto na primeira linha vazia encontrada
                Exit For 
        
            End If
    
        End If
    
    Next LINHADIARIO
    
    Application.ScreenUpdating = True

    If Texto <> "" Then
        MsgBox "Não há linhas vazias disponíveis na coluna C.", vbExclamation
    End If

End Sub


