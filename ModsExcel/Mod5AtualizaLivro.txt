Sub AtualizaLivro()
    
    DeletarConteudo
    SelecionarArquivoXLSX
    
    ThisWorkbook.Save
End Sub

Sub ConfirmarAcaoATUALIZALIVRO()
    Dim resposta As VbMsgBoxResult
    
    ' Exibe uma mensagem perguntando se o usuário tem certeza que quer continuar
    resposta = MsgBox("DESEJA SINCRONIZAR O LIVRO ?", vbYesNo + vbQuestion, "Confirmação")

    ' Verifica a resposta do usuário
    If resposta = vbYes Then
        Call AtualizaLivro
        
    Else
        Exit Sub
    End If
End Sub

