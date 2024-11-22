Sub AtualizaAtlas()

    DeletarConteudoAtlas
    SelecionarArquivoXLS
    AtualizarDataAtualAtlas
    
    ThisWorkbook.Save
    
End Sub

Sub AtualizarDataAtualAtlas()
    ' Declarar variáveis
    Dim ws As Worksheet

    ' Definir a planilha de destino
    Set ws = ThisWorkbook.Sheets("Pagina Inicial")

    ' Inserir a data atual na célula P11
    ws.Range("P11").Value = Date
End Sub

Sub ConfirmarAcaoATUALIZAATLAS()
    Dim resposta As VbMsgBoxResult
    
    ' Exibe uma mensagem perguntando se o usuário tem certeza que quer continuar
    resposta = MsgBox("DESEJA SINCRONIZAR O ATLAS ?", vbYesNo + vbQuestion, "Confirmação")

    ' Verifica a resposta do usuário
    If resposta = vbYes Then
        Call AtualizaAtlas
        
    Else
        Exit Sub
    End If
End Sub

