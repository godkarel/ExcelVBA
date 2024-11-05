Sub OnClickRegistrarDiarioAcessorio()
    
    'seleciona Planilha 1 que é o Diario de carga
    Planilha34.Select
    
    'Salva o arquivo como CSV ja com a configuração do teclado
    SalvarCSVAcessorio
    
    'Zera o diario depois de registrar para a criação de um novo Diario
    ZerarDepoisDeSalvarAcessorio
    
    'Seleciona Planilha 7 que é para ir pra pagina inicial
    Planilha7.Select

End Sub

Sub ConfirmarSalvaCSVAcessorio()
    Dim resposta As VbMsgBoxResult
    
    ' Exibe uma mensagem perguntando se o usuário tem certeza que quer continuar
    resposta = MsgBox("Você tem certeza que quer Gerar o Diariao?", vbYesNo + vbQuestion, "Confirmação")

    ' Verifica a resposta do usuário
    If resposta = vbYes Then

        ' Coloque aqui o código que você deseja executar
        Call OnClickRegistrarDiarioAcessorio
        
    Else
        ' Interrompe a execução do código
        Exit Sub
    End If
End Sub
