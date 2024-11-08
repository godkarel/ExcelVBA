Sub OnClickRegistrarDiarioDINR()
    ' Seleciona Planilha 20 que é o Diario de carga para DINR
    Planilha20.Select
    
    ' Salva o arquivo como CSV para DINR
    SalvarCSVDINR
    
    ' Zera o diario de DINR depois de registrar para a criação de um novo Diario
    ZerarDepoisDeSalvarDINR
    
    ' Seleciona Planilha 3 que é para voltar para a página inicial
    Planilha7.Select
End Sub

Sub ConfirmarAcaoSalvarDINR()
    Dim resposta As VbMsgBoxResult
    
    ' Exibe uma mensagem perguntando se o usuário tem certeza que quer continuar
    resposta = MsgBox("Você tem certeza que quer Salvar o Diario de Equipamentos não registrados ?", vbYesNo + vbQuestion, "Confirmação")

    ' Verifica a resposta do usuário
    If resposta = vbYes Then
    ' Coloque aqui o código que você deseja executar
        Call OnClickRegistrarDiarioDINR
        
    Else

        ' Interrompe a execução do código
        Exit Sub
    End If
End Sub
