Sub OnClickRegistrarDiarioMic()
    ' Seleciona Planilha 14 que é o Diario de carga para MIC
    Planilha14.Select
    
    ' Salva o arquivo como CSV para MIC
    SalvarCSVMIC
    
    ' Zera o diario de MIC depois de registrar para a criação de um novo Diario
    ZerarDepoisDeSalvarMIC
    
    ' Seleciona Planilha 7 que é para voltar para a página inicial
    Planilha7.Select
End Sub

Sub ConfirmarSalvaMICCSV()
    Dim resposta As VbMsgBoxResult
    
    ' Exibe uma mensagem perguntando se o usuário tem certeza que quer continuar
    resposta = MsgBox("Você tem certeza que quer Gerar o diario de MIC?", vbYesNo + vbQuestion, "Confirmação")

    ' Verifica a resposta do usuário
    If resposta = vbYes Then

        ' Coloque aqui o código que você deseja executar
        Call OnClickRegistrarDiarioMic
        
    Else
        ' Interrompe a execução do código
        Exit Sub
    End If
End Sub
