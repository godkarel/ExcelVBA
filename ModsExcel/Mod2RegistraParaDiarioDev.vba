Sub OnClickRegistrarDiarioDEV()
    ' Seleciona Planilha 11 que é o Diario de carga para desenvolvimento
    Planilha11.Select
    
    ' Salva o arquivo como CSV para desenvolvimento
    SalvarCSVDEV
    
    ' Zera o diario de desenvolvimento depois de registrar para a criação de um novo Diario
    ZerarDepoisDeSalvarDEV
    
    ' Seleciona Planilha 7 que é para voltar para a página inicial
    Planilha7.Select
End Sub

Sub ConfirmarSalvaDEVCSV()
    Dim resposta As VbMsgBoxResult
    
    ' Exibe uma mensagem perguntando se o usuário tem certeza que quer continuar
    resposta = MsgBox("Você tem certeza que quer Gerar o diario de Devolução?", vbYesNo + vbQuestion, "Confirmação")

    ' Verifica a resposta do usuário
    If resposta = vbYes Then

        ' Coloque aqui o código que você deseja executar
        Call OnClickRegistrarDiarioDEV
        
    Else
        ' Interrompe a execução do código
        Exit Sub
    End If
End Sub
