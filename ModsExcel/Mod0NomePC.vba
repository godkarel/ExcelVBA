Sub MostrarNomeComputador()
    Dim nomeComputador As String
    
    ' Obter o nome do computador
    nomeComputador = Environ("COMPUTERNAME")
    
    ' Exibir o nome do computador em uma caixa de mensagem
    MsgBox "Nome do Computador: " & nomeComputador, vbInformation, "Informações do Computador"
End Sub
