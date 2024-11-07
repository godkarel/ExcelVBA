Sub SelecionarArquivoMICXLSX()
    ' Declarar variáveis
    Dim strCaminhoArquivo As String
    Dim FD As FileDialog
    
    DeletarConteudoMIC
    
    ' Criar objeto FileDialog
    Set FD = Application.FileDialog(msoFileDialogFilePicker)

    ' Configurar propriedades
    With FD
        .AllowMultiSelect = False ' Selecionar apenas um arquivo
        .Filters.Clear
        .Filters.Add "Arquivos Excel", "*.xlsx" ' Mostrar apenas arquivos XLSX
        .Title = "Selecione o arquivo XLSX"
    End With

    ' Exibir caixa de diálogo
    If FD.Show = -1 Then
        strCaminhoArquivo = FD.SelectedItems(1) ' Obter caminho do arquivo selecionado
    Else
        Exit Sub ' Usuário clicou em cancelar
    End If

    CopiarColunasMisc strCaminhoArquivo
End Sub

Sub ConfirmarAcaoConfirmaMic()
    Dim resposta As VbMsgBoxResult
    
    ' Exibe uma mensagem perguntando se o usuário tem certeza que quer continuar
    resposta = MsgBox("DESEJA CARREGAR A MIC ?", vbYesNo + vbQuestion, "Confirmação")

    ' Verifica a resposta do usuário
    If resposta = vbYes Then
        Call SelecionarArquivoMICXLSX
        
    Else
        Exit Sub
    End If
End Sub

