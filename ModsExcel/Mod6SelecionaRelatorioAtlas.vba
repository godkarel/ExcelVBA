Sub SelecionarArquivoXLS()
    ' Declarar variáveis
    Dim strCaminhoArquivo As String
    Dim FD As FileDialog

    ' Criar objeto FileDialog
    Set FD = Application.FileDialog(msoFileDialogFilePicker)

    ' Configurar propriedades
    With FD
        .AllowMultiSelect = False ' Selecionar apenas um arquivo
        .Filters.Clear
        .Filters.Add "Arquivos Excel", "*.xls" ' Mostrar apenas arquivos XLSX
        .Title = "Selecione o arquivo XLS"
    End With

    ' Exibir caixa de diálogo
    If FD.Show = -1 Then
        strCaminhoArquivo = FD.SelectedItems(1) ' Obter caminho do arquivo selecionado
    Else
        Exit Sub ' Usuário clicou em cancelar
    End If

    CopiarColunas strCaminhoArquivo
End Sub
