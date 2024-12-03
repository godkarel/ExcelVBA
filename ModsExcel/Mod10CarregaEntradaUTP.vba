Sub CarregarEntradaCaboUTP(EnderecoPlan As String)
    Dim wb As Workbook
    Dim wsAtual As Worksheet
    Dim wsNova As Worksheet
    Dim wsDados As Worksheet
    Dim ultimaLinha As Long
    Dim ultimaLinhaNova As Long
    Dim linha As Long
    Dim proximaLinha As Long
    Dim abaExistente As Boolean
    Dim dataAtual As Date
    Dim data As Date
    Dim valorCelula As Variant
    
    ' Desativar atualização da tela
    Application.ScreenUpdating = False
    
    ' Definir a data atual
    dataAtual = Date
    
    ' Abre o arquivo ControlCaboUTP.xlsx
    Set wb = Workbooks.Open(EnderecoPlan)
    Set wsAtual = ThisWorkbook.Sheets(1) ' Planilha atual que carrega os dados
    
    ' Verificar se a aba CONTROLEUTP já existe
    abaExistente = False
    On Error Resume Next
    Set wsNova = ThisWorkbook.Sheets("CONTROLEUTP")
    On Error GoTo 0
    
    If wsNova Is Nothing Then
        ' Adiciona uma nova planilha se a aba não existir
        Set wsNova = ThisWorkbook.Sheets.Add(After:=wsAtual)
        wsNova.Name = "CONTROLEUTP"
    Else
        ' Se a aba já existir, marca a flag para utilizar a aba existente
        abaExistente = True
    End If
    
    ' Acessa a aba Dados na planilha ControlCaboUTP
    Set wsDados = wb.Sheets("Dados")
    
    ' Encontrar a última linha na aba CONTROLEUTP
    If abaExistente Then
        ultimaLinhaNova = wsNova.Cells(wsNova.Rows.Count, "A").End(xlUp).Row
        proximaLinha = ultimaLinhaNova + 1 ' Próxima linha na aba CONTROLEUTP
    Else
        proximaLinha = 1 ' Próxima linha na nova planilha
    End If
    
    ' Encontrar a última linha na aba Dados
    ultimaLinha = wsDados.Cells(wsDados.Rows.Count, "A").End(xlUp).Row
    
    ' Copiar os dados da aba Dados para a aba CONTROLEUTP
    For linha = 1 To ultimaLinha
        ' Obter o valor da célula na coluna F da aba Dados
        valorCelula = wsDados.Cells(linha, "F").Value
        ' Verificar se o valor é uma data válida
        If IsDate(valorCelula) Then
            ' Converter o valor para data
            data = CDate(valorCelula)
            ' Verificar se a data está dentro dos últimos 14 dias
            If data >= dataAtual - 6 And data <= dataAtual Then
                ' Copiar a linha inteira para a nova planilha
                wsDados.Rows(linha).Copy Destination:=wsNova.Rows(proximaLinha)
                proximaLinha = proximaLinha + 1
            End If
        End If
    Next linha
    
    ' Fechar o arquivo sem salvar alterações
    wb.Close SaveChanges:=False
    
    ' Ativar atualização da tela
    Application.ScreenUpdating = True
End Sub


Sub SelecionarArquivo()
    Dim dlgOpen As FileDialog
    Dim EnderecoPlan As String
    
    ' Configuração da caixa de diálogo de arquivo
    Set dlgOpen = Application.FileDialog(msoFileDialogOpen)
    With dlgOpen
        .Title = "Selecione a planilha de controle de cabos"
        .Filters.Clear
        .Filters.Add "Arquivos Excel", "*.xlsx; *.xls"
        .AllowMultiSelect = False
        If .Show = -1 Then ' Se o usuário selecionar um arquivo
            EnderecoPlan = .SelectedItems(1)
            CarregarEntradaCaboUTP EnderecoPlan
        Else ' Se o usuário cancelar a seleção
            MsgBox "Operação cancelada pelo usuário.", vbExclamation, "Selecionar Arquivo"
        End If
    End With
End Sub

