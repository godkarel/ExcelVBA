Sub CarregarSaidaCaboUTP()
    Dim wb As Workbook
    Dim wsAtual As Worksheet
    Dim wsNova As Worksheet
    Dim wsDev As Worksheet
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
    Set wb = Workbooks.Open("C:\Users\r.santos12.CONDUMEX\OneDrive - Grupo Carso\Meu Drive\almox\Controle de Estoque\Controle de Cabos\ControlCaboUTP.xlsx")
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
    
    ' Acessa a aba Dev na planilha ControlCaboUTP
    Set wsDev = wb.Sheets("Dev")
    
    ' Encontrar a última linha na aba CONTROLEUTP
    If abaExistente Then
        ultimaLinhaNova = wsNova.Cells(wsNova.Rows.Count, "A").End(xlUp).Row
        proximaLinha = ultimaLinhaNova + 1 ' Próxima linha na aba CONTROLEUTP
    Else
        proximaLinha = 1 ' Próxima linha na nova planilha
    End If
    
    ' Encontrar a última linha na aba Dev
    ultimaLinha = wsDev.Cells(wsDev.Rows.Count, "A").End(xlUp).Row
    
    ' Copiar os dados da aba Dev para a aba CONTROLEUTP
    For linha = 1 To ultimaLinha
        ' Obter o valor da célula na coluna F da aba Dev
        valorCelula = wsDev.Cells(linha, "F").Value
        ' Verificar se o valor é uma data válida
        If IsDate(valorCelula) Then
            ' Converter o valor para data
            data = CDate(valorCelula)
            ' Verificar se a data está dentro dos últimos 2 dias
            If data >= dataAtual - 7 And data <= dataAtual Then
                ' Copiar a linha inteira para a nova planilha
                wsDev.Rows(linha).Copy Destination:=wsNova.Rows(proximaLinha)
                proximaLinha = proximaLinha + 1
            End If
        End If
    Next linha
    
    ' Fechar o arquivo sem salvar alterações
    wb.Close SaveChanges:=False
    
    ' Ativar atualização da tela
    Application.ScreenUpdating = True
    ThisWorkbook.ActiveSheet.Range("A1").Select
    
End Sub

