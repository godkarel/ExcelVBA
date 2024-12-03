Sub CopiaColunasKitSeiLa(strLocalLivro As String)
    ' Declarar variáveis
    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim arrOrigem As Variant
    Dim arrDestino As Variant
    Dim intUltimaLinha As Long
    Dim i As Long

    ' Desativar atualizações de tela e cálculos automáticos para aumentar a velocidade
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Abrir o arquivo selecionado
    Set wbOrigem = Workbooks.Open(strLocalLivro)

    ' Definir a planilha de origem (ajuste o nome da planilha conforme necessário)
    Set wsOrigem = wbOrigem.Sheets(1) ' Substitua "NomeDaPlanilha" pelo nome da planilha correta
    Set wsDestino = ThisWorkbook.Sheets("KIT DE SEI LA")

    ' Carregar os dados das colunas de origem em um array
    arrOrigem = wsOrigem.Range("C4:C36").Value

    ' Copiar dados para a planilha de destino
    wsDestino.Range("C4:C36").Value = arrOrigem

    ' Fechar o arquivo de origem sem salvar
    wbOrigem.Close False

    ' Reativar atualizações de tela e cálculos automáticos
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
