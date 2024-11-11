Sub CopiarColunas(strLocalLivro As String)
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
    Set wsDestino = ThisWorkbook.Sheets("Claro")

    ' Definir última linha
    intUltimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, 4).End(xlUp).Row

    ' Carregar os dados das colunas de origem em arrays
    arrOrigem = wsOrigem.Range("D2:L" & intUltimaLinha).Value

    ' Inicializar o array de destino
    ReDim arrDestino(1 To UBound(arrOrigem, 1), 1 To 6)

    ' Copiar dados para o array de destino
    For i = 1 To UBound(arrOrigem, 1)
        arrDestino(i, 1) = arrOrigem(i, 1) ' Coluna D para Coluna A
        arrDestino(i, 2) = arrOrigem(i, 2) ' Coluna E para Coluna B
        arrDestino(i, 3) = arrOrigem(i, 3) ' Coluna F para Coluna C
        arrDestino(i, 4) = arrOrigem(i, 4) ' Coluna G para Coluna D
        arrDestino(i, 5) = arrOrigem(i, 5) ' Coluna H para Coluna E
        arrDestino(i, 6) = arrOrigem(i, 9) ' Coluna K para Coluna F
    Next i

    ' Colar o array de destino na planilha de destino
    wsDestino.Range("A2").Resize(UBound(arrDestino, 1), UBound(arrDestino, 2)).Value = arrDestino

    ' Fechar o arquivo de origem sem salvar
    wbOrigem.Close False

    ' Reativar atualizações de tela e cálculos automáticos
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub