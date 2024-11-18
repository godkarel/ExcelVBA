Sub CopiarLinhasComCriterios(strLocalLivro As String)
    ' Declarar variáveis
    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim arrOrigem As Variant
    Dim arrDestino As Variant
    Dim intUltimaLinha As Long
    Dim i As Long
    Dim j As Long

    ' Desativar atualizações de tela e cálculos automáticos para aumentar a velocidade
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Abrir o arquivo selecionado
    Set wbOrigem = Workbooks.Open(strLocalLivro)

    ' Definir a planilha de origem (ajuste o nome da planilha conforme necessário)
    Set wsOrigem = wbOrigem.Sheets(1) ' Substitua "NomeDaPlanilha" pelo nome da planilha correta
    Set wsDestino = ThisWorkbook.Sheets("Equipamentos")

    ' Definir última linha
    intUltimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, 1).End(xlUp).Row

    ' Carregar os dados da planilha de origem em um array
    arrOrigem = wsOrigem.Range("A2:G" & intUltimaLinha).Value

    ' Inicializar o array de destino
    ReDim arrDestino(1 To UBound(arrOrigem, 1), 1 To UBound(arrOrigem, 2))

    ' Copiar dados com critérios para o array de destino
    j = 1
    For i = 1 To UBound(arrOrigem, 1)
        If arrOrigem(i, 6) <> "" And arrOrigem(i, 7) <> "" Then
            arrDestino(j, 1) = arrOrigem(i, 1)
            arrDestino(j, 2) = arrOrigem(i, 2)
            arrDestino(j, 3) = arrOrigem(i, 3)
            arrDestino(j, 4) = arrOrigem(i, 4)
            arrDestino(j, 5) = arrOrigem(i, 5)
            arrDestino(j, 6) = arrOrigem(i, 6)
            arrDestino(j, 7) = arrOrigem(i, 7)
            j = j + 1
        End If
    Next i

    ' Colar o array de destino na planilha de destino
    wsDestino.Range("B2").Resize(j - 1, UBound(arrDestino, 2)).Value = arrDestino

    ' Fechar o arquivo de origem sem salvar
    wbOrigem.Close False

    ' Reativar atualizações de tela e cálculos automáticos
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub