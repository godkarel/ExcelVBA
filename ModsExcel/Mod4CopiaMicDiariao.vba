Sub CopiarColunasMisc(strLocalLivro As String)
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
    Set wsOrigem = wbOrigem.Sheets("SUBIR") ' Substitua "NomeDaPlanilha" pelo nome da planilha correta
    Set wsDestino = ThisWorkbook.Sheets("Diario Mic")

    ' Definir última linha
    intUltimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, 4).End(xlUp).Row

    ' Carregar os dados das colunas de origem em arrays
    arrOrigem = wsOrigem.Range("J2:R" & intUltimaLinha).Value

    ' Inicializar o array de destino
    ReDim arrDestino(1 To UBound(arrOrigem, 1), 1 To 12)

    ' Copiar dados para o array de destino
    For i = 1 To UBound(arrOrigem, 1)
        arrDestino(i, 6) = arrOrigem(i, 1) ' COPIA QUATIDADE para coluna F
        arrDestino(i, 12) = arrOrigem(i, 6) ' COPIA NOME DO OTARIO para coluna L
        arrDestino(i, 1) = arrOrigem(i, 9) ' COPIA CODAX para coluna A
    Next i

    ' Colar o array de destino na planilha de destino sem sobrescrever as colunas H, I, K
    ' Copia apenas colunas A, F, L
    wsDestino.Range("A2").Resize(UBound(arrDestino, 1), 1).Value = Application.Index(arrDestino, 0, 1) ' Coluna A
    wsDestino.Range("F2").Resize(UBound(arrDestino, 1), 1).Value = Application.Index(arrDestino, 0, 6) ' Coluna F
    wsDestino.Range("L2").Resize(UBound(arrDestino, 1), 1).Value = Application.Index(arrDestino, 0, 12) ' Coluna L

    ' Fechar o arquivo de origem sem salvar
    wbOrigem.Close False

    ' Reativar atualizações de tela e cálculos automáticos
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub



