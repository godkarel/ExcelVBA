Sub CopiarDadosParaDiarioKit()
    Dim wsFonte As Worksheet
    Dim wsDestino As Worksheet
    Dim dadosA As Variant
    Dim dadosE As Variant
    Dim i As Long
    
    Application.ScreenUpdating = False
    
    ' Definindo as planilhas
    Set wsFonte = ThisWorkbook.Sheets("BIPAGEM DO KIT")
    Set wsDestino = ThisWorkbook.Sheets("DIARIO KIT")

    ' Lendo os dados da coluna A e E da planilha fonte para arrays
    dadosA = wsFonte.Range("A2:A34").Value
    dadosE = wsFonte.Range("E2:E34").Value

    ' Copiando os dados do array para a planilha destino
    For i = 1 To UBound(dadosA, 1)
        wsDestino.Cells(i + 1, 3).Value = dadosA(i, 1) ' Coluna C
        wsDestino.Cells(i + 1, 12).Value = dadosE(i, 1) ' Coluna L
    Next i
    ZeraDiarioPraNovoKit
    Application.ScreenUpdating = True
    SelecionaDiarioKit
End Sub
