Sub ImprimirComprovante()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Comprovante de Entrega")

    ws.Range("A1:F47").PrintOut Copies:=1, Collate:=True
End Sub
