
Sub AplicarFormulaProcurados()
    Dim ws As Worksheet
    Dim i As Integer
    
    ' Define a planilha "PROCURADOS"
    Set ws = ThisWorkbook.Sheets("PROCURADOS")
    
    ' Aplica a fórmula em C2 a C30
    For i = 2 To 30
        ws.Cells(i, 1).Formula = "=CONTROLEUTP!C" & i
        ws.Cells(i, 2).Formula = "=CONTROLEUTP!E" & i
    Next i
    
    
End Sub

