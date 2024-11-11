Sub ZerarDepoisDeSalvarDEV()
    'MACRO PARA ZERAR TODOS OS VALORES
    'APÓS EXECUTAR O SALVAMENTO PARA DEIXAR A PLANILHA ZERADA.

    Dim LINHADIARIO As Long
    Dim LINHANS As Long

    LINHANS = 2
    
    Application.ScreenUpdating = False
    
    For LINHADIARIO = 2 To 1000
        If Planilha11.Cells(LINHADIARIO, 3) <> "" Then
            Planilha11.Cells(LINHADIARIO, 3) = ""
            Planilha11.Cells(LINHADIARIO, 1) = ""
            Planilha11.Cells(LINHADIARIO, 12) = ""
            LINHANS = LINHANS + 1
        End If
    Next LINHADIARIO
    
    Application.ScreenUpdating = True

    ThisWorkbook.Save
End Sub
