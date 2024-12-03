Sub ZerarDepoisDeSalvarDiarioKit()
    'MACRO PARA ZERAR TODOS OS VALORES
    'APÓS EXECUTAR O SALVAMENTO PARA DEIXAR A PLANILHA ZERADA.

    Dim LINHADIARIO As Long
    Dim LINHANS As Long

    LINHANS = 4
    
    Application.ScreenUpdating = False

    For LINHADIARIO = 2 To 300
        If Planilha27.Cells(LINHADIARIO, 3) <> "" Then
            Planilha27.Cells(LINHADIARIO, 3) = ""
            Planilha27.Cells(LINHADIARIO, 12) = ""
            LINHANS = LINHANS + 1
        End If
    Next LINHADIARIO
    
    Application.ScreenUpdating = True

    ThisWorkbook.Save
End Sub

