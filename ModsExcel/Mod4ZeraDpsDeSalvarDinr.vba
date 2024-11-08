Sub ZerarDepoisDeSalvarDINR()
    'MACRO PARA ZERAR TODOS OS VALORES
    'APÓS EXECUTAR O SALVAMENTO PARA DEIXAR A PLANILHA ZERADA.

    Dim LINHADIARIO As Long
    Dim LINHANS As Long

    LINHANS = 4
    
    Application.ScreenUpdating = False

    For LINHADIARIO = 2 To 1000
        If Planilha20.Cells(LINHADIARIO, 3) <> "" Then
            Planilha20.Cells(LINHADIARIO, 3) = ""
            Planilha20.Cells(LINHADIARIO, 1) = ""
            Planilha20.Cells(LINHADIARIO, 10) = ""
            LINHANS = LINHANS + 1
        End If
    Next LINHADIARIO
    
    Application.ScreenUpdating = True

    ThisWorkbook.Save
End Sub

