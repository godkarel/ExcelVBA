Sub Loop_REDINR()
    Dim LINHADIARIO As Long
    Dim LINHANS As Long

    LINHANS = 4

    ' Desativar atualização de tela para melhorar o desempenho
    Application.ScreenUpdating = False

    For LINHADIARIO = 3 To 700
        If Planilha20.Cells(LINHADIARIO, 1) = "" Then
            Planilha20.Cells(LINHADIARIO, 3) = Planilha19.Cells(LINHANS, 2)
            Planilha20.Cells(LINHADIARIO, 1) = Planilha19.Cells(LINHANS, 4)

            If Planilha20.Cells(LINHADIARIO, 3) <> "" Then
                Planilha20.Cells(LINHADIARIO, 10) = Planilha19.Cells(1, 4)
            End If

            LINHANS = LINHANS + 1
        End If
    Next LINHADIARIO

    ' Chama subrotina ZeraDiarioPraNovoAtendimentoDINR
    ZeraDiarioPraNovoAtendimentoDINR

    ' Salva a pasta de trabalho
    ThisWorkbook.Save

    ' Reativar atualização de tela
    Application.ScreenUpdating = True
End Sub



