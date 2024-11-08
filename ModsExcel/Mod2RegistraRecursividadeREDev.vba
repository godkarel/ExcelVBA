Sub Loop_REDEV()
    Dim LINHADIARIO As Long
    Dim LINHANS As Long

    LINHANS = 4

    ' Desativar atualização de tela para melhorar o desempenho
    Application.ScreenUpdating = False

    For LINHADIARIO = 3 To 1000
        If Planilha1.Cells(LINHADIARIO, 1) = "" Then
            Planilha1.Cells(LINHADIARIO, 3) = Planilha3.Cells(LINHANS, 2)
            Planilha1.Cells(LINHADIARIO, 1) = Planilha3.Cells(LINHANS, 4)

            If Planilha1.Cells(LINHADIARIO, 3) <> "" Then
                Planilha1.Cells(LINHADIARIO, 12) = "001AB"
            End If

            LINHANS = LINHANS + 1
        End If
    Next LINHADIARIO

    ' Chama subrotina ImprimirDEV
    ImprimirDEV

    ' Chama subrotina ZeraDiarioPraNovoAtendimentoDEV
    ZeraDiarioPraNovoAtendimentoDEV

    ' Salva a pasta de trabalho
    ThisWorkbook.Save

    ' Chama subrotina SelecionaPaginaInicial
    SelecionaPaginaInicial

    ' Reativar atualização de tela
    Application.ScreenUpdating = True
End Sub


