Sub SequenciaUTP()
    ' Desativa o cálculo automático e a atualização da tela
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    
    DeletaAntesUTPProcurados
    SelecionarArquivo
    CarregarSaidaCaboUTP
    FiltrarCaboUTP
    OrdenarData
    ExcluirDevolucoes
    DeletarLinhasPorValor
    AplicarFormulaProcurados
    
    ' Reativa o cálculo automático e a atualização da tela
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    'If Err.Number <> 0 Then
        ' Define a cor da fonte como vermelho na célula Planilha7.Cells(7, 16)
        'Planilha7.Cells(7, 16).Font.Color = RGB(255, 0, 0)
        'MsgBox "Ocorreu um erro durante a sincronização dos cabos UTPs. Verifique e tente novamente.", vbExclamation
        'Planilha7.Select
        'Exit Sub
    'End If

    Planilha7.Select
    Planilha7.Cells(7, 16).Font.ColorIndex = xlAutomatic ' Restaura a cor padrão da fonte
    
    Dim ws As Worksheet
    ' Definir a planilha de destino
    Set ws = ThisWorkbook.Sheets("PAGINA INICIAL")

    ' Inserir a data atual na célula P11
    ws.Range("P7").Value = Date
    
    ThisWorkbook.Save
End Sub

Sub ConfirmarAcaoATUALIZAUTP()
    Dim resposta As VbMsgBoxResult
    
    ' Exibe uma mensagem perguntando se o usuário tem certeza que quer continuar
    resposta = MsgBox("DESEJA SINCRONIZAR O UTP ?", vbYesNo + vbQuestion, "Confirmação")

    ' Verifica a resposta do usuário
    If resposta = vbYes Then
        Call SequenciaUTP
        
    Else
        Exit Sub
    End If
End Sub

