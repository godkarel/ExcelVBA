Sub Imprimir()
    ' Desativa atualização da tela para melhorar o desempenho
    Application.ScreenUpdating = False
    
    ' Chama a função Previa e imprime se o usuário escolher Sim
    If Previa() = vbYes Then
        ' Abre o diálogo de seleção de impressora
        Application.Dialogs(xlDialogPrinterSetup).Show
        
        ' Define a área de impressão
                With Planilha5.PageSetup
            .PrintArea = Planilha5.Range("B1:K57").Address
            .Orientation = xlPortrait ' Define a orientação da página como retrato
        End With
        
        ' Imprime a área definida
        Planilha5.PrintOut Copies:=1, Collate:=True
    End If
    
    ' Ativa a atualização da tela novamente
    Application.ScreenUpdating = True
End Sub

Function Previa() As VbMsgBoxResult
    ' Pergunta ao usuário se deseja imprimir
    Previa = MsgBox("Deseja imprimir o comprovante do técnico selecionado?", vbYesNo)
End Function


