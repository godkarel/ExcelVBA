Sub SalvarCSVDIARIAOKIT()
    Dim FileName As String

    ' Selecionar a Planilha27 e copiar para um novo workbook
    Planilha27.Select
    ActiveSheet.Copy
    
    ' Definir o nome do arquivo com a caixa de diálogo Salvar Como
    FileName = Application.GetSaveAsFilename(InitialFileName:="Diario de Kit de Sei la oque", _
                                             FileFilter:="CSV (separado por vírgula) (*.csv), *.csv")
    
    ' Desativar alertas, salvar como CSV e fechar o workbook
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs FileName:=FileName, FileFormat:=xlCSV
    ActiveWorkbook.Close SaveChanges:=False
    Application.DisplayAlerts = True
    
    ' Chamar a subrotina ZerarDepoisDeSalvarDiarioKit
    ZerarDepoisDeSalvarDiarioKit
    
    SelecionaPaginaInicial
End Sub



