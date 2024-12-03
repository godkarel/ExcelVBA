Sub SalvarKitSeiLaImpressao()
    Dim FileName As String

    ' Selecionar a Planilha28 e copiar para um novo workbook
    Planilha28.Select
    ActiveSheet.Copy
    
    ' Definir o nome do arquivo com a caixa de diálogo Salvar Como
    FileName = Application.GetSaveAsFilename(InitialFileName:="Diario de Kit de Sei la oque", _
                                             FileFilter:="Arquivo do Excel (*.xlsx), *.xlsx")
    
    ' Desativar alertas, salvar como .xlsx e fechar o workbook
    If FileName <> "Falso" Then ' Verifica se o usuário não cancelou o Save As
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs FileName:=FileName, FileFormat:=xlOpenXMLWorkbook
        ActiveWorkbook.Close SaveChanges:=False
        Application.DisplayAlerts = True
    End If
    
    ZerarParaNovoKitSeiLa
    SelecionaPaginaInicial
End Sub

