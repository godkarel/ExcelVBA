Sub SalvarCSVMIC()
    Dim FileName As String
    Dim NamePcValido1 As String

    ' Copiar a planilha ativa para um novo workbook (ou seja pega essa merda)
    ActiveSheet.Copy
    
    ' Obter o nome do computador ( valida )
    NamePcValido1 = Environ$("ComputerName")
    
    ' Atualizar células na Planilha16 ( tem que ser a 16 )
    With Planilha16
        .Cells(3, 3) = Date
        .Cells(4, 3) = Now
        .Cells(5, 3) = Format(Now, "hh:mm")
    End With
    
    ' Obter o nome do arquivo da célula específica nesse caralho
    FileName = Planilha16.Cells(2, 3)
    
    ' Determinar o nome do arquivo baseado na condição ( CONDIÇÃO DE MAMADA )
    If Planilha16.Cells(5, 3) > Planilha16.Cells(6, 3) Then
        FileName = Application.GetSaveAsFilename(InitialFileName:="MIC " & FileName, _
                                                 FileFilter:="CSV (separado por vírgula) (*.csv), *.csv")
    Else
        Select Case NamePcValido1
            Case "PRSPPE04EFCK", "DESKTOP-DBKEF0A", "PRSP41Y0QK2", "DESKTOP-ROMUVAJ", "PRSTPE04EFD5", "PRSPPE04EFCS"
                FileName = "MIC TODOS"
        End Select
        FileName = Application.GetSaveAsFilename(InitialFileName:=FileName, _
                                                 FileFilter:="CSV (separado por vírgula) (*.csv), *.csv")
    End If

    ' Desativar alertas, salvar como CSV e fechar o workbook ( N ESQUECE DE DESATIVAR )
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs FileName:=FileName, FileFormat:=xlCSV
    ActiveWorkbook.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Sub
