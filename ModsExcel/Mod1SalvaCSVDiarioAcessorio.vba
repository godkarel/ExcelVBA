Sub SalvarCSVAcessorio()
    Dim FileName As String
    Dim NamePcValido1 As String

    ' Copiar a planilha ativa para um novo workbook
    ActiveSheet.Copy
    
    ' Obter o nome do computador
    NamePcValido1 = Environ$("ComputerName")
    
    ' Atualizar células na Planilha16
    With Planilha16
        .Cells(3, 3) = Date
        .Cells(4, 3) = Now
        .Cells(5, 3) = Format(Now, "hh:mm")
    End With
    
    ' Obter o nome do arquivo da célula específica
    FileName = Planilha16.Cells(2, 3)
    
    ' Determinar o nome do arquivo baseado na condição
    If Planilha16.Cells(5, 3) > Planilha16.Cells(6, 3) Then
        FileName = Application.GetSaveAsFilename(InitialFileName:="ACE - " & FileName, _
                                                 FileFilter:="CSV (separado por vírgula) (*.csv), *.csv")
    Else
        Select Case NamePcValido1
            Case "PRSPPE04EFCK"
                FileName = "DIARIAO ACESSORIOS"
            Case "DESKTOP-DBKEF0A"
                FileName = "BAIA 2 ACESSORIOS"
            Case "PRSP41Y0QK2"
                FileName = "BAIA 3 ACESSORIOS"
            Case "DESKTOP-ROMUVAJ", "PRSTPE04EFD5", "PRSPPE04EFCS"
                FileName = "BAIA 0 ACESSORIOS"
        End Select
        FileName = Application.GetSaveAsFilename(InitialFileName:=FileName, _
                                                 FileFilter:="CSV (separado por vírgula) (*.csv), *.csv")
    End If

    ' Desativar alertas, salvar como CSV e fechar o workbook
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs FileName:=FileName, FileFormat:=xlCSV
    ActiveWorkbook.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Sub


