Sub ProcurarTextoEmArquivos()
    Dim FSO As Object
    Dim FD As FileDialog
    Dim caminhoPasta As String
    Dim txtProcurado As String
    Dim resultado As String
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    
    ' Texto a ser procurado
    txtProcurado = InputBox("Digite o texto a ser procurado:", "Procurar Texto")
    
    ' Selecionar a pasta
    Set FD = Application.FileDialog(msoFileDialogFolderPicker)
    FD.Title = "Selecione a pasta"
    
    If FD.Show = -1 Then
        caminhoPasta = FD.SelectedItems(1)
    Else
        MsgBox "Nenhuma pasta selecionada.", vbExclamation
        Exit Sub
    End If
    
    ' Inicializar o FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Configurar a planilha
    Set ws = ThisWorkbook.Sheets("BOT DE PESQUISA")
    
    ' Limpar coluna A a partir da célula A2
    ws.Range("A2:A20" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).ClearContents
    ws.Range("B2:B20" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row).ClearContents
    
    ultimaLinha = 2
    
    ' Procurar texto nos arquivos na pasta selecionada
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Call ProcurarTextoEmPasta(FSO.GetFolder(caminhoPasta), txtProcurado, ws, ultimaLinha)
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Exibir mensagem final
    If ultimaLinha = 2 Then
        MsgBox "Texto '" & txtProcurado & "' não encontrado em nenhum arquivo na pasta selecionada.", vbExclamation
    Else
        CriarHyperlinks
        MsgBox "Pesquisa concluída. Resultados salvos na planilha BOT DE PESQUISA.", vbInformation
    End If
End Sub

Sub ProcurarTextoEmPasta(pasta As Object, txtProcurado As String, ws As Worksheet, ByRef ultimaLinha As Long)
    Dim arquivo As Object
    Dim subPasta As Object
    Dim ts As Object
    Dim wb As Workbook
    Dim wsTemp As Worksheet
    Dim cell As Range
    Dim linha As String
    Dim extension As String
    Dim FSO As Object
    
    ' Inicializar o FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Processar arquivos na pasta atual
    For Each arquivo In pasta.Files
        extension = LCase(FSO.GetExtensionName(arquivo.Name))
        Select Case extension
            Case "txt"
                ' Processar arquivos de texto
                Set ts = arquivo.OpenAsTextStream(1, -2) ' 1 = ForReading, -2 = TristateUseDefault
                Do While Not ts.AtEndOfStream
                    linha = ts.ReadLine
                    If InStr(1, linha, txtProcurado, vbTextCompare) > 0 Then
                        ws.Cells(ultimaLinha, 1).Value = arquivo.Path
                        ultimaLinha = ultimaLinha + 1
                        Exit Do
                    End If
                Loop
                ts.Close
            Case "csv"
                ' Processar arquivos CSV
                Set wb = Workbooks.Open(arquivo.Path, ReadOnly:=True)
                For Each wsTemp In wb.Worksheets
                    For Each cell In wsTemp.UsedRange
                        If Not IsError(cell.Value) And Not IsEmpty(cell.Value) Then
                            If InStr(1, CStr(cell.Value), txtProcurado, vbTextCompare) > 0 Then
                                ws.Cells(ultimaLinha, 1).Value = arquivo.Path
                                ultimaLinha = ultimaLinha + 1
                                Exit For
                            End If
                        End If
                    Next cell
                    If InStr(ws.Cells(ultimaLinha - 1, 1).Value, arquivo.Path) > 0 Then Exit For
                Next wsTemp
                wb.Close False
            Case "xlsx", "xls"
                ' Processar arquivos Excel
                Set wb = Workbooks.Open(arquivo.Path, ReadOnly:=True)
                For Each wsTemp In wb.Worksheets
                    For Each cell In wsTemp.UsedRange
                        If Not IsError(cell.Value) And Not IsEmpty(cell.Value) Then
                            If InStr(1, CStr(cell.Value), txtProcurado, vbTextCompare) > 0 Then
                                ws.Cells(ultimaLinha, 1).Value = arquivo.Path
                                ultimaLinha = ultimaLinha + 1
                                Exit For
                            End If
                        End If
                    Next cell
                    If InStr(ws.Cells(ultimaLinha - 1, 1).Value, arquivo.Path) > 0 Then Exit For
                Next wsTemp
                wb.Close False
        End Select
    Next
    
    ' Processar subpastas
    For Each subPasta In pasta.SubFolders
        Call ProcurarTextoEmPasta(subPasta, txtProcurado, ws, ultimaLinha)
    Next subPasta
    
    
End Sub
