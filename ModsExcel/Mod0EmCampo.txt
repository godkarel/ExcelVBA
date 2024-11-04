Sub ImportarDadosDoAtlasEmCampo()
    Dim ws As Worksheet
    Dim wbImport As Workbook
    Dim filePath As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim destRow As Long

    ' Desativa a atualização da tela e o cálculo automático para melhorar a performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Seleciona a planilha "FabioMamado"
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("FabioMamado")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "A aba 'FabioMamado' não existe.", vbExclamation
        ' Restaura as configurações antes de sair
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Exit Sub
    End If
    
    ' Limpa o intervalo da célula B2 até R10000 na planilha de destino
    ws.Range("B2:R10000").ClearContents
    
    ' Abre a caixa de diálogo para selecionar o arquivo
    With Application.FileDialog(msoFileDialogOpen)
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx", 1
        .Title = "Selecione o arquivo Excel"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            MsgBox "Nenhum arquivo selecionado.", vbExclamation
            ' Restaura as configurações antes de sair
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
            Exit Sub
        End If
    End With

    ' Abre o arquivo selecionado
    Set wbImport = Workbooks.Open(filePath)

    ' Define a última linha e última coluna da planilha de importação
    With wbImport.Sheets(1)
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
    End With
    
    ' Inicializa a linha de destino na segunda linha
    destRow = 2
    
    ' Copia os dados do arquivo importado para a planilha "FabioMamado"
    ' Deslocando uma coluna para a direita e começando na segunda linha
    For i = 2 To lastRow
        If wbImport.Sheets(1).Cells(i, 6).Value = "INICIALIZADO" And _
           wbImport.Sheets(1).Cells(i, 7).Value = "PROCISA DO BRASIL PROJETOS CONSTRUC" Then
            For j = 1 To lastCol
                ws.Cells(destRow, j + 1).Value = wbImport.Sheets(1).Cells(i, j).Value
            Next j
            destRow = destRow + 1
        End If
    Next i
    
    ' Fecha o arquivo importado sem salvar
    wbImport.Close SaveChanges:=False
    
    Sheets("FabioMamado").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("B2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=True, OtherChar:= _
        "/", FieldInfo:=Array(Array(1, 9), Array(2, 2), Array(3, 9)), _
        TrailingMinusNumbers:=True
    Range("B1").Select

    ' Restaura as configurações
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Dados importados com sucesso!", vbInformation
End Sub



