Sub CombineCSVFiles()
    Dim ws As Worksheet
    Dim wsNew As Worksheet
    Dim fDialog As FileDialog
    Dim FileChosen As Integer
    Dim FileName As Variant
    Dim lastRow As Long
    Dim NewLastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    Dim wb As Workbook
    Dim wbNew As Workbook
    Dim ultimaLinha As Long
    
    Dim rng As Range
    Dim celula As Range
    Dim dict As Object
    Dim totalDiferentes As Long
    Dim horaAtual As String

    ' Create a new workbook
    Set wbNew = Workbooks.Add
    Set wsNew = wbNew.Sheets(1)
    wsNew.Name = "Combinado"

    ' Set up file dialog
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .Title = "Select CSV Files"
        FileChosen = .Show

        If FileChosen <> -1 Then Exit Sub ' User canceled

        ' Loop through each selected file
        For Each FileName In .SelectedItems
            ' Open the file
            Set wb = Workbooks.Open(FileName)
            Set ws = wb.Sheets(1)
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
            ' Loop through the rows starting from the second row
            For i = 2 To lastRow
                ' Verificar se a célula contém um valor não vazio e não é um erro
                If Not IsError(ws.Cells(i, 1).Value) And Trim(CStr(ws.Cells(i, 1).Value)) <> "" Then
                    NewLastRow = wsNew.Cells(wsNew.Rows.Count, "A").End(xlUp).Row + 1
                    ' Se a célula A1 da nova planilha estiver vazia, não adicionar +1
                    If wsNew.Cells(1, 1).Value = "" Then NewLastRow = 1
                    ' Loop through the columns and copy the data as text
                    For j = 1 To lastCol
                        ' Tratando especificamente a coluna C como texto
                        If j = 3 Then
                            wsNew.Cells(NewLastRow, j).Value = "'" & ws.Cells(i, j).Value
                        Else
                            wsNew.Cells(NewLastRow, j).Value = ws.Cells(i, j).Text
                        End If
                    Next j
                End If
            Next i
        
            ' Close the file without saving
            wb.Close SaveChanges:=False
        Next FileName

    End With
    
    ' Formatar a coluna C como texto
    wsNew.Columns("C").NumberFormat = "@"

    ' Ajustar o formato das colunas A, H e K
    wsNew.Columns("A").NumberFormat = "00000000"
    wsNew.Columns("H").NumberFormat = "000"
    wsNew.Columns("K").NumberFormat = "000"
    
    ' Inserir linha de cabeçalho
    wsNew.Rows(1).Insert Shift:=xlDown
    wsNew.Range("A1").Value = "Nº do item"
    wsNew.Range("B1").Value = "Nº do lote"
    wsNew.Range("C1").Value = "Nº de série"
    wsNew.Range("D1").Value = "Dimensão 1"
    wsNew.Range("E1").Value = "Dimensão 2"
    wsNew.Range("F1").Value = "Quantidade"
    wsNew.Range("G1").Value = "Site"
    wsNew.Range("H1").Value = "Depósito"
    wsNew.Range("I1").Value = "Localização"
    wsNew.Range("J1").Value = "Até Site"
    wsNew.Range("K1").Value = "Até Depósito"
    wsNew.Range("L1").Value = "Até Localização"
    wsNew.Range("M1").Value = ""

    ' Criar um Dictionary para armazenar valores únicos
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Definir o intervalo da coluna L com células preenchidas
    Set rng = wsNew.Range("L1:L" & wsNew.Cells(wsNew.Rows.Count, "L").End(xlUp).Row)
    
    ' Percorrer as células do intervalo
    For Each celula In rng
        If Not IsEmpty(celula.Value) Then
            ' Adicionar valor ao Dictionary se não estiver presente
            If Not dict.exists(celula.Value) Then
                dict.Add celula.Value, Nothing
            End If
        End If
    Next celula
    
    ' Contar o número de itens distintos
    totalDiferentes = dict.Count
    
    ' Exibir o resultado
    horaAtual = Time
    Planilha7.Cells(13, 16) = "Tec: " & totalDiferentes - 1 & "|" & horaAtual
    
    ' Salvar o novo workbook como CSV
    FileName = "DIARIO UNIFICADO"
    FileName = Application.GetSaveAsFilename(InitialFileName:=FileName, _
                                             FileFilter:="CSV (separado por vírgula) (*.csv), *.csv")
    If FileName <> "False" Then
        Application.DisplayAlerts = False
        wbNew.SaveAs FileName:=FileName, FileFormat:=xlCSV
        wbNew.Close SaveChanges:=False
        Application.DisplayAlerts = True
        MsgBox "CSV files combined and saved successfully as " & FileName & "!"
    Else
        wbNew.Close SaveChanges:=False
        MsgBox "Operation cancelled. No file was saved."
    End If

End Sub



