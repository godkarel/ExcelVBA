Sub SalvarLogAbastecimento()
    Dim ws As Worksheet
    Dim logFile As String
    Dim i As Long
    Dim dataLine As String
    Dim fileNumber As Integer
    
    ' Defina a planilha onde os dados estão localizados
    Set ws = ThisWorkbook.Sheets("DIARIO CARGA")
    
    ' Caminho do arquivo de log
    logFile = Environ("USERPROFILE") & "\Documents\LOGABASTECIMENTO.txt"
    
    ' Obtenha um número de arquivo
    fileNumber = FreeFile
    
    ' Abra o arquivo de log para adicionar (Append)
    Open logFile For Append As #fileNumber
    
    ' Loop para percorrer todas as linhas até o último preenchido
    For i = 2 To ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
        ' Verifique se a célula na coluna C não está vazia
        If ws.Cells(i, "C").Value <> "" Then
            ' Formate a linha com o dado da coluna C e o correspondente na coluna L
            dataLine = ws.Cells(i, "C").Value & " - " & ws.Cells(i, "L").Value
            ' Escreva a linha no arquivo de log
            Print #fileNumber, dataLine
        End If
    Next i
    
    ' Feche o arquivo
    Close #fileNumber
End Sub

Sub SalvarLogDevolution()
    Dim ws As Worksheet
    Dim logFile As String
    Dim i As Long
    Dim dataLine As String
    Dim fileNumber As Integer
    
    ' Defina a planilha onde os dados estão localizados
    Set ws = ThisWorkbook.Sheets("DIARIO DEVOLUÇÃO")
    
    ' Caminho do arquivo de log
    logFile = Environ("USERPROFILE") & "\Documents\LOGDEVOLUÇÃO.txt"
    
    ' Obtenha um número de arquivo
    fileNumber = FreeFile
    
    ' Abra o arquivo de log para adicionar (Append)
    Open logFile For Append As #fileNumber
    
    ' Loop para percorrer todas as linhas até o último preenchido
    For i = 2 To ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
        ' Verifique se a célula na coluna C não está vazia
        If ws.Cells(i, "C").Value <> "" Then
            ' Formate a linha com o dado da coluna C e o correspondente na coluna L
            dataLine = ws.Cells(i, "C").Value & " - " & ws.Cells(i, "I").Value
            ' Escreva a linha no arquivo de log
            Print #fileNumber, dataLine
        End If
    Next i
    
    ' Feche o arquivo
    Close #fileNumber
End Sub