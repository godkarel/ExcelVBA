Sub EmCampo()
    Dim wsAutonomia As Worksheet
    Dim wsClaro As Worksheet
    Dim criterioA2 As String
    Dim ultimaLinhaClaro As Long
    Dim contador As Long
    Dim i As Long
    Dim j As Long
    
    
    Application.ScreenUpdating = False
    
    ' Definir as planilhas
    Set wsAutonomia = ThisWorkbook.Sheets("AUTONOMIA EM CAMPO")
    Set wsClaro = ThisWorkbook.Sheets("CLARO")
    
    ' Encontrar a última linha usada na planilha "CLARO"
    ultimaLinhaClaro = wsClaro.Cells(wsClaro.Rows.Count, "A").End(xlUp).Row
    
    ' Loop através das células A2 até A26 na planilha "AUTONOMIA EM CAMPO"
    For j = 2 To 26
        ' Obter o valor de A2 até A26 na planilha "AUTONOMIA EM CAMPO"
        criterioA2 = wsAutonomia.Cells(j, 1).Value
        
        ' Inicializar o contador
        contador = 0
        
        ' Loop através das linhas da planilha "CLARO"
        For i = 2 To ultimaLinhaClaro
            ' Verificar se os critérios são atendidos
            If wsClaro.Cells(i, "F") = "INICIALIZADO" Then
                If Not IsError(wsClaro.Cells(i, "G")) Then
                    If (Right(wsClaro.Cells(i, "G"), 3) = "EST" Or _
                        Right(wsClaro.Cells(i, "G"), 3) = "EDT") And _
                        wsClaro.Cells(i, "I") = criterioA2 Then
                        
                        ' Incrementar o contador
                        contador = contador + 1
                    End If
                End If
            End If
        Next i
        
        ' Colocar o resultado em B2 até B26 na planilha "AUTONOMIA EM CAMPO"
        wsAutonomia.Cells(j, 2).Value = contador
    Next j
    
    Application.ScreenUpdating = True
    
End Sub


