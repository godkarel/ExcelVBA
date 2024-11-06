Sub Loop_NomeDEV()
    Dim LINHADIARIO As Long
    Dim LINHANS As Long
    Dim wsDiario As Worksheet
    Dim wsNS As Worksheet
    Dim wsAux As Worksheet
    Dim ultimaLinhaNS As Long
    
    ' Define as planilhas de trabalho
    Set wsDiario = Planilha11 ' Defina o nome correto da Planilha11
    Set wsNS = Planilha9 ' Defina o nome correto da Planilha9
    Set wsAux = Planilha23 ' Defina o nome correto da Planilha23
    
    ' Inicializa a linha de NS
    LINHANS = 4
    
    ' Determina a última linha com dados na Planilha9
    ultimaLinhaNS = wsNS.Cells(wsNS.Rows.Count, 2).End(xlUp).Row
    
    ' Loop pelas linhas da Planilha11
    For LINHADIARIO = 1 To 1000
        ' Verifica se a célula na coluna 3 está vazia
        If wsDiario.Cells(LINHADIARIO, 3).Value = "" Then
            ' Preenche as informações na coluna 3 da Planilha11 com base na Planilha9
            wsDiario.Cells(LINHADIARIO, 3).Value = wsNS.Cells(LINHANS, 2).Value
            
            ' Verifica se a coluna 3 da Planilha11 não está vazia
            If wsDiario.Cells(LINHADIARIO, 3).Value <> "" Then
                wsDiario.Cells(LINHADIARIO, 1).Value = wsNS.Cells(LINHANS, 4).Value
            End If
            
            ' Preenche as informações na Planilha23 com base na Planilha9
            wsAux.Cells(LINHADIARIO, 1).Value = wsNS.Cells(LINHANS, 2).Value
            
            ' Verifica se a coluna 1 da Planilha23 não está vazia e preenche a coluna 2 com base na Planilha9
            If wsAux.Cells(LINHADIARIO, 1).Value <> "" Then
                wsAux.Cells(LINHADIARIO, 2).Value = wsNS.Cells(1, 3).Value
            End If
            
            ' Verifica se a coluna 3 da Planilha11 não está vazia e preenche a coluna 12 com base na condição
            If wsDiario.Cells(LINHADIARIO, 3).Value <> "" Then
                If wsNS.Cells(1, 3).Value = "BOM PRA USO" Then
                    wsDiario.Cells(LINHADIARIO, 12).Value = "001AA"
                Else
                    wsDiario.Cells(LINHADIARIO, 12).Value = "001AB"
                End If
            End If
            
            ' Incrementa a linha da Planilha9
            LINHANS = LINHANS + 1
            
            ' Verifica se ultrapassou a última linha da Planilha9 e reinicia
            If LINHANS > ultimaLinhaNS Then
                Exit For ' Sai do loop se não houver mais dados em Planilha9
            End If
        End If
    Next LINHADIARIO
    
    ' Imprime a Planilha11
    ImprimirDEV
    
    ' Zera o diário de carga para novo atendimento
    ZeraDiarioPraNovoAtendimentoDEV
    
    ' Salva o arquivo
    ThisWorkbook.Save
    
    ' Seleciona a página inicial
    SelecionaPaginaInicial
End Sub




