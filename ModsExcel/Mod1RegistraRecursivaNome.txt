Sub Loop_Nome()
    Dim LINHADIARIO As Long
    Dim LINHANS As Long
    Dim wsDiario As Worksheet
    Dim wsDiarioAcessorio As Worksheet
    Dim wsNS As Worksheet
    Dim wsDBAcessorio As Worksheet
    Dim ultimaLinhaNS As Long
    Dim ultimaLinhaDiarioAcessorio As Long
    Dim linhaAcessorio As Long
    Dim valorNS As String
    Dim achouItem As Range
    
    ' Define as planilhas de trabalho
    Set wsDiario = Planilha1 ' Define o nome correto da Planilha1
    Set wsNS = Planilha3 ' Define o nome correto da Planilha3
    Set wsDiarioAcessorio = Planilha34 ' Define o nome correto da Planilha34
    Set wsDBAcessorio = Planilha35 ' Define o nome correto da Planilha35 (DB Acessorio)
    
    ' Inicializa a linha de NS
    LINHANS = 4
    
    ' Determina a última linha com dados na Planilha3
    ultimaLinhaNS = wsNS.Cells(wsNS.Rows.Count, 2).End(xlUp).Row
    
    ' Loop pelas linhas da Planilha1
    For LINHADIARIO = 2 To 700
        ' Verifica se a célula na coluna 3 está vazia
        If wsDiario.Cells(LINHADIARIO, 3) = "" Then
        
            ' Preenche as informações nas colunas 3 e 1 da Planilha1 com base na Planilha3
            wsDiario.Cells(LINHADIARIO, 3) = wsNS.Cells(LINHANS, 2)
            wsDiario.Cells(LINHADIARIO, 1) = wsNS.Cells(LINHANS, 4)
            
            ' Verifica se a coluna 3 da Planilha1 não está vazia e preenche a coluna 12 com base na Planilha3
            If wsDiario.Cells(LINHADIARIO, 3) <> "" Then
                wsDiario.Cells(LINHADIARIO, 12) = wsNS.Cells(1, 4)
            End If
            
           On Error Resume Next ' Ignora erros
            
            ' Verifica se o valor de wsNS.Cells(Linhas, 4) está na coluna A da Planilha34
            valorNS = wsNS.Cells(LINHANS, 4)
            Set achouItem = wsDBAcessorio.Columns(1).Find(valorNS, LookIn:=xlValues, LookAt:=xlWhole)
            
            ' Retorna ao comportamento normal do erro depois da busca
            On Error GoTo 0 ' Volta a tratar os erros normalmente
            
             '
             ' CODIGO COMENTADO POIS ESTAVA ATRAPALHANDO A DESTRIBUIÇÃO DO SALDO
             '
             
'            If Not achouItem Is Nothing Then
'                ' Encontra a próxima linha vazia na coluna A da Planilha34 (wsDiarioAcessorio)
'                ultimaLinhaDiarioAcessorio = wsDiarioAcessorio.Cells(wsDiarioAcessorio.Rows.Count, 1).End(xlUp).Row + 1
'
'                ' Preenche a próxima célula vazia da coluna A com o valor da célula C correspondente na Planilha35 (DB Acessorio)
'                wsDiarioAcessorio.Cells(ultimaLinhaDiarioAcessorio, 1) = wsDBAcessorio.Cells(achouItem.Row, 3)
'                wsDiarioAcessorio.Cells(ultimaLinhaDiarioAcessorio, 12) = wsNS.Cells(1, 4)
'
'                ' Verifica se existe valor na coluna 5 e preenche na próxima linha
'                If wsDBAcessorio.Cells(achouItem.Row, 5).Value <> "" Then
'                    ultimaLinhaDiarioAcessorio = ultimaLinhaDiarioAcessorio + 1
'                    wsDiarioAcessorio.Cells(ultimaLinhaDiarioAcessorio, 1) = wsDBAcessorio.Cells(achouItem.Row, 5)
'                    wsDiarioAcessorio.Cells(ultimaLinhaDiarioAcessorio, 12) = wsNS.Cells(1, 4)
'                End If
'
'                ' Verifica se existe valor na coluna 7 e preenche na próxima linha
'                If wsDBAcessorio.Cells(achouItem.Row, 7) <> "" Then
'                    ultimaLinhaDiarioAcessorio = ultimaLinhaDiarioAcessorio + 1
'                    wsDiarioAcessorio.Cells(ultimaLinhaDiarioAcessorio, 1) = wsDBAcessorio.Cells(achouItem.Row, 7)
'                    wsDiarioAcessorio.Cells(ultimaLinhaDiarioAcessorio, 12) = wsNS.Cells(1, 4)
'                End If
'
'                ' Verifica se existe valor na coluna 9 e preenche na próxima linha
'                If wsDBAcessorio.Cells(achouItem.Row, 9) <> "" Then
'                    ultimaLinhaDiarioAcessorio = ultimaLinhaDiarioAcessorio + 1
'                    wsDiarioAcessorio.Cells(ultimaLinhaDiarioAcessorio, 1) = wsDBAcessorio.Cells(achouItem.Row, 9)
'                    wsDiarioAcessorio.Cells(ultimaLinhaDiarioAcessorio, 12) = wsNS.Cells(1, 4)
'                End If
'            End If
            
            ' Incrementa a linha da Planilha3
            LINHANS = LINHANS + 1
            
            ' Verifica se ultrapassou a última linha da Planilha3 e reinicia
            If LINHANS > ultimaLinhaNS Then
                Exit For ' Sai do loop se não houver mais dados em Planilha3
            End If
        End If
    Next LINHADIARIO
    
    ' Imprime a Planilha1
    Imprimir
    
    ' Zera o diário de carga para novo atendimento
    ZeraDiarioPraNovoAtendimento
    
    ' Salva o arquivo
    ThisWorkbook.Save
    
    ' Seleciona a página inicial
    SelecionaPaginaInicial
End Sub

