Sub CriarAbasPorFamilia()
    Dim wsOrigem As Worksheet
    Dim wbNovo As Workbook
    Dim rngDados As Range
    Dim dictFamilias As Object
    Dim ultimaLinha As Long
    Dim celula As Range
    Dim Key As Variant
    Dim wsDestino As Worksheet
    Dim UltimaLinhaDestino As Long
    Dim rngLinhaOrigem As Range
    Dim nomeAba As String
    Dim dataAtual As String
    Dim nomeSugerido As String

    ' Desativar atualização de tela e cálculos para aumentar a velocidade
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Definir a planilha de origem
    Set wsOrigem = ThisWorkbook.Sheets("FabioMamado") ' Ajuste para a aba correta

    ' Determinar a última linha da coluna T (coluna 20)
    ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, 20).End(xlUp).Row

    ' Definir o intervalo dos dados na coluna T
    Set rngDados = wsOrigem.Range("T2:T" & ultimaLinha) ' Apenas a coluna T sem o cabeçalho

    ' Criar um dicionário para armazenar as famílias únicas
    Set dictFamilias = CreateObject("Scripting.Dictionary")

    ' Loop para capturar as famílias únicas
    For Each celula In rngDados
        If Not dictFamilias.exists(celula.Value) And celula.Value <> "" And celula.Value <> "N/D" Then
            dictFamilias.Add celula.Value, celula.Value
        End If
    Next celula

    ' Criar um novo workbook
    Set wbNovo = Workbooks.Add

    ' Loop para criar uma aba para cada família
    For Each Key In dictFamilias.Keys
        ' Substituir caracteres inválidos no nome da aba
        nomeAba = Key
        nomeAba = Replace(nomeAba, "/", "_") ' Substitui a barra por underline
        nomeAba = Replace(nomeAba, "\", "_") ' Substitui barra invertida por underline
        nomeAba = Replace(nomeAba, ":", "_") ' Substitui dois pontos por underline
        nomeAba = Replace(nomeAba, "*", "_") ' Substitui asterisco por underline
        nomeAba = Replace(nomeAba, "?", "_") ' Substitui ponto de interrogação por underline
        nomeAba = Replace(nomeAba, """", "_") ' Substitui aspas por underline
        nomeAba = Replace(nomeAba, "<", "_") ' Substitui menor que por underline
        nomeAba = Replace(nomeAba, ">", "_") ' Substitui maior que por underline
        nomeAba = Replace(nomeAba, "|", "_") ' Substitui barra vertical por underline

        On Error Resume Next ' Para evitar erro com nomes de abas inválidos
        wbNovo.Sheets.Add(After:=wbNovo.Sheets(wbNovo.Sheets.Count)).Name = nomeAba
        On Error GoTo 0
    Next Key

    ' Copiar as linhas correspondentes para as abas respectivas
    For Each Key In dictFamilias.Keys
        ' Substituir caracteres inválidos no nome da aba
        nomeAba = Key
        nomeAba = Replace(nomeAba, "/", "_")
        nomeAba = Replace(nomeAba, "\", "_")
        nomeAba = Replace(nomeAba, ":", "_")
        nomeAba = Replace(nomeAba, "*", "_")
        nomeAba = Replace(nomeAba, "?", "_")
        nomeAba = Replace(nomeAba, """", "_")
        nomeAba = Replace(nomeAba, "<", "_")
        nomeAba = Replace(nomeAba, ">", "_")
        nomeAba = Replace(nomeAba, "|", "_")

        ' Configurar a aba de destino
        Set wsDestino = wbNovo.Sheets(nomeAba)

        ' Adicionar cabeçalho como valores
        wsDestino.Rows(1).Value = wsOrigem.Rows(1).Value

        ' Definir a última linha da aba de destino onde será colada a próxima linha
        UltimaLinhaDestino = 2 ' A primeira linha já é o cabeçalho

        ' Loop para encontrar e copiar as linhas correspondentes de forma recursiva
        For Each celula In rngDados
            If celula.Value = Key Then
                ' Verificar se a célula na coluna S não tem erro
                If Not IsError(wsOrigem.Cells(celula.Row, 19).Value) Then
                    ' Copiar a linha inteira como valores
                    Set rngLinhaOrigem = wsOrigem.Rows(celula.Row)
                    wsDestino.Rows(UltimaLinhaDestino).Value = rngLinhaOrigem.Value

                    ' Colar a coluna B como texto para evitar notação científica
                    wsDestino.Cells(UltimaLinhaDestino, 2).NumberFormat = "@" ' Define como texto
                    wsDestino.Cells(UltimaLinhaDestino, 2).Value = wsOrigem.Cells(celula.Row, 2).Value

                    ' Atualizar a última linha de destino
                    UltimaLinhaDestino = UltimaLinhaDestino + 1
                End If
            End If
        Next celula
    Next Key

    ' Obter a data atual formatada
    dataAtual = Format(Date, "DD-MM-YYYY")
    nomeSugerido = "Arquivo em campo detalhado - " & dataAtual & ".xlsx"

    ' Restaurar as configurações após o processo
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' Salvar o novo arquivo em uma pasta específica com o nome sugerido
    Dim caminhoArquivo As String
    caminhoArquivo = Application.GetSaveAsFilename(InitialFileName:=nomeSugerido, FileFilter:="Arquivos Excel (*.xlsx), *.xlsx")
    If caminhoArquivo <> "False" Then
        wbNovo.SaveAs FileName:=caminhoArquivo, FileFormat:=xlOpenXMLWorkbook
        MsgBox "Novo arquivo criado e salvo em: " & caminhoArquivo, vbInformation
    Else
        MsgBox "Arquivo não foi salvo.", vbExclamation
    End If
End Sub
