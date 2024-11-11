Sub PesquisaAvancada()

    Dim Pesquisa As String
    Dim Texto As String
    Dim UltimaCelula As Range
    Dim celula As Range

    On Error Resume Next ' Ignora erros para evitar problemas caso a célula não tenha texto

    Pesquisa = UserForm6.TextBox1.Value
    If Pesquisa = "" Then
        MsgBox "Por favor, insira um valor para pesquisar.", vbExclamation, "Pesquisa Avançada"
        Exit Sub
    End If

    ' Define o intervalo da pesquisa
    Set UltimaCelula = Planilha15.Range("A200").End(xlUp)
    If UltimaCelula.Row < 1 Then
        MsgBox "A planilha está vazia.", vbInformation, "Pesquisa Avançada"
        Exit Sub
    End If

    ' Realiza a pesquisa
    For Each celula In Planilha15.Range("A1:A" & UltimaCelula.Row)
        Texto = celula.Text
        If UCase(Texto) Like UCase("*" & Pesquisa & "*") Then
            UserForm6.TextBox1.Text = Texto
            celula.Select
            Exit For
        End If
    Next celula

    If celula.Row > UltimaCelula.Row Then
        MsgBox "Texto não encontrado.", vbInformation, "Pesquisa Avançada"
    End If

End Sub
