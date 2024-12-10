Private Sub UserForm_Activate()

Dim totlLIN As Long
Dim LIN As Long

cmbTecnicoEmprestimo.Clear
totlLIN = Planilha9.Range("a" & Rows.Count).End(xlUp).Row

For LIN = 3 To totlLIN

    UserForm3.cmbTecnicoEmprestimo.AddItem Planilha3.Cells(LIN, 1)

Next LIN

End Sub

Private Sub cmbTecnicoEmprestimo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 13 Then
           
    If cmbTecnicoEmprestimo.Value <> "" Then
    
        PreencherNomes
        
        UserForm3.Hide
    Else
        MsgBox " TA FALTANDO NADA NÃO ARROMADO ?"
    End If
            
End If

End Sub

Private Sub cmbTecnicoEmprestimo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim totlLIN As Long
Dim LIN As Long

If cmbTecnicoEmprestimo.TextLength > 0 Then

    totlLIN = Planilha2.Range("a" & Rows.Count).End(xlUp).Row
    
    For LIN = 3 To totlLIN
    
        cmbTecnicoEmprestimo.AddItem Planilha2.Cells(LIN, 1)
    
    Next LIN

Else

    totlLIN = Planilha2.Range("a" & Rows.Count).End(xlUp).Row
    
    For LIN = 3 To totlLIN
    
        cmbTecnicoEmprestimo.AddItem Planilha2.Cells(LIN, 2)
    
    Next LIN
    
End If


End Sub

Private Sub PreencherNomes()
    Dim nomeSelecionado As String
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    
    nomeSelecionado = cmbTecnicoEmprestimo.Value
    
    ' Planilha onde você deseja preencher (planilha 24)
    Set ws = Planilha24
    
    ' Intervalo de busca (linha 2 a 20 da coluna 1)
    Set rng = ws.Range("A2:A20")
    
    ' Encontre a próxima célula vazia na coluna A
    For Each cell In rng
        If cell.Value = "" Then
            cell.Value = nomeSelecionado
            PreencherDevendo cell
            Exit For
        End If
    Next cell
    
    ' Limpe o TextBox após adicionar
    cmbTecnicoEmprestimo.Value = ""
End Sub

Private Sub PreencherDevendo(ByVal celula As Range)
    Dim itemSelecionado As String
    
    ' Item selecionado no ComboBox
    If cmbDevendo.Value <> "" Then
        itemSelecionado = cmbDevendo.Value
    Else
        MsgBox " O CARA TA DEVENDO OQUE O ARROMBADO ? "
        cell.Value = Empty
        End
    End If
    
    ' Preenche a célula adjacente na coluna B com o valor selecionado no ComboBox "cmbDevendo"
    celula.Offset(0, 1).Value = itemSelecionado
End Sub

Private Sub UserForm_Initialize()
    ' Adicione os itens ao ComboBox
    cmbDevendo.AddItem "CABO UTP"
    cmbDevendo.AddItem "EQUIPAMENTO"
End Sub


'
