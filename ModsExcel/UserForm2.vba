
Private Sub cmbTecnico_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 13 Then

Dim nomeSelecionado As String
Dim ws As Worksheet
Dim rng As Range
Dim cell As Range
Dim Encontrado As Boolean
           
    If cmbTecnico.Value <> "" Then
    
    nomeSelecionado = cmbTecnico.Value
    
    Set ws = Planilha24
    
    Set rng = ws.Range("A2:A20")
    
    Encontrado = False
    
        For Each cell In rng
            If cell.Value = nomeSelecionado Then
                MsgBox "Nome: " & nomeSelecionado & vbCrLf & "DEVENDO: " & cell.Offset(0, 1).Value
                Encontrado = True
            Exit For
            End If
        Next cell
        
        Planilha9.Cells(1, 3).Value = cmbTecnico.Value
        
        Planilha16.Cells(2, 3).Value = cmbTecnico.Value
        
        Planilha13.Cells(2, 3).Value = cmbTecnico.Value
        
        Planilha16.Cells(2, 3).Value = cmbTecnico.Value
        
        Range("B4").Select
        
        UserForm2.Hide
    End If
            
End If

End Sub

Private Sub cmbTecnico_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim totlLIN As Long
Dim LIN As Long

If cmbTecnico.TextLength > 0 Then

    totlLIN = Planilha2.Range("a" & Rows.Count).End(xlUp).Row
    
    For LIN = 3 To totlLIN
    
        cmbTecnico.AddItem Planilha2.Cells(LIN, 1)
    
    Next LIN

Else

    totlLIN = Planilha2.Range("a" & Rows.Count).End(xlUp).Row
    
    For LIN = 3 To totlLIN
    
        cmbTecnico.AddItem Planilha2.Cells(LIN, 2)
    
    Next LIN
    
End If


End Sub


Private Sub UserForm_Activate()

Dim totlLIN As Long
Dim LIN As Long

cmbTecnico.Clear
totlLIN = Planilha9.Range("a" & Rows.Count).End(xlUp).Row

For LIN = 3 To totlLIN

    cmbTecnico.AddItem Planilha3.Cells(LIN, 1)

Next LIN

End Sub



Private Sub UserForm_Initialize()
    If Planilha3.Cells(1, 3).Value <> "" Or _
       Planilha9.Cells(1, 3).Value <> "" Or _
       Planilha13.Cells(2, 3).Value <> "" Or _
       Planilha16.Cells(2, 3).Value <> "" Then
    
        cmbTecnico.Value = Planilha3.Cells(1, 3)
    End If
End Sub