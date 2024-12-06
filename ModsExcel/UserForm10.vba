Private Sub cmbTecnico_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 13 Then
           
    If cmbTecnico.Value <> "" Then
    
        Planilha19.Cells(1, 3).Value = cmbTecnico.Value
        
        UserForm10.Hide
        
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

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Activate()

Dim totlLIN As Long
Dim LIN As Long

cmbTecnico.Clear
totlLIN = Planilha3.Range("a" & Rows.Count).End(xlUp).Row

For LIN = 3 To totlLIN

    cmbTecnico.AddItem Planilha3.Cells(LIN, 1)

Next LIN

End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 27 Then
    ' A tecla Esc foi pressionada
    UserForm10.Hide

End If

End Sub



