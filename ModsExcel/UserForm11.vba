Private Sub cmbTecnicoEmprestimo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 13 Then
           
    If cmbTecnicoEmprestimo.Value <> "" Then
    
        Planilha17.Cells(6, 2).Value = cmbTecnicoEmprestimo.Value
        
        AchaRE
        
        Planilha17.Cells(5, 2).Value = edtRETecnicoEmprestimo.Value
        
        Range("B4").Select
        
        UserForm11.Hide
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

Private Sub AchaRE()

On Error GoTo Erro

Dim codigo_produto As Variant

Dim resultado_procv As Variant
    
    If UserForm11.cmbTecnicoEmprestimo.Value <> "" Then
        codigo_produto = UserForm11.cmbTecnicoEmprestimo.Value
        codigo_produto = CStr(codigo_produto)
        resultado_procv = Application.VLookup(codigo_produto, Sheets("TECNICOS").Range("E:F"), 2, False)
        edtRETecnicoEmprestimo.Value = resultado_procv
    End If

Exit Sub
Erro:
MsgBox "VAI PAGAR A CONTA DE AGUA DA PROCISA ? TA BEBENDO UMA AGUA DO CARALHO ESSE 'RE' NÃO EXISTE", vbCritical, "RE DESCONHECIDO"
UserForm9.TREUTP.Value = ""
UserForm9.cmbNomeUTP.Value = ""


End Sub

Private Sub AchaNome()

On Error GoTo Erro

Dim codigo_produto As Variant

Dim resultado_procv As Variant
    
    If UserForm11.edtRETecnicoEmprestimo.Value <> "" Then
        codigo_produto = UserForm11.edtRETecnicoEmprestimo.Value
        codigo_produto = CLng(codigo_produto)
        resultado_procv = Application.VLookup(codigo_produto, Sheets("TECNICOS").Range("C:E"), 3, False)
        cmbTecnicoEmprestimo.Value = resultado_procv
    End If

Exit Sub
Erro:
MsgBox "VAI PAGAR A CONTA DE AGUA DA PROCISA ? TA BEBENDO UMA AGUA DO CARALHO ESSE 'RE' NÃO EXISTE", vbCritical, "RE DESCONHECIDO"
UserForm9.TREUTP.Value = ""
UserForm9.cmbNomeUTP.Value = ""


End Sub

Private Sub edtRETecnicoEmprestimo_Exit(ByVal Cancel As MSForms.ReturnBoolean)

 If edtRETecnicoEmprestimo.Value <> "" Then
        
        Planilha17.Cells(5, 2).Value = edtRETecnicoEmprestimo.Value
        
        AchaNome
        
        Planilha17.Cells(6, 2).Value = cmbTecnicoEmprestimo.Value
        
        Range("B4").Select
        
        UserForm11.Hide
    End If

End Sub


Private Sub UserForm_Activate()

Dim totlLIN As Long
Dim LIN As Long

cmbTecnicoEmprestimo.Clear
totlLIN = Planilha3.Range("a" & Rows.Count).End(xlUp).Row

For LIN = 3 To totlLIN

    cmbTecnicoEmprestimo.AddItem Planilha3.Cells(LIN, 1)

Next LIN

End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 27 Then
    ' A tecla Esc foi pressionada
    UserForm11.Hide

End If

End Sub

