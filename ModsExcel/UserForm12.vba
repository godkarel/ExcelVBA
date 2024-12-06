Private Sub cmbTecnico_Change()

Dim valorPesquisado As String
    Dim resultado As Variant
    
    If cmbTecnico.Value <> "" Then
        ' Obtenha o valor da caixa de texto
        valorPesquisado = cmbTecnico.Value
        ' Execute o PROCV na planilha desejada (substitua "Planilha1" pelo nome da sua planilha)
        resultado = Application.WorksheetFunction.VLookup(valorPesquisado, ThisWorkbook.Sheets("TECNICOS").Range("B2:D100"), 3, False)
        ' Exiba o resultado em uma caixa de mensagem
        TCodigo.Value = resultado
    End If

End Sub

Private Sub TCodigo_Change()

If TCodigo.Value <> Empty Then

    THDNG.Value = Application.WorksheetFunction.CountIfs(Planilha22.Range("G:G"), TCodigo.Value, Planilha22.Range("E:E"), "PROCISA DO BRASIL PROJETOS CONSTRUC", Planilha22.Range("F:F"), "INICIALIZADO", Planilha22.Range("I:I"), "DECODER HDNG")
    
    T31.Value = Application.WorksheetFunction.CountIfs(Planilha22.Range("G:G"), TCodigo.Value, Planilha22.Range("E:E"), "PROCISA DO BRASIL PROJETOS CONSTRUC", Planilha22.Range("F:F"), "INICIALIZADO", Planilha22.Range("I:I"), "EMTA WIFI 3.1")
    
    T1GB.Value = Application.WorksheetFunction.CountIfs(Planilha22.Range("G:G"), TCodigo.Value, Planilha22.Range("E:E"), "PROCISA DO BRASIL PROJETOS CONSTRUC", Planilha22.Range("F:F"), "INICIALIZADO", Planilha22.Range("I:I"), "EMTA 3.1 1GB")
    
    TDUAL.Value = Application.WorksheetFunction.CountIfs(Planilha22.Range("G:G"), TCodigo.Value, Planilha22.Range("E:E"), "PROCISA DO BRASIL PROJETOS CONSTRUC", Planilha22.Range("F:F"), "INICIALIZADO", Planilha22.Range("I:I"), "EMTA 3.0 DUAL BAND")
    
    TMesh.Value = Application.WorksheetFunction.CountIfs(Planilha22.Range("G:G"), TCodigo.Value, Planilha22.Range("E:E"), "PROCISA DO BRASIL PROJETOS CONSTRUC", Planilha22.Range("F:F"), "INICIALIZADO", Planilha22.Range("I:I"), "EXTENSOR MESH")
    
    TMeshW6.Value = Application.WorksheetFunction.CountIfs(Planilha22.Range("G:G"), TCodigo.Value, Planilha22.Range("E:E"), "PROCISA DO BRASIL PROJETOS CONSTRUC", Planilha22.Range("F:F"), "INICIALIZADO", Planilha22.Range("I:I"), "EXTENSOR MESH WIFI 6")
    
    TONT.Value = Application.WorksheetFunction.CountIfs(Planilha22.Range("G:G"), TCodigo.Value, Planilha22.Range("E:E"), "PROCISA DO BRASIL PROJETOS CONSTRUC", Planilha22.Range("F:F"), "INICIALIZADO", Planilha22.Range("I:I"), "ONT")
    
    TONTGB.Value = Application.WorksheetFunction.CountIfs(Planilha22.Range("G:G"), TCodigo.Value, Planilha22.Range("E:E"), "PROCISA DO BRASIL PROJETOS CONSTRUC", Planilha22.Range("F:F"), "INICIALIZADO", Planilha22.Range("I:I"), "ONT WIFI 6")
    
    TIPTV.Value = Application.WorksheetFunction.CountIfs(Planilha22.Range("G:G"), TCodigo.Value, Planilha22.Range("E:E"), "PROCISA DO BRASIL PROJETOS CONSTRUC", Planilha22.Range("F:F"), "INICIALIZADO", Planilha22.Range("I:I"), "DECODER 4K - IPTV")
    
    TChip.Value = Application.WorksheetFunction.CountIfs(Planilha22.Range("G:G"), TCodigo.Value, Planilha22.Range("E:E"), "PROCISA DO BRASIL PROJETOS CONSTRUC", Planilha22.Range("F:F"), "INICIALIZADO", Planilha22.Range("I:I"), "CHIP DA CLARO")
    
    TCardless.Value = Application.WorksheetFunction.CountIfs(Planilha22.Range("G:G"), TCodigo.Value, Planilha22.Range("E:E"), "PROCISA DO BRASIL PROJETOS CONSTRUC", Planilha22.Range("F:F"), "INICIALIZADO", Planilha22.Range("I:I"), "4K CARDLESS")
    
    T4KLegado.Value = Application.WorksheetFunction.CountIfs(Planilha22.Range("G:G"), TCodigo.Value, Planilha22.Range("E:E"), "PROCISA DO BRASIL PROJETOS CONSTRUC", Planilha22.Range("F:F"), "INICIALIZADO", Planilha22.Range("I:I"), "DECODER 4K")
    
    If THDNG.Value <> Empty Then

        TTHDNG.Value = THDNG.Value - 3
        
        TT31.Value = T31.Value - 3
        
        TT1GB.Value = T1GB.Value - 3
        
        TTDUAL.Value = TDUAL.Value - 5
        
        TTMESH.Value = TMesh.Value - 4
        
        TTMESHW6.Value = TMeshW6.Value - 3
        
        TTONT.Value = TONT.Value - 3
        
        TTONTGB.Value = TONTGB.Value - 1
        
        TTIPTV.Value = TIPTV.Value - 1
        
        TTCHIP.Value = TChip.Value - 2
        
        TTCARDLESS.Value = TCardless.Value - 1
        
        TT4K.Value = T4KLegado.Value - 1
        
    End If
    
End If

AUXILIOPORCOR

End Sub

Private Sub TTHDNG_Change()

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

Private Sub UserForm_Initialize()
If Planilha3.Cells(1, 3).Value <> "" Or _
   Planilha9.Cells(1, 3).Value <> "" Or _
   Planilha13.Cells(2, 3).Value <> "" Or _
   Planilha16.Cells(2, 3).Value <> "" Then

    cmbTecnico.Value = Planilha3.Cells(1, 3)
End If
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 27 Then
    ' A tecla Esc foi pressionada
    UserForm12.Hide

End If

End Sub
