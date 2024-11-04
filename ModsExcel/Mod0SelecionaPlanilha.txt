Sub SelecionaCargaTrino()
    ' Seleciona a Planilha3
    Planilha3.Select
    
    ' Verifica se a célula (1, 3) está vazia
    If Planilha3.Cells(1, 3) = "" Then
        ' Mostra o UserForm1 e define o foco no cmbTecnico
        UserForm1.Show
        UserForm1.cmbTecnico.SetFocus
    End If
End Sub

Sub SelecionaDiarioCargaTrino()
    ' Seleciona a Planilha1
    Planilha1.Select
End Sub

Sub SelecionaDevolucaoTrino()
    ' Seleciona a Planilha9
    Planilha9.Select
    
    ' Verifica se a célula (1, 3) está vazia
    If Planilha9.Cells(1, 3) = "" Then
        ' Mostra o UserForm2 e define o foco no cmbTecnico
        UserForm2.Show
        UserForm2.cmbTecnico.SetFocus
    End If
End Sub

Sub SelecionaDiarioDevolucaoTrino()
    ' Seleciona a Planilha11
    Planilha11.Select
End Sub

Sub SelecionaPaginaInicial()
    ' Seleciona a Planilha7
    Planilha7.Select
End Sub

Sub SelecionaLivroTrino()
    ' Seleciona a Planilha4
    Planilha4.Select
End Sub

Sub SelecionaTecnicoTrino()
    ' Seleciona a Planilha2
    Planilha2.Select
End Sub

Sub SelecionaCargaMic()
    ' Seleciona a Planilha13
    Planilha13.Select
    
    ' Verifica se a célula (2, 3) está vazia
    If Planilha13.Cells(2, 3) = "" Then
        ' Mostra o UserForm5 e define o foco no cmbTecnico
        UserForm5.Show
        UserForm5.cmbTecnico.SetFocus
    End If
End Sub

Sub SelecionaDiarioMic()
    ' Seleciona a Planilha14
    Planilha14.Select
End Sub

Sub SelecionaBibMic()
    ' Seleciona a Planilha15
    Planilha15.Select
End Sub

Sub SelecionaCODEQUIP()
    ' Seleciona a Planilha8
    Planilha8.Select
End Sub

Sub SelecionaAlmoxarife()
    ' Seleciona a Planilha16
    Planilha16.Select
End Sub

Sub SelecionaComprovante()
    ' Seleciona a Planilha17
    Planilha17.Select
    
    ' Seleciona a célula B5
    Range("B5").Select
End Sub

Sub SelecionaProcurado()
    ' Seleciona a Planilha24
    Planilha24.Select
    
    ' Seleciona a célula A2
    Range("A2").Select
End Sub

Sub AbreFormProcurado()
    ' Mostra o UserForm3 e define o foco no cmbDevendo
    UserForm3.Show
    UserForm3.cmbDevendo.SetFocus
End Sub

Sub SelecionaFormularioCabo()
    ' Define os valores iniciais do UserForm7
    UserForm7.TData.Value = Date
    UserForm7.cmbTIPO.AddItem "Cabo Branco"
    UserForm7.cmbTIPO.AddItem "Cabo Preto"
    UserForm7.cmbTIPO.AddItem "Bobina Preta"
    UserForm7.cmbTIPO.AddItem "Fibra Branca"
    UserForm7.cmbTIPO.AddItem "Fibra Cinza"
    UserForm7.cmbMetragem.AddItem "1KM"
    UserForm7.cmbMetragem.AddItem "500m"
    UserForm7.cmbMetragem.AddItem "305m"
    UserForm7.cmbMetragem.AddItem "100m"
    UserForm7.cmbTIPO.Value = "Cabo Branco"
    
    ' Mostra o UserForm7 e define o foco no cmbNome
    UserForm7.Show
    UserForm7.cmbNome.SetFocus
End Sub

Sub SelecionaUTPCabo()
    ' Define os valores iniciais do UserForm9
    UserForm9.TData.Value = Date
    UserForm9.cmbTIPO.AddItem "CABO UTP"
    UserForm9.cmbTIPO.AddItem "CONECTOR RJ45"
    UserForm9.cmbTIPO.AddItem "DEV CABO UTP"
    UserForm9.cmbTIPO.AddItem "DEV CONEC RJ45"
    UserForm9.cmbTIPO.AddItem "CB TELEFONICO"
    UserForm9.cmbTIPO.AddItem "DEV CB TELEFONICO"
    
    ' Mostra o UserForm9 e define o foco no cmbNomeUTP
    UserForm9.Show
    UserForm9.cmbNomeUTP.SetFocus
End Sub

Sub aindanaofaznada()
    ' Exibe uma mensagem de aviso
    MsgBox "Ainda não configurado"
End Sub

Sub SelecionaEntradaSeriais()
    ' Seleciona a Planilha19
    Planilha19.Select
    
    ' Seleciona a célula B4
    Range("B4").Select
End Sub

Sub SelecionaEntradaDINR()
    ' Seleciona a Planilha20
    Planilha20.Select
End Sub

Sub SelecionaTecnicoDINR()
    ' Mostra o UserForm10 e define o foco no cmbTecnico
    UserForm10.Show
    UserForm10.cmbTecnico.SetFocus
End Sub

Sub SelecionaTecnicoEmprestimo()
    ' Mostra o UserForm11 e define o foco no cmbTecnicoEmprestimo
    UserForm11.Show
    UserForm11.edtRETecnicoEmprestimo.Value = ""
    UserForm11.cmbTecnicoEmprestimo.Value = ""
    UserForm11.cmbTecnicoEmprestimo.SetFocus
End Sub

Sub SelecionaNaoEraPraClicar()
    ' Seleciona a Planilha12
    Planilha12.Select
End Sub

Sub SelecionaStatus()
    ' Seleciona a Planilha22
    Planilha22.Select
End Sub

Sub SelecionaAtendidos()
    ' Seleciona a Planilha23
    Planilha23.Select
End Sub

Sub SelecionaThaisCarla()
    ' Mostra o UserForm12
    UserForm12.Show
End Sub

Sub AtivaFormProcura()
    ' Mostra o UserForm1
    UserForm1.Show
    
    ' Seleciona a célula (4, 2) da Planilha3
    Planilha3.Cells(4, 2).Select
End Sub

Sub AtivaFormProcuraDEV()
    ' Mostra o UserForm2
    UserForm2.Show
    
    ' Seleciona a célula (4,
End Sub

Sub AtivaFormAtendente()
    ' Mostra o UserForm2
    UserForm8.Show
    
    ' Seleciona a célula (4,
End Sub

Sub SelecionaBotPesquisa()
    ' Seleciona a Planilha23
    Planilha26.Select
End Sub

Sub SelecionaDiarioKit()
    ' seleciona o Kit de domingo
    Planilha27.Select
End Sub


Sub SelecionaKitDeSeiLa()
    ' seleciona o Kit de domingo
    Planilha28.Select
End Sub

Sub SelecionaBipagemDoKit()
    ' seleciona o Kit de domingo
    Planilha29.Select
End Sub


Sub SelecionaAutonomiaDinamica()
    ' seleciona o Kit de domingo
    Planilha31.Select
End Sub

Sub SelecionaEmCampoFabio()
    ' seleciona o Kit de domingo
    Planilha32.Select
    ActiveWorkbook.RefreshAll
End Sub

Sub SelecionaEmCampo()
    ' seleciona o Kit de domingo
    Planilha30.Select
End Sub

Sub SelecionaDiarioAcessorio()
    ' seleciona o Kit de domingo
    Planilha34.Select
End Sub

Sub SelecionaBDAcessorio()
    ' seleciona o Kit de domingo
    Planilha35.Select
End Sub

Sub SelecionaCriacaoMicBTP()
    ' seleciona o Kit de domingo
    Planilha37.Select
End Sub


