Public conexaoUTP As ADODB.Connection
Public rsUTP As ADODB.Recordset

Sub ConectarPlanilhaUTP()

    Dim EnderecoPlan As String
    Dim Provider As String, Ex As String
    Dim NamePcValido1 As String

    NamePcValido1 = Environ$("ComputerName")

    On Error GoTo Erro

    Provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
    Ex = "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"

    Select Case NamePcValido1
        Case "PRSPPE04EFCK"
            EnderecoPlan = "C:\Users\r.santos12.CONDUMEX\OneDrive - Grupo Carso\Meu Drive\almox\Controle de Estoque\Controle de Cabos" & "\ControlCaboUTP.xlsx;"
        Case "DESKTOP-DBKEF0A"
            EnderecoPlan = "C:\Users\BAIA 2\OneDrive - Grupo Carso\Meu Drive\almox\Controle de Estoque\Controle de Cabos" & "\ControlCaboUTP.xlsx;"
        Case "PRSP41Y0QK2"
            EnderecoPlan = "C:\Users\e.silva29\OneDrive - Grupo Carso\Meu Drive\almox\Controle de Estoque\Controle de Cabos" & "\ControlCaboUTP.xlsx;"
        Case "DESKTOP-ROMUVAJ"
            EnderecoPlan = "C:\Users\Poderoso Deus Karel\OneDrive - Grupo Carso\Meu Drive\almox\Controle de Estoque\Controle de Cabos" & "\ControlCaboUTP.xlsx;"
        Case "PRSTPE04EFD5"
            EnderecoPlan = "C:\Users\g.lima2\OneDrive - Grupo Carso\Meu Drive\almox\Controle de Estoque\Controle de Cabos" & "\ControlCaboUTP.xlsx;"
        Case Else
            MsgBox "Nome do computador não reconhecido", vbCritical, "CONECTAR"
            Exit Sub
    End Select

    Set conexaoUTP = New ADODB.Connection
    conexaoUTP.Open Provider & EnderecoPlan & Ex

    Exit Sub

Erro:
    MsgBox "Erro ao conectar a planilha de cabos", vbCritical, "CONECTAR"

End Sub

Sub DesconectarPlanUTP()
    If Not conexaoUTP Is Nothing Then
        conexaoUTP.Close
        Set conexaoUTP = Nothing
    End If
End Sub

