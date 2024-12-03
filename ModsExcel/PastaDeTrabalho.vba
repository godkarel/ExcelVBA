Private Sub Workbook_Open()

Dim MyDate As Date
Dim DatePassou As String
Dim RegisterDate As Date
Dim NamePcValido1 As String

SelecionaPaginaInicial

NamePcValido1 = Environ$("ComputerName")

??????????? = VBA.Date
?????? = "??/??/???/"

If NamePcValido1 = "PRSPPE04EFCK" Or NamePcValido1 = "DESKTOP-DBKEF0A" Or NamePcValido1 = "PRSP41Y0QK2" Or NamePcValido1 = "PRSTPE04EFD5" Or NamePcValido1 = "PRSPPE04EFCS" Or NamePcValido1 = "DESKTOP-ROMUVAJ" Or NamePcValido1 = "PRSPPE043VEK" Or NamePcValido1 = "PRSPPE04EFD5" Or NamePcValido1 = "DESKTOP-NU71HVS" Then

    If Planilha16.Cells(7, 3) = "VERDADEIRO" Then
        
        UserForm4.Show
        
    End If

    If MyDate > RegisterDate Then
    
        Planilha16.Cells(7, 3) = "VERDADEIRO"
        
        ThisWorkbook.Save
        
        UserForm4.Show
        
    Else
    
        Planilha16.Cells(7, 3) = "FALSO"
        
        ThisWorkbook.Save
        
    End If

Else

    MsgBox ("Computador Não Registrado Envie um Email para ???????????")
    
    UserForm4.Show

End If

UserForm8.Show

End Sub