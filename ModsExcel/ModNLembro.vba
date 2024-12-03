Sub TratamentoDesligado()
'
' Macro1 Macro
'

'
    Sheets("FabioMamado").Select
    Columns("B:B").Select
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=True, OtherChar:= _
        "/", FieldInfo:=Array(Array(1, 9), Array(2, 2), Array(3, 9)), _
        TrailingMinusNumbers:=True
    Range("B1").Select
End Sub