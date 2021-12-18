Public Function RegExpExtract(Text As String, Pattern As String, Optional Item As Integer = 1) As String
    On Error GoTo ErrHandl
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = Pattern
    regex.Global = True
    If regex.Test(Text) Then
        Set matches = regex.Execute(Text)
        RegExpExtract = matches.Item(Item - 1)
        Exit Function
    End If
ErrHandl:
    RegExpExtract = CVErr(xlErrValue)
End Function
