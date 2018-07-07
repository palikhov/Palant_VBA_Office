Sub MassProtect1()

Dim wsh As Worksheet
Dim strpassword As String
Dim i As Integer

strpassword = InputBox("Введите пароль")

i = 0

For Each wsh In ActiveWorkbook.Worksheets

wsh.Protect Password:=strpassword

i = i + 1
Application.StatusBar = "Установка защиты на лист" & wsh.Name & "¦ ¦ ¦ ¦ ¦ Общий прогресс ¦" & i & "¦ из ¦" & ActiveWorkbook.Sheets.Count

Next wsh

MsgBox ("Защита установлена. Пароль:" & strpassword)
End Sub
Sub MassUnProtect1()

Dim wsh As Worksheet
Dim strpassword As String
Dim i As Integer

strpassword = InputBox("Введите пароль")

i = 0

For Each wsh In ActiveWorkbook.Worksheets

wsh.Unprotect Password:=strpassword

i = i + 1
Application.StatusBar = "Снятие защиты с листа" & wsh.Name & "¦ ¦ ¦ ¦ ¦ Общий прогресс ¦" & i & "¦ из¦" & ActiveWorkbook.Sheets.Count

Next wsh

MsgBox ("Защита снята.")
End Sub
