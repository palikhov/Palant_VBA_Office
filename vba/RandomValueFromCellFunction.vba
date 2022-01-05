Function RandomValueFromCell(r As Range, separator As String)

Dim cell As Range
Dim i As Integer
Dim x As Double
Dim ArraySiza As Integer

List = Split(r, separator)

ArraySize = UBound(List) + 1

RandomValueFromCell = List(WorksheetFunction.RandBetween(1, ArraySize) - 1)


End Function

