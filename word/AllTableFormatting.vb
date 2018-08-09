Sub AllTableFormatting()

Dim pT As Word.Table
For Each pT In ActiveDocument.Tables
    pT.Style = "Стиль1"
    pT.ApplyStyleHeadingRows = Not pT. _
    ApplyStyleHeadingRows
    pT.ApplyStyleFirstColumn = Not pT. _
    ApplyStyleFirstColumn
    pT.Range.Style = ActiveDocument.Styles("Table Text")
Next
End Sub
