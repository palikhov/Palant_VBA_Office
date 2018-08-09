Sub tableformatting()
'
' tableformatting Ìàêðîñ
'
Dim pT As Word.Table
For Each pT In ActiveDocument.Tables
    pT.Style = "Ñòèëü1"
    pT.ApplyStyleHeadingRows = Not pT. _
    ApplyStyleHeadingRows
    pT.ApplyStyleFirstColumn = Not pT. _
    ApplyStyleFirstColumn
    pT.Columns.PreferredWidthType = wdPreferredWidthAuto
    pT.Columns.PreferredWidth = 0
    pT.Range.Style = ActiveDocument.Styles("Table Text")
Next
End Sub
