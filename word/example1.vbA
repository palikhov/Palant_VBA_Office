Sub test()
Dim doc As Document
Dim p As Paragraph
Dim s As String
Set doc = ActiveDocument
For Each p In doc.Paragraphs
    Debug.Print p.Range.Sentences(1)
Next
//Here is an example of extracting the first SENTENCE from each paragraph in a document. This would be code written in and called from Word Application.

End Sub
