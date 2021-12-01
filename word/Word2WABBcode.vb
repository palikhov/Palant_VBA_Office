'Word2BBCode-Converter v0.2, December 1, 2021
'Palant edition
'Some parts adapted from
'Word2Wiki-Converter V0.4, May 28, 2006 and Matthew Kruer project
'http://de.wikipedia.org/wiki/Wikipedia:Helferlein/Word2MediaWikiPlus
'Original Version by InfPro: http://www.infpro.com/downloads/downloads/wordmedia.htm
'Major improvements by Gunter Schmidt, Mail me: Word2MediaWikiPlus@beadsoft.de
'Works only with Word 2000 and above
'License: GPL: Feel free to use and modify. Keep the credits and do not sell.
'Functionality:
'
Sub wikiformattingclearing()

    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Font.Reset
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[edit]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Copy
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " • "
        .Replacement.Text = " - "
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Cut
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "• "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Private Sub ReplaceP_PwithP()

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p ^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Private Sub ReplacePwith2P()

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "^p^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Private Sub ClearingNonVisibleSymbols()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^p "
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub Word2BBCode()
    
    Application.ScreenUpdating = False
    ClearingNonVisibleSymbols
    ReplaceP_PwithP
    ReplacePwith2P
    'Heading 1 to Heading 5
    ConvertParagraphStyle wdStyleHeading1, "[h1]", "[/h1]"
    ConvertParagraphStyle wdStyleHeading2, "[h2]", "[/h2]"
    ConvertParagraphStyle wdStyleHeading3, "[h3]", "[/h3]"
    ConvertParagraphStyle wdStyleHeading4, "[h4]", "[/h4]"
    ClearingNonVisibleSymbols
    ReplaceP_PwithP
    ReplacePwith2P
    ConvertItalic
    ConvertBold
    ConvertUnderline
    ConvertLists

   
    ActiveDocument.Content.Copy
    
    Application.ScreenUpdating = True
End Sub
Private Sub ConvertParagraphStyle(styleToReplace As WdBuiltinStyle, _
                                    preText As String, _
                                    postText As String)
    Dim normalStyle As Style
    Set normalStyle = ActiveDocument.Styles(wdStyleNormal)
    
    ActiveDocument.Select
    
    With Selection.Find
    
        .ClearFormatting
        .Style = ActiveDocument.Styles(styleToReplace)
        .Text = ""
        
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        .Forward = True
        .Wrap = wdFindContinue
        
        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                       
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore preText
                    .InsertAfter postText
                End If
                
                .Style = normalStyle
            End With
        Loop
    End With
End Sub
Private Sub ConvertBold()
    ActiveDocument.Select
    
    With Selection.Find
    
        .ClearFormatting
        .Font.Bold = True
        .Text = ""
        
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        .Forward = True
        .Wrap = wdFindContinue
        
        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                          .Font.Bold = False
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                       
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore "[b]"
                    .InsertAfter "[/b]"
                End If
                
                .Font.Bold = False
            End With
        Loop
    End With
End Sub
Private Sub ConvertItalic()
    ActiveDocument.Select
    
    With Selection.Find
    
        .ClearFormatting
        .Font.Italic = True
        .Text = ""
        
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        .Forward = True
        .Wrap = wdFindContinue
        
        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Font.Italic = False
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                       
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore "[i]"
                    .InsertAfter "[/i]"
                End If
                
                .Font.Italic = False
            End With
        Loop
    End With
End Sub
Private Sub ConvertUnderline()
    ActiveDocument.Select
    
    With Selection.Find
    
        .ClearFormatting
        .Font.Underline = True
        .Text = ""
        
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        .Forward = True
        .Wrap = wdFindContinue
        
        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Font.Underline = False
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                       
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore "[u]"
                    .InsertAfter "[/u]"
                End If
                
                .Font.Underline = False
            End With
        Loop
    End With
End Sub
 
Private Sub ConvertLists()
   Dim para As Paragraph
    For Each para In ActiveDocument.ListParagraphs
        With para.Range
            For i = 1 To .ListFormat.ListLevelNumber
                .InsertBefore "[li]"
                .InsertAfter "[/li]"
            Next i
            .InsertBefore "[ul]"
            .InsertAfter "[/ul]"
            .ListFormat.RemoveNumbers
            
        End With
    Next para
End Sub
Private Sub ConvertHyperlinks()
    'converts Hyperlinks
    '24-MAY-2006: only convert http..., mark others with error marker

    Dim hyperCount&
    Dim i&
    Dim addr$ ', title$

    hyperCount = ActiveDocument.Hyperlinks.Count

    For i = 1 To hyperCount

        With ActiveDocument.Hyperlinks(1) 'must be 1, since the delete changes count and position

            addr = .Address
            If Trim$(addr) = "" Then addr = "no hyperlink found"
            'title = .Range.Text
           
            'http, ftp
            If LCase(Left$(addr, 4)) = "http" Or LCase(Left$(addr, 3)) = "ftp" Then
                .Delete 'hyperlink
                .Range.InsertBefore "[url=" & addr & "]"
                .Range.InsertAfter "[/url]"
               
                GoTo ConvertHyperlinks_Next
            End If
           
            'mailto:
            If LCase(Left$(addr, 7)) = "mailto:" Then
                .Delete 'hyperlink
                .Range.InsertBefore "[email]" & addr & " "
                .Range.InsertAfter "[/email]"
               
                GoTo ConvertHyperlinks_Next
            End If
           
            'file guess
            If Len(addr) > 4 Then 'the reason for not nice goto
                If Mid$(addr, Len(addr) - 3, 1) = "." Then
                    .Delete
                    .Range.InsertBefore "[file://" & Replace(addr, " ", "_") & " "
                    .Range.InsertAfter "]"
                   
                    GoTo ConvertHyperlinks_Next
                End If
            End If
           
            'unidentified
            .Delete
            .Range.InsertBefore UnableToConvertMarker & "[" & addr & " "
            .Range.InsertAfter "]"

ConvertHyperlinks_Next:
        End With

    Next i

End Sub
Sub ConvertTables()

    Dim myRange As Word.Range
    Dim tTable  As Word.Table
    Dim tRow    As Word.Row
    Dim tCell   As Word.Cell
    Dim strText As String
    
    Dim i   As Long 'tTable.Rows
    Dim j   As Long 'tRow.Cells
    Dim k   As Long
    Dim l   As Long
    
    For Each tTable In ActiveDocument.Tables
            
            'Memorize table text
            ReDim x(1 To tTable.Rows.Count, 1 To tTable.Columns.Count)
            i = 0
            For Each tRow In tTable.Rows
                i = i + 1
                j = 0
                For Each tCell In tRow.Cells
                    j = j + 1
                    strText = tCell.Range.Text
                    x(i, j) = Left(strText, Len(strText) - 2)
                Next tCell
            Next tRow
            
            'Delete table and position after table
            Set myRange = tTable.Range
            myRange.Collapse Direction:=wdCollapseEnd
            tTable.Delete
            
            'Rewrite table with memorized text
            myRange.InsertParagraphAfter
            myRange.InsertAfter ("[table]")
            myRange.InsertParagraphAfter
            For k = 1 To i 'Rows
            
            myRange.InsertAfter ("[tr]")
            myRange.InsertParagraphAfter
                     For l = 1 To j 'Cells
                          myRange.InsertAfter "[td]" + x(k, l) + "[/td]"
                          myRange.InsertParagraphAfter
                     Next l
            myRange.InsertAfter ("[/tr]")
            myRange.InsertParagraphAfter
            
            
            Next k
            
            myRange.InsertAfter ("[/table]")
            myRange.InsertParagraphAfter
            
    Next tTable

End Sub


