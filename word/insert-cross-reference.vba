
Sub selecttables()
Dim mytable As Table
 
For Each mytable In ActiveDocument.Tables
mytable.Range.Editors.Add wdEditorEveryone
Next
ActiveDocument.SelectAllEditableRanges (wdEditorEveryone)
ActiveDocument.DeleteAllEditableRanges (wdEditorEveryone)
End Sub
Sub InsertCrossRef()

    Dim RefList As Variant
    Dim LookUp As String
    Dim Ref As String
    Dim s As Integer, t As Integer
    Dim i As Integer

    On Error GoTo ErrExit
    With Selection.Range
        ' discard leading blank spaces
        Do While (Asc(.Text) = 32) And (.End > .Start)
            .MoveStart wdCharacter
        Loop
        ' discard trailing blank spaces, full stops and CRs
        Do While ((Asc(Right(.Text, 1)) = 46) Or _
                  (Asc(Right(.Text, 1)) = 32) Or _
                  (Asc(Right(.Text, 1)) = 11) Or _
                  (Asc(Right(.Text, 1)) = 13)) And _
                  (.End > .Start)
            .MoveEnd wdCharacter, -1
        Loop

ErrExit:
        If Len(.Text) = 0 Then
            MsgBox "Please select a reference.", _
                   vbExclamation, "Invalid selection"
            Exit Sub
        End If

        LookUp = .Text
    End With
    On Error GoTo 0

    With ActiveDocument
        ' Use WdRefTypeHeading to retrieve Headings
        RefList = .GetCrossReferenceItems(wdRefTypeHeading)
        For i = UBound(RefList) To 1 Step -1
            Ref = Trim(RefList(i))
            If InStr(1, Ref, LookUp, vbTextCompare) = 1 Then
                '  s = InStr(2, Ref, " ")
                ' t = InStr(2, Ref, Chr(9))
               '   If (s = 0) Or (t = 0) Then
               '       s = IIf(s > 0, s, t)
             '     Else
             '         s = IIf(s < t, s, t)
             '     End If
                If LookUp = Ref Then Exit For
            End If
        Next i

        If i Then
            Selection.InsertCrossReference ReferenceType:="Заголовок", _
                                           ReferenceKind:=wdContentText, _
                                           ReferenceItem:=CStr(i), _
                                           InsertAsHyperlink:=True, _
                                           IncludePosition:=False, _
                                           SeparateNumbers:=False, _
                                           SeparatorString:=" "
        Else
            MsgBox "A cross reference to """ & LookUp & """ couldn't be set" & vbCr & _
                   "because a paragraph with that number couldn't" & vbCr & _
                   "be found in the document.", _
                   vbInformation, "Invalid cross reference"
        End If
    End With
End Sub
