Sub ConvertTrackChangesToHardline()
    'Removes comments, converts track changes to hardline, and saves as new file.
    Dim wdDoc As Word.Document
    Set wdDoc = ActiveDocument
    
    Result = MsgBox("Do you want double underline/strikethrough? If not, single underline/strikethrough will be applied.", vbYesNo, "Yes No Example")

    Select Case Result
    Case vbYes
        Call TrackChangesToHardline_Single(wdDoc)
    Case vbNo
        Call TrackChangesToHardline_Double(wdDoc)
    End Select
    
    wdDoc.TrackRevisions = False
    With wdDoc
        If .Comments.Count > 0 Then .DeleteAllComments
    End With
    
    Call RemoveAddedAndDeleted(wdDoc)
    Call SaveAs_Hardline(wdDoc)

End Sub
Sub TrackChangesToHardline_Single(wdDoc As Document)
    'Converts tracked changes to hardline - single underline, single strikethrough.
    Dim rev As Word.Revision
    For Each rev In wdDoc.Revisions
        Select Case rev.Type
            Case wdRevisionDelete
                rev.Range.Font.StrikeThrough = True
                rev.Reject
            Case wdRevisionInsert
                rev.Range.Underline = wdUnderlineSingle
        End Select
    Next
    wdDoc.AcceptAllRevisions
End Sub
Sub TrackChangesToHardline_Double(wdDoc As Document)
    'Converts tracked changes to hardline - double underline, double strikethrough.
    Dim rev As Word.Revision
    For Each rev In wdDoc.Revisions
        Select Case rev.Type
            Case wdRevisionDelete
                rev.Range.Font.DoubleStrikeThrough = True
                rev.Reject
            Case wdRevisionInsert
                rev.Range.Underline = wdUnderlineDouble
        End Select
    Next
    wdDoc.AcceptAllRevisions
End Sub
Sub RemoveAddedAndDeleted(wdDoc As Document)
    'Removes text that was added and also deleted in the same tracked changes session, for both single and double lines.
    'For example, if Person A added text and Person B deleted it, the text would be removed.
    Set rng = wdDoc.Range
    
    'Single underline and single strikethrough text.
    rng.Find.ClearFormatting
    With rng.Find.Font
        .Underline = wdUnderlineSingle
        .StrikeThrough = True
        .DoubleStrikeThrough = False
    End With
    rng.Find.Replacement.ClearFormatting
    With rng.Find
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    rng.Find.Execute Replace:=wdReplaceAll
    
    'Double underline and double strikethrough text.
    rng.Find.ClearFormatting
    With rng.Find.Font
        .Underline = wdUnderlineDouble
        .StrikeThrough = False
        .DoubleStrikeThrough = True
    End With
    rng.Find.Replacement.ClearFormatting
    With rng.Find
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    rng.Find.Execute Replace:=wdReplaceAll
End Sub
Sub SaveAs_Hardline(wdDoc As Document)
    Dim docName As String
    'Append filename with:
    append = "_Hardline"
        
    docName = wdDoc.FullName
    docName = Left(docName, (InStr(docName, ".") - 1))
    docName = docName & append
    
    wdDoc.SaveAs2 FileName:=docName, FileFormat:=wdFormatDocumentDefault
    
    MsgBox ("Saved hardline copy as new file.")
End Sub
