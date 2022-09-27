Sub PrettyPaper()

'
' PrettyPaper Macro
' Author: Ozelot Vanilla
'


' ------------------------------------------------------------------------------------------------
' Var declare
    
    Dim cursorPosition As Integer
    cursorPosition = Selection.Start
' ------------------------------------------------------------------------------------------------


' ------------------------------------------------------------------------------------------------
' Two space to one space

    With Selection.Find
            .Text = "  "
            .ClearFormatting
            With .Replacement
                .Text = " "
                .ClearFormatting
            End With
            .Execute Replace:=wdReplaceAll
        End With
' ------------------------------------------------------------------------------------------------

' ------------------------------------------------------------------------------------------------
' Update Fields

    With ActiveDocument
        For Each Field In .Fields
            Field.Locked = False
        Next
        .Fields.Update
    End With
' ------------------------------------------------------------------------------------------------


' ------------------------------------------------------------------------------------------------
' Page Layout Setting
    
    With ActiveDocument.PageSetup
        .LeftMargin = InchesToPoints(1)
        .RightMargin = InchesToPoints(1)
        .TopMargin = InchesToPoints(1)
        .BottomMargin = InchesToPoints(1)
    End With
' ------------------------------------------------------------------------------------------------


' ------------------------------------------------------------------------------------------------
' Style of text
    
    Selection.WholeStory
    With Selection.Font
        .NameFarEast = "UD Digi Kyokasho NK-R"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Name = "Times New Roman"
        .Size = 12
    End With
    Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceDouble
    With Selection.Find
        .Text = ""
        .ClearFormatting
        .ParagraphFormat.Alignment = wdAlignParagraphJustify
        With .Replacement
            .Text = ""
            .ClearFormatting
            .ParagraphFormat.Alignment = wdAlignParagraphLeft
        End With
        .Execute Replace:=wdReplaceAll
    End With
' ------------------------------------------------------------------------------------------------


' ------------------------------------------------------------------------------------------------
' Check Placeholder (Text)
    Dim countPlaceholder As Integer
    With ActiveDocument.Content.Find
        Do While .Execute(FindText:="(Placeholder)") = True
            countPlaceholder = countPlaceholder + 1
        Loop
    End With
' ------------------------------------------------------------------------------------------------

' ------------------------------------------------------------------------------------------------
' Ending

    Selection.Start = cursorPosition
    Selection.End = cursorPosition
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
' ------------------------------------------------------------------------------------------------


' ------------------------------------------------------------------------------------------------
' End Message
    
    Dim stillPlaceholder As String
    Dim returnEndMsg As Integer
    stillPlaceholder = ""
    
    If countPlaceholder <> 0 Then
        stillPlaceholder = "Found placeholder:" + Str(countPlaceholder)
    End If
    
    returnEndMsg = MsgBox("Check over. Is the result/previous result right?" + vbCrLf + vbCrLf + stillPlaceholder + vbCrLf, vbYesNo, "Finished!")
    
    If returnEndMsg = vbNo Then
        x = MsgBox("If the macro damage your document," + vbCrLf + vbCrLf + "DO NOT SAVE or EXIT, use Ctrl + Z.", , "Need contact to author?")
    End If
    
    
    
    
' ------------------------------------------------------------------------------------------------

End Sub
