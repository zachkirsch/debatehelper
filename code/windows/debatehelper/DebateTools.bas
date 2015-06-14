Attribute VB_Name = "DebateTools"
Sub GoToTop()

    On Error GoTo Handler

    Selection.GoTo What:=wdGoToSection, Which:=wdGoToFirst

    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub

Sub GoToBottom()

    On Error GoTo Handler

    Selection.MoveDown Unit:=wdParagraph, Count:=ActiveDocument.Paragraphs.Count, Extend:=wdMove

    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub

Sub CopyCard()


    Dim CurrentDoc
    Dim OLevel

    'Save active document name
    On Error GoTo Handler

    CurrentDoc = ActiveDocument.Name


    'Turn off screen updating for the heavy-lifting
    Application.ScreenUpdating = False

    'If text is selected, copy and send it.  Add a return if not in the selection.
    If Selection.End > Selection.Start Then
        Selection.Copy

    End If

    'If nothing is selected, select the current card, block, hat or pocket
    'Move to start of current paragraph and collapse the selection
    Selection.StartOf Unit:=wdParagraph
    Selection.Collapse

    'Move backwards through each paragraph to find the first tag, block title, hat, pocket or the top of the document
    Do While True
        If Selection.Paragraphs.OutlineLevel = wdOutlineLevel8 Then Exit Do    'Heading 8
        If Selection.Paragraphs.OutlineLevel = wdOutlineLevel9 Then Exit Do    'Heading 9
        If Selection.Start <= ActiveDocument.Range.Start Then    'Top of document
            MsgBox "Nothing found to send"
            Exit Sub
        End If
        Selection.Move Unit:=wdParagraph, Count:=-1
    Loop

    'Get current outline level
    OLevel = Selection.Paragraphs.OutlineLevel

    'Extend selection until hitting the bottom or a bigger outline level
    Selection.MoveEnd Unit:=wdParagraph, Count:=1
    Do While True
        Selection.MoveEnd Unit:=wdParagraph, Count:=1
        If Selection.End + 1 >= ActiveDocument.Range.End Then Exit Do    'Bottom of doc
        If Selection.Paragraphs.Last.OutlineLevel <= OLevel Then
            Selection.MoveEnd Unit:=wdParagraph, Count:=-1
            Exit Do    'Bigger Outline Level
        End If
    Loop

    'Copy the unit
    Selection.Copy

    'Reset Selection
    Selection.Collapse


    Application.ScreenUpdating = True

    Exit Sub


    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub

Sub PasteAndCondense()
'

    On Error GoTo Handler

    Application.ScreenUpdating = False
    Selection.Delete
    Selection.Collapse

    Dim CurrentStyle
    CurrentStyle = Selection.Style

    Dim oRng As Range, oStart As Range
    Set oRng = Selection.Range
    Set oStart = Selection.Range
    With oRng
        .PasteSpecial _
                DataType:=wdPasteText, _
                Placement:=wdInLine
        .Start = oStart.Start
        .Select
    End With

    Selection.ClearFormatting
    Call Condense
    Selection.Style = ActiveDocument.Styles(CurrentStyle)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True


    Exit Sub

Handler:

    If Err.Number = 5342 Then
        MsgBox "There's a problem with the current clipboard. Try pasting " & _
               "the text normally, then select it and click condense."
    Else
        MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"
    End If

End Sub

Sub Condense()
    On Error Resume Next
    Dim CurrentStyle
    CurrentStyle = Selection.Style

    Selection.ClearFormatting

    'Condense
    Dim CondenseRange As Range

    'If selection is too short, exit
    If Len(Selection) < 2 Then Exit Sub

    'If end of selection is a line break, shorten it
    If Selection.Characters.Last = vbCr Then Selection.MoveEnd , -1

    'Save selection
    Set CondenseRange = Selection.Range

    'Condense everything except hard returns
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Wrap = wdFindStop

        .Text = "^m"                    'page breaks
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll

        .Text = "^t"                    'tabs
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll

        .Text = "^s"                    'non-breaking space
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll

        .Text = "^b"                    'section break
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll

        .Text = "^l"                    'new line
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll

        .Text = "^n"                    'column break
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
    End With

    With Selection.Find
        .Text = "^p"
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll

        .Text = "  "
        .Replacement.Text = " "

        While InStr(Selection, "  ")
            .Execute Replace:=wdReplaceAll
        Wend

        If Selection.Characters(1) = " " And _
           Selection.Paragraphs(1).Range.Start = Selection.Start Then _
           Selection.Characters(1).Delete
    End With

    'Clear find dialogue
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting

    'Change style to evidence
    Selection.Style = ActiveDocument.Styles(CurrentStyle)

End Sub
Sub InsertAutoText(autotextName As String)

    On Error GoTo Handler

    Dim TemplatePath As String
    TemplatePath = Options.DefaultFilePath(wdUserTemplatesPath) & "\DebateHelper.dotm"

    Application.Templates( _
            TemplatePath). _
            BuildingBlockEntries(autotextName).Insert Where:=Selection.Range, _
                                                      RichText:=True

Handler:
    If Err.Number = 5941 Then
        MsgBox "No autotext to insert. Please re-install DebateHelper."
        End
    End If


End Sub


Sub MoveToNewLine()

    On Error GoTo Handler

    If Selection.Start <> Selection.Paragraphs(1).Range.Start Then
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.MoveDown Unit:=wdParagraph, Count:=1
    End If
    If Selection.Start <> Selection.Paragraphs(1).Range.Start Then
        Selection.TypeParagraph
    End If

    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub

Sub ShowStyle()

    Dim StyleName As String
    On Error GoTo Handler

    Selection.Collapse
    StyleName = Selection.Paragraphs.Style

    If Selection.Paragraphs.Style = wdStyleHeading1 Then StyleName = "Section Level 1"
    If Selection.Paragraphs.Style = wdStyleHeading2 Then StyleName = "Section Level 2"
    If Selection.Paragraphs.Style = wdStyleHeading3 Then StyleName = "Section Level 3"
    If Selection.Paragraphs.Style = wdStyleHeading4 Then StyleName = "Block"
    If Selection.Paragraphs.Style = wdStyleHeading5 Then StyleName = "Responses Level 1"
    If Selection.Paragraphs.Style = wdStyleHeading6 Then StyleName = "Responses Level 2"
    If Selection.Paragraphs.Style = wdStyleHeading7 Then StyleName = "Responses Level 3"
    If Selection.Paragraphs.Style = wdStyleHeading8 Then StyleName = "Tag"
    If Selection.Paragraphs.Style = wdStyleHeading9 Then StyleName = "Sub Tag"

    Application.StatusBar = "Style: " & StyleName
    Exit Sub

    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub

Sub UpdateFields()


'Update Fields
    Dim oStory As Range
    On Error GoTo Handler

    For Each oStory In ActiveDocument.StoryRanges
        oStory.Fields.Update
        If oStory.StoryType <> wdMainTextStory Then
            While Not (oStory.NextStoryRange Is Nothing)
                Set oStory = oStory.NextStoryRange
                oStory.Fields.Update
            Wend
        End If
    Next oStory
    Set oStory = Nothing


    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub


Sub InsertCard()

    On Error GoTo Handler

    Call MoveToNewLine
    Selection.Style = ActiveDocument.Styles("Normal")
    'Define the required building block entry

    If GetSetting("DebateHelper", "Main", "UseBlockedCite", False) = True Then
        Call InsertAutoText("CardWithBlockedCite")
    Else
        Call InsertAutoText("CardWithoutBlockedCite")
    End If

    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub

Sub InsertCardWithPreviousCite()

'This inserts an autotext built with a styleref to insert the previous citation, then retreives the citation-part of the tag

'Save Current Location
    Dim StartLocation
    StartLocation = Selection.Start

    'Find previous cite
    With Selection.Find
        .ClearFormatting
        .Text = ""
        .Wrap = wdFindStop
        .Format = True
        .Style = ActiveDocument.Styles("Citation")
        .Forward = False
        .Execute
    End With

    'If nothing found, exit
    If Selection.Find.Found = False Then
        MsgBox "No cite found. Make sure the previous card's cite is using the 'Citation' style."
        Exit Sub
    Else
        Selection.Start = StartLocation
    End If

    Call MoveToNewLine
    Selection.ClearFormatting


    With Selection.Find
        .ClearFormatting
        .Text = "("
        .Wrap = wdFindStop
        .Format = True
        .Style = ActiveDocument.Styles("Tag")
        .Forward = False
        .Execute
    End With
    Selection.MoveEnd Unit:=wdParagraph, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    
    Selection.Fields.Unlink
    Dim AuthorInfo As String
    AuthorInfo = Selection.Text

    'Return to original location
    Selection.Start = StartLocation

    With Selection.Find
        .ClearFormatting
        .Text = ""
        .Wrap = wdFindStop
        .Format = True
        .Style = ActiveDocument.Styles("Citation")
        .Forward = False
        .Execute
    End With
    
    Selection.HomeKey Unit:=wdLine
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    
    Dim CiteInfo As String
    Selection.Copy

    'Return to original location
    Selection.Start = StartLocation
    
    Call InsertAutoText("Card_With_Previous_Cite")

    With Selection.Find
        .ClearFormatting
        .Text = "("
        .Wrap = wdFindStop
        .Format = True
        .Style = ActiveDocument.Styles("Tag")
        .Forward = False
        .Execute
    End With
    Selection.MoveEnd Unit:=wdParagraph, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Delete
    Selection.InsertAfter (AuthorInfo)

    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Delete
    Selection.Paste
    
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveRight Unit:=wdWord, Count:=1
    
    Call MoveToNewLine 'Move cursor to end of paragraph
    
    'Clear clipboard
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Copy
    Selection.MoveLeft Unit:=wdCharacter, Count:=1

    Exit Sub
    
Handler:
    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub
Sub InsertBlock()

    On Error GoTo Handler

    Call MoveToNewLine
    Selection.Style = ActiveDocument.Styles("Normal")

    Selection.InsertBreak Type:=wdPageBreak
    Selection.Style = ActiveDocument.Styles("Block")
    Selection.TypeText Text:="A2: Argument"
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.Style = ActiveDocument.Styles("Resonses Level 1")
    Selection.MoveDown Unit:=wdLine, Count:=1

    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub

Sub UpdateTOC()
    On Error Resume Next
    ActiveDocument.TablesOfContents(1).Update
End Sub

Sub NewRebuttal()
'Creates a new Speech document
'August 2010 by Aaron Hardy
'Edited by Zach Kirsch


'Trap for user cancelling the save
    On Error GoTo Handler

    'Create filename
    Dim h As String
    Dim FileName As String
    If Hour(Now) > 12 Then h = Hour(Now) - 12 & "PM"
    If Hour(Now) <= 12 Then h = Hour(Now) & "AM"
    FileName = "Rebuttal " & Month(Now) & "-" & Day(Now) & " " & h

    Documents.Add (Options.DefaultFilePath(wdUserTemplatesPath) & "\Rebuttal.dotm")

    With Application.Dialogs(wdDialogFileSaveAs)
        .Format = Word.WdSaveFormat.wdFormatXMLDocument
        .Name = FileName
        .Show
    End With

    Exit Sub


Handler:
    If Err.Number = 4198 Then
        With Application.Dialogs(wdDialogFileSaveAs)
            .Format = Word.WdSaveFormat.wdFormatXMLDocument
            .Name = "*"
            .Show
        End With
    End If

End Sub

Sub SendToRebuttal()

'Edited by Zach Kirsch
'Sends content to the Speech doc.  Sends currently selected text,
'or if nothing is selected, the current tag, block, hat, or pocket
'If in reading view, enters a stopped reading marker at the current location
'September 2011 by Aaron Hardy
'Updated to reflect v. 4 Heading Level changes
'Also fixed infinite loop bug on Speech Doc Creation

    Dim CurrentDoc
    Dim Doc
    Dim RebuttalDoc As Document
    Dim FoundDoc
    Dim OLevel
    Dim CreateRebuttal
    Dim d As Document

    On Error GoTo Handler

    'Save active document name
    CurrentDoc = ActiveDocument.Name

    'Turn off screen updating for the heavy-lifting
    Application.ScreenUpdating = False

    'If nothing is selected, select the current card, block, hat or pocket
    If Selection.Start = Selection.End Then
        'Move to start of current paragraph and collapse the selection
        Selection.StartOf Unit:=wdParagraph
        Selection.Collapse

        ' If cursor is at a block or response, select the block
        If Selection.Paragraphs.OutlineLevel > wdOutlineLevel3 And Selection.Paragraphs.OutlineLevel < wdOutlineLevel7 Then
            'Move backwards through each paragraph to find the first block title or the top of the document
            Do While True
                If Selection.Paragraphs.OutlineLevel = wdOutlineLevel4 Then Exit Do    'Heading 4
                If Selection.Start <= ActiveDocument.Range.Start Then    'Top of document
                    MsgBox "Nothing found to send"
                    Exit Sub
                End If
                Selection.Move Unit:=wdParagraph, Count:=-1
            Loop
        Else ' Then look for a card
            'Move backwards through each paragraph to find the first tag or the top of the document
            Do While True
                If Selection.Paragraphs.OutlineLevel = wdOutlineLevel8 Then Exit Do    'Heading 8
                If Selection.Paragraphs.OutlineLevel = wdOutlineLevel9 Then Exit Do    'Heading 9
                If Selection.Start <= ActiveDocument.Range.Start Then    'Top of document
                    MsgBox "Nothing found to send"
                    Exit Sub
                End If
                Selection.Move Unit:=wdParagraph, Count:=-1
            Loop
        End If
        
        'Get current outline level
        OLevel = Selection.Paragraphs.OutlineLevel

        'Extend selection until hitting the bottom or a greater/equal outline level
        Selection.MoveEnd Unit:=wdParagraph, Count:=1
        Do While True
            Selection.MoveEnd Unit:=wdParagraph, Count:=1
            If Selection.End + 1 >= ActiveDocument.Range.End Then Exit Do    'Bottom of doc
            If Selection.Paragraphs.Last.OutlineLevel <= OLevel Then
                Selection.MoveEnd Unit:=wdParagraph, Count:=-1
                Exit Do    'Bigger Outline Level
            End If
        Loop
    End If
    
    'Copy the unit
    Selection.Copy
    
RebuttalDocCheck:
    FoundDoc = 0
    'If there's an active speech doc, use it
    'Check if a document with "rebuttal.dotm" template is open.
    For Each Doc In Application.Documents
        If InStr(LCase(Doc.AttachedTemplate), "rebuttal") Then
            FoundDoc = FoundDoc + 1
            If FoundDoc = 1 Then Set RebuttalDoc = Doc
        End If
    Next Doc


    'If no Rebuttal doc is found, prompt whether to create one.
    'If yes, create a new document based on the current template to save, then retry
    If FoundDoc = 0 Then
        CreateRebuttal = MsgBox("Rebuttal document is not open - create one?", vbYesNo, "Create Rebuttal?")
        If CreateRebuttal = vbNo Then
            Exit Sub
        Else
            'Open New Rebuttal Doc
            Call NewRebuttal

            'Switch focus back after save
            Documents(CurrentDoc).Activate
            GoTo RebuttalDocCheck
        End If
    End If

    'If multiple Rebuttal docs are open, warn the user.
    If FoundDoc > 1 Then
        If MsgBox("There are " + CStr(FoundDoc) + " possible rebuttal documents open. This feature only " + _
                  "works when there is only 1 open. Would you like to close the other rebuttal documents and " + _
                  "start over?", vbYesNo) = vbNo Then
            Exit Sub
        Else
            For Each Doc In Application.Documents
                If InStr(LCase(Doc.AttachedTemplate), "rebuttal") Then
                    Doc.Saved = True
                    Doc.Close
                End If
            Next Doc

            'Open New Rebuttal Doc
            Call NewRebuttal

            For Each Doc In Application.Documents
                If InStr(LCase(Doc.AttachedTemplate), "rebuttal") Then Set RebuttalDoc = Doc
            Next Doc

            'Switch focus back after save
            Documents(CurrentDoc).Activate
        End If

    End If

    Windows(RebuttalDoc).Activate
    Call MoveToNewLine
    Windows(CurrentDoc).Activate

    'Paste it
    RebuttalDoc.ActiveWindow.Selection.PasteAndFormat (wdFormatOriginalFormatting)

    Set RebuttalDoc = Nothing

    Application.ScreenUpdating = True

    Exit Sub

    'Error handler for user cancelling the save.  4198 is a generic runtime error.
Handler:
    If Err.Number = 4198 Then
        MsgBox "You messed up saving the Rebuttal document, start over."
        Exit Sub
    Else: MsgBox Err.Number & Err.Description
    End If

End Sub




