Attribute VB_Name = "DebateHelper"

Sub AutoNew()
    On Error GoTo AutoNew_Error

    Call SaveDocument

    Call BuildKeyboardShortcuts

    ThisDocument.AttachedTemplate.Saved = True
    ActiveDocument.Save

AutoNew_Error:
    If Err.Number = 4198 Then
        Resume Next
    End If

End Sub

Sub AutoOpen()

    If System.OperatingSystem = "Macintosh" Then
        MsgBox ("This is a Windows only template. While on Mac, there will be limited functionality.")
        Exit Sub
    End If
    Call AutoOpenWindows

End Sub

Sub AutoOpenWindows()

    Call BuildKeyboardShortcuts

    ThisDocument.AttachedTemplate.Saved = True
    ActiveDocument.Save

End Sub

Sub BuildKeyboardShortcuts()


'Change keyboard shortcuts in template

    Application.CustomizationContext = ActiveDocument.AttachedTemplate
    KeyBindings.ClearAll
    KeyBindings.Add wdKeyCategoryMacro, "InsertFootnote", BuildKeyCode(wdKeyF1)
    KeyBindings.Add wdKeyCategoryMacro, "PasteAndCondense", BuildKeyCode(wdKeyF2)
    KeyBindings.Add wdKeyCategoryMacro, "CountFont12", BuildKeyCode(wdKeyF3)
    KeyBindings.Add wdKeyCategoryStyle, "Normal", BuildKeyCode(wdKeyF4)
    KeyBindings.Add wdKeyCategoryStyle, "Card", BuildKeyCode(wdKeyF5)
    KeyBindings.Add wdKeyCategoryMacro, "DebateHelper.Size12Underlined", BuildKeyCode(wdKeyF6)
    KeyBindings.Add wdKeyCategoryMacro, "DebateHelper.Size7NotUnderlined", BuildKeyCode(wdKeyF7)
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, _
                                          wdKeyS), KeyCategory:=wdKeyCategoryMacro, _
                                          Command:="DebateHelper.ShowStyle"

    Application.CustomizationContext = ThisDocument


End Sub

Sub PasteAndCondense()
'

    Application.ScreenUpdating = False

    Call MoveToNewLine

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
    Selection.Style = ActiveDocument.Styles("Card")
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True

End Sub

Sub Condense()


    Selection.ClearFormatting

    'Condense
    Dim CondenseRange As Range
    Dim ParagraphIntegrity As Boolean
    Dim Pilcrows As Boolean

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
    Selection.Style = ActiveDocument.Styles("Card")

End Sub
Sub CountFont12()
'
' Count_Font_12 Macro
'
'
    Dim Answer
    Dim i As Integer
    Dim TimeFormat As String
    i = ActiveDocument.Words.Count * 0.047
    If i > 60 Then
        seconds = Fix(i)
        mins = Fix(i / 60)
        secs = i - (mins * 60)
        If secs > 30 Then mins = mins + 1
        TimeFormat = mins & " minutes"
        If mins = 1 Then TimeFormat = "1 minute"
    Else
        TimeFormat = i & " seconds"
    End If

    Answer = MsgBox("This word count should take about " & TimeFormat & "." & vbCrLf _
                  & "Proceed?", vbYesNo)
    If Answer = vbNo Then Exit Sub
    ActiveDocument.Save

    Dim lngWord As Long
    Dim lngCountIt As Long
    Const ChosenFontSize As Integer = 12

    For lngWord = 1 To ActiveDocument.Words.Count
        'Ignore any document "Words" that aren't real words (CR, LF etc)
        If Len(Trim(ActiveDocument.Words(lngWord))) > 1 Or _
           IsLetter(ActiveDocument.Words(lngWord)) Then
            If ActiveDocument.Words(lngWord).Font.Size = ChosenFontSize Then
                lngCountIt = lngCountIt + 1
            End If
        End If
    Next lngWord

    MsgBox "Words (size 12): " & lngCountIt, vbOKOnly, "Case Word Count"

End Sub

Function IsLetter(strValue As String) As Boolean
    Dim intPos As Integer
    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, intPos, 1))
        Case 65 To 90, 97 To 122
            IsLetter = True
        Case Else
            IsLetter = False
            Exit For
        End Select
    Next
End Function

Sub NewCase()

    Documents.Add Application.NormalTemplate.Path & ":Case.dotm"

End Sub

Sub InsertFootnote()

    Selection.Footnotes.Add Range:=Selection.Range, Reference:=""

End Sub

Sub MoveToNewLine()

    If Selection.Start <> Selection.Paragraphs(1).Range.Start Then
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.MoveDown Unit:=wdParagraph, Count:=1
    End If
    If Selection.Start <> Selection.Paragraphs(1).Range.Start Then
        Selection.TypeParagraph
    End If

End Sub

Sub SaveDocument()

    With Application.Dialogs(wdDialogFileSaveAs)
        .Format = Word.WdSaveFormat.wdFormatXMLDocument
        .Name = "*"
        .Show
    End With

End Sub

Sub ShowStyle()

    Application.StatusBar = "Style: " & Selection.Paragraphs.Style

End Sub

Sub Size12Underlined()
'
    If Selection.Paragraphs.Style <> "Card" Then
        MsgBox "Style of evidence must be 'Card'. (The style of this paragraph is '" & Selection.Paragraphs.Style & "').", vbOKOnly, ""
        Exit Sub
    End If

    Selection.Font.Size = 12
    Selection.Font.UnderlineColor = wdColorAutomatic
    Selection.Font.Underline = wdUnderlineSingle
End Sub

Sub Size7NotUnderlined()

'
    If Selection.Paragraphs.Style <> "Card" Then
        MsgBox "Style of evidence must be 'Card'. (The style of this paragraph is '" & Selection.Paragraphs.Style & "').", vbOKOnly, ""
        Exit Sub
    End If

    Selection.Font.Size = 7
    Selection.Font.UnderlineColor = wdColorAutomatic
    Selection.Font.Underline = wdUnderlineNone


End Sub

Sub UpdateFields()

'Update Fields
    Dim oStory As Range
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

End Sub

