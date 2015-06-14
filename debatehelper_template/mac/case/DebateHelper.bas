Attribute VB_Name = "DebateHelper"
Sub AutoOpen()

    If System.OperatingSystem <> "Macintosh" Then
        MsgBox ("This is a Mac-only template. While on Windows, you will have very little functionality.")
        Exit Sub
    End If

    Call AutoOpenMac

    Selection.Collapse

End Sub

Sub AutoOpenMac()

    Call BuildKeyboardShortcuts

    Call BuildDHCaseToolbar
    Call ShowDHCaseToolbar

    ThisDocument.AttachedTemplate.Saved = True
    ActiveDocument.Save

End Sub

Sub AutoNew()

    Call SaveDocument

    Call BuildKeyboardShortcuts

    ThisDocument.AttachedTemplate.Saved = True
    ActiveDocument.Save

    Call BuildDHCaseToolbar
    Call ShowDHCaseToolbar

AutoNew_Error:
    If Err.Number = 4198 Then
        Resume Next
    End If

End Sub

Sub ShowDHCaseToolbar()

    CommandBars("DHCase").Visible = True
    CommandBars("DHCase").Position = msoBarTop

End Sub

Sub BuildDHCaseToolbar()

'Build dynamic toolbar with labels from registry

    Dim Toolbar As CommandBar
    Dim DHCaseToolbar As CommandBar
    Dim ButtonControl As CommandBarButton
    Dim MenuControl As CommandBarControl
    Dim MenuItem As CommandBarButton

    CustomizationContext = ThisDocument

    'Delete any preexisting toolbars to start from scratch
    For Each Toolbar In Application.CommandBars
        If Toolbar.Name = "DHCase" Then
            Toolbar.Protection = msoBarNoProtection
            Toolbar.Delete
        End If
    Next Toolbar

    'Create Toolbar
    Set DHCaseToolbar = CommandBars.Add(Name:="DHCase", Position:=msoBarTop)
    DHCaseToolbar.Visible = False

    'Create buttons

    'Update Styles
    Set ButtonControl = DHCaseToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Style = msoButtonIcon
    ButtonControl.OnAction = "UpdateStyles"
    ButtonControl.TooltipText = "Update styles from template"
    ButtonControl.FaceId = 254

    'New footnote
    Set ButtonControl = DHCaseToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Style = msoButtonIcon
    ButtonControl.OnAction = "InsertFootnote"
    ButtonControl.TooltipText = "F1 - Insert a footnote here"
    ButtonControl.FaceId = 3429
    ButtonControl.BeginGroup = True

    'Paste Card
    Set ButtonControl = DHCaseToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Style = msoButtonIcon
    ButtonControl.OnAction = "PasteAndCondense"
    ButtonControl.TooltipText = "F2 - Paste the clipboard here, as a single card"
    ButtonControl.FaceId = 22
    ButtonControl.BeginGroup = True

    'Word Count
    Set ButtonControl = DHCaseToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "F3 Word Count"
    ButtonControl.Style = msoButtonCaption
    ButtonControl.OnAction = "CountFont12"
    ButtonControl.TooltipText = "Count the number of words that are Size 12 font"
    ButtonControl.BeginGroup = True

    'Page Break
    Set ButtonControl = DHCaseToolbar.Controls.Add(Type:=msoControlButton, ID:=509)
    ButtonControl.TooltipText = "Insert Page Break"
    ButtonControl.BeginGroup = True

    'Styles Group
    Set MenuControl = DHCaseToolbar.Controls.Add(Type:=msoControlPopup)
    MenuControl.Caption = "Styles"
    MenuControl.BeginGroup = True

    'Case Style
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "F4 Case Style"
    MenuItem.Parameter = "Normal"
    MenuItem.OnAction = "FKeyAssignAction"

    'Card style
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "F5 Card Style"
    MenuItem.Parameter = "Card"
    MenuItem.OnAction = "FKeyAssignAction"

    'Big Card Evidence
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "F6 Big Evidence"
    MenuItem.OnAction = "Size12Underlined"
    MenuItem.BeginGroup = True

    'Small Card Evidence
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "F7 Small Evidence"
    MenuItem.OnAction = "Size7NotUnderlined"

    'Show Current Style
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Show Current Style (Alt+S)"
    MenuItem.OnAction = "ShowStyle"
    MenuItem.BeginGroup = True

    'Clear Formatting
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton, ID:=8099)
    MenuItem.Caption = "Clear Formatting"
    MenuItem.FaceId = 2822

    'Update Styles
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton)
    MenuItem.OnAction = "UpdateStyles"
    MenuItem.Caption = "Update Styles from Template"
    MenuItem.FaceId = 254
    MenuItem.BeginGroup = True

    'Comments Group
    Set ButtonControl = DHCaseToolbar.Controls.Add(Type:=msoControlButton, ID:=1594)
    ButtonControl.BeginGroup = True
    Set ButtonControl = DHCaseToolbar.Controls.Add(Type:=msoControlButton, ID:=1589)
    Set ButtonControl = DHCaseToolbar.Controls.Add(Type:=msoControlButton, ID:=1590)
    Set ButtonControl = DHCaseToolbar.Controls.Add(Type:=msoControlButton, ID:=1591)
    Set ButtonControl = DHCaseToolbar.Controls.Add(Type:=msoControlButton, ID:=1592)

    DHCaseToolbar.Width = 458
    DHCaseToolbar.Protection = msoBarNoCustomize + msoBarNoResize
    DHCaseToolbar.Visible = True

    'Save template if editing it
    ThisDocument.AttachedTemplate.Saved = True
End Sub

Sub FKeyAssignAction()

    Dim FKey As CommandBarControl
    Dim FKeyAction As String
    Dim StartLocation

    Set FKey = CommandBars.ActionControl
    If FKey Is Nothing Then Exit Sub
    FKeyAction = FKey.Parameter

    Select Case FKeyAction

    Case Is = "Normal"
        StartLocation = Selection.Start
        Selection.GoTo What:=wdGoToBookmark, Name:="\para"
        Selection.ClearFormatting
        Selection.Paragraphs.Style = "Normal"
        Selection.Start = StartLocation
        Selection.Collapse
    Case Is = "Card"
        StartLocation = Selection.Start
        Selection.GoTo What:=wdGoToBookmark, Name:="\para"
        Selection.ClearFormatting
        Selection.Paragraphs.Style = "Card"
        Selection.Start = StartLocation
        Selection.Collapse
    End Select

End Sub

Sub CountFont12()
'
' Count_Font_12 Macro


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
' Size7NotUnderlined Macro
'
    If Selection.Paragraphs.Style <> "Card" Then
        MsgBox "Style of evidence must be 'Card'. (The style of this paragraph is '" & Selection.Paragraphs.Style & "').", vbOKOnly, ""
        Exit Sub
    End If

    Selection.Font.Size = 7
    Selection.Font.UnderlineColor = wdColorAutomatic
    Selection.Font.Underline = wdUnderlineNone


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

Sub UpdateStyles()
    ActiveDocument.UpdateStyles
End Sub
