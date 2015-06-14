Attribute VB_Name = "Startup"
Sub AutoOpen()

    If System.OperatingSystem <> "Macintosh" Then
        MsgBox ("This is a Mac-only template. While on Windows, you will have very little functionality.")
        Exit Sub
    End If

    Call AutoOpenMac

    Selection.Collapse

End Sub
Sub AutoOpenMac()

On Error GoTo Handler

    Call BuildKeyboardShortcuts

    Call BuildToolbar
    Call UpdateTOC
    Call BuildKeyboardShortcuts
    ThisDocument.AttachedTemplate.Saved = True

    'Refresh document styles from template if setting checked and not editing template itself
    If GetSetting("DebateHelper", "Main", "AutoUpdateStyles", True) = True Then ActiveDocument.UpdateStyles

    'Check for new messages every three days
    If DateDiff("d", GetSetting("DebateHelper", "Main", "LastMessageCheck", "8/15/2004 11:58:25 AM"), Now) > 3 Then MessageCheck

    'Check for updates if auto-checking enabled in settings (checks in Ribbon startup)
    If GetSetting("DebateHelper", "Main", "AutoUpdateCheck", True) = True Then UpdateCheck

    'Analytics
    UpdateAnalytics

    ActiveDocument.Save

    Exit Sub
    
Handler:
    If Err.Number = 4198 Then
        MsgBox "Error with saving!", vbExclamation, "DebateHelper Error"
        Exit Sub
    End If
    
    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"
End Sub

Sub AutoNew()

    If System.OperatingSystem <> "Macintosh" Then
        MsgBox "This template is not intended for Windows."
        ActiveDocument.Close
        Exit Sub
    End If

    Call SaveDocument
    Call BuildToolbar
    Call BuildKeyboardShortcuts

    ActiveDocument.Save

End Sub


Sub AttachTemplate()

    If InStr(ActiveDocument.Name, "dotm") Then Exit Sub

    Dim OriginalTemplate As String
    OriginalTemplate = ActiveDocument.AttachedTemplate

    ActiveDocument.AttachedTemplate = Application.NormalTemplate.Path & ":DebateHelper.dotm"

    If ActiveDocument.AttachedTemplate <> OriginalTemplate Then
        Application.Documents.Open (ActiveDocument.FullName)
    End If

    ActiveDocument.UpdateStyles

    On Error Resume Next

End Sub
Sub SaveDocument()

    On Error GoTo Handler

    With Application.Dialogs(wdDialogFileSaveAs)
        .Format = Word.WdSaveFormat.wdFormatXMLDocument
        .Name = "*"
        .Show
    End With

Handler:
    If Err.Number = 4198 Then
        With Application.Dialogs(wdDialogFileSaveAs)
            .Format = Word.WdSaveFormat.wdFormatXMLDocument
            .Name = "*"
            .Show
        End With
    End If

End Sub

Sub BuildKeyboardShortcuts()
'Change keyboard shortcuts in template

    Application.CustomizationContext = ActiveDocument.AttachedTemplate

    'KeyBindings.ClearAll
    KeyBindings.Add wdKeyCategoryStyle, "Section Title 1", BuildKeyCode(wdKeyF1)
    KeyBindings.Add wdKeyCategoryStyle, "Section Title 2", BuildKeyCode(wdKeyF2)
    KeyBindings.Add wdKeyCategoryStyle, "Section Title 3", BuildKeyCode(wdKeyF3)
    KeyBindings.Add wdKeyCategoryStyle, "Block", BuildKeyCode(wdKeyF4)
    KeyBindings.Add wdKeyCategoryStyle, "Tag", BuildKeyCode(wdKeyF5)
    KeyBindings.Add wdKeyCategoryStyle, "Citation", BuildKeyCode(wdKeyF6)
    KeyBindings.Add wdKeyCategoryStyle, "Normal", BuildKeyCode(wdKeyF7)
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF8), KeyCategory:=wdKeyCategoryMacro, Command:="InsertCard"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF9), KeyCategory:=wdKeyCategoryMacro, Command:="InsertCardWithPreviousCite"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF10), KeyCategory:=wdKeyCategoryMacro, Command:="CopyCurrentCard"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF11), KeyCategory:=wdKeyCategoryMacro, Command:="SendToRebuttal"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF12), KeyCategory:=wdKeyCategoryMacro, Command:="PasteAndCondense"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, _
                                          wdKeyS), KeyCategory:=wdKeyCategoryMacro, _
                                          Command:="ShowCurrentStyle"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, _
                                          wdKeyR), KeyCategory:=wdKeyCategoryMacro, _
                                          Command:="SendToRebuttal"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, _
                                          wdKeyC), KeyCategory:=wdKeyCategoryMacro, _
                                          Command:="OpenCitationMaker"

    Application.CustomizationContext = ThisDocument

    Exit Sub

End Sub

Sub BuildToolbar()
    On Error Resume Next
    'Build dynamic toolbar with labels from registry

    Dim Toolbar As CommandBar
    Dim DHToolbar As CommandBar
    Dim ButtonControl As CommandBarButton
    Dim MenuControl As CommandBarControl
    Dim MenuItem As CommandBarButton

    CustomizationContext = ThisDocument

    'Delete any preexisting toolbars to start from scratch

    For Each Toolbar In Application.CommandBars
        If Toolbar.Name = "DH" Then
            Toolbar.Protection = msoBarNoProtection
            Toolbar.Delete
        End If
    Next Toolbar

    'Create Toolbar
    Set DHToolbar = CommandBars.Add(Name:="DH", Position:=msoBarTop)
    DHToolbar.Visible = False

    'Create buttons

    'Update Styles
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Style = msoButtonIcon
    ButtonControl.OnAction = "UpdateStyles"
    ButtonControl.TooltipText = "Update Styles from Template"
    ButtonControl.FaceId = 254

    'Headings Group
    Set MenuControl = DHToolbar.Controls.Add(Type:=msoControlPopup)
    MenuControl.Caption = "Headings"
    MenuControl.TooltipText = "Heading Styles"

    'Section Title 1
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "F1 Section Title 1"
    MenuItem.Parameter = "Section Title 1"
    MenuItem.OnAction = "FKeyAssignAction"

    'Section Title 2
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "F2 Section Title 2"
    MenuItem.Parameter = "Section Title 2"
    MenuItem.OnAction = "FKeyAssignAction"

    'Section Title 3
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "F3 Section Title 3"
    MenuItem.Parameter = "Section Title 3"
    MenuItem.OnAction = "FKeyAssignAction"

    'Block Style
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "F4 Block"
    MenuItem.Parameter = "Block"
    MenuItem.OnAction = "FKeyAssignAction"

    'Response Level 1 Style
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "       Response Level 1"
    MenuItem.Parameter = "Response Level 1"
    MenuItem.OnAction = "FKeyAssignAction"

    'Response Level 2 Style
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "       Response Level 2"
    MenuItem.Parameter = "Response Level 2"
    MenuItem.OnAction = "FKeyAssignAction"

    'Response Level 3 Style
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "       Response Level 3"
    MenuItem.Parameter = "Response Level 3"
    MenuItem.OnAction = "FKeyAssignAction"

    'Show Current Style
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Show Current Style (Alt+S)"
    MenuItem.OnAction = "ShowCurrentStyle"
    MenuItem.BeginGroup = True

    'Restart Numbering
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Style = msoButtonIcon
    ButtonControl.OnAction = "RestartNumbering"
    ButtonControl.TooltipText = "Restart Numerbing"
    ButtonControl.FaceId = 3362

    'Card Group
    Set MenuControl = DHToolbar.Controls.Add(Type:=msoControlPopup)
    MenuControl.Caption = "Card"
    MenuControl.TooltipText = "Card Styles"

    'Tag Style
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "F5 Tag"
    MenuItem.Parameter = "Tag"
    MenuItem.OnAction = "FKeyAssignAction"

    'Sub Tag Style
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "     Sub Tag"
    MenuItem.Parameter = "Sub Tag"
    MenuItem.OnAction = "FKeyAssignAction"

    'Citation Style
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "F6 Citation"
    MenuItem.Parameter = "Citation"
    MenuItem.OnAction = "FKeyAssignAction"

    'Evidence (Normal) Style
    Set MenuItem = MenuControl.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "F7 Evidence"
    MenuItem.Parameter = "Evidence"
    MenuItem.OnAction = "FKeyAssignAction"

    'New Block
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "New Block"
    ButtonControl.Style = msoButtonCaption
    ButtonControl.OnAction = "InsertBlock"
    ButtonControl.TooltipText = "Insert a new block with a response"
    ButtonControl.BeginGroup = True

    'New Card
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "F8 New Card"
    ButtonControl.Style = msoButtonCaption
    ButtonControl.OnAction = "InsertCard"
    ButtonControl.TooltipText = "Insert a new card"

    ' Generate Cite Wizard
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Style = msoButtonIcon
    ButtonControl.OnAction = "OpenCitationMaker"
    ButtonControl.TooltipText = "Generate a citation using a wizard (Alt+C)"
    ButtonControl.FaceId = 3272

    'New Card with Previous Cite
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "F9 Same Cite"
    ButtonControl.Style = msoButtonCaption
    ButtonControl.OnAction = "InsertCardWithPreviousCite"
    ButtonControl.TooltipText = "Insert a new card with the same citation as the previous card"

    'Copy Card
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "F10 Copy Card"
    ButtonControl.Style = msoButtonCaption
    ButtonControl.OnAction = "CopyCurrentCard"
    ButtonControl.TooltipText = "Copy the current card"
    ButtonControl.BeginGroup = True

    ' Send to Rebuttal
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.FaceId = 8694
    ButtonControl.Style = msoButtonIconAndCaption
    ButtonControl.Caption = "F11"
    ButtonControl.OnAction = "SendToRebuttal"
    ButtonControl.TooltipText = "Send to Rebuttal"

    'Update TOC
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.OnAction = "UpdateTOC"
    ButtonControl.FaceId = 688
    ButtonControl.TooltipText = "Update the table of contents"
    ButtonControl.BeginGroup = True

    'Document Map
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton, ID:=1714)
    ButtonControl.Caption = " "
    ButtonControl.TooltipText = "Show the document map"

    'Page Break
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton, ID:=509)
    ButtonControl.TooltipText = "Insert Page Break"

    ' Windows
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton, ID:=959)
    ButtonControl.FaceId = 303

    ' Field Codes
    'Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton, ID:=288)

    'Clear Formatting
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton, ID:=8099)
    ButtonControl.Caption = "Clear Formatting"
    ButtonControl.FaceId = 2822
    ButtonControl.BeginGroup = True

    ' Paste and Condense
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "F12 Paste && Condense"
    ButtonControl.Style = msoButtonCaption
    ButtonControl.OnAction = "PasteAndCondense"
    ButtonControl.TooltipText = "tooltip"

    ' Condense
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Caption = "Condense"
    ButtonControl.Style = msoButtonCaption
    ButtonControl.OnAction = "Condense"
    ButtonControl.TooltipText = "tooltip"

    'Comments Group
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton, ID:=1594)
    ButtonControl.BeginGroup = True
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton, ID:=1589)
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton, ID:=1590)
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton, ID:=1591)
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton, ID:=1592)

    'DH Settings
    Set ButtonControl = DHToolbar.Controls.Add(Type:=msoControlButton)
    ButtonControl.Style = msoButtonIcon
    ButtonControl.OnAction = "OpenSettings"
    ButtonControl.FaceId = 8876
    ButtonControl.BeginGroup = True
    ButtonControl.TooltipText = "DebateHelper Settings"

    DHToolbar.Width = 558
    DHToolbar.Height = 62

    DHToolbar.Protection = msoBarNoCustomize + msoBarNoResize

    DHToolbar.Visible = True
    'Save template if editing it
    ThisDocument.AttachedTemplate.Saved = True

End Sub
Sub FKeyAssignAction()

    Dim FKey As CommandBarControl
    Dim FKeyAction As String

    Set FKey = CommandBars.ActionControl
    If FKey Is Nothing Then Exit Sub

    FKeyAction = FKey.Parameter
    Select Case FKeyAction

    Case Is = "Tag"
        Selection.Style = ActiveDocument.Styles(wdStyleHeading8)
    Case Is = "Sub Tag"
        Selection.Style = ActiveDocument.Styles(wdStyleHeading9)
    Case Is = "Citation"
        Selection.Style = ActiveDocument.Styles("Citation")
    Case Is = "Evidence"
        Selection.Style = ActiveDocument.Styles("Normal")
    Case Is = "Section Title 1"
        Selection.Style = ActiveDocument.Styles(wdStyleHeading1)
    Case Is = "Section Title 2"
        Selection.Style = ActiveDocument.Styles(wdStyleHeading2)
    Case Is = "Section Title 3"
        Selection.Style = ActiveDocument.Styles(wdStyleHeading3)
    Case Is = "Block"
        Selection.Style = ActiveDocument.Styles(wdStyleHeading4)
    Case Is = "Response Level 1"
        Selection.Style = ActiveDocument.Styles(wdStyleHeading5)
    Case Is = "Response Level 2"
        Selection.Style = ActiveDocument.Styles(wdStyleHeading6)
    Case Is = "Response Level 3"
        Selection.Style = ActiveDocument.Styles(wdStyleHeading7)
    Case Else
        'Nothing
    End Select

End Sub

