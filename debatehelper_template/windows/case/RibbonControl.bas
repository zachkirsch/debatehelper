Attribute VB_Name = "RibbonControl"
Sub RibbonMain(ByVal control As IRibbonControl)

    CustomizationContext = ActiveDocument.AttachedTemplate

    Select Case control.ID

    Case Is = "DHCaseSettings1"

    Case Is = "DHCaseSettings2"

    Case Is = "btnInsertFootnote"
        DebateHelper.InsertFootnote
    Case Is = "btnShowStyle"
        Application.StatusBar = "Style: " & Selection.Paragraphs.Style
    Case Is = "btnPasteCard"
        DebateHelper.PasteAndCondense
    Case Is = "btnWordCount"
        DebateHelper.CountFont12
    Case Is = "btnGrowFont"
        DebateHelper.Size12Underlined
    Case Is = "btnShrinkFont"
        DebateHelper.Size7NotUnderlined
    Case Is = "btnCase"
        Selection.Style = ActiveDocument.Styles("Normal")
    Case Is = "btnCard"
        Selection.Style = ActiveDocument.Styles("Card")
    Case Is = "UpdateStyles"
        ActiveDocument.UpdateStyles
    Case Else
        'Do Nothing

    End Select

    CustomizationContext = ActiveDocument

End Sub



