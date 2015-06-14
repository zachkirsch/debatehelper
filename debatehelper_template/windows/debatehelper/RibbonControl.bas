Attribute VB_Name = "RibbonControl"
'API Declarations for saving a pointer to the Ribbon
#If Win64 Then
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef Source As Any, ByVal length As Long)
#Else
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef Source As Any, ByVal length As Long)
#End If

Public DebateRibbon As IRibbonUI
Public DynamicUpdateLabel As Boolean
Public UpdateAvailable As Boolean
Public UpdateFailure As Boolean

Sub OnLoad(Ribbon As IRibbonUI)
    Dim SavedState As Boolean
    On Error GoTo Handler

    Set DebateRibbon = Ribbon

    'Save a pointer to the Ribbon in case it gets lost
    SavedState = ActiveDocument.Saved
    ActiveDocument.Variables("RibbonPointer") = ObjPtr(Ribbon)
    ActiveDocument.Saved = SavedState

    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"
End Sub

#If VBA7 Then
Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
#Else
Function GetRibbon(ByVal lRibbonPointer As Long) As Object
#End If

Dim objRibbon As Object
CopyMemory objRibbon, lRibbonPointer, LenB(lRibbonPointer)
Set GetRibbon = objRibbon
Set objRibbon = Nothing
End Function

Public Sub RefreshRibbon()

    On Error GoTo Handler

    If DebateRibbon Is Nothing Then
        Set DebateRibbon = GetRibbon(ActiveDocument.Variables("RibbonPointer"))
        DebateRibbon.Invalidate
    Else
        DebateRibbon.Invalidate
    End If

    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub

Sub getUpdateLabel(control As IRibbonControl, ByRef returnVal)

'Return the label for the selected control to the Ribbon

    On Error GoTo Handler

    If DynamicUpdateLabel = False Then
        returnVal = "Check for DH updates"
        Exit Sub
    Else
        Call Settings.UpdateCheck
    End If

    If UpdateAvailable = True Then returnVal = "DH Update Available!"
    If UpdateAvailable = False Then returnVal = "No DH Update Available"
    If UpdateFailure = True Then returnVal = "Error checking updates"

    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub

Sub RibbonMain(ByVal control As IRibbonControl)

    On Error GoTo Handler

    CustomizationContext = ActiveDocument.AttachedTemplate

    Select Case control.ID

    Case Is = "DHSettings1"
        Dim SettingsForm1 As frmSettings
        Set SettingsForm1 = New frmSettings
        SettingsForm1.Show
    Case Is = "DHSettings2"
        Dim SettingsForm2 As frmSettings
        Set SettingsForm2 = New frmSettings
        SettingsForm2.Show
    Case Is = "btnSectionLevel1"
        Selection.Style = ActiveDocument.Styles(wdStyleHeading1)
    Case Is = "btnSectionLevel2"
        Selection.Style = ActiveDocument.Styles(wdStyleHeading2)
    Case Is = "btnSectionLevel3"
        Selection.Style = ActiveDocument.Styles(wdStyleHeading3)
    Case Is = "btnBlockStyle"
        Selection.Style = ActiveDocument.Styles(wdStyleHeading4)
    Case Is = "btnResponseLevel1"
        Selection.Style = ActiveDocument.Styles(wdStyleHeading5)
    Case Is = "btnResponseLevel2"
        Selection.Style = ActiveDocument.Styles(wdStyleHeading6)
    Case Is = "btnResponseLevel3"
        Selection.Style = ActiveDocument.Styles(wdStyleHeading7)
    Case Is = "btnTag"
        Selection.Style = ActiveDocument.Styles(wdStyleHeading8)
    Case Is = "btnSubTag"
        Selection.Style = ActiveDocument.Styles(wdStyleHeading9)
    Case Is = "btnCitation"
        Selection.Style = ActiveDocument.Styles("Citation")
    Case Is = "btnEvidence"
        Selection.Style = ActiveDocument.Styles("Normal")
    Case Is = "btnShowStyle"
        ShowStyle
    Case Is = "btnInsertBlock"
        InsertBlock
    Case Is = "btnInsertCard"
        InsertCard
    Case Is = "btnInsertCardWithPreviousCitation"
        InsertCardWithPreviousCite
    Case Is = "btnCopyCard"
        CopyCard
    Case Is = "btnPasteAndCondense"
        PasteAndCondense
    Case Is = "btnCondense"
        Condense
    Case Is = "btnClearFormatting"
        Selection.ClearFormatting
    Case Is = "UpdateStyles"
        ActiveDocument.UpdateStyles
    Case Is = "btnSendToRebuttal"
        SendToRebuttal
    Case Is = "btnCheckUpdate"
        DynamicUpdateLabel = True
        RefreshRibbon
    Case Is = "btnCitationWizard"
        OpenCitationMaker
    Case Else
        'Do Nothing

    End Select

    CustomizationContext = ActiveDocument

    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub


