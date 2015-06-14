Attribute VB_Name = "Settings"
Function Error(ErrNumber As String, ErrMessage As String)

    Error = "You've received an error in DebateHelper." & vbNewLine & "Error: " & ErrNumber & " " & ErrMessage & "." & _
            vbNewLine & "Go to http://www.debatehelper.com/contact.html to report the bug and fix it."
End Function

Sub OpenSettings()

    Dim x As frmSettings
    Set x = New frmSettings
    x.Show

End Sub

Sub OpenCitationMaker()

    On Error GoTo Handler

    'Save Current Location
    Dim StartLocation
    StartLocation = Selection.Start

    Selection.MoveDown Unit:=wdLine, Count:=2

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
        MsgBox "No cite found. Make sure the card's cite is using the 'Citation' style."
        Exit Sub
    Else
        Selection.Start = StartLocation
        Selection.Collapse
    End If


    Dim x As CitationMaker
    Set x = New CitationMaker
    x.Show

Handler:
    If Err.Number = 0 Then Exit Sub
    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"
End Sub

Sub UpdateCheck(Optional Popup As Boolean)

    Dim newestVersion As String
    Dim URL As String

    On Error GoTo Handler

    SavedStatus = ActiveDocument.Saved
    Application.StatusBar = "Checking for DebateHelper updates..."

    newestVersion = MacScript("(do shell script ""curl http://hosting.debatehelper.com/macversion.txt"")")
    If newestVersion > Settings.GetVersion Then
        If MsgBox("Do you want to download it now?", vbYesNo, "DebateHelper Update Available!") = vbYes Then
            URL = MacScript("(do shell script ""curl http://hosting.debatehelper.com/macurl.txt"")")
            ThisDocument.FollowHyperlink (URL)
            SaveSetting "DebateHelper", "Main", "LastUpdateCheck", Now
        End If
    Else
        If Popup = True Then
            MsgBox "No DebateHelper updates found."
        Else
            Application.StatusBar = "No DebateHelper updates found."
        End If
    End If

    If SavedStatus = True Then ActiveDocument.Save

Handler:
    'Application.StatusBar = "There may have been a problem checking for DH updates."
    SaveSetting "DebateHelper", "Main", "AutoUpdateCheck", False
    If SavedStatus = True Then ActiveDocument.Save
    Application.StatusBar = "Update Check Failed. Automatic update checking for DH has been disabled."


End Sub

Sub MessageCheck()

    Dim message As String
    On Error GoTo Handler
    SavedStatus = ActiveDocument.Saved

    If MacScript("(do shell script ""curl http://hosting.debatehelper.com/macmessage.txt"")") = "" Then
        'Do nothing
    Else
        'If there a message online and it's different than the last message, then show the new message
        If GetSetting("DebateHelper", "Main", "LastMessage", "") <> MacScript("(do shell script ""curl http://hosting.debatehelper.com/macmessage.txt"")") Then
            MsgBox MacScript("(do shell script ""curl http://hosting.debatehelper.com/macmessage.txt"")"), vbInformation, "Update from Developer"
            ' Save setting as the last message that was shown
            SaveSetting "DebateHelper", "Main", "LastMessage", _
                        MacScript("(do shell script ""curl http://hosting.debatehelper.com/macmessage.txt"")")
        End If
    End If

    ' Save setting as the last time there was a check for new messages
    SaveSetting "DebateHelper", "Main", "LastMessageCheck", Now

    If SavedStatus = True Then ActiveDocument.Save

Handler:
    'Application.StatusBar = "There may have been a problem checking for DH updates."
    If SavedStatus = True Then ActiveDocument.Save


End Sub

Function GetVersion() As String
    On Error GoTo Handler

    GetVersion = ActiveDocument.AttachedTemplate.BuiltInDocumentProperties(wdPropertyKeywords)

    Exit Function

Handler:
    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"
End Function

Sub LaunchWebsite(URL As String)

    On Error GoTo Handler
    ActiveDocument.FollowHyperlink (URL)

    Exit Sub

Handler:

End Sub

Sub UpdateAnalytics()

    On Error GoTo Handler

    MacScript ("(do shell script ""curl http://hosting.debatehelper.com/app%20tracker%20mac.html"")")
Handler:
End Sub

Function ChangeTOC(Min As Integer, Max As Integer, ByRef ExtraStyles() As Boolean)


' Go To TOC Bookmark (manually put right before the TOC on page 1)
    On Error GoTo Handler
    Dim b As Bookmark
    Set b = ActiveDocument.Bookmarks("TOC")
    b.Select

    ActiveWindow.View.ShowFieldCodes = True

    ' Delete TOC
    Selection.MoveEnd Unit:=wdLine
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
    Selection.TypeBackspace

    ActiveWindow.View.ShowFieldCodes = False

    ' For each style, include or exclude based on user prefs (via passed-in array)
    Dim N As Integer
    Dim EndOfTOCField As String
    EndOfTOCField = ""
    For N = 1 To 9
        If ExtraStyles(N) Then
            EndOfTOCField = EndOfTOCField & "Heading " & N & "," & N & ","
        End If
    Next N
    EndOfTOCField = Left(EndOfTOCField, Len(EndOfTOCField) - 1)

    ' Build TOC
BuildTOC:

    ActiveDocument.TablesOfContents.Add _
            Range:=Selection.Range, _
            UseFields:=False, _
            UseHeadingStyles:=True, _
            LowerHeadingLevel:=Max, _
            UpperHeadingLevel:=Min, _
            AddedStyles:=EndOfTOCField


    GoToTop

    Exit Function

Handler:

    If Err.Number = 5941 Then    ' no bookmark
        MsgBox "This feature only works with" + _
             " documents created with DebateHelper 1.6 or higher"
        Selection.Start = StartLocation
        Exit Function
    End If

    If Err.Number = 5 Then    ' array is empty, so let's only have Heading 1
        EndOfTOCField = ""
        GoTo BuildTOC
    End If

    Selection.Start = StartLocation
    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"


End Function




