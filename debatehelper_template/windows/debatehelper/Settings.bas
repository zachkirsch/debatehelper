Attribute VB_Name = "Settings"
'Declarations for Win API calls
'ShellExecute needed to launch installer package
#If Win64 Then
    Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

'For internet connection test
Private Const FLAG_ICC_FORCE_CONNECTION = &H1
#If Win64 Then
    Private Declare PtrSafe Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
#Else
    Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
#End If

Sub OpenTemplatesFolder()
    On Error GoTo Handler

    Shell "explorer.exe " & CStr(Environ("USERPROFILE")) & "\AppData\Roaming\Microsoft\Templates", vbNormalFocus

    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"
End Sub

Function GetVersion() As String
    On Error GoTo Handler

    GetVersion = ActiveDocument.AttachedTemplate.BuiltInDocumentProperties(wdPropertyKeywords)

    Exit Function
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"
End Function

Sub UpdateCheck()

'Check for updates
'Taken from Verbatim

    Dim HttpReq As MSXML2.ServerXMLHTTP60
    Dim XMLDoc As MSXML2.DOMDocument60
    Dim FileStream As ADODB.Stream
    Dim TempFile As String
    Dim retval
    Dim SavedStatus As Boolean

    'Turn on error checking
    On Error GoTo Handler

    SavedStatus = ActiveDocument.Saved

    'If IsInternetConnected = False Then
    '    Application.StatusBar = "Can't check for updates; Internet connection failed."
    '    GoTo SubExit
    'End If

    Application.StatusBar = "Checking for DebateHelper updates..."

    'Create and send HttpReq
    Set HttpReq = New ServerXMLHTTP60
    HttpReq.Open "GET", "http://hosting.debatehelper.com/updates.xml", False
    HttpReq.SetRequestHeader "Content-Type", "application/xml"
    HttpReq.SetRequestHeader "Accept", "application/xml"
    HttpReq.Send

    'Exit if the request fails
    If HttpReq.Status <> 200 Then
        Application.StatusBar = "Update Check Failed. Automatic update checking has been disabled."
        Set HttpReq = Nothing
        UpdateFailure = True
        SaveSetting "DebateHelper", "Main", "AutoUpdateCheck", False
        GoTo SubExit
    End If

    SaveSetting "DebateHelper", "Main", "LastUpdateCheck", Now

    'Process XML
    Set XMLDoc = HttpReq.responseXML

    DynamicUpdateLabel = True


    'If newer version is found
    If XMLDoc.getElementsByTagName("pcversion").Item(0).Text > Settings.GetVersion Then
        UpdateAvailable = True
        'RibbonControl.RefreshRibbon

        'MsgBox "There is a newer version of DebateHelper available for download." & vbNewLine & "Visit www.DebateHelper.com for the latest version."
        'Set HttpReq = Nothing
        'Set XMLDoc = Nothing
        'GoTo SubExit

        'Confirm update
        If MsgBox("There is a newer version of DebateHelper available for download. Would you like to download the newest version?", vbYesNo) = vbNo Then GoTo SubExit

        Application.StatusBar = "Downloading updates..."

        'Get the URL for latest PC version
        HttpReq.Open "GET", XMLDoc.getElementsByTagName("pcurl").Item(0).Text, False
        HttpReq.Send

        'Save file to disk
        Set FileStream = CreateObject("ADODB.Stream")
        FileStream.Open
        FileStream.Type = 1
        FileStream.Write HttpReq.ResponseBody
        TempFile = CStr(Environ("TEMP")) & "\" & "DebateHelper.msi"
        FileStream.SaveToFile TempFile, 2    '1 = no overwrite, 2 = overwrite
        FileStream.Close
        Set FileStream = Nothing

        'Launch installer
        Application.StatusBar = "Launching installer..."
        MsgBox "Please make sure all word documents are closed before you download the newest DebateHelper."
        retval = ShellExecute(0, "OPEN", TempFile, "", "", 0)

    Else
        Application.StatusBar = "No DebateHelper updates found."
        UpdateAvailable = False
        'RibbonControl.RefreshRibbon

        'Close HttpReq
        Set HttpReq = Nothing
        Set XMLDoc = Nothing

        GoTo SubExit
    End If

SubExit:

    If SavedStatus = True Then ActiveDocument.Save
    Exit Sub

Handler:

    Set HttpReq = Nothing
    Set XMLDoc = Nothing
    Set FileStream = Nothing
    Application.StatusBar = "There may have been a problem checking for DH updates."

    SaveSetting "DebateHelper", "Main", "AutoUpdateCheck", False

    If SavedStatus = True Then ActiveDocument.Save

    If Err.Number = 0 Then Exit Sub

    Application.StatusBar = "Update Check Failed. Automatic update checking for DH has been disabled."
    UpdateFailure = True

End Sub

'Testing for internet connection
Public Function IsInternetConnected() As Boolean
'KPD-Team 2001
'E-Mail: KPDTeam@Allapi.net

    On Error GoTo Handler

    If InternetCheckConnection("http://www.google.com/", FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
        IsInternetConnected = False
    Else
        IsInternetConnected = True
    End If

    Exit Function
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Function

Sub LaunchWebsite(URL As String)

    On Error GoTo Handler
    ActiveDocument.FollowHyperlink (URL)

    Exit Sub

Handler:
    If Err.Number = 4198 Then
        MsgBox "Opening website failed. Check your internet connection."
    Else
        MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"
    End If

End Sub

Function Error(ErrNumber As String, ErrMessage As String)

    Error = "You've received an error in DebateHelper." & vbNewLine & "Error: " & ErrNumber & " " & ErrMessage & "." & _
            vbNewLine & "Go to http://www.debatehelper.com/contact.html to report the bug and fix it."
End Function

Sub OpenCitationMaker()

'Save Current Location
    Dim StartLocation
    On Error GoTo Handler

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
        MsgBox "No cite found. Make sure the previous card's cite is using the 'Citation' style."
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


Sub AddCitation(Citation As String)

    Selection.Collapse
    Selection.MoveDown Unit:=wdLine, Count:=2

    If Selection.Paragraphs.OutlineLevel = wdOutlineLevel8 Or _
       Selection.Paragraphs.OutlineLevel = wdOutlineLevel9 Then
        Selection.MoveEnd Unit:=wdLine, Count:=-1
    Else
        With Selection.Find
            .ClearFormatting
            .Text = ""
            .Wrap = wdFindStop
            .Format = True
            .Style = ActiveDocument.Styles("Tag")
            .Forward = False
            .Execute
        End With
        Selection.MoveRight Unit:=wdCharacter, Count:=-1
    End If

    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Delete
    Selection.InsertBefore (Citation)

    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1

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

