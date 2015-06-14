Attribute VB_Name = "Startup"
Option Explicit
Public XML As String
Public FSO As Scripting.FileSystemObject

' Following declaration are for checking internet connections
Private Const FLAG_ICC_FORCE_CONNECTION = &H1
#If Win64 Then
    Private Declare PtrSafe Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
    Private Declare PtrSafe Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
#Else
    Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
    Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
#End If
Private Const strSite As String = "http://hosting.debatehelper.com"
Dim strISPName As String * 255


Sub AutoOpen()

    If System.OperatingSystem = "Macintosh" Then
        MsgBox ("This is a Windows-only template. While on Mac, you will have very little functionality.")
        Exit Sub
    End If

    Call AutoOpenWindows

    Selection.Collapse

End Sub
Sub AutoOpenWindows()

    Application.TaskPanes(wdTaskPaneFormatting).Visible = False

    Call BuildKeyboardShortcuts

    ActiveDocument.FormFields.Shaded = False

    Call UpdateTOC

    'Refresh document styles from template if setting checked and not editing template itself
    If GetSetting("DebateHelper", "Main", "AutoUpdateStyles", True) = True Then ActiveDocument.UpdateStyles

    'Check for new messages every three days
    If DateDiff("d", GetSetting("DebateHelper", "Main", "LastMessageCheck", "8/15/2004 11:58:25 AM"), Now) > 3 Then MessageCheck

    'Check for updates if auto-checking enabled in settings (checks in Ribbon startup)
    DynamicUpdateLabel = False
    If GetSetting("DebateHelper", "Main", "AutoUpdateCheck", True) = True Then DynamicUpdateLabel = True

    'Analytics
    UpdateAnalytics

    'Reinstall DebateHelper
    'Call SendModuleToNormal 'Sends DebateHelper_Normal to Normal Template
    'Application.NormalTemplate.Save

    ThisDocument.AttachedTemplate.Saved = True
    ActiveDocument.Save

End Sub

Sub AutoNew()

On Error GoTo Handler
    If System.OperatingSystem = "Macintosh" Then
        MsgBox "This template is not intended for Mac."
        ActiveDocument.Saved = True
        ActiveDocument.Close
        Exit Sub
    End If

    Call SaveDocument

    Call BuildKeyboardShortcuts

    Call UpdateTOC

    ThisDocument.AttachedTemplate.Saved = True

    ActiveDocument.Save

    Selection.Collapse

    Exit Sub

Handler:
    If Err.Number = 4198 Then
        MsgBox "Error with saving!", vbExclamation, "DebateHelper Error"
        Exit Sub
    End If
    
    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub

Sub AttachTemplate()

    On Error GoTo Handler

    If InStr(ActiveDocument.Name, "dotm") Then Exit Sub

    Dim CurrentTemplate As String
    CurrentTemplate = ActiveDocument.AttachedTemplate

    ActiveDocument.AttachedTemplate = Options.DefaultFilePath(wdUserTemplatesPath) & "\DebateHelper.dotm"

    If ActiveDocument.AttachedTemplate <> CurrentTemplate Then
        Application.Documents.Open (ActiveDocument.FullName)
    End If

    ActiveDocument.UpdateStyles


    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub

Sub SaveDocument()

    On Error GoTo Handler

    On Error GoTo Handler

    With Application.Dialogs(wdDialogFileSaveAs)
        .Format = Word.WdSaveFormat.wdFormatXMLDocument
        .Name = "*"
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
        Exit Sub
    End If
    If Err.Number = 0 Then Exit Sub

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub

Sub BuildKeyboardShortcuts()
'Change keyboard shortcuts in template

    On Error GoTo Handler

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
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF10), KeyCategory:=wdKeyCategoryMacro, Command:="CopyCard"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF11), KeyCategory:=wdKeyCategoryMacro, Command:="SendToRebuttal"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF12), KeyCategory:=wdKeyCategoryMacro, Command:="PasteAndCondense"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, _
                                          wdKeyS), KeyCategory:=wdKeyCategoryMacro, _
                                          Command:="DebateTools.ShowStyle"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, _
                                          wdKeyC), KeyCategory:=wdKeyCategoryMacro, _
                                          Command:="OpenCitationMaker"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, _
                                          wdKeyR), KeyCategory:=wdKeyCategoryMacro, _
                                          Command:="DebateTools.SendToRebuttal"

    Application.CustomizationContext = ThisDocument


    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub

Function CheckConnection()
    Dim Ret As Long
    Ret = InternetGetConnectedStateEx(Ret, strISPName, 254, 0)

    If InternetCheckConnection(strSite, FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
        CheckConnection = False
    Else
        CheckConnection = True
    End If
End Function

Sub UpdateAnalytics()
    On Error GoTo Handler
    If CheckConnection = False Then Exit Sub

    Dim HttpReq As MSXML2.ServerXMLHTTP60
    Set HttpReq = New ServerXMLHTTP60
    HttpReq.Open "GET", "http://hosting.debatehelper.com/app%20tracker%20windows.html", False
    HttpReq.Send

Handler:
End Sub


Sub MessageCheck()

'Check for new message from DH Website
'Taken from Verbatim

    Dim HttpReq As MSXML2.ServerXMLHTTP60
    Dim XMLDoc As MSXML2.DOMDocument60
    Dim FileStream As ADODB.Stream
    Dim SavedStatus As Boolean

    'Turn on error checking
    On Error GoTo Handler

    SavedStatus = ActiveDocument.Saved

    'Create and send HttpReq
    Set HttpReq = New ServerXMLHTTP60
    HttpReq.Open "GET", "http://hosting.debatehelper.com/updates.xml", False
    HttpReq.SetRequestHeader "Content-Type", "application/xml"
    HttpReq.SetRequestHeader "Accept", "application/xml"
    HttpReq.Send

    'Exit if the request fails
    If HttpReq.Status <> 200 Then
        Application.StatusBar = "Update Check Failed"
        Set HttpReq = Nothing
    End If

    'Process XML
    Set XMLDoc = HttpReq.responseXML

    If XMLDoc.getElementsByTagName("pcmessage").Item(0).Text = "" Then
        'Do nothing
    Else
        'If there a message online and it's different than the last message, then show the new message
        If GetSetting("DebateHelper", "Main", "LastMessage", "") <> XMLDoc.getElementsByTagName("pcmessage").Item(0).Text Then
            MsgBox XMLDoc.getElementsByTagName("pcmessage").Item(0).Text, vbInformation, "Update from Developer"
            ' Save setting as the last message that was shown
            SaveSetting "DebateHelper", "Main", "LastMessage", _
                        XMLDoc.getElementsByTagName("pcmessage").Item(0).Text
        End If
    End If

    ' Save setting as the last time there was a check for new messages
    SaveSetting "DebateHelper", "Main", "LastMessageCheck", Now

    Set HttpReq = Nothing
    Set XMLDoc = Nothing
    Set FileStream = Nothing
    If SavedStatus = True Then ActiveDocument.Save
    Exit Sub

Handler:

    Set HttpReq = Nothing
    Set XMLDoc = Nothing
    Set FileStream = Nothing

    If SavedStatus = True Then ActiveDocument.Save

End Sub
