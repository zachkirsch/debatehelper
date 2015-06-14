VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "DebateHelper Settings"
   ClientHeight    =   6080
   ClientLeft      =   -5560
   ClientTop       =   -12440
   ClientWidth     =   4880
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public XML As String
Public SuppressMessages As Boolean
Option Explicit
Private Sub UpdateTOC_Click()

    UpdateTOCFromSettings

End Sub

Private Sub UserForm_Initialize()

    Dim f
    Dim MacroArray

    'Turn on Error handling
    On Error GoTo Handler


    'Main tab
    SuppressMessages = True
    Me.chkAutoUpdateCheck.Value = GetSetting("DebateHelper", "Main", "AutoUpdateCheck", True)
    SuppressMessages = False
    Me.chkBlockedCitation.Value = GetSetting("DebateHelper", "Main", "UseBlockedCite", False)
    'Me.lblLastUpdateCheck.Caption = "Last Update Check:  " & _
     Format(GetSetting("DebateHelper", "Main", "LastUpdateCheck", ""), "mm-dd-yy hh:mm")

    Me.chkAutoUpdateStyles.Value = GetSetting("DebateHelper", "Main", "AutoUpdateStyles", True)

    ' TOC Tab
    Me.CheckBox_Heading1.Value = GetSetting("DebateHelper", "Main", "Heading1inTOC", True)
    Me.CheckBox_Heading2.Value = GetSetting("DebateHelper", "Main", "Heading2inTOC", True)
    Me.CheckBox_Heading3.Value = GetSetting("DebateHelper", "Main", "Heading3inTOC", True)
    Me.CheckBox_Heading4.Value = GetSetting("DebateHelper", "Main", "Heading4inTOC", True)
    Me.CheckBox_Heading5.Value = GetSetting("DebateHelper", "Main", "Heading5inTOC", False)
    Me.CheckBox_Heading6.Value = GetSetting("DebateHelper", "Main", "Heading6inTOC", False)
    Me.CheckBox_Heading7.Value = GetSetting("DebateHelper", "Main", "Heading7inTOC", False)
    Me.CheckBox_Heading8.Value = GetSetting("DebateHelper", "Main", "Heading8inTOC", False)
    Me.CheckBox_Heading9.Value = GetSetting("DebateHelper", "Main", "Heading9inTOC", False)

    'About Tab
    Me.lblAbout2.Caption = "DebateHelper v. " & Settings.GetVersion

    Exit Sub

Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub

Private Sub btnSave_Click()
'Save Settings to Registry

    Dim DebateTemplate As Document
    Dim CloseDebateTemplate As Boolean

    'Turn on Error handling
    On Error GoTo Handler

    'Main Tab

    SaveSetting "DebateHelper", "Main", "AutoUpdateCheck", Me.chkAutoUpdateCheck.Value
    SaveSetting "DebateHelper", "Main", "AutoUpdateStyles", Me.chkAutoUpdateStyles.Value
    SaveSetting "DebateHelper", "Main", "UseBlockedCite", Me.chkBlockedCitation.Value

    'Close template if opened separately
    If CloseDebateTemplate = True Then
        DebateTemplate.Close SaveChanges:=wdSaveChanges
    End If

    ActiveDocument.UpdateStyles

    ' TOC Tab

    ' Check if settings have been changed
    Dim TocSettingsChanged
    TocSettingsChanged = False
    If Me.CheckBox_Heading1.Value <> GetSetting("DebateHelper", "Main", "Heading1inTOC", True) Or _
       Me.CheckBox_Heading2.Value <> GetSetting("DebateHelper", "Main", "Heading2inTOC", True) Or _
       Me.CheckBox_Heading3.Value <> GetSetting("DebateHelper", "Main", "Heading3inTOC", True) Or _
       Me.CheckBox_Heading4.Value <> GetSetting("DebateHelper", "Main", "Heading4inTOC", True) Or _
       Me.CheckBox_Heading5.Value <> GetSetting("DebateHelper", "Main", "Heading5inTOC", False) Or _
       Me.CheckBox_Heading6.Value <> GetSetting("DebateHelper", "Main", "Heading6inTOC", False) Or _
       Me.CheckBox_Heading7.Value <> GetSetting("DebateHelper", "Main", "Heading7inTOC", False) Or _
       Me.CheckBox_Heading8.Value <> GetSetting("DebateHelper", "Main", "Heading8inTOC", False) Or _
       Me.CheckBox_Heading9.Value <> GetSetting("DebateHelper", "Main", "Heading9inTOC", False) Then
        TocSettingsChanged = True
    End If

    If TocSettingsChanged Then
        'Save new settings and change toc
        UpdateTOCFromSettings
    End If

    'About Tab
    SaveSetting "DebateHelper", "Main", "Version", Settings.GetVersion

    'Unload the form
    Unload Me

    ActiveDocument.Save

    Exit Sub

Handler:
    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub

Private Sub btnCancel_Click()
    On Error GoTo Handler

    Unload Me

    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"
End Sub

'*************************************************************************************
'* MAIN TAB                                                                                                            *
'*************************************************************************************

Private Sub checkUpdates_Click()
    On Error GoTo Handler

    Settings.UpdateCheck

    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub

Private Sub chkAutoUpdateCheck_Click()

    On Error GoTo Handler

    If SuppressMessages = False And chkAutoUpdateCheck.Value = True Then
        If MsgBox("This will cause a short delay upon opening DebateHelper documents in the future. Proceed?", _
                  vbYesNo + vbQuestion) = vbNo Then chkAutoUpdateCheck.Value = False
    End If

    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub

Private Sub btnRemoveDH_Click()

'Removes Custom UI
    On Error GoTo Handler

    Application.CustomizationContext = Application.NormalTemplate

    Dim Toolbar As CommandBar
    Dim StandardToolbar As CommandBar

    Set StandardToolbar = Application.CommandBars("Standard")
    StandardToolbar.Visible = True
    StandardToolbar.Reset

    Exit Sub

Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"

End Sub

'*************************************************************************************
'* TOC TAB                                                                         *
'*************************************************************************************

Sub UpdateTOCFromSettings()

    SaveSetting "DebateHelper", "Main", "Heading1inTOC", Me.CheckBox_Heading1.Value
    SaveSetting "DebateHelper", "Main", "Heading2inTOC", Me.CheckBox_Heading2.Value
    SaveSetting "DebateHelper", "Main", "Heading3inTOC", Me.CheckBox_Heading3.Value
    SaveSetting "DebateHelper", "Main", "Heading4inTOC", Me.CheckBox_Heading4.Value
    SaveSetting "DebateHelper", "Main", "Heading5inTOC", Me.CheckBox_Heading5.Value
    SaveSetting "DebateHelper", "Main", "Heading6inTOC", Me.CheckBox_Heading6.Value
    SaveSetting "DebateHelper", "Main", "Heading7inTOC", Me.CheckBox_Heading7.Value
    SaveSetting "DebateHelper", "Main", "Heading8inTOC", Me.CheckBox_Heading8.Value
    SaveSetting "DebateHelper", "Main", "Heading9inTOC", Me.CheckBox_Heading9.Value

    Dim ExtraStyles(1 To 9) As Boolean
    ExtraStyles(1) = False
    ExtraStyles(2) = Me.CheckBox_Heading2.Value
    ExtraStyles(3) = Me.CheckBox_Heading3.Value
    ExtraStyles(4) = Me.CheckBox_Heading4.Value
    ExtraStyles(5) = Me.CheckBox_Heading5.Value
    ExtraStyles(6) = Me.CheckBox_Heading6.Value
    ExtraStyles(7) = Me.CheckBox_Heading7.Value
    ExtraStyles(8) = Me.CheckBox_Heading8.Value
    ExtraStyles(9) = Me.CheckBox_Heading9.Value

    Dim x
    x = ChangeTOC(1, 1, ExtraStyles)

End Sub


'*************************************************************************************
'* ABOUT TAB                                                                         *
'*************************************************************************************

Private Sub lblAbout5_Click()
    On Error GoTo Handler

    Settings.LaunchWebsite ("http://www.debatehelper.com/")

    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"
End Sub
