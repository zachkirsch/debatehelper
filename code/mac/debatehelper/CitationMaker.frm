VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CitationMaker 
   Caption         =   "Generate Citation"
   ClientHeight    =   6000
   ClientLeft      =   -3800
   ClientTop       =   -2760
   ClientWidth     =   4720
   OleObjectBlob   =   "CitationMaker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CitationMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    On Error GoTo Handler

    Unload Me

    Exit Sub
Handler:

    MsgBox Error(Err.Number, Err.Description), vbExclamation, "DebateHelper Error"
End Sub

Private Sub btnOkay_Click()
    Dim Citation As String
    On Error Resume Next

    If Me.TextBox1.Value <> "" Then Citation = Me.TextBox1.Value

    If Me.TextBox2.Value <> "" Then
        Citation = Citation & " (" & Me.TextBox2.Value & "). "
    Else
        Citation = Citation & ". "
    End If

    If Me.TextBox3.Value <> "" Then Citation = Citation & Me.TextBox3.Value & ". "

    If Me.TextBox4.Value <> "" Then Citation = Citation + "Accessed " & Me.TextBox4.Value & ". "

    If Me.TextBox5.Value <> "" Then Citation = Citation + "Published " & Me.TextBox5.Value & ". "

    If Me.TextBox6.Value <> "" Then Citation = Citation + Me.TextBox6.Value & "."

    AddCitation (Citation)
    Unload Me
End Sub

