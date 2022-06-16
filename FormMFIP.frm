Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Dim OFName As OPENFILENAME

Private Function ShowSave(filter) As String
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = Me.hWnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = filter
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = "C:\"
    OFName.lpstrTitle = "Save File - Yoga"
    OFName.flags = 0
    If GetSaveFileName(OFName) Then
        ShowSave = Trim$(OFName.lpstrFile)
    Else
        Exit Function
    End If
End Function

Private Function ShowOpen(filter, text)
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = Me.hWnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = filter
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = "C:\"
    OFName.lpstrTitle = "Created By Yoga"
    OFName.flags = 0
    If GetOpenFileName(OFName) Then
        text.text = Trim$(OFName.lpstrFile)
    Else
        Exit Function
    End If
End Function

Private Sub Command1_Click()
On Error GoTo NotF
If Text1.text = "" Then
MsgBox "Error, Please Insert Archive File", vbCritical, "Notice"
Command2_Click
ElseIf Text2.text = "" Then
MsgBox "Error, Please Insert Picture File", vbCritical, "Notice"
Command3_Click
ElseIf Text3.text = "" Then
MsgBox "Error, Please Insert Output File", vbCritical, "Notice"
Command4_Click
Else
Shell "cmd /c copy /b " & Chr$(34) & Text2.text & Chr$(34) & " + " & Chr$(34) & Text1.text & Chr$(34) & " " & Chr$(34) & Text3.text & Chr$(34), vbHide
MsgBox "Done !!!", vbInformation, "Done"
End If
Exit Sub
NotF:
MsgBox "Error, Please Try Again Later !!!", vbCritical, "ErrHandler"
End Sub
Private Sub Command2_Click()
ShowOpen "Archive Files (*.rar)" + Chr$(0) + "*.rar", Text1
End Sub
Private Sub Command3_Click()
ShowOpen "Image Files (*.jpg)" + Chr$(0) + "*.jpg", Text2
End Sub

Private Sub Command4_Click()
Text3.text = ShowSave("Image (*.jpg)" + Chr$(0) + "*.jpg")
Text3.text = Text3.text & ".jpg"
End Sub

Private Sub Command5_Click()
If MsgBox("Really Want Quit ?", vbOKCancel, "Notice") = vbOK Then
End
Else
Cancel = True
End If
End Sub

Private Sub Form_Load()
Dim R
Dim register
R = GetSetting("Cicorp", "Reg", "yoga15s")
register = Val(R)
If register >= 1 Then
Label99.Caption = "Registered To: " & GetSetting("Cicorp", "Project51", "Name") & ""
Else
MsgBox "You Are Not Registered," + vbCrLf + "This is Only For Pro Version !", vbInformation, "Sorry"
End
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Really Want Quit ?", vbOKCancel, "Notice") = vbOK Then
End
Else
Cancel = True
End If
End Sub
