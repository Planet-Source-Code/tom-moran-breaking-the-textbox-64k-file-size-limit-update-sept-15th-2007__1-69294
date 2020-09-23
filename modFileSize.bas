Attribute VB_Name = "modFileSize"
Option Explicit

 'API's and Const to load large files
 
   Private Declare Function SendMessage Lib "user32" _
      Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
      ByVal wParam As Long, lParam As Any) As Long
      
   Private Declare Function GetWindowTextLength Lib "user32" _
      Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

   Private Const WM_SETTEXT = &HC
   Private Const WM_GETTEXT = &HD
   Private Const WM_GETTEXTLENGTH = &HE
Sub SaveFileAs(Filename)
    On Error Resume Next
    Dim strContents As String
    Dim FileNum As Integer
    
    Screen.MousePointer = 11
    
    FileNum = FreeFile
    
    ' Open the file.
    Open Filename For Output As #FileNum
    ' Place the contents of the notepad into a variable.
    strContents = frmHugeText.Text1.Text
    ' Display the hourglass mouse pointer.
    Screen.MousePointer = 11
    ' Write the variable contents to a saved file.
    Print #FileNum, strContents
    Close #FileNum
    ' Reset the mouse pointer.
    Screen.MousePointer = 0

End Sub


Function GetFileName(Filename As Variant)
    ' Display a Save As dialog box and return a filename.
    ' If the user chooses Cancel, return an empty string.
   ' On Error Resume Next
    'frmHugeText.CMDialog1.CancelError = TruE
    frmHugeText.CMDialog1.CancelError = True
    frmHugeText.CMDialog1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    frmHugeText.CMDialog1.Filename = Filename
    frmHugeText.CMDialog1.ShowSave
    If Err <> 32755 Then    ' User chose Cancel.
        GetFileName = frmHugeText.CMDialog1.Filename
    Else
        GetFileName = ""
    End If
End Function
Sub FileOpenProc()
    
    On Error Resume Next
    frmHugeText.CMDialog1.CancelError = True
    frmHugeText.CMDialog1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    frmHugeText.CMDialog1.FilterIndex = 1
    frmHugeText.CMDialog1.Filename = ""
    frmHugeText.CMDialog1.ShowOpen
    If Err = 32755 Then   ' User chose Cancel.
       Exit Sub
    Else
       OpenFile (frmHugeText.CMDialog1.Filename)
    End If
    
End Sub
Sub OpenFile(Filename)

    Dim TempText As String
    Dim FileNum As Integer
    Dim iret As Long
    
    On Error GoTo Errhandler
    
    FileNum = FreeFile
    ' Open the selected file.
    Open Filename For Input As #FileNum
    If Err Then
        MsgBox "Can't open file: " + Filename
        Exit Sub
    End If
    ' Change the mouse pointer to an hourglass.
    Screen.MousePointer = 11
    
'==============================================================
' The following code is where loading is executed.  If you attempt
' to load using the normal VB method you'll get the error which is
' the file is too big.  If you've chosen the API option method then
' the file is loaded with no error.
'================================================================

  If frmHugeText.optLoadFile(0).Value = True Then 'load with normal VB method
  
    
      If FileLen(Filename) > 65000 Then
        Screen.MousePointer = 0
        MsgBox "The file is too large to open."
        Close #FileNum
        Exit Sub
      Else
        TempText = StrConv(InputB(LOF(FileNum), FileNum), vbUnicode)
        frmHugeText.Text1.Text = TempText 'StrConv(InputB(LOF(FileNum), FileNum), vbUnicode)
      End If
    
  Else 'load using API method
  
    TempText = StrConv(InputB(LOF(FileNum), FileNum), vbUnicode)
    DoEvents
    frmHugeText.Text1.Text = ""
    
    iret = SendMessage(frmHugeText.Text1.hwnd, WM_SETTEXT, 0&, ByVal TempText)
    Debug.Print "WM_SETTEXT: " & iret
   
    iret = SendMessage(frmHugeText.Text1.hwnd, WM_GETTEXTLENGTH, 0&, ByVal 0&)
    Debug.Print "WM_GETTEXTLENGTH: " & iret
    
    TempText = ""
  
  End If
  
    Close #FileNum
    ' Reset the mouse pointer.
    Screen.MousePointer = 0
    Exit Sub
    
Errhandler:

     Close #FileNum
    ' Reset the mouse pointer.
     Screen.MousePointer = 0
     MsgBox "Error:" & Str(Err) & " - " & Error$, vbCritical, "Textbox Error"
     Exit Sub
 
 
 
End Sub
