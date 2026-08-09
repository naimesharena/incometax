Attribute VB_Name = "pwd"
Sub sbUnProtectAll()
  
    On Error GoTo ErrorOccured
   
    Dim pwd1 As String
    pwd1 = InputBox("Please Enter the password")
    If pwd1 = "" Then Exit Sub
    For Each ws In Worksheets
        ws.Unprotect Password:=pwd1
    Next
    MsgBox "All sheets Unprotected."

    Exit Sub
     
ErrorOccured:
    MsgBox "Sheets could not be UnProtected - Password Incorrect"
    Exit Sub
    
End Sub

Sub sbProtectAll()
  
    On Error GoTo ErrorOccured
   
    Dim pwd1 As String
    pwd1 = InputBox("Please Enter the password")
    If pwd1 = "" Then Exit Sub
    For Each ws In Worksheets
        ws.Protect Password:=pwd1
    Next
    MsgBox "All sheets Protected."

    Exit Sub
     
ErrorOccured:
    MsgBox "Sheets could not be Protected - Password Incorrect"
    Exit Sub
    
End Sub

Sub PasswordBreaker()
'Breaks worksheet password protection.

Dim i As Integer, j As Integer, k As Integer
Dim l As Integer, m As Integer, n As Integer
Dim i1 As Integer, i2 As Integer, i3 As Integer
Dim i4 As Integer, i5 As Integer, i6 As Integer

On Error Resume Next

For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
 ThisWorkbook.Unprotect Chr(i) & Chr(j) & Chr(k) & _
    Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
    Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
    MsgBox "Password is " & Chr(i) & Chr(j) & _
    Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
    Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
 'For Each ws In Worksheets
  '  ws.Unprotect Chr(i) & Chr(j) & Chr(k) & _
   ' Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
    'Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
    
    'If ThisWorkbook.ProtectContents = False Then
   ' MsgBox "Password is " & Chr(i) & Chr(j) & _
    'Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
    'Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
   ' Dim password As String
    'password = Chr(i) & Chr(j) & _
    'Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
    'Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
    'ws.Unprotect password:=password
'Next
'Exit Sub

'End If

Next: Next: Next: Next: Next: Next
Next: Next: Next: Next: Next: Next
End Sub


