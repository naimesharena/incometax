Attribute VB_Name = "mfMessage"
Public Sub fmsgbox(IntMsg As Variant)
'With MessageBox
'    .Height = 300
'    .LMessagebox.Height = 300
'    .CommandButton1.Top = 300
'    .LMessagebox = IntMsg
'    .LMessagebox.TextAlign = fmTextAlignLeft
'    .Show
'End With
MsgBox IntMsg, , "Alert"
End Sub

Public Sub fmsgboxStatus(IntMsg1 As Variant)
With MessageBox
    .Height = 150
    .LMessagebox.Height = 65
    .CommandButton1.Top = 78
    .LMessagebox = "" + IntMsg1
    .LMessagebox.TextAlign = fmTextAlignLeft
    .Show
End With

End Sub

Public Sub fmsgboxsmall(IntMsg2 As Variant)
With MessageBox
     .Height = 255
     .LMessagebox.Height = 174
     .CommandButton1.Top = 191
     .LMessagebox = IntMsg2
     .LMessagebox.TextAlign = fmTextAlignLeft
     .Show
End With
End Sub

