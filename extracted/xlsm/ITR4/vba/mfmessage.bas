Attribute VB_Name = "mfmessage"
Public Sub fmsgbox(IntMsg As Variant)
With MessageBox
    .Height = 389
    .LMessagebox.Height = 312
    .CommandButton1.Top = 324
    .LMessagebox = IntMsg
    .LMessagebox.TextAlign = fmTextAlignLeft
    .Show
End With
End Sub
Public Sub fmsgboxStatus(IntMsg1 As Variant)
With MessageBox
    .Height = 265
    .LMessagebox.Height = 200
    .CommandButton1.Top = 210
    .LMessagebox = "*" + IntMsg1
    .LMessagebox.TextAlign = fmTextAlignLeft
    .Show
End With

End Sub

Public Sub fmsgboxsmall(IntMsg2 As Variant)
With MessageBox
     .Height = 265
     .LMessagebox.Height = 184
     .CommandButton1.Top = 200
     .LMessagebox = IntMsg2
     .LMessagebox.TextAlign = fmTextAlignLeft
     .Show
End With
End Sub
'Malli-----AY_2025-26
Public Sub fmsgboxsmall_LTCG(IntMsg2 As Variant)
     MsgBox IntMsg2, , "Error(S)"
End Sub
