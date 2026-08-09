Attribute VB_Name = "Sheet49"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Private Sub Worksheet_Activate()
On Error Resume Next
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet49.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Sheet49.Shapes("Button 1").Left = Range("TradingAcc_Button1").Left
Sheet49.Shapes("Button 1").Top = Range("TradingAcc_Button1").Top + 3

Sheet49.Shapes("Button 3").Left = Range("TradingAcc_Button2").Left
Sheet49.Shapes("Button 3").Top = Range("TradingAcc_Button2").Top + 3



Sheet49.Protect Password:=sPassword
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)


CalculateTaxCheckFlag = False



End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error Resume Next
If Not Range(ActiveCell.Address).Locked = True Then
       If ActiveCell.Validation.Type = 6 Then
    Application.OnKey "%{DOWN}", ""
    'Sheet1.Range(targetaddress).Value = ""
ElseIf ActiveCell.Validation.Type = 1 Then
 Application.OnKey "%{DOWN}", ""
ElseIf ActiveCell.Validation.Type = 3 Then

    Application.OnKey "%{DOWN}"
End If
Else
Application.OnKey "%{DOWN}", ""
End If

End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    Cancel = True
End Sub




