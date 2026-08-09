Attribute VB_Name = "Sheet18"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Dim ActiveCellRange As String

Private Sub Worksheet_Activate()
On Error Resume Next
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet18.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Sheet18.Shapes("Button 11").Left = Range("Button_10A").Left
Sheet18.Shapes("Button 11").Top = Range("Button_10A").Top

Sheet18.Shapes("Button 12").Left = Range("Button_10AA").Left + 2
Sheet18.Shapes("Button 12").Top = Range("Button_10AA").Top + 2


Sheet18.Protect Password:=sPassword

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
