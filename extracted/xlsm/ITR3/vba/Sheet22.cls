Attribute VB_Name = "Sheet22"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_ChangeX(ByVal Target As Range)
On Error Resume Next

CalculateTaxCheckFlag = False

Application.EnableEvents = False
Application.ScreenUpdating = False
Sheet22.Unprotect Password:=getmsgstate

If Not Application.Intersect(Sheet22.Range("AMT.AdjustedUnderSec115JC3a"), Sheet22.Range(Target.Address)) Is Nothing Then
        If Sheet22.Range("AMT.AdjustedUnderSec115JC").value = 0 Then
            fmsgbox "* Sr No 3a can not exceed 0 since SrNo 3=0"
            Sheet22.Range("AMT.AdjustedUnderSec115JC3a").value = 0
        End If
    End If
Sheet22.Protect Password:=getmsgstate

Sheet22.Activate
Application.EnableEvents = True

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
