Attribute VB_Name = "Sheet16"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Activate()

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
'added by chandru
If Sheet16.Range("sheet16.CYLA_Edit").value = "Yes" Or Sheet16.Range("sheet16.CYLA_Edit").value = "No" Then
Sheet16.Unprotect Password:=getmsgstate

Call setOffHPLossCYLA
Call setOffOthSrcLossCYLA

Call CYLACalculations.calcBFLA
'Call setOffBussLossCYLA
Sheet16.Protect Password:=getmsgstate
End If
'
'If Sheet16.Range("sheet16.BFLA_Edit").value = "No" Then
'
'Sheet16.Unprotect Password:=getmsgstate
'Call CYLACalculations.calcBFLA
'
'Sheet16.Protect Password:=getmsgstate
'End If


'Chandru
If Sheet16.Range("sheet16.CYLA_Edit").value = "Yes" Then
Sheet16.Unprotect Password:=getmsgstate
Sheet16.Range("sheet16.BFLA_Edit").Validation.Delete
Sheet16.Range("sheet16.BFLA_Edit").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=sheet16.CYLA_Edit"
Application.EnableEvents = False
Sheet16.Range("sheet16.BFLA_Edit").value = "Yes"
Sheet16.Protect Password:=getmsgstate
ElseIf Sheet16.Range("sheet16.CYLA_Edit").value = "No" Then
Sheet16.Unprotect Password:=getmsgstate
Sheet16.Range("sheet16.BFLA_Edit").Validation.Delete
Sheet16.Range("sheet16.BFLA_Edit").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=CYLA_TempEdit"
Sheet16.Protect Password:=getmsgstate
End If
CalculateTaxCheckFlag = False
Application.EnableEvents = True

End Sub


Private Sub Worksheet_Deactivate()
On Error Resume Next
Dim activesheet1 As Variant
    activesheet1 = ActiveSheet.name
'        CYLACalculations.CYLABFLASetOff

'Updated by Bindu
If (UCase(Sheet16.Range("sheet16.CYLA_Edit").value) = UCase("No") And UCase(Sheet16.Range("sheet16.BFLA_Edit").value) = UCase("No")) Then
    CYLACalculations.CYLABFLASetOff
ElseIf (UCase(Sheet16.Range("sheet16.CYLA_Edit").value) = UCase("Yes") Or UCase(Sheet16.Range("sheet16.BFLA_Edit").value) = UCase("Yes")) Then
    CYLA_BFLA_Editable
End If
'--

    Worksheets(activesheet1).Activate
End Sub
'
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
