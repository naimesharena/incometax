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

On Error GoTo endline
Sheet49.Unprotect Password:=getmsgstate
Application.EnableEvents = False

CalculateTaxCheckFlag = False
'Change.12.03.2023.102.IDS
If Not Application.Intersect(Sheet49.Range("turnoverintraTrading_12a"), Sheet49.Range(Target.Address)) Is Nothing Then
    If Sheet49.Range("turnoverintraTrading_12a").value < Sheet49.Range("IncomeIntraTrading_12b").value Then
    'changed by Chetan C M for AY 2026-27
        'start--
'        MsgBox ("Income from Intraday Trading can not be more than Turnover from Intraday Trading")
        MsgBox ("Income from Intraday Trading can not be more than Turnover from Intraday Trading in Schedule Trading Account")
        '--end
        Target.value = ""
        Target.Select
    End If
End If
If Not Application.Intersect(Sheet49.Range("IncomeIntraTrading_12b"), Sheet49.Range(Target.Address)) Is Nothing Then
    If Sheet49.Range("turnoverintraTrading_12a").value < Sheet49.Range("IncomeIntraTrading_12b").value Then
    'changed by Chetan C M for AY 2026-27
        'start--
'        MsgBox ("Income from Intraday Trading can not be more than Turnover from Intraday Trading")
        MsgBox ("Income from Intraday Trading can not be more than Turnover from Intraday Trading in Schedule Trading Account")
        '--end
        Target.value = ""
        Target.Select
    End If
End If

'added by Chetan C M for AY 2026-27
'start--
If Not Application.Intersect(Sheet49.Range("turnoverFO_Trading_12c"), Sheet49.Range(Target.Address)) Is Nothing Then
    If Sheet49.Range("turnoverFO_Trading_12c").value < Sheet49.Range("IncomeFO_Trading_12d").value Then
        MsgBox ("Income from Futures & Options Trading  can not be more than Turnover from Futures & Options Trading in Schedule Trading Account")
        Target.value = ""
        Target.Select
    End If
End If
If Not Application.Intersect(Sheet49.Range("IncomeFO_Trading_12d"), Sheet49.Range(Target.Address)) Is Nothing Then
    If Sheet49.Range("turnoverFO_Trading_12c").value < Sheet49.Range("IncomeFO_Trading_12d").value Then
        MsgBox ("Income from Futures & Options Trading  can not be more than Turnover from Futures & Options Trading in Schedule Trading Account")
        Target.value = ""
        Target.Select
    End If
End If
'--end



endline:
Application.EnableEvents = True
Sheet49.Protect Password:=getmsgstate

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




