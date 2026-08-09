Attribute VB_Name = "Sheet3"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit



Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error Resume Next
Application.EnableEvents = False

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

If Sheet3.Range("sheet2.NetBlock").value < 0 Then
    Sheet3.Range("sheet2.NetBlock").value = 0
End If

Application.EnableEvents = True
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)

CalculateTaxCheckFlag = False


End Sub


Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    Cancel = True
End Sub
Sub Worksheet_Activate()
sPassword = EfilingCommon.getmsgstate
Sheet3.Unprotect Password:=sPassword
MsgBox "Warning message: Field no. 4 Sources of funds(1c + 2c +3) and field no. 6 In a case where regular books of account of business or profession are not maintained -(furnish the following information as on 31st day of March, 2023, in respect of business or profession). in schedule Balance sheet is equal to zero And Income chargeable under the head 'Profits and Gains from Business or Profession at Sl.No D in Schedule BP is greater than 2.50 Lakh." 'dpk1601
Sheet3.Protect Password:=sPassword
End Sub



Sub Worksheet_Deactivate()
If (Sheet3.Range("sheet2.TotFundSrc").value = 0) And (Sheet3.Range("sheet3.TotSundryDbtAmt").value = 0) And (Sheet3.Range("sheet3.TotSundryCrdAmt").value = 0) And (Sheet3.Range("sheet3.TotStkInTradAmt").value = 0) And (Sheet3.Range("sheet3.CashBalAmt").value = 0) And (Sheet9.Range("sheet12.IncChrgUnHdProftGain").value > 250000) Then
Dim QLine1, QLine2, Nchar As String
QLine1 = "Please ensure that the relevant fields of Profit & Loss and Balance Sheet are filled else Return of Income may be treated as defective u/s 139(9)"
QLine2 = "Note - If Nature of business code is only 00001-Share of Income from firm only, then above Pop up should not come"
Nchar = "\n"
MsgBox QLine1 + Nchar + QLine2
End If
End Sub

