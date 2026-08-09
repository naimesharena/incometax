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
'Change-16.02.2023.103.PAG.02
'MsgBox "Warning message: Field no. 4 Sources of funds(1c + 2c +3) and field no. 6 In a case where regular books of account of business or profession are not maintained -(furnish the following information as on 31st day of March, 2022, in respect of business or profession). in schedule Balance sheet is equal to zero And Income chargeable under the head 'Profits and Gains from Business or Profession at Sl.No D in Schedule BP is greater than 2.50 Lakh."
'AY_2024_25 OLD 'Malli
'MsgBox "Warning message: Field no. 4 Sources of funds(1c + 2c + 3 + 4) and field no. 6 In a case where regular books of account of business or profession are not maintained -(furnish the following information as on 31st day of March, 2023, in respect of business or profession). in schedule Balance sheet is equal to zero And Income chargeable under the head 'Profits and Gains from Business or Profession at Sl.No D in Schedule BP is greater than 2.50 Lakh."
'AY_2024_25 Change 'Malli
'MsgBox "Warning message: Field no. 4 ""Sources of funds(1c + 2c + 3 + 4)"" and field no.6 ""In a case where regular books of account of business or profession are not maintained -(furnish the following information as on 31st day of March, 2024, in respect of business or profession)."" in schedule Balance sheet is equal to zero And ""Income chargeable under the head 'Profits and Gains from Business or Profession"" at Sl.No D in Schedule BP is greater than 2.50 Lakh."

'MsgBox "Warning message: Field no. 4 ""Sources of funds(1c + 2c + 3 + 4)"" and field no.6 ""In a case where regular books of account of business or profession are not maintained -(furnish the following information as on 31st day of March, 2025, in respect of business or profession)."" in schedule Balance sheet is equal to zero And ""Income chargeable under the head 'Profits and Gains from Business or Profession"" at Sl.No D in Schedule BP is greater than 2.50 Lakh."

'MsgBox "Warning message: Field no. 5 ""Sources of funds(1c + 2c + 3 + 4)"" and field no.6 " & _
'           """In a case where regular books of account of business or profession are not maintained -(furnish the following information as on 31st day of March, 2025, in respect of business or profession).""in schedule Balance sheet is equal to zero" & vbCrLf & _
'           "And" & vbCrLf & _
'           """Income chargeable under the head 'Profits and Gains from Business or Profession"" at Sl.No D in Schedule BP is greater than 2.50 Lakh."
'fmsgbox
          
'MsgBox "Warning message: Field no. 5 ""Sources of funds(1c + 2c + 3 + 4)"" and field no.6 " & _
           """In a case where regular books of account of business or profession are not maintained -(furnish the following information as on 31st day of March, 2026, in respect of business or profession).""in schedule Balance sheet is equal to zero" & vbCrLf & _
           "And" & vbCrLf & _
           """Income chargeable under the head 'Profits and Gains from Business or Profession"" at Sl.No D in Schedule BP is greater than 2.50 Lakh."
           
           
           
fmsgbox "Warning message: Field no. 5 ""Sources of funds(1c + 2c + 3 + 4)"" and field no.6 " & _
           """In a case where regular books of account of business or profession are not maintained -(furnish the following information as on 31st day of March, 2026, in respect of business or profession).""in schedule Balance sheet is equal to zero" & vbCrLf & _
           "And" & vbCrLf & _
           """Income chargeable under the head 'Profits and Gains from Business or Profession"" at Sl.No D in Schedule BP is greater than 2.50 Lakh."
                
          
Dim x As Integer
Dim i As Integer
Dim chk_00001 As Variant
Dim Notchk_00001 As Variant

Notchk_00001 = 0
chk_00001 = 0

x = Sheet2.Range("NOB.Code").Cells.count
For i = 1 To x
'If Sheet2.Range("NOB.Code").Cells(i, 1).value <> "(Select)" And Sheet2.Range("NOB.Code").Cells(i, 1).value <> "" Then
If Sheet2.Range("NOB.Code").Cells(i, 1).value = "00001-Share of Income from firm only" Then


chk_00001 = chk_00001 + 1

Else

If Sheet2.Range("NOB.Code").Cells(i, 1).value <> "(Select)" And Sheet2.Range("NOB.Code").Cells(i, 1).value <> "" Then
Notchk_00001 = Notchk_00001 + 1

End If

'Exit Sub
End If
'End If

Next i


'If Not chk_00001 = 1 And Not Notchk_00001 = 0 Then
' MsgBox "Please ensure that the relevant fields of Profit & Loss and Balance Sheet are filled else Return of Income may be treated as defective u/s 139(9)."
'
' End If

If (chk_00001 = 1 And Notchk_00001 = 0) Then
Exit Sub
Else:
'    MsgBox "Please ensure that the relevant fields of Profit & Loss and Balance Sheet are filled else Return of Income may be treated as defective u/s 139(9)."
     fmsgbox "Please ensure that the relevant fields of Profit & Loss and Balance Sheet are filled else Return of Income may be treated as defective u/s 139(9)."
End If

'MsgBox "Please ensure that the relevant fields of Profit & Loss and Balance Sheet are filled else Return of Income may be treated as defective u/s 139(9)."

Sheet3.Protect Password:=sPassword
End Sub



Sub Worksheet_Deactivate()
If (Sheet3.Range("sheet2.TotFundSrc").value = 0) And (Sheet3.Range("sheet3.TotSundryDbtAmt").value = 0) And (Sheet3.Range("sheet3.TotSundryCrdAmt").value = 0) And (Sheet3.Range("sheet3.TotStkInTradAmt").value = 0) And (Sheet3.Range("sheet3.CashBalAmt").value = 0) And (Sheet9.Range("sheet12.IncChrgUnHdProftGain").value > 250000) Then
Dim QLine1, QLine2, Nchar As String
QLine1 = "Please ensure that the relevant fields of Profit & Loss and Balance Sheet are filled else Return of Income may be treated as defective u/s 139(9)"
QLine2 = "Note - If Nature of business code is only 00001-Share of Income from firm only, then above Pop up should not come"
Nchar = "\n"
'Change-23.05.2023.101 -> Issue number 4 in production build 1
'MsgBox QLine1 + Nchar + QLine2
MsgBox QLine1

End If

End Sub

