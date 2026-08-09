Attribute VB_Name = "SchBP"
Option Explicit
Dim end_name1, end_amt1, end_BP, end_BPOE As Long
Dim MsgBx_BP1 As String
Dim msgBx_name, msgBx_BP2, BP_Drpdn, OEName, OEAmt As Variant
Public rngname_35AD_BP, rngname_BPOE As Variant
Sub Cmd_AddRows_BP_Click()
On Error Resume Next
Dim vRows As Long
Sheets("BP").Activate
setTableinfo_EId
ActiveCellRange = EfilingCommon.searchLastRow("EI.Drpdn")
vRows = InsertRowsAndFillFormulas()
Call ExendRangeNameToTable(vRows, rngname_35AD_BP)
Dim RngHOI, rrCount, rr
Set RngHOI = Selection.Resize(rowsize:=vRows + 1)
        rrCount = 0
        For Each rr In RngHOI.Rows
            If Not rrCount = 0 Then
                Range("G" & rr.row).value = "(Select)"
            End If
            rrCount = rrCount + 1
        Next
End Sub
Sub Cmd_AddRows_BP_Clickold1()
Dim vRows As Long
Sheets("BP").Activate
setTableinfo_EId
ActiveCellRange = EfilingCommon.searchLastRow("EI.Drpdn")
vRows = InsertRowsAndFillFormulas()
Call ExendRangeNameToTable(vRows, rngname_35AD_BP)
End Sub

Sub setTableinfo_EId()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet9.Range("EI.Drpdn").count
    Set rangecells = Sheet9.Range("EI.Drpdn").Cells
     For mIntCtr = 1 To mIntCells
            If Not ((rangecells.item(mIntCtr).value = "")) Then
                ccount = ccount + 1
            End If
    Next
 end_BP = ccount
 rngname_35AD_BP = "EI.Drpdn;"
End Sub

Sub ValidateSheetBP()
ValidateSchBP
fmsgboxStatus "Sheet BP is ok"
End Sub



Sub ValidateSchBP()
If Not ValidateExemptInc Then
    Sheet9.Activate
    fmsgboxsmall (MsgBx_BP1)
    CloseMsg
End If

If Not CheckNames_BP Then
  Sheet9.Activate
  fmsgboxsmall (msgBx_name)
  CloseMsg
End If
'Change-28.01.2023.101.BP.XA
If Sheet9.Range("sheet10.IncRecCredPLOthHeads4Div").value > Sheet4.Range("PL.Dividends").value Then
         msgBx_BP2 = msgBx_BP2 + "* Divident Income can not be higher then Profit and loss divident income" & Chr(13)
            Sheet9.Range("sheet10.IncRecCredPLOthHeads4Div").value = "0"
    End If

If Sheet9.Range("BP_Chargeable_rule7").value > Sheet9.Range("BP_4b_ProfitRule7").value Then
    msgBx_BP2 = msgBx_BP2 + "* Chargeable income under Rule 7 at Point 37 cannot be more than income shown at Point 4b in Schedule BP. " & Chr(13)
End If

If Sheet9.Range("BP_Chargeable_rule7A").value < (0.35 * Sheet9.Range("BP_4b_ProfitRule7A").value) Then
    msgBx_BP2 = msgBx_BP2 + "* Chargeable income under Rule 7A at Point 37 cannot be less than 35% of income shown at Point 4b in Schedule BP. " & Chr(13)
End If

If Sheet9.Range("BP_Chargeable_rule7B1").value < (0.25 * Sheet9.Range("BP_4b_ProfitRule7B1").value) Then
    msgBx_BP2 = msgBx_BP2 + "* Chargeable income under Rule 7B(1) at Point 37 cannot be less than 25% of income shown at Point 4b in Schedule BP. " & Chr(13)
End If

If Sheet9.Range("BP_Chargeable_rule7B1A").value < (0.4 * Sheet9.Range("BP_4b_ProfitRule7B2").value) Then
    msgBx_BP2 = msgBx_BP2 + "* Chargeable income under Rule 7B(1A) at Point 37 cannot be less than 40% of income shown at Point 4b in Schedule BP. " & Chr(13)
End If

If Sheet9.Range("BP_Chargeable_rule8").value < (0.4 * Sheet9.Range("BP_4b_ProfitRule8").value) Then
    msgBx_BP2 = msgBx_BP2 + "* Chargeable income under Rule 8 at Point 37 cannot be less than 40% of income shown at Point 4b in Schedule BP. " & Chr(13)
End If

If Sheet9.Range("sheet11.TotExempIncPL").value > (Sheet24.Range("EI.Total").value) Then
    msgBx_BP2 = msgBx_BP2 + "* In schedule BP Amount of exempt income reduced from Profits and Gains of Business and Profession does not tally with income offered in Schedule EI. " & Chr(13)
End If

'chandru
If Sheet9.Range("sheet10.NetPLFromSpecBus").value = "" Then

'Sheet9.Range("sheet10.NetPLFromSpecBus").formula = "=PL_65iv"
Sheet9.Range("sheet10.NetPLFromSpecBus").formula = "=IncomeIntraTrading_12b+PL_65iv"
End If

'Added by Aavula For AY 2026-27 SIT-116699
'SIT-119717 18/05/2026

    If Sheet1.Range("Sheet1.BusiProfAssyr").value = "No" And (Sheet9.Range("sheet12.NetPLBusOthThanSpec7A7B7C").value <> 0 Or Sheet9.Range("sheet12.AdjustedPLFrmSpecuBus").value <> 0 Or Sheet9.Range("sheet12.AdjustedPLFrmSpecifiedBus").value <> 0) Then
        msgBx_BP2 = msgBx_BP2 + "*""Please select option as ""Yes"" at Sl.no.A19(b) of Part A General as there is Business income / loss""" & Chr(13)
   End If


If Not msgBx_BP2 = "" Then
Sheet9.Activate
  fmsgboxsmall (msgBx_BP2)
  CloseMsg
End If

If Not ValidateDrpdn Then
  Sheet9.Activate
  fmsgboxsmall (msgBx_BP2)
  CloseMsg
End If

'ValidateBP_EI 'Malli dt:28/03/24
' If Not ValidateBP_EI Then
'  Sheet9.Activate
'   CloseMsg
'  End If
End Sub

Function CheckNames_BP() As Boolean
    CheckNames_BP = True
Dim rangecell As Range
Dim tar As Variant
Set rangecell = Sheet9.Range("Name_1").Cells

For Each tar In rangecell.value
    If Not checkfieldspecialcharacter(tar) Then
        CheckNames_BP = False
        msgBx_name = msgBx_name + "* Nature in A5 cannot contain special charachters <> ' '&' " '
        Exit Function
    End If
Next

End Function

Function ValidateExemptInc()
ValidateExemptInc = True

Dim i As Long
setTableinfo_name1
setTableinfo_amt1

Dim rangecells, rangecells1 As Range
Set rangecells = Sheet9.Range("Name_1").Cells
Set rangecells1 = Sheet9.Range("Amt_1").Cells

ReDim OEName(end_name1)
ReDim OEAmt(end_name1)
 
 For i = 1 To end_name1
 OEName(i) = rangecells.item(i).value
 OEAmt(i) = rangecells1.item(i).value


If Len(OEName(i)) > 50 Then
    msgBx_BP2 = msgBx_BP2 + "* Nature cannot be exceed 50 Characters  in Sheet BP" & Chr(13)
    ValidateExemptInc = False
    Exit Function
End If


If Len(OEAmt(i)) > 14 Then
    msgBx_BP2 = msgBx_BP2 + "* Amount in Any Other Exempt cannot exceed 14 digits in Sheet BP" & Chr(13)
    ValidateExemptInc = False
    Exit Function
End If

Next

If end_name1 > end_amt1 Then
    ValidateExemptInc = False
    Sheet9.Activate
    Sheet9.Range("Amt_1").Select
'    MsgBx_BP1 = "* Please Enter the Amount"
    MsgBx_BP1 = "* ""Please enter the amount in schedule BP"""
ElseIf end_name1 < end_amt1 Then
    ValidateExemptInc = False
'    MsgBx_BP1 = "* Please Enter the Nature"
    MsgBx_BP1 = "* ""Please specify the nature in schedule BP"""

    Sheet9.Activate
    Sheet9.Range("Name_1").Select
End If

End Function

Function ValidateDrpdn()
ValidateDrpdn = True

Dim rangecells As Range
Set rangecells = Sheet9.Range("EI.Drpdn").Cells
setTableinfo_EId
Dim i As Long
'dpkp

If Sheet9.Range("Sheet12.NetPLFrmSpecifiedBus").value <> 0 And SheetALL.Range("bacValue").value = 2 Then

If Not end_BP > 0 Then
If Sheet9.Range("EI.Drpdn").Locked = False Then
'    msgBx_BP2 = msgBx_BP2 + "* Please select the Relevant clause of sub-section (5) of section 35AD in Sheet BP. " & Chr(13)
    msgBx_BP2 = msgBx_BP2 + "* ""Please select the option from the drop down at Sl No 49 in schedule BP""" & Chr(13)

    ValidateDrpdn = False
    Exit Function
    End If
    
End If
End If

ReDim BP_Drpdn(end_BP)
For i = 1 To end_BP
BP_Drpdn(i) = rangecells.item(i).value
If Len(BP_Drpdn(i)) = 0 Then
End If

If Sheet9.Range("Sheet12.NetPLFrmSpecifiedBus").value <> 0 And SheetALL.Range("bacValue").value = 2 Then
If Sheet9.Range("sheet12.DeductUs35AD").value > 0 Then
If BP_Drpdn(i) = "(Select)" Or BP_Drpdn(i) = "" Then
'If BP_Drpdn(i) = "" Then

'    msgBx_BP2 = msgBx_BP2 + "* Please select the Relevant clause of sub-section (5) of section 35AD in Sheet BP." & Chr(13)
    msgBx_BP2 = msgBx_BP2 + "* ""Please select the option from the drop down at Sl No 49 in schedule BP""" & Chr(13)
    ValidateDrpdn = False
    Exit Function
End If
End If
End If
Next

End Function
Sub setTableinfo_name1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet9.Range("Name_1").count
    Set rangecells = Sheet9.Range("Name_1").Cells
     For mIntCtr = 1 To mIntCells
            If Not ((rangecells.item(mIntCtr).value = "")) Then
                ccount = ccount + 1
            End If
    Next
 end_name1 = ccount
 rngname_BPOE = "Name_1;Amt_1;"
End Sub


Sub setTableinfo_amt1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet9.Range("Amt_1").count
    Set rangecells = Sheet9.Range("Amt_1").Cells
     For mIntCtr = 1 To mIntCells
            If Not ((rangecells.item(mIntCtr).value = "")) Then
                ccount = ccount + 1
            End If
    Next
 end_amt1 = ccount
End Sub
Sub AddRows_BPOE()
    Dim vRows As Long
Sheets("BP").Activate
setTableinfo_name1
ActiveCellRange = EfilingCommon.searchLastRow("Name_1")
vRows = InsertRowsAndFillFormulas()
Call ExendRangeNameToTable(vRows, rngname_BPOE)
Dim RngHOI, rrCount, rr
Set RngHOI = Selection.Resize(rowsize:=vRows + 1)
        rrCount = 0
        For Each rr In RngHOI.Rows
            If Not rrCount = 0 Then
                Range("G" & rr.row).ClearContents
                Range("I" & rr.row).ClearContents
            End If
            rrCount = rrCount + 1
        Next
End Sub


Sub AddRows_BPOEold1()
    Dim vRows As Long
Sheets("BP").Activate
setTableinfo_name1
ActiveCellRange = EfilingCommon.searchLastRow("Name_1")
vRows = InsertRowsAndFillFormulas()
Call ExendRangeNameToTable(vRows, rngname_BPOE)
End Sub


Sub LockBp()
 Application.ScreenUpdating = False
  Application.EnableEvents = False
  Dim Rangename, name As Variant
  Dim i As Variant
  Rangename = "sheet10.Inc115BBF||sheet11.ExpDebToPLOth115BF"
  Rangename = Split(Rangename, "||")
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet9.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
For i = 0 To UBound(Rangename)
    Sheet9.Range(Rangename(i)).value = ""
    Sheet9.Range(Rangename(i)).Interior.Color = (&HD8D8D8)
    Sheet9.Range(Rangename(i)).Locked = True
Next

 Sheet9.Protect Password:=sPassword
 
 Application.ScreenUpdating = True
 Application.EnableEvents = True
 End Sub

Sub UnLockBp()
Application.ScreenUpdating = False
  Application.EnableEvents = False
  Dim Rangename, name As Variant
  Dim i As Variant
  Rangename = "sheet10.Inc115BBF||sheet11.ExpDebToPLOth115BF"
  Rangename = Split(Rangename, "||")
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet9.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
For i = 0 To UBound(Rangename)
    Sheet9.Range(Rangename(i)).value = ""
    Sheet9.Range(Rangename(i)).Interior.Color = (&HCCFFCC)
    Sheet9.Range(Rangename(i)).Locked = False
Next

'-----------------Lock Password-----------------------------
 Sheet9.Protect Password:=sPassword
'-----------------------------------------------------------
 Application.ScreenUpdating = True
 Application.EnableEvents = True
End Sub

'Updated by Malli on 28/03/2024
'Function ValidateBP_EI()
'ValidateBP_EI = True
'
'
'Dim dividends_tot As Variant
'dividends_tot = Sheet9.Range("sheet10.IncRecCredPLOthHeads4Div").value + Sheet9.Range("DivIncome_1").value
'If dividends_tot > Sheet4.Range("PL.Dividends").value Then
'
''  MsgBox """Amount mentioned in this field plus A.5c(1) - Dividend Income"""" cannot be more than the amount of Dividend income mentioned in the Sl. No. 14(iii) of Part A Profit & loss account""", , "Error(s)"
'  MsgBox """Total amount mentioned in A.3d(i) - Dividend Income and A.5c(i) - Dividend Income cannot be more than the amount of Dividend income mentioned in the Sl. No. 14(iii) of Part A Profit & loss account""", , "Error(s)"
'  ValidateBP_EI = False
'  Exit Function
'End If
''-----------------------------
'End Function
