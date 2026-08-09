Attribute VB_Name = "mdNOBBP"
Option Explicit
Public end_Others, end_Amt As Variant
Public end_GSTIN As Variant
Public end_GSTAMOUNT As Variant
Public GSTIN_BP As Variant
Public GSTINAMOUNT_BP As Variant

Dim end_noofrow As Variant
Dim end_noofvalues As Variant
Dim errorMsg_NOB, errorMsg_NOB1 As Variant
Dim errorMsg_description, errorMsg_amount As Variant

Public BPCode As Variant
Public BPTradeName1 As Variant
Public BPTradeName2 As Variant
Public BPTradeName3 As Variant

'For E1 Table
Public GrossTurnover_Receipts As Variant 'AY 2024-25 Development
Public GrossTurnover_a As Variant
Public GrossTurnover_a_b As Variant      'AY 2024-25 Development
Public GrossTurnover_b As Variant
'Malli_AY_2025_26
Public GrossTurnoverE1bandE1C As Variant

'For E3 Table
Public GrossReceipts_E3_a  As Variant  'AY 2024-25 Development
Public GrossReceipts_E3_b As Variant   'AY 2024-25 Development
Public GrossReceipts_E3_c  As Variant  'AY 2024-25 Development


Public GrossTurnover As Variant
Public PresumptiveIncome44AD_a As Variant
Public PresumptiveIncome44AD_b As Variant
Public PresumptiveIncome44AD As Variant
Public GrossTurnover44ADA As Variant
Public PresumptiveIncome44ADa As Variant
Public SalaryInterestPaid44AD As Variant
Public TotalBPE15 As Variant
Public TotalBPE23 As Variant


Public PartnerMemberOwnCapital As Variant
Public TotSecuredLoans As Variant
Public TotUnSecuredLoans As Variant
Public TotAdvances As Variant
Public TotSundryCreditors As Variant
Public TotOthrCurrLiab As Variant
Public FixedAssets As Variant
Public Inventories As Variant
Public SundryDebtors As Variant
Public BalWithBanks As Variant
Public TotCashInHand As Variant
Public LoansAndAdvances As Variant
Public OtherAssets As Variant
Public GSTRNo As Variant
Public AmonutGSTR As Variant


Public IncomeChargeabeBPUs44AE As Variant
Public IncomeChargeabeBPUs As Variant

Public TotalPresumptiveIncome44AD As Variant
Public PresumptiveIncomeHV As Variant
Public IncomeChargeabe As Variant
Public AmountSundryDebators As Variant
Public AmountSundryCreditors As Variant
Public AmountTotalStockInTrade As Variant
Public AmountCashBalance As Variant
Public end_NOB_Code, end_NOB_Trade1, end_NOB_Trade2, end_NOB_Trade3 As Long
Public MsgBx_Grid As String
Public msgBx14NOBBP As String

Public end_NOBC As Variant
Public end_NOBName As Variant
Public end_NOBDescription As Variant
Public end_NOBC1 As Variant
Public end_NOBName1 As Variant
Public end_NOBDescription1 As Variant
Public end_NOBC2 As Variant
Public end_NOBName2 As Variant
Public end_NOBDescription2 As Variant
Public errorMsg_NOB2 As Variant
Public errorMsg_NOB3 As Variant
Public end_NOBC3 As Variant
Public end_NOBName3 As Variant
Public end_NOBDescription3 As Variant
Public errorMsg_NOB4, errorMsg_NOB5 As Variant
Public end44AE_RegNo, end44AE_OLH, end44AE_TC, end44AE_Months, end44AE_PIncome As Variant
Public SEC44AERegNo, SEC44AEOLH, SEC44AETC, SEC44AEMonth, SEC44AEPIncome, Greater12_BP, Lesser12_BP As Variant

Public NOBGSTIN, NOBCode3, NOBCode2, NOBCode1, NOBCode As Variant
Public NOBName3, NOBName2, NOBName1, NOBName As Variant
Public NOBDescription3, NOBDescription2, NOBDescription1, NOBDescription As Variant
Public end44AD, end44ADA, end44AE As Variant

'Newly added by Bindu
Public GrossReceipts_E3_b_c As Variant



Sub cmdValidateNOBBP_Click()
ValidateSheetNOBBP
fmsgboxStatus "Sheet BP is Ok" ', vbOKOnly, vbMessageCaption
End Sub

Sub cmdPrev_Click_NOB()

    If Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" Then
       Sheet201.Activate
    ElseIf Sheet22.Visible = True Then
        Sheet22.Activate
    ElseIf Sheet15.Visible = True Then
        Sheet15.Activate
    Else
       Sheet1.Activate
   
    End If
End Sub
Sub cmdNext_Click()
Sheet4.Activate
End Sub
Sub cmdHelp_Click()
ThisWorkbook.Unprotect Password:=getmsgstate
Sheet12.Visible = xlSheetVisible
Sheet12.Activate
ThisWorkbook.Protect Password:=getmsgstate
End Sub

Sub ValidateSheetNOBBP()
msgError = ""
mdNOBBP.AssignValuesNOBBP


If Not mdNOBBP.ValidateSheetBP Then
Sheet3.Activate
fmsgboxsmall msgError ', vbOKOnly, "Error(s)!"
CloseMsg
End If


If Not ValidateNOBBP14Digits Then
Sheet3.Activate
fmsgboxsmall msgBx14NOBBP ', vbOKOnly, "Error(s)!"
CloseMsg
End If

End Sub


Function ValidateNOBBP14Digits() As Boolean
ValidateNOBBP14Digits = True
    msgBx14NOBBP = ""
If Len(Sheet3.Range("BP_E3").value) > 14 Then
    ValidateNOBBP14Digits = False
    msgBx14NOBBP = msgBx14NOBBP + "* Total Presumptive Income from goods Carriage u/s 44AE cannot exceed 14 digits in Sheet BP" & Chr(13)
End If
If (Len(Sheet3.Range("BP_E9").value) > 14 Or Len(Sheet3.Range("BP_E10").value) > 14 Or Len(Sheet3.Range("BP_E11").value) > 14 Or Len(Sheet3.Range("BP_E12").value) > 14 Or Len(Sheet3.Range("BP_E13").value) > 14 Or Len(Sheet3.Range("BP_E14").value) > 14) Then
    ValidateNOBBP14Digits = False
    msgBx14NOBBP = msgBx14NOBBP + "* Amount entered cannot exceed 14 digits in Sheet BP" & Chr(13)
End If

If Len(Sheet3.Range("BP_E15_Total").value) > 14 Then
    ValidateNOBBP14Digits = False
    msgBx14NOBBP = msgBx14NOBBP + "* Total sources of funds cannot exceed 14 digits in Sheet BP" & Chr(13)
End If

If (Len(Sheet3.Range("BP_E16").value) > 14 Or Len(Sheet3.Range("BP_E17").value) > 14 Or Len(Sheet3.Range("BP_E18").value) > 14 Or Len(Sheet3.Range("BP_E19").value) > 14 Or Len(Sheet3.Range("BP_E20").value) > 14 Or Len(Sheet3.Range("BP_E21").value) > 14 Or Len(Sheet3.Range("BP_E15_Total").value) > 22) Then
    ValidateNOBBP14Digits = False
    msgBx14NOBBP = msgBx14NOBBP + "* Amount entered cannot exceed 14 digits in Sheet BP" & Chr(13)
End If

If Len(Sheet3.Range("BP_E23_total").value) > 14 Then
    ValidateNOBBP14Digits = False
    msgBx14NOBBP = msgBx14NOBBP + "* Total application of funds cannot exceed 14 digits in Sheet BP" & Chr(13)
End If

If Len(Sheet3.Range("TotalGST").value) > 14 Then
    ValidateNOBBP14Digits = False
    msgBx14NOBBP = msgBx14NOBBP + "Total of value of Outward Supplies as per the GST returns filed cannot exceed 14 digits" & Chr(13)
End If

'
'If Len(Sheet3.Range("BP.GSTRAmount").value) > 14 Then
'    ValidateNOBBP14Digits = False
'    msgBx14NOBBP = msgBx14NOBBP + "GSTR Amount cannot exceed 14 digits" & Chr(13)
'End If




'If Len(Sheet3.Range("BP_E6").value) > 14 Then
'    ValidateNOBBP14Digits = False
'    msgBx14NOBBP = msgBx14NOBBP + "Income chargeble under Business cannot exceed 14 digits" & Chr(13)
'End If

End Function

Function AssignValuesNOBBP()
On Error Resume Next
'Table E1 Starts-------------------
GrossTurnover_Receipts = Range("BP_E1_GR").value  ' AY 2024-25
GrossTurnover_a = Range("BP_E1a").value
GrossTurnover_a_b = Range("BP_E1a_E1b").value     ' BP_C3 AY 2024-25
GrossTurnover_b = Range("BP_E1b").value
GrossTurnoverE1bandE1C = Range("BP_E1a_E1b") + Range("BP_E1b").value
'Table E1 End ---------------------

GrossTurnover = Range("BP_E1").value
PresumptiveIncome44AD_a = Range("BP_E2a").value
PresumptiveIncome44AD_b = Range("BP_E2b").value
PresumptiveIncome44AD = Range("BP_E2_PI").value

'Table E3 Starts------------------------

GrossTurnover44ADA = Range("BP_E3_GR").value
'AY 2024-25 Development
GrossReceipts_E3_a = Range("BP_E3_a").value  'BP_C10 AY 2024-25
GrossReceipts_E3_b = Range("BP_E3_b").value  'BP_C10 AY 2024-25
GrossReceipts_E3_c = Range("BP_E3_c").value  'BP_C10 AY 2024-25


'Newly added by Bindu post confirmation with Aparna on 29th May 25
GrossReceipts_E3_b_c = Range("BP_E3_b").value + Range("BP_E3_c").value


'Table E3 End---------------------------

PresumptiveIncome44ADa = Range("BP_E4_PI44ADA").value

SalaryInterestPaid44AD = Range("BP_E3_2").value
IncomeChargeabeBPUs44AE = Range("BP_E7_IC44AE").value
IncomeChargeabeBPUs = Range("BP_E8_IC_BP").value
'TotalPresumptiveIncome44AD = Range("BP_E4_2").value

PresumptiveIncomeHV = Range("BP_E3").value
'IncomeChargeabe = Range("BP_E6").value
'AmountSundryDebators = Range("BP_E7").value
'AmountSundryCreditors = Range("BP_E8").value
'AmountTotalStockInTrade = Range("BP_E9").value
'AmountCashBalance = Range("BP_E10").value
PartnerMemberOwnCapital = Range("BP_E9").value
TotSecuredLoans = Range("BP_E10").value
TotUnSecuredLoans = Range("BP_E11").value
TotAdvances = Range("BP_E12").value
TotSundryCreditors = Range("BP_E13").value
TotOthrCurrLiab = Range("BP_E14").value
TotalBPE15 = Range("BP_E15_total").value
FixedAssets = Range("BP_E16").value
Inventories = Range("BP_E17").value
SundryDebtors = Range("BP_E18").value
BalWithBanks = Range("BP_E19").value
TotCashInHand = Range("BP_E20").value
LoansAndAdvances = Range("BP_E21").value
OtherAssets = Range("BP_E22").value
TotalBPE23 = Range("BP_E23_total").value
GSTRNo = Range("BP.GSTRno").value
AmonutGSTR = Range("BP.GSTRAmount").value


End Function

Function ValidateSheetBP() As Boolean
On Error Resume Next
ValidateSheetBP = True

'BP_E2 AY 2023-24 Old
'If (IIf(GrossTurnover_a = "", 0, GrossTurnover_a) + IIf(GrossTurnover_b = "", 0, GrossTurnover_b)) > 20000000 Then msgError = msgError & "* If Gross Receipts are more than Rs 2 crores , it is mandatory to have a tax audit under 44AB. Please use the regular ITR 3/5 form" & Chr(13)

'BP_C2 AY 2024-25 Change
If (IIf(GrossTurnover_a = "", 0, GrossTurnover_a) + IIf(GrossTurnover_a_b = "", 0, GrossTurnover_a_b) + IIf(GrossTurnover_b = "", 0, GrossTurnover_b)) > 30000000 Then msgError = msgError & "* If Gross Receipts are more than Rs 3 crores , it is mandatory to have a tax audit under 44AB. Please use the regular ITR 3/5 form" & Chr(13)

'''''BP_E8 AY 2023-24  old
'If IIf(GrossTurnover44ADA = "", 0, GrossTurnover44ADA) > 5000000 Then msgError = msgError & "*  If Gross Receipts are more than Rs.50 Lakh, it is mandatory to have a tax audit under 44AB and Regular ITR 3/5 form has to be filled and not this form" & Chr(13)

''''''BP_C8 AY 2024_2025 Change
'If IIf(GrossTurnover44ADA = "", 0, GrossTurnover44ADA) > 7500000 Then msgError = msgError & "*  If Gross Receipts are more than Rs.75 Lakh, it is mandatory to have a tax audit under 44AB and Regular ITR 3/5 form has to be filled and not this form" & Chr(13)

'If Not ValidateNOB Then msgError = msgError & errorMsg_NOB & Chr(13)
If Not ValidateNOB1 Then msgError = msgError & errorMsg_NOB1 & Chr(13)
If Not ValidateNOB2 Then msgError = msgError & errorMsg_NOB2 & Chr(13)
If Not ValidateNOB3 Then msgError = msgError & errorMsg_NOB3 & Chr(13)
'If Not ValidateNOB4 Then msgError = msgError & errorMsg_NOB4 & Chr(13)
If Not ValidateSec44AE Then msgError = msgError & errorMsg_NOB5 & Chr(13)
'
'setTblinfo_BPNatureItem
'setTblinfo_BPNatureAmount

'end_Others = WorksheetFunction.Max(0, end_Others, end_Amt)
'If Not ValidateGSTNature Then msgError = msgError & errorMsg_description & Chr(13)
'If Not ValidateGSTNewAmount Then msgError = msgError & errorMsg_amount & Chr(13)



'If IIf(GrossTurnover_a = "", 0, GrossTurnover_a) < PresumptiveIncome44AD_a Then msgError = msgError & "E2(a) should not be greater than E1(a)" & Chr(13)
'If IIf(GrossTurnover_b = "", 0, GrossTurnover_b) < PresumptiveIncome44AD_b Then msgError = msgError & "E2(b) should not be greater than E1(c)" & Chr(13)

'BP_C6 AY 2024-25 Change
If (IIf(GrossTurnover_b = "", 0, GrossTurnover_b) + IIf(GrossTurnover_a_b = "", 0, GrossTurnover_a_b)) < PresumptiveIncome44AD_b Then msgError = msgError & "E2(b) should not be greater than E1(b+c)" & Chr(13)

'BP_E6 AY 2023-24 OLD

'If GrossTurnover_a = 0 And GrossTurnover_b = 0 Then
'
'If (IIf(GrossTurnover_a = "", 0, GrossTurnover_a) + IIf(GrossTurnover_b = "", 0, GrossTurnover_b)) < (IIf(PresumptiveIncome44AD_a = "", 0, PresumptiveIncome44AD_a) + IIf(PresumptiveIncome44AD_b = "", 0, PresumptiveIncome44AD_b)) Then msgError = msgError & "* Invalid presumptive income under 44AD, as gross turnover is zero in Sheet BP" & Chr(13)

'BP_C3 AY 2024-25 Change

If GrossTurnover_a = 0 And GrossTurnover_b = 0 And GrossTurnover_a_b = 0 Then

'If (IIf(GrossTurnover_a = "", 0, GrossTurnover_a) + IIf(GrossTurnover_a_b = "", 0, GrossTurnover_a_b) + IIf(GrossTurnover_b = "", 0, GrossTurnover_b)) < (IIf(PresumptiveIncome44AD_a = "", 0, PresumptiveIncome44AD_a) + IIf(PresumptiveIncome44AD_b = "", 0, PresumptiveIncome44AD_b)) Then msgError = msgError & "* Invalid presumptive income under 44AD, as gross turnover is zero in Sheet BP" & Chr(13)
' AY 2024-25 Change
If (IIf(GrossTurnover_a = "", 0, GrossTurnover_a) + IIf(GrossTurnover_a_b = "", 0, GrossTurnover_a_b) + IIf(GrossTurnover_b = "", 0, GrossTurnover_b)) < (IIf(PresumptiveIncome44AD_a = "", 0, PresumptiveIncome44AD_a) + IIf(PresumptiveIncome44AD_b = "", 0, PresumptiveIncome44AD_b)) Then msgError = msgError & "* Invalid presumptive income under 44AD, as gross turnover is zero in 44AD section" & Chr(13)
End If

If IIf(PresumptiveIncome44AD_a = "", 0, PresumptiveIncome44AD_a) < (0.06 * IIf(GrossTurnover_a = "", 0, GrossTurnover_a)) Then msgError = msgError & "* If the presumptive income E2(a) is less than 6% of Gross Receipts through a/c payee or electronic clearing system, it is mandatory to have a tax audit under 44AB. Please use the regular ITR 3 or ITR 5 form." & Chr(13)
'If IIf(PresumptiveIncome44AD_b = "", 0, PresumptiveIncome44AD_b) < (0.08 * IIf(GrossTurnover_b = "", 0, GrossTurnover_b)) Then msgError = msgError & "* If the presumptive income E2(b) is less than 8% of Gross Receipts through any other mode, it is mandatory to have a tax audit under 44AB. Please use the regular ITR 3 or ITR 5 form." & Chr(13)

'BP_C3 AY 2024-25 Change
If IIf(PresumptiveIncome44AD_b = "", 0, PresumptiveIncome44AD_b) < (0.08 * (IIf(GrossTurnover_a_b = "", 0, GrossTurnover_a_b) + IIf(GrossTurnover_b = "", 0, GrossTurnover_b))) Then msgError = msgError & "* If the income is less than 8% of E1(b+c), it is mandatory to have a tax audit under 44AB. Please use  ITR 3 if Individual/HUF and ITR 5 if Firm." & Chr(13)
''''''''''''''''''''''

'Chandru
   If Range("BP_E2a").value > Range("BP_E1a").value Then
      msgError = msgError & "* E2(a)should not be greater than E1(a)" & Chr(13)
     ' Range("BP_E2a").value = ""
   End If



'AY 2024-25 CHANGES STARTS----------------

 ValidateSecGROSSTURNOVER
 ValidateSecGROSS_Receipts
 
 'AY 2024-25 CHANGES ENDS----------------


 'BP_C8 AY 2024-25 Change
 If (GrossTurnover44ADA = "") Or GrossTurnover44ADA = "0" And (PresumptiveIncome44ADa > 0) Then msgError = msgError & "* Invalid presumptive income under 44ADA, as gross turnover is zero" & Chr(13)
 
''''''''''''''''''testing
If (GrossTurnover44ADA = "") And (GrossTurnover44ADA < 0) Then msgError = msgError & "* Invalid presumptive income under 44ADA, as gross turnover is zero in 44ADA section in Sheet BP" & Chr(13)
  

''''''''''''''''
If GrossTurnover44ADA < PresumptiveIncome44ADa Then msgError = msgError & "* Presumptive Income under section 44ADA  of Schedule BP should not be greater than Gross Receipts ." & Chr(13)

''For testing purpose of AY 2024_25 change for Table 3 related

'BP_C8,BP_C9,BP_C10 AY 2024_25 Change
'If (IIf(GrossReceipts_E3_a = "", 0, GrossReceipts_E3_a) + IIf(GrossReceipts_E3_b = "", 0, GrossReceipts_E3_b) + IIf(GrossReceipts_E3_c = "", 0, GrossReceipts_E3_c)) > 7500000 Then msgError = msgError & "* If Gross Receipts are more than Rs 7500000, it is mandatory to have a tax audit under 44AB. Please use the regular ITR 3/5 form" & Chr(13)


If IIf(PresumptiveIncome44ADa = "", 0, PresumptiveIncome44ADa) < (0.5 * IIf(GrossTurnover44ADA = "", 0, GrossTurnover44ADA)) Then msgError = msgError & "* If the income entered at E4 is less than 50% of Gross Receipts entered at E3, it is mandatory to have a tax audit under 44AB. Please use the regular ITR 3/5 form." & Chr(13)

Dim income_BP As Variant
 income_BP = Sheet3.Range("BP_E2_PI").value + IIf(Sheet3.Range("BP_E4_PI44ADA").value = "", 0, Sheet3.Range("BP_E4_PI44ADA").value) + Sheet3.Range("BP_E3").value
    
' If (TotalBPE15 <> TotalBPE23) Then msgError = msgError & "E15(Sources of fund) shall be equal to E23 (Total application of funds)." & Chr(13)
'If (income_BP = "" Or income_BP = 0) Then msgError = msgError & "For filling ITR 4 presumptive business & Profession income is Mandatory. If there is no presumptive business Income please fill other ITR." & Chr(13)
If msgError <> "" Then ValidateSheetBP = False

 
 

'If ((Sheet3.Range("BP_E1a").value + (Sheet3.Range("BP_E1b").value) + (Sheet3.Range("BP_E3_GR").value)) > 2000000) Then
'    If Not chkCompulsory(GSTRNo) Then
'             msgError = msgError + "GSTR Number is Mandatory" & Chr(13)
'            ValidateSheetBP = False
'            Exit Function
'        End If
'End If
'
'
'If ((Sheet3.Range("BP_E1a").value + (Sheet3.Range("BP_E1b").value) + (Sheet3.Range("BP_E3_GR").value)) > 2000000) Then
'If (Sheet3.Range("BP.GSTRno").value) <> "" Then
'    If Not chkCompulsory(AmonutGSTR) Then
'             msgError = msgError + " Please enter the amount of turnover /gross reciept as per the GST return filed ." & Chr(13)
'            ValidateSheetBP = False
'            Exit Function
'        End If
'End If
'End If

'If (Sheet3.Range("BP.GSTRno").value) <> "" Then
'    If Not chkCompulsory(AmonutGSTR) Then
'             msgError = msgError + " Please enter the amount of turnover /gross reciept as per the GST return filed ." & Chr(13)
'            ValidateSheetBP = False
'            Exit Function
'        End If
'End If
'
'If AmonutGSTR <> "" Then
'    If Not chkCompulsory(Sheet3.Range("BP.GSTRno").value) Then
'             msgError = msgError + " Please enter the GSTR Number." & Chr(13)
'            ValidateSheetBP = False
'            Exit Function
'        End If
'End If

setTblinfo_GSTAMOUNT
setTblinfo_GSTIN
If Not ValidateGSTIN Then ValidateSheetBP = False
'If Not ValidateGSTAmount Then ValidateSheetBP = False



If Sheet3.Range("TotalGST").value > 30000000 Then
            msgError = msgError + "* As per GST, your total turnover exceeds 3 Cr, please check if you are eligible for ITR-4"
            ValidateSheetBP = False
            Exit Function
End If

If Sheet3.Range("BP_E8_IC_BP").value = 0 Then
msgError = msgError + "*You cannot file ITR 4 if you do not have income from business u/s 44AD or 44AE or income from profession u/s 44ADA"
ValidateSheetBP = False
Exit Function
End If

If end_GSTIN <> end_GSTAMOUNT Then
            msgError = msgError + "* All the fields in Information Regarding Turnover/Gross Reciept Reported for GST is Mandatory in Sheet BP" & Chr(13)
            ValidateSheetBP = False
            Exit Function
End If



If ((Sheet3.Range("BP_E1a").value + (Sheet3.Range("BP_E1b").value) + (Sheet3.Range("BP_E3_GR").value) + (Sheet3.Range("BP_E3").value)) > 0) Then

'issue no 65 15032022
'If (TotSundryCreditors = "" Or Inventories = "" Or Sundrycreditors = "" Or TotCashInHand = "") Then
'           msgError = msgError + "* enter all the deatail that given ( totsundrycreditors,inventories,sundrydebtors,totcashinhand)"
'           ValidateSheetBP = False
'           Exit Function
'       End If
        
If TotSundryCreditors = "" Then
            msgError = msgError + "* Please enter amount in Sundry creditors"
            ValidateSheetBP = False
            Exit Function
        End If
If Inventories = "" Then
            msgError = msgError + "* Please enter amount in Inventories"
            ValidateSheetBP = False
            Exit Function
        End If
If SundryDebtors = "" Then
            msgError = msgError + "* Please enter amount in Sundry Debtors"
            ValidateSheetBP = False
            Exit Function
        End If
        
    If BalWithBanks = "" Then
        msgError = msgError + "* Please enter amount in ""Balance with Banks"""
            ValidateSheetBP = False
            Exit Function
    End If
If TotCashInHand = "" Then
            msgError = msgError + "* Please enter amount in 'Cash-In-Hand'"
            ValidateSheetBP = False
            Exit Function
        End If
        
'issue no 65
End If



If Len(Sheet3.Range("BP_E8_IC_BP").value) > 14 Then
    msgError = msgError + "*  Income chargeable under Business or Profession (E2c+ E4+ E7) must not exceed 14 Digits"
    ValidateSheetBP = False
End If


End Function

'Function ValidateNOB() As Boolean
'On Error Resume Next
'Dim cntr, cntr2, value As Variant
'value = 0
'ValidateNOB = True
'setTableInfo
'ReDim BPCode(end_noofvalues)
'ReDim BPTradeName1(end_noofvalues)
'ReDim BPTradeName2(end_noofvalues)
'ReDim BPTradeName3(end_noofvalues)
'cntr2 = 0
'For cntr = 1 To end_noofrow
'
' If Not CheckEmptyDDList(Range("NOB.Code").item(cntr).value) Then
' value = value + 1
' BPCode(cntr2) = Range("NOB.Code").item(cntr).value
' BPCode(cntr2) = Mid(BPCode(cntr2), 1, WorksheetFunction.Search("-", BPCode(cntr2)) - 1)
' BPTradeName1(cntr2) = Range("NOB.TradeNameOrDescription").item(cntr).value
' BPTradeName2(cntr2) = Range("NOB.TradeNameOrDescription2").item(cntr).value
' BPTradeName3(cntr2) = Range("NOB.TradeNameOrDescription3").item(cntr).value
' cntr2 = cntr2 + 1
' Else
'
' End If
'Next
'
'If value = 0 Then
'    Dim income_BP As Variant
'    income_BP = Sheet3.Range("BP_E2_PI").value + IIf(Sheet3.Range("BP_E4_PI44ADA").value = "", 0, Sheet3.Range("BP_E4_PI44ADA").value) + Sheet3.Range("BP_E3").value
'    If income_BP > 0 Then
'        errorMsg_NOB = errorMsg_NOB & "Atleast One Entry is Required in Nature of Business"
'        ValidateNOB = False
'    End If
'Else
'For cntr = 0 To value - 1
'    If Not ChkTradeName(BPTradeName1(cntr)) Then errorMsg_NOB = errorMsg_NOB & "Trade Name at Sr No. " & cntr + 1 & " Charachters < > are not allowed" & Chr(13)
'    If Not ChkTradeName(BPTradeName2(cntr)) Then errorMsg_NOB = errorMsg_NOB & "Trade Name at Sr No. " & cntr + 1 & " Charachters < > are not allowed" & Chr(13)
'    If Not ChkTradeName(BPTradeName3(cntr)) Then errorMsg_NOB = errorMsg_NOB & "Trade Name at Sr No. " & cntr + 1 & " Charachters < > are not allowed" & Chr(13)
'Next cntr
'If errorMsg_NOB <> "" Then ValidateNOB = False
'End If
'
'End Function

Function ChkTradeName(name As Variant) As Boolean
On Error Resume Next
Dim chararr As Variant
Dim cntr1, cntr2 As Variant
chararr = Array("<", ">")
ChkTradeName = True
For cntr1 = 1 To Len(name)
    For cntr2 = 0 To UBound(chararr) - 1
        If Mid(name, cntr1, 1) = chararr(cntr2) Then
        ChkTradeName = False
        Exit Function
        End If
    Next cntr2
    Next cntr1

End Function

'Sub setTableInfo()
'
'Dim rangecells As Range
'Dim mIntCells As Long
'Dim mIntCtr As Long
'Dim ccount As Long
'ccount = 0
'mIntCells = Sheet3.Range("NOB.Code").Count
'Set rangecells = Sheet3.Range("NOB.Code").Cells
'Dim countrycd As Variant
'
'For mIntCtr = 1 To mIntCells
'    countrycd = rangecells.item(mIntCtr).value
'    If Not CheckEmptyDDList(countrycd) Then ccount = ccount + 1
'Next
'end_noofvalues = ccount
'end_noofrow = mIntCells
'
'End Sub



Sub Lock44AD()
On Error Resume Next
    Dim strpassword As String
    
    strpassword = mdCommon.getmsgstate
    Sheet3.Unprotect Password:=strpassword
    
    Dim rngname44AD As Variant
    'rngname44AD = "BP_E1;BP_E2;"
    rngname44AD = "BP_E1a;BP_E1b;BP_E2a;BP_E2b;BP_E4_PI44ADA;BP_E3_GR;"
    LockRange44AD (rngname44AD)
    
    Sheet3.Protect Password:=strpassword
End Sub


Sub Unlock44AD()
On Error Resume Next
    Dim strpassword As String
    
    strpassword = mdCommon.getmsgstate
    Sheet3.Unprotect Password:=strpassword
    Dim rngname44AD As Variant
    
    'rngname44AD = "BP_E1;BP_E2;"
    rngname44AD = "BP_E1a;BP_E1b;BP_E2a;BP_E2b;BP_E4_PI44ADA;BP_E3_GR;"
    UnLockRange44AD (rngname44AD)
    
    Sheet3.Protect Password:=strpassword
End Sub

Sub LockRange44AD(rangenamestring As Variant)
Dim i As Long
Dim chcells As Range
    rangenamestring = Split(rangenamestring, ";")
        For i = 0 To UBound(rangenamestring) - 1
            For Each chcells In Sheet3.Range(rangenamestring(i))
                chcells.Locked = True
                chcells.Interior.Color = (&HD8D8D8)
                chcells.value = ""
            Next
        Next
End Sub

Sub UnLockRange44AD(rangenamestring As Variant)
Dim i As Long
Dim chcells As Range
    rangenamestring = Split(rangenamestring, ";")
        For i = 0 To UBound(rangenamestring) - 1
            For Each chcells In Sheet3.Range(rangenamestring(i))
            chcells.Locked = False
            chcells.Interior.Color = (&HCCFFCC)
            Next
        Next
End Sub

Function CheckGSTR(GSTRNo As Variant) As Boolean
On Error Resume Next
'GSTR No : Consist of 10 characters
'GSTR No: First 2 numeric, 5 Alphabets, next 4 numeric , next 1 Alphabet, then next 3 alphanumeric.

    CheckGSTR = True
    
'Change-06.05.2023.101 -> Prod Issue
    If Len(GSTRNo) > 15 Then
        CheckGSTR = False
        Exit Function
    End If
'---
    If Len(GSTRNo) > 0 Then
        If Not IsNumeric(Mid(GSTRNo, 1, 1)) Then
            CheckGSTR = False
            Exit Function
        End If
        If Not IsNumeric(Mid(GSTRNo, 2, 1)) Then
            CheckGSTR = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(GSTRNo, 3, 1)) Then
            CheckGSTR = False
            Exit Function
        End If
                
       If Not ChkAlphabet(Mid(GSTRNo, 4, 1)) Then
            CheckGSTR = False
            Exit Function
        End If
        
        If Not ChkAlphabet(Mid(GSTRNo, 5, 1)) Then
            CheckGSTR = False
            Exit Function
        End If
        
        If Not ChkAlphabet(Mid(GSTRNo, 6, 1)) Then
            CheckGSTR = False
            Exit Function
        End If
        
        If Not ChkAlphabet(Mid(GSTRNo, 7, 1)) Then
            CheckGSTR = False
            Exit Function
        End If
         
        If Not IsNumeric(Mid(GSTRNo, 8, 4)) Then
            CheckGSTR = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(GSTRNo, 12, 1)) Then
            CheckGSTR = False
            Exit Function
        End If
        If (Not ChkAlphabet(Mid(GSTRNo, 13, 1))) Then
        If (Not IsNumeric(Mid(GSTRNo, 13, 1))) Then
            CheckGSTR = False
            Exit Function
        End If
        End If
        
        
        If (Not ChkAlphabet(Mid(GSTRNo, 14, 1))) Then
        If (Not IsNumeric(Mid(GSTRNo, 14, 1))) Then
            CheckGSTR = False
            Exit Function
        End If
        End If
        
        If (Not ChkAlphabet(Mid(GSTRNo, 15, 1))) Then
        If (Not IsNumeric(Mid(GSTRNo, 15, 1))) Then
            CheckGSTR = False
            Exit Function
        End If
        End If
       
    End If
End Function

Sub NOBBP_ADD()
    Sheets("BP").Activate
    mdCommon.DefinedgridNameRange = "BP.GSTRno||BP.GSTRAmount"
    ActiveCellRange = mdCommon.searchLastRow("BP.GSTRno")
    mdCommon.insertRowUnderSectionWithFormula
End Sub
Sub NOB_AddRow()
    Sheets("BP").Activate
    mdCommon.DefinedgridNameRange = "Sec44AD.NOBName||Sec44AD.NOBCode||Sec44AD.NOBDescription"
    ActiveCellRange = mdCommon.searchLastRow("Sec44AD.NOBName")
    mdCommon.insertRowUnderSectionWithFormula
End Sub
Sub NOB_AddRow1()
    Sheets("BP").Activate
    mdCommon.DefinedgridNameRange = "Sec44AD.NOBName1||Sec44AD.NOBCode1||Sec44AD.NOBDescription1"
    ActiveCellRange = mdCommon.searchLastRow("Sec44AD.NOBName1")
    mdCommon.insertRowUnderSectionWithFormula
End Sub
Sub NOB_AddRow2()
    Sheets("BP").Activate
    mdCommon.DefinedgridNameRange = "Sec44AD.NOBName2||Sec44AD.NOBCode2||Sec44AD.NOBDescription2"
    ActiveCellRange = mdCommon.searchLastRow("Sec44AD.NOBName2")
    mdCommon.insertRowUnderSectionWithFormula
End Sub
'Sub NOB_AddRow3()
'    Sheets("NOB BP").Activate
'    mdCommon.DefinedgridNameRange = "Sec44AD.NOBName3||Sec44AD.NOBCode3||Sec44AD.NOBDescription3"
'    ActiveCellRange = mdCommon.searchLastRow("Sec44AD.NOBName3")
'    mdCommon.insertRowUnderSectionWithFormula
'End Sub
Sub setTblinfo_GSTIN()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet3.Range("BP.GSTRno").count
 Set rangecells = Sheet3.Range("BP.GSTRno").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_GSTIN = ccount
 
 End Sub
 
 
 Sub setTblinfo_GSTAMOUNT()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet3.Range("BP.GSTRAmount").count
 Set rangecells = Sheet3.Range("BP.GSTRAmount").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_GSTAMOUNT = ccount
 
 End Sub
 
 
 Function ValidateGSTIN() As Boolean
 
    ValidateGSTIN = True
    setTblinfo_GSTIN
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim i As Long
    Dim cellRange1 As String
    Set rangecells1 = Sheet3.Range("BP.GSTRno").Cells
    Set rangecells = Sheet3.Range("BP.GSTRAmount").Cells
    ReDim GSTINAMOUNT_BP(end_GSTIN)
    ReDim GSTIN_BP(end_GSTIN)
    For i = 1 To end_GSTIN
    
        cellRange1 = GetMergedAddressCell(rangecells1, i)
        GSTIN_BP(i) = Sheet3.Range(cellRange1).value
         GSTINAMOUNT_BP(i) = rangecells.item(i).value
       
        
        If Not mdNOBBP.CheckGSTR(GSTIN_BP(i)) Then
        msgError = msgError + "* Invalid GSTIN. GSTIN format should be First 2 numeric, 5 Alphabets, next 4 numeric , next 1 Alphabet, then next 3 alphanumeric  at Sr. No  " & i & " in Sheet BP " & Chr(13)
            ValidateGSTIN = False
            Exit Function
        End If
        
        
        If Not IsNumeric(GSTINAMOUNT_BP(i)) Then
        msgError = msgError + "* Amount of turnover/Gross receipt for the year corresponding to the GSTIN  at Sr. No  " & i & "in Sheet BP" & Chr(13)
            ValidateGSTIN = False
            Exit Function
        End If
        
        If end_GSTIN > 0 Then
         If Not chkCompulsory(GSTIN_BP(i)) Then
             msgError = msgError + "* Please enter the Value of Outward Supplies as per the GST return filed   at Sr. No  " & i & "in Sheet BP" & Chr(13)
            ValidateGSTIN = False
            Exit Function
        End If
        
        If Not chkCompulsory(GSTINAMOUNT_BP(i)) Then
             msgError = msgError + "* Please enter the Value of Outward Supplies as per the GSTIN return filed  at Sr. No  " & i & "in Sheet BP" & Chr(13)
            ValidateGSTIN = False
            Exit Function
        End If
        End If
        
        
        
     Next
End Function


'Function ValidateGSTAmount() As Boolean
'
'    ValidateGSTAmount = True
'    setTblinfo_GSTIN
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet3.Range("BP.GSTRAmount").Cells
'    ReDim GSTINAMOUNT_BP(end_GSTIN)
'    For i = 1 To end_GSTIN
'        GSTINAMOUNT_BP(i) = rangecells.item(i).value
'        If Not chkCompulsory(GSTINAMOUNT_BP(i)) Then
'             msgError = msgError + "Amount of turnover/Gross receipt for the year corresponding to the GSTIN at Sr. No  " & i & "  is Mandatory" & Chr(13)
'            ValidateGSTAmount = False
'            Exit Function
'        End If
'
'        If Not IsNumeric(GSTINAMOUNT_BP(i)) Then
'        msgError = msgError + "Amount of turnover/Gross receipt for the year corresponding to the GSTIN at Sr. No  " & i & "  is Mandatory" & Chr(13)
'            ValidateGSTAmount = False
'            Exit Function
'        End If
'
'
'
' Next
'End Function
'Sub GST_NewAddRow()
'    Sheets("BP").Activate
'    mdCommon.DefinedgridNameRange = "BP.NatureItem||BP.NatureAmount"
'    ActiveCellRange = mdCommon.searchLastRow("BP.NatureItem")
'    mdCommon.insertRowUnderSectionWithFormula
'End Sub

'Sub setTblinfo_BPNatureItem()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("BP.NatureItem").Count
'    Set rangecells = Range("BP.NatureItem").Cells
'    Dim countrycd As Variant
'    For mIntCtr = 1 To mIntCells
'            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
'                ccount = ccount + 1
'            End If
'    Next
' end_Others = ccount
' End Sub
'
' Sub setTblinfo_BPNatureAmount()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("BP.NatureAmount").Count
'    Set rangecells = Range("BP.NatureAmount").Cells
'    Dim countrycd As Variant
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_Amt = ccount
'End Sub
'
'Function ValidateGSTNature() As Boolean
'ValidateGSTNature = True
'
'    'setTblinfo_BPNatureItem
'    Dim rangecells As Range
'    Dim rangecells1 As Range
'    Dim cellrange As String
'    Dim cellrange1 As String
'    Set rangecells = Range("BP.NatureItem").Cells
'    Set rangecells1 = Range("BP.NatureDescription").Cells
'    Dim i As Long
'    ReDim Others_NOI_BP(end_Others)
'    ReDim Others_NOI1_BP(end_Others)
'
'
'    For i = 1 To end_Others
'
'    Others_NOI_BP(i) = rangecells.item(i).value
'     Others_NOI1_BP(i) = rangecells1.item(i).value
'
'
'
'
'         If isdropdownblank(Others_NOI_BP(i)) Then
'              errorMsg_description = errorMsg_description + "Please select an option from the drop down of nature of items at point E10(b) at Sr. No  " & i & "  in Sheet BP  is mandatory" & Chr(13)
'             ValidateGSTNature = False
'             Exit Function
'         End If
'
'          If (Others_NOI_BP(i) = "Any Other") Then
'          If Others_NOI1_BP(i) = "" Then
'             errorMsg_description = errorMsg_description & "Description of Nature of Items at point E10(b) at Sr. No  " & i & "  in Sheet BP  " & Chr(13)
'             ValidateGSTNature = False
'             Exit Function
'             End If
'         End If
'
'
'
''
''         If (Others_NOI(i) = "") Then
''         If Others_Amt(i) <> "" Then
''             errorMsg_description = errorMsg_description & "Please enter description of Nature of Items at point E10(b)" & Chr(13)
''             ValidateGSTNature = False
''             Exit Function
''         End If
''         End If
'
'
'
'
'    Next
'
'End Function
'Function ValidateGSTNewAmount() As Boolean
'ValidateGSTNewAmount = True
'
'    'setTblinfo_BPNatureItem
'    Dim rangecells As Range
'    Dim cellrange As String
'    Set rangecells = Range("BP.NatureAmount").Cells
'    Dim i As Long
'    ReDim Others_Amt_BP(end_Others)
'    For i = 1 To end_Others
'
'
'    Others_Amt_BP(i) = rangecells.item(i).value
'        If Not chkCompulsory(Others_Amt_BP(i)) Then
'            errorMsg_amount = errorMsg_amount & "Please enter Amount at Point E10(b)" & Chr(13)
'            ValidateGSTNewAmount = False
'            Exit Function
'        End If
'
''
''        If Not IsNumeric(Others_Amt(i)) Then
''            errorMsg_amount = errorMsg_amount & "Amount at Sr. No  " & i & "  in Sheet NOB BP should be Numeric value" & Chr(13)
''            ValidateGSTNewAmount = False
''            Exit Function
''        End If
''
''        If Others_Amt(i) > 99999999999999# Then
''            errorMsg_amount = errorMsg_amount & "Amount at Sr. No  " & i & "  in Sheet NOB BP cannot exceed 14 digits" & Chr(13)
''            ValidateGSTNewAmount = False
''            Exit Function
''        End If
'    Next
'End Function

Sub setTblinfo_NOBC()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AD.NOBCode").count
    Set rangecells = Range("Sec44AD.NOBCode").Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
 end_NOBC = ccount
 End Sub
 
 Sub setTblinfo_NOBName()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AD.NOBName").count
    Set rangecells = Range("Sec44AD.NOBName").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_NOBName = ccount
End Sub
Sub setTblinfo_NOBDescription()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AD.NOBDescription").count
    Set rangecells = Range("Sec44AD.NOBDescription").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_NOBDescription = ccount
End Sub


Function ValidateNOB1() As Boolean
On Error Resume Next

ValidateNOB1 = True

Dim cellrange As String
Dim cellRange1 As String
Dim cellrange2 As String
Dim rangecells As Range
Dim rangecells1 As Range
Dim rangecells2 As Range
Dim value As Variant
Dim i As Long
Set rangecells = Range("Sec44AD.NOBCode").Cells
Set rangecells1 = Range("Sec44AD.NOBName").Cells
Set rangecells2 = Range("Sec44AD.NOBDescription").Cells
setTblinfo_NOBC
setTblinfo_NOBName
setTblinfo_NOBDescription




end44AD = WorksheetFunction.Max(0, end_NOBName, end_NOBC, end_NOBDescription)

ReDim NOBCode(end44AD)
ReDim NOBName(end44AD)
ReDim NOBDescription(end44AD)

For i = 1 To end44AD

 cellrange = GetMergedAddressCell(rangecells, i)
 cellRange1 = GetMergedAddressCell(rangecells1, i)
 cellrange2 = GetMergedAddressCell(rangecells2, i)
 
value = value + 1
 NOBCode(i) = LTrim(Sheet3.Range(cellrange).value)
 NOBCode(i) = Mid(NOBCode(i), 1, WorksheetFunction.Search("-", NOBCode(i)) - 1)
 NOBDescription(i) = Sheet3.Range(cellrange2).value
 NOBName(i) = Sheet3.Range(cellRange1).value
 
 If NOBDescription(i) <> "" Or NOBName(i) <> "" Then
 If NOBCode(i) = "" Or NOBCode(i) = "(Select)" Then
        errorMsg_NOB1 = errorMsg_NOB1 & "* Please select Business code from the drop down of 44AD section  at Sr. No  " & i & " in Sheet BP" & Chr(13)
        ValidateNOB1 = False
 End If
 End If
 
 If NOBDescription(i) <> "" Or (NOBCode(i) <> "" Or NOBCode(i) <> "(Select)") Then
 If NOBName(i) = "" Then
        errorMsg_NOB1 = errorMsg_NOB1 & "* Please enter Name of Business of 44AD section  at Sr. No  " & i & " in Sheet BP" & Chr(13)
        ValidateNOB1 = False
 End If
 End If
 

Next

If value = 0 Then
    Dim income_BP As Variant
    income_BP = IIf(Sheet3.Range("BP_E1b").value = "", 0, Sheet3.Range("BP_E1b").value) + IIf(Sheet3.Range("BP_E1a").value = "", 0, Sheet3.Range("BP_E1a").value)
    If income_BP > 0 Then
       If Not end_NOBC > 0 Then
        errorMsg_NOB1 = errorMsg_NOB1 & "* Please select Business code from the drop down of 44AD section in Sheet BP" & Chr(13)
        ValidateNOB1 = False
       End If
    End If
End If

'If end_NOBName <> end_NOBC Then
'        errorMsg_NOB1 = errorMsg_NOB1 & "Please fill all the mandatory fields in table under section 44AD."
'        ValidateNOB1 = False
'End If
'
''If end_NOBName <> 0 And end_NOBC <> 0 Then
'If end_NOBDescription > 0 Then
'   If end_NOBName <> end_NOBC Or end_NOBName <> end_NOBDescription Or end_NOBC <> end_NOBDescription Then
'        errorMsg_NOB1 = errorMsg_NOB1 & "Please fill all the mandatory fields in  table under section 44AD."
'        ValidateNOB1 = False
'    End If
'End If
''End If

If errorMsg_NOB1 <> "" Then
ValidateNOB1 = False
End If


End Function

Sub setTblinfo_NOBC1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AD.NOBCode1").count
    Set rangecells = Range("Sec44AD.NOBCode1").Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
 end_NOBC1 = ccount
 End Sub
 
 Sub setTblinfo_NOBName1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AD.NOBName1").count
    Set rangecells = Range("Sec44AD.NOBName1").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_NOBName1 = ccount
End Sub
Sub setTblinfo_NOBDescription1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AD.NOBDescription1").count
    Set rangecells = Range("Sec44AD.NOBDescription1").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_NOBDescription1 = ccount
End Sub

Function ValidateNOB2() As Boolean
On Error Resume Next

ValidateNOB2 = True

Dim cellrange As String
Dim cellRange1 As String
Dim cellrange2 As String
Dim rangecells As Range
Dim rangecells1 As Range
Dim rangecells2 As Range
Dim value As Variant
Dim i As Long
Set rangecells = Range("Sec44AD.NOBCode1").Cells
Set rangecells1 = Range("Sec44AD.NOBName1").Cells
Set rangecells2 = Range("Sec44AD.NOBDescription1").Cells
setTblinfo_NOBC1
setTblinfo_NOBName1
setTblinfo_NOBDescription1

end44ADA = WorksheetFunction.Max(0, end_NOBName1, end_NOBC1, end_NOBDescription1)

ReDim NOBCode1(end44ADA)
ReDim NOBName1(end44ADA)
ReDim NOBDescription1(end44ADA)

For i = 1 To end44ADA

 cellrange = GetMergedAddressCell(rangecells, i)
 cellRange1 = GetMergedAddressCell(rangecells1, i)
 cellrange2 = GetMergedAddressCell(rangecells2, i)
 
  value = value + 1
 NOBCode1(i) = Sheet3.Range(cellrange).value
 NOBCode1(i) = Mid(NOBCode1(i), 1, WorksheetFunction.Search("-", NOBCode1(i)) - 1)
 NOBDescription1(i) = Sheet3.Range(cellrange2).value
 NOBName1(i) = Sheet3.Range(cellRange1).value
 
 If NOBDescription1(i) <> "" Or NOBName1(i) <> "" Then
 If NOBCode1(i) = "" Or NOBCode1(i) = "(Select)" Then
        errorMsg_NOB2 = errorMsg_NOB2 & "* Please select Business code from the drop down of 44ADA section  at Sr. No  " & i & "in Sheet BP " & Chr(13)
        ValidateNOB2 = False
 End If
 End If
 
 If NOBDescription1(i) <> "" Or (NOBCode1(i) <> "" Or NOBCode1(i) <> "(Select)") Then
 If NOBName1(i) = "" Then
        errorMsg_NOB2 = errorMsg_NOB2 & "* Please fill all the mandatory fields in  table  at Sr. No  " & i & "in Sheet BP " & Chr(13)
        ValidateNOB2 = False
 End If
 End If
 
Next

'--------------------Commented by Mallikarjun on 10-12-2023
If value = 0 Then
    Dim income_BP As Variant
    income_BP = IIf(Sheet3.Range("BP_E3_GR").value = "", 0, Sheet3.Range("BP_E3_GR").value)
    If income_BP > 0 Then
        errorMsg_NOB2 = errorMsg_NOB2 & "* Please select Business code from the drop down of 44ADA section in Sheet BP"
        ValidateNOB2 = False
    End If
End If
'--------------------Commented by Mallikarjun on 10-12-2023







'If end_NOBC1 <> end_NOBName1 Then
'        errorMsg_NOB2 = errorMsg_NOB2 & "Please fill all the mandatory fields in table under section 44ADA."
'        ValidateNOB2 = False
'End If
'
'If end_NOBDescription1 > 0 Then
'If end_NOBC1 <> end_NOBDescription1 Or end_NOBName1 <> end_NOBDescription1 Then
'        errorMsg_NOB2 = errorMsg_NOB2 & "Please fill all the mandatory fields in  table under section 44ADA."
'        ValidateNOB2 = False
'End If
'End If

If errorMsg_NOB2 <> "" Then
ValidateNOB2 = False
End If

End Function
Sub setTblinfo_NOBC2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AD.NOBCode2").count
    Set rangecells = Range("Sec44AD.NOBCode2").Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
 end_NOBC2 = ccount
 End Sub
 
 Sub setTblinfo_NOBName2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AD.NOBName2").count
    Set rangecells = Range("Sec44AD.NOBName2").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_NOBName2 = ccount
End Sub
Sub setTblinfo_NOBDescription2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AD.NOBDescription2").count
    Set rangecells = Range("Sec44AD.NOBDescription2").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_NOBDescription2 = ccount
End Sub

Function ValidateNOB3() As Boolean
On Error Resume Next

ValidateNOB3 = True

Dim cellrange As String
Dim cellRange1 As String
Dim cellrange2 As String
Dim rangecells As Range
Dim rangecells1 As Range
Dim rangecells2 As Range
Dim value As Variant
Dim i As Long
Set rangecells = Range("Sec44AD.NOBCode2").Cells
Set rangecells1 = Range("Sec44AD.NOBName2").Cells
Set rangecells2 = Range("Sec44AD.NOBDescription2").Cells
setTblinfo_NOBC2
setTblinfo_NOBName2
setTblinfo_NOBDescription2

end44AE = WorksheetFunction.Max(0, end_NOBName2, end_NOBC2, end_NOBDescription2)


ReDim NOBCode2(end44AE)
ReDim NOBName2(end44AE)
ReDim NOBDescription2(end44AE)

For i = 1 To end44AE

 cellrange = GetMergedAddressCell(rangecells, i)
 cellRange1 = GetMergedAddressCell(rangecells1, i)
 cellrange2 = GetMergedAddressCell(rangecells2, i)
 
  value = value + 1
 NOBCode2(i) = Sheet3.Range(cellrange).value
 NOBCode2(i) = Mid(NOBCode2(i), 1, WorksheetFunction.Search("-", NOBCode2(i)) - 1)
 NOBDescription2(i) = Sheet3.Range(cellrange2).value
 NOBName2(i) = Sheet3.Range(cellRange1).value
 
 If NOBDescription2(i) <> "" Or NOBName2(i) <> "" Then
 If NOBCode2(i) = "" Or NOBCode2(i) = "(Select)" Then
        errorMsg_NOB3 = errorMsg_NOB3 & "* Please select Business code from the drop down of 44AE section at Sr. No  " & i & " in Sheet BP" & Chr(13)
        ValidateNOB3 = False
 End If
 End If
 
 If NOBDescription2(i) <> "" Or (NOBCode2(i) <> "" Or NOBCode2(i) <> "(Select)") Then
 If NOBName2(i) = "" Then
        errorMsg_NOB3 = errorMsg_NOB3 & "* Please enter Name of Business of 44AE section  at Sr. No  " & i & " in Sheet BP" & Chr(13)
        ValidateNOB3 = False
 End If
 End If
 

Next

If value = 0 Then
    Dim income_BP As Variant
    income_BP = IIf(Sheet3.Range("BP_E3").value = "", 0, Sheet3.Range("BP_E3").value)
    If income_BP > 0 Then
        errorMsg_NOB3 = errorMsg_NOB3 & "* Please select at least 1 drop down from Business Code of 44AE section in Sheet BP"
        ValidateNOB3 = False
    End If
End If

'If end_NOBC2 <> end_NOBName2 Then
'        errorMsg_NOB3 = errorMsg_NOB3 & "Please fill all the mandatory fields in table under section 44AE."
'        ValidateNOB3 = False
'End If
'
'If end_NOBDescription2 > 0 Then
'If end_NOBC2 <> end_NOBDescription2 Or end_NOBName2 <> end_NOBDescription2 Then
'        errorMsg_NOB3 = errorMsg_NOB3 & "Please fill all the mandatory fields in  table under section 44AE."
'        ValidateNOB3 = False
'End If
'End If


If errorMsg_NOB3 <> "" Then
ValidateNOB3 = False
End If

End Function

'Sub setTblinfo_NOBC3()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("Sec44AD.NOBCode3").Count
'    Set rangecells = Range("Sec44AD.NOBCode3").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
'                ccount = ccount + 1
'            End If
'    Next
' end_NOBC3 = ccount
' End Sub
 
' Sub setTblinfo_NOBName3()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("Sec44AD.NOBName3").Count
'    Set rangecells = Range("Sec44AD.NOBName3").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_NOBName3 = ccount
'End Sub
'Sub setTblinfo_NOBDescription3()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("Sec44AD.NOBDescription3").Count
'    Set rangecells = Range("Sec44AD.NOBDescription3").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_NOBDescription3 = ccount
'End Sub

'Function ValidateNOB4() As Boolean
'On Error Resume Next
'
'ValidateNOB4 = True
'
'Dim cellrange As String
'Dim cellrange1 As String
'Dim cellrange2 As String
'Dim cellrange3 As String
'Dim rangecells As Range
'Dim rangecells1 As Range
'Dim rangecells2 As Range
'Dim rangecells3 As Range
'Dim value As Variant
'Dim i As Long
'Set rangecells = Range("Sec44AD.NOBCode3").Cells
'Set rangecells1 = Range("Sec44AD.NOBName3").Cells
'Set rangecells2 = Range("BP.GSTRAmount").Cells
'Set rangecells3 = Range("BP.GSTRno").Cells
'setTblinfo_NOBC3
'setTblinfo_NOBName3
'setTblinfo_GSTAMOUNT
'setTblinfo_GSTIN
'
'ReDim NOBCode3(end_NOBC3)
'ReDim NOBName3(end_NOBC3)
'ReDim NOBDescription3(end_NOBC3)
'ReDim NOBGSTIN(end_NOBC3)
'
'For i = 1 To end_NOBC3
'
' cellrange = GetMergedAddressCell(rangecells, i)
' cellrange1 = GetMergedAddressCell(rangecells1, i)
' cellrange2 = GetMergedAddressCell(rangecells2, i)
' cellrange3 = GetMergedAddressCell(rangecells3, i)
'
'value = value + 1
' NOBCode3(i) = Sheet3.Range(cellrange).value
' NOBCode3(i) = Mid(NOBCode3(i), 1, WorksheetFunction.Search("-", NOBCode3(i)) - 1)
' NOBDescription3(i) = Sheet3.Range(cellrange2).value
' NOBName3(i) = Sheet3.Range(cellrange1).value
' NOBGSTIN(i) = Sheet3.Range(cellrange3).value
'
'        If Not chkCompulsory(NOBName3(i)) Then
'                errorMsg_NOB4 = errorMsg_NOB4 + "Please enter Name of Business at Sr. No  " & i & " " & Chr(13)
'                ValidateNOB4 = False
'                Exit Function
'        End If
'
'        If Not chkCompulsory(NOBDescription3(i)) Then
'                errorMsg_NOB4 = errorMsg_NOB4 + "Please enter the annual Value of outward supplies reported for GST at Sr. No  " & i & " " & Chr(13)
'                ValidateNOB4 = False
'                Exit Function
'        End If
'
'        If Not chkCompulsory(NOBGSTIN(i)) Then
'                errorMsg_NOB4 = errorMsg_NOB4 + "Please enter GSTIN No at Sr. No  " & i & " " & Chr(13)
'                ValidateNOB4 = False
'                Exit Function
'        End If
'
'        If isdropdownblank(NOBCode3(i)) Then
'                errorMsg_NOB4 = errorMsg_NOB4 + "Please select  Business code from the drop down at Sr. No  " & i & " " & Chr(13)
'                ValidateNOB4 = False
'                Exit Function
'        End If
'
'
'
'
'Next
'
'If end_NOBC3 <> end_NOBName3 Or end_NOBC3 <> end_GSTIN Or end_NOBName3 <> end_GSTIN Or end_GSTIN <> end_GSTAMOUNT Or end_NOBC3 <> end_GSTAMOUNT Or end_NOBName3 <> end_GSTAMOUNT Then
'        errorMsg_NOB4 = errorMsg_NOB4 & "Please fill all the mandatory fields in table under INFORMATION REGARDING TURNOVER/GROSS RECEIPT REPORTED FOR GST"
'        ValidateNOB4 = False
'End If
'
'If errorMsg_NOB4 <> "" Then
'ValidateNOB4 = False
'End If
'
'
'End Function

Sub setTableInfo_RegNo()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet3.Range("Sec44AE.RegNo").count
 Set rangecells = Sheet3.Range("Sec44AE.RegNo").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end44AE_RegNo = ccount
End Sub
Sub setTblinfo_OLH()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AE.OLH").count
    Set rangecells = Range("Sec44AE.OLH").Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
 end44AE_OLH = ccount
 End Sub
Sub setTableInfo_TC()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet3.Range("Sec44AE.TonnageCpct").count
 Set rangecells = Sheet3.Range("Sec44AE.TonnageCpct").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end44AE_TC = ccount
End Sub
Sub setTableInfo_Month()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet3.Range("Sheet44AE.NoOfMonths").count
 Set rangecells = Sheet3.Range("Sheet44AE.NoOfMonths").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end44AE_Months = ccount
End Sub
Sub setTableInfo_PIncome()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet3.Range("Sheet44AE.PresumptiveIncome").count
 Set rangecells = Sheet3.Range("Sheet44AE.PresumptiveIncome").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end44AE_PIncome = ccount
End Sub


Function ValidateSec44AE() As Boolean
On Error Resume Next
Dim cntr, cntr2, value As Variant
Dim i As Variant
value = 0
ValidateSec44AE = True
setTableInfo_RegNo
ReDim SEC44AERegNo(end44AE_RegNo)
ReDim SEC44AEOLH(end44AE_RegNo)
ReDim SEC44AETC(end44AE_RegNo)
ReDim SEC44AEMonth(end44AE_RegNo)
ReDim SEC44AEPIncome(end44AE_RegNo)
ReDim Greater12_BP(end44AE_RegNo)
ReDim Lesser12_BP(end44AE_RegNo)


setTblinfo_OLH
setTableInfo_TC
setTableInfo_Month
setTableInfo_PIncome

end44AE_RegNo = WorksheetFunction.Max(0, end44AE_RegNo, end44AE_OLH, end44AE_TC, end44AE_Months, end44AE_PIncome)


For cntr = 1 To end44AE_RegNo

 If Not CheckEmptyDDList(Range("Sec44AE.RegNo").item(cntr).value) Then
 value = value + 1
 SEC44AERegNo(cntr) = Range("Sec44AE.RegNo").item(cntr).value
 'SEC44AEOLH(cntr2) = Mid(NOBCode3(cntr2), 1, WorksheetFunction.Search("-", NOBCode3(cntr2)) - 1)
 SEC44AEOLH(cntr) = Range("Sec44AE.OLH").item(cntr).value
 If SEC44AEOLH(cntr) = "Owned" Then
    SEC44AEOLH(cntr) = "OWN"
ElseIf SEC44AEOLH(cntr) = "Leased" Then
    SEC44AEOLH(cntr) = "LEASE"
ElseIf SEC44AEOLH(cntr) = "Hired" Then
    SEC44AEOLH(cntr) = "HIRED"
End If
 SEC44AETC(cntr) = Range("Sec44AE.TonnageCpct").item(cntr).value
 SEC44AEMonth(cntr) = Range("Sheet44AE.NoOfMonths").item(cntr).value
 SEC44AEPIncome(cntr) = Range("Sheet44AE.PresumptiveIncome").item(cntr).value
 
 Greater12_BP(cntr) = Range("greater12").item(cntr).value
 Lesser12_BP(cntr) = Range("lesser12").item(cntr).value
 Else

 End If
Next

For i = 1 To end44AE_RegNo

        If Not chkCompulsory(SEC44AERegNo(i)) Then
                errorMsg_NOB5 = errorMsg_NOB5 + "* Please enter the Registration no of goods carriage in Sheet BP at Sr. No  " & i & " " & Chr(13)
                ValidateSec44AE = False
                Exit Function
        End If
        
        'Ayush_24/03/2026_SIT-111668
        If SEC44AERegNo(i) <> "" Then
            If Not checkfieldspecialcharacter80DD_80U(SEC44AERegNo(i)) Then
                 errorMsg_NOB5 = errorMsg_NOB5 + "* Registration No. of goods carriage cannot contain special characters in Schedule BP" & Chr(13)
                 ValidateSec44AE = False
                 Exit Function
            End If
        End If
        
        If Not chkCompulsory(SEC44AETC(i)) Then
                errorMsg_NOB5 = errorMsg_NOB5 + "* Please enter the tonnage capacity of goods carriage(in MT) in Sheet BP at Sr. No  " & i & " " & Chr(13)
                ValidateSec44AE = False
                Exit Function
        End If
        
        If SEC44AETC(i) <> "" Then
        If Len(SEC44AETC(i)) > 3 Then
                errorMsg_NOB5 = errorMsg_NOB5 + "* tonnage capacity of goods carriage(in MT) in Sheet BP cannot be more than 3 digits at Sr. No  " & i & " " & Chr(13)
                ValidateSec44AE = False
                Exit Function
        End If
        End If
        
        If Not chkCompulsory(SEC44AEMonth(i)) Then
                errorMsg_NOB5 = errorMsg_NOB5 + "* Please enter Period of Holding in Sheet BP at Sr. No  " & i & " " & Chr(13)
                ValidateSec44AE = False
                Exit Function
        End If
        
        If Not chkCompulsory(SEC44AEPIncome(i)) Then
                errorMsg_NOB5 = errorMsg_NOB5 + "* Please enter Presumptive income for the goods carriage in Sheet BP at Sr. No  " & i & " " & Chr(13)
                ValidateSec44AE = False
                Exit Function
        End If
        
        If isdropdownblank(SEC44AEOLH(i)) Then
                errorMsg_NOB5 = errorMsg_NOB5 + "* Please Select an option from the drop down whether goods carriage is owned/leased/hired in Sheet BP at Sr. No  " & i & " " & Chr(13)
                ValidateSec44AE = False
                Exit Function
        End If
        
        If SEC44AETC(i) > 12 Then
        If SEC44AEPIncome(i) < Greater12_BP(i) Then
            errorMsg_NOB5 = errorMsg_NOB5 + "*Income cannot be lower than Rs. 1000 per metric ton per month in case of Heavy Goods vehicle in Sheet BP at Sr. No  " & i & " " & Chr(13)
            ValidateSec44AE = False
                Exit Function
      End If
      End If
      
      If SEC44AETC(i) < 12 Then
        If SEC44AEPIncome(i) < Lesser12_BP(i) Then
            errorMsg_NOB5 = errorMsg_NOB5 + "* Income cannot be less than Rs. 7500 per month in case of other than Heavy Goods Vehicle in Sheet BP at Sr. No  " & i & " " & Chr(13)
            ValidateSec44AE = False
                Exit Function
      End If
      End If

Next


If end44AE_PIncome <> end44AE_RegNo Or end44AE_PIncome <> end44AE_OLH Or end44AE_PIncome <> end44AE_TC Or end44AE_PIncome <> end44AE_Months Then
        errorMsg_NOB5 = errorMsg_NOB5 & "* Please enter all mandatory Fields of 44AE Table in Sheet BP Table"
        ValidateSec44AE = False
End If


'If 120 > Sheet3.Range("Sheet44AE.NoOfMonthsNew").value Then
'errorMsg_NOB5 = errorMsg_NOB5 & "*No of months cannot exceed 120"
'ValidateSec44AE = False
'End If

If errorMsg_NOB5 <> "" Then
ValidateSec44AE = False
End If


End Function

'AY 2024-25 CHANGE---------------------------------

Function ValidateSecGROSSTURNOVER() As Boolean

ValidateSecGROSSTURNOVER = True
Dim fivepersent As Variant
 fivepersent = GrossTurnover_Receipts * 5 / 100
 
 'If GrossTurnover_a_b >= fivepersent Then
'GrossTurnoverE1bandE1C
 If GrossTurnoverE1bandE1C > fivepersent Then ' Changed as per Aparna Confirmation by Bindu on 29th may 25
 

     If GrossTurnover_Receipts >= 20000001 Then
          msgError = msgError & "*  If Gross Receipts are more than Rs.2 Crore and cash receipts are more than 5% of total receipts, it is mandatory to have a tax audit under 44AB. Please use the regular ITR 3/5 form."
     'msgBx14NOBBP = msgBx14NOBBP + "*  If Gross Receipts are more than Rs.2 Crore and cash receipts are more than 5% of total receipts, it is mandatory to have a tax audit under 44AB. Please use the regular ITR 3/5 form."
     ValidateSecGROSSTURNOVER = False
     Exit Function
     End If
'AY_2025_26 'Malli '
'ElseIf GrossTurnover_a_b <= fivepersent Then
'ElseIf GrossTurnoverE1bandE1C <= fivepersent Then ' Commented by Bindu post confirmation with BA team (Aayushi & Aparna)n on 29th may 25
'---------------
  If GrossTurnover_Receipts >= 30000001 Then
         msgError = msgError & "*  If Gross Receipts are more than Rs.3 Crore it is mandatory to have a tax audit under 44AB. Please use the regular ITR 3/5 form."
     ValidateSecGROSSTURNOVER = False
     Exit Function
     
     End If
End If

End Function

'AY 2024-25 Change ------

Function ValidateSecGROSS_Receipts() As Boolean
'GrossTurnover44ADA = Range("BP_E3_GR").value
 
'GrossReceipts_E3_a = Range("BP_E3_a").value  'BP_C10 AY 2024-25
'GrossReceipts_E3_b = Range("BP_E3_b").value  'BP_C10 AY 2024-25
'GrossReceipts_E3_c = Range("BP_E3_c").value  'BP_C10 AY 2024-25
 
ValidateSecGROSS_Receipts = True
Dim fivepersent As Variant
fivepersent = GrossTurnover44ADA * 5 / 100
'If GrossReceipts_E3_b >= fivepersent Then
 
'GrossReceipts_E3_b_c
If GrossReceipts_E3_b_c > fivepersent Then ' Post confirmation with Aparna updated by Bindu on 29th May 25
 
 
     If GrossTurnover44ADA >= 5000001 Then
          msgError = msgError & "*  If Gross Receipts are more than Rs.5000000 and cash receipts are more than 5% of total receipts, it is mandatory to have a tax audit under 44AB. Please use the regular ITR 3/5 form." & Chr(13)
     
     ValidateSecGROSS_Receipts = False
     Exit Function
     End If
 
 '
'ElseIf GrossReceipts_E3_b <= fivepersent Then
ElseIf GrossReceipts_E3_b_c <= fivepersent Then '  Post confirmation with Aparna updated by Bindu on 29th May 25
 
  If GrossTurnover44ADA >= 7500001 Then
         msgError = msgError & "*  If Gross Receipts are more than Rs.7500000, it is mandatory to have a tax audit under 44AB. Please use the regular ITR 3/5 form" & Chr(13)
     ValidateSecGROSS_Receipts = False
     Exit Function
     End If
End If
 
End Function



