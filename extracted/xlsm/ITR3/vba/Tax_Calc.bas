Attribute VB_Name = "Tax_Calc"
Option Explicit
Public rebate_new As Variant 'For Rebate 87A Changes by sai 06/01/2025
Sub calcTaxTest()
UserForm1.Show
End Sub
'Sub calculateTax()
 Sub calculateTax(Optional rebate_flag As Boolean) 'For Rebate 87A Changes by sai 01/01/2025

Dim osDI As Double
Dim osIE As Double
Dim osBFLA As Double
Dim SW1 As Double
Dim os1a As Double
Dim Pre1ai As Double
Dim Post1ai As Double
Dim Div5 As Double
Dim IEOS As Double
Dim sheet As String
sheet = ActiveSheet.name
Sheet14.Unprotect getmsgstate
'Sheet14.Range("os.amountinterestexpenditure").Value = 0
Sheet14.Protect getmsgstate

Dim flag, cell
flag = False
For Each cell In Sheet2.Range("NOB.Code").Cells
    If Not isdropdownblank(cell.value) Then
        flag = True
        Exit For
    End If
Next cell

If Not flag Then
For Each cell In Sheet4.Range("Sec44AD.NOBCode").Cells
    If Not isdropdownblank(cell.value) Then
        flag = True
        Exit For
    End If
Next cell

End If

If Not flag Then
For Each cell In Sheet4.Range("Sec44AD.NOBCode1").Cells
    If Not isdropdownblank(cell.value) Then
        flag = True
        Exit For
    End If
Next cell

End If


If Not flag Then
For Each cell In Sheet4.Range("Sec44AD.NOBCode2").Cells
    If Not isdropdownblank(cell.value) Then
        flag = True
        Exit For
    End If
Next cell

End If

If Not flag Then
    MsgBox "At least one Business Code is mandatory either in Schedule Nature of Business or at Point-61 or 62 or 63 in Schedule Profit & Loss A/c"
    CloseMsg
End If

osBFLA = Sheet16.Range("othsrcincl.IncOfCurYrAfterSetOffBFLosses5").value
IEOS = Sheet14.Range("os.expenditureClaimed").value
SW1 = Sheet7.Range("Sheet8b.AggregateIncomeNew").value

If (Sheet14.Range("ITER").value <> 2) Then
If (Sheet14.Range("osBFLA") <> "") Then
If (osBFLA <> Sheet14.Range("osBFLA").value) Then
Sheet14.Unprotect getmsgstate
Sheet14.Range("os.amountinterestexpenditure") = ""
Sheet14.Unprotect getmsgstate
Sheet14.Range("osBFLA") = ""
Sheet14.Unprotect getmsgstate
Sheet14.Range("OSSW1") = ""
Sheet14.Unprotect getmsgstate
Sheet14.Range("IEOS") = ""
Sheet14.Protect getmsgstate
End If
End If

If (Sheet14.Range("IEOS") <> "") Then
If (IEOS <> Sheet14.Range("IEOS").value) Then
Sheet14.Unprotect getmsgstate
Sheet14.Range("os.amountinterestexpenditure") = ""
Sheet14.Unprotect getmsgstate
Sheet14.Range("osBFLA") = ""
Sheet14.Unprotect getmsgstate
Sheet14.Range("IEOS") = ""
Sheet14.Unprotect getmsgstate
Sheet14.Range("OSSW1") = ""
Sheet14.Protect getmsgstate
End If
End If

If (Sheet14.Range("OSSW1") <> "") Then
If (SW1 <> Sheet14.Range("OSSW1").value) Then
Sheet14.Unprotect getmsgstate
Sheet14.Range("os.amountinterestexpenditure") = ""
Sheet14.Unprotect getmsgstate
Sheet14.Range("osBFLA") = ""
Sheet14.Unprotect getmsgstate
Sheet14.Range("OSSW1") = ""
Sheet14.Unprotect getmsgstate
Sheet14.Range("IEOS") = ""
Sheet14.Protect getmsgstate
End If
End If

End If

If (Sheet14.Range("os.amountinterestexpenditure") = "") Then
Sheet14.Unprotect getmsgstate
Pre1ai = Sheet14.Range("os.DividendGrossai").value



'WcalculateTax
WcalculateTax (rebate_flag) 'For Rebate87A Changes by sai 06/01/2025
Sheet14.Unprotect getmsgstate
Application.EnableEvents = False

osIE = Sheet14.Range("os.expenditureClaimed").value
Sheet14.Unprotect getmsgstate
osBFLA = Sheet16.Range("othsrcincl.IncOfCurYrAfterSetOffBFLosses5").value
Sheet14.Unprotect getmsgstate
IEOS = Sheet14.Range("IEOS").value
Sheet14.Unprotect getmsgstate
SW1 = Sheet7.Range("Sheet8b.AggregateIncomeNew").value
Sheet14.Unprotect getmsgstate

'os1a = Sheet14.Range("os.DividendGross").value
'For AY 2025-26 by sai on 02/03/2025
os1a = Sheet14.Range("os.DividendGross_1ai_1aii").value

Application.EnableEvents = True

osDI = IIf(os1a < osBFLA, os1a, osBFLA)

osDI = IIf(osDI < SW1, osDI, SW1)

osDI = Round(0.2 * osDI)


osIE = IIf(osIE < osDI, osIE, osDI)

Sheet14.Unprotect getmsgstate
Application.EnableEvents = False
Sheet14.Range("os.amountinterestexpenditure").value = osIE
Sheet14.Unprotect getmsgstate
Sheet14.Range("osBFLA").value = osBFLA
Sheet14.Unprotect getmsgstate
Sheet14.Range("IEOS").value = IEOS
Sheet14.Unprotect getmsgstate
Sheet14.Range("ITER").value = 2
Sheet14.Unprotect getmsgstate
Sheet14.Range("OSSW1").value = SW1

Application.EnableEvents = True
Sheet14.Unprotect getmsgstate
Post1ai = Sheet14.Range("os.DividendGrossai").value
If (Pre1ai <> Post1ai) Then
Sheet14.Unprotect getmsgstate
Div5 = Sheet14.Range("BBDA.Up16Of3To31Of3").value
Application.EnableEvents = False
Sheet14.Unprotect getmsgstate
Sheet14.Range("BBDA.Up16Of3To31Of3").value = Div5 + Pre1ai - Post1ai
Application.EnableEvents = True
End If
Sheet16.Protect getmsgstate
End If



'WcalculateTax
WcalculateTax (rebate_flag) 'For Rebate87A Changes by sai 06/01/2025
Sheet14.Unprotect getmsgstate
osIE = Sheet14.Range("os.expenditureClaimed").value
osBFLA = Sheet16.Range("othsrcincl.IncOfCurYrAfterSetOffBFLosses5").value
Sheet14.Unprotect getmsgstate
IEOS = Sheet14.Range("IEOS").value
SW1 = Sheet7.Range("Sheet8b.AggregateIncomeNew").value
Sheet14.Unprotect getmsgstate
Application.EnableEvents = False
Sheet14.Range("osBFLA").value = osBFLA
Sheet14.Unprotect getmsgstate
Sheet14.Range("IEOS").value = Sheet14.Range("os.expenditureClaimed").value
Sheet14.Unprotect getmsgstate
Sheet14.Range("OSSW1").value = SW1
Sheet14.Unprotect getmsgstate
Sheet14.Range("ITER").value = 4

'Added by sai on 30/10/2024

If UCase(Sheet1.Range("sheet1.ReturnFileSec").value) = UCase("139(8A)") Then
 
    If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").value) = UCase("No") Then
 
        If Sheet202.Range("U_AmtPayable") <= 0 Then
           MsgBox "If A5 in Part A gen 139(8A) is ""No"",then updated return can't be filed if there is no payment u/s 140B", vbCritical, "Error"
           CalculateTaxCheckFlag = False
           Exit Sub
        End If
    End If
' End If

'Added by sai on 30/10/2024
 
    If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").value) = UCase("No") Then
 
        If Sheet202.Range("U_AmtRefundable") > 0 Then
           MsgBox "If A5 in Part A gen 139(8A) is ""No"",then updated return can't be filed with Refund", , "Error"
           CalculateTaxCheckFlag = False
           Exit Sub
        End If
    End If
 
'Added by sai on 30/10/2024
    If Sheet202.Range("U_NetPayable") <= 0 Then
        MsgBox "Since you are filing Updated return, you are expected to have ""Net amount payable"" more than 0 to be able to pay tax u/s 140B.", , "Error"
'       CalculateTaxCheckFlag = False
'       Exit Sub
    End If
 
 
 'Added by sai on 13/11/2024
    If Sheet202.Range("U_TaxUS140B").value = 0 Then
        MsgBox "* ""Tax Paid u/s 140B at sl.no.12 should be more than ""0"" to file ITR u/s 139(8A)"" " & Chr(13)
    End If
'==========================================================

'Added by sai on 30/10/2024
 
    If Sheet202.Range("U_TaxUS140B") > Sheet202.Range("U_NetPayable").value Then
       MsgBox "Tax paid u/s 140B is more than ""sl.no.12.Net amount payable"". Hence your updated return may be invalid as it appears to result in refund as provided in clause (c) to first proviso of 139(8A).", vbInformation, "Error"
'      CalculateTaxCheckFlag = False
'       Exit Sub
    End If
End If
'======================================================================================================================

'IPIP-71829 Warning message for 87A incase of new tax regime

        If (SheetALL.Range("bacValue").value) = 1 Then
               If Sheet7.Range("Sheet9.RebateUs88E").value > 0 Then
                      MsgBox """Rebate under Section 87A will be available only on the income taxable at normal rates i.e. after exclusion of the income taxable at special rates under New Tax Regime.""", vbInformation, "Warning"
               End If
        End If

Application.EnableEvents = True
Sheet14.Protect getmsgstate
'If Not Validate_OSIE Then
'        Sheet15.Activate
'        MsgBox "Opps 2"
'        EndProcessing
'    End If
Sheets(sheet).Activate

End Sub
Sub calculateTaxold1()
Dim osDI As Double
Dim osIE As Double
Dim osBFLA As Double
Dim SW1 As Double
Dim os1a As Double
Dim Pre1ai As Double
Dim Post1ai As Double
Dim Div5 As Double
Dim IEOS As Double
Dim sheet As String
sheet = ActiveSheet.name
Sheet14.Unprotect getmsgstate
'Sheet14.Range("os.amountinterestexpenditure").Value = 0
Sheet14.Protect getmsgstate
osBFLA = Sheet16.Range("othsrcincl.IncOfCurYrAfterSetOffBFLosses5").value
IEOS = Sheet14.Range("os.expenditureClaimed").value
SW1 = Sheet7.Range("Sheet8b.AggregateIncome").value

If (Sheet14.Range("ITER").value <> 2) Then
If (Sheet14.Range("osBFLA") <> "") Then
If (osBFLA <> Sheet14.Range("osBFLA").value) Then
Sheet14.Unprotect getmsgstate
Sheet14.Range("os.amountinterestexpenditure") = ""
Sheet14.Range("osBFLA") = ""
Sheet14.Range("OSSW1") = ""
Sheet14.Range("IEOS") = ""
Sheet14.Protect getmsgstate
End If
End If

If (Sheet14.Range("IEOS") <> "") Then
If (IEOS <> Sheet14.Range("IEOS").value) Then
Sheet14.Unprotect getmsgstate
Sheet14.Range("os.amountinterestexpenditure") = ""
Sheet14.Range("osBFLA") = ""
Sheet14.Range("IEOS") = ""
Sheet14.Range("OSSW1") = ""
Sheet14.Protect getmsgstate
End If
End If

If (Sheet14.Range("OSSW1") <> "") Then
If (SW1 <> Sheet14.Range("OSSW1").value) Then
Sheet14.Unprotect getmsgstate
Sheet14.Range("os.amountinterestexpenditure") = ""
Sheet14.Range("osBFLA") = ""
Sheet14.Range("OSSW1") = ""
Sheet14.Range("IEOS") = ""
Sheet14.Protect getmsgstate
End If
End If

End If

If (Sheet14.Range("os.amountinterestexpenditure") = "") Then
Pre1ai = Sheet14.Range("os.DividendGrossai").value



WcalculateTax
Application.EnableEvents = False
osIE = Sheet14.Range("os.expenditureClaimed").value

osBFLA = Sheet16.Range("othsrcincl.IncOfCurYrAfterSetOffBFLosses5").value
IEOS = Sheet14.Range("IEOS").value
SW1 = Sheet7.Range("Sheet8b.AggregateIncome").value
os1a = Sheet14.Range("os.DividendGross").value
Application.EnableEvents = True

osDI = IIf(os1a < osBFLA, os1a, osBFLA)



osDI = IIf(osDI < SW1, osDI, SW1)

osDI = Round(0.2 * osDI)


osIE = IIf(osIE < osDI, osIE, osDI)

Sheet14.Unprotect getmsgstate
Application.EnableEvents = False
Sheet14.Range("os.amountinterestexpenditure").value = osIE
Sheet14.Range("osBFLA").value = osBFLA
Sheet14.Range("IEOS").value = IEOS
Sheet14.Range("ITER").value = 2
Sheet14.Range("OSSW1").value = SW1
Application.EnableEvents = True
Post1ai = Sheet14.Range("os.DividendGrossai").value
If (Pre1ai <> Post1ai) Then
Div5 = Sheet14.Range("BBDA.Up16Of3To31Of3").value
Application.EnableEvents = False
Sheet14.Range("BBDA.Up16Of3To31Of3").value = Div5 + Pre1ai - Post1ai
Application.EnableEvents = True
End If
Sheet16.Protect getmsgstate
End If



WcalculateTax
osIE = Sheet14.Range("os.expenditureClaimed").value
osBFLA = Sheet16.Range("othsrcincl.IncOfCurYrAfterSetOffBFLosses5").value
IEOS = Sheet14.Range("IEOS").value
SW1 = Sheet7.Range("Sheet8b.AggregateIncome").value
Sheet14.Unprotect getmsgstate
Application.EnableEvents = False
Sheet14.Range("osBFLA").value = osBFLA
Sheet14.Range("IEOS").value = Sheet14.Range("os.expenditureClaimed").value
Sheet14.Range("OSSW1").value = SW1
Sheet14.Range("ITER").value = 4
Application.EnableEvents = True
Sheet14.Protect getmsgstate
'If Not Validate_OSIE Then
'        Sheet15.Activate
'        MsgBox "Opps 2"
'        EndProcessing
'    End If
Sheets(sheet).Activate

End Sub
'Sub WcalculateTax()
Sub WcalculateTax(Optional rebate_flag As Boolean) ' For Rebate 87A Changes by sai 06/01/2025

'On Error GoTo endline
'UserForm1.Show vbModeless
'validateAllSheets

rebate_new = Sheet7.Range("Sheet9.RebateUs88E").value 'For Rebate 87A Changes by sai 06/01/2025

ProgressFrameCaption = "Calculating Tax"
mainProcCaption = "Validating All Schedules"
'ValidatAllSchedule
ValidatAllScheduleTaxCalculate
'CG_Calc.doSetoff
'CYLACalculations.CYLABFLASetOff


 CalculateTaxCheckFlag = True 'For Rebate 87A Changes by sai 06/01/2025


 Sheet7.Activate

'Application.EnableEvents = False

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet7.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Sheet33.EnableCalculation = True
Application.EnableEvents = False

'Range("Sheet9.TaxAtNormalRatesOnAggrInc").value = 0
calculate234Amonths (Sheet52.Range("sheet9.Date").value)
calculate234Bmonths (Sheet52.Range("sheet9.Date").value)

Sheet7.Range("Sheet9.TaxAtNormalRatesOnAggrInc").value = Sheet33.Range("taxAggrIncm").value
Sheet7.Range("sheet9.TaxAtSpecialRates").value = Sheet21.Range("SI.TotSplRateIncTax").value

'SIT-121080 Updated by Sai as per Bindu 02/06/2026
'Sheet7.Range("sheet9.RebateOnAgriInc").value = Sheet33.Range("RebateOnAgrIncm").value
 Sheet7.Range("sheet9.RebateOnAgriInc").value = Application.WorksheetFunction.Min(Sheet7.Range("Sheet9.TaxAtNormalRatesOnAggrInc").value, Sheet33.Range("RebateOnAgrIncm").value)
 
'---------------
'To populate Advance and Self Assessment Tax
'Sheet7.Range("Sheet9.AdvanceTax").value = Sheet25.Range("IT.AT").value
'Sheet7.Range("Sheet9.SelfAssessmentTax").value = Sheet25.Range("IT.SAT").value

'Sheet7.Range("Sheet9.SurchargeOnTaxPayable").value = Sheet33.Range("SurchargeTTI").value
Sheet7.Range("Sheet9.Surcharge_i").value = Sheet33.Range("Surcharge_i").value
Sheet7.Range("Sheet9.Surcharge_ii").value = Sheet33.Range("Surcharge_ii").value
Sheet7.Range("Sheet9.EducationCess").value = Round((Range("Sheet9.BalTaxPayable").value + Range("Sheet9.SurchargeOnTaxPayable").value) * 0.04, 0)



    'Sheet7.Range("Sheet9.IntrstPayUs234A").value = Sheet33.Range("section234A").value
    'Sheet7.Range("Sheet9.IntrstPayUs234B").value = Sheet33.Range("interest234B").value
    

'For Rebate 87A Changes by sai on 06/01/2025============================================================

Dim SIincomediff As Variant
SIincomediff = Sheet21.Range("SI.TotSplRateIncome").value

If (SheetALL.Range("bacValue").value) = 1 Then

Application.EnableEvents = False
  If SIincomediff = 0 Then
            Sheet7.Unprotect Password:=getmsgstate
            Sheet7.Range("Sheet9.RebateUs88E").Locked = True
            Application.EnableEvents = False
            Sheet7.Range("Sheet9.RebateUs88E").formula = "=Rebate87Aformula_new"
            Sheet7.Range("Sheet9.RebateUs88E").Interior.ColorIndex = 2
            Sheet7.Protect Password:=getmsgstate
  ElseIf SIincomediff > 0 Then
         If Sheet7.Range("TotalIncome.INCD").value <= 1750000 Then
            Sheet7.Unprotect Password:=getmsgstate
            Sheet7.Range("Sheet9.RebateUs88E").Locked = False
            
         If rebate_flag = True Then
            Application.EnableEvents = False
            Sheet7.Range("Sheet9.RebateUs88E").formula = "=Rebate87Aformula_new"
            ElseIf rebate_flag = False Then
            Application.EnableEvents = False
            Sheet7.Range("Sheet9.RebateUs88E").value = rebate_new
            End If
            
            Sheet7.Range("Sheet9.RebateUs88E").Interior.Color = RGB(255, 255, 204)
            Sheet7.Protect Password:=getmsgstate
        ElseIf Sheet7.Range("TotalIncome.INCD").value > 1750000 Then
            Sheet7.Unprotect Password:=getmsgstate
            Sheet7.Range("Sheet9.RebateUs88E").Locked = True
            Application.EnableEvents = False
            Sheet7.Range("Sheet9.RebateUs88E").formula = "=Rebate87Aformula_new"
            Sheet7.Range("Sheet9.RebateUs88E").Interior.ColorIndex = 2
            Sheet7.Protect Password:=getmsgstate
        End If
        
 End If
      ElseIf (SheetALL.Range("bacValue").value) = 2 Then
            Sheet7.Unprotect Password:=getmsgstate
            Sheet7.Range("Sheet9.RebateUs88E").Locked = True
            Application.EnableEvents = False
            Sheet7.Range("Sheet9.RebateUs88E").formula = "=Rebate87Aformula_new"
            Sheet7.Range("Sheet9.RebateUs88E").Interior.ColorIndex = 2
            Sheet7.Protect Password:=getmsgstate
 
 
End If

'===========================================================================================================



    
    ComputeInterest
    'If Sheet7.Range("sheet9.GrossTaxPayable").value <> Sheet7.Range("Sheet9.GrossTaxLiability").value Then
    '    Sheet7.Range("Sheet9.IntrstPayUs234C").value = Sheet33.Range("Total_234Cii").value
    'Else:
    '    Sheet7.Range("Sheet9.IntrstPayUs234C").value = Sheet33.Range("interest234C").value
    'End If
    
    CalculateTaxCheckFlag = True


    If Sheet7.Range("Sheet9.BalTaxPayable1").value > 0 Then
         Sheet7.Shapes("Button 1930").Visible = True
       Else
           Sheet7.Shapes("Button 1930").Visible = False
    End If
    
    
ProgressBarHide
'----------------Lock Password-------------------START---
   Sheet7.Protect Password:=sPassword
   
'----------------Lock Password-------------------END-----
endline:
Application.EnableEvents = True
Sheet33.EnableCalculation = False
Sheet7.Activate
End Sub

Function calculate234Amonths(dt As Variant) As Long
'On Error Resume Next

Dim actualdate As String
Dim Filingtype As String
Dim filingtype1 As String
Dim UpdatedY As Boolean
    UpdatedY = False
    
'    If (((Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))) = "21") And (Mid(Range("U_PreviouslyFiledForThisAY"), 1, 1) = "Y")) Then
    If (((Trim(Mid(Range("sheet1.ReturnFileSec"), 1, 2))) = "21") And (Mid(Range("U_PreviouslyFiledForThisAY"), 1, 1) = "Y")) Then
    UpdatedY = True
    End If
    
    
        currentDate = Sheet31.Range("Cuuren_Date").value
'        Filingtype = Mid(Sheet1.Range("sheet1.ReturnFileSec1").value, 1, 2)
        Filingtype = Mid(Sheet1.Range("sheet1.ReturnFileSec").value, 1, 2)
        'filingtype1 = Mid(Sheet1.Range("sheet1.ReturnType1").value, 1, 3)

        actualdate = currentDate


If Not (checkFirstDateBefore(dt, currentDate)) Then
    currentDate = dt
End If

Dim originalfilingdate As String
originalfilingdate = Sheet1.Range("sheet1.OrigRetFiledDate").value

        If ((Filingtype = "17" Or Filingtype = "19" Or Filingtype = "18" Or UpdatedY)) Then  ' And originalFilingDate = "") Then
            currentDate = originalfilingdate
        End If

Dim DueDate As String
DueDate = getDueDate

'Commented by Riyaz on 26/02/2026
'Dim MonthsAfterDueDate As String
MonthsAfterDueDate = EfilingCommon.calcNoOfMonths(currentDate, DueDate)

'by sadineni on 29/04/2026
'********************************************************
If DueDate = "31/08/2026" Then
    MonthsAfterDueDate = MonthsAfterDueDate
ElseIf DueDate = "31/10/2026" Then
    MonthsAfterDueDate = MonthsAfterDueDate + 2
ElseIf DueDate = "30/11/2026" Then
    MonthsAfterDueDate = MonthsAfterDueDate + 3
End If
'*******************************************************
calculate234Amonths = MonthsAfterDueDate

End Function

Function calculate234Bmonths(dt As Variant) As Long
On Error Resume Next

Dim CurrDate As Variant
Dim TodayDate As Variant



CurrDate = Sheet52.Range("sheet9.Date").value
TodayDate = Sheet33.Range("TodayDate").value

If (checkFirstDateBefore(CurrDate, TodayDate)) Then
    CurrDate = TodayDate
End If

'calculate234Bmonths = calcNoOfMonths(CurrDate, "01/04/2024") 'Chandru
 calculate234Bmonths = calcNoOfMonths(CurrDate, "01/04/2026") 'by Riyaz on 26/02/2026 for AY 2026-27

End Function


Function calcNoOfMonths(currentDate As Variant, startDate As Variant) As Long
On Error Resume Next

Dim currentYear As Variant
Dim startyear As Variant
Dim currentMonth As Variant
Dim startmonth As Variant

currentYear = val(Mid(currentDate, 7, 4))
startyear = val(Mid(startDate, 7, 4))
currentMonth = val(Mid(currentDate, 4, 2))
startmonth = val(Mid(startDate, 4, 2))

 calcNoOfMonths = 0
 
 
 If currentYear = startyear And currentMonth = startmonth And _
        val(Mid(currentDate, 1, 2)) = val(Mid(startDate, 1, 2)) Then
        calcNoOfMonths = 1
ElseIf (checkFirstDateBefore(currentDate, startDate)) Then
        calcNoOfMonths = 0
Else
    If currentYear = (startyear + 1) Then
        If currentMonth < startmonth Then
            calcNoOfMonths = 12 - (startmonth - currentMonth) + 1
        Else
            calcNoOfMonths = 12 + (currentMonth - startmonth) + 1
        End If
   ElseIf (currentYear = startyear) Then
        calcNoOfMonths = currentMonth - startmonth + 1
   ElseIf currentMonth < startmonth Then
        calcNoOfMonths = Round((currentYear - startyear - 1) * 12) + 12 - startmonth + currentMonth + 1
   ElseIf (currentMonth > startmonth) Then
        calcNoOfMonths = Round(((currentYear - startyear) * 12)) + currentMonth - startmonth + 1
    Else
        calcNoOfMonths = Round(((currentYear - startyear) * 12)) + 1
    End If
End If

End Function
Function getDueDate_old()
Dim DueDate As String
'duedate = "01/09/2017"
'duedate = "08/09/2017"
'duedate = "01/08/2017"



'
Dim rangecells As Range
Dim rangecells2 As Range
Dim mIntCtr, mIntCtr2, i, count As Long
count = 0


Dim rangecells1 As Range
Dim mIntCtr1, j, count1, count2, count3 As Long
count1 = 0
count2 = 0
count3 = 0
Set rangecells1 = Sheet1.Range("sheet1.AuditedSection").Cells
mIntCtr1 = Sheet1.Range("sheet1.AuditedSection").count

Set rangecells2 = Sheet1.Range("Audit.Act").Cells
mIntCtr2 = Sheet1.Range("Audit.Act").count

For i = 1 To mIntCtr1
    If (rangecells1.item(i).value = "92E") Then
        count1 = count + 1
    ElseIf Not ((rangecells1.item(i).value = "") Or (rangecells1.item(i).value = "(Select)")) Then
        count2 = count2 + 1
    End If
Next

For i = 1 To mIntCtr2

    If Not isdropdownblank(rangecells2.item(i).value) Then
        count3 = count3 + 1
    End If
Next


If Mid(Sheet1.Range("sheet1.LiableSec92Eflg").value, 1, 1) = "Y" Then
        count1 = count + 1
End If


'If (Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "14" Or Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "37") Then
'
'    If Count1 > 0 Then
'        DueDate = "01/04/2020"
'        GoTo endline:
'
'    ElseIf (count2 > 0 Or count3 > 0) Then
'        'duedate = "01/10/2017"
'        'duedate = "01/11/2017"
'        'duedate = "01/11/2017"
'        'DueDate = "01/10/2019"
'        DueDate = "01/04/2020"
'        'duedate = "18/10/2017"
'    End If
'
'Else
'    If Count1 > 0 Then
'        DueDate = "01/12/2019"
'        GoTo endline:
'
'    ElseIf (count2 > 0 Or count3 > 0) Then
'        'duedate = "01/10/2017"
'        'duedate = "01/11/2017"
'        'duedate = "01/11/2017"
'        DueDate = "01/10/2019"
'        'duedate = "18/10/2017"
'    End If
'
'End If

Dim auditFlag As String

    auditFlag = Mid(Sheet1.Range("sheet1.LiableSec44ABflg").value, 1, 1)
   
'If (Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "14" Or Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "37") Then
'    If (auditFlag = "Y") Then
'        'duedate = "01/10/2017"
'        'duedate = "01/11/2017"
'        DueDate = "01/04/2020"
'        'duedate = "18/10/2017"
'    End If
'Else
'    If (auditFlag = "Y") Then
'        'duedate = "01/10/2017"
'        'duedate = "01/11/2017"
'        DueDate = "01/10/2019"
'        'duedate = "18/10/2017"
'    End If
'End If

Set rangecells = Sheet21.Range("IF.IsLiableToAudit").Cells
mIntCtr = Sheet21.Range("IF.IsLiableToAudit").count

For i = 1 To mIntCtr
    If (Mid(rangecells.item(i).value, 1, 1) = "Y") Then
        count = count + 1
    End If
Next



'If (Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "14" Or Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "37") Then
'    If Count > 0 Then
'        'duedate = "01/10/2017"
'        'duedate = "01/11/2017"
'        'duedate = "01/11/2017"
'        DueDate = "01/04/2020"
'        'duedate = "18/10/2017"
'    End If
'Else
'    If Count > 0 Then
'        'duedate = "01/10/2017"
'        'duedate = "01/11/2017"
'        'duedate = "01/11/2017"
'        DueDate = "01/10/2019"
'        'duedate = "18/10/2017"
'    End If
'End If

'Chandru - Due date changed

If (Sheet33.Range("Temp_lakhcheck").value) > 100000 Then
    If Mid(Sheet1.Range("Sheet1.LiableSec92Eflg").value, 1, 1) = "Y" Or Sheet21.Range("checkin92E").value = "Y" Or Sheet29.Range("BooksSpouse92EFlg").value = "Yes" Then
'    DueDate = "30/11/2024"
    
'    DueDate = "15/12/2024" 'Changed by sai on 06/12/2024 for Due date extended to 15th Dec 2024
     DueDate = "30/11/2025"  'by sai on 28/03/2025
    
    Else
    If auditFlag = "Y" Or count > 0 Or count2 > 0 Or count3 > 0 Or Sheet21.Range("AuditcheckinIF").value = "Y" Or Sheet29.Range("BooksSpouse44ABFlg").value = "Yes" Then
'    DueDate = "31/10/2024"
    
'     DueDate = "15/11/2024" 'Changed by sai on 16/10/2024 - Bindu 27/11/24
     'DueDate = "31/10/2025"  'by sai on 28/03/2025
     
      DueDate = Sheet33.Range("AuditcaseDate_44AB") 'Sadineni_30_10_2025
     
    'DueDate = "07/11/2023"
    Else
'     DueDate = "31/07/2024"
     'DueDate = "15/09/2025"  'by sai on 28/03/2025
     
      DueDate = Sheet33.Range("NonAudit_DueDate") 'Sadineni_08_09_2025
      
    End If
    End If
End If


If (Sheet33.Range("Temp_lakhcheck").value) <= 100000 Then
    If Mid(Sheet1.Range("Sheet1.LiableSec92Eflg").value, 1, 1) = "Y" Or Sheet21.Range("checkin92E").value = "Y" Or Sheet29.Range("BooksSpouse92EFlg").value = "Yes" Then
'    DueDate = "30/11/2024"

'     DueDate = "15/12/2024" 'Changed by sai on 06/12/2024 for Due date extended to 15th Dec 2024
      DueDate = "30/11/2025"  'by sai on 28/03/2025
    Else
    If auditFlag = "Y" Or count > 0 Or count2 > 0 Or count3 > 0 Or Sheet21.Range("AuditcheckinIF").value = "Y" Or Sheet29.Range("BooksSpouse44ABFlg").value = "Yes" Then
'    DueDate = "31/10/2024"
      
'       DueDate = "15/11/2024" 'Changed by sai on 16/10/2024-Bindu 27/11/24
       'DueDate = "31/10/2025"  'by sai on 28/03/2025
        DueDate = Sheet33.Range("AuditcaseDate_44AB") 'Sadineni_30_10_2025



    'DueDate = "07/11/2023"
    Else
'    DueDate = "31/07/2024"
     'DueDate = "15/09/2025"  'by sai on 28/03/2025
      DueDate = Sheet33.Range("NonAudit_DueDate") 'Sadineni_08_09_2025
    End If
    End If

End If


endline:
        
getDueDate = DueDate
End Function
Function getDueDate()

Dim DueDate As String
Dim rangecells As Range
Dim rangecells2 As Range
Dim mIntCtr, mIntCtr2, i, count As Long
count = 0


Dim rangecells1 As Range
Dim mIntCtr1, j, count1, count2, count3 As Long
count1 = 0
count2 = 0
count3 = 0
Set rangecells1 = Sheet1.Range("sheet1.AuditedSection").Cells
mIntCtr1 = Sheet1.Range("sheet1.AuditedSection").count

Set rangecells2 = Sheet1.Range("Audit.Act").Cells
mIntCtr2 = Sheet1.Range("Audit.Act").count

For i = 1 To mIntCtr1
    If (rangecells1.item(i).value = "92E") Then
        count1 = count + 1
    ElseIf Not ((rangecells1.item(i).value = "") Or (rangecells1.item(i).value = "(Select)")) Then
        count2 = count2 + 1
    End If
Next

For i = 1 To mIntCtr2

    If Not isdropdownblank(rangecells2.item(i).value) Then
        count3 = count3 + 1
    End If
Next


If Mid(Sheet1.Range("sheet1.LiableSec92Eflg").value, 1, 1) = "Y" Then
        count1 = count + 1
End If


Dim auditFlag As String

auditFlag = Mid(Sheet1.Range("sheet1.LiableSec44ABflg").value, 1, 1)

Set rangecells = Sheet21.Range("IF.IsLiableToAudit").Cells
mIntCtr = Sheet21.Range("IF.IsLiableToAudit").count

For i = 1 To mIntCtr
    If (Mid(rangecells.item(i).value, 1, 1) = "Y") Then
        count = count + 1
    End If
Next


If (Sheet33.Range("Temp_lakhcheck").value) > 100000 Then
    
    If Mid(Sheet1.Range("Sheet1.LiableSec92Eflg").value, 1, 1) = "Y" Or Sheet21.Range("checkin92E").value = "Y" Or Sheet29.Range("BooksSpouse92EFlg").value = "Yes" Then
'    DueDate = "30/11/2024"
     DueDate = Sheet33.Range("AuditcaseDate_92E")

    Else
    
    If auditFlag = "Y" Or count > 0 Or count2 > 0 Or count3 > 0 Or Sheet21.Range("AuditcheckinIF").value = "Y" Or Sheet29.Range("BooksSpouse44ABFlg").value = "Yes" Then
'    DueDate = "31/10/2024"
      DueDate = Sheet33.Range("AuditcaseDate_44AB")
      
    Else
      
'     DueDate = "31/07/2024"
      DueDate = Sheet33.Range("NonAudit_DueDate")
      
    End If
  End If
End If


If (Sheet33.Range("Temp_lakhcheck").value) <= 100000 Then
    
    If Mid(Sheet1.Range("Sheet1.LiableSec92Eflg").value, 1, 1) = "Y" Or Sheet21.Range("checkin92E").value = "Y" Or Sheet29.Range("BooksSpouse92EFlg").value = "Yes" Then
'    DueDate = "30/11/2024"
      DueDate = Sheet33.Range("AuditcaseDate_92E")
    
    Else
    
    If auditFlag = "Y" Or count > 0 Or count2 > 0 Or count3 > 0 Or Sheet21.Range("AuditcheckinIF").value = "Y" Or Sheet29.Range("BooksSpouse44ABFlg").value = "Yes" Then
'    DueDate = "31/10/2024"
    DueDate = Sheet33.Range("AuditcaseDate_44AB")
    
    Else
    
'    DueDate = "31/07/2024"
    DueDate = Sheet33.Range("NonAudit_DueDate")
    
    End If
    End If

End If


endline:

getDueDate = DueDate
End Function


Function getSlabbedIncome(totalIncome As Double, value As Long) As Double
On Error Resume Next
Dim resStatus As String
Dim age As String
Dim Status As String
Dim normal10 As Double
Dim normal20 As Double
Dim normal30 As Double
Dim normal10A As Double
Dim normal10B As Double
Dim normal20A As Double
Dim netTxblIncome As Double
 netTxblIncome = totalIncome
 age = Sheet31.Range("age").value
 resStatus = Sheet33.Range("resStatus").value
 Status = Sheet33.Range("taxPayerStatus").value
  If (SheetALL.Range("bacValue").value <> 1) Then
  If (Status = "I" And (resStatus = "RES" Or resStatus = "NOR") And (age > 59 And age < 80)) Then

            If (netTxblIncome >= 0 And netTxblIncome <= 300000) Then
                    normal10 = 0
                    normal20 = 0
                    normal30 = 0
                    normal10A = 0
                    normal10B = 0
                    normal20A = 0
            ElseIf (netTxblIncome >= 300001 And netTxblIncome <= 500000) Then
                    normal10 = (netTxblIncome - 300000)
                    normal20 = 0
                    normal30 = 0
                    normal10A = 0
                    normal10B = 0
                    normal20A = 0
            ElseIf (netTxblIncome >= 500001 And netTxblIncome <= 1000000) Then
                    normal10 = 200000
                   normal20 = (netTxblIncome - 500000)
                   normal30 = 0
                   normal10A = 0
                    normal10B = 0
                    normal20A = 0
            ElseIf (netTxblIncome >= 1000001) Then
                   normal10 = 200000
                   normal20 = 500000
                   normal30 = (netTxblIncome - 1000000)
                   normal10A = 0
                    normal10B = 0
                    normal20A = 0
            End If
   
ElseIf (Status = "I" And (resStatus = "RES" Or resStatus = "NOR") And (age >= 80)) Then
                If (netTxblIncome >= 0 And netTxblIncome <= 500000) Then
                    normal10 = 0
                    normal20 = 0
                    normal30 = 0
                    normal10A = 0
                    normal10B = 0
                    normal20A = 0
                ElseIf (netTxblIncome >= 500001 And netTxblIncome <= 1000000) Then
                    normal10 = 0
                    normal20 = (netTxblIncome - 500000)
                    normal30 = 0
                    normal10A = 0
                    normal10B = 0
                    normal20A = 0
                ElseIf (netTxblIncome >= 1000001) Then
                    normal10 = 0
                    normal20 = 500000
                    normal30 = (netTxblIncome - 1000000)
                    normal10A = 0
                    normal10B = 0
                    normal20A = 0
                End If
              
ElseIf (Status = "I" Or Status = "H") Then
                If (netTxblIncome >= 0 And netTxblIncome <= 250000) Then
                    normal10 = 0
                    normal20 = 0
                    normal30 = 0
                    
                ElseIf (netTxblIncome >= 250001 And netTxblIncome <= 500000) Then
                    normal10 = (netTxblIncome - 250000)
                    normal20 = 0
                    normal30 = 0
                ElseIf (netTxblIncome >= 500001 And netTxblIncome <= 1000000) Then
                    normal10 = 250000
                    normal20 = (netTxblIncome - 500000)
                    normal30 = 0
                ElseIf (netTxblIncome >= 1000001) Then
                    normal10 = 250000
                    normal20 = 500000
                    normal30 = (netTxblIncome - 1000000)
                End If
End If
Else
                
'Changes by sai on 28/03/2025 for AY 2025-26
                
If (Status = "I" Or Status = "H") Then
                If (netTxblIncome >= 0 And netTxblIncome <= 400000) Then
                    normal10 = 0
                    normal20 = 0
                    normal30 = 0
                    normal10A = 0
                    normal10B = 0
                    normal20A = 0

                ElseIf (netTxblIncome >= 400001 And netTxblIncome <= 800000) Then
                    normal10 = (netTxblIncome - 400000)
                    '-----------------------------
                    normal20 = 0
                    normal30 = 0
                    normal10A = 0
                    normal10B = 0
                    normal20A = 0

                 ElseIf (netTxblIncome >= 800001 And netTxblIncome <= 1200000) Then

                    normal10 = 400000
                   ' normal10A = (netTxblIncome - 600000)  'AY_2024_25 OLD
                    normal10A = (netTxblIncome - 800000)   'AY_2025_26 NEW
                    normal10B = 0
                    normal20 = 0
                    normal20A = 0
                    normal30 = 0

                 ElseIf (netTxblIncome >= 1200001 And netTxblIncome <= 1600000) Then
                 'Changed from 3 to 4 lac
                    normal10 = 400000
                    normal10A = 400000
                    normal10B = (netTxblIncome - 1200000)  'AY_2025_26 NEW
                    normal20 = 0
                    normal20A = 0
                    normal30 = 0
                  ElseIf (netTxblIncome >= 1600001 And netTxblIncome <= 2000000) Then
                    normal10 = 400000
                    normal10A = 400000
                    normal10B = 400000
                    normal20 = (netTxblIncome - 1600000)
                    normal20A = 0
                    normal30 = 0
                   ElseIf (netTxblIncome >= 2000001 And netTxblIncome <= 2400000) Then
                    normal10 = 400000
                    normal10A = 400000
                    normal10B = 400000
                    normal20 = 400000
                    normal20A = (netTxblIncome - 2000000)
                    normal30 = 0
                ElseIf (netTxblIncome >= 2400001) Then
                    normal10 = 400000
                    normal10A = 400000
                    normal10B = 400000
                    normal20 = 400000
                    normal20A = 400000
                    normal30 = (netTxblIncome - 2400000)
                End If
                
'*************************************************************************************************
                
                        
End If
'xxxx
End If

If value = 1 Then
    getSlabbedIncome = normal10
ElseIf value = 10 Then
    getSlabbedIncome = normal10A
ElseIf value = 15 Then
    getSlabbedIncome = normal10B
ElseIf value = 2 Then
    getSlabbedIncome = normal20
ElseIf value = 25 Then
    getSlabbedIncome = normal20A
ElseIf value = 3 Then
    getSlabbedIncome = normal30
End If


End Function

'Chandru AY2024-25 Dev
Function calculateTaxPayable(netTxblIncome As Double) As Double
On Error Resume Next
Dim resStatus As String
Dim age As String
Dim Status As String

 netTxblIncome = netTxblIncome
 age = Sheet31.Range("age").value
 resStatus = Sheet33.Range("resStatus").value
 Status = Sheet33.Range("taxPayerStatus").value

  If (Status = "I" And (resStatus = "RES" Or resStatus = "NOR") And (age > 59 And age < 80) And (SheetALL.Range("bacValue").value = 2)) Then

            If (netTxblIncome >= 0 And netTxblIncome <= 300000) Then
                    calculateTaxPayable = 0
            ElseIf (netTxblIncome >= 300001 And netTxblIncome <= 500000) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 300000) * 0.05, 0)
            ElseIf (netTxblIncome >= 500001 And netTxblIncome <= 1000000) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 500000) * 0.2 + 10000, 0)
            ElseIf (netTxblIncome >= 1000001) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 1000000) * 0.3 + 110000, 0)
            End If
            
ElseIf (Status = "I" And (resStatus = "RES" Or resStatus = "NOR") And (age >= 80) And (SheetALL.Range("bacValue").value = 2)) Then
                If (netTxblIncome >= 0 And netTxblIncome <= 500000) Then
                    calculateTaxPayable = 0
                ElseIf (netTxblIncome >= 500001 And netTxblIncome <= 1000000) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 500000) * 0.2, 0)
                ElseIf (netTxblIncome >= 1000001) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 1000000) * 0.3 + 100000, 0)
                End If
                
ElseIf (Status = "I" Or Status = "H") And (SheetALL.Range("bacValue").value <> 1) Then
                If (netTxblIncome >= 0 And netTxblIncome <= 250000) Then
                    calculateTaxPayable = 0
                ElseIf (netTxblIncome >= 250001 And netTxblIncome <= 500000) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 250000) * 0.05, 0)
                ElseIf (netTxblIncome >= 500001 And netTxblIncome <= 1000000) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 500000) * 0.2 + 12500, 0)
                ElseIf (netTxblIncome >= 1000001) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 1000000) * 0.3 + 112500, 0)
                End If

ElseIf (Status = "H" Or Status = "H") Then
                If (netTxblIncome >= 0 And netTxblIncome <= 250000) Then
                    calculateTaxPayable = 0
                ElseIf (netTxblIncome >= 250001 And netTxblIncome <= 500000) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 250000) * 0.05, 0)
                ElseIf (netTxblIncome >= 500001 And netTxblIncome <= 1000000) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 500000) * 0.2 + 12500, 0)
                ElseIf (netTxblIncome >= 1000001) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 1000000) * 0.3 + 112500, 0)
                End If
                
                
ElseIf (Status = "I") And (SheetALL.Range("bacValue").value = 1) Then
                If (netTxblIncome >= 0 And netTxblIncome <= 300000) Then
                    calculateTaxPayable = 0

'Changed by Riyaz on 26/02/2026 for AY 2026-27
                ElseIf (netTxblIncome >= 400001 And netTxblIncome <= 800000) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 400000) * 0.05, 0)
                ElseIf (netTxblIncome >= 800001 And netTxblIncome <= 1200000) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 800000) * 0.1 + 20000, 0)
                ElseIf (netTxblIncome >= 1200001 And netTxblIncome <= 1600000) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 1200000) * 0.15 + 60000, 0)
                ElseIf (netTxblIncome >= 1600001 And netTxblIncome <= 2000000) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 1600000) * 0.2 + 140000, 0)
                ElseIf (netTxblIncome >= 2000001 And netTxblIncome <= 2400000) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 2000000) * 0.2 + 200000, 0)
                ElseIf (netTxblIncome >= 2400001) Then
                     calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 2400000) * 0.3 + 300000, 0)
                End If

End If
End Function

Function getSlabbedIncomeold(totalIncome As Double, value As Long) As Double
On Error Resume Next
Dim resStatus As String
Dim age As String
Dim Status As String
Dim normal10 As Double
Dim normal20 As Double
Dim normal30 As Double
Dim netTxblIncome As Double
 netTxblIncome = totalIncome
 age = Sheet31.Range("age").value
 resStatus = Sheet33.Range("resStatus").value
 Status = Sheet33.Range("taxPayerStatus").value

  If (Status = "I" And (resStatus = "RES" Or resStatus = "NOR") And (age > 59 And age < 80)) Then

            If (netTxblIncome >= 0 And netTxblIncome <= 300000) Then
                    normal10 = 0
                    normal20 = 0
                    normal30 = 0
            ElseIf (netTxblIncome >= 300001 And netTxblIncome <= 500000) Then
                    normal10 = (netTxblIncome - 300000)
                    normal20 = 0
                    normal30 = 0
            ElseIf (netTxblIncome >= 500001 And netTxblIncome <= 1000000) Then
                    normal10 = 200000
                   normal20 = (netTxblIncome - 500000)
                   normal30 = 0
            ElseIf (netTxblIncome >= 1000001) Then
                   normal10 = 200000
                   normal20 = 500000
                   normal30 = (netTxblIncome - 1000000)
            End If
   
ElseIf (Status = "I" And (resStatus = "RES" Or resStatus = "NOR") And (age >= 80)) Then
                If (netTxblIncome >= 0 And netTxblIncome <= 500000) Then
                    normal10 = 0
                    normal20 = 0
                    normal30 = 0
                ElseIf (netTxblIncome >= 500001 And netTxblIncome <= 1000000) Then
                    normal10 = 0
                    normal20 = (netTxblIncome - 500000)
                    normal30 = 0
                ElseIf (netTxblIncome >= 1000001) Then
                    normal10 = 0
                    normal20 = 500000
                    normal30 = (netTxblIncome - 1000000)
                End If
              
ElseIf (Status = "I" Or Status = "H") Then
                If (netTxblIncome >= 0 And netTxblIncome <= 250000) Then
                    normal10 = 0
                    normal20 = 0
                    normal30 = 0
                ElseIf (netTxblIncome >= 250001 And netTxblIncome <= 500000) Then
                    normal10 = (netTxblIncome - 250000)
                    normal20 = 0
                    normal30 = 0
                ElseIf (netTxblIncome >= 500001 And netTxblIncome <= 1000000) Then
                    normal10 = 250000
                    normal20 = (netTxblIncome - 500000)
                    normal30 = 0
                ElseIf (netTxblIncome >= 1000001) Then
                    normal10 = 250000
                    normal20 = 500000
                    normal30 = (netTxblIncome - 1000000)
                End If
End If

If value = 1 Then
    getSlabbedIncome = normal10
ElseIf value = 2 Then
    getSlabbedIncome = normal20
ElseIf value = 3 Then
    getSlabbedIncome = normal30
End If


End Function

Function calculateTaxPayableold(netTxblIncome As Double) As Double
On Error Resume Next
Dim resStatus As String
Dim age As String
Dim Status As String

 netTxblIncome = netTxblIncome
 age = Sheet31.Range("age").value
 resStatus = Sheet33.Range("resStatus").value
 Status = Sheet33.Range("taxPayerStatus").value

  If (Status = "I" And (resStatus = "RES" Or resStatus = "NOR") And (age > 59 And age < 80)) Then

            If (netTxblIncome >= 0 And netTxblIncome <= 300000) Then
                    calculateTaxPayable = 0
            ElseIf (netTxblIncome >= 300001 And netTxblIncome <= 500000) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 300000) * 0.05, 0)
            ElseIf (netTxblIncome >= 500001 And netTxblIncome <= 1000000) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 500000) * 0.2 + 10000, 0)
            ElseIf (netTxblIncome >= 1000001) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 1000000) * 0.3 + 110000, 0)
            End If
            
ElseIf (Status = "I" And (resStatus = "RES" Or resStatus = "NOR") And (age >= 80)) Then
                If (netTxblIncome >= 0 And netTxblIncome <= 500000) Then
                    calculateTaxPayable = 0
                ElseIf (netTxblIncome >= 500001 And netTxblIncome <= 1000000) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 500000) * 0.2, 0)
                ElseIf (netTxblIncome >= 1000001) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 1000000) * 0.3 + 100000, 0)
                End If
                
ElseIf (Status = "I" Or Status = "H") Then
                If (netTxblIncome >= 0 And netTxblIncome <= 250000) Then
                    calculateTaxPayable = 0
                ElseIf (netTxblIncome >= 250001 And netTxblIncome <= 500000) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 250000) * 0.05, 0)
                ElseIf (netTxblIncome >= 500001 And netTxblIncome <= 1000000) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 500000) * 0.2 + 12500, 0)
                ElseIf (netTxblIncome >= 1000001) Then
                    calculateTaxPayable = Application.WorksheetFunction.Round((netTxblIncome - 1000000) * 0.3 + 112500, 0)
                End If
End If
                
End Function


Sub calculateTax_New() 'For Rebate 87A Changes by sai 06/01/2025
   Call calculateTax(True)
End Sub










