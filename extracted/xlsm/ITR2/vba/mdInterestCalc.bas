Attribute VB_Name = "mdInterestCalc"
Option Explicit

Dim intrstPayUs234A As Double
Dim intrstPayUs234B As Double
Dim intrstPayUs234C As Double
Dim intrstPayUs234F As Double
Public agestatus As Variant
Public age1 As Variant
Public CONST_IntrstPay234B_Percentage_1, April_1, may_1, June_1 As Variant

'Added by Riyaz on 02/02/2026
Public MonthsAfterDueDate As Long

Dim calcInterestPayable234B As Double
Const CONST_IntrstPay234B_Percentage As Double = 1
Const CONST_NET_Limit As Double = 10000
Const CONST_ATP_Limit As Double = 90
Dim shortFall As Double
Public OrigRetFiledDate_1

Sub ComputeInterest()
'On Error GoTo endline
'Condition
'1. Net Tax Liability -TDS\TCS <10000 :: 234B\234C will be Zero
'2. Advance tax paid is 90 % or more of Net Tax Libilty -TDS\TCS :: 234B will be zero and 234 C will be calculated
    
    Dim SysCalculatedNetTaxLiability As Double
    Dim matchedAdvanceTax As Double
    Dim tdsamtused As Double
    Dim tcsamtused As Double
   
    sPassword = EfilingCommon.getmsgstate
    Sheet32.Unprotect Password:=sPassword

    intrstPayUs234A = 0
    intrstPayUs234B = 0
    intrstPayUs234C = 0
    intrstPayUs234F = 0
            
    filingdate
        
    SysCalculatedNetTaxLiability = Sheet7.Range("Sheet9.NetTaxLiability").value

    matchedAdvanceTax = Sheet7.Range("Sheet9.AdvanceTax").value
    tdsamtused = Sheet7.Range("Sheet9.TDS").value
    tcsamtused = Sheet7.Range("Sheet9.TCS").value
            
            
age1 = AY23_24Changes.calculateAge(Sheet1.Range("sheet1.DOB").value)
If (age1 <= 59) Then
agestatus = "NC"
ElseIf ((age1 > 59) And (age1 <= 79)) Then
agestatus = "SC"
Else
agestatus = "SSC"
End If
        
        
'''''''''''''''''''''''''''''''''''''''''''''Interest234C - Start''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'If Sheet7.Range("sheet9.GrossTaxPayable").value <> Sheet7.Range("Sheet9.GrossTaxLiability").value Then
    
    'Malli
    'If Sheet31.Range("age") > 59 Then
'    If Sheet31.Range("age") > 59 And Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Then

'Added by Riyaz on 29/10/2025 for SC & Senior citizen
    
    If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" And (Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Or Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NOR") And (agestatus = "SC" Or agestatus = "SSC") Then
            
            intrstPayUs234C = 0
            Sheet7.Range("Sheet9.IntrstPayUs234C").value = 0
    Else
        If Sheet7.Range("sheet9.GrossTaxPayable").value <> (Sheet7.Range("Sheet9.GrossTaxLiability").value - Sheet7.Range("sheet9.GrossTaxPayable_3b").value) Then
            intrstPayUs234C = Sheet33.Range("Total_234Cii").value
            Sheet7.Range("Sheet9.IntrstPayUs234C").value = WorksheetFunction.Round(Sheet33.Range("Total_234Cii").value, 0)
        Else
            intrstPayUs234C = Sheet33.Range("interest234C").value
            Sheet7.Range("Sheet9.IntrstPayUs234C").value = WorksheetFunction.Round(Sheet33.Range("interest234C").value, 0)
        End If
    End If
'''''''''''''''''''''''''''''''''''''''''''''Interest234C - Done''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''Interest234F - Start''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


If Mid(Sheet7.Range("AOIFlag_1").value, 1, 1) = "Y" Or Mid(Sheet1.Range("sheet1.SeventhProvisoFlag"), 1, 1) = "Y" Then
Calculate234F

'ElseIf Range("Sheet8b.LossesOfCurrentYearCarriedFwd").value > 0 Then
'Calculate234F

'ElseIf (Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Or Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Or SheetALL.Range("bacValue").value = 1) And (Range("Sheet8b.GrossTotalIncome").value + Range("Sheet20.scei.LTCGWhereSTTPaid").value > 250000) Then
'Chandru change as per AY2024-25
'ElseIf (Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Or Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H" Or SheetALL.Range("bacValue").value = 1) And (Range("Sheet8b.GrossTotalIncome").value + Range("Sheet20.scei.LTCGWhereSTTPaid").value > 250000) Then

'Changed by Riyaz 03/02/2026 3Lac to 4Lac
ElseIf SheetALL.Range("bacValue").value = 1 And (Range("Sheet8b.GrossTotalIncome").value + Range("Sheet20.scei.LTCGWhereSTTPaid").value > 400000) Then
Calculate234F
ElseIf SheetALL.Range("bacValue").value = 2 And (Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Or Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "H") And (Range("Sheet8b.GrossTotalIncome").value + Range("Sheet20.scei.LTCGWhereSTTPaid").value > 250000) Then
Calculate234F
ElseIf SheetALL.Range("bacValue").value = 2 And agestatus = "NC" And (Range("Sheet8b.GrossTotalIncome").value + Range("Sheet20.scei.LTCGWhereSTTPaid").value > 250000) Then
Calculate234F
ElseIf SheetALL.Range("bacValue").value = 2 And agestatus = "SC" And (Range("Sheet8b.GrossTotalIncome").value + Range("Sheet20.scei.LTCGWhereSTTPaid").value > 300000) Then
Calculate234F
ElseIf SheetALL.Range("bacValue").value = 2 And agestatus = "SSC" And (Range("Sheet8b.GrossTotalIncome").value + Range("Sheet20.scei.LTCGWhereSTTPaid").value > 500000) Then
Calculate234F
Else
intrstPayUs234F = 0
End If
Sheet7.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
'''''''''''''''''''''''''''''''''''''''''''''Interest234F - Done''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                   
'''''''''''''''''''''''''''''''''''''''''''''Interest234B - Start''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

     
    intrstPayUs234B = Calculate_InterestPayable234B(SysCalculatedNetTaxLiability, _
                                                     matchedAdvanceTax, _
                                                     tdsamtused, _
                                                     tcsamtused, _
                                                     intrstPayUs234A, _
                                                     intrstPayUs234F, _
                                                     intrstPayUs234C)
                                                            
                                                            
'    Sheet7.Range("Sheet9.IntrstPayUs234B").value = Round(intrstPayUs234B)
'Added by Riyaz on 29/10/2025 for SC  & Senior citizen
 
 If Mid(Sheet1.Range("sheet1.Status").value, 1, 1) = "I" And (Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "RES" Or Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NOR") And (agestatus = "SC" Or agestatus = "SSC") Then
    Sheet7.Range("Sheet9.IntrstPayUs234B").value = 0
 Else
    Sheet7.Range("Sheet9.IntrstPayUs234B").value = Round(intrstPayUs234B)
 End If

'''''''''''''''''''''''''''''''''''''''''''''Interest234B - Done''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''Interest234A - Start''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    intrstPayUs234A = Sheet33.Range("section234A").value
'    Sheet7.Range("Sheet9.IntrstPayUs234A").value = WorksheetFunction.Round(Sheet33.Range("section234A").value, 0)


Dim i As Long
'Dim calcInterestPayable234A As Long
 Dim calcInterestPayable234A As Double


    calcInterestPayable234A = 0
    
    If MonthsAfterDueDate = 0 Then
        calcInterestPayable234A = 0
    Else
        For i = 1 To MonthsAfterDueDate
            calcInterestPayable234A = calcInterestPayable234A + Sheet33.Range("DG" & i + 8).value
        Next i
    End If

    Sheet7.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet7.Range("Sheet9.IntrstPayUs234A").value = calcInterestPayable234A
    Sheet7.Protect Password:=EfilingCommon.getmsgstate
    
    Sheet33.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet33.Range("section234A").value = calcInterestPayable234A
    Sheet33.Protect Password:=EfilingCommon.getmsgstate



'''''''''''''''''''''''''''''''''''''''''''''Interest234A - Done''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


endline:

   Sheet32.Protect Password:=sPassword
End Sub

Function Calculate_InterestPayable234B(SysCalculatedNetTaxLiability As Double, _
                                        matchedAdvanceTax As Double, _
                                        tdsamtused As Double, _
                                        tcsamtused As Double, _
                                        intrstPayUs234A_1 As Double, _
                                        intrstPayUs234F_1 As Double, _
                                        intrstPayUs234C_1 As Double) As Double
        
        Dim balancePrincipal As Double
        Dim balanceInterest As Double
        Dim adjustedPrincipal As Double
        Dim adjustedInterest As Double
        Dim SATPaidAtPeriod As Double
        Dim carryForwardPrinicipal As Double
        Dim carryForwardInterest As Double
        Dim calcIntrst234BOnPeriod As Double
        Dim balancePrincipalR As Double
        Dim balanceInterestR As Double
        Dim calcIntrst234BOnPeriodR As Double
        Dim calcIntrst234BUptoPeriod As Double
        
        Dim i As Long
        Dim AssYear As Long
        Dim dateOfProcessing As String
        Dim yrdop As Long
        Dim mthdop As Long
        Dim Is44AD As String
        Dim NetTaxLiability As Variant
        Dim j As Long
        Dim SATPaidAtlooper As Double
        Dim looperint, looper As Long
        
        looperint = 0
        looper = 0
  
        SysCalculatedNetTaxLiability = Sheet33.Range("TaxLiability").value
        calcInterestPayable234B = 0
            
        Dim AdditionalTax As Variant
        AdditionalTax = 0
        
        If Sheet202.Range("U_Refund").value <> "" And (Sheet202.Range("U_TotRefund").value = "" Or Sheet202.Range("U_TotRefund").value = 0) Then
        AdditionalTax = Sheet202.Range("U_Refund").value
        End If
        
        If Sheet202.Range("U_TotRefund") <> "" Then
        AdditionalTax = AdditionalTax + Sheet202.Range("U_TotRefund").value
        End If
        
        If Mid(Range("sheet1.ReturnFileSec"), 1, 7) = "139(8A)" Then
        SysCalculatedNetTaxLiability = SysCalculatedNetTaxLiability + AdditionalTax
        End If
    
   If ((SysCalculatedNetTaxLiability - (tdsamtused + tcsamtused)) >= CONST_NET_Limit And _
        matchedAdvanceTax < CONST_ATP_Limit / 100 * (SysCalculatedNetTaxLiability - (tdsamtused + tcsamtused))) Then
        
        shortFall = WorksheetFunction.Max(0, SysCalculatedNetTaxLiability - (matchedAdvanceTax + tdsamtused + tcsamtused))
                          
         If (shortFall = 0) Then
             Calculate_InterestPayable234B = 0
          Else
             balancePrincipal = 0
             balanceInterest = 0
             adjustedPrincipal = 0
             adjustedInterest = 0
             SATPaidAtPeriod = 0
             carryForwardPrinicipal = 0
             carryForwardInterest = 0
             calcIntrst234BOnPeriod = 0
             balancePrincipalR = 0
             balanceInterestR = 0
             calcIntrst234BOnPeriodR = 0
             calcIntrst234BUptoPeriod = 0
        End If
        
        Dim Filingtype As String
        
'Commented by Riyaz for testing purpose

              dateOfProcessing = Sheet32.Range("DateOfFiling").value
'             dateOfProcessing = Sheet52.Range("sheet9.Date").value
           
'Changed by Riyaz 2025 to 2026
        AssYear = 2026 'dpk1601
        dateOfProcessing = Dformat(dateOfProcessing, "yyyy-mm-dd")
        
        yrdop = CLng(Mid(dateOfProcessing, 1, 4))
        mthdop = CLng(Mid(dateOfProcessing, 6, 2))
        
        If (yrdop >= AssYear) Then
            calcIntrst234BUptoPeriod = mthdop - 4 + (yrdop - AssYear) * 12
        End If
        
        calcIntrst234BUptoPeriod = calcIntrst234BUptoPeriod + 1
        
       carryForwardPrinicipal = shortFall
       carryForwardPrinicipal = carryForwardPrinicipal
        
'Added by Riyaz on 08/09/2025
       Sheet33.Unprotect Password:=EfilingCommon.getmsgstate
            Sheet33.Range("Month_Wise_234Bvalues").ClearContents
            Sheet33.Range("Balance_Interest").ClearContents
            Sheet33.Range("IT_Interest").ClearContents
       Sheet33.Protect Password:=EfilingCommon.getmsgstate
       
                Dim month_count_1 As Variant
                Dim DueDate_1 As String
                DueDate_1 = getDueDate
                month_count_1 = (CLng(Mid(Dformat(DueDate_1, "yyyy-mm-dd"), 6, 2)) - 2)
        
        For i = 1 To calcIntrst234BUptoPeriod
            
                SATPaidAtPeriod = 0
                SATPaidAtlooper = 0
                balancePrincipal = carryForwardPrinicipal

                calcIntrst234BOnPeriod = Math.Round(CONST_IntrstPay234B_Percentage / 100 * WorksheetFunction.RoundDown(balancePrincipal, -2))

'Changed by Riyaz on 24/04/2024
'                If (age1 > 59) Then        Updated by Riyaz---SIT-119701
                If (age1 > 59) And Sheet1.Range("sheet1.Status").value = "I-INDIVIDUAL" And UCase(Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3)) = "RES" Then
                    calcIntrst234BOnPeriod = 0
                    Sheet33.Unprotect Password:=EfilingCommon.getmsgstate
                    Sheet33.Range("DB" & i + 4).value = calcIntrst234BOnPeriod
                    Sheet33.Protect Password:=EfilingCommon.getmsgstate
                Else
                    Sheet33.Unprotect Password:=EfilingCommon.getmsgstate
                    Sheet33.Range("DB" & i + 4).value = calcIntrst234BOnPeriod
                    Sheet33.Protect Password:=EfilingCommon.getmsgstate
    
                    calcInterestPayable234B = calcInterestPayable234B + calcIntrst234BOnPeriod
                End If
'-------end----

'Changed by Riyaz on 24/04/2024
                If (Mid(Range("sheet1.ReturnFileSec1"), 1, 2) = 17) Or (Mid(Range("sheet1.ReturnFileSec1"), 1, 2) = 18) Or (Mid(Range("sheet1.ReturnFileSec1"), 1, 2) = 19) Or _
                (((Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))) = "21") And (Mid(Range("U_PreviouslyFiledForThisAY"), 1, 1) = "Y")) Then
                    
 'Added by Riyaz on 24/04/2024 for original return date diffrence month calc
                Dim dateOfProcessing1 As String
                Dim yrdop1 As Long
                Dim mthdop1 As Long
                Dim Revisedmonthdiff As Double
                Dim AssYear1 As Long
                
                dateOfProcessing1 = Sheet1.Range("sheet1.OrigRetFiledDate").value
        
                AssYear1 = 2026
                dateOfProcessing1 = Dformat(dateOfProcessing1, "yyyy-mm-dd")
        
                yrdop1 = CInt(Mid(dateOfProcessing1, 1, 4))
                mthdop1 = CInt(Mid(dateOfProcessing1, 6, 2))
        
                If (yrdop1 >= AssYear) Then
                    Revisedmonthdiff = mthdop1 - 4 + (yrdop1 - AssYear1) * 12
                End If
        
                    Revisedmonthdiff = Revisedmonthdiff + 1
'End----
                    
                    If MonthsAfterDueDate = 0 Then
                        If i = 1 Then
                            balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234C_1
                        Else
                            balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                        End If
                    Else
                        If i = 1 Then
                            balanceInterest = Sheet33.Range("DG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234C_1
                        ElseIf i = month_count_1 And getDueDate = Sheet33.Range("G4").value Then
                            balanceInterest = Sheet33.Range("DG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234F_1
                        ElseIf i = month_count_1 And getDueDate = Sheet33.Range("G5").value Then
                            balanceInterest = Sheet33.Range("DG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234F_1
                        ElseIf i = month_count_1 And getDueDate = Sheet33.Range("G6").value Then
                            balanceInterest = Sheet33.Range("DG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234F_1
                        ElseIf i <= Revisedmonthdiff Then
                            balanceInterest = Sheet33.Range("DG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod
                        Else
                            balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                            
                        End If
                    End If
                Else
                    If i = 1 Then
                        balanceInterest = Sheet33.Range("DG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234C_1
                    ElseIf i = month_count_1 And getDueDate = Sheet33.Range("G4").value Then
                        balanceInterest = Sheet33.Range("DG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234F_1
                    ElseIf i = month_count_1 And getDueDate = Sheet33.Range("G5").value Then
                        balanceInterest = Sheet33.Range("DG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234F_1
                    ElseIf i = month_count_1 And getDueDate = Sheet33.Range("G6").value Then
                        balanceInterest = Sheet33.Range("DG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234F_1
                    Else
                        balanceInterest = Sheet33.Range("DG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod
                    End If
                
                End If
'-------end----

'Added by Riyaz on 08/09/2025
                Sheet33.Unprotect Password:=EfilingCommon.getmsgstate
                Sheet33.Range("DD" & i + 5).value = balanceInterest
                Sheet33.Protect Password:=EfilingCommon.getmsgstate

                SATPaidAtPeriod = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & i, Range("IT.Amt"))  'Uncommente by Shrutika05/01/2026


'Added by Riyaz on 08/09/2025
                Sheet33.Unprotect Password:=EfilingCommon.getmsgstate
                Sheet33.Range("DE" & i + 5).value = SATPaidAtPeriod
                Sheet33.Protect Password:=EfilingCommon.getmsgstate

                adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)

        Next
    End If
    
    Sheet33.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet33.Range("DI2").value = getDueDate
    Sheet33.Range("DI3").value = getDueDate
    Sheet33.Range("DI4").value = getDueDate
    Sheet33.Protect Password:=EfilingCommon.getmsgstate

    Calculate_InterestPayable234B = WorksheetFunction.Round(calcInterestPayable234B, 0)
End Function


Sub filingdate()
    Dim todaysdate As String
    Dim newfilingdate As String
    Dim originalfilingdate As String
    Dim origrevised As Boolean
    Dim origrevisedstatus As String
    Dim retfilestatus As String
    
    
    'origrevisedstatus = Mid(Range("sheet1.ReturnType1").value, 1, 1)
    retfilestatus = Mid(Range("sheet1.ReturnFileSec1"), 1, 2)
    If retfilestatus = "18" Or _
       retfilestatus = "17" Then
    origrevisedstatus = "R"
    End If
        
    todaysdate = Sheet32.Range("DateOfProcessing").value
    If (origrevisedstatus = "R") Then
        If Not ValidateOrigRetFiledDate_1() Then
            fmsgboxStatus "Please enter Original Return Date "
        Else
            newfilingdate = OrigRetFiledDate_1
        End If
        newfilingdate = Mid(newfilingdate, 9, 2) + "/" + Mid(newfilingdate, 6, 2) + "/" + Mid(newfilingdate, 1, 4)
        Sheet32.Range("DateOfFiling234A").value = newfilingdate
    End If
    
    'If (origrevisedstatus <> "R") Then
        If Not ValidateDate_9() Then
            fmsgboxStatus "Please enter Verification Date"
            CloseMsg
        Else
            newfilingdate = Date_9
            If Not ChkMinInclusiveDate(Date_9, todaysdate) Then
                newfilingdate = todaysdate
            End If
            newfilingdate = Mid(newfilingdate, 9, 2) + "/" + Mid(newfilingdate, 6, 2) + "/" + Mid(newfilingdate, 1, 4)
            Sheet32.Range("DateOfFiling").value = newfilingdate
        End If
    'End If
End Sub

Function ValidateOrigRetFiledDate_1() As Boolean
    ValidateOrigRetFiledDate_1 = True
    
    OrigRetFiledDate_1 = Range("sheet1.OrigRetFiledDate").value
    If Not CheckDateddmmyyyy(OrigRetFiledDate_1) Then
        ValidateOrigRetFiledDate_1 = False
        'msgbox1 ("OrigRetFiledDate in Sheet : PART A - GENERAL  must be a valid dd/mm/yyyy format")
        Exit Function
    Else
        OrigRetFiledDate_1 = Dformat(Range("sheet1.OrigRetFiledDate"), "yyyy-mm-dd")
        If Len(OrigRetFiledDate_1) > 0 Then
            If Not CheckDateMinDDMMYYYY(OrigRetFiledDate_1, 1, 4, 2021, "Original return filing date cannot be less than 01/04/2021") Then
                ValidateOrigRetFiledDate_1 = False
                'msgbox1 ("Original return filing date to be corrected as it should not be less than 01/04/2017")
                Exit Function
            Else
                OrigRetFiledDate_1 = Dformat(Range("sheet1.OrigRetFiledDate"), "yyyy-mm-dd")
            End If
        End If
    End If
End Function

Function ValidateDate_9() As Boolean
    ValidateDate_9 = True
    Date_9 = Sheet52.Range("sheet9.Date").value
    
    If Date_9 = "" Or IsEmpty(Date_9) Then
        'MsgCalcInterest = MsgCalcInterest & "Verification Date in Sheet : PARTB and Verification is mandatory" & Chr(13)
        ValidateDate_9 = False
        Exit Function
    End If
    
    If Not CheckDateddmmyyyy(Date_9) Then
        ValidateDate_9 = False
        'MsgCalcInterest = MsgCalcInterest & "Date in Sheet : PARTB and Verification   must be a valid dd/mm/yyyy format" & Chr(13)
        Exit Function
    Else
        Date_9 = Dformat(Sheet52.Range("sheet9.Date"), "yyyy-mm-dd")
    End If
    
    If Not CheckDateMinDDMMYYYY(Date_9, 1, 4, 2024, "Verification date cannot be less than 01/04/2024") Then 'dpk1601
        ValidateDate_9 = False
        'MsgCalcInterest = MsgCalcInterest & "Date in Verification , Sheet PARTB and Verification must not be less than 01/04/2017" & Chr(13)
        Exit Function
    Else
        Date_9 = Dformat(Sheet52.Range("sheet9.Date"), "yyyy-mm-dd")
    End If
End Function

Function MonthDiff(startDate As String, endDate As String) As Long
    Dim NoofDaysinmonth As Long
    Dim endyear, startyear, startmonth, endmonth, startday As Long
    Dim noofdaysinendmonth As Long
    
    NoofDaysinmonth = 31
    MonthDiff = 0
    startDate = CStr(Dformat(startDate, "yyyy-mm-dd"))
    endDate = CStr(Dformat(endDate, "yyyy-mm-dd"))
    startyear = Mid(startDate, 1, 4)
    endyear = Mid(endDate, 1, 4)
    startmonth = Mid(startDate, 6, 2)
    If (startmonth = 2) Then
        NoofDaysinmonth = 28
    End If
    
    startday = Mid(startDate, 9, 2)
    If ((startmonth = 4) Or (startmonth = 6) Or (startmonth = 9) Or (startmonth = 11)) Then
        NoofDaysinmonth = 30
    End If
    
    
    noofdaysinendmonth = 31
    endmonth = Mid(endDate, 6, 2)
    If (endmonth = 2) Then
        noofdaysinendmonth = 28
    End If
    
    If ((endmonth = 4) Or (endmonth = 6) Or (endmonth = 9) Or (endmonth = 11)) Then
        noofdaysinendmonth = 30
    End If
    
    MonthDiff = (CLng(endyear) - CLng(startyear)) * 12
    MonthDiff = MonthDiff + (CLng(endmonth) - CLng(startmonth))
    
    If (startDate < endDate) Then
        If (CLng(startday) < Mid(endDate, 9, 2)) Then
            MonthDiff = MonthDiff + 1
            If CLng(startday) = 30 And NoofDaysinmonth = 30 Then
                MonthDiff = MonthDiff - 1
            End If
        ElseIf (CLng(startday)) = 7 And Mid(endDate, 9, 2) <= (CLng(startday)) Then
            MonthDiff = MonthDiff + 1
        End If
    End If
End Function


Function Calculate234F()

Dim Returnfiledstatus As String
Dim VerificationDate As Variant
Dim dateOfFiling As Variant
Dim VerificationDate1 As Variant
Dim DateOfFiling1 As Variant

Dim check92EFlag As Variant
Dim Altduedate As Variant
Dim CheckAuditFLAG As Variant
Dim DueDate As Variant
Dim RepSec92EOtherFlag_2, RepSec92EOtherFlag_2_1 As Variant
Dim CheckAuditFLAGIF As Variant
Dim todaysdate1, Temp234f As Variant
    Dim UpdatedY As Boolean
    UpdatedY = False
    If (((Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))) = "21") And (Mid(Range("U_PreviouslyFiledForThisAY"), 1, 1) = "Y")) Then
    UpdatedY = True
    End If

    Dim UpdatedN As Boolean
    UpdatedN = False
    If (((Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))) = "21") And (Mid(Range("U_PreviouslyFiledForThisAY"), 1, 1) <> "Y")) Then
    UpdatedN = True
    End If

todaysdate1 = Sheet32.Range("DateOfProcessing").value
'Const FilingLastDate As Date = #1/8/2019#
'Const YearEndDate As Date = #12/31/2019#
'Temp234f = 0
'Temp234f = setTblinfo_AuditInfo_Front()
sPassword = EfilingCommon.getmsgstate
           Sheet33.Unprotect Password:=sPassword
'Sheet33.Range("FlagAuditTable").value = setTblinfo_AuditInfo_Front()
Sheet33.Protect Password:=sPassword
RepSec92EOtherFlag_2 = Sheet33.Range("Flag92EO").value
RepSec92EOtherFlag_2_1 = Sheet33.Range("FlagAuditTable").value
If (RepSec92EOtherFlag_2 = "1" Or RepSec92EOtherFlag_2_1 = "1") Then
RepSec92EOtherFlag_2 = True
Else
RepSec92EOtherFlag_2 = False
End If
'
'check92EFlag = Mid(Sheet1.Range("sheet1.LiableSec92Eflg"), 1, 1)
check92EFlag = "N"
'CheckAuditFLAG = Mid(Sheet1.Range("sheet1.LiableSec44ABflg"), 1, 1)
CheckAuditFLAG = "N"
If (CheckAuditFLAG = "Y") Then
CheckAuditFLAG = True
Else
CheckAuditFLAG = False
End If

'CheckAuditFLAGIF = Sheet21.Range("AuditcheckinIF").value
CheckAuditFLAGIF = "N"
If (CheckAuditFLAGIF = "Y") Then
CheckAuditFLAGIF = True
Else
CheckAuditFLAGIF = False
End If

'If check92EFlag = "Y" Then
'If check92EFlag = "Y" Or Sheet21.Range("checkin92E").value = "Y" Or Sheet29.Range("BooksSpouse92EFlg").value = "Yes" Then
If check92EFlag = "Y" Or Sheet21.Range("checkin92E").value = "Y" Then
'    If Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "14" Or Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "37" Then
'        Altduedate = "31/03/2020"
'  Else
'        Altduedate = "30/11/2025" 'dpk1601
        Altduedate = Sheet33.Range("G6").value
'    End If
ElseIf (CheckAuditFLAG Or RepSec92EOtherFlag_2 Or CheckAuditFLAGIF) Then
'ElseIf (CheckAuditFLAG Or RepSec92EOtherFlag_2 Or CheckAuditFLAGIF Or Sheet29.Range("BooksSpouse44ABFlg").value = "Yes") Then
'  If Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "14" Or Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "37" Then
'        Altduedate = "31/03/2020"
'  Else
'        Altduedate = "31/10/2025" 'dpk1601
        Altduedate = Sheet33.Range("G5").value
'    End If
Else
'    If Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "14" Or Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "37" Then
'    Altduedate = "31/03/2020"
'    Else
    'Altduedate = "31/07/2025" 'dpk1601
    'Malli_08/09/2025
   ' Altduedate = "15/09/2025" '15/07/2025
     Altduedate = Sheet33.Range("Duedate_Extended").value
    '----------------------
'    End If
End If

Returnfiledstatus = Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))
VerificationDate = Range("sheet9.Date").value
 If Not ChkMinInclusiveDate(Dformat(VerificationDate, "yyyy-mm-dd"), todaysdate1) Then
                VerificationDate = todaysdate1
 End If
dateOfFiling = Range("sheet1.OrigRetFiledDate").value

DueDate = Dformat(Altduedate, "yyyy-mm-dd")
If VerificationDate <> todaysdate1 Then
VerificationDate1 = Dformat(VerificationDate, "yyyy-mm-dd")
Else
VerificationDate1 = VerificationDate
End If
DateOfFiling1 = Dformat(dateOfFiling, "yyyy-mm-dd")


If Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or UpdatedY Then
VerificationDate1 = Dformat(dateOfFiling, "yyyy-mm-dd")
ElseIf VerificationDate <> todaysdate1 Then
VerificationDate1 = Dformat(VerificationDate, "yyyy-mm-dd")
Else
VerificationDate1 = VerificationDate
End If

If (Returnfiledstatus = "14" Or Returnfiledstatus = "15" Or Returnfiledstatus = "16" Or Returnfiledstatus = "20" Or Returnfiledstatus = "19") Or (VerificationDate1 <= DueDate) Then  'Or (Range("Sheet8b.TotalIncome").value <= 0) Then
' If (Returnfiledstatus = "14" Or Returnfiledstatus = "15" Or Returnfiledstatus = "16" Or Returnfiledstatus = "20" Or Returnfiledstatus = "19") Or (VerificationDate1 <= DueDate Or Range("Sheet8b.LossesOfCurrentYearCarriedFwd").value > 0) Then  'Or (Range("Sheet8b.TotalIncome").value <= 0) Then
 intrstPayUs234F = 0

ElseIf (Returnfiledstatus = "13" Or Returnfiledstatus = "11" Or Returnfiledstatus = "12" Or UpdatedN) And (Range("Sheet8b.TotalIncome").value <= 500000) Then  'And Range("Sheet8b.TotalIncome").value > 0) Then
intrstPayUs234F = 1000

ElseIf (Returnfiledstatus = "13" Or Returnfiledstatus = "11" Or Returnfiledstatus = "12" Or UpdatedN) And Range("Sheet8b.TotalIncome").value > 500000 And (VerificationDate1 > DueDate) And (VerificationDate1 <= "2026-12-31") Then 'dpk1601 Year changed 2025 to 2026 by sadineni
intrstPayUs234F = 5000
'Else
'intrstPayUs234F = 10000
'Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
'End If

'ElseIf (Returnfiledstatus = "17" Or UpdatedY Or Returnfiledstatus = "18" Or Range("Sheet8b.LossesOfCurrentYearCarriedFwd").value > 0) And (DateOfFiling1 <= DueDate) Then
ElseIf (Returnfiledstatus = "17" Or UpdatedY Or Returnfiledstatus = "18") And (DateOfFiling1 <= DueDate) Then
intrstPayUs234F = 0

ElseIf (Returnfiledstatus = "17" Or UpdatedY Or Returnfiledstatus = "18") And Range("Sheet8b.TotalIncome").value <= 500000 Then
intrstPayUs234F = 1000
ElseIf (Returnfiledstatus = "17" Or UpdatedY Or Returnfiledstatus = "18") And Range("Sheet8b.TotalIncome").value > 500000 And (DateOfFiling1 > DueDate) And (DateOfFiling1 <= "2026-12-31") Then 'dpk1601 'Year 2025 to 2026 changed by sadineni on 10/02/2026
intrstPayUs234F = 5000
Else
intrstPayUs234F = 5000
End If

sPassword = EfilingCommon.getmsgstate

Sheet7.Unprotect Password:=sPassword
Sheet7.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
Sheet7.Protect Password:=sPassword

End Function

Function setTblinfo_AuditInfo_Front()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Audit.Act").count
    Set rangecells = Sheet1.Range("Audit.Act").Cells
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    setTblinfo_AuditInfo_Front = ccount
    
End Function


