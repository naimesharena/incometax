Attribute VB_Name = "mdTaxCalc"
Option Explicit

Dim sPassword As String
Dim Date_9 As Variant
Dim Date_D As Variant

Dim altduedate As String
Dim MsgCalcInterest As Variant
Dim IntrstPayUs234A As Double
Dim delayedInMonths As Double
Dim taxbase234A As Double
Dim calcInterestPayable234A As Double
Const CONST_IntrstPay234A_Percentage As Double = 1
Dim taxbase234B As Double
Dim IntrstPayUs234B As Double
Dim calcInterestPayable234B As Double
Public CONST_IntrstPay234B_Percentage_1, April_1, may_1, June_1 As Variant
Const CONST_IntrstPay234B_Percentage As Double = 1
Const CONST_NET_Limit As Double = 10000
Const CONST_ATP_Limit As Double = 90
Dim IntrstPayUs234C As Double
Dim intrstPayUs234F As Double
Dim shortFall As Double
Dim OrigRetFiledDate_1 As Variant
Dim OrigRetFiledDate_11 As Variant

Dim GrpBIncome1   As Variant
Dim GrpBIncome1RATE   As Variant
Dim GrpBIncome2   As Variant
Dim GrpBIncome2RATE   As Variant

Dim SI112LTCGSecurities   As Variant
Dim SI112LTCGSecuritiesRATE   As Variant
Dim SI111ASTCG   As Variant
Dim SI111ASTCGRATE   As Variant
Dim SI112LTCGOthers   As Variant
Dim SI112LTCGOthersRATE   As Variant
Dim SI115BBLotteries   As Variant
Dim SI115BBLotteriesRATE   As Variant
Public age As Long
Dim ageStatus As String



Sub ComputeInterest()
On Error GoTo endline
     Dim Status_3 As String
     
     
     
     
    Status_3 = Sheet1.Range("sheet1.Status").value
    Status_3 = Mid(Status_3, 1, 1)
    InitProgBar
    ProgressFrameCaption = "Tax Calculation"
    mainProcCaption = "Validating Sheets"
    noOfProcessMain = 7
    UserForm1.Show vbModeless
    
    
    UpdateProgressBar
    ValidateSheetIncomeDetails
    
    ValidateSheetHouseProperty  'Konda Added 23-01-2026
    
    If Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" Then

        UpdateProgressBar
        Validate_Gen_1398A
    End If
    UpdateProgressBar
    ValidateSheetNOBBP
    UpdateProgressBar
    ValidateTDS_TCS_IT
    
    UpdateProgressBar
    Validate_IT
    
    UpdateProgressBar
    Validate_TCS
    
'    UpdateProgressBar
'    ValidateSheetSchTDS2
'
'    UpdateProgressBar
'    ValidateSheetSchTDS3

    If (Sheet8.Range("bacValue").value) = 2 Then
        UpdateProgressBar
        ValidateSheet80G
        
        UpdateProgressBar
        Validate80D_All
        
        UpdateProgressBar
        Validate80GGC
        
        UpdateProgressBar
        Validate80DD
        
        UpdateProgressBar
        Validate80U
    
'Newly added by sai on 25/04/2025===============
        UpdateProgressBar
        Validate80C__All
        
'        UpdateProgressBar
'        Validate80C_2_All

    If Sheet1.Range("sheet1.Status").value = "I - INDIVIDUAL" Then
            
            UpdateProgressBar
            Validate80E_All
            
            UpdateProgressBar
            Validate80EE_All
            
            UpdateProgressBar
            Validate80EEA_All
            
            UpdateProgressBar
            Validate80EEB_All
    End If
        
'        UpdateProgressBar
'        Validate24b_All
    '===================================================

    End If
    
    'Malli--------SIT_SIT-93261
    If Not (Sheet8.Range("BacValue").value = 1 And Mid(Sheet1.Range("IncD.TypeOfHP").value, 1, 1) = "S") Then
    UpdateProgressBar
        Validate24b_All
    End If
    '--------------------------
    
     'Malli------08/05/2025
    If Sheet8.Range("BacValue").value = 2 And Left(Sheet1.Range("sheet1.Status").value, 1) = "I" And Mid(Range("sheet1.NatureofEmployment").value, 1, 1) <> "N" Then
     ValidateEA10_13A
    End If
    '-------------------
    
    UpdateProgressBar
    ValidateSheetTaxPaidVerificaton
    UpdateProgressBar
    ValidateBA
    
   UpdateProgressBar
'    ValidateDI
    

        If Sheet5.Range("IncD.BalTaxPayable").value > 0 Then
        Sheet5.Shapes("EPAYButton").Visible = True
          Else
         Sheet5.Shapes("EPAYButton").Visible = False
    End If
    
    'If Status_3 <> "F" Then
'        ValidateSchAL
    'End If
'    If (Sheet5.Range("Sheet5.CTcheck").value = "YES") Then
'    ValidateSheetCT
'    End If
    
    'ValidateSheetGC
    
    
    Dim SysCalculatedNetTaxLiability As Double
    Dim matchedAdvanceTax As Double
    Dim tdsamtused As Double
    Dim tcsamtused As Double
    Dim IncSal, IncHp, IncOS, Inc44AD, Inc44ADA, Inc44AE As Variant
    Dim IsSeniorCitizen, IsSuperSeniorCitizen, IsNRI As Variant
    
    sPassword = mdCommon.getmsgstate
    Sheet9.Unprotect Password:=sPassword

    IntrstPayUs234A = 0
    IntrstPayUs234B = 0
    IntrstPayUs234C = 0
    intrstPayUs234F = 0
        Sheet9.Unprotect Password:=sPassword
    Sheet9.Range("CalculatedValue234A").value = IntrstPayUs234A

    Sheet9.Range("CalculatedValue234C").value = IntrstPayUs234C

    Sheet9.Range("CalculatedValue234B").value = IntrstPayUs234B
    'Sheet9.Range("CalculatedValue234F").value = intrstPayUs234F
    
    filingdate
        
    SysCalculatedNetTaxLiability = Sheet1.Range("IncD.NetTaxLiability").value

    matchedAdvanceTax = Sheet5.Range("IncD.AdvanceTax").value
    tdsamtused = Sheet5.Range("IncD.TDS").value
    tcsamtused = Sheet5.Range("IncD.TCS").value
    
    
    IncSal = Sheet1.Range("IncD.IncomeFromSal").value
    IncHp = Sheet1.Range("IncD.IncomeFromHP").value
    IncOS = Sheet1.Range("IncD.IncomeFromOS").value
    
    Inc44AD = Sheet3.Range("BP_E2_PI").value
    Inc44ADA = Sheet3.Range("BP_E4_PI44ADA").value
    Inc44AE = Sheet3.Range("BP_E3").value
    
    IsSeniorCitizen = Sheet9.Range("GrpA").value
    IsSuperSeniorCitizen = Sheet9.Range("GrpB").value
    IsNRI = Sheet9.Range("IS_NRI").value
    


     '234F Calculation


age = mdCommon.calculateAge(Sheet1.Range("sheet1.DOB").value)
If (age <= 59) Then
ageStatus = "NC"
ElseIf ((age > 59) And (age <= 79)) Then
ageStatus = "SC"
Else
ageStatus = "SSC"
End If

'If (Range("IncD.GrossTotIncome").value) < 0 Then
'    New234F
If Mid(Sheet1.Range(" sheet1.SeventhProvisoFlag"), 1, 1) = "Y" Then
New234F
Else
'If (Mid(Sheet1.Range("sheet1.Status"), 1, 1) = "H" And (Mid(Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR" Or Mid(Range("sheet1.ResidentialStatus1").value, 1, 1) = "R")) Then

' If ((Range("IncD.GrossTotIncome").value + Range("Sec38Amt").value) > 250000) Then


'ElseIf (Mid(Sheet1.Range("sheet1.Status"), 1, 1) = "I" And Mid(Range("sheet1.ResidentialStatus1").value, 1, 2) = "NR") Then
'If ((Range("IncD.GrossTotIncome").value + Range("Sec38Amt").value) > 250000) Then

If (Mid(Sheet1.Range("sheet1.Status"), 1, 1) = "H" Or Mid(Sheet1.Range("sheet1.Status"), 1, 1) = "I") And (Sheet8.Range("bacValue").value = 1) Then
    'If ((Range("IncD.GrossTotIncome").value + Range("Sec38Amt").value) > 300000) Then ''For 234F calculation need to consider GTI in which 112A Included -- updated by Bindu as per BA/DE V3 on 4th Feb 2025
'    If ((Range("IncD.GrossTotIncome_New").value + Range("Sec38Amt").value) > 300000) Then
     If ((Range("IncD.GrossTotIncome_New").value + Range("Sec38Amt").value) > 400000) Then  'Ayush_29/01/2026
       
       New234F
    End If
    ElseIf (Mid(Sheet1.Range("sheet1.Status"), 1, 1) = "H") And (Sheet8.Range("bacValue").value = 2) Then
    'If ((Range("IncD.GrossTotIncome").value + Range("Sec38Amt").value) > 250000) Then ''For 234F calculation need to consider GTI in which 112A Included -- updated by Bindu as per BA/DE V3 on 4th Feb 2025
     If ((Range("IncD.GrossTotIncome_New").value + Range("Sec38Amt").value) > 250000) Then
      New234F
    End If

'chandru
'    ElseIf ((Mid(Sheet1.Range("sheet1.Status"), 1, 1) = "I") And (Sheet8.Range("bacValue").value = 1)) Then
'    If ((Range("IncD.GrossTotIncome").value + Range("Sec38Amt").value) > 250000) Then
'      New234F
'    End If
    

ElseIf (Mid(Sheet1.Range("sheet1.Status"), 1, 1) = "F") Then
    New234F

End If

' If ageStatus = "NC" And ((Range("IncD.GrossTotIncome").value + Range("Sec38Amt").value) > 250000) Then
'    New234F
'    ElseIf ageStatus = "SC" And ((Range("IncD.GrossTotIncome").value + Range("Sec38Amt").value) > 300000) Then
'    New234F
'    ElseIf ageStatus = "SSC" And ((Range("IncD.GrossTotIncome").value + Range("Sec38Amt").value) > 500000) Then
'    New234F



'If ageStatus = "NC" And ((Range("IncD.GrossTotIncome").value + Range("Sec38Amt").value) > 250000) Then
   
   
   
If (Mid(Sheet1.Range("sheet1.Status"), 1, 1) = "I") Then
   
   
   If Sheet8.Range("bacValue").value = 2 Then
   
       'If ageStatus = "NC" And ((Range("IncD.GrossTotIncome").value + Range("Sec38Amt").value) > 300000) And Sheet8.Range("bacValue").value = 1 Then
        'New234F
       'If ageStatus = "NC" And ((Range("IncD.GrossTotIncome").value + Range("Sec38Amt").value) > 250000) And Sheet8.Range("bacValue").value = 2 Then ''For 234F calculation need to consider GTI in which 112A Included -- updated by Bindu as per BA/DE V3 on 4th Feb 2025
       If ageStatus = "NC" And ((Range("IncD.GrossTotIncome_New").value + Range("Sec38Amt").value) > 250000) And Sheet8.Range("bacValue").value = 2 Then
        New234F
        
       'ElseIf ageStatus = "SC" And ((Range("IncD.GrossTotIncome").value + Range("Sec38Amt").value) > 300000) Then ''For 234F calculation need to consider GTI in which 112A Included -- updated by Bindu as per BA/DE V3 on 4th Feb 2025
       ElseIf ageStatus = "SC" And ((Range("IncD.GrossTotIncome_New").value + Range("Sec38Amt").value) > 300000) Then
        New234F
       'ElseIf ageStatus = "SSC" And ((Range("IncD.GrossTotIncome").value + Range("Sec38Amt").value) > 500000) Then ''For 234F calculation need to consider GTI in which 112A Included -- updated by Bindu as per BA/DE V3 on 4th Feb 2025
       ElseIf ageStatus = "SSC" And ((Range("IncD.GrossTotIncome_New").value + Range("Sec38Amt").value) > 500000) Then
        New234F
    
       Else
            intrstPayUs234F = 0
       End If
       
    End If

End If


End If



'Malli------------27/10/2025
'Application.EnableEvents = False
        Sheet9.Unprotect Password:=mdCommon.getmsgstate
        Sheet9.Range("CB2").value = intrstPayUs234F
        'Sheet9.Protect Password:=mdCommon.getmsgstate

'Application.EnableEvents = True
'-----------------







 


'Dim Returnfiledstatus As String
'Dim VerificatonDate As Variant
'Dim DateOfFiling As Variant
'
''Const FilingLastDate As Date = #1/8/2018#
''Const YearEndDate As Date = #12/31/2018#
'Dim dat As Variant
'Dim mont As Variant
'Dim yea As Variant
'Dim dat1 As Variant
'Dim mont1 As Variant
'Dim yea1 As Variant
'
'Returnfiledstatus = Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))
'VerificatonDate = Range("Ver.Date").value
'DateOfFiling = Range("sheet1.OrigRetFiledDate").value
'
'dat = Mid(VerificatonDate, 1, 2)
'mont = Mid(VerificatonDate, 4, 2)
'yea = Mid(VerificatonDate, 7, 4)
'
'dat1 = Mid(DateOfFiling, 1, 2)
'mont1 = Mid(DateOfFiling, 4, 2)
'yea1 = Mid(DateOfFiling, 7, 4)
'
'If (Returnfiledstatus = "13" Or Returnfiledstatus = "14" Or Returnfiledstatus = "15" Or Returnfiledstatus = "16" Or Returnfiledstatus = "20") Or (dat <= "31" And mont <= "07" And yea <= "2018") Then
'intrstPayUs234F = 0
''Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
'ElseIf (Returnfiledstatus = "11" Or Returnfiledstatus = "12") And Range("IncD.TotalIncome").value <= 500000 Then
'intrstPayUs234F = 1000
''Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
'ElseIf (Returnfiledstatus = "11" Or Returnfiledstatus = "12") And Range("IncD.TotalIncome").value > 500000 And (dat <= "31" And mont >= "07" And yea = "2018") And (dat <= "31" And mont <= "12" And yea = "2018") Then
'intrstPayUs234F = 5000
''Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
''Else
''intrstPayUs234F = 10000
''Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
''End If
'ElseIf (Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or Returnfiledstatus = "19") And (dat1 <= "31" And mont1 <= "07" And yea1 <= "2018") Then
'intrstPayUs234F = 0
''Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
'ElseIf (Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or Returnfiledstatus = "19") And Range("IncD.TotalIncome").value <= 500000 Then
'intrstPayUs234F = 1000
''Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
'ElseIf (Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or Returnfiledstatus = "19") And Range("IncD.TotalIncome").value > 500000 And (dat1 <= "31" And mont1 >= "07" And yea1 = "2018") And (dat1 <= "31" And mont1 <= "12" And yea1 = "2018") Then
'intrstPayUs234F = 5000
''Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
'Else
'intrstPayUs234F = 10000
''Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
'End If
'
'
''Sheet9.Range("CalculatedValue234F").value = Round(intrstPayUs234F)
        
'''''''''''''''''''''''''''''''''''''''''''''Interest234C - Start''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
    
    Inc44AD = IIf(Inc44AD = "", 0, Inc44AD)
    Inc44ADA = IIf(Inc44ADA = "", 0, Inc44ADA)
    IncHp = IIf(IncHp = "", 0, IncHp)
    Inc44AE = IIf(Inc44AE = "", 0, Inc44AE)
    IncOS = IIf(IncOS = "", 0, IncOS)
       


    If (SysCalculatedNetTaxLiability - (tdsamtused + tcsamtused)) >= CONST_NET_Limit Then
        IntrstPayUs234C = Sheet9.Range("interest234C").value
    Else
        IntrstPayUs234C = 0
    End If
    
    If UCase(IsNRI) = "FALSE" Then
                 If (Inc44AD + Inc44ADA) > 0 Then
                    'SysCalculatedNetTaxLiability = Sheet9.Range("AssessedTax_2").value
                     'Newly updated Bindu
                     'If (SysCalculatedNetTaxLiability) >= CONST_NET_Limit Then
                     If (SysCalculatedNetTaxLiability - (tdsamtused + tcsamtused)) >= CONST_NET_Limit Then
                        IntrstPayUs234C = Sheet9.Range("interest234C_1").value
                        
                    Else
                        IntrstPayUs234C = 0
                    End If
                 End If
    End If
       
    'If UCase(IsNRI) = "FALSE" Then
    '    If Inc44AD > 0 Then
    '        If (IncSal = 0 And IncHp = 0 And IncOS = 0 And Inc44AE = 0) Then intrstPayUs234C = 0
    '    End If
    'End If
    
    Sheet9.Range("CalculatedValue234C").value = Round(IntrstPayUs234C)

'''''''''''''''''''''''''''''''''''''''''''''Interest234C - Done''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        
           
'''''''''''''''''''''''''''''''''''''''''''''Interest234B - Start''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

     
    If UCase(IsNRI) = "FALSE" Then
                 If (Inc44AD + Inc44ADA) > 0 Then
                    'SysCalculatedNetTaxLiability = Sheet9.Range("AssessedTax_1").value
                 End If
        'If Inc44AD > 0 Then
        '    If (IncSal = 0 And IncHp = 0 And IncOS = 0 And Inc44AE = 0) Then intrstPayUs234B = 0
        'End If
    End If
         
     Call Calculate_InterestPayable234A(SysCalculatedNetTaxLiability, _
                                                     matchedAdvanceTax, _
                                                     tdsamtused, _
                                                     tcsamtused)
       
                                                     
    IntrstPayUs234B = Calculate_InterestPayable234B(SysCalculatedNetTaxLiability, _
                                                     matchedAdvanceTax, _
                                                     tdsamtused, _
                                                     tcsamtused, _
                                                     IntrstPayUs234A, _
                                                     IntrstPayUs234C, _
                                                     intrstPayUs234F)
                                                        
    'If UCase(IsNRI) = "FALSE" Then
    '    If Inc44AD > 0 Then
    '        If (IncSal = 0 And IncHp = 0 And IncOS = 0 And Inc44AE = 0) Then intrstPayUs234B = 0
    '    End If
    'End If
    
    Sheet9.Unprotect Password:=mdCommon.getmsgstate
    Sheet9.Range("CalculatedValue234B").value = Round(IntrstPayUs234B)
    Sheet9.Protect Password:=mdCommon.getmsgstate

'''''''''''''''''''''''''''''''''''''''''''''Interest234B - Done''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'Added by Riyaz on 07/01/2026
''''''''''''''''''''''''''''''''''''''''''''''Interest234A - Start''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    
    IntrstPayUs234A = 0
    Dim i As Long
    If delayedInMonths = 0 Then
        IntrstPayUs234A = 0
    Else
        For i = 1 To delayedInMonths
        
            IntrstPayUs234A = IntrstPayUs234A + Sheet9.Range("CG" & i + 9).value
        
        Next i
    End If
    
    
    Sheet9.Unprotect Password:=mdCommon.getmsgstate
    Sheet9.Range("CalculatedValue234A").value = Round(IntrstPayUs234A)
    Sheet9.Range("section234A").value = Round(IntrstPayUs234A)
    Sheet9.Protect Password:=mdCommon.getmsgstate

''''''''''''''''''''''''''''''''''''''''''''''Interest234A - Done''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''Malli -------------27/10/2025
'
''    Calculate_234A
'    Sheet9.Unprotect Password:=mdCommon.getmsgstate
'    Sheet9.Range("CalculatedValue234A").value = Round(IntrstPayUs234A)
'    Sheet9.Range("section234A").value = Round(IntrstPayUs234A)
'    Sheet9.Protect Password:=mdCommon.getmsgstate
''------------------------------



    Dim income_BP As Variant
    income_BP = Sheet3.Range("BP_E2_PI").value + IIf(Sheet3.Range("BP_E4_PI44ADA").value = "", 0, Sheet3.Range("BP_E4_PI44ADA").value) + Sheet3.Range("BP_E3").value
    
    If UCase(IsNRI) = "FALSE" And Status_3 = "I" And Range("GrpC") = True And income_BP = 0 Then
        IntrstPayUs234B = 0
        IntrstPayUs234C = 0
    End If



    sPassword = mdCommon.getmsgstate
    Sheet1.Unprotect Password:=sPassword
    Sheet1.Range("IncD.IntrstPayUs234A").value = Round(IntrstPayUs234A)
    Sheet1.Unprotect Password:=sPassword
    Sheet1.Range("IncD.IntrstPayUs234B").value = Round(IntrstPayUs234B)
    Sheet1.Unprotect Password:=sPassword
    Sheet1.Range("IncD.IntrstPayUs234C").value = Round(IntrstPayUs234C)
    Sheet1.Unprotect Password:=sPassword
    Sheet1.Range("IncD.IntrstPayUs234F").value = Round(intrstPayUs234F)
    Sheet1.Protect Password:=sPassword
    
'    If Sheet5.Range("IncD.BalTaxPayable").value > 0 Then
'        Sheet1.Range("sheet1.TaxStatus").value = "Tax Payable"
'        ElseIf Sheet5.Range("IncD.RefundDue").value > 0 Then
'            Sheet1.Range("sheet1.TaxStatus").value = "Tax Refundable"
'
'        Else
'        Sheet1.Range("sheet1.TaxStatus").value = "Nil Tax Payable"
'        End If
        
        
      
        CalculateTaxCheckFlag = True
      
        
endline:
ProgressBarHide


'Ayush_UR
If UCase(Sheet1.Range("sheet1.ReturnFileSec").value) = UCase("139(8A)") Then
        If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").value) = UCase("No") Then
        If Sheet202.Range("U_AmtPayable") <= 0 Then
           MsgBox "If A5 in Part A gen 139(8A)  is ""No"",then updated return can't be filed if there is no payment u/s 140B", vbCritical, "Error"
          
           CalculateTaxCheckFlag = False
           
           
           Exit Sub
        End If
    End If
    
    If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").value) = UCase("No") Then
        If Sheet202.Range("U_AmtRefundable") > 0 Then
           MsgBox "If A5 in Part A gen 139(8A)  is ""No"",then updated return can't be filed with Refund", , "Error"
          
           CalculateTaxCheckFlag = False
           
           Exit Sub
        End If
    End If
    
    
     If Sheet202.Range("U_NetPayable") <= 0 Then
       MsgBox "Since you are filing Updated return, you are expected to have ""Net amount payable"" more than 0 to be able to pay tax u/s 140B.", , "Error"
       
       
    End If
    
    
    
    
    If Sheet202.Range("U_TaxUS140B") = 0 Then
       MsgBox """Tax Paid u/s 140B at sl.no.12 should be more than ""0"" to file ITR u/s 139(8A)""", vbInformation, "Error"
    End If
    
     If Sheet202.Range("U_TaxUS140B") > Sheet202.Range("U_NetPayable").value Then
       MsgBox "Tax paid u/s 140B is more than ""sl.no.11.Net amount payable"". Hence your updated return may be invalid as it appears to result in refund as provided in clause (c) to first proviso of 139(8A).", vbInformation, "Error"
    End If
End If
    
    
    
   Sheet9.Protect Password:=sPassword
   Sheet1.Activate
End Sub

Function Calculate_InterestPayable234A(SysCalculatedNetTaxLiability As Double, matchedAdvanceTax As Double, _
                                    tdsamtused As Double, tcsamtused As Double) As Double

    Dim retfilestatus As String
    Dim origrevisedstatus As String
    Dim origreturnfileddate As String
    Dim revisedreturnfilingdate As String
    Dim oldrevisedreturnfilingdate As String
    Dim origreturnfileddate2 As String
    Dim revisedreturnfilingdate2 As String
    
    Dim AuditFLAG As String
    Dim RepSec92EFlag_2 As Variant
    Dim RepSec92EOtherFlag_2 As Variant
    
        ''''''''''''
    Dim startDate, currentdate As Variant
    Dim dueDate As Date
    Dim userFilingDate As Date
    Dim delayedMonths As Integer
    Dim Message As String
    
    Date_D = Sheet5.Range("Ver.Date").value
    
    dueDate = DateSerial(2025, 9, 22)
    userFilingDate = val(Mid(startDate, 7, 4))
    delayedMonths = val(Mid(currentdate, 1, 2))
    
    ''''''''''''''
    '--nEW
    Dim i As Long
    
    
    Dim UpdatedY As Boolean
    UpdatedY = False
    If (((Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))) = "21") And (Mid(Range("U_PreviouslyFiledForThisAY"), 1, 1) = "Y")) Then
    UpdatedY = True
    End If
    
    If Mid(Sheet1.Range("sheet1.ReturnFileSec1"), 1, 2) = "17" Or UpdatedY Then

        origreturnfileddate = Sheet1.Range("sheet1.OrigRetFiledDate").value
    End If
    
    
    retfilestatus = Mid(Sheet1.Range("sheet1.ReturnFileSec1"), 1, 2)

    If retfilestatus = "18" Or retfilestatus = "17" Or UpdatedY Then
        
        origrevisedstatus = "R"
    End If

    If (retfilestatus = "17" Or retfilestatus = "18" Or UpdatedY) Then
        revisedreturnfilingdate = Sheet9.Range("DateOfFiling234A").value
    Else
        revisedreturnfilingdate = Sheet9.Range("DateOfFiling").value
    End If

    origreturnfileddate2 = Dformat(origreturnfileddate, "yyyy-mm-dd")
    revisedreturnfilingdate2 = Dformat(revisedreturnfilingdate, "yyyy-mm-dd")
    oldrevisedreturnfilingdate = revisedreturnfilingdate
      
    If origreturnfileddate2 <> "" Then
        If Not ChkMinInclusiveDate(origreturnfileddate2, revisedreturnfilingdate2) Then
            revisedreturnfilingdate = origreturnfileddate
        End If
    End If
    If (SysCalculatedNetTaxLiability - matchedAdvanceTax - tdsamtused - tcsamtused) > 100000 Then
    altduedate = Sheet9.Range("DueDate_Extend").value   'Bindu 8th Sep 25

    Else
    
    altduedate = Sheet9.Range("DueDate_Extend").value   'Bindu 8th Sep 25

    End If

    If (AuditFLAG = "Yes") Then
        If (RepSec92EFlag_2 Or RepSec92EOtherFlag_2) Then
            If RepSec92EFlag_2 Then
                delayedInMonths = MonthDiff(revisedreturnfilingdate, Sheet9.Range("DueDate_Extend").value) 'Bindu on 8th Sep 25
            
            Else
                delayedInMonths = MonthDiff(revisedreturnfilingdate, altduedate) 'Changes by Bindu on 8th Sep 25
            
            End If
        End If
    Else
            delayedInMonths = MonthDiff(revisedreturnfilingdate, altduedate)
    End If
    
End Function

Function Calculate_InterestPayable234B(SysCalculatedNetTaxLiability As Double, _
                                        matchedAdvanceTax As Double, _
                                        tdsamtused As Double, _
                                        tcsamtused As Double, _
                                        intrstPayUs234A_1 As Double, _
                                        intrstPayUs234C_1 As Double, _
                                        intrstPayUs234F_1 As Double) As Double
        

                    
  '*/ Malli_27/10/2025>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  
  
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
        Dim j As Long
        Dim looperint, looper As Long
        Dim SATPaidAtlooper As Double
        
        looperint = 0
        looper = 0
                
        Dim DueDate_1 As String
        DueDate_1 = Sheet9.Range("DueDate_Extend").value
        

        
        sPassword = mdCommon.getmsgstate
  
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
                          
                If (shortFall > 100) Then
                    shortFall = WorksheetFunction.Floor((shortFall / 100), 1) * 100
                End If
                          
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
        
        dateOfProcessing = Sheet9.Range("DateOfFiling").value
        'Ayush_26-27
'        AssYear = 2025
        AssYear = 2026
        dateOfProcessing = Dformat(dateOfProcessing, "yyyy-mm-dd")
        
        yrdop = CInt(Mid(dateOfProcessing, 1, 4))
        mthdop = CInt(Mid(dateOfProcessing, 6, 2))
        
        If (yrdop >= AssYear) Then
            calcIntrst234BUptoPeriod = mthdop - 4 + (yrdop - AssYear) * 12
        End If
        
        calcIntrst234BUptoPeriod = calcIntrst234BUptoPeriod + 1
        
        carryForwardPrinicipal = shortFall
         
        
         
Application.EnableEvents = False

        Sheet9.Unprotect Password:=mdCommon.getmsgstate
        Sheet9.Range("Month_Wise_234Bvalues").ClearContents
        Sheet9.Range("Balance_Interest").ClearContents
        Sheet9.Range("IT_Interest").ClearContents
        Sheet9.Protect Password:=mdCommon.getmsgstate
        
Application.EnableEvents = True
        
        Dim month_count_1 As Variant
        month_count_1 = (CLng(Mid(Dformat(DueDate_1, "yyyy-mm-dd"), 6, 2)) - 2)
        
        For i = 1 To calcIntrst234BUptoPeriod
        
                SATPaidAtPeriod = 0
                SATPaidAtlooper = 0
                balancePrincipal = carryForwardPrinicipal

                calcIntrst234BOnPeriod = Math.Round(CONST_IntrstPay234B_Percentage / 100 * WorksheetFunction.RoundDown(balancePrincipal, -2))
        'Ankita_07/05/2026
'                If (age > 59) Then
'                    calcIntrst234BOnPeriod = 0
'                    Sheet9.Unprotect Password:=getmsgstate
'                    Sheet9.Range("Month_Wise_234Bvalues").Cells(i, 1).value = calcIntrst234BOnPeriod
'                    Sheet9.Protect Password:=getmsgstate
'                Else
                    Sheet9.Unprotect Password:=getmsgstate
                    Sheet9.Range("Month_Wise_234Bvalues").Cells(i, 1).value = calcIntrst234BOnPeriod
                    Sheet9.Protect Password:=getmsgstate
                    
                    calcInterestPayable234B = calcInterestPayable234B + calcIntrst234BOnPeriod
'                End If
                
'                Sheet9.Unprotect Password:=mdCommon.getmsgstate
'                Sheet9.Range("Month_Wise_234Bvalues").Cells(i, 1).value = calcIntrst234BOnPeriod
'                Sheet9.Protect Password:=mdCommon.getmsgstate
'
'                calcInterestPayable234B = calcInterestPayable234B + calcIntrst234BOnPeriod
'
'                If i = 1 Then
'                    balanceInterest = Sheet9.Range("CG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234C_1
'                ElseIf i = month_count_1 And DueDate_1 = Sheet9.Range("DueDate_Extend").value Then
'                    balanceInterest = Sheet9.Range("CG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234F_1
'                Else
'                    balanceInterest = Sheet9.Range("CG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod
'                End If
'Ayush_23/04
 If Mid(Sheet1.Range("sheet1.ReturnFileSec1").value, 1, 2) = 17 Or Mid(Sheet1.Range("sheet1.ReturnFileSec1").value, 1, 2) = 18 Or (((Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))) = "21") And (Mid(Range("U_PreviouslyFiledForThisAY"), 1, 1) = "Y")) Then
 
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
 '--------------------
                
                
            If delayedInMonths = 0 Then
                    If i = 1 Then
                        balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234C_1
                    Else
                        balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                    End If
                Else
                    If i = 1 Then
                        balanceInterest = Sheet9.Range("CG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234C_1
                     ElseIf i = month_count_1 And DueDate_1 = Sheet9.Range("DueDate_Extend").value Then
                        balanceInterest = Sheet9.Range("CG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234F_1
                    ElseIf i <= Revisedmonthdiff Then
                        balanceInterest = Sheet9.Range("CG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod
                    Else
                        balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                    End If
                End If
                
        Else
                If i = 1 Then
                    balanceInterest = Sheet9.Range("CG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234C_1
                 ElseIf i = month_count_1 And DueDate_1 = Sheet9.Range("DueDate_Extend").value Then
                    balanceInterest = Sheet9.Range("CG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234F_1
                Else
                    balanceInterest = Sheet9.Range("CG" & i + 4).value + carryForwardInterest + calcIntrst234BOnPeriod
                End If
        End If
'
                Sheet9.Unprotect Password:=mdCommon.getmsgstate
                Sheet9.Range("Balance_Interest").Cells(i, 1).value = balanceInterest
                Sheet9.Protect Password:=mdCommon.getmsgstate
                
                
                
'Added by Riyaz on 10/12/2025

'-------------Commented by Aavula Naresh on 07/01/2026 -------------

'                Dim SATPaidAtPeriod1 As Double
'                Dim SATPaidAtPeriod2 As Double
'
'                SATPaidAtPeriod1 = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & i, Range("TaxP.Amt"))
'
'                If Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" Then
'                    SATPaidAtPeriod2 = Application.WorksheetFunction.SumIf(Range("U_DateOfAmount1"), "=" & i, Range("U_Amt1"))
'                Else
'                    SATPaidAtPeriod2 = 0
'                End If
'
'                SATPaidAtPeriod = SATPaidAtPeriod1 + SATPaidAtPeriod2
           
'            Dim SATPaidAtPeriod As Double
'--------------------------------------------------------------------
'Added by Aavula Naresh on 07/01/2026

                SATPaidAtPeriod = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & i, Range("TaxP.Amt"))
           
                
                Sheet9.Unprotect Password:=mdCommon.getmsgstate
                Sheet9.Range("IT_Interest").Cells(i + 1, 1).value = SATPaidAtPeriod
                Sheet9.Protect Password:=mdCommon.getmsgstate
                
                
                adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
             
        Next
        
    End If
    
        
    Calculate_InterestPayable234B = WorksheetFunction.Round(calcInterestPayable234B, 0)
  '------------------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
End Function
'Malli renamed function name is as OLD 06/11/2025
Function Calculate_InterestPayable234B_OLD(SysCalculatedNetTaxLiability As Double, _
                                        matchedAdvanceTax As Double, _
                                        tdsamtused As Double, _
                                        tcsamtused As Double, _
                                        intrstPayUs234A_1 As Double, _
                                        intrstPayUs234C_1 As Double, _
                                        intrstPayUs234F_1 As Double) As Double
        
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
        Dim j As Long
        Dim looperint, looper As Long
        Dim SATPaidAtlooper As Double
        
        looperint = 0
        looper = 0
                
        
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
                          
                If (shortFall > 100) Then
                    shortFall = WorksheetFunction.Floor((shortFall / 100), 1) * 100
                End If
                          
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
        
        dateOfProcessing = Sheet9.Range("DateOfFiling").value
        
        'assYear = 2022 dpk1201
'        assYear = 2023
'Ayush_13
'        assYear = 2024 'AY 2024-25 Change
         AssYear = 2025 'AY 2025-26 Change
        dateOfProcessing = Dformat(dateOfProcessing, "yyyy-mm-dd")
        
        yrdop = CInt(Mid(dateOfProcessing, 1, 4))
        mthdop = CInt(Mid(dateOfProcessing, 6, 2))
        
        If (yrdop >= AssYear) Then
            calcIntrst234BUptoPeriod = mthdop - 4 + (yrdop - AssYear) * 12
        End If
        
        calcIntrst234BUptoPeriod = calcIntrst234BUptoPeriod + 1
        
        carryForwardPrinicipal = shortFall
        carryForwardPrinicipal = WorksheetFunction.RoundDown(carryForwardPrinicipal, -2)
        For i = 1 To calcIntrst234BUptoPeriod
            
            j = i - looper * 12
            looperint = looperint + 1
                        
           If j = 1 Then
        
                SATPaidAtPeriod = 0
                SATPaidAtlooper = 0
                
                balancePrincipal = carryForwardPrinicipal
                'balancePrincipal = WorksheetFunction.RoundDown(balancePrincipal, -2)
                April_1 = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & 1, Range("TaxP.Amt"))
                may_1 = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & 2, Range("TaxP.Amt"))
                June_1 = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & 3, Range("TaxP.Amt"))
                
                
                If June_1 > 0 Then
                CONST_IntrstPay234B_Percentage_1 = 1
                ElseIf may_1 > 0 Then
                CONST_IntrstPay234B_Percentage_1 = 1
                ElseIf April_1 > 0 Then
                CONST_IntrstPay234B_Percentage_1 = 1
                Else
                CONST_IntrstPay234B_Percentage_1 = 1
                End If
                
                calcIntrst234BOnPeriod = Math.Round(CONST_IntrstPay234B_Percentage_1 / 100 * balancePrincipal)
                calcInterestPayable234B = calcInterestPayable234B + calcIntrst234BOnPeriod
                
                If looper = 0 Then
                    balanceInterest = intrstPayUs234A_1 + intrstPayUs234F_1 + calcIntrst234BOnPeriod + intrstPayUs234C_1
                Else
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                End If
                
                SATPaidAtlooper = looperint + looper * 12
                SATPaidAtPeriod = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & SATPaidAtlooper, Range("TaxP.Amt"))
                
                adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
            ElseIf j = 2 Then
                
                SATPaidAtPeriod = 0
                SATPaidAtlooper = 0
                balancePrincipal = carryForwardPrinicipal
                'balancePrincipal = WorksheetFunction.RoundDown(balancePrincipal, -2)
                  April_1 = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & 1, Range("TaxP.Amt"))
                may_1 = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & 2, Range("TaxP.Amt"))
                June_1 = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & 3, Range("TaxP.Amt"))
                
                
                If June_1 > 0 Then
                CONST_IntrstPay234B_Percentage_1 = 1
                ElseIf may_1 > 0 Then
                CONST_IntrstPay234B_Percentage_1 = 1
                ElseIf April_1 > 0 Then
                CONST_IntrstPay234B_Percentage_1 = 1
                Else
                CONST_IntrstPay234B_Percentage_1 = 1
                End If
                calcIntrst234BOnPeriod = Math.Round(CONST_IntrstPay234B_Percentage_1 / 100 * balancePrincipal)
                calcInterestPayable234B = calcInterestPayable234B + calcIntrst234BOnPeriod
                balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                
                SATPaidAtlooper = looperint + looper * 12
                SATPaidAtPeriod = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & SATPaidAtlooper, Range("TaxP.Amt"))
                
                
                adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
                
            ElseIf j = 3 Then
                
                SATPaidAtPeriod = 0
                SATPaidAtlooper = 0
                balancePrincipal = carryForwardPrinicipal
                'balancePrincipal = WorksheetFunction.RoundDown(balancePrincipal, -2)
                April_1 = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & 1, Range("TaxP.Amt"))
                may_1 = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & 2, Range("TaxP.Amt"))
                June_1 = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & 3, Range("TaxP.Amt"))
                
                
                If June_1 > 0 Then
                CONST_IntrstPay234B_Percentage_1 = 1
                ElseIf may_1 > 0 Then
                CONST_IntrstPay234B_Percentage_1 = 1
                ElseIf April_1 > 0 Then
                CONST_IntrstPay234B_Percentage_1 = 1
                Else
                CONST_IntrstPay234B_Percentage_1 = 1
                End If
                
                calcIntrst234BOnPeriod = Math.Round(CONST_IntrstPay234B_Percentage_1 / 100 * balancePrincipal)
                calcInterestPayable234B = calcInterestPayable234B + calcIntrst234BOnPeriod
                balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                
                SATPaidAtlooper = looperint + looper * 12
                SATPaidAtPeriod = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & SATPaidAtlooper, Range("TaxP.Amt"))
                
                adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
                
            ElseIf j = 4 Then
                
                SATPaidAtPeriod = 0
                SATPaidAtlooper = 0
                balancePrincipal = carryForwardPrinicipal
               ' balancePrincipal = WorksheetFunction.RoundDown(balancePrincipal, -2)
                calcIntrst234BOnPeriod = Math.Round(CONST_IntrstPay234B_Percentage / 100 * balancePrincipal)
                calcInterestPayable234B = calcInterestPayable234B + calcIntrst234BOnPeriod
                balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                
                'SATPaidAtPeriod = SATPaidAtPeriod + Sheet4.Range("Sat16").value
                SATPaidAtlooper = looperint + looper * 12
                SATPaidAtPeriod = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & SATPaidAtlooper, Range("TaxP.Amt"))
                
                adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
            ElseIf j = 5 Then
                
                SATPaidAtPeriod = 0
                SATPaidAtlooper = 0
                balancePrincipal = carryForwardPrinicipal
               ' balancePrincipal = WorksheetFunction.RoundDown(balancePrincipal, -2)
                calcIntrst234BOnPeriod = Math.Round(CONST_IntrstPay234B_Percentage / 100 * balancePrincipal)
                calcInterestPayable234B = calcInterestPayable234B + calcIntrst234BOnPeriod
                balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                
                'SATPaidAtPeriod = SATPaidAtPeriod + Sheet4.Range("Sat17").value
                SATPaidAtlooper = looperint + looper * 12
                SATPaidAtPeriod = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & SATPaidAtlooper, Range("TaxP.Amt"))
                
                adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
            ElseIf j = 6 Then
                
                SATPaidAtPeriod = 0
                SATPaidAtlooper = 0
                balancePrincipal = carryForwardPrinicipal
               ' balancePrincipal = WorksheetFunction.RoundDown(balancePrincipal, -2)
                calcIntrst234BOnPeriod = Math.Round(CONST_IntrstPay234B_Percentage / 100 * balancePrincipal)
                calcInterestPayable234B = calcInterestPayable234B + calcIntrst234BOnPeriod
                balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                
                'SATPaidAtPeriod = SATPaidAtPeriod + Sheet4.Range("Sat18").value
                SATPaidAtlooper = looperint + looper * 12
                SATPaidAtPeriod = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & SATPaidAtlooper, Range("TaxP.Amt"))
                
                adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
            ElseIf j = 7 Then
                
                SATPaidAtPeriod = 0
                SATPaidAtlooper = 0
                balancePrincipal = carryForwardPrinicipal
               ' balancePrincipal = WorksheetFunction.RoundDown(balancePrincipal, -2)
                calcIntrst234BOnPeriod = Math.Round(CONST_IntrstPay234B_Percentage / 100 * balancePrincipal)
                calcInterestPayable234B = calcInterestPayable234B + calcIntrst234BOnPeriod
                balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                
                'SATPaidAtPeriod = SATPaidAtPeriod + Sheet4.Range("Sat19").value
                SATPaidAtlooper = looperint + looper * 12
                SATPaidAtPeriod = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & SATPaidAtlooper, Range("TaxP.Amt"))
                
                adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
            ElseIf j = 8 Then
                
                SATPaidAtPeriod = 0
                SATPaidAtlooper = 0
                balancePrincipal = carryForwardPrinicipal
               ' balancePrincipal = WorksheetFunction.RoundDown(balancePrincipal, -2)
                calcIntrst234BOnPeriod = Math.Round(CONST_IntrstPay234B_Percentage / 100 * balancePrincipal)
                calcInterestPayable234B = calcInterestPayable234B + calcIntrst234BOnPeriod
                balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                
                'SATPaidAtPeriod = SATPaidAtPeriod + Sheet4.Range("Sat20").value
                SATPaidAtlooper = looperint + looper * 12
                SATPaidAtPeriod = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & SATPaidAtlooper, Range("TaxP.Amt"))
                
                adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
            ElseIf j = 9 Then
                
                SATPaidAtPeriod = 0
                SATPaidAtlooper = 0
                balancePrincipal = carryForwardPrinicipal
               ' balancePrincipal = WorksheetFunction.RoundDown(balancePrincipal, -2)
                calcIntrst234BOnPeriod = Math.Round(CONST_IntrstPay234B_Percentage / 100 * balancePrincipal)
                calcInterestPayable234B = calcInterestPayable234B + calcIntrst234BOnPeriod
                balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                
                'SATPaidAtPeriod = SATPaidAtPeriod + Sheet4.Range("Sat21").value
                SATPaidAtlooper = looperint + looper * 12
                SATPaidAtPeriod = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & SATPaidAtlooper, Range("TaxP.Amt"))
                
                adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
            ElseIf j = 10 Then
                
                SATPaidAtPeriod = 0
                SATPaidAtlooper = 0
                balancePrincipal = carryForwardPrinicipal
               ' balancePrincipal = WorksheetFunction.RoundDown(balancePrincipal, -2)
                calcIntrst234BOnPeriod = Math.Round(CONST_IntrstPay234B_Percentage / 100 * balancePrincipal)
                calcInterestPayable234B = calcInterestPayable234B + calcIntrst234BOnPeriod
                balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                
                'SATPaidAtPeriod = SATPaidAtPeriod + Sheet4.Range("Sat22").value
                SATPaidAtlooper = looperint + looper * 12
                SATPaidAtPeriod = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & SATPaidAtlooper, Range("TaxP.Amt"))
                
                adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
            ElseIf j = 11 Then
                
                SATPaidAtPeriod = 0
                SATPaidAtlooper = 0
                balancePrincipal = carryForwardPrinicipal
             '   balancePrincipal = WorksheetFunction.RoundDown(balancePrincipal, -2)
                calcIntrst234BOnPeriod = Math.Round(CONST_IntrstPay234B_Percentage / 100 * balancePrincipal)
                calcInterestPayable234B = calcInterestPayable234B + calcIntrst234BOnPeriod
                balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                
                'SATPaidAtPeriod = SATPaidAtPeriod + Sheet4.Range("Sat23").value
                SATPaidAtlooper = looperint + looper * 12
                SATPaidAtPeriod = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & SATPaidAtlooper, Range("TaxP.Amt"))
                
                adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
            ElseIf j = 12 Then
                
                SATPaidAtPeriod = 0
                SATPaidAtlooper = 0
                balancePrincipal = carryForwardPrinicipal
               ' balancePrincipal = WorksheetFunction.RoundDown(balancePrincipal, -2)
                calcIntrst234BOnPeriod = Math.Round(CONST_IntrstPay234B_Percentage / 100 * balancePrincipal)
                calcInterestPayable234B = calcInterestPayable234B + calcIntrst234BOnPeriod
                balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                
                'SATPaidAtPeriod = SATPaidAtPeriod + Sheet4.Range("Sat24").value
                SATPaidAtlooper = looperint + looper * 12
                SATPaidAtPeriod = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & SATPaidAtlooper, Range("TaxP.Amt"))
                
                adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
                
                looper = looper + 1
                looperint = 0
            End If
        Next
        
    End If
        Calculate_InterestPayable234B = WorksheetFunction.Round(calcInterestPayable234B, 0)
End Function

Sub filingdate()
    Dim todaysdate As String
    Dim newfilingdate As String
    Dim originalfilingdate As String
    Dim origrevised As Boolean
    Dim origrevisedstatus As String
    Dim retfilestatus As String
    Dim UpdatedY As Boolean
    UpdatedY = False
    If (((Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))) = "21") And (Mid(Range("U_PreviouslyFiledForThisAY"), 1, 1) = "Y")) Then
    UpdatedY = True
    End If
    
    origrevisedstatus = Mid(Sheet1.Range("sheet1.ReturnType1").value, 1, 1)
    retfilestatus = Mid(Sheet1.Range("sheet1.ReturnFileSec1"), 1, 2)
    If retfilestatus = "18" Or retfilestatus = "17" Or UpdatedY Then
        origrevisedstatus = "R"
    End If
        
    todaysdate = Sheet9.Range("DateOfProcessing").value
    If (origrevisedstatus = "R") Then
        If Not ValidateOrigRetFiledDate_11() Then
            fmsgboxStatus "* Please enter Original Return Date "
        Else
            newfilingdate = OrigRetFiledDate_11
        End If
        newfilingdate = Mid(newfilingdate, 9, 2) + "/" + Mid(newfilingdate, 6, 2) + "/" + Mid(newfilingdate, 1, 4)
        Sheet9.Range("DateOfFiling234A").value = newfilingdate
    End If
    
    'If (origrevisedstatus <> "R") Then
        If Not ValidateDate_9() Then
            fmsgboxStatus "* Please enter Verification Date"
            EndProcessing
        Else
            newfilingdate = Date_9
            If Not ChkMinInclusiveDate(Date_9, todaysdate) Then
                newfilingdate = todaysdate
            End If
            newfilingdate = Mid(newfilingdate, 9, 2) + "/" + Mid(newfilingdate, 6, 2) + "/" + Mid(newfilingdate, 1, 4)
            Sheet9.Range("DateOfFiling").value = newfilingdate
        End If
    'End If
End Sub

Function ValidateDate_9() As Boolean
    ValidateDate_9 = True
    Date_9 = Sheet5.Range("Ver.Date").value
    
    If Date_9 = "" Or IsEmpty(Date_9) Then
        MsgCalcInterest = MsgCalcInterest & "* Verification Date in Sheet PARTB and Verification is Compulsory" & Chr(13)
        ValidateDate_9 = False
        Exit Function
    End If
    
'    If Not CheckDateddmmyyyy(Date_9) Then
'        ValidateDate_9 = False
'        MsgCalcInterest = MsgCalcInterest & "Date in Sheet : PARTB and Verification   must be a valid dd/mm/yyyy format" & Chr(13)
'        Exit Function
'    Else
        Date_9 = Dformat(Sheet5.Range("Ver.Date"), "yyyy-mm-dd")
'    End If
    
    'If Not CheckDateMinDDMMYYYY(Date_9, 1, 4, 2022, "Verification date cannot be less than 01/04/2022") Then dpk1201
    'Ayush_13
'    If Not CheckDateMinDDMMYYYY(Date_9, 1, 4, 2024, "Verification date cannot be less than 01/04/2024") Then
     If Not CheckDateMinDDMMYYYY(Date_9, 1, 4, 2026, "Verification date cannot be less than 01/04/2026") Then
        ValidateDate_9 = False
        'MsgCalcInterest = MsgCalcInterest & "* Date in Verification , Sheet PARTB and Verification must not be less than 01/04/2022" & Chr(13) dpk1201
        'Ayush_13
'        MsgCalcInterest = MsgCalcInterest & "* Date in Verification , Sheet PARTB and Verification must not be less than 01/04/2024" & Chr(13)
        MsgCalcInterest = MsgCalcInterest & "* Date in Verification , Sheet PARTB and Verification must not be less than 01/04/2026" & Chr(13)
        Exit Function
    Else
        Date_9 = Dformat(Sheet5.Range("Ver.Date"), "yyyy-mm-dd")
    End If
End Function
Function Enddateofthemonth(startDate As Variant) As Boolean
Enddateofthemonth = True
 
If Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 1 Then
    Enddateofthemonth = False
ElseIf (Mid(startDate, 1, 2) = 28 Or Mid(startDate, 1, 2) = 29) And Mid(startDate, 4, 2) = 2 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 3 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 30 And Mid(startDate, 4, 2) = 4 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 5 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 30 And Mid(startDate, 4, 2) = 6 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 7 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 8 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 30 And Mid(startDate, 4, 2) = 9 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 10 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 30 And Mid(startDate, 4, 2) = 11 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 12 Then
    Enddateofthemonth = False
Else
    Enddateofthemonth = True
End If
End Function
Function MonthDiff(currentdate As Variant, startDate As Variant) As Long 'By Bindu on 8th Sep 25
On Error Resume Next
 
Dim currentYear As Variant
Dim startyear As Variant
Dim currentMonth As Variant
Dim startmonth As Variant
 
currentYear = val(Mid(currentdate, 7, 4))
startyear = val(Mid(startDate, 7, 4))
'Added by Shrutika 15/07/2025
'Changed by Riyaz on 29/08/2025
'If getDueDate = "15/09/2025" Then
 
If Enddateofthemonth(startDate) Then
    currentMonth = val(Mid(currentdate, 4, 2)) + 1
    'Updated by Ayush on 11/09/2025
Else
    currentMonth = val(Mid(currentdate, 4, 2))
End If

startmonth = val(Mid(startDate, 4, 2))
MonthDiff = 0

If currentYear = startyear And currentMonth = startmonth And _
        val(Mid(currentdate, 1, 2)) = val(Mid(startDate, 1, 2)) Then
        MonthDiff = 0
ElseIf (checkFirstDateBefore(currentdate, startDate)) Then
        MonthDiff = 0
Else
    If currentYear = (startyear + 1) Then
        If currentMonth < startmonth Then
            'MonthDiff = 12 - (startmonth - currentMonth) + 1
            MonthDiff = 12 - (startmonth - currentMonth)
        Else
        'tobetested
            'MonthDiff = 12 + (currentMonth - startmonth) + 1
            MonthDiff = 12 + (currentMonth - startmonth)
        End If
   ElseIf (currentYear = startyear) Then
        'MonthDiff = currentMonth - startmonth + 1
        MonthDiff = currentMonth - startmonth
   If (val(Mid(currentdate, 1, 2)) > val(Mid(startDate, 1, 2))) And (currentMonth = startmonth) Then
        MonthDiff = MonthDiff + 1
   'ElseIf val(Mid(startDate, 1, 2)) = 15 And (currentMonth >= startmonth) Then
   ElseIf val(Mid(startDate, 1, 2)) = 7 And (currentMonth >= startmonth) Then '15/07/2025
        MonthDiff = MonthDiff + 1
     End If
   ElseIf currentMonth < startmonth Then
        'MonthDiff = Round((currentYear - startyear - 1) * 12) + 12 - startmonth + currentMonth + 1
        MonthDiff = Round((currentYear - startyear - 1) * 12) + 12 - startmonth + currentMonth
   ElseIf (currentMonth > startmonth) Then
        'MonthDiff = Round(((currentYear - startyear) * 12)) + currentMonth - startmonth + 1
        MonthDiff = Round(((currentYear - startyear) * 12)) + currentMonth - startmonth
    Else
        'MonthDiff = Round(((currentYear - startyear) * 12)) + 1
        MonthDiff = Round(((currentYear - startyear) * 12))
    End If
End If
'Added by Shrutika 15/07/2025
If currentYear = (startyear + 1) Then
 
If val(Mid(startDate, 1, 2)) = 7 And (currentYear > startyear) Then
        MonthDiff = MonthDiff + 1
End If
End If
 
End Function
Function MonthDiff_Old(startDate As String, endDate As String) As Long
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
    
    MonthDiff = (CInt(endyear) - CInt(startyear)) * 12
    MonthDiff = MonthDiff + (CInt(endmonth) - CInt(startmonth))
    
    If (startDate < endDate) Then
        If (CInt(startday) < Mid(endDate, 9, 2)) Then
            MonthDiff = MonthDiff + 1
            If CInt(startday) = 30 And NoofDaysinmonth = 30 Then
                MonthDiff = MonthDiff - 1
            End If
 
'        ElseIf Mid(endDate, 9, 2) <= (CInt(startday)) And (Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "16") Then
'            MonthDiff = MonthDiff + 1
        ElseIf (CInt(startday)) <= 15 And (CInt(startmonth)) <= (CInt(endmonth)) Then  'Malli 14/07/2025
            MonthDiff = MonthDiff + 1
        ElseIf (CInt(startday)) = 5 And Mid(endDate, 9, 2) <= (CInt(startday)) Then
            MonthDiff = MonthDiff + 1
        End If
    End If
End Function

Function MonthDiffPrev(startDate As String, endDate As String) As Long
    Dim NoofDaysinmonth As Long
    Dim endyear, startyear, startmonth, endmonth, startday As Long
    
    NoofDaysinmonth = 31
    MonthDiffPrev = 0
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
    
    endmonth = Mid(endDate, 6, 2)
    MonthDiffPrev = (CInt(endyear) - CInt(startyear)) * 12
    MonthDiffPrev = MonthDiffPrev + (CInt(endmonth) - CInt(startmonth))
    
    If (startDate < endDate) Then
        If (CInt(startday) < Mid(endDate, 9, 2)) Then
                MonthDiffPrev = MonthDiffPrev + 1
        End If
    End If
End Function

Function CalculateDelayedInMonths(startDate As Date, endDate As Date) As Double
      Dim noOfMonths As Long
      Dim startDateTotDayMonths As Long
      Dim endDateTotDayMonths As Long
      
      noOfMonths = 0
      If (Year(endDate) >= Year(startDate)) Then
          noOfMonths = Math.Abs(month(endDate) - month(startDate)) + (Year(endDate) - Year(startDate)) * 12
          endDateTotDayMonths = month(endDate) + day(endDate)
          startDateTotDayMonths = month(startDate) + day(startDate)

          If ((endDateTotDayMonths > startDateTotDayMonths) And ((endDateTotDayMonths - startDateTotDayMonths) >= 2) And day(endDate) <> day(startDate)) Then
                noOfMonths = noOfMonths + 1
          End If
      End If
   CalculateDelayedInMonths = noOfMonths
End Function


Function ValidateOrigRetFiledDate_1() As Boolean
    ValidateOrigRetFiledDate_1 = True
    OrigRetFiledDate_1 = Sheet1.Range("sheet1.OrigRetFiledDate").value
    
    If Not CheckDateddmmyyyy(OrigRetFiledDate_1) Then
        ValidateOrigRetFiledDate_1 = False
        'Ayush_10/02
'        fmsgboxStatus ("* OrigRetFiledDate in Sheet Income Details  must be a valid dd/mm/yyyy format")
            fmsgboxStatus ("* ""Date of filing of original Return - Please enter valid date in valid format.""")
        Exit Function
    Else
        OrigRetFiledDate_1 = Dformat(Sheet1.Range("sheet1.OrigRetFiledDate"), "yyyy-mm-dd")
    
        If Len(OrigRetFiledDate_1) > 0 Then
            'If Not CheckDateMinDDMMYYYY(OrigRetFiledDate_1, 1, 4, 2022, "Original return filing date cannot be less than 01/04/2022") Then dpk1201
            'If Not CheckDateMinDDMMYYYY(OrigRetFiledDate_1, 1, 4, 2023, "Original return filing date cannot be less than 01/04/2023") Then
            'Ayush_13
'            If Not CheckDateMinDDMMYYYY(OrigRetFiledDate_1, 1, 4, 2024, "Original return filing date cannot be less than 01/04/2024") Then
             If Not CheckDateMinDDMMYYYY(OrigRetFiledDate_1, 1, 4, 2025, "Original return filing date cannot be less than 01/04/2025") Then
                'fmsgboxStatus ("* Original return filing date to be corrected as it should not be less than 01/04/2022") dpk1201
                'fmsgboxStatus ("* Original return filing date to be corrected as it should not be less than 01/04/2023")
                'Ayush_13
'                fmsgboxStatus ("* Original return filing date to be corrected as it should not be less than 01/04/2024")
                fmsgboxStatus ("* Original return filing date to be corrected as it should not be less than 01/04/2025")
                ValidateOrigRetFiledDate_1 = False
                Exit Function
            Else
                OrigRetFiledDate_1 = Dformat(Sheet1.Range("sheet1.OrigRetFiledDate"), "yyyy-mm-dd")
            End If
        End If
    End If
End Function
Function ValidateOrigRetFiledDate_11() As Boolean
Dim Returnfiledstatus As String
Returnfiledstatus = Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))
    ValidateOrigRetFiledDate_11 = True
    Dim UpdatedY As Boolean
    UpdatedY = False
    If (((Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))) = "21") And (Mid(Range("U_PreviouslyFiledForThisAY"), 1, 1) = "Y")) Then
    UpdatedY = True
    End If
    If Returnfiledstatus = "18" Or Returnfiledstatus = "17" Or UpdatedY Then
    OrigRetFiledDate_11 = Sheet1.Range("sheet1.OrigRetFiledDate").value
    End If
    
    If Not CheckDateddmmyyyy(OrigRetFiledDate_11) Then
        ValidateOrigRetFiledDate_11 = False
        'Ayush_10/02
'        fmsgboxStatus ("* OrigRetFiledDate in Sheet Income Details  must be a valid dd/mm/yyyy format")
        fmsgboxStatus ("* ""Date of filing of original Return - Please enter valid date in valid format.""")
        Exit Function
    Else
         If Returnfiledstatus = "18" Or Returnfiledstatus = "17" Or UpdatedY Then
        OrigRetFiledDate_11 = Dformat(Sheet1.Range("sheet1.OrigRetFiledDate"), "yyyy-mm-dd")
        End If
    
        If Len(OrigRetFiledDate_11) > 0 Then
            'If Not CheckDateMinDDMMYYYY(OrigRetFiledDate_11, 1, 4, 2022, "Original return filing date cannot be less than 01/04/2022") Then dpk1201
          'PAG_E30 AY 2023-24 Old
          'If Not CheckDateMinDDMMYYYY(OrigRetFiledDate_11, 1, 4, 2023, "Original return filing date cannot be less than 01/04/2023") Then
               'fmsgboxStatus ("* Original return filing date to be corrected as it should not be less than 01/04/2023")

            'PAG_C30 AY 2024-25 Change
            'Ayush_13
'            If Not CheckDateMinDDMMYYYY(OrigRetFiledDate_11, 1, 4, 2024, "Original return filing date cannot be less than 01/04/2024") Then
'                fmsgboxStatus ("* Original return filing date to be corrected as it should not be less than 01/04/2024")
             If Not CheckDateMinDDMMYYYY(OrigRetFiledDate_11, 1, 4, 2025, "Original return filing date cannot be less than 01/04/2025") Then
                fmsgboxStatus ("* Original return filing date to be corrected as it should not be less than 01/04/2025")
                ValidateOrigRetFiledDate_11 = False
                Exit Function
            Else
        If Returnfiledstatus = "18" Or Returnfiledstatus = "17" Or UpdatedY Then
        OrigRetFiledDate_11 = Dformat(Sheet1.Range("sheet1.OrigRetFiledDate"), "yyyy-mm-dd")
        End If
            End If
        End If
    End If
End Function




Function calculate_SATfor234A()
Dim i As Long
Dim Date_IT, SAT_Amt As Variant
Dim rangecells1, rangecells2 As Range
Dim SATfor234a As Double
SATfor234a = 0
setTblinfo_IT
ReDim Date_IT(end_IT)
ReDim SAT_Amt(end_IT)

Set rangecells1 = Sheet18.Range("TaxP.DateDep").Cells
Set rangecells2 = Sheet18.Range("TaxP.Amt").Cells


For i = 1 To end_IT
 Date_IT(i) = rangecells1.item(i).value
 SAT_Amt(i) = rangecells2.item(i).value


If (Sheet1.Range("IncD.NetTaxLiability").value - Sheet5.Range("IncD.AdvanceTax").value - Sheet5.Range("IncD.TDS").value - Sheet5.Range("IncD.TCS").value) > 100000 Then

'If ((checkFirstDateBefore("01/04/2022", Date_IT(i))) And (checkFirstDateBefore(Date_IT(i), "31/07/2022"))) Then dpk1201

'IT-E2 2023-24 Old
'If ((checkFirstDateBefore("01/04/2023", Date_IT(i))) And (checkFirstDateBefore(Date_IT(i), "31/07/2023"))) Then

'IT-C2 2024-25 AY Change
'Ayush_13
'If ((checkFirstDateBefore("01/04/2024", Date_IT(i))) And (checkFirstDateBefore(Date_IT(i), "31/07/2024"))) Then
'If ((checkFirstDateBefore("01/04/2025", Date_IT(i))) And (checkFirstDateBefore(Date_IT(i), "31/07/2025"))) Then
'If ((checkFirstDateBefore("01/04/2025", Date_IT(i))) And (checkFirstDateBefore(Date_IT(i), "15/09/2025"))) Then 'Ayush Due Date Changes AY2025-26
If ((checkFirstDateBefore("01/04/2025", Date_IT(i))) And (checkFirstDateBefore(Date_IT(i), Sheet9.Range("DueDate_Extend").value))) Then 'Bindu 8th Sep 25
        SATfor234a = SATfor234a + SAT_Amt(i)
End If
Else
'If ((checkFirstDateBefore("01/04/2022", Date_IT(i))) And (checkFirstDateBefore(Date_IT(i), "31/07/2022"))) Then dpk1201

'IT-E2 2023-24 Old
'If ((checkFirstDateBefore("01/04/2023", Date_IT(i))) And (checkFirstDateBefore(Date_IT(i), "31/07/2023"))) Then

'IT-C2 2024-25 AY Change
'Ayush_13
'If ((checkFirstDateBefore("01/04/2024", Date_IT(i))) And (checkFirstDateBefore(Date_IT(i), "31/07/2024"))) Then
'If ((checkFirstDateBefore("01/04/2025", Date_IT(i))) And (checkFirstDateBefore(Date_IT(i), "31/07/2025"))) Then
'If ((checkFirstDateBefore("01/04/2025", Date_IT(i))) And (checkFirstDateBefore(Date_IT(i), "15/09/2025"))) Then  'Ayush Due Date Changes AY2025-26
If ((checkFirstDateBefore("01/04/2025", Date_IT(i))) And (checkFirstDateBefore(Date_IT(i), Sheet9.Range("DueDate_Extend").value))) Then 'Bindu 8th Sep 25
       
        SATfor234a = SATfor234a + SAT_Amt(i)
End If
End If


'If ((checkFirstDateBefore("01/04/2018", Date_IT(i))) And (checkFirstDateBefore(Date_IT(i), "05/08/2018"))) Then
'        SATfor234a = SATfor234a + SAT_Amt(i)
'End If

Next

calculate_SATfor234A = SATfor234a
End Function




Function checkFirstDateBefore(firstDate As Variant, secondDate As Variant) As Boolean
On Error Resume Next

If val(Mid(firstDate, 7, 4)) < val(Mid(secondDate, 7, 4)) Then
        checkFirstDateBefore = True
ElseIf val(Mid(firstDate, 7, 4)) = val(Mid(secondDate, 7, 4)) Then

        If val(Mid(firstDate, 4, 2)) < val(Mid(secondDate, 4, 2)) Then
        
             checkFirstDateBefore = True
             
        ElseIf val(Mid(firstDate, 4, 2)) = val(Mid(secondDate, 4, 2)) Then
        
                If val(Mid(firstDate, 1, 2)) < val(Mid(secondDate, 1, 2)) Then
                
                        checkFirstDateBefore = True
                        
                ElseIf val(Mid(firstDate, 1, 2)) = val(Mid(secondDate, 1, 2)) Then
                
                        checkFirstDateBefore = True
            
                Else
                      checkFirstDateBefore = False
                End If
        Else
            checkFirstDateBefore = False
        End If
Else

    checkFirstDateBefore = False
End If
  
End Function
Sub New234F()


Dim Returnfiledstatus As String
Dim VerificationDate, VerificationDate1 As Variant
Dim dateOfFiling, DateOfFiling1, dueDate As Variant

'Const FilingLastDate As Date = #1/8/2018#
'Const YearEndDate As Date = #12/31/2018#
Dim dat, altduedate As Variant
Dim mont As Variant
Dim yea As Variant
Dim dat1 As Variant
Dim mont1 As Variant
Dim yea1 As Variant
Dim todaysdate1 As Variant
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

todaysdate1 = Sheet9.Range("DateOfProcessing").value

Returnfiledstatus = Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))
VerificationDate = Range("Ver.Date").value
VerificationDate = Dformat(VerificationDate, "yyyy-mm-dd")


If Mid(Sheet1.Range("sheet1.ReturnFileSec1").value, 1, 2) = "17" Or Mid(Sheet1.Range("sheet1.ReturnFileSec1").value, 1, 2) = "18" Or UpdatedY Then
        dateOfFiling = Sheet1.Range("sheet1.OrigRetFiledDate").value
End If

'DateOfFiling = Range("sheet1.OrigRetFiledDate").value

'If Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "14" Or Mid(Sheet1.Range("sheet1.StateCode1").value, 1, 2) = "37" Then
'altduedate = "31/07/2022" dpk1201
'Ayush_13
'altduedate = "31/07/2024"
'altduedate = "31/07/2025"
'altduedate = "15/09/2025" 'Ayush Due Date Changes AY2025-26
altduedate = Sheet9.Range("DueDate_Extend").value   'Bindu 8th Sep 25


'Else
'altduedate = "31/07/2020"
'End If

'altduedate = "31/08/2019"

If Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or UpdatedY Then
VerificationDate = Dformat(dateOfFiling, "yyyy-mm-dd")
Else
VerificationDate = VerificationDate
End If

If Returnfiledstatus <> "17" And Returnfiledstatus <> "18" Or UpdatedY Then
If Not ChkMinInclusiveDate(VerificationDate, todaysdate1) Then
                VerificationDate = todaysdate1
End If
End If
 
dueDate = Dformat(altduedate, "yyyy-mm-dd")
VerificationDate1 = VerificationDate
DateOfFiling1 = Dformat(dateOfFiling, "yyyy-mm-dd")

 
If (Returnfiledstatus = "14" Or Returnfiledstatus = "15" Or Returnfiledstatus = "16" Or Returnfiledstatus = "19" Or Returnfiledstatus = "20") Or (VerificationDate1 <= dueDate) Then
intrstPayUs234F = 0
'Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F

'ElseIf (Returnfiledstatus = "13" Or Returnfiledstatus = "11" Or Returnfiledstatus = "12" Or UpdatedN) And Range("IncD.TotalIncome").value <= 500000 Then ''For 234F calculation need to consider Total in which 112A Included -- updated by Bindu as per BA/DE V3 on 4th Feb 2025
ElseIf (Returnfiledstatus = "13" Or Returnfiledstatus = "11" Or Returnfiledstatus = "12" Or UpdatedN) And Range("IncD.TotalIncome_New").value <= 500000 Then
intrstPayUs234F = 1000
'Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
'ElseIf (Returnfiledstatus = "13" Or Returnfiledstatus = "11" Or Returnfiledstatus = "12" Or UpdatedN) And Range("IncD.TotalIncome").value > 500000 And (VerificationDate1 > DueDate) And (VerificationDate1 <= "2022-12-31") Then dpk1201
'Ayush_13 updated year
'ElseIf (Returnfiledstatus = "13" Or Returnfiledstatus = "11" Or Returnfiledstatus = "12" Or UpdatedN) And Range("IncD.TotalIncome").value > 500000 And (VerificationDate1 > DueDate) And (VerificationDate1 <= "2025-12-31") Then
ElseIf (Returnfiledstatus = "13" Or Returnfiledstatus = "11" Or Returnfiledstatus = "12" Or UpdatedN) And Range("IncD.TotalIncome_New").value > 500000 And (VerificationDate1 > dueDate) And (VerificationDate1 <= "2025-12-31") Then ''For 234F calculation need to consider Total in which 112A Included -- updated by Bindu as per BA/DE V3 on 4th Feb 2025

intrstPayUs234F = 5000
'Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
'Else
'intrstPayUs234F = 10000
'Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
'End If
ElseIf (Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or UpdatedY) And (DateOfFiling1 <= dueDate) Then
intrstPayUs234F = 0
'Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
'ElseIf (Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or UpdatedY) And Range("IncD.TotalIncome").value <= 500000 Then ''For 234F calculation need to consider Total in which 112A Included -- updated by Bindu as per BA/DE V3 on 4th Feb 2025
ElseIf (Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or UpdatedY) And Range("IncD.TotalIncome_New").value <= 500000 Then
intrstPayUs234F = 1000
'Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
'ElseIf (Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or UpdatedY) And Range("IncD.TotalIncome").value > 500000 And (DateOfFiling1 > DueDate) And (DateOfFiling1 <= "2022-12-31") Then dpk1201
'Ayush_13 Year updated
'ElseIf (Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or UpdatedY) And Range("IncD.TotalIncome").value > 500000 And (DateOfFiling1 > DueDate) And (DateOfFiling1 <= "2025-12-31") Then''For 234F calculation need to consider Total in which 112A Included -- updated by Bindu as per BA/DE V3 on 4th Feb 2025
ElseIf (Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or UpdatedY) And Range("IncD.TotalIncome_New").value > 500000 And (DateOfFiling1 > dueDate) And (DateOfFiling1 <= "2025-12-31") Then
intrstPayUs234F = 5000
'Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
'ElseIf (Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or UpdatedY) And Range("IncD.GrossTotIncome").value < 0 And (DateOfFiling1 > DueDate) Then 'For 234F calculation need to consider GTI in which 112A Included -- updated by Bindu as per BA/DE V3 on 4th Feb 2025
ElseIf (Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or UpdatedY) And Range("IncD.GrossTotIncome_New").value < 0 And (DateOfFiling1 > dueDate) Then
intrstPayUs234F = 1000
Else
intrstPayUs234F = 5000
'Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
End If
End Sub
'---------------------------------------------------------------------------------------------------------








