Attribute VB_Name = "SchOS"
Option Explicit


Dim msgavalidateSheet14

Public end_PTIDes, end_PTIIncm, os_PTIAmount, os_PTIDescription As Variant
Public rngname_PTIOS As Variant

Public FamilyPension_15 As Variant
Public Deduction57iia_15 As Variant

Public rngname_os, rngname_os_e As Variant
Public rngname_os111 As Variant
Public rngname_os2 As Variant
Public end_os As Variant
Public end_os111 As Variant
Public end_os2 As Variant
Public end_os1 As Variant
Public end_os12 As Variant
Public rngname_os1 As Variant
Public rngname_os12 As Variant
Public end_os111a As Variant
Public end_os111b As Variant

Public end_osDTAA, end_osDTAA1, end_osDTAA2, end_osDTAA3, end_osDTAA4, end_osDTAA5, end_osDTAA6, end_osDTAA7 As Variant
Public rngname_osDTAA As Variant
Public DividendGross_15 As Variant
Public InterestGross_15 As Variant
Public InterestGross_i As Variant
Public InterestGross_ii As Variant
Public InterestGross_iii As Variant
Public InterestGross_iv As Variant
Public RentFromMachPlantBldgs_15 As Variant
Public OthersGross_15 As Variant
Public TotalOSGross_15 As Variant
Public fi_WinningFrmLotteries_os As Variant
Public fii_SecXIIOth_os As Variant
Public fiii_TotalOSGrossChargblSplRate_os As Variant
Public g_GrossAmtChargblNormalRate_os As Variant

Public os_Country As Variant
Public os_CountryName As Variant
Public os_CountryCode As Variant
Public os_Article_DTAA As Variant
Public os_Rate_DTAA  As Variant
Public os_TRC As Variant
Public os_SectionAct As Variant
Public os_Amount_income As Variant
Public os_DTAAAmount As Variant
Public os_DTAAAmount1 As Variant
Public os_DTAAAmount2a As Variant
'Added by AYush
Public os_DTAAAmount2aii As Variant



Public os_DTAAAmount2d As Variant
Public os_DTAAAmount2e As Variant
Public os_DTAAAmount5A1ai As Variant
Public os_DTAAAmount5ACA1a As Variant
Public os_DTAAAmount5AC As Variant
Public os_DTAAAmount5AD1i As Variant
Public os1ai As Variant
Public os1aie As Variant

Public os_Rate_ITAct, ApplRate As Variant
Public os_Nature_income As Variant



Public Expenses_15 As Variant
Public Depreciation_15 As Variant
Public TotDeductions_15 As Variant
Public BalanceNoRaceHorse_15 As Variant
Public WinLottRacePuzz_15 As Variant
'Newly added
Public WinOnlineGames_15 As Variant

'Newly added by sai on 20/02/2025 for AY 2025-26==========
Public WinOnlineGames_15_2aii_i As Variant
Public WinOnlineGames_15_2aii_ii As Variant
Public BBDA_1aiii_Upto15Of6, BBDA_1aiii_Upto15Of9, BBDA_1aiii_Up16Of9To15Of12, BBDA_1aiii_Up16Of12To15Of3, BBDA_1aiii_Up16Of3To31Of3 As Variant
'*********************************************************

Public os_DTAAAmount5A1aA As Variant
Public CheckDTAA5A1aA As Variant
Public CheckOS5A1aA As Variant
Public CalculatedOS5A1aA As Variant
Public dtaa_5A1aA As Variant
Public Source_5A1aA As Variant
Public PTI_5A1aA_OS As Variant
Public PTI_5A1aA_DTAA As Variant


Public TotOthSrcNoRaceHorse_15 As Variant
Public Receipts_15 As Variant
Public DeductSec57_15 As Variant
Public BalanceOwnRaceHorse_15 As Variant
Public OthersSourceSection_OS As Variant
Public OthersSourceAmount_OS As Variant
Public SourceDescription_OS As Variant
Public SourceAmount_OS As Variant

Public AmtNotDedSec58 As Variant
Public ProfitCharSec59 As Variant
Public AmtNotDed As Variant
Public ProfitSec59 As Variant

Public os_AssYear As Variant
Public os_IncomeBenefit As Variant
Public os_TaxBenefit As Variant
Sub ValidateOSClick()
ValidateSheetOtherSource
fmsgboxStatus "Sheet Others Source (OS) is OK" ', vbOKOnly, "ITR-3"
End Sub
Sub ValidateSheetOtherSource()
    If Not ValidatesheetOS Then
        Sheet14.Activate
        fmsgboxsmall getmsgos ', vbOKOnly, "Error(s)!"
        CloseMsg
    Else
        'MsgBox "Sheet is ok", vbOKOnly, "ITR-3"
    End If
End Sub
Sub AddRowsOS_DTAA()
Dim vRows As Long
Sheets("OS").Activate
EfilingCommon.DefinedgridNameRange = "os.Amount_income||os.NatureOfIncome||os.Country||os.Article_DTAA||os.Rate_DTAA||os.TRC||os.SectionAct||os.Rate_ITAct||os.ApplRate||os.DTAAcheck||OS_DTAASectionCode||os.DTAA_AmtChk"
ActiveCellRange = EfilingCommon.searchLastRow("os.Amount_income")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub AddRows_os()
'Change 03, CDB- , 30.06.2022
    Dim vRows As Long
    Sheets("OS").Activate
    EfilingCommon.DefinedgridNameRange = "os.SourceDescription||os.SourceAmount||os.check||SRCamt||OS_Section.Validate||Os_Value.Validate"
    ActiveCellRange = EfilingCommon.searchLastRow("os.SourceDescription")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula
    
'    Dim numberofrows As Long
'    setTblinfo_OS
'    SelectLastRow ("os.SourceDescription")
'    numberofrows = InsertRowsAndFillFormulas()
'    Call ExendRangeNameToTable(numberofrows, rngname_os)
'---End Change

End Sub

Sub AddRows_os_e()
    Dim numberofrows As Long
    setTblinfo_OS_e
    SelectLastRow ("os.PTIDescription")
    numberofrows = InsertRowsAndFillFormulas()
    Call ExendRangeNameToTable(numberofrows, rngname_os_e)
End Sub

Sub AddRows_os111()
    Dim numberofrows As Long
    setTblinfo_OS111
    SelectLastRow ("os.AssYear")
    numberofrows = InsertRowsAndFillFormulas()
    Call ExendRangeNameToTable(numberofrows, rngname_os111)
End Sub
Sub AddRows_os1()
    Dim numberofrows As Long
    setTblinfo_OS1
    SelectLastRow ("os.othDesc")
    numberofrows = InsertRowsAndFillFormulas()
    Call ExendRangeNameToTable(numberofrows, rngname_os1)
End Sub

Sub settblinfo_OSDTAA()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet14.Range("os.Amount_income").count
    Set rangecells = Sheet14.Range("os.Amount_income").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_osDTAA = ccount
    rngname_osDTAA = "os.Amount_income;os.NatureOfIncome;os.Country;os.Article_DTAA;os.Rate_DTAA;os.TRC;os.SectionAct;os.Rate_ITAct;os.ApplRate;os.DTAAcheck;OS_DTAASectionCode;os.DTAA_AmtChk;"
 End Sub
 
 Sub settblinfo_OSDTAA1()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet14.Range("os.NatureOfIncome").count
    Set rangecells = Sheet14.Range("os.NatureOfIncome").Cells
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    end_osDTAA1 = ccount
    
 End Sub
 
  Sub settblinfo_OSDTAA2()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet14.Range("os.Country").count
    Set rangecells = Sheet14.Range("os.Country").Cells
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    end_osDTAA2 = ccount
    
 End Sub
 
   Sub settblinfo_OSDTAA3()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet14.Range("os.Article_DTAA").count
    Set rangecells = Sheet14.Range("os.Article_DTAA").Cells
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    end_osDTAA3 = ccount
    
 End Sub
 
    Sub settblinfo_OSDTAA4()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet14.Range("os.Rate_DTAA").count
    Set rangecells = Sheet14.Range("os.Rate_DTAA").Cells
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    end_osDTAA4 = ccount
    
 End Sub
 
   Sub settblinfo_OSDTAA5()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet14.Range("os.TRC").count
    Set rangecells = Sheet14.Range("os.TRC").Cells
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    end_osDTAA5 = ccount
    
 End Sub
 
    Sub settblinfo_OSDTAA6()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet14.Range("os.SectionAct").count
    Set rangecells = Sheet14.Range("os.SectionAct").Cells
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    end_osDTAA6 = ccount
    
 End Sub
 
     Sub settblinfo_OSDTAA7()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet14.Range("os.Rate_ITAct").count
    Set rangecells = Sheet14.Range("os.Rate_ITAct").Cells
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    end_osDTAA7 = ccount
    
 End Sub

'dpk1
Sub setTblinfo_OS()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet14.Range("os.SourceDescription").count
    Set rangecells = Sheet14.Range("os.SourceDescription").Cells
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    end_os = ccount
    rngname_os = "os.SourceDescription;os.SourceAmount;os.check;SRCamt;OS_Section.Validate;Os_Value.Validate;"
 End Sub
 Sub setTblinfo_OS_e()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet14.Range("os.PTIDescription").count
    Set rangecells = Sheet14.Range("os.PTIDescription").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_os = ccount
    rngname_os_e = "os.PTIDescription;os.PTIAmount;os_PTI_check;OS_PTI_Amount;OS_PTISectionCode;"
 End Sub
 
Sub setTblinfo_OS111()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet14.Range("os.AssYear").count
    Set rangecells = Sheet14.Range("os.AssYear").Cells
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
            ccount = ccount + 1
        End If
    Next
    end_os111 = ccount
    rngname_os111 = "os.AssYear;os.IncomeBenefit;os.TaxBenefit;"
 End Sub
 
 Sub setTblinfo_os111a()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet14.Range("os.IncomeBenefit").count
    Set rangecells = Sheet14.Range("os.IncomeBenefit").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_os111a = ccount
    
 End Sub
 Sub setTblinfo_os111b()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet14.Range("os.TaxBenefit").count
    Set rangecells = Sheet14.Range("os.TaxBenefit").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_os111b = ccount
    
 End Sub

Sub setTblinfo_os2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet14.Range("os.SourceAmount").count
    Set rangecells = Sheet14.Range("os.SourceAmount").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_os2 = ccount
    rngname_osDTAA = "os.SourceDescription;os.SourceAmount;os.check;SRCamt;OS_Section.Validate;Os_Value.Validate;"
 End Sub


Sub setTblinfo_OS1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet14.Range("os.othDesc").count
    Set rangecells = Sheet14.Range("os.othDesc").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_os1 = ccount
    rngname_os1 = "os.othDesc;os.OtherSections;"
 End Sub


Sub setTblinfo_os12()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet14.Range("os.OtherSections").count
    Set rangecells = Sheet14.Range("os.OtherSections").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_os12 = ccount
    rngname_os12 = "os.othDesc;os.OtherSections;"
 End Sub

Function getmsgos() As String
    getmsgos = msgavalidateSheet14
End Function


Function ValidatesheetOS() As Boolean
    ValidatesheetOS = True

    If Not ValidateDividendGross_15() Then ValidatesheetOS = False
    If Not ValidateInterestGross_15() Then ValidatesheetOS = False
    If Not ValidateRentFromMachPlantBldgs_15() Then ValidatesheetOS = False
    'If Not ValidateOthersGross_15() Then ValidatesheetOS = False
    
    setTblinfo_OS
    setTblinfo_os2
    end_os = WorksheetFunction.Max(0, end_os, end_os2)
    If Not ValidateSourceDescription_OS() Then ValidatesheetOS = False
    If Not ValidateSourceAmount_OS Then ValidatesheetOS = False
    setTblinfo_OS1
    setTblinfo_os12
    
    end_os1 = WorksheetFunction.Max(0, end_os1, end_os12)
    If Not ValidateOthersSectionOS() Then ValidatesheetOS = False
    If Not ValidateOthersSectionAmtOS() Then ValidatesheetOS = False
 
    If Not ValidateExpenses_15() Then ValidatesheetOS = False
    If Not ValidateDepreciation_15() Then ValidatesheetOS = False
    If Not ValidateAmtNotDedSec58() Then ValidatesheetOS = False
    If Not ValidateProfitCharSec59() Then ValidatesheetOS = False
    
    If Not ValidateTotDeductions_15() Then ValidatesheetOS = False
    If Not ValidateBalanceNoRaceHorse_15() Then ValidatesheetOS = False
    If Not ValidateWinLottRacePuzz_15() Then ValidatesheetOS = False
    'Chandru Change as per OS_C7 AY24-25
    If Not ValidateWinOnlineGames_15() Then ValidatesheetOS = False
    
    If Not ValidateTotOthSrcNoRaceHorse_15() Then ValidatesheetOS = False
    If Not ValidateReceipts_15() Then ValidatesheetOS = False
    If Not ValidateDeductSec57_15() Then ValidatesheetOS = False
    If Not ValidateProfitSec59() Then ValidatesheetOS = False
    If Not ValidateNotDeductSec58() Then ValidatesheetOS = False
    
   setTblinfo_PTIDes
    setTblinfo_PTIIncm
    end_PTIDes = WorksheetFunction.Max(0, end_PTIDes, end_PTIIncm)
    
    If Not Validateos_PTIDescription() Then ValidatesheetOS = False
     If Not Validateos_PTIAmount() Then ValidatesheetOS = False
    
    If Not ValidateBalanceOwnRaceHorse_15() Then ValidatesheetOS = False
    
    If Not ValidateOS_Section1() Then ValidatesheetOS = False
    If Not ValidateOS_Amount1() Then ValidatesheetOS = False
    
    '********************************* For AY 2025-26 by Sai
    If Not ValidateBBDA_1aiii() Then ValidatesheetOS = False
    '***************************************************************
    
    If Not validate_OSDTAA() Then ValidatesheetOS = False
    If Not validate_OS111() Then ValidatesheetOS = False
    If Not ValidateSectionAmount() Then ValidatesheetOS = False
    
     If (Sheet14.Range("os.DividendGrossaii").value = "") Then
        msgabox15 (" Sl. No. 1aii should not be blank in schedule OS" & Chr(13))
        ValidatesheetOS = False
        Exit Function
     End If
     
'*******************************************************************************
'commented by sadineni on 24/04/2026 as per DE V0.2
'      If (Sheet14.Range("os.DividendGrossaiii").value <= 0) Then
'        If (Sheet13.Range("LTCG.BII_CapitalLoss").value < 0) Or (Sheet13.Range("STCG.AII_CapitalLoss").value < 0) Then
'        'msgabox15 ("* ""Dividend income u/s 2(22)(f) 1aiii in OS should be filled (1aiii > 0 )""" & Chr(13))
''        msgabox15 ("* ""Please ensure to offer Dividend Income u/s 2(22)(f) against the loss claimed in Schedule CG.""" & Chr(13)) 'Updated as per enhancement by sai
'        msgabox15 ("* ""Please ensure to offer Dividend Income u/s 2(22)(f) against the loss claimed in Schedule CG""" & Chr(13)) 'Updated by Aavula Naresh
'        ValidatesheetOS = False
'        Exit Function
'        End If
'     End If


'*********************************************************************************************************

'by sadineni on 29/04/2026 as Per DE V0.2
Dim sumof1aiii2d2e As Variant
 
sumof1aiii2d2e = Sheet14.Range("os.DividendGrossaiii").value + Sheet14.Range("os.TotalChargeableSpecialRate").value + Sheet14.Range("os.TotalPassThroughIncome").value
 
    If (Sheet13.Range("LTCG.BII_CapitalLoss").value < 0) Or (Sheet13.Range("STCG.AII_CapitalLoss").value < 0) Then
            If sumof1aiii2d2e = "" Or sumof1aiii2d2e = 0 Then
                msgabox15 ("* ""Please ensure to offer Dividend Income u/s 2(22)(f) or at Sl. No. 2d/2e against the loss claimed in Schedule CG""" & Chr(13)) 'Updated by Aavula Naresh
                ValidatesheetOS = False
                Exit Function
            End If
     End If
'********************************************************************************************************
     
     
     'Ayush_13/06
     If Sheet14.Range("os.expenditureClaimed").value > (Sheet14.Range("os.DividendGross").value) * 0.2 Then
        msgabox15 ("* ""Interest expenditure claimed u/s 57(1) cannot be more than 20% of dividend income.""" & Chr(13))
                ValidatesheetOS = False
        Exit Function
    End If
     
     
     
     
    
End Function

Function ValidateOS_DTAA_Section() As Boolean
    ValidateOS_DTAA_Section = True
    Dim rangecells1 As Range
    Dim i As Long
    Dim OS_DTAASection, OS_DTAASectionVal_1 As String
    
    'settblinfo_OSDTAA
    
    Set rangecells1 = Sheet14.Range("os.SectionAct").Cells
    
    For i = 1 To end_osDTAA
        If Not isdropdownblank(os_Country(i)) Then
            OS_DTAASection = Sheet14.Range(Replace(rangecells1.item(i).Address, "I", "P")).value 'rangecells1.item(i).value
            OS_DTAASectionVal_1 = Sheet14.Range(Replace(rangecells1.item(i).Address, "I", "Q")).value 'rangecells1.item(i).value

        End If
    Next
End Function



Function ValidateOS_Section1() As Boolean
    ValidateOS_Section1 = True
    Dim rangecells1, rangecells3, rangecells5 As Range
    Dim i As Long
    Dim OS_Section As String
    Dim OS_income As Variant
    Dim section As Variant
    
    setTblinfo_OS
    
    Set rangecells1 = Sheet14.Range("OS_Section.Validate").Cells
    Set rangecells3 = Sheet14.Range("Os_Value.Validate").Cells
    Set rangecells5 = Sheet14.Range("os.check").Cells
    
    For i = 1 To end_os
        OS_Section = rangecells1.item(i).value
        OS_income = rangecells3.item(i).value
        section = rangecells5.item(i).value
        
        If OS_Section = 1 And OS_income = 0 Then
            msgabox15 ("* Income against sections:  " & section & " added is Mandatory in Sheet OS ")
            ValidateOS_Section1 = False
            Exit Function
        End If
    Next
End Function

Function ValidateOS_Amount1() As Boolean
    ValidateOS_Amount1 = True
    Dim rangecells1, rangecells3, rangecells5 As Range
    Dim i As Long
    Dim OS_Section As String
    Dim OS_income As Variant
    Dim section As Variant
    
    setTblinfo_os2
    
    Set rangecells1 = Sheet14.Range("OS_Section.Validate").Cells
    Set rangecells3 = Sheet14.Range("Os_Value.Validate").Cells
    Set rangecells5 = Sheet14.Range("os.check").Cells
    
    For i = 1 To end_os2
        OS_Section = rangecells1.item(i).value
        OS_income = rangecells3.item(i).value
        section = rangecells5.item(i).value
        
        If OS_income = 1 And OS_Section = 0 Then
            msgabox15 ("* Section against income added in Sheet OS  is Mandatory")
            ValidateOS_Amount1 = False
            Exit Function
        End If
    Next
End Function
Function validate_OS111() As Boolean
validate_OS111 = True
Dim rangecells, rangecells1, rangecells2, rangecells3 As Range
Dim i As Variant

Set rangecells = Sheet14.Range("os.AssYear")
Set rangecells1 = Sheet14.Range("os.IncomeBenefit")
Set rangecells2 = Sheet14.Range("os.TaxBenefit")


setTblinfo_OS111
setTblinfo_os111a
setTblinfo_os111b
end_os111 = WorksheetFunction.Max(0, end_os111, end_os111a, end_os111b)
ReDim os_AssYear(end_os111)
ReDim os_IncomeBenefit(end_os111)
ReDim os_TaxBenefit(end_os111)


For i = 1 To end_os111

os_AssYear(i) = rangecells.item(i).value
os_IncomeBenefit(i) = rangecells1.item(i).value
os_TaxBenefit(i) = rangecells2.item(i).value

Next

For i = 1 To end_os111

        If os_IncomeBenefit(i) = "" Then
            msgabox15 ("* Please enter the amount in Income Benefit at Sr no. " & i & " in Schedule OS.")
            validate_OS111 = False
        End If
        If Len(os_IncomeBenefit(i)) > 14 Then
            msgabox15 ("* Amount in Income Benefit at Sr no. " & i & " cannot be greater than 14 digits in Schedule OS.")
            validate_OS111 = False
        End If
        If os_TaxBenefit(i) = "" Then
            msgabox15 ("* Please enter the amount in Tax Benefit at Sr no. " & i & " in Schedule OS.")
            validate_OS111 = False
        End If
        If Len(os_TaxBenefit(i)) > 14 Then
            msgabox15 ("* Amount in Tax Benefit at Sr no. " & i & " cannot be greater than 14 digits in Schedule OS.")
            validate_OS111 = False
        End If
        
'Added by sudha as per SIT-95519

        If isdropdownblank(os_AssYear(i)) Then
            'msgabox15 ("* Please select the Assessment Year at Sr no. " & i & " in Schedule OS.")
            
             msgabox15 ("*Please select the Assessment Year in 2c of OS schedule at Sr no. " & i)
            validate_OS111 = False
     End If

Next

End Function
Function validate_OSDTAA() As Boolean
validate_OSDTAA = True
Dim rangecells, rangecells1, rangecells2, rangecells3, rangecells4, rangecells5, rangecells6, rangecells7, rangecells8 As Range
Dim i As Variant
Dim CheckDivDTAA As Variant

'For AY 2025-26 added by sai on 20/02/2025=====================
 Dim CheckDivDTAA_1aiii As Variant
'==============================================================
Dim CheckDTAA5ACA1a As Variant
Dim CheckDTAA5AC As Variant
Dim CheckDTAA5A1ai As Variant
Dim CheckDTAA5AD1i As Variant

Dim CheckTotalDivDTAA As Variant
Dim CheckLottery As Variant
'Added by Ayush
Dim CheckOnlineGame As Variant

Dim CheckOSInc89A As Variant
Dim CheckOS5A1ai As Variant
Dim CheckOS5ACA1a As Variant
Dim CheckOS5AC As Variant
Dim CheckOS5AD1i As Variant
Dim CalculatedLottery As Variant
'Added by Ayush
Dim CalculatedOnlineGame As Variant

Dim CalculatedOSInc89A As Variant
Dim CalculatedOS5A1ai As Variant
Dim CalculatedOS5ACA1a As Variant
Dim CalculatedOS5AC As Variant
Dim CalculatedOS5AD1i As Variant
Dim DTAA_5BB_OS As Variant

Set rangecells = Sheet14.Range("os.Country")
Set rangecells1 = Sheet14.Range("os.Article_DTAA")
Set rangecells2 = Sheet14.Range("os.Rate_DTAA")
Set rangecells3 = Sheet14.Range("os.TRC")
Set rangecells4 = Sheet14.Range("os.SectionAct")
Set rangecells5 = Sheet14.Range("os.Amount_income")
Set rangecells6 = Sheet14.Range("os.NatureOfIncome")
Set rangecells7 = Sheet14.Range("os.Rate_ITAct")
Set rangecells8 = Sheet14.Range("os.ApplRate")


settblinfo_OSDTAA
settblinfo_OSDTAA1
settblinfo_OSDTAA2
settblinfo_OSDTAA3
settblinfo_OSDTAA4
settblinfo_OSDTAA5
settblinfo_OSDTAA6
settblinfo_OSDTAA7

end_osDTAA = WorksheetFunction.Max(0, end_osDTAA, end_osDTAA1, end_osDTAA2, end_osDTAA3, end_osDTAA4, end_osDTAA5, end_osDTAA6, end_osDTAA7)

ReDim os_Country(end_osDTAA)
ReDim os_CountryName(end_osDTAA)
ReDim os_CountryCode(end_osDTAA)
ReDim os_Article_DTAA(end_osDTAA)
ReDim os_Rate_DTAA(end_osDTAA)
ReDim os_Rate_ITAct(end_osDTAA)
ReDim os_TRC(end_osDTAA)
ReDim os_SectionAct(end_osDTAA)
ReDim os_Amount_income(end_osDTAA)
ReDim os_Nature_income(end_osDTAA)
ReDim ApplRate(end_osDTAA)

For i = 1 To end_osDTAA
os_Country(i) = rangecells.item(i).value
If Not isdropdownblank(os_Country(i)) Then
os_CountryCode(i) = Mid(os_Country(i), WorksheetFunction.Search(":", os_Country(i)) + 1, Len(os_Country(i)))
os_CountryName(i) = Mid(os_Country(i), 1, WorksheetFunction.Search(":", os_Country(i)) - 1)
End If

os_Article_DTAA(i) = rangecells1.item(i).value
os_Rate_DTAA(i) = rangecells2.item(i).text
os_TRC(i) = Mid(rangecells3.item(i).value, 1, 1)
os_SectionAct(i) = Trim(rangecells4.item(i).value)
os_Amount_income(i) = rangecells5.item(i).value
os_Nature_income(i) = rangecells6.item(i).value
os_Rate_ITAct(i) = rangecells7.item(i).value
ApplRate(i) = rangecells8.item(i).value

Next

'If Not isdropdownblank(rangecells.item(1).value) Then
        If Not validateArticleDTAA() Then validate_OSDTAA = False
        If Not validateRateTAXDTAA() Then validate_OSDTAA = False
        If Not validateCorrespondingSection() Then validate_OSDTAA = False
        If Not validateTRC() Then validate_OSDTAA = False
        If Not validateIncomeAmount() Then validate_OSDTAA = False
        If Not ValidateOS_DTAA_Section() Then validate_OSDTAA = False
        If Not ValidateOS_Nature_Income() Then validate_OSDTAA = False
        If Not validateRateITActDTAA() Then validate_OSDTAA = False
        
'End If
os_DTAAAmount = Sheet14.Range("os.DTAA_Amt").value

If validate_OSDTAA = True Then
For i = 1 To end_osDTAA
   
          If isdropdownblank(os_Country(i)) Then
            msgabox15 ("* Country name and Code at Sr no. " & i & " is mandatory in Schedule OS.")
            validate_OSDTAA = False
          End If
          If os_Article_DTAA(i) = "" Then
            msgabox15 ("* Article of DTAA at Sr no. " & i & " is mandatory in Schedule OS.")
            validate_OSDTAA = False
          End If
          
          If os_Rate_DTAA(i) = "" Then
            msgabox15 ("* Rate of Tax under DTAA at Sr no. " & i & " is mandatory in Schedule OS.")
            validate_OSDTAA = False
          End If
          
          If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
          If isdropdownblank(os_TRC(i)) Then
            msgabox15 ("* whether TRC obtained at Sr no. " & i & " is mandatory in Schedule OS.")
            validate_OSDTAA = False
          End If
          End If
          
          If isdropdownblank(os_SectionAct(i)) Then
            msgabox15 ("* Section Code at Sr no. " & i & " is mandatory in Schedule OS.")
            validate_OSDTAA = False
          End If
          If os_Amount_income(i) = "" Then
            msgabox15 ("* Amount of income at Sr no. " & i & " is mandatory in Schedule OS.")
            validate_OSDTAA = False
          End If
          If os_Nature_income(i) = "" Or os_Nature_income(i) = "(Select)" Then
            msgabox15 ("* Please select an option from dtaa table column 2 at  Sr no. " & i & " is mandatory in Schedule OS.")
            validate_OSDTAA = False
          End If
          
          If os_Rate_ITAct(i) = "" Then
            msgabox15 ("* Rate as per I.T. Act at Sr no. " & i & " is mandatory in Schedule OS.")
            validate_OSDTAA = False
          End If
          
    
Next
End If

'dpk
os_DTAAAmount = Sheet14.Range("os.DTAA_Amt").value
os_DTAAAmount1 = Sheet14.Range("DivDTAA").value + Sheet14.Range("InterestDTAA").value + Sheet14.Range("RentDTAA").value + Sheet14.Range("Sec56DTAA").value
os_DTAAAmount2a = Sheet14.Range("WinLottRacePuzzDTAA").value
'Added by Ayush
os_DTAAAmount2aii = Sheet14.Range("WinOnlineGameDTAA").value

os_DTAAAmount2d = Sheet14.Range("Total562TaxDTAA").value
os_DTAAAmount2e = Sheet14.Range("TotalPassThroughIncomeDTAA").value
os_DTAAAmount5A1ai = Sheet14.Range("DTAA_5A1ai_OS").value + Sheet14.Range("Os_PTI_5A1ai_DTAA").value

'Chandru
os_DTAAAmount5A1aA = Sheet14.Range("DTAA_5A1aA_OS").value + Sheet14.Range("Os_PTI_5A1aA_DTAA").value

os_DTAAAmount5ACA1a = Sheet14.Range("DTAA_5ACA1a_OS").value + Sheet14.Range("Os_PTI_5ACA1a_DTAA").value
os_DTAAAmount5AC = Sheet14.Range("DTAA_5AC1b_OS").value + Sheet14.Range("Os_PTI_5AC1b_DTAA").value
os_DTAAAmount5AD1i = Sheet14.Range("DTAA_5A1D1i_OS").value + Sheet14.Range("Os_PTI_5AD1i_DTAA").value

'DTAA_5BB_OS = Sheet14.Range("DTAA_5BB_OS").value 'sachin
CheckDivDTAA = Sheet14.Range("DivDTAA").value

'For AY 2025-26 added by sai on 20/02/2025====================
CheckDivDTAA_1aiii = Sheet14.Range("DivDTAA_1aiii").value
'==============================================================


CheckDTAA5ACA1a = Sheet14.Range("DTAA5ACA1a").value
CheckDTAA5AC = Sheet14.Range("DTAA5AC").value
CheckDTAA5A1ai = Sheet14.Range("DTAA5A1ai").value

'Chandru
CheckDTAA5A1aA = Sheet14.Range("DTAA5A1aA").value


CheckDTAA5AD1i = Sheet14.Range("DTAA5AD1i").value
CheckTotalDivDTAA = Sheet14.Range("DivDTAATotal").value
CheckLottery = Sheet14.Range("Temp_IncomeFromLottery").value
'Chandru Change as per OS_C7 AY24-25
CheckOnlineGame = Sheet14.Range("Temp_IncomeFromOnlineGame").value

CheckOSInc89A = Sheet14.Range("IncD_qTotOS1").value
CheckOS5A1ai = Sheet14.Range("IncD_115A1aiQTot").value

'Chandru
CheckOS5A1aA = Sheet14.Range("IncD_115A1aAQTot").value

CheckOS5ACA1a = Sheet14.Range("IncD_115ACA1aQTot").value
CheckOS5AC = Sheet14.Range("IncD_115ACQTot").value
CheckOS5AD1i = Sheet14.Range("IncD_115AD1iQTot").value
'CalculatedLottery = Sheet14.Range("os.WinLottRacePuzz").value - DTAA_5BB_OS
CalculatedLottery = Sheet14.Range("os.WinLottRacePuzz").value - os_DTAAAmount2a
'Added by Ayush
CalculatedOnlineGame = Sheet14.Range("os.WinOnlineGame").value - os_DTAAAmount2aii

CalculatedOSInc89A = Sheet14.Range("OSIncomeNotified89A").value - Sheet14.Range("OSIncreliefus89A").value
CalculatedOS5A1ai = Sheet14.Range("Source_5A1ai").value + Sheet14.Range("Os_PTI_5A1ai").value - os_DTAAAmount5A1ai
'Chandru
CalculatedOS5A1aA = Sheet14.Range("Source_5A1aA").value + Sheet14.Range("Os_PTI_5A1aA").value - os_DTAAAmount5A1aA

CalculatedOS5ACA1a = Sheet14.Range("Source_5ACA1a_OS").value + Sheet14.Range("OS_PTI_5ACA1a").value - Sheet14.Range("DTAA5ACA1a").value
'CalculatedOS5AC = Sheet14.Range("Source_5A1C1ab_OS").value + Sheet14.Range("OS_PTI_5AC1ab").value + Sheet14.Range("Source_5A1C1b_OS").value + Sheet14.Range("OS_PTI_5AC1b").value - os_DTAAAmount5AC
CalculatedOS5AC = Sheet14.Range("Source_5A1C1b_OS").value + Sheet14.Range("OS_PTI_5AC1b").value - os_DTAAAmount5AC
'CalculatedOS5AD1i = Sheet14.Range("Source_5AD1i_OS").value + Sheet14.Range("OS_PTI_5AD1i").value - os_DTAAAmount5AD1i
CalculatedOS5AD1i = Sheet14.Range("Source_5AD1iDiv_OS").value + Sheet14.Range("OS_PTI_5AD1iDiv").value - Sheet14.Range("DTAA5AD1i").value
os1aie = Sheet14.Range("os.DividendGrossai").value - Sheet14.Range("os.amountinterestexpenditure1").value
os1ai = Sheet14.Range("os.DividendGrossai").value - Sheet14.Range("os.expenditureClaimed1").value
    If (Sheet14.Range("os.amountinterestexpenditure") = "") Then
        'If os1ai <> Sheet14.Range("tempBBDA").value + Sheet14.Range("DivDTAATotal").value Then
        If os1ai <> Sheet14.Range("tempBBDA").value + Sheet14.Range("DivDTAA").value Then
            validate_OSDTAA = False
            msgabox15 ("Please correct your Quarterly Breakup of Dividend/DTAA in SrNo.10. 1ai should match with SrNo.10+3c (Claimed Interest Expenditure).")
            Exit Function
            End If
        Else
        'If os1aie <> Sheet14.Range("tempBBDA").value + Sheet14.Range("DivDTAATotal").value Then
        If os1aie <> Sheet14.Range("tempBBDA").value + Sheet14.Range("DivDTAA").value Then
            validate_OSDTAA = False
            msgabox15 ("Please correct your Quarterly Breakup of Dividend/DTAA in SrNo.10. 1ai should match with SrNo.10+3c (Eligible Interest Expenditure).")
            Exit Function
            End If
        End If

    If (Sheet14.Range("os.InterestGross").value < 0) Then
     If os_DTAAAmount1 > (Sheet14.Range("os.DividendGross").value + _
                        Sheet14.Range("os.RentFromMachPlantBldgs").value) Then
            validate_OSDTAA = False
           msgabox15 ("Total amount of income chargeable to tax under DTAA cannot be greater than value of 1a+1b+1c+1d")
        
        If os_DTAAAmount2a > (Sheet14.Range("os.WinLottRacePuzz").value) Then
             validate_OSDTAA = False
            msgabox15 ("Total amount of income chargeable to tax under DTAA (Winning from Lottry) cannot be greater than value of 2a")
            Exit Function
            End If
        
         'Added by Ayush
         If os_DTAAAmount2aii > (Sheet14.Range("os.WinOnlineGame").value) Then
             validate_OSDTAA = False
            msgabox15 ("Total amount of income chargeable to tax under DTAA (Winning from Online Games) cannot be greater than value of 2aii")
            Exit Function
            End If
            
            
        If os_DTAAAmount2d > (Sheet14.Range("os.TotalChargeableSpecialRate").value) Then
             validate_OSDTAA = False
            msgabox15 ("Total amount of income chargeable to tax under DTAA (Any Other Income Chargeable At Special Rate) cannot be greater than value of 2d")
            Exit Function
            End If
        If os_DTAAAmount2e > (Sheet16.Range("os.TotalPassThroughIncome").value) Then
             validate_OSDTAA = False
            msgabox15 ("Total amount of income chargeable to tax under DTAA (Pass through income in the nature of income from other sources chargeable at special rates) cannot be greater than value of 2e")
            Exit Function
            End If
            
Exit Function
    End If
Else
If os_DTAAAmount1 > (Sheet14.Range("os.DividendGross").value + _
                        Sheet14.Range("os.InterestGross").value + _
                        Sheet14.Range("os.RentFromMachPlantBldgs").value) Then
            validate_OSDTAA = False
           msgabox15 ("Total amount of income chargeable to tax under DTAA cannot be greater than value of 1a+1b+1c+1d")
        
        If os_DTAAAmount2a > (Sheet14.Range("os.WinLottRacePuzz").value) Then
             validate_OSDTAA = False
            msgabox15 ("Total amount of income chargeable to tax under DTAA (Winning from Lottry) cannot be greater than value of 2a")
            Exit Function
            End If
            
        'Chandru Change as per OS_C7 AY24-25
        If os_DTAAAmount2aii > (Sheet14.Range("os.WinOnlineGame").value) Then
             validate_OSDTAA = False
            msgabox15 ("Total amount of income chargeable to tax under DTAA (Winning from Online Games) cannot be greater than value of 2aii")
            Exit Function
            End If
            
            
        If os_DTAAAmount2d > (Sheet14.Range("os.TotalChargeableSpecialRate").value) Then
             validate_OSDTAA = False
            msgabox15 ("Total amount of income chargeable to tax under DTAA (Any Other Income Chargeable At Special Rate) cannot be greater than value of 2d")
            Exit Function
            End If
        If os_DTAAAmount2e > (Sheet14.Range("os.TotalPassThroughIncome").value) Then
             validate_OSDTAA = False
            msgabox15 ("Total amount of income chargeable to tax under DTAA (Pass through income in the nature of income from other sources chargeable at special rates) cannot be greater than value of 2e")
            Exit Function
            End If
            
Exit Function
    End If
End If



If (CheckLottery <> CalculatedLottery) Then
msgabox15 ("The total of all the quarter Income by way of winnings from Lottery etc. must be equal to 115BB income in Schedule SI")
validate_OSDTAA = False
Exit Function
End If

'Added by Ayush
If (CheckOnlineGame <> CalculatedOnlineGame) Then
msgabox15 ("The total of all the quarter Income by way of winnings from Online Games etc. must be equal to 115BBJ income in Schedule SI")
validate_OSDTAA = False
Exit Function
End If

If (CheckOSInc89A <> CalculatedOSInc89A) Then
msgabox15 ("The quarterly break up of Income from retirement benefit account maintained in a notified country u/s 89A should be equal to [Income from retirement benefit account maintained in a notified country u/s 89A-5a. Income claimed for relief from taxation u/s 89A] ")
validate_OSDTAA = False
Exit Function
End If

If (CheckOS5A1ai <> CalculatedOS5A1ai) Then
msgabox15 ("The quarterly break up of Dividend Income u/s 115A(1)(a)(i)  @ 20% ( Including PTI Income) should be equal to Dividend Income u/s 115A(1)(a)(i)( Including PTI Income) - DTAA ")
validate_OSDTAA = False
Exit Function
End If

'Chandru
If (CheckOS5A1aA <> CalculatedOS5A1aA) Then
msgabox15 ("The quarterly break up of Dividend Income u/s 115A(1)(a)(A)  @ 10% ( Including PTI Income) should be equal to Dividend Income u/s 115A(1)(a)(A)( Including PTI Income) - DTAA ")
validate_OSDTAA = False
Exit Function
End If



If (CheckOS5ACA1a <> CalculatedOS5ACA1a) Then
msgabox15 ("The quarterly break up of Dividend Income u/s 115ACA (1)(a) @ 10% (Including PTI Income) should be equal to Dividend Income u/s 115ACA (1)(a) (Including PTI Income) - DTAA ")
validate_OSDTAA = False
Exit Function
End If

If (CheckOS5AD1i <> CalculatedOS5AD1i) Then
msgabox15 ("The quarterly break up of Dividend Income (other than units referred to in section 115AB) u/s 115AD(1)(i) @ 20% (Including PTI Income) should be equal to Dividend Income (other than units referred to in section 115AB) u/s 115AD(1)(i)(Including PTI Income) - DTAA ")
validate_OSDTAA = False
Exit Function
End If

If (CheckOS5AC <> CalculatedOS5AC) Then
msgabox15 ("The quarterly break up of Dividend Income u/s 115AC (1)(b) @ 10% (Including PTI Income) should be equal to Dividend Income u/s 115AC (1)(b)(Including PTI Income) - DTAA ")
validate_OSDTAA = False
Exit Function
End If

Dim OSTotDTAA As Variant

'OSTotDTAA = CheckDivDTAA + CheckDTAA5ACA1a + CheckDTAA5AC + CheckDTAA5A1ai + CheckDTAA5AD1i
'Chandru

'OSTotDTAA = CheckDivDTAA + CheckDTAA5ACA1a + CheckDTAA5AC + CheckDTAA5A1ai + CheckDTAA5A1aA + CheckDTAA5AD1i

'Newly modified by sai on 20/02/2025 for AY 2025-26
OSTotDTAA = CheckDivDTAA + CheckDivDTAA_1aiii + CheckDTAA5ACA1a + CheckDTAA5AC + CheckDTAA5A1ai + CheckDTAA5A1aA + CheckDTAA5AD1i
'=====================================================================================================================

If (OSTotDTAA <> CheckTotalDivDTAA) Then
msgabox15 ("Please check Quarterly Breackup of DTAA")
validate_OSDTAA = False
Exit Function
End If

'If os_DTAAAmount <> 0 Then
'If os_DTAAAmount > (Sheet14.Range("os.DividendGross").value + Sheet14.Range("os.InterestGross").value + Sheet14.Range("os.RentFromMachPlantBldgs").value + Sheet14.Range("os.Totalsec56").value + Sheet14.Range("os.WinLottRacePuzz").value + Sheet14.Range("os.Total115BE").value + Sheet14.Range("os.IncomeBenefitTotal").value + Sheet14.Range("os.TotalChargeableSpecialRate").value + Sheet14.Range("os.TotalPassThroughIncome").value) Then
'validate_OSDTAA = False
'msgabox15 ("* Total amount of income chargeable to tax under DTAA cannot be greater than value of 1a+1b+1c+1d+2a+2b+2c+2d+2e")
'End If
'End If


End Function

Function validateArticleDTAA() As Boolean
Dim i As Variant
validateArticleDTAA = True
For i = 1 To end_osDTAA
'If Not isdropdownblank(os_Country(i)) Then
    If os_Article_DTAA(i) = "" Then
    msgabox15 ("* Article Of DTAA is mandatory at SrNo." & i & " in Schedule OS.")
    validateArticleDTAA = False
'End If
End If
Next

End Function


Function ValidateOS_Nature_Income() As Boolean
Dim i As Variant
ValidateOS_Nature_Income = True
For i = 1 To end_osDTAA
'If Not isdropdownblank(os_Country(i)) Then
    If os_Nature_income(i) = "" Or Mid(os_Nature_income(i), 1, 1) = "(" Then
    msgabox15 ("* Please select an option from dtaa table column no.2 Sr No." & i & " in Schedule OS.")
    ValidateOS_Nature_Income = False
'End If
End If

If Sheet14.Range("os.InterestGross").value < 0 Then
  If os_Nature_income(i) = "1b" Then
    msgabox15 ("* Item No 1b cannot be selected in DTAA Table if amount of interest Gross(1b) is negative  Sr No." & i & " in Schedule OS.")
    ValidateOS_Nature_Income = False
'End If
End If
End If
Next

End Function
Function validateRateTAXDTAA() As Boolean
validateRateTAXDTAA = True
Dim i As Variant
For i = 1 To end_osDTAA
'If Not isdropdownblank(os_Country(i)) Then
    If os_Rate_DTAA(i) = "" Then
    msgabox15 ("* Please select an option from dtaa table at column no." & i & " in Schedule OS.")
    validateRateTAXDTAA = False
    End If
'End If
Next
End Function

Function validateRateITActDTAA() As Boolean
validateRateITActDTAA = True
Dim i As Variant
For i = 1 To end_osDTAA
'If Not isdropdownblank(os_Country(i)) Then
    If os_Rate_ITAct(i) = "" Then
    msgabox15 ("* Rate as per IT Act is Mandatory at SrNo." & i & " in Schedule OS.")
    validateRateITActDTAA = False
    End If
'End If
Next
End Function

Function validateCorrespondingSection() As Boolean
    validateCorrespondingSection = True
    
    Dim i As Variant
    
    For i = 1 To end_osDTAA
    
        
        'If Not isdropdownblank(os_Country(i)) Then

          
            If isdropdownblank(os_SectionAct(i)) Then
          
                msgabox15 ("* Section Act is Mandatory at SrNo." & i & " in Schedule OS.")
                validateCorrespondingSection = False
                Exit Function
            End If
        

    Next
End Function

Function validateIncomeAmount() As Boolean
validateIncomeAmount = True
Dim i As Variant
For i = 1 To end_osDTAA
If Not isdropdownblank(os_Country(i)) Then
    If os_Amount_income(i) = "" Then
    msgabox15 ("* Amount of Income is Mandatory at SrNo." & i & " in Schedule OS.")
    validateIncomeAmount = False
    End If
End If
Next

End Function

Function validateTRC() As Boolean
validateTRC = True
Dim i As Variant
For i = 1 To end_osDTAA
'If Not isdropdownblank(os_Country(i)) Then
  If Mid(Sheet1.Range("sheet1.ResidentialStatus1").value, 1, 3) = "NRI" Then
    If isdropdownblank(os_TRC(i)) Then
    msgabox15 ("* Whether TRC obtained is Mandatory at SrNo." & i & " in Schedule OS.")
    validateTRC = False
    End If
  End If
'End If
Next

End Function

Function ValidateDividendGross_15() As Boolean
    Dim i As Long
    
    ValidateDividendGross_15 = True
    DividendGross_15 = Sheet14.Range("os.DividendGross").value
    
    If Len(DividendGross_15) > 14 Then
        msgabox15 ("* DividendGross in Sheet OS cannot be greater than 14 digits ")
        ValidateDividendGross_15 = False
        Exit Function
    End If
    
    For i = 1 To Len(DividendGross_15)
        If Not IsNumeric(DividendGross_15) Then
            msgabox15 ("* DividendGross in Sheet OS must contain only digits from 0 to 9")
            ValidateDividendGross_15 = False
            Exit Function
        End If
    Next
End Function

Function ValidateInterestGross_15() As Boolean
    Dim i As Long
    
    ValidateInterestGross_15 = True
    InterestGross_15 = Sheet14.Range("os.InterestGross").value
    InterestGross_i = Sheet14.Range("os.SavingBank").value
    InterestGross_ii = Sheet14.Range("os.TermDeposit").value
    InterestGross_iii = Sheet14.Range("os.IncomeTaxRefund").value
    InterestGross_iv = Sheet14.Range("os.Others").value
    
    
    If Len(InterestGross_15) > 14 Then
        msgabox15 ("* InterestGross in Sheet OS cannot be greater than 14 digits ")
        ValidateInterestGross_15 = False
        Exit Function
    End If
    
'    For i = 1 To Len(InterestGross_15)
'        If Not IsNumeric(InterestGross_15) Then
'            msgabox15 ("* InterestGross in Sheet OS must contain only digits from 0 to 9")
'            ValidateInterestGross_15 = False
'            Exit Function
'        End If
'    Next
    
    If Len(InterestGross_i) > 14 Then
        msgabox15 ("* From SavingBank in Sheet OS cannot be greater than 14 digits ")
        ValidateInterestGross_15 = False
        Exit Function
    End If
    
    For i = 1 To Len(InterestGross_i)
        If Not IsNumeric(InterestGross_i) Then
            msgabox15 ("* From SavingBank in Sheet OS must contain only digits from 0 to 9")
            ValidateInterestGross_15 = False
            Exit Function
        End If
    Next
    
    If Len(InterestGross_ii) > 14 Then
        msgabox15 ("* From Term Deposit in Sheet OS cannot be greater than 14 digits ")
        ValidateInterestGross_15 = False
        Exit Function
    End If
    
    For i = 1 To Len(InterestGross_ii)
        If Not IsNumeric(InterestGross_ii) Then
            msgabox15 ("* From Term Deposit in Sheet OS must contain only digits from 0 to 9")
            ValidateInterestGross_15 = False
            Exit Function
        End If
    Next
    
    If Len(InterestGross_iii) > 14 Then
        msgabox15 ("* From Income Tax Refund in Sheet OS cannot be greater than 14 digits ")
        ValidateInterestGross_15 = False
        Exit Function
    End If
    
    For i = 1 To Len(InterestGross_iii)
        If Not IsNumeric(InterestGross_iii) Then
            msgabox15 ("* From Income Tax Refund in Sheet OS must contain only digits from 0 to 9")
            ValidateInterestGross_15 = False
            Exit Function
        End If
    Next
    
    If Len(InterestGross_iv) > 14 Then
        msgabox15 ("* From Others in Sheet OS cannot be greater than 14 digits ")
        ValidateInterestGross_15 = False
        Exit Function
    End If
    
    For i = 1 To Len(InterestGross_iv)
        If Not IsNumeric(InterestGross_iv) Then
            msgabox15 ("* From Others in Sheet OS must contain only digits from 0 to 9")
            ValidateInterestGross_15 = False
            Exit Function
        End If
    Next
    
    
End Function

Function ValidateRentFromMachPlantBldgs_15() As Boolean
    Dim i As Long
    
    ValidateRentFromMachPlantBldgs_15 = True
    RentFromMachPlantBldgs_15 = Sheet14.Range("os.RentFromMachPlantBldgs").value
    
    If Len(RentFromMachPlantBldgs_15) > 14 Then
        msgabox15 ("* RentFromMachPlantBldgs in Sheet OS cannot be greater than 14 digits ")
        ValidateRentFromMachPlantBldgs_15 = False
        Exit Function
    End If
    
    For i = 1 To Len(RentFromMachPlantBldgs_15)
        If Not IsNumeric(RentFromMachPlantBldgs_15) Then
            msgabox15 ("* RentFromMachPlantBldgs in Sheet OS must contain only digits from 0 to 9")
            ValidateRentFromMachPlantBldgs_15 = False
            Exit Function
        End If
    Next
End Function


Function ValidateSourceDescription_OS() As Boolean
    Dim i, j As Long
    Dim rangecells As Range
    Dim OtherSourceAmount_OS, srcdes As Variant
    
    ValidateSourceDescription_OS = True
    'setTblinfo_OS
    Set rangecells = Sheet14.Range("os.SourceDescription").Cells
    ReDim SourceDescription_OS(end_os)

    
    For i = 1 To end_os
        SourceDescription_OS(i) = rangecells.item(i).value
        For j = 1 To end_os
            If j <> i And SourceDescription_OS(i) = rangecells.item(j).value Then
                If Not (rangecells.item(j).value = "(Select)" Or rangecells.item(j).value = "") Then
                    msgabox15 ("* Source Description Duplicates not allowed")
                    ValidateSourceDescription_OS = False
                    Exit Function
                End If
            End If
        Next
       
       If isdropdownblank(SourceDescription_OS(i)) Then
       msgabox15 ("* Nature is Mandatory at Sl No" & i)
                    ValidateSourceDescription_OS = False
                    Exit Function
       End If
    
    
        If Not (SourceDescription_OS(i) = "(Select)" Or SourceDescription_OS(i) = "") Then
            If WorksheetFunction.Search("-", SourceDescription_OS(i)) > 0 Then
                SourceDescription_OS(i) = Mid(SourceDescription_OS(i), 1, WorksheetFunction.Search("-", SourceDescription_OS(i)) - 1)
            Else
                'SourceDescription_OS(i + 8) = "Others"
            End If

        End If
    Next
End Function

Function ValidateSourceAmount_OS() As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim osdesc As Variant
    
    ValidateSourceAmount_OS = True
    
    Set rangecells = Sheet14.Range("os.SourceAmount").Cells
    ReDim SourceAmount_OS(end_os)

    For i = 1 To end_os
        SourceAmount_OS(i) = rangecells.item(i).value
        
        If (SourceAmount_OS(i) < 0) Then
            msgabox15 ("* 1.d  Amount at Sr. No  " & i & "  in Sheet OS  should be Non negative, no decimal, upto 99,999,999,999,999")
            ValidateSourceAmount_OS = False
            Exit Function
        End If
        
        If (Len(SourceAmount_OS(i)) > 14) Then
            msgabox15 ("* 1.d  Amount at Sr. No  " & i & "  in Sheet OS  should be Numeric,Non negative, no decimal, cannot exceeds 14 digits")
            ValidateSourceAmount_OS = False
            Exit Function
        End If
    Next
End Function


Function ValidateOthersSectionOS() As Boolean
    ValidateOthersSectionOS = True
  
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim cellrange As Variant
    Dim i As Long
    
    'setTblinfo_OS1
        
    Set rangecells = Sheet14.Range("os.othDesc").Cells
    
    ReDim OthersSourceSection_OS(end_os1)
    For i = 1 To end_os1
    cellrange = GetMergedAddressCell(rangecells, i)
    OthersSourceSection_OS(i) = Sheet14.Range(cellrange).value
          
            
            If isdropdownblank(OthersSourceSection_OS(i)) Then
              msgabox15 ("* Nature field at Sr. No  " & i + 1 & " is Mandatory in schedule OS")
              ValidateOthersSectionOS = False
              Exit Function
            End If
            
            If Not checkfieldspecialcharacter(OthersSourceSection_OS(i)) Then
                msgabox15 ("* Nature at Sr. No " & i + 1 & "   characters < > & ' " & Chr(34) & " are not allowed in schedule OS ")
                ValidateOthersSectionOS = False
                Exit Function
            End If
            
            
       
    Next
End Function

Function ValidateOthersSectionAmtOS() As Boolean
    ValidateOthersSectionAmtOS = True
    Dim rangecells, rangecells1 As Range
    Dim i As Long
    
    'setTblinfo_OS1
        
    Set rangecells = Sheet14.Range("os.OtherSections").Cells
    
    ReDim OthersSourceAmount_OS(end_os1)
    For i = 1 To end_os1
    OthersSourceAmount_OS(i) = rangecells.item(i).value
          
            
            If isdropdownblank(OthersSourceAmount_OS(i)) Then
                msgabox15 ("* Please enter Amount at Sr. No  " & i + 1 & "in schedule OS")
                ValidateOthersSectionAmtOS = False
                Exit Function
            End If
             
        
    Next
End Function



Function ValidateExpenses_15() As Boolean
    Dim i As Long
    
    ValidateExpenses_15 = True
    Expenses_15 = Sheet14.Range("os.Expenses").value
    
    'fi_WinningFrmLotteries_os = Sheet14.Range("os.WinningFrmLotteries").value
    'fii_SecXIIOth_os = Sheet14.Range("os.SecXIIOth").value
    'fiii_TotalOSGrossChargblSplRate_os = Sheet14.Range("os.TotalOSGrossChargblSplRate").value
   ' g_GrossAmtChargblNormalRate_os = Sheet14.Range("os.GrossAmtChargblNormalRate").value
    
    If Len(Expenses_15) > 14 Then
        msgabox15 ("* Expenses cannot be greater than 14 digits in schedule OS ")
        ValidateExpenses_15 = False
        Exit Function
    End If
    For i = 1 To Len(Expenses_15)
        If Not IsNumeric(Expenses_15) Then
            msgabox15 ("* Expenses must contain only digits from 0 to 9 in schedule OS")
            ValidateExpenses_15 = False
            Exit Function
        End If
    Next
End Function

Function ValidateDepreciation_15() As Boolean
    Dim i As Long
    ValidateDepreciation_15 = True
    Depreciation_15 = Sheet14.Range("os.Depreciation").value
    
    If Len(Depreciation_15) > 14 Then
        msgabox15 ("* Depreciation cannot be greater than 14 digits in schedule OS")
        ValidateDepreciation_15 = False
        Exit Function
    End If
    
    For i = 1 To Len(Depreciation_15)
        If Not IsNumeric(Depreciation_15) Then
            msgabox15 ("* Depreciation must contain only digits from 0 to 9 in schedule OS")
            ValidateDepreciation_15 = False
            Exit Function
        End If
    Next
End Function
Function ValidateAmtNotDedSec58() As Boolean
    Dim i As Long
    ValidateAmtNotDedSec58 = True
    AmtNotDedSec58 = Sheet14.Range("os.AmountnotDed58").value
    
    If Len(AmtNotDedSec58) > 14 Then
        msgabox15 ("* Amount Not Deductible u/s 58 cannot be greater than 14 digits in schedule OS")
        ValidateAmtNotDedSec58 = False
        Exit Function
    End If
    
    For i = 1 To Len(AmtNotDedSec58)
        If Not IsNumeric(AmtNotDedSec58) Then
            msgabox15 ("* Amount Not Deductible u/s 58 must contain only digits from 0 to 9 in schedule OS")
            ValidateAmtNotDedSec58 = False
            Exit Function
        End If
    Next
End Function



Function ValidateProfitCharSec59() As Boolean
    Dim i As Long
    ValidateProfitCharSec59 = True
    ProfitCharSec59 = Sheet14.Range("os.Profit59").value
    
    If Len(ProfitCharSec59) > 14 Then
        msgabox15 ("* Profits Chargeble to tax u/s 59 cannot be greater than 14 digits in schedule OS")
        ValidateProfitCharSec59 = False
        Exit Function
    End If
    
    For i = 1 To Len(ProfitCharSec59)
        If Not IsNumeric(ProfitCharSec59) Then
            msgabox15 ("* Profits Chargeble to tax u/s 59 must contain only digits from 0 to 9 in schedule OS")
            ValidateProfitCharSec59 = False
            Exit Function
        End If
    Next
End Function


Function ValidateTotDeductions_15() As Boolean
    Dim i As Long
    ValidateTotDeductions_15 = True
    TotDeductions_15 = Sheet14.Range("os.TotDeductions").value
    Deduction57iia_15 = Sheet14.Range("os_Deduction57iia").value
'    FamilyPension_15 = Int((Sheet14.Range("os_FamilyPension").value) / 3)

'     If Deduction57iia_15 > WorksheetFunction.Min(FamilyPension_15, 15000) Then
'     msgabox15 "Deductions under section 57(iia) cannot exceed lower of 1/3rd of Family Pension or Rs. 15,000/- in Schedule OS" & Chr(13)
'     ValidateTotDeductions_15 = False
'     Exit Function
'      End If

    FamilyPension_15 = WorksheetFunction.RoundUp((Sheet14.Range("os_FamilyPension").value / 3), 0)
 'by sai on 20/02/2025 for AY 2025-26
If (SheetALL.Range("bacValue").value = 1 And Deduction57iia_15 > WorksheetFunction.Min(FamilyPension_15, 25000)) Or _
 (SheetALL.Range("bacValue").value = 2 And Deduction57iia_15 > WorksheetFunction.Min(FamilyPension_15, 15000)) Then
 
   ' If Deduction57iia_15 > WorksheetFunction.Min(FamilyPension_15, 15000) Then
        'msgabox15 "Deductions under section 57(iia) cannot exceed lower of 1/3rd of Family Pension or Rs. 15,000/- in Schedule OS" & Chr(13)
        
         'msgabox15 "Deduction u/s 57(iia) cannot exceed lower of 1/3rd of Family pension or Rs. 15,000 in Point 3a(ii) of Schedule OS" & Chr(13)
         
    'modified by sai on 20/02/2025 for AY 2025-26
         msgabox15 " Deduction u/s 57(iia) cannot exceed lower of 1/3rd of Family pension or Rs. 15,000 or 25,000 in Point 3a(ii) of Schedule OS" & Chr(13)

        ValidateTotDeductions_15 = False
        Exit Function
    End If
    

    
    If Len(TotDeductions_15) > 14 Then
        msgabox15 ("* TotDeductions cannot be greater than 14 digits in schedule OS")
        ValidateTotDeductions_15 = False
        Exit Function
    End If
    
    For i = 1 To Len(TotDeductions_15)
        If Not IsNumeric(TotDeductions_15) Then
            msgabox15 ("* TotDeductions must contain only digits from 0 to 9 in schedule OS")
            ValidateTotDeductions_15 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBalanceNoRaceHorse_15() As Boolean
    Dim i As Long
    ValidateBalanceNoRaceHorse_15 = True
    BalanceNoRaceHorse_15 = Sheet14.Range("os.BalanceNoRaceHorse").value
    
    If Len(BalanceNoRaceHorse_15) > 14 Then
        msgabox15 ("* BalanceNoRaceHorse cannot be greater than 14 digits in schedule OS")
        ValidateBalanceNoRaceHorse_15 = False
        Exit Function
    End If
    
    If Not IsNumeric(BalanceNoRaceHorse_15) Then
        msgabox15 ("* BalanceNoRaceHorse must contain only digits from 0 to 9 in schedule OS")
        ValidateBalanceNoRaceHorse_15 = False
        Exit Function
    End If
End Function

Function ValidateWinLottRacePuzz_15() As Boolean
    ValidateWinLottRacePuzz_15 = True
    WinLottRacePuzz_15 = Sheet14.Range("os.WinLottRacePuzz").value
    
    If Len(WinLottRacePuzz_15) > 14 Then
        msgabox15 ("* WinLottRacePuzz cannot be greater than 14 digits in schedule OS")
        ValidateWinLottRacePuzz_15 = False
        Exit Function
    End If
    
    If Not IsNumeric(WinLottRacePuzz_15) Then
        msgabox15 ("* WinLottRacePuzz must contain only digits from 0 to 9 in schedule OS")
        ValidateWinLottRacePuzz_15 = False
        Exit Function
    End If
End Function
'Chandru Change as per OS_C7 AY24-25
Function ValidateWinOnlineGames_15() As Boolean
    ValidateWinOnlineGames_15 = True
    WinOnlineGames_15 = Sheet14.Range("os.WinOnlineGame").value
    
'added by sai on 20/02/2025 for AY 2025-26===============================================
'    WinOnlineGames_15_2aii_i = Sheet14.Range("os.WinOnlineGame2aii_i").value
'    WinOnlineGames_15_2aii_ii = Sheet14.Range("os.WinOnlineGame2aii_ii").value
'=========================================================================================
    
    
    If Len(WinOnlineGames_15) > 14 Then
        msgabox15 ("* WinOnlineGames cannot be greater than 14 digits in schedule OS")
        ValidateWinOnlineGames_15 = False
        Exit Function
    End If
    
    If Not IsNumeric(WinOnlineGames_15) Then
        msgabox15 ("* WinOnlineGames must contain only digits from 0 to 9 in schedule OS")
        ValidateWinOnlineGames_15 = False
        Exit Function
    End If
    
'     If Len(WinOnlineGames_15_2aii_i) > 14 Then
'        msgabox15 ("* WinOnlineGames cannot be greater than 14 digits in schedule OS")
'        ValidateWinOnlineGames_15 = False
'        Exit Function
'    End If
'
'    If Not IsNumeric(WinOnlineGames_15_2aii_i) Then
'        msgabox15 ("* WinOnlineGames must contain only digits from 0 to 9 in schedule OS")
'        ValidateWinOnlineGames_15 = False
'        Exit Function
'    End If
'
'    If Len(WinOnlineGames_15_2aii_ii) > 14 Then
'        msgabox15 ("* Loss from online games cannot be greater than 14 digits in schedule OS")
'        ValidateWinOnlineGames_15 = False
'        Exit Function
'    End If
'
'    If Not IsNumeric(WinOnlineGames_15_2aii_ii) Then
'        msgabox15 ("* Loss from online games must contain only digits from 0 to 9 in schedule OS")
'        ValidateWinOnlineGames_15 = False
'        Exit Function
'    End If
    
    
    
    
    
    
    
    
End Function

Function ValidateTotOthSrcNoRaceHorse_15() As Boolean
    ValidateTotOthSrcNoRaceHorse_15 = True
    TotOthSrcNoRaceHorse_15 = Sheet14.Range("os.TotOthSrcNoRaceHorse").value
    
    If Len(TotOthSrcNoRaceHorse_15) > 14 Then
        msgabox15 ("* TotOthSrcNoRaceHorse cannot be greater than 14 digits in schedule OS")
        ValidateTotOthSrcNoRaceHorse_15 = False
        Exit Function
    End If
    
    If Not IsNumeric(TotOthSrcNoRaceHorse_15) Then
        msgabox15 ("* TotOthSrcNoRaceHorse must contain only digits from 0 to 9 in schedule OS")
        ValidateTotOthSrcNoRaceHorse_15 = False
        Exit Function
    End If
End Function

Function ValidateReceipts_15() As Boolean
    Dim i As Long
    
    ValidateReceipts_15 = True
    Receipts_15 = Sheet14.Range("os.Receipts").value
    
    If Len(Receipts_15) > 14 Then
        msgabox15 ("* Receipts cannot be greater than 14 digits in schedule OS")
        ValidateReceipts_15 = False
        Exit Function
    End If
    
    For i = 1 To Len(Receipts_15)
        If Not IsNumeric(Receipts_15) Then
            msgabox15 ("* Receipts must contain only digits from 0 to 9 in schedule OS")
            ValidateReceipts_15 = False
            Exit Function
        End If
    Next
End Function

Function ValidateDeductSec57_15() As Boolean
    Dim i As Long
    
    ValidateDeductSec57_15 = True
    DeductSec57_15 = Sheet14.Range("os.DeductSec57").value
    
    
    If DeductSec57_15 > 0 Then
    If Not Sheet14.Range("os.Receipts").value > 0 Then
        msgabox15 ("* To claim deduction u/s 57 at Point-8b , receipts at Point-8a must be greater than zero in schedule OS ")
        ValidateDeductSec57_15 = False
        Exit Function
    End If
    End If
    
    If Len(DeductSec57_15) > 14 Then
        msgabox15 ("* DeductSec57 cannot be greater than 14 digits in schedule OS ")
        ValidateDeductSec57_15 = False
        Exit Function
    End If
    
    For i = 1 To Len(DeductSec57_15)
        If Not IsNumeric(DeductSec57_15) Then
            msgabox15 ("* DeductSec57 must contain only digits from 0 to 9 in schedule OS")
            ValidateDeductSec57_15 = False
            Exit Function
        End If
    Next
End Function
Function ValidateNotDeductSec58() As Boolean
    Dim i As Long
    
    ValidateNotDeductSec58 = True
    AmtNotDed = Sheet14.Range("os.notDeductSec58").value
    If Len(AmtNotDed) > 14 Then
        msgabox15 ("* Amount cannot be greater than 14 digits  in schedule OS")
        ValidateNotDeductSec58 = False
        Exit Function
    End If
    
    For i = 1 To Len(AmtNotDed)
        If Not IsNumeric(AmtNotDed) Then
            msgabox15 ("* Amount Not Deductible u/s 58 must contain only digits from 0 to 9 in schedule OS")
            ValidateNotDeductSec58 = False
            Exit Function
        End If
    Next
End Function


Function ValidateProfitSec59() As Boolean
    Dim i As Long
    
    ValidateProfitSec59 = True
    ProfitSec59 = Sheet14.Range("os.ProfitSec59").value
    If Len(ProfitSec59) > 14 Then
        msgabox15 ("* Amount Not Deductible u/s 58 cannot be greater than 14 digits in schedule OS")
        ValidateProfitSec59 = False
        Exit Function
    End If
    
    For i = 1 To Len(ProfitSec59)
        If Not IsNumeric(ProfitSec59) Then
            msgabox15 ("* Amount Not Deductible u/s 58 must contain only digits from 0 to 9 in schedule OS")
            ValidateProfitSec59 = False
            Exit Function
        End If
    Next
End Function
Function ValidateBalanceOwnRaceHorse_15() As Boolean
    ValidateBalanceOwnRaceHorse_15 = True
    BalanceOwnRaceHorse_15 = Sheet14.Range("os.BalanceOwnRaceHorse").value
    
    If Len(BalanceOwnRaceHorse_15) > 14 Then
        msgabox15 ("* BalanceOwnRaceHorse cannot be greater than 14 digits in schedule OS ")
        ValidateBalanceOwnRaceHorse_15 = False
        Exit Function
    End If
    
    If Not IsNumeric(BalanceOwnRaceHorse_15) Then
        msgabox15 ("* BalanceOwnRaceHorse must contain only digits from 0 to 9 in schedule OS")
        ValidateBalanceOwnRaceHorse_15 = False
        Exit Function
    End If
End Function

Function DefaultDividendGross_15() As String
    DividendGross_15 = "0"
End Function

Function DefaultInterestGross_15() As String
    InterestGross_15 = "0"
End Function

Function DefaultRentFromMachPlantBldgs_15() As String
    RentFromMachPlantBldgs_15 = "0"
End Function

Function DefaultOthersGross_15() As String
    OthersGross_15 = "0"
End Function

Function DefaultTotalOSGross_15() As String
    TotalOSGross_15 = "0"
End Function

Function DefaultExpenses_15() As String
    Expenses_15 = "0"
End Function

Function DefaultDepreciation_15() As String
    Depreciation_15 = "0"
End Function

Function DefaultTotDeductions_15() As String
    TotDeductions_15 = "0"
End Function

Function DefaultBalanceNoRaceHorse_15() As String
    BalanceNoRaceHorse_15 = "0"
End Function

Function DefaultWinLottRacePuzz_15() As String
    WinLottRacePuzz_15 = "0"
End Function

Function DefaultTotOthSrcNoRaceHorse_15() As String
    TotOthSrcNoRaceHorse_15 = "0"
End Function

Function DefaultReceipts_15() As String
    Receipts_15 = "0"
End Function

Function DefaultDeductSec57_15() As String
    DeductSec57_15 = "0"
End Function

Function DefaultBalanceOwnRaceHorse_15() As String
    BalanceOwnRaceHorse_15 = "0"
End Function


Function msgabox15(strmsg As String) As String
    msgavalidateSheet14 = msgavalidateSheet14 & strmsg & Chr(13)
End Function
Sub ResetOSDropDown(Status As String)
On Error Resume Next
Dim i As Long
  
    For i = 1 To (Sheet14.Range("os.SourceDescription").count)
        Range("os.SourceDescription").item(i).value = ""
    Next
    
    For i = 1 To (Sheet14.Range("os.SourceAmount").count)
        Range("os.SourceAmount").item(i).value = ""
    Next
    
      For i = 1 To (Sheet14.Range("os.PTIDescription").count)
        Range("os.PTIDescription").item(i).value = ""
    Next
    
    For i = 1 To (Sheet14.Range("os.PTIAmount").count)
        Range("os.PTIAmount").item(i).value = ""
    Next
    
      For i = 1 To (Sheet14.Range("os.SectionAct").count)
        Range("os.SectionAct").item(i).value = ""
    Next
    
    
End Sub

Sub ResetOSDropDownold(Status As String)
On Error Resume Next
Dim i As Long
  
    For i = 1 To (Sheet14.Range("os.SourceDescription").count)
        Range("os.SourceDescription").item(i).value = ""
    Next
    
      For i = 1 To (Sheet14.Range("os.PTIDescription").count)
        Range("os.PTIDescription").item(i).value = ""
    Next
    
      For i = 1 To (Sheet14.Range("os.SectionAct").count)
        Range("os.SectionAct").item(i).value = ""
    Next
    
    
End Sub


 Function ValidateSectionAmount() As Boolean
 
 ValidateSectionAmount = True
 
Dim dtaa_1a As Variant

'Newly added by Sai on 25/02/2025
'****************************************
Dim dtaa_1aiii As Variant
Dim Source_1aiii As Variant
'****************************************

Dim dtaa_1b As Variant
Dim dtaa_1c As Variant
Dim dtaa_1d As Variant
Dim dtaa_5A1ai As Variant
Dim dtaa_5A1aii As Variant
Dim dtaa_5A1aiia As Variant
Dim dtaa_5A1aiiaa As Variant
Dim dtaa_5A1aiiab As Variant
Dim dtaa_5A1aiiac As Variant
Dim dtaa_5A1aiii As Variant
Dim dtaa_5A1bA As Variant
Dim dtaa_5A1bB As Variant
Dim dtaa_5AC1ab As Variant
Dim dtaa_5AD1i As Variant
Dim dtaa_5AD1ip As Variant
Dim dtaa_5BBA As Variant
Dim dtaa_5Ea As Variant
Dim dtaa_111 As Variant
Dim dtaa_115Aca As Variant
Dim dtaa_5BBC As Variant
Dim dtaa_5BBF As Variant
Dim dtaa_5BBG As Variant
Dim dtaa_5BBDA As Variant
Dim dtaa_5BB As Variant
'Added by Ayush
Dim dtaa_5BBJ As Variant

Dim Total_dtaa As Variant

Dim Source_1a As Variant
Dim Source_1b As Variant
Dim Source_1c As Variant
Dim Source_1d As Variant
Dim Source_5A1ai As Variant
Dim Source_5A1aii As Variant
Dim Source_5A1aiia As Variant
Dim Source_5A1aiiaa As Variant
Dim Source_5A1aiiab As Variant
Dim Source_5A1aiiac As Variant
Dim Source_5A1aiii As Variant
Dim Source_5A1bA As Variant
Dim Source_5A1bB As Variant
Dim Source_5AC1ab As Variant
Dim Source_5AD1i As Variant
Dim Source_5AD1ip As Variant
Dim Source_5BBA As Variant
Dim Source_5Ea As Variant
Dim Source_111 As Variant
Dim Source_115Aca As Variant
Dim Source_5BBC As Variant
Dim Source_5BBF As Variant
Dim Source_5BBG As Variant
Dim Source_5BB As Variant
'Added by Ayush
Dim Source_5BBJ As Variant

Dim Source_5BBDA As Variant
Dim Total_Deductions As Variant
Dim Total_Specialrates As Variant

Dim TemptotalBBDA As Variant
Dim Tempincomefromlottery As Variant
'Added by Ayush
Dim TempincomefromOnlineGames As Variant


Dim PTI_5BB_OS As Variant
'Added by Ayush
Dim PTI_5BBJ_OS As Variant

Dim PTI_5BBE_OS As Variant
Dim PTI_5A1ai_OS As Variant
Dim PTI_5A1aii_OS As Variant
Dim PTI_5A1aiia_OS As Variant
Dim PTI_5A1aiiaa_OS As Variant
Dim PTI_5A1aiiab_OS As Variant
Dim PTI_5A1aiiac_OS As Variant
Dim PTI_5A1aiii_OS As Variant
Dim PTI_5A1bA_OS As Variant
Dim PTI_5A1bB_OS As Variant
Dim PTI_5AC1ab_OS As Variant
Dim PTI_5ACA1a_OS As Variant
Dim PTI_5AD1i_OS As Variant
Dim PTI_5AD1iP_OS As Variant
Dim PTI_5BBA_OS As Variant
Dim PTI_5BBC_OS As Variant
Dim PTI_5BBDA_OS As Variant
Dim PTI_5BBF_OS As Variant
Dim PTI_5BBG_OS As Variant
Dim PTI_5Ea_OS As Variant
Dim PTI_1_OS As Variant

Dim PTI_5BB_DTAA As Variant
'Added by Ayush
Dim PTI_5BBJ_DTAA As Variant

Dim PTI_5BBE_DTAA As Variant
Dim PTI_5A1ai_DTAA As Variant
Dim PTI_5A1aii_DTAA As Variant
Dim PTI_5A1aiia_DTAA As Variant
Dim PTI_5A1aiiaa_DTAA As Variant
Dim PTI_5A1aiiab_DTAA As Variant
Dim PTI_5A1aiiac_DTAA As Variant
Dim PTI_5A1aiii_DTAA As Variant
Dim PTI_5A1bA_DTAA As Variant
Dim PTI_5A1bB_DTAA As Variant
Dim PTI_5AC1ab_DTAA As Variant
Dim PTI_5ACA1a_DTAA As Variant
Dim PTI_5AD1i_DTAA As Variant
Dim PTI_5AD1iP_DTAA As Variant
Dim PTI_5BBA_DTAA As Variant
Dim PTI_5BBC_DTAA As Variant
Dim PTI_5BBDA_DTAA As Variant
Dim PTI_5BBF_DTAA As Variant
Dim PTI_5BBG_DTAA As Variant
Dim PTI_5Ea_DTAA As Variant
Dim PTI_1_DTAA As Variant



dtaa_1a = Sheet14.Range("DTAA_56i_OS").value

'For AY 2025-26 by sai on 25/02/2025
'*********************************************************
dtaa_1aiii = Sheet14.Range("DTAA_56i_f_OS").value
'*********************************************************

dtaa_1b = Sheet14.Range("DTAA_56_OS").value
dtaa_1c = Sheet14.Range("DTAA_1c_OS").value
dtaa_1d = Sheet14.Range("DTAA_1d_OS").value

dtaa_5A1ai = Sheet14.Range("DTAA_5A1ai_OS").value
'Chandru
dtaa_5A1aA = Sheet14.Range("DTAA_5A1aA_OS").value

dtaa_5A1aii = Sheet14.Range("DTAA_5A1aii_OS").value
dtaa_5A1aiia = Sheet14.Range("DTAA_5A1aiia_OS").value
dtaa_5A1aiiaa = Sheet14.Range("DTAA_5A1aiiaa_OS").value
dtaa_5A1aiiab = Sheet14.Range("DTAA_5A1aiiab_OS").value
dtaa_5A1aiiac = Sheet14.Range("DTAA_5A1aiiac_OS").value
dtaa_5A1aiii = Sheet14.Range("DTAA_5A1aiii_OS").value
dtaa_5A1bA = Sheet14.Range("DTAA_5A1bA_OS").value
dtaa_5A1bB = Sheet14.Range("DTAA_5A1bB_OS").value
dtaa_5AC1ab = Sheet14.Range("DTAA_5AC1ab_OS").value
dtaa_5AD1i = Sheet14.Range("DTAA_5A1D1i_OS").value
dtaa_5AD1ip = Sheet14.Range("DTAA_5AD1ip_OS").value
dtaa_5BBA = Sheet14.Range("DTAA_5BBA_OS").value
dtaa_5Ea = Sheet14.Range("DTAA_5Eaa_OS").value
dtaa_111 = Sheet14.Range("DTAA_111_OS").value
dtaa_115Aca = Sheet14.Range("DTAA_5ACA1a_OS").value
dtaa_5BBC = Sheet14.Range("DTAA_5BBC_OS").value
dtaa_5BBF = Sheet14.Range("DTAA_5BBF_OS").value
dtaa_5BBG = Sheet14.Range("DTAA_5BBG_OS").value
dtaa_5BB = Sheet14.Range("DTAA_5BB_OS").value
'Added by Ayush
dtaa_5BBJ = Sheet14.Range("DTAA_5BBJ_OS").value

dtaa_5BBDA = Sheet14.Range("DTAA_5BBDA_OS").value

'Total_dtaa = Sheet14.Range("os.TotalOSGross").value

Source_1a = Sheet14.Range("Os.1a").value
Source_1b = Sheet14.Range("Os.1b").value
Source_1c = Sheet14.Range("os.RentFromMachPlantBldgs").value
Source_1d = Sheet14.Range("os.Totalsec56").value

Source_5A1ai = Sheet14.Range("Source_5A1ai").value
'Chandru
Source_5A1aA = Sheet14.Range("Source_5A1aA").value

'********************************************** For AY 2025-26 by Sai
Source_1aiii = Sheet14.Range("Os.1aiii").value
'*******************************************************************


Source_5A1aii = Sheet14.Range("Source_5A1aii_OS").value
Source_5A1aiia = Sheet14.Range("Source_5A1aiia_OS").value
Source_5A1aiiaa = Sheet14.Range("Source_5A1aiiaa_OS").value
Source_5A1aiiab = Sheet14.Range("Source_5A1aiiab_OS").value
Source_5A1aiiac = Sheet14.Range("Source_5A1aiiac_OS").value
Source_5A1aiii = Sheet14.Range("Source_5A1aiii_OS").value
Source_5A1bA = Sheet14.Range("Source_5A1ba_OS").value
Source_5A1bB = Sheet14.Range("Source_5A1bB_OS").value
Source_5AC1ab = Sheet14.Range("Source_5A1C1ab_OS").value
Source_5AD1i = Sheet14.Range("Source_5AD1i_OS").value
Source_5AD1ip = Sheet14.Range("Source_5AD1iP_OS").value
Source_5BBA = Sheet14.Range("Source_5BBA_OS").value
Source_5Ea = Sheet14.Range("Source_5Ea_OS").value
Source_111 = Sheet14.Range("Source_111_OS").value
Source_115Aca = Sheet14.Range("Source_5ACA1a_OS").value
Source_5BBC = Sheet14.Range("Source_5BBC_OS").value
Source_5BBF = Sheet14.Range("Source_5BBF_OS").value
Source_5BBG = Sheet14.Range("Source_5BBG_OS").value
Source_5BB = Sheet14.Range("os.WinLottRacePuzz").value
'Added by Ayush
Source_5BBJ = Sheet14.Range("os.WinOnlineGame").value

Source_5BBDA = Sheet14.Range("Source_5BBDA_OS").value
Total_Deductions = Sheet14.Range("os.DTAA_Amt").value
Total_Specialrates = Sheet14.Range("os.DTAA_Amt").value

TemptotalBBDA = Sheet14.Range("tempBBDA").value
Tempincomefromlottery = Sheet14.Range("Temp_IncomeFromLottery").value
'Added by Ayush
TempincomefromOnlineGames = Sheet14.Range("Temp_IncomeFromOnlineGame").value


PTI_5BB_OS = Range("OS_PTI_5BB").value
PTI_5BBE_OS = Range("OS_PTI_5BBE").value
PTI_5A1ai_OS = Range("Os_PTI_5A1ai").value
'Chandru
PTI_5A1aA_OS = Range("Os_PTI_5A1aA").value

PTI_5A1aii_OS = Range("Os_PTI_5A1aii").value
PTI_5A1aiia_OS = Range("Os_PTI_5A1aiia").value
PTI_5A1aiiaa_OS = Range("Os_PTI_5A1aiiaa").value
PTI_5A1aiiab_OS = Range("Os_PTI_5A1aiiab").value
PTI_5A1aiiac_OS = Range("Os_PTI_5A1aiiac").value
PTI_5A1aiii_OS = Range("Os_PTI_5A1aiii").value
PTI_5A1bA_OS = Range("Os_PTI_5A1bA").value
PTI_5A1bB_OS = Range("Os_PTI_5A1bB").value
PTI_5AC1ab_OS = Range("OS_PTI_5AC1ab").value
PTI_5ACA1a_OS = Range("OS_PTI_5ACA1a").value
PTI_5AD1i_OS = Range("OS_PTI_5AD1i").value
PTI_5AD1iP_OS = Range("OS_PTI_5AD1iP").value
PTI_5BBA_OS = Range("OS_PTI_5BBA").value
PTI_5BBC_OS = Range("OS_PTI_5BBC").value
PTI_5BBDA_OS = Range("OS_PTI_5BBDA").value
PTI_5BBF_OS = Range("OS_PTI_5BBF").value
PTI_5BBG_OS = Range("OS_PTI_5BBG").value
PTI_5Ea_OS = Range("OS_PTI_5Ea").value
PTI_1_OS = Range("OS_PTI_1").value

PTI_5BB_DTAA = Range("Os_PTI_5BB_DTAA").value


PTI_5BBE_DTAA = Range("Os_PTI_5BBE_DTAA").value
PTI_5A1ai_DTAA = Range("Os_PTI_5A1ai_DTAA").value
'Chandru
PTI_5A1aA_DTAA = Range("Os_PTI_5A1aA_DTAA").value

PTI_5A1aii_DTAA = Range("Os_PTI_5A1aii_DTAA").value
PTI_5A1aiia_DTAA = Range("Os_PTI_5A1aiia_DTAA").value
PTI_5A1aiiaa_DTAA = Range("Os_PTI_5A1aiiaa_DTAA").value
PTI_5A1aiiab_DTAA = Range("Os_PTI_5A1aiiab_DTAA").value
PTI_5A1aiiac_DTAA = Range("Os_PTI_5A1aiiac_DTAA").value
PTI_5A1aiii_DTAA = Range("Os_PTI_5A1aiii_DTAA").value
PTI_5A1bA_DTAA = Range("Os_PTI_5A1aibA_DTAA").value
PTI_5A1bB_DTAA = Range("Os_PTI_5A1aibB_DTAA").value
PTI_5AC1ab_DTAA = Range("Os_PTI_5AC1ab_DTAA").value
PTI_5ACA1a_DTAA = Range("Os_PTI_5ACA1a_DTAA").value
PTI_5AD1i_DTAA = Range("Os_PTI_5AD1i_DTAA").value
PTI_5AD1iP_DTAA = Range("Os_PTI_5AD1iP_DTAA").value
PTI_5BBA_DTAA = Range("Os_PTI_5BBA_DTAA").value
PTI_5BBC_DTAA = Range("Os_PTI_5BBC_DTAA").value
PTI_5BBDA_DTAA = Range("Os_PTI_5BBDA_DTAA").value
PTI_5BBF_DTAA = Range("Os_PTI_5BBF_DTAA").value
PTI_5BBG_DTAA = Range("Os_PTI_5BBG_DTAA").value
PTI_5Ea_DTAA = Range("Os_PTI_5Ea_DTAA").value
PTI_1_DTAA = Range("Os_PTI_1_DTAA").value

'If (TemptotalBBDA <> (Source_5BBDA + Sheet14.Range("OS_PTI_5BBDA").value - dtaa_5BBDA - Sheet14.Range("Os_PTI_5BBDA_DTAA").value)) Then
'msgabox15 ("* The total of all the quarter of Dividend Income u/s 115BBDA must be equal to (2d(xv) + 115BBDA dropdown of 2e-DTAA 115BBDA - DTAA 115BBDA PTI) of Schedule OS " & Chr(13))
'ValidateSectionAmount = False
'Exit Function
'End If

'If (Tempincomefromlottery <> Sheet14.Range("os.WinningFrmLotteries").value) Then
'dpk115bbissue
If (Tempincomefromlottery <> (Source_5BB + Sheet14.Range("OS_PTI_5BB").value - dtaa_5BB - Sheet14.Range("Os_PTI_5BB_DTAA").value)) Then
msgabox15 ("* The total of all the quarter of Income from Lottery must be equal to 2a of Schedule OS-DTAA 115BB" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If


''''''''''''
Dim TempincomeFromaiii As Variant 'tempBBDA_aiii

TempincomeFromaiii = Range("tempBBDA_aiii").value
If (TempincomeFromaiii <> (Source_1aiii - dtaa_1aiii)) Then
    msgabox15 ("* The quarterly break up of Dividend Income should be equal to [1a(iii ) Less DTAA of 1a(i) subject to TRC flag  attributable to Sl.No. 1a(iii)] of Schedule OS" & Chr(13))
    ValidateSectionAmount = False
    Exit Function
End If
 
'''''''''''''
'Added by Ayush
If (TempincomefromOnlineGames <> (Source_5BBJ + Sheet14.Range("OS_PTI_5BBJ").value - dtaa_5BBJ - Sheet14.Range("Os_PTI_5BBJ_DTAA").value)) Then
msgabox15 ("* The total of all the quarter of Income from Online Games must be equal to 2aii of Schedule OS-DTAA 115BBJ" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_111 > Source_111) Then
msgabox15 ("* Please ensure that Amount of income u/s 111 in DTAA table cannot exceed the Income filled in Pt. 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_5BB > Source_5BB) Then
msgabox15 ("* Please ensure that Amount of income u/s 115BB in DTAA table cannot exceed the Income filled in Pt. 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

'Chandru
If (dtaa_5BBJ > Source_5BBJ) Then
msgabox15 ("* Please ensure that Amount of income u/s 115BB in DTAA table cannot exceed the Income filled in Pt. 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_5BBC > Source_5BBC) Then
msgabox15 ("* Please ensure that Amount of income u/s 115BBC in DTAA table cannot exceed the Income filled in Pt. 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_5BBF > Source_5BBF) Then
msgabox15 ("* Please ensure that Amount of income u/s 115BBF in DTAA table cannot exceed the Income filled in Pt. 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_5BBG > Source_5BBG) Then
msgabox15 ("* Please ensure that Amount of income u/s 115BBG in DTAA table cannot exceed the Income filled in Pt. 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_5BBDA > Source_5BBDA) Then
msgabox15 ("* Please ensure that Amount of income u/s 115BBDA in DTAA table cannot exceed the Income filled in Pt. 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_115Aca > Source_115Aca) Then
msgabox15 ("* Please ensure that Amount of income u/s 115ACA(1)a in DTAA table cannot exceed the Income filled in Pt. 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If


If (dtaa_1a > Source_1a) Then
msgabox15 ("* Please ensure that Amount of income u/s 56(2)(i)-Dividend in DTAA table cannot exceed the Income filled in Pt. 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

'*******************For AY 2025-26 by Sai
If (dtaa_1aiii > Source_1aiii) Then
msgabox15 ("* Please ensure that Amount of income u/s 56(2)(i)-Dividend u/s 2(22)(f) in DTAA table cannot exceed the Income filled in Pt. 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If
'****************************************


If dtaa_1b > 0 Then
If (dtaa_1b > Source_1b) Then
msgabox15 ("* Please ensure that Amount of income u/s 56(2)-Interest in DTAA table cannot exceed the Income filled in 2d.Interest of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If
End If

If (dtaa_1c > Source_1c) Then
msgabox15 ("* Please ensure that Amount of income 56(2)(iii)-Rental income from machinery, plants, buildings in DTAA table cannot exceed the Income filled in Pt. 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_1d > Source_1d) Then
msgabox15 ("* Please ensure that Amount of income Income under section 56(2)(x)  in DTAA table cannot exceed the Income filled in 2d.Interest of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If


If (dtaa_5A1ai > Source_5A1ai) Then
msgabox15 ("* Please ensure that Amount of income u/s 115A(1)(a)(i) in DTAA table does not exceed the income filled in corresponding section in pt 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

'Chandru
If (dtaa_5A1aA > Source_5A1aA) Then
msgabox15 ("* Please ensure that Amount of income u/s 115A(1)(a)(A) in DTAA table does not exceed the income filled in corresponding section in pt 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_5A1aii > Source_5A1aii) Then
msgabox15 ("* Please ensure that Amount of income u/s 115A(1)(a)(ii) in DTAA table does not exceed the income filled in corresponding section in pt 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_5A1aiia > Source_5A1aiia) Then
msgabox15 ("* Please ensure that Amount of income u/s 115A(1)(a)(iia) in DTAA table does not exceed the income filled in corresponding section in pt 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_5A1aiiaa > Source_5A1aiiaa) Then
msgabox15 ("* Please ensure that Amount of income u/s 115A(1)(a)(iiaa) in DTAA table does not exceed the income filled in corresponding section in pt 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_5A1aiiab > Source_5A1aiiab) Then
msgabox15 ("* Please ensure that Amount of income u/s 115A(1)(a)(iiab) in DTAA table does not exceed the income filled in corresponding section in pt 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_5A1aiiac > Source_5A1aiiac) Then
msgabox15 ("* Please ensure that Amount of income u/s 115A(1)(a)(iiac) in DTAA table does not exceed the income filled in corresponding section in pt 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_5A1aiii > Source_5A1aiii) Then
msgabox15 ("* Please ensure that Amount of income u/s 115A(1)(a)(iii) in DTAA table does not exceed the income filled in corresponding section in pt 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_5A1bA > Source_5A1bA) Then
msgabox15 ("* Please ensure that Amount of income u/s 115A(1)(b)(A) in DTAA table does not exceed the income filled in corresponding section in pt 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_5A1bB > Source_5A1bB) Then
msgabox15 ("* Please ensure that Amount of income u/s 115A(1)(b)(B) in DTAA table does not exceed the income filled in corresponding section in pt 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_5AC1ab > Source_5AC1ab) Then

msgabox15 ("* Please ensure that Amount of income u/s 115AC(1)(a&b) in DTAA table cannot exceed the Income filled in Pt. 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_5AD1i > Source_5AD1i) Then
msgabox15 ("* Please ensure that Amount of income u/s 115AD(1)(i) in DTAA table does not exceed the income filled in corresponding section in pt 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_5AD1ip > Source_5AD1ip) Then
msgabox15 ("* Please ensure that Amount of income u/s 115AD(1)(i) in DTAA table does not exceed the income filled in corresponding section in pt 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_5BBA > Source_5BBA) Then
msgabox15 ("* Please ensure that Amount of income u/s 115BBA in DTAA table does not exceed the income filled in corresponding section in pt 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (dtaa_5Ea > Source_5Ea) Then
msgabox15 ("* Please ensure that Amount of income u/s 115E(a) in DTAA table does not exceed the income filled in corresponding section in pt 2d of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5BB_DTAA > PTI_5BB_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115BB in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

''Added by Ayush
'If (PTI_5BBJ_DTAA > PTI_5BBJ_OS) Then
'msgabox15 ("* Please ensure that Amount of income u/s PTI_115BBJ in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
'ValidateSectionAmount = False
'Exit Function
'End If

If (PTI_5BBE_DTAA > PTI_5BBE_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115BBE in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5A1ai_DTAA > PTI_5A1ai_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115A1ai in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

'Chandru
If (PTI_5A1aA_DTAA > PTI_5A1aA_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115A1aA in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If


If (PTI_5A1aii_DTAA > PTI_5A1aii_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115A1aii in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5A1aiia_DTAA > PTI_5A1aiia_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115A1aiia in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5A1aiiaa_DTAA > PTI_5A1aiiaa_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115A1aiiaa in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5A1aiiab_DTAA > PTI_5A1aiiab_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115A1aiiab in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5A1aiiac_DTAA > PTI_5A1aiiac_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115A1aiiac in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5A1aiii_DTAA > PTI_5A1aiii_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115A1aiii in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5A1bA_DTAA > PTI_5A1bA_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115A1bA in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5A1bB_DTAA > PTI_5A1bB_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115A1bB in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5AC1ab_DTAA > PTI_5AC1ab_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115AC1ab in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5ACA1a_DTAA > PTI_5ACA1a_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115ACA1a in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5AD1i_DTAA > PTI_5AD1i_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115AD1i in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5AD1iP_DTAA > PTI_5AD1iP_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115AD1iP in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5BBA_DTAA > PTI_5BBA_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115BBA in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5BBC_DTAA > PTI_5BBC_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115BBC in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5BBDA_DTAA > PTI_5BBDA_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115BBDA in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5BBF_DTAA > PTI_5BBF_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115BBF in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5BBG_DTAA > PTI_5BBG_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115BBG in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_5Ea_DTAA > PTI_5Ea_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_115Ea in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

If (PTI_1_DTAA > PTI_1_OS) Then
msgabox15 ("* Please ensure that Amount of income u/s PTI_111 in DTAA table cannot exceed the Income filled in corresponding section in pt 2e of Schedule OS" & Chr(13))
ValidateSectionAmount = False
Exit Function
End If

End Function

Function Validateos_PTIDescription() As Boolean
    Validateos_PTIDescription = True
    Dim i As Long
    Dim j As Long
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim cellrange As Variant
    
    
    ReDim os_PTIDescription(end_PTIDes)
    Set rangecells = Sheet14.Range("OS_PTISectionCode").Cells
    Set rangecells1 = Sheet14.Range("os.PTIAmount").Cells
    
    For i = 1 To end_PTIDes
    cellrange = GetMergedAddressCell(rangecells, i)
    os_PTIDescription(i) = Sheet14.Range(cellrange).value
    Next
    
    For j = 1 To end_PTIIncm
        ReDim os_PTIDescription(end_PTIIncm)
        cellrange = GetMergedAddressCell(rangecells, j)
        os_PTIDescription(j) = Sheet14.Range(cellrange).value
        
        
            If isdropdownblank(os_PTIDescription(j)) Then
                msgabox15 ("* Nature of Pass through income at Sr. No  " & j & "  is Mandatory in Schedule OS")
                Validateos_PTIDescription = False
                Exit Function
            End If
       
    Next
    
End Function

Function Validateos_PTIAmount() As Boolean
    Validateos_PTIAmount = True
    Dim i As Long
    Dim j As Long
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim cellrange As Variant
    
    'setTblinfo_PTIDes
    Set rangecells = Sheet14.Range("os.PTIAmount")
    Set rangecells1 = Sheet14.Range("OS_PTISectionCode").Cells
     
    ReDim os_PTIAmount(end_PTIDes)
    ReDim os_PTIDescription(end_PTIDes)
    
    For i = 1 To end_PTIDes
    cellrange = GetMergedAddressCell(rangecells1, i)
    os_PTIDescription(i) = Sheet14.Range(cellrange).value
    os_PTIAmount(i) = rangecells.item(i).value
    
    
        If Not chkMandatory(os_PTIAmount(i)) Then
            msgabox15 ("* Please enter amount in Pass through income at Point 2e.")
            Validateos_PTIAmount = False
            Exit Function
        End If
    
    If Len(os_PTIAmount(i)) > 14 Then
        msgabox15 ("* Amount in Pass through income cannot be greater than 14 digits in Schedule OS")
        Validateos_PTIAmount = False
        Exit Function
    End If
    
    For j = 1 To Len(os_PTIAmount(i))
        If Not IsNumeric(Mid(os_PTIAmount(i), j, 1)) Then
           msgabox15 ("* Amount in Pass through income cannot must contain only digits from 0 to 9 in Schedule OS")
            Validateos_PTIAmount = False
            Exit Function
        End If
    Next
    Next
End Function

Sub setTblinfo_PTIDes()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 Dim dtaa_section As Variant
 ccount = 0
 mIntCells = Sheet14.Range("os.PTIDescription").count
 Set rangecells = Sheet14.Range("os.PTIDescription").Cells
 For mIntCtr = 1 To mIntCells
dtaa_section = rangecells.item(mIntCtr).value

        If isdropdownblank(dtaa_section) Then
        dtaa_section = ""
        End If
    
     If Not dtaa_section = "" Then
         ccount = ccount + 1
     End If
 Next
 end_PTIDes = ccount
rngname_PTIOS = "os.PTIDescription;os.PTIAmount;os_PTI_check;OS_PTI_Amount;OS_PTISectionCode;"
End Sub

Sub setTblinfo_PTIIncm()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet14.Range("os.PTIAmount").count
 Set rangecells = Sheet14.Range("os.PTIAmount").Cells
 For mIntCtr = 1 To mIntCells
     If rangecells.item(mIntCtr).value <> "" Then
         ccount = ccount + 1
     End If
 Next
 end_PTIIncm = ccount
End Sub

'Newly added by sai for AY 2025-26
Function ValidateBBDA_1aiii() As Boolean
    ValidateBBDA_1aiii = True
     
'    BBDA_1aiii_Upto15Of6 = Sheet14.Range("BBDA_1aiii.Upto15Of6").value
'    BBDA_1aiii_Upto15Of9 = Sheet14.Range("BBDA_1aiii.Upto15Of9").value
    BBDA_1aiii_Up16Of9To15Of12 = Sheet14.Range("BBDA_1aiii.Up16Of9To15Of12").value
    BBDA_1aiii_Up16Of12To15Of3 = Sheet14.Range("BBDA_1aiii.Up16Of12To15Of3").value
    BBDA_1aiii_Up16Of3To31Of3 = Sheet14.Range("BBDA_1aiii.Up16Of3To31Of3").value
    
'   If BBDA_1aiii_Upto15Of6 <> "" Then
'
'    If Len(BBDA_1aiii_Upto15Of6) > 14 Then
'        msgabox15 ("* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999""" & Chr(13))
'        ValidateBBDA_1aiii = False
'        Exit Function
'    End If
'
'    If Not IsNumeric(BBDA_1aiii_Upto15Of6) Then
'        msgabox15 ("* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999""" & Chr(13))
'        ValidateBBDA_1aiii = False
'        Exit Function
'    End If
'
'    End If
    
    
'   If BBDA_1aiii_Upto15Of9 <> "" Then
'    If Len(BBDA_1aiii_Upto15Of9) > 14 Then
'        msgabox15 ("* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999""" & Chr(13))
'        ValidateBBDA_1aiii = False
'        Exit Function
'    End If
    
'    If Not IsNumeric(BBDA_1aiii_Upto15Of9) Then
'        msgabox15 ("* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999""" & Chr(13))
'        ValidateBBDA_1aiii = False
'        Exit Function
'    End If
'    End If
    
    If BBDA_1aiii_Up16Of9To15Of12 <> "" Then
    If Len(BBDA_1aiii_Up16Of9To15Of12) > 14 Then
        msgabox15 ("* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999""" & Chr(13))
        ValidateBBDA_1aiii = False
        Exit Function
    End If
    
    If Not IsNumeric(BBDA_1aiii_Up16Of9To15Of12) Then
        msgabox15 ("* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999""" & Chr(13))
        ValidateBBDA_1aiii = False
        Exit Function
    End If
    End If
    
    If BBDA_1aiii_Up16Of12To15Of3 <> "" Then
    If Len(BBDA_1aiii_Up16Of12To15Of3) > 14 Then
        msgabox15 ("* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999""" & Chr(13))
        ValidateBBDA_1aiii = False
        Exit Function
    End If
    
    If Not IsNumeric(BBDA_1aiii_Up16Of12To15Of3) Then
        msgabox15 ("* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999""" & Chr(13))
        ValidateBBDA_1aiii = False
        Exit Function
    End If
    End If
    
    If BBDA_1aiii_Up16Of3To31Of3 <> "" Then
    If Len(BBDA_1aiii_Up16Of3To31Of3) > 14 Then
        msgabox15 ("* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999""" & Chr(13))
        ValidateBBDA_1aiii = False
        Exit Function
    End If
    
    If Not IsNumeric(BBDA_1aiii_Up16Of3To31Of3) Then
        msgabox15 ("* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999""" & Chr(13))
        ValidateBBDA_1aiii = False
        Exit Function
    End If
    End If

End Function
