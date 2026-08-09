Attribute VB_Name = "SchTR_FA"
Option Explicit

'TR

Public Country_TR As Variant
'Public Country_FSI As Variant
Public CountryName_TR As Variant
'Public CountryName_FSI As Variant
Public TIN_TR As Variant
Public TaxPaidOutsideIndia_TR As Variant
Public TaxReliefOutsideIndia_TR As Variant
Public ReliefClaimedUsSection_TR As Variant

Public TotalTaxPaidOutsideIndia_TR As Variant
Public TaxReliefOutsideIndiaDTAA_TR As Variant
Public TaxReliefOutsideIndiaNotDTAA_TR As Variant
Public TaxPaidOutsideIndFlg_TR As Variant
Public AmtTaxRefunded_TR As Variant
Public AssmtYrTaxRelief_TR As Variant

Public rngname_TR As Variant
Public end_TR As Variant

Public rngname_TR2 As Variant
Public end_TR2 As Variant

Dim msgValidateSheetTR As String

'FA1
Public CountryCodeA1_FA As Variant
Public CountryNameA1_FA As Variant
Public BankNameA1_FA As Variant
Public BankAddA1_FA As Variant
Public ZipCodeA1_FA As Variant
Public AccountNameA1_FA As Variant
Public StatusA1_FA As Variant
Public AccountOpeningDateA1_FA As Variant
Public PeakBalA1_FA As Variant
Public CloseBalA1_FA As Variant
Public GrossInterest_A1_FA As Variant


'FA2
Public CountryCodeA2_FA As Variant
Public CountryNameA2_FA As Variant
Public BankNameA2_FA As Variant
Public BankAddA2_FA As Variant
Public ZipCodeA2_FA As Variant
Public AccountNameA2_FA As Variant
Public StatusA2_FA As Variant
Public AccountOpeningDateA2_FA As Variant
Public PeakBalA2_FA As Variant
Public CloseBalA2_FA As Variant
Public GrossInterest_A2_FA As Variant
Public GrossInterestNature_A2_FA As Variant


'FA3
Public CountryCodeA3_FA As Variant
Public CountryNameA3_FA As Variant
Public BankNameA3_FA As Variant
Public BankAddA3_FA As Variant
Public ZipCodeA3_FA As Variant
Public NatureOfEntityA3_FA As Variant
Public AccountOpeningDateA3_FA As Variant
Public Initialvalue_Investment As Variant
Public PeakBalA3_FA As Variant
Public CloseBalA3_FA As Variant
Public TotalGrossAmount_A3_FA As Variant
Public TotalGrossProceeds_A3_FA As Variant


'FA4
Public CountryCodeA4_FA As Variant
Public CountryNameA4_FA As Variant
Public BankNameA4_FA As Variant
Public BankAddA4_FA As Variant
Public ZipCodeA4_FA As Variant
Public AccountOpeningDateA4_FA As Variant
Public TotalGrossAmount_A4_FA As Variant
Public TheCashValue_A4_FA As Variant


Public CountryCodeB_FA As Variant
Public CountryNameB_FA As Variant
Public ZipCodeB_FA As Variant
Public EntityNatureB_FA As Variant
Public EntityNameB_FA As Variant
Public EntityAddB_FA As Variant
Public TotalInvB_FA As Variant
Public NatureOfInterestB_FA As Variant
Public DateHeldB_FA As Variant
Public IncomeIncurredInAccB_FA As Variant
Public NatureOfIncomeB_FA As Variant
Public AmountB_FA As Variant
Public Schedule_B_FA As Variant
Public ItemNumber_B_FA As Variant
Public InvTotalB_FA As Variant

Public CountryCodeC_FA As Variant
Public CountryNameC_FA As Variant
Public ZipCodeC_FA As Variant
Public PropertyAddC_FA As Variant
Public OwnershipC_FA As Variant
Public DateOfAcquistionC_FA As Variant
Public TotalInvC_FA As Variant
Public IncomeFromPropC_FA As Variant
Public NatureOfIncomeC_FA As Variant
Public AmountC_FA As Variant
Public ScheduleOfferedC_FA As Variant
Public ItemNumofScheduleC_FA As Variant
Public InvTotalC_FA As Variant

Public CountryCodeF_FA As Variant
Public CountryNameF_FA As Variant
Public ZipCodeF_FA As Variant
Public NameOfTrustF_FA As Variant
Public AddressOfTrustF_FA As Variant
Public NameOfOtherTrusteesF_FA As Variant
Public AddressOfOtherTrusteesF_FA As Variant
Public NameOfSettlorF_FA As Variant
Public AddressOfSettlorF_FA As Variant
Public NameOfBeneficiariesF_FA As Variant
Public AddressOfBeneficiariesF_FA As Variant
Public DateSinceHeldF_FA As Variant
Public WheteherTaxableIncomeF_FA As Variant
Public IncomeDerivedF_FA As Variant
Public AmountF_FA As Variant
Public ScheduleOfferedF_FA As Variant
Public ItemNumOfScheduleF_FA As Variant

Public CountryCodeD_FA As Variant
Public CountryNameD_FA As Variant
Public ZipCodeD_FA As Variant
Public AssetNatureD_FA As Variant
Public OwnershipD_FA As Variant
Public DateofAcquistionD_FA As Variant
Public IncomeFrmAsset_FA As Variant
Public NatureOfIncome_FA As Variant
Public AmountD_FA As Variant
Public ScheduleofferedD_FA As Variant
Public ItemNumSchedule_FA As Variant
Public TotalInvD_FA As Variant
Public InvTotalD_FA As Variant

Public InstNameE_FA As Variant
Public InstAddE_FA As Variant

Public CountryCodeE_FA As Variant
Public CountryNameE_FA As Variant
Public ZipCodeE_FA As Variant

Public AccountNameE_FA As Variant
Public InstitutionAccountNumberE_FA As Variant
Public AccountTaxableE_FA As Variant
Public IncomeAccuredE_FA As Variant
Public AmountE_FA As Variant
Public ScheduleOfferedE_FA As Variant
Public ItemNumScheduleE_FA As Variant
Public PeakBalInvE_FA As Variant
Public InvPBalTotE_FA As Variant

Public CountryCodeG_FA As Variant
Public CountryNameG_FA As Variant
Public ZipCodeG_FA As Variant
Public PersonNameG_FA As Variant
Public PersonAddressG_FA As Variant
Public IncomeDerivedG_FA As Variant
Public NatureIncomeG_FA As Variant
Public WhetherTaxableIncomeG_FA As Variant
Public AmountG_FA As Variant
Public ScheduleOfferedG_FA As Variant
Public ItemNumSchG_FA As Variant

Public rngname_A_FA As Variant
Public end_A_FA As Variant

Public rngname_A1_FA As Variant
Public end_A1_FA As Variant
Public end_A11_FA As Variant

Public rngname_A2_FA As Variant
Public end_A2_FA As Variant
Public end_A21_FA As Variant

Public rngname_A3_FA As Variant
Public end_A3_FA As Variant
Public end_A31_FA As Variant

Public rngname_A4_FA As Variant
Public end_A4_FA As Variant
Public end_A41_FA As Variant

Public rngname_A_FA2 As Variant
Public end_A_FA2 As Variant

Public incBy_A_FA As Variant
Dim msgValidateSheetAFA As String


Dim msgValidateSheetA1FA As String
Dim msgValidateSheetA2FA As String
Dim msgValidateSheetA3FA As String
Dim msgValidateSheetA4FA As String

Public rngname_B_FA As Variant
Public end_B_FA As Variant

Public rngname_B_FA2 As Variant
Public end_B_FA2 As Variant

Dim msgValidateSheetBFA As String

Public rngname_C_FA As Variant
Public end_C_FA As Variant

Public rngname_C_FA2 As Variant
Public end_C_FA2 As Variant

Dim msgValidateSheetCFA As String
Public rngname_D_FA As Variant
Public end_D_FA As Variant

Public rngname_D_FA2 As Variant
Public end_D_FA2 As Variant

Dim msgValidateSheetDFA As String
Dim msgValidateSheetFFA As String

Public rngname_F_FA As Variant
Public end_F_FA As Variant

Public rngname_F_FA2 As Variant
Public end_F_FA2 As Variant

Public rngname_E_FA As Variant
Public end_E_FA As Variant

Public rngname_E_FA2 As Variant
Public end_E_FA2 As Variant

Public rngname_G_FA As Variant
Public end_G_FA As Variant

Public rngname_G_FA2 As Variant
Public end_G_FA2 As Variant

Public CountA1, CountA11 As Variant
Public CountA2, CountA21 As Variant
Public CountA3, CountA31 As Variant
Public CountA4, CountA41 As Variant
Public CountB1, CountB11 As Variant
Public CountC1, CountC11 As Variant
Public CountD1, CountD11 As Variant
Public CountE1, CountE11 As Variant
Public CountF1, CountF11 As Variant
Public CountG1, CountG11 As Variant



Public sourceSheet As Worksheet
Dim msgValidateSheetGFA As String
Dim msgValidateSheetEFA As String

Sub AddRowsSchCG_STCGA7()
Dim vRows As Long
Sheets("CG").Activate
EfilingCommon.DefinedgridNameRange = "STCG.A7_PreviousYer_1||STCG.A7_DeductionSection_1||STCG.A7_AsstAcquiredYer_1||STCG.A7_AmtUtilised_1||STCG.A7_AmtNotUsed_1"
ActiveCellRange = EfilingCommon.searchLastRow("STCG.A7_PreviousYer_1")
vRows = EfilingCommon.insertRowUnderSectionWithFormula_CG

'prepopulating year and select for the added rows
Dim rangecells, rangecells1, rangecells2 As Range
    Dim mIntCells, q As Long
Set rangecells = Range("STCG.A7_PreviousYer_1").Cells
Set rangecells1 = Range("STCG.A7_DeductionSection_1").Cells
Set rangecells2 = Range("STCG.A7_Sl_1").Cells

mIntCells = Range("STCG.A7_PreviousYer_1").count

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Application.EnableEvents = False

For q = 1 To mIntCells


    If rangecells.item(q).value = "" Then rangecells.item(q).value = "(Select)"
    If rangecells1.item(q).value = "" Then rangecells1.item(q).value = "(Select)"
    If rangecells2.item(q).value = "" Then rangecells2.item(q).value = rangecells2.item(q).Offset(-1, 0).value + 1
Next

'Sheet13.Range("CG_STCGA7.Sl.").value = ((Sheet13.Range("CG_STCGA7.Sl.").Offset(-1, 0).value) + 1)

Application.EnableEvents = True
'----------------Lock Password-------------------START---
  Sheet13.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
End Sub
Sub AddRowsCG_Deduction54()
Dim vRows  As Long
Sheets("CG").Activate
'Newly changed as per DE 0.2 By Bindu
EfilingCommon.DefinedgridNameRange = "Deduction.54.DateOfTransfer||Deduction.54.CostOfHouse||Deduction.54.DateOfPursConst||Deduction.54.AmtdepCG||Deduction.54.DateofDeposit||Deduction.54.AcctNum||Deduction.54.IFSC||Deduction.54.AmtDedClaimed"
ActiveCellRange = EfilingCommon.searchLastRow("Deduction.54.DateOfTransfer")
vRows = insertRowUnderSectionWithFormula
End Sub

Sub AddRowsCG_Deduction54B()
Dim vRows  As Long
Sheets("CG").Activate
'Newly changed by Bindu as per DE 0.2
EfilingCommon.DefinedgridNameRange = "Deduction.54B.DateOfTransfer||Deduction.54B.CostOfAgriLand||Deduction.54B.DateOfPurLand||Deduction.54B.AmtdepCG||Deduction.54B.DateofDeposit||Deduction.54B.AcctNum||Deduction.54B.IFSCDeduction.54B.AmtDedClaimed"

ActiveCellRange = EfilingCommon.searchLastRow("Deduction.54B.DateOfTransfer")
vRows = insertRowUnderSectionWithFormula
End Sub

Sub AddRowsCG_Deduction54D()
Dim vRows  As Long
Sheets("CG").Activate
EfilingCommon.DefinedgridNameRange = "Deduction.54D.DateOfTransfer||Deduction.54D.CostOfPurcConst||Deduction.54B.DateOfPurLandB||Deduction.54D.AmtdepCG||Deduction.54D.AmtDedClaimed"
ActiveCellRange = EfilingCommon.searchLastRow("Deduction.54D.DateOfTransfer")
vRows = insertRowUnderSectionWithFormula
End Sub

Sub AddRowsCG_Deduction54EC()
Dim vRows  As Long
Sheets("CG").Activate
EfilingCommon.DefinedgridNameRange = "Deduction.54EC.DateOfTransfer||Deduction.54EC.AmountInvestd||Deduction.54EC.DateOfInvestment||Deduction.54EC.AmtDedClaimed"
ActiveCellRange = EfilingCommon.searchLastRow("Deduction.54EC.DateOfTransfer")
vRows = insertRowUnderSectionWithFormula
End Sub


Sub AddRowsCG_Deduction54F()
Dim vRows  As Long
Sheets("CG").Activate
EfilingCommon.DefinedgridNameRange = "Deduction.54F.DateOfTransfer||Deduction.54F.CostOfHouse||Deduction.54F.DateOfPurc||Deduction.54F.AmtdepCG||Deduction.54F.DateofDeposit||Deduction.54F.AcctNum||Deduction.54F.IFSC||Deduction.54F.AmtDedClaimed"
ActiveCellRange = EfilingCommon.searchLastRow("Deduction.54F.DateOfTransfer")
vRows = insertRowUnderSectionWithFormula
End Sub

Sub AddRowsCG_Deduction54G()
Dim vRows  As Long
Sheets("CG").Activate
EfilingCommon.DefinedgridNameRange = "Deduction.54G.DateOfTransfer||Deduction.54G.CostOfExpenses||Deduction.54G.DateOfPurc||Deduction.54G.AmtdepCG||Deduction.54G.DateofDeposit||Deduction.54G.AcctNum||Deduction.54G.IFSC||Deduction.54G.AmtDedClaimed" '||Deduction.54GB.DateOfPurchase||Deduction.54GB.AmountDeposited||Deduction.54GB.AmountClaimed"
ActiveCellRange = EfilingCommon.searchLastRow("Deduction.54G.DateOfTransfer")
vRows = insertRowUnderSectionWithFormula
End Sub

Sub AddRowsCG_Deduction54GA()
Dim vRows  As Long
Sheets("CG").Activate
EfilingCommon.DefinedgridNameRange = "Deduction.54GA.DateOfTransfer||Deduction.54GA.CostOfExpenses||Deduction.54GA.DateOfPurc||Deduction.54GA.AmtdepCG||Deduction.54GA.DateofDeposit||Deduction.54GA.AcctNum||Deduction.54GA.IFSC||Deduction.54GA.AmtDedClaimed" '||Deduction.54GB.DateOfPurchase||Deduction.54GB.AmountDeposited||Deduction.54GB.AmountClaimed"
ActiveCellRange = EfilingCommon.searchLastRow("Deduction.54GA.DateOfTransfer")
vRows = insertRowUnderSectionWithFormula
End Sub

Sub AddRowsCG_Deduction54115F()
Dim vRows  As Long
Sheets("CG").Activate
EfilingCommon.DefinedgridNameRange = "Deduction.115F.DateOfTransfer||Deduction.115F.AmtInvt||Deduction.115F.DateOfInvest||Deduction.115F.AmtDedClaimed"
ActiveCellRange = EfilingCommon.searchLastRow("Deduction.115F.DateOfTransfer")
vRows = insertRowUnderSectionWithFormula
End Sub
Sub setTblinfo_TR()

Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 
 ccount = 0
 mIntCells = sourceSheet.Range("TR_Country").count
 ccount = mIntCells
 GoTo endtR
 Set rangecells = sourceSheet.Range("TR_Country").Cells
 
 Dim tr_countrycd As Variant
 
 For mIntCtr = 1 To mIntCells
     
     tr_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(tr_countrycd) Then
        tr_countrycd = ""
        End If
     If Not tr_countrycd = "" Then
         ccount = ccount + 1
     End If
 Next

endtR:
 end_TR = ccount
 rngname_TR = "TR_Country;TR_TIN;TR_TaxPaidOutsideIndia;TR_TaxReliefOutsideIndia;TR_ReliefClaimedUsSection;"
 
End Sub

Sub setTblinfo_TR2()
 
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = sourceSheet.Range("TR_Country").count
 Set rangecells = sourceSheet.Range("TR_Country").Cells
 Dim tr_countrycd As Variant
 
 For mIntCtr = 1 To mIntCells
     
     tr_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(tr_countrycd) Then
        tr_countrycd = ""
        End If
     If Not tr_countrycd = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TR2 = ccount
 rngname_TR2 = "TR_Country;TR_TIN;TR_TaxPaidOutsideIndia;TR_TaxReliefOutsideIndia;TR_ReliefClaimedUsSection;"
End Sub
Sub ValidateTR_FAClick()
  ValidateSheetSchTR_FA
  VAlidateSheetFA
  
  fmsgboxStatus "Sheet TR-FA is OK"

End Sub

Sub ValidateSheetSchTR_FA()
'TR
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    
    If Not ValidatesheetTR Then
        sourceSheet.Activate
        fmsgboxsmall msgValidateSheetTR ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
End Sub

Sub VAlidateSheetFA()
'FA
    subProcCaption = "Validating FA - A"
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    
    subProcCaption = "Validating FA - A1"
    If Not ValidatesheetA1FA Then
        sourceSheet.Activate
        fmsgboxsmall "* Sch FA-A1: " & Chr(13) & Chr(13) & msgValidateSheetA1FA ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    
    subProcCaption = "Validating FA - A2"
    If Not ValidatesheetA2FA Then
        sourceSheet.Activate
        fmsgboxsmall "* Sch FA-A2: " & Chr(13) & Chr(13) & msgValidateSheetA2FA ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    
     subProcCaption = "Validating FA - A3"
    If Not ValidatesheetA3FA Then
        sourceSheet.Activate
        fmsgboxsmall "* Sch FA-A3: " & Chr(13) & Chr(13) & msgValidateSheetA3FA ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    
     subProcCaption = "Validating FA - A4"
    If Not ValidatesheetA4FA Then
        sourceSheet.Activate
        fmsgboxsmall "* Sch FA-A4: " & Chr(13) & Chr(13) & msgValidateSheetA4FA ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    
    
    subProcCaption = "Validating FA - B"
    If Not ValidatesheetBFA Then
        sourceSheet.Activate
        fmsgboxsmall "* Sch FA-B: " & Chr(13) & Chr(13) & msgValidateSheetBFA ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    subProcCaption = "Validating FA - C"
    If Not ValidatesheetCFA Then
        sourceSheet.Activate
        fmsgboxsmall "* Sch FA-C: " & Chr(13) & Chr(13) & msgValidateSheetCFA ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    subProcCaption = "Validating FA - D"
    If Not ValidatesheetDFA Then
        sourceSheet.Activate
        fmsgboxsmall "* Sch FA-D: " & Chr(13) & Chr(13) & msgValidateSheetDFA ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    subProcCaption = "Validating FA - E"
    If Not ValidatesheetEFA Then
        sourceSheet.Activate
        fmsgboxsmall "*Sch FA-E: " & Chr(13) & Chr(13) & msgValidateSheetEFA ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    subProcCaption = "Validating FA - F"
    If Not ValidatesheetFFA Then
        sourceSheet.Activate
        fmsgboxsmall "*Sch FA-F: " & Chr(13) & Chr(13) & msgValidateSheetFFA ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
   subProcCaption = "Validating FA - G"
   If Not ValidatesheetGFA Then
        sourceSheet.Activate
        fmsgboxsmall "*Sch FA-G: " & Chr(13) & Chr(13) & msgValidateSheetGFA ', vbOKOnly, "Error(s)!"
        CloseMsg
   End If
   
   If Sheet7.Range("AOIFlag_1").value = "Yes" Then
    setTblinfo_A1_FA
    setTblinfo_A2_FA
    setTblinfo_A3_FA
    setTblinfo_A4_FA
    setTblinfo_B_FA
    setTblinfo_C_FA
    setTblinfo_D_FA
    setTblinfo_E_FA
    setTblinfo_F_FA
    setTblinfo_G_FA
    If (end_A1_FA = 0 And end_A2_FA = 0 And end_A3_FA = 0 And end_A4_FA = 0 And end_B_FA = 0 And end_C_FA = 0 And end_D_FA = 0 And end_E_FA = 0 And end_F_FA = 0 And end_G_FA = 0) Then
        sourceSheet.Activate
        fmsgboxStatus "*Please Fill Schedule FA" ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
   End If
End Sub

Function ValidatesheetTR() As Boolean
 
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ValidatesheetTR = True
    If Not ValidateCountry_TR() Then ValidatesheetTR = False
    If Not isdropdownblank(sourceSheet.Range("TR_Country")(1).value) And (Len(sourceSheet.Range("TR_Country").item(1).value) > 0) Then
        If Not ValidateTIN_TR() Then ValidatesheetTR = False
        If Not ValidateIncome_TR() Then ValidatesheetTR = False
    End If
    If Not ValidateTotalTaxPaidOutsideIndia_TR() Then ValidatesheetTR = False
End Function

Function ValidateCountry_TR() As Boolean
   
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    Dim CountryName As String
    Dim i As Long
    Dim rangecells As Range
    
    ValidateCountry_TR = True
    setTblinfo_TR
    Set rangecells = sourceSheet.Range("TR_Country").Cells
    ReDim Country_TR(end_TR)
    ReDim CountryName_TR(end_TR)
    For i = 1 To end_TR
        CountryName = rangecells.item(i).value
'        Country_TR(I) = Mid(CountryName, WorksheetFunction.Search(":", CountryName) + 1, Len(CountryName))
'        CountryName_TR(I) = Mid(CountryName, 1, WorksheetFunction.Search(":", CountryName) - 1)
        If Len(Country_TR(i)) = 0 Then
            Exit For
        End If
        
        If Not chkCompulsory(Country_TR(i)) Then
            msgbox_TR ("* CountryName at Sr. No  " & i & "  in Sheet TR_FA  is Mandatory")
            ValidateCountry_TR = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(Country_TR(i)) Then
            msgbox_TR ("* CountryName at Sr. No  " & i & " in Sheet TR_FA  characters < > & ' " & Chr(34) & " are not allowed")
            ValidateCountry_TR = False
            Exit Function
        End If
    Next
End Function

Function ValidateTIN_TR() As Boolean
    
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    Dim rangecells As Range
    Dim i As Long
    
    ValidateTIN_TR = True
    setTblinfo_TR
    
    Set rangecells = sourceSheet.Range("TR_TIN").Cells
    ReDim TIN_TR(end_TR)
    For i = 1 To end_FSI
        'TIN_TR(i) = rangecells.Item(i).value
        If Not checkfieldspecialcharacter(TIN_TR(i)) Then
            msgbox_TR ("* TIN " & TIN_TR(i) & " in Sheet TR_FA  characters < > & ' " & Chr(34) & " are not allowed")
            ValidateTIN_TR = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncome_TR() As Boolean
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    Dim i As Long
    Dim rangecells As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Dim rangecellsAA As Range
    Dim trcount As Long
    
    ValidateIncome_TR = True
    setTblinfo_TR
    
    Set rangecellsAA = sourceSheet.Range("TR_TIN").Cells
    Set rangecells = sourceSheet.Range("TR_TaxPaidOutsideIndia").Cells
    Set rangecells2 = sourceSheet.Range("TR_TaxReliefOutsideIndia").Cells
    Set rangecells3 = sourceSheet.Range("TR_ReliefClaimedUsSection").Cells
    
    trcount = 0
    ReDim TIN_TR(end_TR)
    
    ReDim TaxPaidOutsideIndia_TR(end_TR)
    ReDim TaxReliefOutsideIndia_TR(end_TR)
    ReDim ReliefClaimedUsSection_TR(end_TR)
    
    For i = 1 To end_TR
        TIN_TR(i) = rangecellsAA.item(i).value
        TaxPaidOutsideIndia_TR(i) = rangecells.item(i).value
        TaxReliefOutsideIndia_TR(i) = rangecells2.item(i).value
        ReliefClaimedUsSection_TR(i) = rangecells3.item(i).value
        
        If (Len(TaxPaidOutsideIndia_TR(i)) > 0 Or Len(TaxReliefOutsideIndia_TR(i)) > 0) Then
            If (TaxPaidOutsideIndia_TR(i) > 0 Or TaxReliefOutsideIndia_TR(i) > 0) Then
                trcount = trcount + 1
                If Len(ReliefClaimedUsSection_TR(i)) = 0 Or ReliefClaimedUsSection_TR(i) = "(Select)" Then
                    If TaxReliefOutsideIndia_TR(i) > 0 Then
                        msgbox_TR ("* Relief claimed under Section at Sr. No  " & trcount & "  in Sheet TR_FA  is Mandatory")
                        ValidateIncome_TR = False
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
End Function

Function ValidateTaxPaid_TR() As Boolean
    
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    Dim rangecells As Range
    Dim i As Long
    
    ValidateTaxPaid_TR = True
    setTblinfo_TR
    
    Set rangecells = sourceSheet.Range("TR_TaxPaid").Cells
    ReDim TaxPaid_TR(end_TR)
    For i = 1 To end_TR
        TaxPaid_TR(i) = rangecells.item(i).value
        If Len(TaxPaid_TR(i)) = 0 Then
            'Exit For
        End If
        
        If Not chkCompulsory(TaxPaid_TR(i)) Then
            msgbox_TR ("* Tax Paid at Sr. No  " & i & "  in Sheet TR_FA  is Mandatory")
            ValidateTaxPaid_TR = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(TaxPaid_TR(i)) Then
            msgbox_TR ("* Tax Paid at Sr. No  " & i & " in Sheet TR_FA  characters < > & ' " & Chr(34) & " are not allowed")
            ValidateTaxPaid_TR = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotTaxRelief_TR() As Boolean
    Dim rangecells As Range
    Dim i As Long

    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ValidateTotTaxRelief_TR = True
    setTblinfo_TR
    
    Set rangecells = sourceSheet.Range("TR_TotTaxRelief").Cells
    ReDim TotTaxRelief_TR(end_TR)
    For i = 1 To end_TR
        TotTaxRelief_TR(i) = rangecells.item(i).value
        If Len(TotTaxRelief_TR(i)) = 0 Then
        'Exit For
        End If
        
        If Not chkCompulsory(TotTaxRelief_TR(i)) Then
            msgbox_TR ("* Relief claimed u/s 90/90A at Sr. No  " & i & "  in Sheet TR_FA  is Mandatory")
            ValidateTotTaxRelief_TR = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(TotTaxRelief_TR(i)) Then
            msgbox_TR ("* Relief claimed u/s 90/90A at Sr. No  " & i & " in Sheet TR_FA  characters < > & ' " & Chr(34) & " are not allowed")
            ValidateTotTaxRelief_TR = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotIncome_TR() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ValidateTotIncome_TR = True
    setTblinfo_TR
    
    Set rangecells = sourceSheet.Range("TR_TotIncome").Cells
    ReDim TotIncome_TR(end_TR)
    For i = 1 To end_TR
        TotIncome_TR(i) = rangecells.item(i).value
        If Len(TotIncome_TR(i)) = 0 Then
            Exit For
        End If
        
        If Not chkCompulsory(TotIncome_TR(i)) Then
            msgbox_TR ("* Total Income at Sr. No  " & i & "  in Sheet TR_FA  is Mandatory")
            ValidateTotIncome_TR = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(TotIncome_TR(i)) Then
            msgbox_TR ("* Total Income at Sr. No  " & i & " in Sheet TR_FA  characters < > & ' " & Chr(34) & " are not allowed")
            ValidateTotIncome_TR = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotTaxPaid_TR() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ValidateTotTaxPaid_TR = True
    setTblinfo_TR
    Set rangecells = sourceSheet.Range("TR_TotTaxPaid").Cells
    ReDim TotTaxPaid_TR(end_TR)
    For i = 1 To end_TR
        TotTaxPaid_TR(i) = rangecells.item(i).value
        If Len(TotTaxPaid_TR(i)) = 0 Then
            Exit For
        End If
        
        If Not chkCompulsory(TotTaxPaid_TR(i)) Then
            msgbox_TR ("* Total Tax Paid at Sr. No  " & i & "  in Sheet TR_FA  is Mandatory")
            ValidateTotTaxPaid_TR = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(TotTaxPaid_TR(i)) Then
            msgbox_TR ("* Total Tax Paid at Sr. No  " & i & " in Sheet TR_FA  characters < > & ' " & Chr(34) & " are not allowed")
            ValidateTotTaxPaid_TR = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotalTaxPaidOutsideIndia_TR() As Boolean
    
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ValidateTotalTaxPaidOutsideIndia_TR = True
    TotalTaxPaidOutsideIndia_TR = sourceSheet.Range("TR_TotalTaxPaidOutsideIndia").value
    TaxReliefOutsideIndiaDTAA_TR = sourceSheet.Range("TR_TaxReliefOutsideIndiaDTAA").value
    TaxReliefOutsideIndiaNotDTAA_TR = sourceSheet.Range("TR_TaxReliefOutsideIndiaNotDTAA").value
    TaxPaidOutsideIndFlg_TR = sourceSheet.Range("TR_TaxPaidOutsideIndFlg").value
    
    If isdropdownblank(TaxPaidOutsideIndFlg_TR) Then
        TaxPaidOutsideIndFlg_TR = ""
    End If
    
    setTblinfo_FSI
    
    If end_FSI > 0 Then
        If TaxPaidOutsideIndFlg_TR = "" Then
            msgbox_TR ("*select whether any tax paid outside India, on which tax relief was allowed in India, has been refunded/credited by the foreign tax authority during the Year in Schedule TR")
            ValidateTotalTaxPaidOutsideIndia_TR = False
            Exit Function
        End If
    End If
    
    AmtTaxRefunded_TR = sourceSheet.Range("TR_AmtTaxRefunded").value
    AssmtYrTaxRelief_TR = sourceSheet.Range("TR_AssmtYrTaxRelief").value
    
    Dim halfval, SecHval, firstval, secval As String
    firstval = ""
    secval = ""
    If AssmtYrTaxRelief_TR <> "" Then
        halfval = Mid(sourceSheet.Range("TR_AssmtYrTaxRelief").value, 1, 4)
        SecHval = Mid(sourceSheet.Range("TR_AssmtYrTaxRelief").value, 5, 1)
        firstval = Mid(sourceSheet.Range("TR_AssmtYrTaxRelief").value, 3, 2)
        secval = Mid(sourceSheet.Range("TR_AssmtYrTaxRelief").value, 6, 2)
        
        If SecHval <> "-" Then
            msgbox_TR ("* Enter Assessment Year in format YYYY-YY, eg 2023-24")
            ValidateTotalTaxPaidOutsideIndia_TR = False
            Exit Function
        End If
        
        If Not IsNumeric(halfval) Then
            msgbox_TR ("* Enter Assessment Year in format YYYY-YY, eg 2023-24")
            ValidateTotalTaxPaidOutsideIndia_TR = False
            Exit Function
        End If
        
        If Not IsNumeric(secval) Then
            msgbox_TR ("* Enter Assessment Year in format YYYY-YY, eg 2023-24")
            ValidateTotalTaxPaidOutsideIndia_TR = False
            Exit Function
        End If
        Dim firstval_1 As String
             firstval_1 = Mid(firstval + 1, 1, 2)
             If (Len(firstval_1) = 1) Then
             firstval_1 = "0" & firstval_1
             End If
       If secval <> IIf(Len(firstval + 1) > 2, Mid(firstval + 1, 2, 2), firstval_1) Then
            msgbox_TR ("* Assessment Year should be consecutive")
            ValidateTotalTaxPaidOutsideIndia_TR = False
            Exit Function
        End If
    End If
    
    If TaxPaidOutsideIndFlg_TR = "YES" Then
        If (AmtTaxRefunded_TR = "") Then  'Or AssmtYrTaxRelief_TR = "") Then
            msgbox_TR ("* Enter the amount of Tax refunded in Schedule TR")
            ValidateTotalTaxPaidOutsideIndia_TR = False
            Exit Function
        End If
        
        If (AssmtYrTaxRelief_TR = "") Then  'Or AssmtYrTaxRelief_TR = "") Then
            msgbox_TR ("* Enter assessment year in which tax relief allowed in Indian in Schedule TR")
            ValidateTotalTaxPaidOutsideIndia_TR = False
            Exit Function
        End If
    End If
End Function

'Function ValidateTaxPaidOutsideIndiaDTAA_TR() As Boolean
'    ValidateTaxPaidOutsideIndiaDTAA_TR = True
'    TaxPaidOutsideIndiaDTAA_TR = sourceSheet.Range("TR_TaxPaidOutsideIndiaDTAA").value
'End Function

'Function ValidateTaxPaidOutsideIndiaNotDTAA_TR() As Boolean
'    ValidateTaxPaidOutsideIndiaNotDTAA_TR = True
 '   TaxPaidOutsideIndiaNotDTAA_TR = sourceSheet.Range("TR_TaxPaidOutsideIndiaNotDTAA").value
 
'End Function

'Function ValidateTotTRelief_TR() As Boolean
'    ValidateTotTRelief_TR = True
 '   TotTRelief_TR = sourceSheet.Range("TR_TotTRelief").value
'End Function

Sub AddRows_A_FA()
    Dim numberofrows As Long
    
    setTblinfo_A_FA
    SelectLastRow ("FA_A_Country")
    numberofrows = InsertRowsAndFillFormulasBPA()
    Call ExendRangeNameToTable(numberofrows, rngname_A_FA)
End Sub

Sub AddRows_TR_FSI()
    Dim vRows As Long
    Sheet27.Activate
    EfilingCommon.DefinedgridNameRange = "TR_SrNo||TR_Country||TR_TIN||TR_TaxPaidOutsideIndia||TR_TaxReliefOutsideIndia||TR_ReliefClaimedUsSection"
    ActiveCellRange = EfilingCommon.searchLastRow("TR_Country")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula(1)
End Sub
Sub AddRows_A1_FA()
    Dim numberofrows As Long
    
    setTblinfo_A1_FA
    SelectLastRow ("FA_A1_Country")
    numberofrows = InsertRowsAndFillFormulasBPA()
    Call ExendRangeNameToTable(numberofrows, rngname_A1_FA)
End Sub

Sub AddRows_A2_FA()
    Dim numberofrows As Long
    
    setTblinfo_A2_FA
    SelectLastRow ("FA_A2_Country")
    numberofrows = InsertRowsAndFillFormulasBPA()
    Call ExendRangeNameToTable(numberofrows, rngname_A2_FA)
End Sub

Sub AddRows_A3_FA()
    Dim numberofrows As Long
    
    setTblinfo_A3_FA
    SelectLastRow ("FA_A3_Country")
    numberofrows = InsertRowsAndFillFormulasBPA()
    Call ExendRangeNameToTable(numberofrows, rngname_A3_FA)
End Sub

Sub AddRows_A4_FA()
    Dim numberofrows As Long
    
    setTblinfo_A4_FA
    SelectLastRow ("FA_A4_Country")
    numberofrows = InsertRowsAndFillFormulasBPA()
    Call ExendRangeNameToTable(numberofrows, rngname_A4_FA)
End Sub

Sub AddRows_B_FA()
    Dim numberofrows As Long
    
    setTblinfo_B_FA
    SelectLastRow ("FA_B_Country")
    numberofrows = InsertRowsAndFillFormulasBPA()
    Call ExendRangeNameToTable(numberofrows, rngname_B_FA)
End Sub

Sub AddRows_C_FA()
    Dim numberofrows As Long
    
    setTblinfo_C_FA
    SelectLastRow ("FA_C_Country")
    numberofrows = InsertRowsAndFillFormulasBPA()
    Call ExendRangeNameToTable(numberofrows, rngname_C_FA)
End Sub

Sub AddRows_D_FA()
    Dim numberofrows As Long
    
    setTblinfo_D_FA
    SelectLastRow ("FA_D_Country")
    numberofrows = InsertRowsAndFillFormulasBPA()
    Call ExendRangeNameToTable(numberofrows, rngname_D_FA)
End Sub

Sub AddRows_F_FA()
    Dim numberofrows As Long
    
    setTblinfo_F_FA
    SelectLastRow ("FA_F_Country")
    numberofrows = InsertRowsAndFillFormulasBPA()
    Call ExendRangeNameToTable(numberofrows, rngname_F_FA)
End Sub

Sub AddRows_G_FA()
    Dim numberofrows As Long
    
    setTblinfo_G_FA
    SelectLastRow ("FA_G_Country")
    numberofrows = InsertRowsAndFillFormulasBPA()
    Call ExendRangeNameToTable(numberofrows, rngname_G_FA)
End Sub

Sub AddRows_E_FA()
    Dim numberofrows As Long
    
    setTblinfo_E_FA
    SelectLastRow ("FA_E_InstName")
    numberofrows = InsertRowsAndFillFormulasBPA()
    Call ExendRangeNameToTable(numberofrows, rngname_E_FA)
End Sub



Sub setTblinfo_A1_FA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim fa_a_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_A1_Country").count
    Set rangecells = sourceSheet.Range("FA_A1_Country").Cells
    For mIntCtr = 1 To mIntCells
        fa_a_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(fa_a_countrycd) Then
            fa_a_countrycd = ""
        End If
        
        If Not fa_a_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_A1_FA = ccount
    rngname_A1_FA = "FA_A1_Country;FA_A1_BankName;FA_A1_BankAdd;FA_A1_ZipCode;FA_A1_ForeignAccountNumber;FA_A1_StatusBeneficiary;FA_A1_AccOpeningDate;FA_A1_PeakBal;FA_A1_ClosingBal;FA_A1_Grossinterest;"
End Sub

Sub setTblinfo_A11_FA()
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim ccount1 As Long
    Dim fa_b_countrycd As Variant
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim fa_a_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_A1_BankName").count
    Set rangecells = sourceSheet.Range("FA_A1_BankName").Cells
    For mIntCtr = 1 To mIntCells
        fa_a_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(fa_a_countrycd) Then
            fa_a_countrycd = ""
        End If
        
        If Not fa_a_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_A11_FA = ccount
    rngname_A1_FA = "FA_A1_Country;FA_A1_BankName;FA_A1_BankAdd;FA_A1_ZipCode;FA_A1_ForeignAccountNumber;FA_A1_StatusBeneficiary;FA_A1_AccOpeningDate;FA_A1_PeakBal;FA_A1_ClosingBal;FA_A1_Grossinterest;"
Dim rn As Variant
    Dim i As Long
    rn = Split(rngname_A1_FA, ";")
    For i = 0 To UBound(rn) - 1
        Set rangecells1 = sourceSheet.Range(rn(i)).Cells
        ccount1 = 0
        For mIntCtr = 1 To mIntCells
            fa_b_countrycd = rangecells1.item(mIntCtr).value
            If isdropdownblank(fa_b_countrycd) Then
                fa_b_countrycd = ""
            End If
            
            If Not fa_b_countrycd = "" Then
                ccount1 = ccount1 + 1
            End If
        Next
        If ccount1 > ccount Then Exit For
    Next
    
    CountA1 = ccount
    CountA11 = ccount1

End Sub

Sub setTblinfo_A2_FA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim fa_a_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_A2_Country").count
    Set rangecells = sourceSheet.Range("FA_A2_Country").Cells
    For mIntCtr = 1 To mIntCells
        fa_a_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(fa_a_countrycd) Then
            fa_a_countrycd = ""
        End If
        
        If Not fa_a_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_A2_FA = ccount
    rngname_A2_FA = "FA_A2_Country;FA_A2_BankName;FA_A2_BankAdd;FA_A2_ZipCode;FA_A2_ForeignAccountNumber;FA_A2_StatusBeneficiary;FA_A2_AccOpeningDate;FA_A2_PeakBal;FA_A2_ClosingBal;FA_A2_Grossinterest;FA_A2_Grossinterest_Nature;"
End Sub

Sub setTblinfo_A21_FA()
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim ccount1 As Long
    Dim fa_b_countrycd As Variant
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim fa_a_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_A2_BankName").count
    Set rangecells = sourceSheet.Range("FA_A2_BankName").Cells
    For mIntCtr = 1 To mIntCells
        fa_a_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(fa_a_countrycd) Then
            fa_a_countrycd = ""
        End If
        
        If Not fa_a_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_A21_FA = ccount
    rngname_A2_FA = "FA_A2_Country;FA_A2_BankName;FA_A2_BankAdd;FA_A2_ZipCode;FA_A2_ForeignAccountNumber;FA_A2_StatusBeneficiary;FA_A2_AccOpeningDate;FA_A2_PeakBal;FA_A2_ClosingBal;FA_A2_Grossinterest;FA_A2_Grossinterest_Nature;"
    Dim rn As Variant
    Dim i As Long
    rn = Split(rngname_A2_FA, ";")
    For i = 0 To UBound(rn) - 1
        Set rangecells1 = sourceSheet.Range(rn(i)).Cells
        ccount1 = 0
        For mIntCtr = 1 To mIntCells
            fa_b_countrycd = rangecells1.item(mIntCtr).value
            If isdropdownblank(fa_b_countrycd) Then
                fa_b_countrycd = ""
            End If
            
            If Not fa_b_countrycd = "" Then
                ccount1 = ccount1 + 1
            End If
        Next
        If ccount1 > ccount Then Exit For
    Next
    
    CountA2 = ccount
    CountA21 = ccount1

End Sub

Sub setTblinfo_A3_FA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim fa_a_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_A3_Country").count
    Set rangecells = sourceSheet.Range("FA_A3_Country").Cells
    For mIntCtr = 1 To mIntCells
        fa_a_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(fa_a_countrycd) Then
            fa_a_countrycd = ""
        End If
        
        If Not fa_a_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_A3_FA = ccount
    rngname_A3_FA = "FA_A3_Country;FA_A3_BankName;FA_A3_BankAdd;FA_A3_ZipCode;FA_A3_NatureOfEntity;FA_A3_AccOpeningDate;FA_A3_initialvalue;FA_A3_PeakBal;FA_A3_ClosingBal;FA_A3_Totalgrossamount;FA_A3_Totalgrosproceeds;"
End Sub

Sub setTblinfo_A31_FA()
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim ccount1 As Long
    Dim fa_b_countrycd As Variant
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim fa_a_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_A3_BankName").count
    Set rangecells = sourceSheet.Range("FA_A3_BankName").Cells
    For mIntCtr = 1 To mIntCells
        fa_a_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(fa_a_countrycd) Then
            fa_a_countrycd = ""
        End If
        
        If Not fa_a_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_A31_FA = ccount
    rngname_A3_FA = "FA_A3_Country;FA_A3_BankName;FA_A3_BankAdd;FA_A3_ZipCode;FA_A3_NatureOfEntity;FA_A3_AccOpeningDate;FA_A3_initialvalue;FA_A3_PeakBal;FA_A3_ClosingBal;FA_A3_Totalgrossamount;FA_A3_Totalgrosproceeds;"


    
    Dim rn As Variant
    Dim i As Long
    rn = Split(rngname_A3_FA, ";")
    For i = 0 To UBound(rn) - 1
        Set rangecells1 = sourceSheet.Range(rn(i)).Cells
        ccount1 = 0
        For mIntCtr = 1 To mIntCells
            fa_b_countrycd = rangecells1.item(mIntCtr).value
            If isdropdownblank(fa_b_countrycd) Then
                fa_b_countrycd = ""
            End If
            
            If Not fa_b_countrycd = "" Then
                ccount1 = ccount1 + 1
            End If
        Next
        If ccount1 > ccount Then Exit For
    Next
    
    CountA3 = ccount
    CountA31 = ccount1
End Sub

Sub setTblinfo_A4_FA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim fa_a_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_A4_Country").count
    Set rangecells = sourceSheet.Range("FA_A4_Country").Cells
    For mIntCtr = 1 To mIntCells
        fa_a_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(fa_a_countrycd) Then
            fa_a_countrycd = ""
        End If
        
        If Not fa_a_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_A4_FA = ccount
    rngname_A4_FA = "FA_A4_Country;FA_A4_BankName;FA_A4_BankAdd;FA_A4_ZipCode;FA_A4_AccOpeningDate;FA_A4_the_cashvalue;FA_A4_Total_grossamount;"
End Sub

Sub setTblinfo_A41_FA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim fa_a_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_A4_BankName").count
    Set rangecells = sourceSheet.Range("FA_A4_BankName").Cells
    For mIntCtr = 1 To mIntCells
        fa_a_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(fa_a_countrycd) Then
            fa_a_countrycd = ""
        End If
        
        If Not fa_a_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_A41_FA = ccount
    rngname_A4_FA = "FA_A4_Country;FA_A4_BankName;FA_A4_BankAdd;FA_A4_ZipCode;FA_A4_AccOpeningDate;FA_A4_the_cashvalue;FA_A4_Total_grossamount;"

Dim rangecells1 As Range
    Dim ccount1 As Long
    Dim fa_b_countrycd As Variant
    
    Dim rn As Variant
    Dim i As Long
    rn = Split(rngname_A4_FA, ";")
    For i = 0 To UBound(rn) - 1
        Set rangecells1 = sourceSheet.Range(rn(i)).Cells
        ccount1 = 0
        For mIntCtr = 1 To mIntCells
            fa_b_countrycd = rangecells1.item(mIntCtr).value
            If isdropdownblank(fa_b_countrycd) Then
                fa_b_countrycd = ""
            End If
            
            If Not fa_b_countrycd = "" Then
                ccount1 = ccount1 + 1
            End If
        Next
        If ccount1 > ccount Then Exit For
    Next
    
    CountA4 = ccount
    CountA41 = ccount1
End Sub

Sub setTblinfo_A_FA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim fa_a_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_A_Country").count
    Set rangecells = sourceSheet.Range("FA_A_Country").Cells
    For mIntCtr = 1 To mIntCells
        fa_a_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(fa_a_countrycd) Then
            fa_a_countrycd = ""
        End If
        
        If Not fa_a_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_A_FA = ccount
    rngname_A_FA = "FA_A_Country;FA_A_ZipCode;FA_A_BankName;FA_A_BankAdd;FA_A_AccountName;FA_A_StatusBeneficiary;FA_A_ForeignAccountNumber;FA_A_IBAN;FA_A_AccOpeningDate;FA_A_PeakBal;FA_A_AccountInterest;FA_A_Amount;FA_A_ScheduleOffered;FA_A_ItemNumOfSchedule;"
End Sub

Sub setTblinfo_A_FA2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim fa_a_countrycd As Variant
    
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_A_Country").count
    Set rangecells = sourceSheet.Range("FA_A_Country").Cells
    For mIntCtr = 1 To mIntCells
        fa_a_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(fa_a_countrycd) Then
            fa_a_countrycd = ""
        End If
        
        If Not fa_a_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_A_FA2 = ccount
    rngname_A_FA2 = "FA_A_Country;FA_A_ZipCode;FA_A_BankName;FA_A_BankAdd;FA_A_AccountName;FA_A_StatusBeneficiary;FA_A_ForeignAccountNumber;FA_A_IBAN;FA_A_AccOpeningDate;FA_A_PeakBal;FA_A_AccountInterest;FA_A_Amount;FA_A_ScheduleOffered;FA_A_ItemNumOfSchedule;"
End Sub

Sub setTblinfo_B_FA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim fa_b_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_B_Country").count
    Set rangecells = sourceSheet.Range("FA_B_Country").Cells
    
    For mIntCtr = 1 To mIntCells
        fa_b_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(fa_b_countrycd) Then
            fa_b_countrycd = ""
        End If
        
        If Not fa_b_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_B_FA = ccount
    rngname_B_FA = "FA_B_Country;FA_B_ZipCode;FA_B_EntityNature;FA_B_EntityName;FA_B_EntityAdd;FA_B_NatureOfInterest;FA_B_DateSinceHeld;FA_B_TotalInv;FA_B_IncomeInterest;FA_B_NatureOfIncome;FA_B_Amount;FA_B_ScheduleOffered;FA_B_ItemNumOfSchedule;"
End Sub

Sub setTblinfo_B_FA2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim fa_b_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_B_Country").count
    Set rangecells = sourceSheet.Range("FA_B_Country").Cells
    For mIntCtr = 1 To mIntCells
        fa_b_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(fa_b_countrycd) Then
            fa_b_countrycd = ""
        End If
        
        If Not fa_b_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_B_FA2 = ccount
    rngname_B_FA2 = "FA_B_Country;FA_B_ZipCode;FA_B_EntityNature;FA_B_EntityName;FA_B_EntityAdd;FA_B_NatureOfInterest;FA_B_DateSinceHeld;FA_B_TotalInv;FA_B_IncomeInterest;FA_B_NatureOfIncome;FA_B_Amount;FA_B_ScheduleOffered;FA_B_ItemNumOfSchedule;"

Dim rangecells1 As Range
    Dim ccount1 As Long
    'Dim fa_b_countrycd As Variant
    
    Dim rn As Variant
    Dim i As Long
    rn = Split(rngname_B_FA2, ";")
    For i = 0 To UBound(rn) - 1
        Set rangecells1 = sourceSheet.Range(rn(i)).Cells
        ccount1 = 0
        For mIntCtr = 1 To mIntCells
            fa_b_countrycd = rangecells1.item(mIntCtr).value
            If isdropdownblank(fa_b_countrycd) Then
                fa_b_countrycd = ""
            End If
            
            If Not fa_b_countrycd = "" Then
                ccount1 = ccount1 + 1
            End If
        Next
        If ccount1 > ccount Then Exit For
    Next
    
    CountB1 = ccount
    CountB11 = ccount1
End Sub

Sub setTblinfo_C_FA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim fa_c_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_C_Country").count
    
    Set rangecells = sourceSheet.Range("FA_C_Country").Cells
    For mIntCtr = 1 To mIntCells
        fa_c_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(fa_c_countrycd) Then
            fa_c_countrycd = ""
        End If
        
        If Not fa_c_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_C_FA = ccount
    rngname_C_FA = "FA_C_Country;FA_C_ZipCode;FA_C_PropertyAdd;FA_C_OwnershipBeneficiary;FA_C_DateOfAcquisition;FA_C_TotalInv;FA_C_IncomeFromProperty;FA_C_NatureOfIncome;FA_C_Amount;FA_C_ScheduleOffered;FA_C_ItemNumOfSchedule;"
End Sub

Sub setTblinfo_C_FA2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim fa_c_countrycd As Variant
   
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_C_Country").count
    Set rangecells = sourceSheet.Range("FA_C_Country").Cells
    
    Set rangecells = sourceSheet.Range("FA_C_Country").Cells
    For mIntCtr = 1 To mIntCells
        fa_c_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(fa_c_countrycd) Then
            fa_c_countrycd = ""
        End If
        
        If Not fa_c_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_C_FA2 = ccount
    rngname_C_FA2 = "FA_C_Country;FA_C_ZipCode;FA_C_PropertyAdd;FA_C_OwnershipBeneficiary;FA_C_DateOfAcquisition;FA_C_TotalInv;FA_C_IncomeFromProperty;FA_C_NatureOfIncome;FA_C_Amount;FA_C_ScheduleOffered;FA_C_ItemNumOfSchedule;"

Dim rangecells1 As Range
    Dim ccount1 As Long
    Dim fa_b_countrycd As Variant
    
    Dim rn As Variant
    Dim i As Long
    rn = Split(rngname_C_FA2, ";")
    For i = 0 To UBound(rn) - 1
        Set rangecells1 = sourceSheet.Range(rn(i)).Cells
        ccount1 = 0
        For mIntCtr = 1 To mIntCells
            fa_b_countrycd = rangecells1.item(mIntCtr).value
            If isdropdownblank(fa_b_countrycd) Then
                fa_b_countrycd = ""
            End If
            
            If Not fa_b_countrycd = "" Then
                ccount1 = ccount1 + 1
            End If
        Next
        If ccount1 > ccount Then Exit For
    Next
    
    CountC1 = ccount
    CountC11 = ccount1
End Sub

Sub setTblinfo_D_FA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim fa_d_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_D_Country").count
    
    Set rangecells = sourceSheet.Range("FA_D_Country").Cells
    
    For mIntCtr = 1 To mIntCells
        fa_d_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(fa_d_countrycd) Then
            fa_d_countrycd = ""
        End If
        
        If Not fa_d_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_D_FA = ccount
    rngname_D_FA = "FA_D_Country;FA_D_ZipCode;FA_D_AssetNature;FA_D_OwnershipBeneficiary;FA_D_DateOfAcquisition;FA_D_TotalInv;FA_D_IncomeFromProperty;FA_D_NatureOfIncome;FA_D_Amount;FA_D_ScheduleOffered;FA_D_ItemNumOfSchedule;"
End Sub

Sub setTblinfo_D_FA2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim fa_d_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_D_Country").count
    Set rangecells = sourceSheet.Range("FA_D_Country").Cells
        
    For mIntCtr = 1 To mIntCells
        fa_d_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(fa_d_countrycd) Then
            fa_d_countrycd = ""
        End If
        
        If Not fa_d_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_D_FA2 = ccount
    rngname_D_FA2 = "FA_D_Country;FA_D_ZipCode;FA_D_AssetNature;FA_D_OwnershipBeneficiary;FA_D_DateOfAcquisition;FA_D_TotalInv;FA_D_IncomeFromProperty;FA_D_NatureOfIncome;FA_D_Amount;FA_D_ScheduleOffered;FA_D_ItemNumOfSchedule;"

Dim rangecells1 As Range
    Dim ccount1 As Long
    Dim fa_b_countrycd As Variant
    
    Dim rn As Variant
    Dim i As Long
    rn = Split(rngname_D_FA2, ";")
    For i = 0 To UBound(rn) - 1
        Set rangecells1 = sourceSheet.Range(rn(i)).Cells
        ccount1 = 0
        For mIntCtr = 1 To mIntCells
            fa_b_countrycd = rangecells1.item(mIntCtr).value
            If isdropdownblank(fa_b_countrycd) Then
                fa_b_countrycd = ""
            End If
            
            If Not fa_b_countrycd = "" Then
                ccount1 = ccount1 + 1
            End If
        Next
        If ccount1 > ccount Then Exit For
    Next
    
    CountD1 = ccount
    CountD11 = ccount1

End Sub

Sub setTblinfo_F_FA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim FA_F_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_F_Country").count
    
    Set rangecells = sourceSheet.Range("FA_F_Country").Cells
    
    For mIntCtr = 1 To mIntCells
        FA_F_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(FA_F_countrycd) Then
            FA_F_countrycd = ""
        End If
        
        If Not FA_F_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_F_FA = ccount
    rngname_F_FA = "FA_F_Country;FA_F_ZipCode;FA_F_NameOfTrust;FA_F_AddressOfTrust;FA_F_NameOfOtherTrustees;FA_F_AddressOfOtherTrustees;FA_F_NameOfSettlor;FA_F_AddressOfSettlor;FA_F_NameOfBeneficiaries;FA_F_AddressOfBeneficiaries;FA_F_DateSinceHeld;FA_F_IncomeDerivedIsTaxable;FA_F_IncomeDerivedAmount;FA_F_Amount;FA_F_Scheduleoffered;FA_F_ItemNumberSchedule;"
End Sub

Sub setTblinfo_F_FA2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim FA_F_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_F_Country").count
    Set rangecells = sourceSheet.Range("FA_F_Country").Cells
        
    For mIntCtr = 1 To mIntCells
        FA_F_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(FA_F_countrycd) Then
            FA_F_countrycd = ""
        End If
        
        If Not FA_F_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_F_FA2 = ccount
    rngname_F_FA2 = "FA_F_IncomeDerivedAmount;FA_F_Amount;FA_F_Scheduleoffered;FA_F_ItemNumberSchedule;FA_F_Country;FA_F_ZipCode;FA_F_NameOfTrust;FA_F_AddressOfTrust;FA_F_NameOfOtherTrustees;FA_F_AddressOfOtherTrustees;FA_F_NameOfSettlor;FA_F_AddressOfSettlor;FA_F_NameOfBeneficiaries;FA_F_AddressOfBeneficiaries;FA_F_DateSinceHeld;FA_F_IncomeDerivedIsTaxable;"

   Dim rangecells1 As Range
    Dim ccount1 As Long
    Dim fa_b_countrycd As Variant
    
    Dim rn As Variant
    Dim i As Long
    rn = Split(rngname_F_FA2, ";")
    For i = 0 To UBound(rn) - 1
        Set rangecells1 = sourceSheet.Range(rn(i)).Cells
        ccount1 = 0
        For mIntCtr = 1 To mIntCells
            fa_b_countrycd = rangecells1.item(mIntCtr).value
            If isdropdownblank(fa_b_countrycd) Then
                fa_b_countrycd = ""
            End If
            
            If Not fa_b_countrycd = "" Then
                ccount1 = ccount1 + 1
            End If
        Next
        If ccount1 > ccount Then Exit For
    Next
    
    CountF1 = ccount
    CountF11 = ccount1

   
End Sub

Sub setTblinfo_E_FA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim FA_E_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_E_InstName").count
    Set rangecells = sourceSheet.Range("FA_E_InstName").Cells
    For mIntCtr = 1 To mIntCells
    FA_E_countrycd = rangecells.item(mIntCtr).value
    
    If isdropdownblank(FA_E_countrycd) Then
            FA_E_countrycd = ""
        End If
        
        If Not FA_E_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_E_FA = ccount
    rngname_E_FA = "FA_E_InstName;FA_E_InstAdd;FA_E_Country;FA_E_ZipCode;FA_E_AccountName;FA_E_InstitutionAccountNumber;FA_E_PeakBalInv;FA_E_IncomeAccured;FA_E_IncomeAccuredAmount;FA_E_Amount;FA_E_Scheduleoffered;FA_E_ItemNumberSchedule;"
End Sub

Sub setTblinfo_E_FA2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim FA_E_countrycd As Variant
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_E_InstName").count
    Set rangecells = sourceSheet.Range("FA_E_InstName").Cells
    For mIntCtr = 1 To mIntCells
        FA_E_countrycd = rangecells.item(mIntCtr).value
    
    If isdropdownblank(FA_E_countrycd) Then
            FA_E_countrycd = ""
        End If
        
        If Not FA_E_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_E_FA2 = ccount
    rngname_E_FA2 = "FA_E_IncomeAccuredAmount;FA_E_Amount;FA_E_Scheduleoffered;FA_E_ItemNumberSchedule;FA_E_InstName;FA_E_InstAdd;FA_E_Country;FA_E_ZipCode;FA_E_AccountName;FA_E_InstitutionAccountNumber;FA_E_PeakBalInv;FA_E_IncomeAccured;"

Dim rangecells1 As Range
    Dim ccount1 As Long
    Dim fa_b_countrycd As Variant
    
    Dim rn As Variant
    Dim i As Long
    rn = Split(rngname_E_FA2, ";")
    For i = 0 To UBound(rn) - 1
        Set rangecells1 = sourceSheet.Range(rn(i)).Cells
        ccount1 = 0
        For mIntCtr = 1 To mIntCells
            fa_b_countrycd = rangecells1.item(mIntCtr).value
            If isdropdownblank(fa_b_countrycd) Then
                fa_b_countrycd = ""
            End If
            
            If Not fa_b_countrycd = "" Then
                ccount1 = ccount1 + 1
            End If
        Next
        If ccount1 > ccount Then Exit For
    Next
    
    CountE1 = ccount
    CountE11 = ccount1


End Sub



Sub setTblinfo_G_FA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim fa_g_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_G_Country").count
    Set rangecells = sourceSheet.Range("FA_G_Country").Cells
    For mIntCtr = 1 To mIntCells
       fa_g_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(fa_g_countrycd) Then
            fa_g_countrycd = ""
        End If
        
        If Not fa_g_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_G_FA = ccount
    rngname_G_FA = "FA_G_Country;FA_G_ZipCode;FA_G_NameOfPerson;FA_G_AddressOfPerson;FA_G_IncomeDerived;FA_G_NatureOfIncome;FA_G_WhetherTaxable;FA_G_Amount;FA_G_Scheduleoffered;FA_G_ItemNumberSchedule;"
End Sub


Sub setTblinfo_G_FA2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim fa_g_countrycd As Variant
     
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    ccount = 0
    mIntCells = sourceSheet.Range("FA_G_Country").count
    Set rangecells = sourceSheet.Range("FA_G_Country").Cells
    For mIntCtr = 1 To mIntCells
        fa_g_countrycd = rangecells.item(mIntCtr).value
        If isdropdownblank(fa_g_countrycd) Then
            fa_g_countrycd = ""
        End If
        
        If Not fa_g_countrycd = "" Then
            ccount = ccount + 1
        End If
    Next
    end_G_FA2 = ccount
    rngname_G_FA2 = "FA_G_Amount;FA_G_Scheduleoffered;FA_G_ItemNumberSchedule;FA_G_Country;FA_G_ZipCode;FA_G_NameOfPerson;FA_G_AddressOfPerson;FA_G_IncomeDerived;FA_G_NatureOfIncome;FA_G_WhetherTaxable;"

   Dim rangecells1 As Range
    Dim ccount1 As Long
    Dim fa_b_countrycd As Variant
    
    Dim rn As Variant
    Dim i As Long
    rn = Split(rngname_G_FA2, ";")
    For i = 0 To UBound(rn) - 1
        Set rangecells1 = sourceSheet.Range(rn(i)).Cells
        ccount1 = 0
        For mIntCtr = 1 To mIntCells
            fa_b_countrycd = rangecells1.item(mIntCtr).value
            If isdropdownblank(fa_b_countrycd) Then
                fa_b_countrycd = ""
            End If
            
            If Not fa_b_countrycd = "" Then
                ccount1 = ccount1 + 1
            End If
        Next
        If ccount1 > ccount Then Exit For
    Next
    
    CountG1 = ccount
    CountG11 = ccount1

End Sub

Function ValidatesheetA1FA() As Boolean
 
    ValidatesheetA1FA = True
    If Not ValidateCountry_A1_FA() Then ValidatesheetA1FA = False
    If Not isdropdownblank(sourceSheet.Range("FA_A1_Country").item(1).value) And (Len(sourceSheet.Range("FA_A1_Country").item(1).value) > 0) Then
    If Not ValidateBankName_A1_FA() Then ValidatesheetA1FA = False
    If Not ValidateBankAddress_A1_FA() Then ValidatesheetA1FA = False
    If Not ValidateZipCode_A1_FA() Then ValidatesheetA1FA = False
    If Not ValidateAccountName_A1_FA() Then ValidatesheetA1FA = False
    If Not ValidateStatus_A1_FA() Then ValidatesheetA1FA = False
    If Not ValidateAccountOpeningDate_A1() Then ValidatesheetA1FA = False
    If Not ValidatePeakBal_A1_FA() Then ValidatesheetA1FA = False
    If Not ValidateCloseBalance_A1_FA() Then ValidatesheetA1FA = False
    If Not ValidateGrossInterest_A1_FA() Then ValidatesheetA1FA = False
    End If
    
    setTblinfo_A11_FA
    If (end_A1_FA <> end_A11_FA) Then
'        msgbox_A1_FA ("* Details of Country in Table A1 of Schedule FA are Mandatory")

'SIT-69523 Changed by sai

        msgbox_A1_FA ("*Please select the Country Name and Code in Table A1 of Schedule FA.")
        ValidatesheetA1FA = False
        Exit Function
    End If
    
    If (CountA1 <> CountA11) Then
        msgbox_A1_FA ("* Enter all mandatory field in Table A1 of Schedule FA")
        ValidatesheetA1FA = False
        Exit Function
    End If
    
    
End Function


Function ValidateCountry_A1_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim CountryName As String
    
    ValidateCountry_A1_FA = True
    setTblinfo_A1_FA
    
    Set rangecells = sourceSheet.Range("FA_A1_Country").Cells
    ReDim CountryCodeA1_FA(end_A1_FA)
    ReDim CountryNameA1_FA(end_A1_FA)
    For i = 1 To end_A1_FA
    noOfProcessSub = end_A1_FA
        CountryName = rangecells.item(i).value
        
        If isdropdownblank(CountryName) Then
            msgbox_A1_FA ("* Please select the Country Name and Code in Table A1  at Sr. No  " & i & " in Table A1 of Schedule FA")
            ValidateCountry_A1_FA = False
            Exit Function
        End If
        
        If CountryName <> "" Then
        CountryNameA1_FA(i) = Mid(CountryName, WorksheetFunction.Search("-", CountryName) + 1, Len(CountryName))
         CountryCodeA1_FA(i) = Mid(CountryName, 1, WorksheetFunction.Search("-", CountryName) - 1)
         End If
        
        If Len(CountryCodeA1_FA(i)) = 0 Then
            Exit For
        End If
        
        If Not chkCompulsory(CountryCodeA1_FA(i)) Then
            msgbox_A1_FA ("* Please select the Country Name and Code in Table A1 at Sr. No  " & i & "  in Table A1 of Schedule FA")
            ValidateCountry_A1_FA = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(CountryCodeA1_FA(i)) Then
            msgbox_A1_FA ("* Country Name at Sr. No  " & i & "characters < > & ' " & Chr(34) & " are not allowed in Table A1 of Schedule FA")
            ValidateCountry_A1_FA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function


Function ValidateBankName_A1_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
  
    ValidateBankName_A1_FA = True
    setTblinfo_A1_FA
    
    Set rangecells = sourceSheet.Range("FA_A1_BankName").Cells
    ReDim BankNameA1_FA(end_A1_FA)
    For i = 1 To end_A1_FA
        BankNameA1_FA(i) = rangecells.item(i).value
        If Len(BankNameA1_FA(i)) = 0 Then
            'Exit For
        End If
        
        If Not chkCompulsory(BankNameA1_FA(i)) Then
            msgbox_A1_FA ("* Please enter Name of  financial institution at Sr. No  " & i & " in Table A1 of Schedule FA")
            ValidateBankName_A1_FA = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(BankNameA1_FA(i)) Then
            msgbox_A1_FA ("* Name of  financial institution at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A1 of Schedule FA")
            ValidateBankName_A1_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateBankAddress_A1_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateBankAddress_A1_FA = True
    setTblinfo_A1_FA
    
    Set rangecells = sourceSheet.Range("FA_A1_BankAdd").Cells
    ReDim BankAddA1_FA(end_A1_FA)
    For i = 1 To end_A1_FA
        BankAddA1_FA(i) = rangecells.item(i).value
        If Len(BankAddA1_FA(i)) = 0 Then
            'Exit For
        End If
        
        If Not chkCompulsory(BankAddA1_FA(i)) Then
            msgbox_A1_FA ("* Please enter the Address of financial institution at Sr. No  " & i & " in Table A1 of Schedule FA")
            ValidateBankAddress_A1_FA = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(BankAddA1_FA(i)) Then
            msgbox_A1_FA ("* Address of  financial institution at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Table A1 of Schedule FA")
            ValidateBankAddress_A1_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateZipCode_A1_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateZipCode_A1_FA = True
    setTblinfo_A1_FA
    
    Set rangecells = sourceSheet.Range("FA_A1_ZipCode").Cells
    ReDim ZipCodeA1_FA(end_A1_FA)
    For i = 1 To end_A1_FA
        ZipCodeA1_FA(i) = rangecells.item(i).value
        If Len(ZipCodeA1_FA(i)) = 0 Then
            'Exit For
        End If
        
        If Not chkCompulsory(ZipCodeA1_FA(i)) Then
            msgbox_A1_FA ("* Please enter the ZipCode at Sr. No  " & i & " If  ZIP code is not available, then enter XXXXXX in Table A1 of Schedule FA")
            ValidateZipCode_A1_FA = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(ZipCodeA1_FA(i)) Then
            msgbox_A1_FA ("* ZipCode at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A1 of Schedule FA")
            ValidateZipCode_A1_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateAccountName_A1_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateAccountName_A1_FA = True
    setTblinfo_A1_FA
    
    Set rangecells = sourceSheet.Range("FA_A1_ForeignAccountNumber").Cells
    ReDim AccountNameA1_FA(end_A1_FA)
    For i = 1 To end_A1_FA
        AccountNameA1_FA(i) = rangecells.item(i).value
        If Len(AccountNameA1_FA(i)) = 0 Then
        'Exit For
        End If
        
        If Not chkCompulsory(AccountNameA1_FA(i)) Then
            msgbox_A1_FA ("* Please enter the Account Number at Sr. No  " & i & "  in Table A1 of Schedule FA")
            ValidateAccountName_A1_FA = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(AccountNameA1_FA(i)) Then
            msgbox_A1_FA ("* Account Number at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A1 of Schedule FA")
            ValidateAccountName_A1_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateStatus_A1_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateStatus_A1_FA = True
    setTblinfo_A1_FA
    Set rangecells = sourceSheet.Range("FA_A1_StatusBeneficiary").Cells
    ReDim StatusA1_FA(end_A1_FA)
    For i = 1 To end_A1_FA
        StatusA1_FA(i) = UCase(rangecells.item(i).value)
        'If StatusA1_FA(i) = "Owner" Then
        '    StatusA1_FA(i) = "OWNER"
        '    Else
            
        'End If
        
        If isdropdownblank(StatusA1_FA(i)) Then
            msgbox_A1_FA ("* Please select the Status at Sr. No  " & i & " in Table A1 of Schedule FA")
            ValidateStatus_A1_FA = False
            Exit Function
        End If
    Next
    
End Function

Function ValidateAccountOpeningDate_A1()
Dim rangecells As Range
Dim i As Long
    
ValidateAccountOpeningDate_A1 = True
setTblinfo_A1_FA
Set rangecells = sourceSheet.Range("FA_A1_AccOpeningDate").Cells
ReDim AccountOpeningDateA1_FA(end_A1_FA)
For i = 1 To end_A1_FA

AccountOpeningDateA1_FA(i) = rangecells.item(i).value

If Not chkCompulsory(AccountOpeningDateA1_FA(i)) Then
 msgbox_A1_FA ("* Please enter the Account Opening Date at Sr. No  " & i & "  in Table A1 of Schedule FA")
 ValidateAccountOpeningDate_A1 = False
 Exit Function
End If

If Not CheckDateddmmyyyy(AccountOpeningDateA1_FA(i)) Then
 msgbox_A1_FA ("* Account Opening Date at Sr. No  " & i & "  is invalid in Table A1 of Schedule FA")
 ValidateAccountOpeningDate_A1 = False
 Exit Function
End If

Next
    
End Function

Function ValidatePeakBal_A1_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidatePeakBal_A1_FA = True
    setTblinfo_A1_FA
    
    Set rangecells = sourceSheet.Range("FA_A1_PeakBal").Cells
    ReDim PeakBalA1_FA(end_A1_FA)
    For i = 1 To end_A1_FA
        PeakBalA1_FA(i) = rangecells.item(i).value
        If Len(PeakBalA1_FA(i)) = 0 Then
        'Exit For
        End If
        
        If Not chkCompulsory(PeakBalA1_FA(i)) Then
            msgbox_A1_FA ("* Please enter the Peak Balance During the  Period at Sr. No  " & i & " in Table A1 of Schedule FA")
            ValidatePeakBal_A1_FA = False
            Exit Function
        End If
       
        If Not checkfieldspecialcharacter(PeakBalA1_FA(i)) Then
            msgbox_A1_FA ("* Peak BalanceDuring the  Period at Sr. No  " & i & "   characters < > & ' " & Chr(34) & " are not allowed in Table A1 of Schedule FA")
            ValidatePeakBal_A1_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateCloseBalance_A1_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateCloseBalance_A1_FA = True
    setTblinfo_A1_FA
    
    Set rangecells = sourceSheet.Range("FA_A1_ClosingBal").Cells
    ReDim CloseBalA1_FA(end_A1_FA)
    For i = 1 To end_A1_FA
        CloseBalA1_FA(i) = rangecells.item(i).value
        If Len(CloseBalA1_FA(i)) = 0 Then
        'Exit For
        End If
        
        If Not chkCompulsory(CloseBalA1_FA(i)) Then
            msgbox_A1_FA ("* Please enter the Closing Balance at Sr. No  " & i & " in Table A1 of Schedule FA")
            ValidateCloseBalance_A1_FA = False
            Exit Function
        End If
       
        If Not checkfieldspecialcharacter(CloseBalA1_FA(i)) Then
            msgbox_A1_FA ("* Closing Balance at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A1 of Schedule FA")
            ValidateCloseBalance_A1_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateGrossInterest_A1_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateGrossInterest_A1_FA = True
    setTblinfo_A1_FA
    
    Set rangecells = sourceSheet.Range("FA_A1_Grossinterest").Cells
    ReDim GrossInterest_A1_FA(end_A1_FA)
    For i = 1 To end_A1_FA
        GrossInterest_A1_FA(i) = rangecells.item(i).value
        
        
        If Not chkCompulsory(GrossInterest_A1_FA(i)) Then
            msgbox_A1_FA ("* Please enter the Gross interest paid or credited at Sr. No  " & i & " in Table A1 of Schedule FA")
            ValidateGrossInterest_A1_FA = False
            Exit Function
        End If
        
'        If (GrossInterest_A1_FA(i) < 0) Then
'            msgbox_A1_FA ("* Gross interest paid or credited at Sr. No  " & i & "  in Sheet TR_FA  should be Non negative, no decimal, upto 99,999,999,999,999")
'            ValidateGrossInterest_A1_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(GrossInterest_A1_FA(i)) Then
            msgbox_A1_FA ("* Gross interest paid or credited  at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Table A1 of Schedule FA")
            ValidateGrossInterest_A1_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidatesheetA2FA() As Boolean
 Dim i As Long
 
    ValidatesheetA2FA = True
    If Not ValidateCountry_A2_FA() Then ValidatesheetA2FA = False
    If Not isdropdownblank(sourceSheet.Range("FA_A2_Country").item(1).value) And (Len(sourceSheet.Range("FA_A2_Country").item(1).value) > 0) Then
    If Not ValidateBankName_A2_FA() Then ValidatesheetA2FA = False
    If Not ValidateBankAddress_A2_FA() Then ValidatesheetA2FA = False
    If Not ValidateZipCode_A2_FA() Then ValidatesheetA2FA = False
    If Not ValidateAccountName_A2_FA() Then ValidatesheetA2FA = False
    If Not ValidateStatus_A2_FA() Then ValidatesheetA2FA = False
    If Not ValidateAccountOpeningDate_A2() Then ValidatesheetA2FA = False
    If Not ValidatePeakBal_A2_FA() Then ValidatesheetA2FA = False
    If Not ValidateCloseBalance_A2_FA() Then ValidatesheetA2FA = False
    If Not ValidateGrossInterest_A2_FA() Then ValidatesheetA2FA = False
    End If
    
    
    setTblinfo_A21_FA
'    If (end_A2_FA <> end_A21_FA) Then
''        msgbox_A2_FA ("* Details of Country in Table A2 of Schedule FA are Mandatory")
'
'      'SIT-69222 changed by sai
'        msgbox_A2_FA ("* Please select the Country Name and Code in Table A2 of Schedule FA.")
'
'        ValidatesheetA2FA = False
'        Exit Function
'    End If
    
'    If (CountA2 <> CountA21) Then
'        msgbox_A2_FA ("* Enter all mandatory field in Table A2 of Schedule FA")
'        ValidatesheetA2FA = False
'        Exit Function
'    End If


'Added by Aavula on 08/08/2024 SIT-69222

For i = 1 To Sheet27.Range("FA_A2_Country").count


        If Sheet27.Range("FA_A2_Country").Cells(i, 1).value <> "" And Sheet27.Range("FA_A2_Country").Cells(i, 1).value <> "(Select)" Or Sheet27.Range("FA_A2_BankName").Cells(i, 1).value <> "" Or Sheet27.Range("FA_A2_BankAdd").Cells(i, 1).value <> "" Or Sheet27.Range("FA_A2_ZipCode").Cells(i, 1).value <> "" Or _
        (Sheet27.Range("FA_A2_ForeignAccountNumber").Cells(i, 1).value <> "" Or Sheet27.Range("FA_A2_StatusBeneficiary").Cells(i, 1).value <> "" And Sheet27.Range("FA_A2_StatusBeneficiary").Cells(i, 1).value <> "(Select)") Or _
        Sheet27.Range("FA_A2_AccOpeningDate").Cells(i, 1).value <> "" Or Sheet27.Range("FA_A2_PeakBal").Cells(i, 1).value <> "" Or Sheet27.Range("FA_A2_ClosingBal").Cells(i, 1).value <> "" Or Sheet27.Range("FA_A2_Grossinterest_Nature").Cells(i, 1).value <> "" Then ' Or Sheet27.Range("FA_A2_Grossinterest").Cells(i, 1).value <> ""
   
        
        If Sheet27.Range("FA_A2_Country").Cells(i, 1).value = "" Or Sheet27.Range("FA_A2_Country").Cells(i, 1).value = "(Select)" Then
            ValidatesheetA2FA = False
            msgbox_A2_FA ("Please select the Country Name and Code in Table A2 of Schedule FA at Sr. No " & i & ".")
        End If


         If Sheet27.Range("FA_A2_BankName").Cells(i, 1).value = "" Then
            ValidatesheetA2FA = False
             msgbox_A2_FA ("*Please enter the Name of the financial institution in Table A2 of Schedule FA at Sr. No " & i & ".")
        End If


        If Sheet27.Range("FA_A2_BankAdd").Cells(i, 1).value = "" Then
            ValidatesheetA2FA = False
            msgbox_A2_FA ("Please enter the Address of the financial institution in Table A2 of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_A2_ZipCode").Cells(i, 1).value = "" Then
            ValidatesheetA2FA = False
            msgbox_A2_FA ("Please enter ZIP code, if  ZIP code is not available, then enter XXXXXX in Table A2 of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_A2_ForeignAccountNumber").Cells(i, 1).value = "" Then
            ValidatesheetA2FA = False
            msgbox_A2_FA ("Please enter the Account Number in Table A2 of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_A2_StatusBeneficiary").Cells(i, 1).value = "" Or Sheet27.Range("FA_A2_StatusBeneficiary").Cells(i, 1).value = "(Select)" Then
            ValidatesheetA2FA = False
            msgbox_A2_FA ("Please select the status in Table A2 of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_A2_AccOpeningDate").Cells(i, 1).value = "" Then
            ValidatesheetA2FA = False
            msgbox_A2_FA ("Please enter the date of account opening in Table A2 of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_A2_PeakBal").Cells(i, 1).value = "" Then
            ValidatesheetA2FA = False
            msgbox_A2_FA ("Please enter the Peak Balance During the Period in Table A2 of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_A2_ClosingBal").Cells(i, 1).value = "" Then
            ValidatesheetA2FA = False
            msgbox_A2_FA ("Please enter the Closing Balance in Table A2 of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_A2_Grossinterest_Nature").Cells(i, 1).value = "" Then
            ValidatesheetA2FA = False
            msgbox_A2_FA ("Please enter amount against the nature of amount in Table A2 of Schedule FA at Sr. No " & i & ".")
        End If
        
'        If Sheet27.Range("FA_A2_Grossinterest").Cells(i, 1).value = "" Then
'            ValidatesheetA2FA = False
'            msgbox_A2_FA ("Please enter the gross amount paid/credited to the account during the period in Table A2 of Schedule FA at Sr. No " & i & ".")
'        End If
        
    End If
Next i








    
End Function


Function ValidateCountry_A2_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim CountryName As String
    
    ValidateCountry_A2_FA = True
    setTblinfo_A2_FA
    
    Set rangecells = sourceSheet.Range("FA_A2_Country").Cells
    ReDim CountryCodeA2_FA(end_A2_FA)
    ReDim CountryNameA2_FA(end_A2_FA)
    For i = 1 To end_A2_FA
    noOfProcessSub = end_A2_FA
        CountryName = rangecells.item(i).value
        
'        If isdropdownblank(CountryName) Then
'            msgbox_A2_FA ("* Please select Country Name and Code at Sr. No  " & i & "  in Table A2 of Schedule FA")
'            ValidateCountry_A2_FA = False
'            Exit Function
'        End If
        
        
        If CountryName <> "" Then
        CountryNameA2_FA(i) = Mid(CountryName, WorksheetFunction.Search("-", CountryName) + 1, Len(CountryName))
         CountryCodeA2_FA(i) = Mid(CountryName, 1, WorksheetFunction.Search("-", CountryName) - 1)
         End If
        
        If Len(CountryCodeA2_FA(i)) = 0 Then
            Exit For
        End If
        
        If Not chkCompulsory(CountryCodeA2_FA(i)) Then
            msgbox_A2_FA ("* Country Name at Sr. No  " & i & "  in Table A2 of Schedule FA")
            ValidateCountry_A2_FA = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(CountryCodeA2_FA(i)) Then
            msgbox_A2_FA ("* Country Name at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Table A2 of Schedule FA")
            ValidateCountry_A2_FA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function


Function ValidateBankName_A2_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
  
    ValidateBankName_A2_FA = True
    setTblinfo_A2_FA
    
    Set rangecells = sourceSheet.Range("FA_A2_BankName").Cells
    ReDim BankNameA2_FA(end_A2_FA)
    For i = 1 To end_A2_FA
        BankNameA2_FA(i) = rangecells.item(i).value
        If Len(BankNameA2_FA(i)) = 0 Then
            'Exit For
        End If
        
'        If Not chkCompulsory(BankNameA2_FA(i)) Then
''            msgbox_A2_FA ("* Please enetr the Name of  financial institution at Sr. No  " & i & " in Table A2 of Schedule FA")
'
'             msgbox_A2_FA ("* Please enter the Name of the financial institution in Table A2 of Schedule FA.")
'
'
'            ValidateBankName_A2_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(BankNameA2_FA(i)) Then
            msgbox_A2_FA ("* Name of  financial institution at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A2 of Schedule FA")
            ValidateBankName_A2_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateBankAddress_A2_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateBankAddress_A2_FA = True
    setTblinfo_A2_FA
    
    Set rangecells = sourceSheet.Range("FA_A2_BankAdd").Cells
    ReDim BankAddA2_FA(end_A2_FA)
    For i = 1 To end_A2_FA
        BankAddA2_FA(i) = rangecells.item(i).value
        If Len(BankAddA2_FA(i)) = 0 Then
            'Exit For
        End If
        
'        If Not chkCompulsory(BankAddA2_FA(i)) Then
'            msgbox_A2_FA ("* Please enter the Address of  financial institution at Sr. No  " & i & "  in Table A2 of Schedule FA")
'            ValidateBankAddress_A2_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(BankAddA2_FA(i)) Then
            msgbox_A2_FA ("* Address of  financial institution at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Table A2 of Schedule FA")
            ValidateBankAddress_A2_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateZipCode_A2_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateZipCode_A2_FA = True
    setTblinfo_A2_FA
    
    Set rangecells = sourceSheet.Range("FA_A2_ZipCode").Cells
    ReDim ZipCodeA2_FA(end_A2_FA)
    For i = 1 To end_A2_FA
        ZipCodeA2_FA(i) = rangecells.item(i).value
        If Len(ZipCodeA2_FA(i)) = 0 Then
            'Exit For
        End If
        
'        If Not chkCompulsory(ZipCodeA2_FA(i)) Then
'            msgbox_A2_FA ("* Please enter the ZipCode at Sr. No  " & i & "If  ZIP code is not available, then enter XXXXXX in Table A2 of Schedule FA")
'            ValidateZipCode_A2_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(ZipCodeA2_FA(i)) Then
            msgbox_A2_FA ("* ZipCode at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A2 of Schedule FA")
            ValidateZipCode_A2_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateAccountName_A2_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateAccountName_A2_FA = True
    setTblinfo_A2_FA
    
    Set rangecells = sourceSheet.Range("FA_A2_ForeignAccountNumber").Cells
    ReDim AccountNameA2_FA(end_A2_FA)
    For i = 1 To end_A2_FA
        AccountNameA2_FA(i) = rangecells.item(i).value
        If Len(AccountNameA2_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(AccountNameA2_FA(i)) Then
'            msgbox_A2_FA ("* Please enter the Account Number at Sr. No  " & i & "  in Table A2 of Schedule FA")
'            ValidateAccountName_A2_FA = False
'            Exit Function
'        End If
'
        If Not checkfieldspecialcharacter(AccountNameA2_FA(i)) Then
            msgbox_A2_FA ("* Account Number at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A2 of Schedule FA")
            ValidateAccountName_A2_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateStatus_A2_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateStatus_A2_FA = True
    setTblinfo_A2_FA
    Set rangecells = sourceSheet.Range("FA_A2_StatusBeneficiary").Cells
    ReDim StatusA2_FA(end_A2_FA)
    For i = 1 To end_A2_FA
        StatusA2_FA(i) = UCase(rangecells.item(i).value)
        'If StatusA2_FA(i) = "Owner" Then
        '    StatusA2_FA(i) = "OWNER"
        '    Else
            
        'End If
        
'        If isdropdownblank(StatusA2_FA(i)) Then
'            msgbox_A2_FA ("* Please select the Status at Sr. No  " & i & "  in Table A2 of Schedule FA")
'            ValidateStatus_A2_FA = False
'            Exit Function
'        End If
    Next
    
End Function

Function ValidateAccountOpeningDate_A2()
Dim rangecells As Range
Dim i As Long
    
ValidateAccountOpeningDate_A2 = True
setTblinfo_A2_FA
Set rangecells = sourceSheet.Range("FA_A2_AccOpeningDate").Cells
ReDim AccountOpeningDateA2_FA(end_A2_FA)
For i = 1 To end_A2_FA

AccountOpeningDateA2_FA(i) = rangecells.item(i).value

'If Not chkCompulsory(AccountOpeningDateA2_FA(i)) Then
'' msgbox_A2_FA ("* Please enter teh Account Opening Date at Sr. No  " & i & "  in Table A2 of Schedule FA")
'
' 'SIT-69222 changed by sai
'      msgbox_A2_FA ("* Please enter the date of account opening in Table A2 of Schedule FA.")
'
' ValidateAccountOpeningDate_A2 = False
' Exit Function
'End If

If Not CheckDateddmmyyyy(AccountOpeningDateA2_FA(i)) Then
 msgbox_A2_FA ("* Account Opening Date at Sr. No  " & i & "is invalid in Table A2 of Schedule FA")
 ValidateAccountOpeningDate_A2 = False
 Exit Function
End If

Next
    
End Function

Function ValidatePeakBal_A2_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidatePeakBal_A2_FA = True
    setTblinfo_A2_FA
    
    Set rangecells = sourceSheet.Range("FA_A2_PeakBal").Cells
    ReDim PeakBalA2_FA(end_A2_FA)
    For i = 1 To end_A2_FA
        PeakBalA2_FA(i) = rangecells.item(i).value
        If Len(PeakBalA2_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(PeakBalA2_FA(i)) Then
'            msgbox_A2_FA ("* Please enter the Peak Balance During the  Period at Sr. No  " & i & "in Table A2 of Schedule FA")
'            ValidatePeakBal_A2_FA = False
'            Exit Function
'        End If
       
        If Not checkfieldspecialcharacter(PeakBalA2_FA(i)) Then
            msgbox_A2_FA ("* Peak Balance During the  Period at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Table A2 of Schedule FA")
            ValidatePeakBal_A2_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateCloseBalance_A2_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateCloseBalance_A2_FA = True
    setTblinfo_A2_FA
    
    Set rangecells = sourceSheet.Range("FA_A2_ClosingBal").Cells
    ReDim CloseBalA2_FA(end_A2_FA)
    For i = 1 To end_A2_FA
        CloseBalA2_FA(i) = rangecells.item(i).value
        If Len(CloseBalA2_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(CloseBalA2_FA(i)) Then
'            msgbox_A2_FA ("* Please enter the Closing Balance at Sr. No  " & i & " in Table A2 of Schedule FA")
'            ValidateCloseBalance_A2_FA = False
'            Exit Function
'        End If
       
        If Not checkfieldspecialcharacter(CloseBalA2_FA(i)) Then
            msgbox_A2_FA ("* Closing Balance at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A2 of Schedule FA")
            ValidateCloseBalance_A2_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateGrossInterest_A2_FA() As Boolean
    Dim rangecells, rangecells1 As Range
    Dim i As Long
    
    ValidateGrossInterest_A2_FA = True
    setTblinfo_A2_FA
    
    Set rangecells = sourceSheet.Range("FA_A2_Grossinterest").Cells
    Set rangecells1 = sourceSheet.Range("FA_A2_Grossinterest_Nature").Cells
    ReDim GrossInterest_A2_FA(end_A2_FA)
    ReDim GrossInterestNature_A2_FA(end_A2_FA)
    For i = 1 To end_A2_FA
        GrossInterest_A2_FA(i) = rangecells.item(i).value
        GrossInterestNature_A2_FA(i) = rangecells1.item(i).value
        
'        If isdropdownblank(GrossInterestNature_A2_FA(i)) Then
'            msgbox_A2_FA ("* Please Select the nature of amount from the drop downs provided at Sr. No  " & i & "  in Table A2 of Schedule FA")
'            ValidateGrossInterest_A2_FA = False
'            Exit Function
'        End If
        
        If GrossInterestNature_A2_FA(i) <> "" And GrossInterestNature_A2_FA(i) <> "(Select)" Then
        If Not chkCompulsory(GrossInterest_A2_FA(i)) Then
            msgbox_A2_FA ("* Please enter amount against the nature of amount at Sr. No  " & i & "  in Table A2 of Schedule FA ")
            ValidateGrossInterest_A2_FA = False
            Exit Function
        End If
        End If
        
        If (GrossInterest_A2_FA(i) < 0) Then
            msgbox_A2_FA ("* Amount in interest paid/credited in the account at Sr. No  " & i & "    should be Non negative, no decimal, upto 99,999,999,999,999 in Table A2 of Schedule FA")
            ValidateGrossInterest_A2_FA = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(GrossInterest_A2_FA(i)) Then
            msgbox_A2_FA ("* Amount in interest paid/credited in the account at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A2 of Schedule FA")
            ValidateGrossInterest_A2_FA = False
            Exit Function
        End If
    Next
End Function


Function ValidatesheetA3FA() As Boolean
 
 Dim i As Long
 
    ValidatesheetA3FA = True
    If Not ValidateCountry_A3_FA() Then ValidatesheetA3FA = False
    If Not isdropdownblank(sourceSheet.Range("FA_A3_Country").item(1).value) And (Len(sourceSheet.Range("FA_A3_Country").item(1).value) > 0) Then
    If Not ValidateBankName_A3_FA() Then ValidatesheetA3FA = False
    If Not ValidateBankAddress_A3_FA() Then ValidatesheetA3FA = False
    If Not ValidateZipCode_A3_FA() Then ValidatesheetA3FA = False
    If Not validateNatureOfEntity_A3_FA() Then ValidatesheetA3FA = False
    If Not ValidateAccountOpeningDate_A3() Then ValidatesheetA3FA = False
    If Not ValidateInitialvalue_Investment_A3_FA() Then ValidatesheetA3FA = False
    If Not ValidatePeakBal_A3_FA() Then ValidatesheetA3FA = False
    If Not ValidateCloseBalance_A3_FA() Then ValidatesheetA3FA = False
    If Not ValidateTotalGrossAmount_A3_FA() Then ValidatesheetA3FA = False
    If Not ValidateTotalGrossProceeds_A3_FA() Then ValidatesheetA3FA = False
    End If
    
   setTblinfo_A31_FA
'    If (end_A3_FA <> end_A31_FA) Then
''        msgbox_A3_FA ("* Details of Country in Table A3 of Schedule FA are mandatory")
'
'       'SIT-69222 changed by sai
'        msgbox_A3_FA ("*Please select the Country Name and Code in Table A3 of Schedule FA.")
'
'        ValidatesheetA3FA = False
'        Exit Function
'    End If
    
'    If (CountA3 <> CountA31) Then
'        msgbox_A3_FA ("* Enter all mandatory field in Table A3 of Schedule FA")
'        ValidatesheetA3FA = False
'        Exit Function
'    End If

'Added by Aavula on 08/08/2024 SIT-69222

 For i = 1 To Sheet27.Range("FA_A3_Country").count

        If Sheet27.Range("FA_A3_Country").Cells(i, 1).value <> "" And Sheet27.Range("FA_A3_Country").Cells(i, 1).value <> "(Select)" Or Sheet27.Range("FA_A3_BankName").Cells(i, 1).value <> "" Or Sheet27.Range("FA_A3_BankAdd").Cells(i, 1).value <> "" Or Sheet27.Range("FA_A3_ZipCode").Cells(i, 1).value <> "" Or _
        (Sheet27.Range("FA_A3_NatureOfEntity").Cells(i, 1).value <> "" Or Sheet27.Range("FA_A3_AccOpeningDate").Cells(i, 1).value <> "") Or _
        Sheet27.Range("FA_A3_initialvalue").Cells(i, 1).value <> "" Or Sheet27.Range("FA_A3_PeakBal").Cells(i, 1).value <> "" Or Sheet27.Range("FA_A3_ClosingBal").Cells(i, 1).value <> "" Or Sheet27.Range("FA_A3_Totalgrossamount").Cells(i, 1).value <> "" Or Sheet27.Range("FA_A3_Totalgrosproceeds").Cells(i, 1).value <> "" Then
   
        
        If Sheet27.Range("FA_A3_Country").Cells(i, 1).value = "" Or Sheet27.Range("FA_A3_Country").Cells(i, 1).value = "(Select)" Then
            ValidatesheetA3FA = False
            msgbox_A3_FA ("*Please select the Country Name and Code in Table A3 of Schedule FA at Sr. No " & i & ".")
        End If


         If Sheet27.Range("FA_A3_BankName").Cells(i, 1).value = "" Then
            ValidatesheetA3FA = False
             msgbox_A3_FA ("*Please enter the Name of the  entity.in Table A3 of Schedule FA at Sr. No " & i & ".")
        End If


        If Sheet27.Range("FA_A3_BankAdd").Cells(i, 1).value = "" Then
            ValidatesheetA3FA = False
            msgbox_A3_FA ("*Please enter the Address of the  entity.in Table A3 of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_A3_ZipCode").Cells(i, 1).value = "" Then
            ValidatesheetA3FA = False
            msgbox_A3_FA ("*Please enter ZIP code, if  ZIP code is not available, then enter XXXXXX. in Table A3 of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_A3_NatureOfEntity").Cells(i, 1).value = "" Then
            ValidatesheetA3FA = False
            msgbox_A3_FA ("*Please enter the Nature of entity. in Table A3 of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_A3_AccOpeningDate").Cells(i, 1).value = "" Then
            ValidatesheetA3FA = False
            msgbox_A3_FA ("*Please enter the Date of acquiring the interest in Table A3 of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_A3_initialvalue").Cells(i, 1).value = "" Then
            ValidatesheetA3FA = False
            msgbox_A3_FA ("*Please enter the Initial value of the investment. in Table A3 of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_A3_PeakBal").Cells(i, 1).value = "" Then
            ValidatesheetA3FA = False
            msgbox_A3_FA ("*Please enter the Peak value of investment during the Period. in Table A3 of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_A3_ClosingBal").Cells(i, 1).value = "" Then
            ValidatesheetA3FA = False
            msgbox_A3_FA ("*Please enter the Closing value. in Table A3 of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_A3_Totalgrossamount").Cells(i, 1).value = "" Then
            ValidatesheetA3FA = False
            msgbox_A3_FA ("*Please enter the Total gross amount paid/credited with respect to the holding during the period in Table A3 of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_A3_Totalgrosproceeds").Cells(i, 1).value = "" Then
            ValidatesheetA3FA = False
            msgbox_A3_FA ("*Please enter the Total gross proceeds from sale or redemption of investment during the period in Table A3 of Schedule FA at Sr. No " & i & ".")
        End If
        
    End If
Next i

End Function


Function ValidateCountry_A3_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim CountryName As String
    
    ValidateCountry_A3_FA = True
    setTblinfo_A3_FA
    
    Set rangecells = sourceSheet.Range("FA_A3_Country").Cells
    ReDim CountryCodeA3_FA(end_A3_FA)
    ReDim CountryNameA3_FA(end_A3_FA)
    For i = 1 To end_A3_FA
    noOfProcessSub = end_A3_FA
        CountryName = rangecells.item(i).value
        
'        If isdropdownblank(CountryName) Then
'           msgbox_A3_FA ("* Please select the Country Name and  at Sr. No  " & i & " in Table A3 of Schedule FA")
'            ValidateCountry_A3_FA = False
'            Exit Function
'        End If
        
        If CountryName <> "" Then
        CountryNameA3_FA(i) = Mid(CountryName, WorksheetFunction.Search("-", CountryName) + 1, Len(CountryName))
         CountryCodeA3_FA(i) = Mid(CountryName, 1, WorksheetFunction.Search("-", CountryName) - 1)
         End If
        
        If Len(CountryCodeA3_FA(i)) = 0 Then

            Exit For
        End If
        
        If Not chkCompulsory(CountryCodeA3_FA(i)) Then
            msgbox_A3_FA ("* Please select the Country Name and Code  at Sr. No  " & i & " in Table A3 of Schedule FA")
            ValidateCountry_A3_FA = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(CountryCodeA3_FA(i)) Then
            msgbox_A3_FA ("* Country Name and Code at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A2 of Schedule FA")
            ValidateCountry_A3_FA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function


Function ValidateBankName_A3_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
  
    ValidateBankName_A3_FA = True
    setTblinfo_A3_FA
    
    Set rangecells = sourceSheet.Range("FA_A3_BankName").Cells
    ReDim BankNameA3_FA(end_A3_FA)
    For i = 1 To end_A3_FA
        BankNameA3_FA(i) = rangecells.item(i).value
        If Len(BankNameA3_FA(i)) = 0 Then
            'Exit For
        End If
        
'        If Not chkCompulsory(BankNameA3_FA(i)) Then
''            msgbox_A3_FA ("* Please enter the Name of entity at Sr. No  " & i & "  in Table A3 of Schedule FA")
'
'
'            msgbox_A3_FA ("*Please enter the Name of the entity in Table A3 of Schedule FA.")
'
'
'            ValidateBankName_A3_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(BankNameA3_FA(i)) Then
            msgbox_A3_FA ("* Name of entity at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A3 of Schedule FA")
            ValidateBankName_A3_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateBankAddress_A3_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateBankAddress_A3_FA = True
    setTblinfo_A3_FA
    
    Set rangecells = sourceSheet.Range("FA_A3_BankAdd").Cells
    ReDim BankAddA3_FA(end_A3_FA)
    For i = 1 To end_A3_FA
        BankAddA3_FA(i) = rangecells.item(i).value
        If Len(BankAddA3_FA(i)) = 0 Then
            'Exit For
        End If
        
'        If Not chkCompulsory(BankAddA3_FA(i)) Then
'            msgbox_A3_FA ("* Please enter the Address of entity at Sr. No  " & i & " in Table A3 of Schedule FA")
'            ValidateBankAddress_A3_FA = False
'            Exit Function
'        End If
'
        If Not checkfieldspecialcharacter(BankAddA3_FA(i)) Then
            msgbox_A3_FA ("* Address of entity at Sr. No  " & i & "   characters < > & ' " & Chr(34) & " are not allowed in Table A3 of Schedule FA")
            ValidateBankAddress_A3_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateZipCode_A3_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateZipCode_A3_FA = True
    setTblinfo_A3_FA
    
    Set rangecells = sourceSheet.Range("FA_A3_ZipCode").Cells
    ReDim ZipCodeA3_FA(end_A3_FA)
    For i = 1 To end_A3_FA
        ZipCodeA3_FA(i) = rangecells.item(i).value
        If Len(ZipCodeA3_FA(i)) = 0 Then
            'Exit For
        End If
'
'        If Not chkCompulsory(ZipCodeA3_FA(i)) Then
'            msgbox_A3_FA ("* Please enetr the ZipCode at Sr. No  " & i & "If  ZIP code is not available, then enter XXXXXX in Table A3 of Schedule FA")
'            ValidateZipCode_A3_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(ZipCodeA3_FA(i)) Then
            msgbox_A3_FA ("* ZipCode at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A3 of Schedule FA")
            ValidateZipCode_A3_FA = False
            Exit Function
        End If
    Next
End Function

Function validateNatureOfEntity_A3_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    validateNatureOfEntity_A3_FA = True
    setTblinfo_C_FA
    Set rangecells = sourceSheet.Range("FA_A3_NatureOfEntity").Cells
    ReDim NatureOfEntityA3_FA(end_A3_FA)
    For i = 1 To end_A3_FA
        NatureOfEntityA3_FA(i) = rangecells.item(i).value
'        If Not chkCompulsory(NatureOfEntityA3_FA(i)) Then
'            msgbox_A3_FA ("* Please enter the Nature Of Entity at Sr. No  " & i & " in Table A3 of Schedule FA")
'            validateNatureOfEntity_A3_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(NatureOfEntityA3_FA(i)) Then
        msgbox_A3_FA ("* Nature Of Entity  at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A3 of Schedule FA")
        validateNatureOfEntity_A3_FA = False
        Exit Function
        End If
        
        Next
    
End Function

Function ValidateAccountOpeningDate_A3()
Dim rangecells As Range
Dim i As Long
    
ValidateAccountOpeningDate_A3 = True
setTblinfo_A3_FA
Set rangecells = sourceSheet.Range("FA_A3_AccOpeningDate").Cells
ReDim AccountOpeningDateA3_FA(end_A3_FA)
For i = 1 To end_A3_FA

AccountOpeningDateA3_FA(i) = rangecells.item(i).value

'If Not chkCompulsory(AccountOpeningDateA3_FA(i)) Then
' msgbox_A3_FA ("* Please enter the Date of acquiring the interest at Sr. No  " & i & " in Table A3 of Schedule FA")
' ValidateAccountOpeningDate_A3 = False
' Exit Function
'End If

If Not CheckDateddmmyyyy(AccountOpeningDateA3_FA(i)) Then
 msgbox_A3_FA ("* Date of acquiring the interest at Sr. No  " & i & "is invalid in Table A3 of Schedule FA")
 ValidateAccountOpeningDate_A3 = False
 Exit Function
End If

Next
    
End Function

Function ValidateInitialvalue_Investment_A3_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateInitialvalue_Investment_A3_FA = True
    setTblinfo_A3_FA
    
    Set rangecells = sourceSheet.Range("FA_A3_initialvalue").Cells
    ReDim Initialvalue_Investment(end_A3_FA)
    For i = 1 To end_A3_FA
        Initialvalue_Investment(i) = rangecells.item(i).value
        
        
'        If Not chkCompulsory(Initialvalue_Investment(i)) Then
'            msgbox_A3_FA ("* Please enter the Initial value of the investment. in Table A3 of Schedule FA at Sr. No  " & i & ".")
'            ValidateInitialvalue_Investment_A3_FA = False
'            Exit Function
'        End If
        
        If (Initialvalue_Investment(i) < 0) Then
            msgbox_A3_FA ("* Initial value of the investment at Sr. No  " & i & "  should be Non negative, no decimal, upto 99,999,999,999,999 in Table A3 of Schedule FA")
            ValidateInitialvalue_Investment_A3_FA = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(Initialvalue_Investment(i)) Then
            msgbox_A3_FA ("* Initial value of the investment  at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A3 of Schedule FA")
            ValidateInitialvalue_Investment_A3_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidatePeakBal_A3_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidatePeakBal_A3_FA = True
    setTblinfo_A3_FA
    
    Set rangecells = sourceSheet.Range("FA_A3_PeakBal").Cells
    ReDim PeakBalA3_FA(end_A3_FA)
    For i = 1 To end_A3_FA
        PeakBalA3_FA(i) = rangecells.item(i).value
        If Len(PeakBalA3_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(PeakBalA3_FA(i)) Then
'            msgbox_A3_FA ("* Please enter the Peak value of investment during the Period at Sr. No  " & i & " in Table A3 of Schedule FA")
'            ValidatePeakBal_A3_FA = False
'            Exit Function
'        End If
       
        If Not checkfieldspecialcharacter(PeakBalA3_FA(i)) Then
            msgbox_A3_FA ("* Peak value of investment during the Period at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A3 of Schedule FA")
            ValidatePeakBal_A3_FA = False
            Exit Function
        End If
    Next
End Function


Function ValidateCloseBalance_A3_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateCloseBalance_A3_FA = True
    setTblinfo_A3_FA
    
    Set rangecells = sourceSheet.Range("FA_A3_ClosingBal").Cells
    ReDim CloseBalA3_FA(end_A3_FA)
    For i = 1 To end_A3_FA
        CloseBalA3_FA(i) = rangecells.item(i).value
        If Len(CloseBalA3_FA(i)) = 0 Then
        'Exit For
        End If
'
'        If Not chkCompulsory(CloseBalA3_FA(i)) Then
''            msgbox_A3_FA ("* Please enter the Closing Balance at Sr. No  " & i & " in Table A3 of Schedule FA")
'
'          'SIT-69222 changed by sai
'            msgbox_A3_FA ("*Please enter the Closing value in Table A3 of Schedule FA.")
'
'            ValidateCloseBalance_A3_FA = False
'            Exit Function
'        End If
       
        If Not checkfieldspecialcharacter(CloseBalA3_FA(i)) Then
            msgbox_A3_FA ("* Closing Balance at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A3 of Schedule FA")
            ValidateCloseBalance_A3_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotalGrossAmount_A3_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateTotalGrossAmount_A3_FA = True
    setTblinfo_A3_FA
    
    Set rangecells = sourceSheet.Range("FA_A3_Totalgrossamount").Cells
    ReDim TotalGrossAmount_A3_FA(end_A3_FA)
    For i = 1 To end_A3_FA
        TotalGrossAmount_A3_FA(i) = rangecells.item(i).value
        
        
'        If Not chkCompulsory(TotalGrossAmount_A3_FA(i)) Then
''            msgbox_A3_FA ("* Please enter the Total gross amount paid/credited at Sr. No  " & i & " in Table A3 of Schedule FA")
'        'SIT-69222 changed by sai
'            msgbox_A3_FA ("*Please enter the Total gross amount paid/credited with respect to the holding during the period in Table A3 of Schedule FA.")
'
'            ValidateTotalGrossAmount_A3_FA = False
'            Exit Function
'        End If
       
        If Not checkfieldspecialcharacter(TotalGrossAmount_A3_FA(i)) Then
            msgbox_A3_FA ("* Total gross amount paid/credited at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Table A3 of Schedule FA")
            ValidateTotalGrossAmount_A3_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotalGrossProceeds_A3_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateTotalGrossProceeds_A3_FA = True
    setTblinfo_A3_FA
    
    Set rangecells = sourceSheet.Range("FA_A3_Totalgrosproceeds").Cells
    ReDim TotalGrossProceeds_A3_FA(end_A3_FA)
    For i = 1 To end_A3_FA
        TotalGrossProceeds_A3_FA(i) = rangecells.item(i).value
        
        
'        If Not chkCompulsory(TotalGrossProceeds_A3_FA(i)) Then
''            msgbox_A3_FA ("* Please enter the Total gross proceeds from sale or redemption at Sr. No  " & i & " in Table A3 of Schedule FA")
'           'SIT-69222 changed by sai
'            msgbox_A3_FA ("* Please enter the Total gross proceeds from sale or redemption of investment during the period in Table A3 of Schedule FA.")
'
'            ValidateTotalGrossProceeds_A3_FA = False
'            Exit Function
'        End If
       
        If Not checkfieldspecialcharacter(TotalGrossProceeds_A3_FA(i)) Then
            msgbox_A3_FA ("* Total gross proceeds from sale or redemption at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A3 of Schedule FA")
            ValidateTotalGrossProceeds_A3_FA = False
            Exit Function
        End If
    Next
End Function


Function ValidatesheetA4FA() As Boolean
 Dim i As Long
    ValidatesheetA4FA = True
    If Not ValidateCountry_A4_FA() Then ValidatesheetA4FA = False
    If Not isdropdownblank(sourceSheet.Range("FA_A4_Country").item(1).value) And (Len(sourceSheet.Range("FA_A4_Country").item(1).value) > 0) Then
    If Not ValidateBankName_A4_FA() Then ValidatesheetA4FA = False
    If Not ValidateBankAddress_A4_FA() Then ValidatesheetA4FA = False
    If Not ValidateZipCode_A4_FA() Then ValidatesheetA4FA = False
    If Not ValidateAccountOpeningDate_A4() Then ValidatesheetA4FA = False
    If Not ValidateTheCashValue_A4_FA() Then ValidatesheetA4FA = False
    If Not ValidateTotalGrossAmount_A4_FA() Then ValidatesheetA4FA = False
    End If
    
    setTblinfo_A41_FA
'    If (end_A4_FA <> end_A41_FA) Then
''        msgbox_A4_FA ("* Details of Country in Table A4 of Schedule FA are Mandatory")
'
'        'SIT-69222 changed by sai
'        msgbox_A4_FA ("* Please select the Country Name and Code in Table A4 of Schedule FA.")
'
'
'
'        ValidatesheetA4FA = False
'        Exit Function
'    End If
'
'    If (CountA4 <> CountA41) Then
'        msgbox_A4_FA ("* Enter all mandatory field in Table A4 of Schedule FA")
'        ValidatesheetA4FA = False
'        Exit Function
'    End If

'Added by Aavula on 08/08/2024 SIT-69222
 For i = 1 To Sheet27.Range("FA_A4_Country").count

        If Sheet27.Range("FA_A4_Country").Cells(i, 1).value <> "" And Sheet27.Range("FA_A4_Country").Cells(i, 1).value <> "(Select)" Or Sheet27.Range("FA_A4_BankName").Cells(i, 1).value <> "" Or Sheet27.Range("FA_A4_BankAdd").Cells(i, 1).value <> "" Or Sheet27.Range("FA_A4_ZipCode").Cells(i, 1).value <> "" Or _
        (Sheet27.Range("FA_A4_AccOpeningDate").Cells(i, 1).value <> "" Or Sheet27.Range("FA_A4_the_cashvalue").Cells(i, 1).value <> "") Or _
        Sheet27.Range("FA_A4_Total_grossamount").Cells(i, 1).value <> "" Then
   
        
        If Sheet27.Range("FA_A4_Country").Cells(i, 1).value = "" Or Sheet27.Range("FA_A4_Country").Cells(i, 1).value = "(Select)" Then
            ValidatesheetA4FA = False
            msgbox_A4_FA ("*Please select the Country Name and Code in Table A4 of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_A4_BankName").Cells(i, 1).value = "" Then
            ValidatesheetA4FA = False
            msgbox_A4_FA ("*Please enter the Name of financial institution in which insurance contract held in Table A4 of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_A4_BankAdd").Cells(i, 1).value = "" Then
            ValidatesheetA4FA = False
             msgbox_A4_FA ("*Please enter the Address of financial institution in Table A4 of Schedule FA at Sr. No " & i & "")
        End If
        ''''''''''''''
        
        If Sheet27.Range("FA_A4_ZipCode").Cells(i, 1).value = "" Then
            ValidatesheetA4FA = False
            msgbox_A4_FA ("*Please enter the Zipcode, if  ZIP code is not available, then enter XXXXXX in Table A4 of Schedule FA at Sr. No " & i & ".")
        End If


        If Sheet27.Range("FA_A4_AccOpeningDate").Cells(i, 1).value = "" Then
            ValidatesheetA4FA = False
            msgbox_A4_FA ("*Please enter the Date of contract  in Table A4 of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_A4_the_cashvalue").Cells(i, 1).value = "" Then
            ValidatesheetA4FA = False
            msgbox_A4_FA ("*Please enter the cash value or surrender value of the contract in Table A4 of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_A4_Total_grossamount").Cells(i, 1).value = "" Then
            ValidatesheetA4FA = False
            msgbox_A4_FA ("*Please enter total gross amount paid/credited with respect to the contract during the period in Table A4 of Schedule FA at Sr. No " & i & ".")
        End If
        
  End If
Next i



End Function

Function ValidateCountry_A4_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim CountryName As String
    
    ValidateCountry_A4_FA = True
    setTblinfo_A4_FA
    
    Set rangecells = sourceSheet.Range("FA_A4_Country").Cells
    ReDim CountryCodeA4_FA(end_A4_FA)
    ReDim CountryNameA4_FA(end_A4_FA)
    For i = 1 To end_A4_FA
    noOfProcessSub = end_A4_FA
        CountryName = rangecells.item(i).value
        
'         If isdropdownblank(CountryName) Then
'           msgbox_A4_FA ("* Please select the Country Name and Code at Sr. No  " & i & "  in Table A4 of Schedule FA")
'            ValidateCountry_A4_FA = False
'            Exit Function
'        End If
        
        
        
        If CountryName <> "" Then
        CountryNameA4_FA(i) = Mid(CountryName, WorksheetFunction.Search("-", CountryName) + 1, Len(CountryName))
         CountryCodeA4_FA(i) = Mid(CountryName, 1, WorksheetFunction.Search("-", CountryName) - 1)
         End If
        
        If Len(CountryCodeA4_FA(i)) = 0 Then
            Exit For
        End If
        
        If Not chkCompulsory(CountryCodeA4_FA(i)) Then
            msgbox_A4_FA ("* Please select the Country Name and Code at Sr. No  " & i & "  in Table A4 of Schedule FA")
            ValidateCountry_A4_FA = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(CountryCodeA4_FA(i)) Then
            msgbox_A4_FA ("* Country Name and Code  at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A4 of Schedule FA")
            ValidateCountry_A4_FA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function


Function ValidateBankName_A4_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
  
    ValidateBankName_A4_FA = True
    setTblinfo_A4_FA
    
    Set rangecells = sourceSheet.Range("FA_A4_BankName").Cells
    ReDim BankNameA4_FA(end_A4_FA)
    For i = 1 To end_A4_FA
        BankNameA4_FA(i) = rangecells.item(i).value
        If Len(BankNameA4_FA(i)) = 0 Then
            'Exit For
        End If
'
'        If Not chkCompulsory(BankNameA4_FA(i)) Then
''            msgbox_A4_FA ("* Please enter the Name of  financial institution at Sr. No  " & i & " in Table A4 of Schedule FA")
'
'        'SIT-69222 changed by sai
'            msgbox_A4_FA ("* Please enter the Name of financial institution in which insurance contract held in Table A4 of Schedule FA.")
'
'            ValidateBankName_A4_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(BankNameA4_FA(i)) Then
            msgbox_A4_FA ("* Name of  financial institution at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Table A4 of Schedule FA")
            ValidateBankName_A4_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateBankAddress_A4_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateBankAddress_A4_FA = True
    setTblinfo_A4_FA
    
    Set rangecells = sourceSheet.Range("FA_A4_BankAdd").Cells
    ReDim BankAddA4_FA(end_A4_FA)
    For i = 1 To end_A4_FA
        BankAddA4_FA(i) = rangecells.item(i).value
        If Len(BankAddA4_FA(i)) = 0 Then
            'Exit For
        End If
        
'        If Not chkCompulsory(BankAddA4_FA(i)) Then
'            msgbox_A4_FA ("* Please enter the Address of  financial institution at Sr. No  " & i & "  in Table A4 of Schedule FA")
'            ValidateBankAddress_A4_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(BankAddA4_FA(i)) Then
            msgbox_A4_FA ("* Address of  financial institution at Sr. No  " & i & "characters < > & ' " & Chr(34) & " are not allowed in Table A4 of Schedule FA")
            ValidateBankAddress_A4_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateZipCode_A4_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateZipCode_A4_FA = True
    setTblinfo_A4_FA
    
    Set rangecells = sourceSheet.Range("FA_A4_ZipCode").Cells
    ReDim ZipCodeA4_FA(end_A4_FA)
    For i = 1 To end_A4_FA
        ZipCodeA4_FA(i) = rangecells.item(i).value
        If Len(ZipCodeA4_FA(i)) = 0 Then
            'Exit For
        End If
'
'        If Not chkCompulsory(ZipCodeA4_FA(i)) Then
'            msgbox_A4_FA ("* Please enter the ZipCode at Sr. No  " & i & "If  ZIP code is not available, then enter XXXXXX in Table A4 of Schedule FA")
'            ValidateZipCode_A4_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(ZipCodeA4_FA(i)) Then
            msgbox_A4_FA ("* ZipCode at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Table A4 of Schedule FA")
            ValidateZipCode_A4_FA = False
            Exit Function
        End If
    Next
    End Function
    
Function ValidateAccountOpeningDate_A4()
Dim rangecells As Range
Dim i As Long
    
ValidateAccountOpeningDate_A4 = True
setTblinfo_A4_FA
Set rangecells = sourceSheet.Range("FA_A4_AccOpeningDate").Cells
ReDim AccountOpeningDateA4_FA(end_A4_FA)
For i = 1 To end_A4_FA

AccountOpeningDateA4_FA(i) = rangecells.item(i).value
'
'If Not chkCompulsory(AccountOpeningDateA4_FA(i)) Then
' msgbox_A4_FA ("* Please enter the Date of contract at Sr. No  " & i & "  in Table A4 of Schedule FA")
' ValidateAccountOpeningDate_A4 = False
' Exit Function
'End If

If Not CheckDateddmmyyyy(AccountOpeningDateA4_FA(i)) Then
 msgbox_A4_FA ("* Date of contract at Sr. No  " & i & " is invalid in Table A4 of Schedule FA")
 ValidateAccountOpeningDate_A4 = False
 Exit Function
End If

Next
    
End Function

Function ValidateTotalGrossAmount_A4_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateTotalGrossAmount_A4_FA = True
    setTblinfo_A4_FA
    
    Set rangecells = sourceSheet.Range("FA_A4_Total_grossamount").Cells
    ReDim TotalGrossAmount_A4_FA(end_A4_FA)
    For i = 1 To end_A4_FA
        TotalGrossAmount_A4_FA(i) = rangecells.item(i).value
        
        
'        If Not chkCompulsory(TotalGrossAmount_A4_FA(i)) Then
''            msgbox_A4_FA ("* Please enter the Total gross amount paid/credited at Sr. No  " & i & " in Table A4 of Schedule FA")
'
'          'SIT-69129 Changed by sai
'            msgbox_A4_FA ("* Please enter total gross amount paid/credited with respect to the contract during the period in Table A4 of Schedule FA at Sr. No " & i & "")
'
'
'            ValidateTotalGrossAmount_A4_FA = False
'            Exit Function
'        End If
       
        If Not checkfieldspecialcharacter(TotalGrossAmount_A4_FA(i)) Then
            msgbox_A4_FA ("* Total gross amount paid/credited at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A4 of Schedule FA")
            ValidateTotalGrossAmount_A4_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateTheCashValue_A4_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateTheCashValue_A4_FA = True
    setTblinfo_A4_FA
    
    Set rangecells = sourceSheet.Range("FA_A4_the_cashvalue").Cells
    ReDim TheCashValue_A4_FA(end_A4_FA)
    For i = 1 To end_A4_FA
        TheCashValue_A4_FA(i) = rangecells.item(i).value
        
        
'        If Not chkCompulsory(TheCashValue_A4_FA(i)) Then
''            msgbox_A4_FA ("* Please enter The cash value or surrender value at Sr. No  " & i & " in Table A4 of Schedule FA")
'        'SIT-69130 Changed by sai
'            msgbox_A4_FA ("* Please enter the cash value or surrender value of the contract in Table A4 of Schedule FA at Sr. No " & i & "")
'
'
'            ValidateTheCashValue_A4_FA = False
'            Exit Function
'        End If
       
        If Not checkfieldspecialcharacter(TheCashValue_A4_FA(i)) Then
            msgbox_A4_FA ("* The cash value or surrender value at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table A4 of Schedule FA")
            ValidateTheCashValue_A4_FA = False
            Exit Function
        End If
    Next
End Function



Function ValidatesheetBFA() As Boolean
    ValidatesheetBFA = True
    Dim i As Long
    If Not ValidateCountry_B_FA() Then ValidatesheetBFA = False
    'If Not isdropdownblank(sourceSheet.Range("FA_B_Country").item(1).value) And (Len(sourceSheet.Range("FA_B_Country").item(1).value) > 0) Then
        If Not ValidateZipCode_B_FA() Then ValidatesheetBFA = False
        If Not ValidateEntityNature_B_FA() Then ValidatesheetBFA = False
        If Not ValidateEntityName_B_FA() Then ValidatesheetBFA = False
        If Not ValidateEntityAdd_B_FA() Then ValidatesheetBFA = False
        If Not ValidateTotalInv_B_FA() Then ValidatesheetBFA = False
        If Not ValidateNatureOfInterest() Then ValidatesheetBFA = False
        If Not ValidateDateHeld() Then ValidatesheetBFA = False
        If Not Validateincomeaccured() Then ValidatesheetBFA = False
        If Not ValidateNatureOfIncome() Then ValidatesheetBFA = False
        If Not ValidateAmount_B_FA() Then ValidatesheetBFA = False
        If Not ValidateScheduleName_B_FA() Then ValidatesheetBFA = False
        If Not ValidateScheduleItemNum_B_FA() Then ValidatesheetBFA = False
        If Not ValidateInvTot_B_FA() Then ValidatesheetBFA = False
    
    'End If
    
    setTblinfo_B_FA2
'    If (end_B_FA <> end_B_FA2) Then
'        msgbox_B_FA ("* Details of Country  in Table B of Schedule FA are Mandatory")
'        ValidatesheetBFA = False
'        Exit Function
'    End If
    
    'Added by Aavula 16/07/2024

    For i = 1 To Sheet27.Range("FA_B_Country").count

        If Sheet27.Range("FA_B_Country").Cells(i, 1).value <> "" And Sheet27.Range("FA_B_Country").Cells(i, 1).value <> "(Select)" Or Sheet27.Range("FA_B_ZipCode").Cells(i, 1).value <> "" Or Sheet27.Range("FA_B_EntityNature").Cells(i, 1).value <> "" Or Sheet27.Range("FA_B_EntityName").Cells(i, 1).value <> "" Or _
        (Sheet27.Range("FA_B_EntityAdd").Cells(i, 1).value <> "" Or Sheet27.Range("FA_B_NatureOfInterest").Cells(i, 1).value <> "" And Sheet27.Range("FA_B_NatureOfInterest").Cells(i, 1).value <> "(Select)") Or _
        Sheet27.Range("FA_B_DateSinceHeld").Cells(i, 1).value <> "" Or Sheet27.Range("FA_B_TotalInv").Cells(i, 1).value <> "" Or Sheet27.Range("FA_B_IncomeInterest").Cells(i, 1).value <> "" Or Sheet27.Range("FA_B_NatureOfIncome").Cells(i, 1).value <> "" Or Sheet27.Range("FA_B_Amount").Cells(i, 1).value <> "" Or Sheet27.Range("FA_B_ScheduleOffered").Cells(i, 1).value <> "" And Sheet27.Range("FA_B_ScheduleOffered").Cells(i, 1).value <> "(Select)" Or Sheet27.Range("FA_B_ItemNumOfSchedule").Cells(i, 1).value <> "" Then
   
        
        If Sheet27.Range("FA_B_Country").Cells(i, 1).value = "" Or Sheet27.Range("FA_B_Country").Cells(i, 1).value = "(Select)" Then
            ValidatesheetBFA = False
            msgbox_B_FA ("Please select the Country Name and Code in Table B of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_B_ZipCode").Cells(i, 1).value = "" Then
            ValidatesheetBFA = False
            msgbox_B_FA ("Please enter ZIP code, if  ZIP code is not available, then enter XXXXXX in Table B of Schedule FA at Sr. No " & i & ".")
        End If
        
        
        'SIT-69125 Changed by sai
        
        If Sheet27.Range("FA_B_EntityNature").Cells(i, 1).value = "" Then
            ValidatesheetBFA = False
             msgbox_B_FA ("*Please enter the Nature of entity in Table B of Schedule FA at Sr. No " & i & "")
        End If
        ''''''''''''''
        
        If Sheet27.Range("FA_B_EntityName").Cells(i, 1).value = "" Then
            ValidatesheetBFA = False
            msgbox_B_FA ("Please enter the name  of the entity in Table B of Schedule FA at Sr. No " & i & ".")
        End If


        If Sheet27.Range("FA_B_EntityAdd").Cells(i, 1).value = "" Then
            ValidatesheetBFA = False
            msgbox_B_FA ("Please enter the Address of the entity in Table B of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_B_NatureOfInterest").Cells(i, 1).value = "" Then
            ValidatesheetBFA = False
            msgbox_B_FA ("Please provide the nature of interest in Table B of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_B_DateSinceHeld").Cells(i, 1).value = "" Then
            ValidatesheetBFA = False
            msgbox_B_FA ("Please enter the date since financial interest in entity is held in Table B of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_B_TotalInv").Cells(i, 1).value = "" Then
            ValidatesheetBFA = False
            msgbox_B_FA ("Please enter the Total Investment in Table B of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_B_IncomeInterest").Cells(i, 1).value = "" Then
            ValidatesheetBFA = False
            msgbox_B_FA ("Please enter the income accrued from such financial interest in Table B of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_B_NatureOfIncome").Cells(i, 1).value = "" Then
            ValidatesheetBFA = False
            msgbox_B_FA ("Please enter the nature of the income in Table B of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_B_Amount").Cells(i, 1).value = "" Then
            ValidatesheetBFA = False
            msgbox_B_FA ("Please enter the amount of the income taxable in Table B of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_B_ScheduleOffered").Cells(i, 1).value = "" Then
            ValidatesheetBFA = False
            msgbox_B_FA ("Please enter the Schedule where income was offered in Table B of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_B_ItemNumOfSchedule").Cells(i, 1).value = "" Then
             ValidatesheetBFA = False
             msgbox_B_FA ("Please enter the Item number of schedule in Table B of Schedule FA at Sr. No " & i & ".")
        End If
    End If
Next i
 ''''''''''''
'    If (CountB1 <> CountB11) Then
'        msgbox_B_FA ("* Enter all mandatory field  in Table B of Schedule FA")
'        ValidatesheetBFA = False
'        Exit Function
'    End If
End Function

Function ValidateCountry_B_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim CountryName As String
    
    ValidateCountry_B_FA = True
    setTblinfo_B_FA
    
    Set rangecells = sourceSheet.Range("FA_B_Country").Cells
    ReDim CountryCodeB_FA(end_B_FA)
    ReDim CountryNameB_FA(end_B_FA)
    noOfProcessSub = end_B_FA
    For i = 1 To end_B_FA
        CountryName = rangecells.item(i).value
        
'     If isdropdownblank(CountryName) Then
'            msgbox_B_FA ("* Please select the CountryName and Code at Sr. No  " & i & " in Table B of Schedule FA")
'            ValidateCountry_B_FA = False
'            Exit Function
'        End If
        
        If CountryName <> "" Then
        CountryNameB_FA(i) = Mid(CountryName, WorksheetFunction.Search("-", CountryName) + 1, Len(CountryName))
        CountryCodeB_FA(i) = Mid(CountryName, 1, WorksheetFunction.Search("-", CountryName) - 1)
        End If
        
'        If Len(CountryCodeB_FA(i)) = 0 Then
'            Exit For
'        End If
        
'        If Not chkCompulsory(CountryCodeB_FA(i)) Then
'            msgbox_B_FA ("* CountryName and Code at Sr. No  " & i & "  in Table B of Schedule FA")
'            ValidateCountry_B_FA = False
'            Exit Function
'        End If
        
'        If Not chkCompulsory(CountryCodeB_FA(i)) Then
'            msgbox_B_FA ("* CountryName at Sr. No  " & i & "  in Sheet TR_FA  is Mandatory")
'            ValidateCountry_B_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(CountryCodeB_FA(i)) Then
            msgbox_B_FA ("* CountryName and Code at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed  in Table B of Schedule FA")
            ValidateCountry_B_FA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function

Function ValidateZipCode_B_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateZipCode_B_FA = True
    setTblinfo_B_FA
    
    Set rangecells = sourceSheet.Range("FA_B_ZipCode").Cells
    ReDim ZipCodeB_FA(end_B_FA)
    For i = 1 To end_B_FA
        ZipCodeB_FA(i) = rangecells.item(i).value
        If Len(ZipCodeB_FA(i)) = 0 Then
            'Exit For
        End If
        
'        If Not chkCompulsory(ZipCodeB_FA(i)) Then
'            msgbox_B_FA ("* Please enter the ZipCode at Sr. No  " & i & " If  ZIP code is not available, then enter XXXXXX  in Table B of Schedule FA")
'            ValidateZipCode_B_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(ZipCodeB_FA(i)) Then
            msgbox_B_FA ("* ZipCode at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed  in Table B of Schedule FA")
            ValidateZipCode_B_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateEntityNature_B_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateEntityNature_B_FA = True
    setTblinfo_B_FA
    
    Set rangecells = sourceSheet.Range("FA_B_EntityNature").Cells
    ReDim EntityNatureB_FA(end_B_FA)
    For i = 1 To end_B_FA
        EntityNatureB_FA(i) = rangecells.item(i).value
        If Len(EntityNatureB_FA(i)) = 0 Then
            'Exit For
        End If
        
'        If Not chkCompulsory(EntityNatureB_FA(i)) Then
''            msgbox_B_FA ("* Please enter the Entity Nature at Sr. No  " & i & "   in Table B of Schedule FA")
'
'       'SIT-69125 Changed by sai
'            msgbox_B_FA ("*Please enter the Nature of entity in Table B of Schedule FA at Sr. No " & i & "")
'
'            ValidateEntityNature_B_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(EntityNatureB_FA(i)) Then
            msgbox_B_FA ("* Entity Nature at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed  in Table B of Schedule FA")
            ValidateEntityNature_B_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateEntityName_B_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateEntityName_B_FA = True
    setTblinfo_B_FA
    
    Set rangecells = sourceSheet.Range("FA_B_EntityName").Cells
    ReDim EntityNameB_FA(end_B_FA)
    For i = 1 To end_B_FA
        EntityNameB_FA(i) = rangecells.item(i).value
        If Len(EntityNameB_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(EntityNameB_FA(i)) Then
''            msgbox_B_FA ("* Please enter the Entity Name at Sr. No  " & i & "   in Table B of Schedule FA")
'
'          'SIT-69124 changed by sai
'            msgbox_B_FA ("* Please enter the name of the entity in Table B of Schedule FA at Sr. No " & i & "")
'
'
'            ValidateEntityName_B_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(EntityNameB_FA(i)) Then
            msgbox_B_FA ("* Entity Name at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed  in Table B of Schedule FA")
            ValidateEntityName_B_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateEntityAdd_B_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateEntityAdd_B_FA = True
    setTblinfo_B_FA
    
    Set rangecells = sourceSheet.Range("FA_B_EntityAdd").Cells
    ReDim EntityAddB_FA(end_B_FA)
    For i = 1 To end_B_FA
        EntityAddB_FA(i) = rangecells.item(i).value
        If Len(EntityAddB_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(EntityAddB_FA(i)) Then
''            msgbox_B_FA ("* Please enter the Entity Address at Sr. No  " & i & "  in Table B of Schedule FA")
'
'       'SIT-69122 Changed by sai
'            msgbox_B_FA ("* Please enter the Address of the entity in Table B of Schedule FA at Sr. No " & i & "")
'
'            ValidateEntityAdd_B_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(EntityAddB_FA(i)) Then
            msgbox_B_FA ("* Entity Address at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed  in Table B of Schedule FA")
            ValidateEntityAdd_B_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotalInv_B_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateTotalInv_B_FA = True
    setTblinfo_B_FA

    Set rangecells = sourceSheet.Range("FA_B_TotalInv").Cells
    ReDim TotalInvB_FA(end_B_FA)
    For i = 1 To end_B_FA
        TotalInvB_FA(i) = rangecells.item(i).value
        If Len(TotalInvB_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(TotalInvB_FA(i)) Then
'            msgbox_B_FA ("* Please enter the Total Investment at Sr. No  " & i & "  in Table B of Schedule FA")
'            ValidateTotalInv_B_FA = False
'            Exit Function
'        End If
        
        If (TotalInvB_FA(i) < 0) Then
            msgbox_B_FA ("* Total Investment at Sr. No  " & i & "  should be Non negative, no decimal, upto 99,999,999,999,999  in Table B of Schedule FA")
            ValidateTotalInv_B_FA = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(TotalInvB_FA(i)) Then
            msgbox_B_FA ("* Total Investment at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed  in Table B of Schedule FA")
            ValidateTotalInv_B_FA = False
            Exit Function
        End If
    Next
End Function
Function ValidateNatureOfInterest() As Boolean

Dim rangecells As Range
Dim i As Long
ValidateNatureOfInterest = True
setTblinfo_B_FA
Set rangecells = sourceSheet.Range("FA_B_NatureOfInterest").Cells
ReDim NatureOfInterestB_FA(end_B_FA)

For i = 1 To end_B_FA
NatureOfInterestB_FA(i) = rangecells.item(i).value
' If isdropdownblank(NatureOfInterestB_FA(i)) Then
'            msgbox_B_FA ("* Please select the Nature of Interest at Sr. No  " & i & "  in Table B of Schedule FA")
'            ValidateNatureOfInterest = False
'            Exit Function
'End If

Next


End Function

Function ValidateDateHeld() As Boolean
Dim rangecells As Range
Dim i As Long
ValidateDateHeld = True
setTblinfo_B_FA

Set rangecells = sourceSheet.Range("FA_B_DateSinceHeld").Cells
ReDim DateHeldB_FA(end_B_FA)
For i = 1 To end_B_FA
DateHeldB_FA(i) = rangecells.item(i).value
'If Not chkCompulsory(DateHeldB_FA(i)) Then
''        msgbox_B_FA ("* Please enter the Date Held at Sr. No  " & i & "  in Table B of Schedule FA")
'       'SIT-69121 Changed by sai
'        msgbox_B_FA ("* Please enter the date since financial interest in entity is held in Table B of Schedule FA at Sr. No " & i & "")
'
'
'        ValidateDateHeld = False
'        Exit Function
'    End If
    
    If Not CheckDateddmmyyyy(DateHeldB_FA(i)) Then
    msgbox_B_FA ("* Date Held at Sr. No  " & i & " is invalid  in Table B of Schedule FA")
        ValidateDateHeld = False
        Exit Function
    End If
Next


End Function

Function Validateincomeaccured() As Boolean
Dim rangecells As Range
Dim i As Long

Validateincomeaccured = True
setTblinfo_B_FA

Set rangecells = sourceSheet.Range("FA_B_IncomeInterest").Cells
ReDim IncomeIncurredInAccB_FA(end_B_FA)

For i = 1 To end_B_FA
IncomeIncurredInAccB_FA(i) = rangecells.item(i).value
'    If Not chkCompulsory(IncomeIncurredInAccB_FA(i)) Then
''        msgbox_B_FA ("* Please enter the Income Accured at Sr. No  " & i & "   in Table B of Schedule FA")
'        'SIT-69118 Changed by sai
'        msgbox_B_FA ("* Please enter the income accrued from such financial interest in Table B of Schedule FA. at Sr. No " & i & "")
'
'
'        Validateincomeaccured = False
'        Exit Function
'    End If
Next

End Function

Function ValidateNatureOfIncome() As Boolean
Dim rangecells As Range
Dim i As Long

ValidateNatureOfIncome = True
setTblinfo_B_FA
Set rangecells = sourceSheet.Range("FA_B_NatureOfIncome").Cells
ReDim NatureOfIncomeB_FA(end_B_FA)

For i = 1 To end_B_FA
NatureOfIncomeB_FA(i) = rangecells.item(i).value
'If Not chkCompulsory(NatureOfIncomeB_FA(i)) Then
'        msgbox_B_FA ("* Please enter the Nature Of Income at Sr. No  " & i & "  in Table B of Schedule FA")
'        ValidateNatureOfIncome = False
'        Exit Function
'End If
 
If Not checkfieldspecialcharacter(NatureOfIncomeB_FA(i)) Then
        msgbox_B_FA ("* Nature Of Income at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed  in Table B of Schedule FA")
        ValidateNatureOfIncome = False
        Exit Function
End If
    
Next
End Function

Function ValidateAmount_B_FA() As Boolean
Dim rangecells As Range
Dim i As Long

ValidateAmount_B_FA = True
setTblinfo_B_FA
Set rangecells = sourceSheet.Range("FA_B_Amount").Cells
ReDim AmountB_FA(end_B_FA)
For i = 1 To end_B_FA
AmountB_FA(i) = rangecells.item(i).value
'If Not chkCompulsory(AmountB_FA(i)) Then
'       ' msgbox_B_FA ("* Please enter the Amount at Sr. No  " & i & "  in Table B of Schedule FA")
'
'        'Added by Aavula SIT-69113
'          msgbox_B_FA ("* Please enter the amount of the income taxable at Sr. No  " & i & " in Table B of Schedule FA")
'
'        ValidateAmount_B_FA = False
'        Exit Function
'End If

Next
End Function

Function ValidateScheduleName_B_FA() As Boolean
Dim rangecells As Range
Dim i As Long

ValidateScheduleName_B_FA = True
setTblinfo_B_FA
Set rangecells = sourceSheet.Range("FA_B_ScheduleOffered").Cells
ReDim Schedule_B_FA(end_B_FA)
For i = 1 To end_B_FA
Schedule_B_FA(i) = rangecells.item(i).value

'If isdropdownblank(Schedule_B_FA(i)) Then
''        msgbox_B_FA ("* Please enter the Schedule where offered at Sr. No  " & i & "  in Table B of Schedule FA")
'     'SIT-69133 Changed by sai
'        msgbox_B_FA ("* Please enter the Schedule where income was offered in Table B of Schedule FA at Sr. No  " & i & "")
'
'
'        ValidateScheduleName_B_FA = False
'        Exit Function
'End If

'If Not checkfieldspecialcharacter(Schedule_B_FA(i)) Then
'        msgbox_B_FA ("* Schedule where offered at Sr. No  " & i & " in Sheet TR_FA  characters < > & ' " & Chr(34) & " are not allowed")
'        ValidateScheduleName_B_FA = False
'        Exit Function
'End If

Next

End Function

Function ValidateScheduleItemNum_B_FA() As Boolean
Dim rangecells As Range
Dim i As Long

ValidateScheduleItemNum_B_FA = True
setTblinfo_B_FA
Set rangecells = sourceSheet.Range("FA_B_ItemNumOfSchedule").Cells
ReDim ItemNumber_B_FA(end_B_FA)
For i = 1 To end_B_FA
ItemNumber_B_FA(i) = rangecells.item(i).value

'If Not chkCompulsory(ItemNumber_B_FA(i)) Then
'        msgbox_B_FA ("* Please enter the Item number of schedule at Sr. No  " & i & "   in Table B of Schedule FA")
'        ValidateScheduleItemNum_B_FA = False
'        Exit Function
'End If

If Not checkfieldspecialcharacter(ItemNumber_B_FA(i)) Then
        msgbox_B_FA ("* Item number of schedule at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed  in Table B of Schedule FA")
        ValidateScheduleItemNum_B_FA = False
        Exit Function
End If

Next

End Function
Function ValidateInvTot_B_FA() As Boolean
    ValidateInvTot_B_FA = True
    InvTotalB_FA = sourceSheet.Range("FA_B_InvTotal").value
End Function

Function ValidatesheetCFA() As Boolean
    ValidatesheetCFA = True
    Dim i As Long
    If Not ValidateCountry_C_FA() Then ValidatesheetCFA = False
'Konda-------SIT-95679---28/08/2025----
    If Not isdropdownblank(sourceSheet.Range("FA_C_Country").item(1).value) And (Len(sourceSheet.Range("FA_C_Country").item(1).value) > 0) Then
'----------------SIT-95679---28/08/2025----
        If Not ValidateZipCode_C_FA() Then ValidatesheetCFA = False
        If Not ValidatePropertyAdd_C_FA() Then ValidatesheetCFA = False
        If Not validateOwnershipC_FA() Then ValidatesheetCFA = False
        If Not ValidateDateOfAcquisitionC_FA() Then ValidatesheetCFA = False
        If Not ValidateIncomeFromPropC_FA() Then ValidatesheetCFA = False
        If Not validateNatureOfIncomeC_FA() Then ValidatesheetCFA = False
        If Not ValidateAmount_C_FA() Then ValidatesheetCFA = False
        If Not ValidateScheduleOffered_C_FA() Then ValidatesheetCFA = False
        If Not VAlidateItemNumOfScheduleC_FA() Then ValidatesheetCFA = False
        If Not ValidateTotalInv_C_FA() Then ValidatesheetCFA = False
        If Not ValidateInvTotal_C_FA() Then ValidatesheetCFA = False
    End If
    
    setTblinfo_C_FA2
'    If (end_C_FA <> end_C_FA2) Then
'        msgbox_C_FA ("* Details of Country  in Table C of Schedule FA are Mandatory")
'        ValidatesheetCFA = False
'        Exit Function
'    End If
    
    'Added by Aavula 16/07/2024

    For i = 1 To Sheet27.Range("FA_C_Country").count

    If Sheet27.Range("FA_C_Country").Cells(i, 1).value <> "" And Sheet27.Range("FA_C_Country").Cells(i, 1).value <> "(Select)" Or Sheet27.Range("FA_C_ZipCode").Cells(i, 1).value <> "" Or _
        (Sheet27.Range("FA_C_PropertyAdd").Cells(i, 1).value <> "" Or Sheet27.Range("FA_C_OwnershipBeneficiary").Cells(i, 1).value <> "" And Sheet27.Range("FA_C_OwnershipBeneficiary").Cells(i, 1).value <> "(Select)") Or _
        Sheet27.Range("FA_C_DateOfAcquisition").Cells(i, 1).value <> "" Or Sheet27.Range("FA_C_TotalInv").Cells(i, 1).value <> "" Or Sheet27.Range("FA_C_IncomeFromProperty").Cells(i, 1).value <> "" Or Sheet27.Range("FA_C_NatureOfIncome").Cells(i, 1).value <> "" Or Sheet27.Range("FA_C_Amount").Cells(i, 1).value <> "" Or Sheet27.Range("FA_C_ScheduleOffered").Cells(i, 1).value <> "" And Sheet27.Range("FA_C_ScheduleOffered").Cells(i, 1).value <> "(Select)" Or Sheet27.Range("FA_C_ItemNumOfSchedule").Cells(i, 1).value <> "" Then
   
        
        If Sheet27.Range("FA_C_Country").Cells(i, 1).value = "" Or Sheet27.Range("FA_C_Country").Cells(i, 1).value = "(Select)" Then
            ValidatesheetCFA = False
            msgbox_C_FA ("Please select the Country Name and Code in Table C of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_C_ZipCode").Cells(i, 1).value = "" Then
            ValidatesheetCFA = False
            msgbox_C_FA ("Please enter ZIP code, if  ZIP code is not available, then enter XXXXXX in Table C of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_C_PropertyAdd").Cells(i, 1).value = "" Then
            ValidatesheetCFA = False
            msgbox_C_FA ("Please enter the Address of the Property in Table C of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_C_OwnershipBeneficiary").Cells(i, 1).value = "" Or Sheet27.Range("FA_C_OwnershipBeneficiary").Cells(i, 1).value = "(Select)" Then
            ValidatesheetCFA = False
            msgbox_C_FA ("Please provide the details of ownership in Table C of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_C_DateOfAcquisition").Cells(i, 1).value = "" Then
            ValidatesheetCFA = False
            msgbox_C_FA ("Please enter the date since financial interest in entity held in Table C of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_C_TotalInv").Cells(i, 1).value = "" Then
            ValidatesheetCFA = False
            msgbox_C_FA ("Please enter the Total Investment amount in Table C of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_C_IncomeFromProperty").Cells(i, 1).value = "" Then
            ValidatesheetCFA = False
            msgbox_C_FA ("Please enter the Income derived from the property in Table C of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_C_NatureOfIncome").Cells(i, 1).value = "" Then
            ValidatesheetCFA = False
            msgbox_C_FA ("Please enter the nature of the income in Table C of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_C_Amount").Cells(i, 1).value = "" Then
            ValidatesheetCFA = False
            msgbox_C_FA ("Please enter the amount of the income taxable in Table C of Schedule FA at Sr. No " & i & ".")
        End If
        
        'SIT-69133
        If Sheet27.Range("FA_C_ScheduleOffered").Cells(i, 1).value = "" Or Sheet27.Range("FA_C_ScheduleOffered").Cells(i, 1).value = "(Select)" Then
            ValidatesheetCFA = False
            msgbox_C_FA ("Please enter the Schedule where income was offered in Table C of Schedule FA at Sr. No " & i & ".")
        End If
            If Sheet27.Range("FA_C_ItemNumOfSchedule").Cells(i, 1).value = "" Then
             ValidatesheetCFA = False
             msgbox_C_FA ("Please enter the Item number of schedule in Table C of Schedule FA at Sr. No " & i & ".")
        End If
    End If
Next i
     ''''''''''
    
    'End If
    
'    If (CountC1 <> CountC11) Then
'        msgbox_C_FA ("* Enter all mandatory field in Table C of Schedule FA")
'        ValidatesheetCFA = False
'        Exit Function
'    End If
End Function

Function ValidateCountry_C_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim CountryName As String
    
    ValidateCountry_C_FA = True
    setTblinfo_C_FA
    
    Set rangecells = sourceSheet.Range("FA_C_Country").Cells
    'Malli-----17/11/2025-------------IPIP-66714
   ' CountryName = rangecells.item(i).value
    '-------------------------
    ReDim CountryCodeC_FA(end_C_FA)
    ReDim CountryNameC_FA(end_C_FA)
    noOfProcessSub = end_C_FA
    For i = 1 To end_C_FA
        
        CountryName = rangecells.item(i).value
'        If isdropdownblank(CountryName) Then
'            msgbox_C_FA ("* Please select the Country Name and Code at Sr. No  " & i & " in Table C of Schedule FA")
'            ValidateCountry_C_FA = False
'            Exit Function
'        End If
        
        If CountryName <> "" Then
        ' 'Malli & Sudhakar----17/11/2025-------------IPIP-66714
'        CountryNameC_FA(i) = Mid(CountryName, WorksheetFunction.Search("-", CountryName) + 1, Len(CountryName))
'        CountryCodeC_FA(i) = Mid(CountryName, 1, WorksheetFunction.Search("-", CountryName) - 1)
        
         Dim CN As Long
        CN = InStr(1, CountryName, "-", vbTextCompare)
        
        If CN > 0 Then
        CountryNameC_FA(i) = Trim(Mid(CountryName, CN + 1))
        CountryCodeC_FA(i) = Mid(CountryName, 1, CN - 1)
        End If
        '-----------------------------
        End If
        
'        If Not chkCompulsory(CountryCodeC_FA(i)) Then
'            msgbox_C_FA ("* CountryName and Code at Sr. No  " & i & " in Table C of Schedule FA")
'            ValidateCountry_C_FA = False
'            Exit Function
'        End If
'
'        If Not chkCompulsory(CountryCodeC_FA(i)) Then
'            msgbox_C_FA ("* CountryName at Sr. No  " & i & "  in Sheet TR_FA  is Mandatory")
'            ValidateCountry_C_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(CountryCodeC_FA(i)) Then
            msgbox_C_FA ("* CountryName and Code at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table C of Schedule FA")
            ValidateCountry_C_FA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function

Function ValidateZipCode_C_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateZipCode_C_FA = True
    setTblinfo_C_FA
    
    Set rangecells = sourceSheet.Range("FA_C_ZipCode").Cells
    ReDim ZipCodeC_FA(end_C_FA)
    For i = 1 To end_C_FA
        ZipCodeC_FA(i) = rangecells.item(i).value
        If Len(ZipCodeC_FA(i)) = 0 Then
            'Exit For
        End If
        
'        If Not chkCompulsory(ZipCodeC_FA(i)) Then
''            msgbox_C_FA ("* Please enter the ZipCode at Sr. No  " & i & "  is Mandatory.If  ZIP code is not available, then enter XXXXXX in Table C of Schedule FA")
'          'SIT-69128 Changed by sai
'            msgbox_C_FA ("* Please enter ZIP code, if  ZIP code is not available, then enter XXXXXX in Table C of Schedule FA at Sr. No  " & i & "")
'
'            ValidateZipCode_C_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(ZipCodeC_FA(i)) Then
            msgbox_C_FA ("* ZipCode at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table C of Schedule FA")
            ValidateZipCode_C_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidatePropertyAdd_C_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidatePropertyAdd_C_FA = True
    setTblinfo_C_FA
    
    Set rangecells = sourceSheet.Range("FA_C_PropertyAdd").Cells
    ReDim PropertyAddC_FA(end_C_FA)
    For i = 1 To end_C_FA
        PropertyAddC_FA(i) = rangecells.item(i).value
        If Len(PropertyAddC_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(PropertyAddC_FA(i)) Then
''            msgbox_C_FA ("* Please enter the Property Address at Sr. No  " & i & "  in Table C of Schedule FA")
'         'SIT-69127 Changed by sai
'            msgbox_C_FA ("* Please enter the Address of the Property in Table C of Schedule FA at Sr. No " & i & "")
'
'            ValidatePropertyAdd_C_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(PropertyAddC_FA(i)) Then
            msgbox_C_FA ("* Property Address at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Table C of Schedule FA")
            ValidatePropertyAdd_C_FA = False
            Exit Function
        End If
    Next
End Function

Function validateOwnershipC_FA() As Boolean
 Dim rangecells As Range
    Dim i As Long
 
    validateOwnershipC_FA = True
    setTblinfo_C_FA
    
    Set rangecells = sourceSheet.Range("FA_C_OwnershipBeneficiary").Cells
    ReDim OwnershipC_FA(end_C_FA)
    For i = 1 To end_C_FA
        OwnershipC_FA(i) = rangecells.item(i).value
                
'        If isdropdownblank(OwnershipC_FA(i)) Then
''            msgbox_C_FA ("* Please select the ownership at Sr. No  " & i & " in Table C of Schedule FA")
'
'         'SIT-69123 Changed By sai
'            msgbox_C_FA ("* Please provide the details of ownership in Table C of Schedule FA"" at Sr. No  " & i & "")
'
'            validateOwnershipC_FA = False
'            Exit Function
'        End If
        Next
End Function

Function ValidateDateOfAcquisitionC_FA() As Boolean
 Dim rangecells As Range
    Dim i As Long
 
    ValidateDateOfAcquisitionC_FA = True
    setTblinfo_C_FA
    
    Set rangecells = sourceSheet.Range("FA_C_DateOfAcquisition").Cells
    ReDim DateOfAcquistionC_FA(end_C_FA)
    For i = 1 To end_C_FA
        DateOfAcquistionC_FA(i) = rangecells.item(i).value
'        If Not chkCompulsory(DateOfAcquistionC_FA(i)) Then
''            msgbox_C_FA ("* Please enter the Date Of Acquistion at Sr. No  " & i & " in Table C of Schedule FA")
'
'          'SIT-69120  Changed by sai
'            msgbox_C_FA (" * Please enter the date since financial interest in entity held in Table C of Schedule FA. at Sr. No " & i & "")
'
'            ValidateDateOfAcquisitionC_FA = False
'            Exit Function
'        End If
        
        If Not CheckDateddmmyyyy(DateOfAcquistionC_FA(i)) Then
            msgbox_C_FA ("* Date of Acquistion at Sr. No  " & i & "is Invalid in Table C of Schedule FA")
            ValidateDateOfAcquisitionC_FA = False
            Exit Function
        End If
        Next
End Function

Function ValidateIncomeFromPropC_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateIncomeFromPropC_FA = True
    setTblinfo_C_FA
    Set rangecells = sourceSheet.Range("FA_C_IncomeFromProperty").Cells
    ReDim IncomeFromPropC_FA(end_C_FA)
    For i = 1 To end_C_FA
        IncomeFromPropC_FA(i) = rangecells.item(i).value
'        If Not chkCompulsory(IncomeFromPropC_FA(i)) Then
'            msgbox_C_FA ("* Please enter the Income Derived from property at Sr. No  " & i & "  in Table C of Schedule FA")
'            ValidateIncomeFromPropC_FA = False
'            Exit Function
'        End If
        
        
        Next
    
End Function

Function validateNatureOfIncomeC_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    validateNatureOfIncomeC_FA = True
    setTblinfo_C_FA
    Set rangecells = sourceSheet.Range("FA_C_NatureOfIncome").Cells
    ReDim NatureOfIncomeC_FA(end_C_FA)
    For i = 1 To end_C_FA
        NatureOfIncomeC_FA(i) = rangecells.item(i).value
'        If Not chkCompulsory(NatureOfIncomeC_FA(i)) Then
'            msgbox_C_FA ("* Please enter the Nature Of Income at Sr. No  " & i & " in Table C of Schedule FA")
'            validateNatureOfIncomeC_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(NatureOfIncomeC_FA(i)) Then
        msgbox_B_FA ("* Nature Of Income at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table C of Schedule FA")
        validateNatureOfIncomeC_FA = False
        Exit Function
        End If
        
        Next
    
End Function

Function ValidateAmount_C_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateAmount_C_FA = True
    setTblinfo_C_FA
    Set rangecells = sourceSheet.Range("FA_C_Amount").Cells
    ReDim AmountC_FA(end_C_FA)
    For i = 1 To end_C_FA
        AmountC_FA(i) = rangecells.item(i).value
'        If Not chkCompulsory(AmountC_FA(i)) Then
'            'msgbox_C_FA ("* Please enter the Amount at Sr. No  " & i & " in Table C of Schedule FA")
'
'            'Added by Aavula
'            msgbox_C_FA ("* Please enter the amount of the income taxable at Sr. No " & i & " in Table C of Schedule FA.")
'
'            ValidateAmount_C_FA = False
'            Exit Function
'        End If
        
        
        Next
End Function

Function ValidateScheduleOffered_C_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateScheduleOffered_C_FA = True
    setTblinfo_C_FA
    Set rangecells = sourceSheet.Range("FA_C_ScheduleOffered").Cells
    ReDim ScheduleOfferedC_FA(end_C_FA)
    For i = 1 To end_C_FA
        ScheduleOfferedC_FA(i) = rangecells.item(i).value
'        If isdropdownblank(ScheduleOfferedC_FA(i)) Then
''            msgbox_C_FA ("* Please enter the Schedule offered at Sr. No  " & i & "  in Table C of Schedule FA")
'
'             'Added by Aavula
'              msgbox_C_FA ("* Please enter the Schedule where income was offered at Sr. No " & i & " in Table C of Schedule FA.")
'
'            ValidateScheduleOffered_C_FA = False
'            Exit Function
'        End If
        
'        If Not checkfieldspecialcharacter(ScheduleOfferedC_FA(i)) Then
'        msgbox_B_FA ("* Schedule offered at Sr. No  " & i & " in Sheet TR_FA  characters < > & ' " & Chr(34) & " are not allowed")
'        ValidateScheduleOffered_C_FA = False
'        Exit Function
'        End If
        
        
        Next
End Function

Function VAlidateItemNumOfScheduleC_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    VAlidateItemNumOfScheduleC_FA = True
    setTblinfo_C_FA
    Set rangecells = sourceSheet.Range("FA_C_ItemNumOfSchedule").Cells
    ReDim ItemNumofScheduleC_FA(end_C_FA)
    For i = 1 To end_C_FA
        ItemNumofScheduleC_FA(i) = rangecells.item(i).value
'        If Not chkCompulsory(ItemNumofScheduleC_FA(i)) Then
'            msgbox_C_FA ("* Please enter the Item No of Schedule at Sr. No  " & i & "in Table C of Schedule FA")
'            VAlidateItemNumOfScheduleC_FA = False
'            Exit Function
'        End If
        
         If Not checkfieldspecialcharacter(ItemNumofScheduleC_FA(i)) Then
        msgbox_B_FA ("* Item No of Schedule at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table C of Schedule FA")
        VAlidateItemNumOfScheduleC_FA = False
        Exit Function
        End If
        
        Next
End Function

Function ValidateTotalInv_C_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateTotalInv_C_FA = True
    setTblinfo_C_FA
    
    Set rangecells = sourceSheet.Range("FA_C_TotalInv").Cells
    ReDim TotalInvC_FA(end_C_FA)
    For i = 1 To end_C_FA
        TotalInvC_FA(i) = rangecells.item(i).value
        If Len(TotalInvC_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(TotalInvC_FA(i)) Then
'            msgbox_C_FA ("* Please enter the Total Investment at Sr. No  " & i & " in Table C of Schedule FA")
'            ValidateTotalInv_C_FA = False
'            Exit Function
'        End If
        
        
        If (TotalInvC_FA(i) < 0) Then
            msgbox_C_FA ("* Please enter the Total Investment at Sr. No  " & i & "  should be Non negative, no decimal, upto 99,999,999,999,999 in Table C of Schedule FA")
            ValidateTotalInv_C_FA = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(TotalInvC_FA(i)) Then
            msgbox_C_FA ("* Total Investment at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table C of Schedule FA")
            ValidateTotalInv_C_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateInvTotal_C_FA() As Boolean
    ValidateInvTotal_C_FA = True
    InvTotalC_FA = sourceSheet.Range("FA_C_InvTotal").value
End Function

Function ValidatesheetDFA() As Boolean
    ValidatesheetDFA = True
    If Not ValidateCountry_D_FA() Then ValidatesheetDFA = False
    If Not isdropdownblank(sourceSheet.Range("FA_D_Country").item(1).value) And (Len(sourceSheet.Range("FA_D_Country").item(1).value) > 0) Then
        If Not ValidateZipCode_D_FA() Then ValidatesheetDFA = False
        If Not ValidateAssetNature_D_FA() Then ValidatesheetDFA = False
        If Not validateownership_D_FA() Then ValidatesheetDFA = False
        If Not validatedateofAcquistion_D_FA() Then ValidatesheetDFA = False
        If Not ValidateincFromProp_D_FA() Then ValidatesheetDFA = False
        If Not ValidateNatureOfIncome_D_FA() Then ValidatesheetDFA = False
        If Not ValidateAmount_D_FA() Then ValidatesheetDFA = False
        If Not ValidateScheduleOffered_D_FA() Then ValidatesheetDFA = False
        If Not ValidateItemNumSchedule_D_FA() Then ValidatesheetDFA = False
        If Not ValidateTotalInv_D_FA() Then ValidatesheetDFA = False
        If Not ValidateInvTotal_D_FA() Then ValidatesheetDFA = False
    End If
    
    setTblinfo_D_FA2
    If (end_D_FA <> end_D_FA2) Then
        msgbox_D_FA ("* Details of Country in Table D of Schedule FA are Mandatory")
        ValidatesheetDFA = False
        Exit Function
    End If
    
    If (CountD1 <> CountD11) Then
        msgbox_D_FA ("* Enter all mandatory field in Table D of Schedule FA")
        ValidatesheetDFA = False
        Exit Function
    End If
End Function


Function ValidateCountry_D_FA() As Boolean
    Dim CountryName As String
    Dim rangecells As Range
    Dim i As Long
    
    ValidateCountry_D_FA = True
    setTblinfo_D_FA
    
    Set rangecells = sourceSheet.Range("FA_D_Country").Cells
    ReDim CountryCodeD_FA(end_D_FA)
    ReDim CountryNameD_FA(end_D_FA)
    noOfProcessSub = end_D_FA
    For i = 1 To end_D_FA
        CountryName = rangecells.item(i).value
        
        If isdropdownblank(CountryName) Then
            msgbox_D_FA ("* Please enter the CountryName and Code at Sr. No  " & i & " in Table D of Schedule FA")
            ValidateCountry_D_FA = False
            Exit Function
        End If
        
        
        If CountryName <> "" Then
       CountryNameD_FA(i) = Mid(CountryName, WorksheetFunction.Search("-", CountryName) + 1, Len(CountryName))
       CountryCodeD_FA(i) = Mid(CountryName, 1, WorksheetFunction.Search("-", CountryName) - 1)
       End If
        
        If Not chkCompulsory(CountryCodeD_FA(i)) Then
            msgbox_D_FA ("* CountryName and Code  at Sr. No  " & i & " in Table D of Schedule FA")
            ValidateCountry_D_FA = False
            Exit Function
        End If
        
'        If Not chkCompulsory(CountryCodeD_FA(i)) Then
'            msgbox_D_FA ("* CountryName at Sr. No  " & i & "  in Sheet TR_FA  is Mandatory")
'            ValidateCountry_D_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(CountryCodeD_FA(i)) Then
            msgbox_D_FA ("* CountryName and Code at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Table D of Schedule FA")
            ValidateCountry_D_FA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function

Function ValidateZipCode_D_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateZipCode_D_FA = True
    setTblinfo_D_FA
    
    Set rangecells = sourceSheet.Range("FA_D_ZipCode").Cells
    ReDim ZipCodeD_FA(end_D_FA)
    For i = 1 To end_D_FA
        ZipCodeD_FA(i) = rangecells.item(i).value
        If Len(ZipCodeD_FA(i)) = 0 Then
            'Exit For
        End If
        
        If Not chkCompulsory(ZipCodeD_FA(i)) Then
'            msgbox_D_FA ("* Please enter the ZipCode at Sr. No  " & i & "   is Mandatory.If  ZIP code is not available, then enter XXXXXX in Table D of Schedule FA")
            msgbox_D_FA ("* Please enter ZIP code, if  ZIP code is not available, then enter XXXXXX at Sr. No " & i & " in Table D of Schedule FA.")
            
            ValidateZipCode_D_FA = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(ZipCodeD_FA(i)) Then
            msgbox_D_FA ("* ZipCode at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table D of Schedule FA")
            ValidateZipCode_D_FA = False
            Exit Function
        End If
    Next
End Function


Function ValidateAssetNature_D_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateAssetNature_D_FA = True
    setTblinfo_D_FA
    
    Set rangecells = sourceSheet.Range("FA_D_AssetNature").Cells
    ReDim AssetNatureD_FA(end_D_FA)
    For i = 1 To end_D_FA
        AssetNatureD_FA(i) = rangecells.item(i).value
        If Len(AssetNatureD_FA(i)) = 0 Then
        'Exit For
        End If
        
        If Not chkCompulsory(AssetNatureD_FA(i)) Then
            msgbox_D_FA ("* Please enter the Nature of Asset at Sr. No " & i & " in Table D of Schedule FA.")
            
            
            ValidateAssetNature_D_FA = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(AssetNatureD_FA(i)) Then
            msgbox_D_FA ("* Asset Nature at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table D of Schedule FA")
            ValidateAssetNature_D_FA = False
            Exit Function
        End If
    Next
End Function

Function validateownership_D_FA() As Boolean
 Dim rangecells As Range
    Dim i As Long
 
    validateownership_D_FA = True
    setTblinfo_D_FA
    
    Set rangecells = sourceSheet.Range("FA_D_OwnershipBeneficiary").Cells
    ReDim OwnershipD_FA(end_D_FA)
    For i = 1 To end_D_FA
        OwnershipD_FA(i) = rangecells.item(i).value
               
        If isdropdownblank(OwnershipD_FA(i)) Then
            msgbox_D_FA ("* Please select the ownership at Sr. No  " & i & " in Table D of Schedule FA")
            validateownership_D_FA = False
            Exit Function
        End If
        Next
End Function

Function validatedateofAcquistion_D_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    validatedateofAcquistion_D_FA = True
    setTblinfo_D_FA
    
    Set rangecells = sourceSheet.Range("FA_D_DateOfAcquisition").Cells
    ReDim DateofAcquistionD_FA(end_D_FA)
    For i = 1 To end_D_FA
        DateofAcquistionD_FA(i) = rangecells.item(i).value
           If Not chkCompulsory(DateofAcquistionD_FA(i)) Then
            msgbox_D_FA ("* Please enter the Date of Acquistion at Sr. No  " & i & " in Table D of Schedule FA")
            validatedateofAcquistion_D_FA = False
            Exit Function
        End If
        If Not CheckDateddmmyyyy(DateofAcquistionD_FA(i)) Then
            msgbox_D_FA ("* Date of Acquistion at Sr. No  " & i & "  in Table D of Schedule FA")
            validatedateofAcquistion_D_FA = False
            Exit Function
        End If
        Next
End Function

Function ValidateincFromProp_D_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateincFromProp_D_FA = True
    setTblinfo_D_FA
    
    Set rangecells = sourceSheet.Range("FA_D_IncomeFromProperty").Cells
    ReDim IncomeFrmAsset_FA(end_D_FA)
    For i = 1 To end_D_FA
        IncomeFrmAsset_FA(i) = rangecells.item(i).value
               
        If Not chkCompulsory(IncomeFrmAsset_FA(i)) Then
            'msgbox_D_FA ("* Please enter the Income From Property at Sr. No  " & i & " in Table D of Schedule FA")
           msgbox_D_FA ("* Please enter the Income derived from the asset at Sr. No " & i & " in Table D of Schedule FA.")

            ValidateincFromProp_D_FA = False
            Exit Function
        End If
        Next
End Function

Function ValidateNatureOfIncome_D_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateNatureOfIncome_D_FA = True
    setTblinfo_D_FA
    
    Set rangecells = sourceSheet.Range("FA_D_NatureOfIncome").Cells
    ReDim NatureOfIncome_FA(end_D_FA)
    For i = 1 To end_D_FA
        NatureOfIncome_FA(i) = rangecells.item(i).value
               
        If Not chkCompulsory(NatureOfIncome_FA(i)) Then
            msgbox_D_FA ("* Please enter the Nature Of Income at Sr. No  " & i & " in Table D of Schedule FA")
            

            ValidateNatureOfIncome_D_FA = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(NatureOfIncome_FA(i)) Then
            msgbox_D_FA ("* Nature Of income at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table D of Schedule FA")
            ValidateNatureOfIncome_D_FA = False
            Exit Function
        End If
        
        Next
End Function

Function ValidateAmount_D_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateAmount_D_FA = True
    setTblinfo_D_FA
    
    Set rangecells = sourceSheet.Range("FA_D_Amount").Cells
    ReDim AmountD_FA(end_D_FA)
    For i = 1 To end_D_FA
        AmountD_FA(i) = rangecells.item(i).value
               
        If Not chkCompulsory(AmountD_FA(i)) Then
            'msgbox_D_FA ("* Please enter the Amount at Sr. No  " & i & "  in Sheet TR_FA  is Mandatory in Table D of Schedule FA")
            
            ''Added by Aavula SIT-SIT-69091
            
            msgbox_D_FA ("* Please enter the amount of the income taxable at Sr. No " & i & " in Table D of Schedule FA.")
            ValidateAmount_D_FA = False
            Exit Function
        End If
        Next
End Function

Function ValidateScheduleOffered_D_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateScheduleOffered_D_FA = True
    setTblinfo_D_FA
    
    Set rangecells = sourceSheet.Range("FA_D_ScheduleOffered").Cells
    ReDim ScheduleofferedD_FA(end_D_FA)
    For i = 1 To end_D_FA
        ScheduleofferedD_FA(i) = rangecells.item(i).value
               
        If isdropdownblank(ScheduleofferedD_FA(i)) Then
            'msgbox_D_FA ("* Please select the Schedule Offered at Sr. No  " & i & " in Table D of Schedule FA")
            
            
            'Added by Aavula SIT-SIT-69092
            msgbox_D_FA ("*Please enter the Schedule where income was offered at Sr. No " & i & " in Table D of Schedule FA.")

            ValidateScheduleOffered_D_FA = False
            Exit Function
        End If
        
'        If Not checkfieldspecialcharacter(ScheduleofferedD_FA(i)) Then
'            msgbox_D_FA ("* Schedule Offered at Sr. No  " & i & " in Sheet TR_FA  characters < > & ' " & Chr(34) & " are not allowed")
'            ValidateScheduleOffered_D_FA = False
'            Exit Function
'        End If
        
        Next
End Function

Function ValidateItemNumSchedule_D_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateItemNumSchedule_D_FA = True
    setTblinfo_D_FA
    
    Set rangecells = sourceSheet.Range("FA_D_ItemNumOfSchedule").Cells
    ReDim ItemNumSchedule_FA(end_D_FA)
    For i = 1 To end_D_FA
        ItemNumSchedule_FA(i) = rangecells.item(i).value
               
        If Not chkCompulsory(ItemNumSchedule_FA(i)) Then
            msgbox_D_FA ("* Please enter the Item number of schedule at Sr. No  " & i & " in Table D of Schedule FA")
            ValidateItemNumSchedule_D_FA = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(ItemNumSchedule_FA(i)) Then
            msgbox_D_FA ("* Item No. of Schedule at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Table D of Schedule FA")
            ValidateItemNumSchedule_D_FA = False
            Exit Function
        End If
        Next
End Function



Function ValidateTotalInv_D_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateTotalInv_D_FA = True
    setTblinfo_D_FA
    
    Set rangecells = sourceSheet.Range("FA_D_TotalInv").Cells
    ReDim TotalInvD_FA(end_D_FA)
    For i = 1 To end_D_FA
        TotalInvD_FA(i) = rangecells.item(i).value
        If Len(TotalInvD_FA(i)) = 0 Then
        'Exit For
        End If
        
        If Not chkCompulsory(TotalInvD_FA(i)) Then
            msgbox_D_FA ("* Please enter the Total Investment at Sr. No  " & i & " in Table D of Schedule FA")
            ValidateTotalInv_D_FA = False
            Exit Function
        End If
        
        If (TotalInvD_FA(i) < 0) Then
            msgbox_D_FA ("* Total Investment at Sr. No  " & i & "  should be Non negative, no decimal, upto 99,999,999,999,999 in Table D of Schedule FA")
            ValidateTotalInv_D_FA = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(TotalInvD_FA(i)) Then
            msgbox_D_FA ("* Total Investment  at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table D of Schedule FA")
            ValidateTotalInv_D_FA = False
            Exit Function
        End If
    Next
End Function


Function ValidateInvTotal_D_FA() As Boolean
    ValidateInvTotal_D_FA = True
    InvTotalD_FA = sourceSheet.Range("FA_D_InvTotal").value
End Function


Function ValidatesheetFFA() As Boolean
    ValidatesheetFFA = True
    
    Dim i As Long
    If Not ValidateCountry_F_FA() Then ValidatesheetFFA = False
    'If Not isdropdownblank(sourceSheet.Range("FA_F_Country")(1).value) And (Len(sourceSheet.Range("FA_F_Country").item(1).value) > 0) Then
        If Not ValidateZipCode_F_FA() Then ValidatesheetFFA = False
        If Not ValidateNameOfTrust_F_FA() Then ValidatesheetFFA = False
        If Not ValidateAddressOfTrust_F_FA() Then ValidatesheetFFA = False
        If Not ValidateNameOfOtherTrustees_F_FA() Then ValidatesheetFFA = False
        If Not ValidateAddressOfOtherTrustees_F_FA() Then ValidatesheetFFA = False
        If Not ValidateNameOfSettlor_F_FA() Then ValidatesheetFFA = False
        If Not ValidateAddressOfSettlor_F_FA() Then ValidatesheetFFA = False
        If Not ValidateNameOfBeneficiaries_F_FA() Then ValidatesheetFFA = False
        If Not ValidateAddressOfBeneficiaries_F_FA() Then ValidatesheetFFA = False
        If Not ValidateDateSincePositionHeld_F_FA() Then ValidatesheetFFA = False
        If Not ValidateWhetherincomeTaxable_F_FA() Then ValidatesheetFFA = False
        If Not ValidateIncomeDerived_F_FA() Then ValidatesheetFFA = False
        If Not ValidateAmount_F_FA() Then ValidatesheetFFA = False
        If Not ValidateScheduleoffered_F_FA() Then ValidatesheetFFA = False
        If Not ValidateItemNumSchedule_F_FA() Then ValidatesheetFFA = False
    'End If
    
    setTblinfo_F_FA2
    
'    If (end_F_FA <> end_F_FA2) Then
'        msgbox_F_FA ("* Details of Country in Table F of Schedule FA are Mandatory")
'        ValidatesheetFFA = False
'        Exit Function
'    End If
    
    'Added by Aavula 16/07/2024
    For i = 1 To Sheet27.Range("FA_F_Country").count

    If Sheet27.Range("FA_F_Country").Cells(i, 1).value <> "" And Sheet27.Range("FA_F_Country").Cells(i, 1).value <> "(Select)" Or Sheet27.Range("FA_F_ZipCode").Cells(i, 1).value <> "" Or Sheet27.Range("FA_F_NameOfTrust").Cells(i, 1).value <> "" Or Sheet27.Range("FA_F_AddressOfTrust").Cells(i, 1).value <> "" Or _
        (Sheet27.Range("FA_F_NameOfOtherTrustees").Cells(i, 1).value <> "" And Sheet27.Range("FA_F_AddressOfOtherTrustees").Cells(i, 1).value <> "") Or _
        Sheet27.Range("FA_F_NameOfSettlor").Cells(i, 1).value <> "" Or Sheet27.Range("FA_F_AddressOfSettlor").Cells(i, 1).value <> "" Or Sheet27.Range("FA_F_NameOfBeneficiaries").Cells(i, 1).value <> "" Or Sheet27.Range("FA_F_AddressOfBeneficiaries").Cells(i, 1).value <> "" Or Sheet27.Range("FA_F_DateSinceHeld").Cells(i, 1).value <> "" Or Sheet27.Range("FA_F_IncomeDerivedIsTaxable").Cells(i, 1).value <> "" And Sheet27.Range("FA_F_IncomeDerivedIsTaxable").Cells(i, 1).value <> "(Select)" Or Sheet27.Range("FA_F_IncomeDerivedAmount").Cells(i, 1).value <> "" Or Sheet27.Range("FA_F_Amount").Cells(i, 1).value <> "" Or Sheet27.Range("FA_F_Scheduleoffered").Cells(i, 1).value <> "" And Sheet27.Range("FA_F_Scheduleoffered").Cells(i, 1).value <> "(Select)" Or Sheet27.Range("FA_F_ItemNumberSchedule").Cells(i, 1).value <> "" Then
   
        
        If Sheet27.Range("FA_F_Country").Cells(i, 1).value = "" Or Sheet27.Range("FA_F_Country").Cells(i, 1).value = "(Select)" Then
            ValidatesheetFFA = False
            msgbox_F_FA ("Please select the Country Name and Code in Table F of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_F_ZipCode").Cells(i, 1).value = "" Then
            ValidatesheetFFA = False
            msgbox_F_FA ("Please enter ZIP code, if  ZIP code is not available, then enter XXXXXX in Table F of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_F_NameOfTrust").Cells(i, 1).value = "" Then
            ValidatesheetFFA = False
            msgbox_F_FA ("Please enter the Name of the trust in Table F of Schedule FA at Sr. No " & i & ".")
        End If


        If Sheet27.Range("FA_F_AddressOfTrust").Cells(i, 1).value = "" Then
            ValidatesheetFFA = False
            msgbox_F_FA ("Please enter the Address of the trust in Table F of Schedule FA at Sr. No " & i & ".")
        End If

       If Sheet27.Range("FA_F_NameOfOtherTrustees").Cells(i, 1).value = "" Then
            ValidatesheetFFA = False
            msgbox_F_FA ("Please enter the Name of the trustees in Table F of Schedule FA at Sr. No " & i & ".")
        End If

       If Sheet27.Range("FA_F_AddressOfOtherTrustees").Cells(i, 1).value = "" Then
            ValidatesheetFFA = False
            msgbox_F_FA ("Please enter the Address of the trustees in Table F of Schedule FA at Sr. No " & i & ".")
        End If


        If Sheet27.Range("FA_F_NameOfSettlor").Cells(i, 1).value = "" Then
            ValidatesheetFFA = False
            msgbox_F_FA ("Please enter the Name of the Settlor in Table F of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_F_AddressOfSettlor").Cells(i, 1).value = "" Then
            ValidatesheetFFA = False
            msgbox_F_FA ("Please enter the Address of the Settlor in Table F of Schedule FA at Sr. No " & i & ".")
        End If


       If Sheet27.Range("FA_F_NameOfBeneficiaries").Cells(i, 1).value = "" Then
            ValidatesheetFFA = False
            msgbox_F_FA ("Please enter the Name of the  Beneficiaries in Table F of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_F_AddressOfBeneficiaries").Cells(i, 1).value = "" Then
            ValidatesheetFFA = False
            msgbox_F_FA ("Please enter the Address of the  Beneficiaries in Table F of Schedule FA at Sr. No " & i & ".")
        End If


        If Sheet27.Range("FA_F_DateSinceHeld").Cells(i, 1).value = "" Then
            ValidatesheetFFA = False
            msgbox_F_FA ("Please enter the date of account opening in Table F of Schedule FA at Sr. No " & i & ".")
        End If
        
        
        
        If Sheet27.Range("FA_F_IncomeDerivedIsTaxable").Cells(i, 1).value = "" Then
            ValidatesheetFFA = False
            msgbox_F_FA ("Please select the appropriate option in Table F of Schedule FA at Sr. No " & i & ".")
        End If
        
If Sheet27.Range("FA_F_IncomeDerivedIsTaxable").Cells(i, 1).value = "Yes" Then
       
        If Sheet27.Range("FA_F_IncomeDerivedAmount").Cells(i, 1).Locked = False Then
        If Sheet27.Range("FA_F_IncomeDerivedAmount").Cells(i, 1).value = "" Then
            ValidatesheetFFA = False
            msgbox_F_FA ("Please enter the details of the income derived in the account in Table F of Schedule FA at Sr. No " & i & ".")
        End If
    End If
        
    If Sheet27.Range("FA_F_Amount").Cells(i, 1).Locked = False Then
        If Sheet27.Range("FA_F_Amount").Cells(i, 1).value = "" Then
            ValidatesheetFFA = False
            msgbox_F_FA ("Please enter the amount of the income taxable in Table F of Schedule FA at Sr. No " & i & ".")
        End If
    End If
        
    If Sheet27.Range("FA_F_Scheduleoffered").Cells(i, 1).Locked = False Then
        If Sheet27.Range("FA_F_Scheduleoffered").Cells(i, 1).value = "" Or Sheet27.Range("FA_F_Scheduleoffered").Cells(i, 1).value = "(Select)" Then
            ValidatesheetFFA = False
            msgbox_F_FA ("Please enter the Schedule where income was offered in Table F of Schedule FA at Sr. No " & i & ".")
        End If
    End If
        
    If Sheet27.Range("FA_F_ItemNumberSchedule").Cells(i, 1).Locked = False Then
        If Sheet27.Range("FA_F_ItemNumberSchedule").Cells(i, 1).value = "" Then
            ValidatesheetFFA = False
            msgbox_F_FA ("Please enter the Item number of schedule in Table F of Schedule FA at Sr. No " & i & ".")
        End If
        End If
End If
        End If
Next i
    ''''''''''
    
    
    
'    If (CountF1 <> CountF11) Then
'        msgbox_F_FA ("* Enter all mandatory field in Table F of Schedule FA")
'        ValidatesheetFFA = False
'        Exit Function
'    End If
End Function

Function ValidateCountry_F_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim CountryName As String
    ValidateCountry_F_FA = True
    setTblinfo_F_FA
    
    Set rangecells = sourceSheet.Range("FA_F_Country").Cells
    ReDim CountryCodeF_FA(end_F_FA)
    ReDim CountryNameF_FA(end_F_FA)
    noOfProcessSub = end_F_FA
    For i = 1 To end_F_FA
        CountryName = rangecells.item(i).value
        
'        If isdropdownblank(CountryName) Then
'            msgbox_F_FA ("* Please select the CountryName and Code at Sr. No  " & i & " in Table F of Schedule FA")
'            ValidateCountry_F_FA = False
'            Exit Function
'        End If
'
        
        If CountryName <> "" Then
        CountryNameF_FA(i) = Mid(CountryName, WorksheetFunction.Search("-", CountryName) + 1, Len(CountryName))
        CountryCodeF_FA(i) = Mid(CountryName, 1, WorksheetFunction.Search("-", CountryName) - 1)
        End If
        
        
'        If Not chkCompulsory(CountryCodeF_FA(i)) Then
'            msgbox_F_FA ("* CountryName and Code at Sr. No  " & i & " in Table F of Schedule FA")
'            ValidateCountry_F_FA = False
'            Exit Function
'        End If
        
'        If Not chkCompulsory(CountryCodeF_FA(i)) Then
'            msgbox_F_FA ("* CountryName at Sr. No  " & i & "  in Sheet TR_FA  is Mandatory")
'            ValidateCountry_F_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(CountryCodeF_FA(i)) Then
            msgbox_F_FA ("* CountryName and Code at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table F of Schedule FA")
            ValidateCountry_F_FA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function


Function ValidateZipCode_F_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateZipCode_F_FA = True
    setTblinfo_F_FA
    
    Set rangecells = sourceSheet.Range("FA_F_ZipCode").Cells
    ReDim ZipCodeF_FA(end_F_FA)
    For i = 1 To end_F_FA
        ZipCodeF_FA(i) = rangecells.item(i).value
        If Len(ZipCodeF_FA(i)) = 0 Then
            'Exit For
        End If
'
'        If Not chkCompulsory(ZipCodeF_FA(i)) Then
'            'msgbox_F_FA ("* Please enter the ZipCode at Sr. No  " & i & " is Mandatory.If  ZIP code is not available, then enter XXXXXX in Table F of Schedule FA")
'
'            msgbox_F_FA ("* Please enter ZIP code, if  ZIP code is not available, then enter XXXXXX at Sr. No " & i & " in Table F of Schedule FA.")
'
'            ValidateZipCode_F_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(ZipCodeF_FA(i)) Then
            msgbox_F_FA ("* ZipCode at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table F of Schedule FA")
            ValidateZipCode_F_FA = False
            Exit Function
        End If
    Next
End Function


Function ValidateNameOfTrust_F_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateNameOfTrust_F_FA = True
    setTblinfo_F_FA
    
    Set rangecells = sourceSheet.Range("FA_F_NameOfTrust").Cells
    ReDim NameOfTrustF_FA(end_F_FA)
    
    For i = 1 To end_F_FA
        NameOfTrustF_FA(i) = rangecells.item(i).value
        If Len(NameOfTrustF_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(NameOfTrustF_FA(i)) Then
'            msgbox_F_FA ("* Please enter the Name of Trust  at Sr. No  " & i & "  in Table F of Schedule FA")
'            ValidateNameOfTrust_F_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(NameOfTrustF_FA(i)) Then
            msgbox_F_FA ("* Name of Trust at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Table F of Schedule FA")
            ValidateNameOfTrust_F_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateAddressOfTrust_F_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateAddressOfTrust_F_FA = True
    setTblinfo_F_FA
    
    Set rangecells = sourceSheet.Range("FA_F_AddressOfTrust").Cells
    ReDim AddressOfTrustF_FA(end_F_FA)
    For i = 1 To end_F_FA
        AddressOfTrustF_FA(i) = rangecells.item(i).value
        If Len(AddressOfTrustF_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(AddressOfTrustF_FA(i)) Then
'            msgbox_F_FA ("* Please enter the Address of Trust  at Sr. No  " & i & " in Table F of Schedule FA")
'            ValidateAddressOfTrust_F_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(AddressOfTrustF_FA(i)) Then
            msgbox_F_FA ("* Address of Trust at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table F of Schedule FA")
            ValidateAddressOfTrust_F_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateNameOfOtherTrustees_F_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateNameOfOtherTrustees_F_FA = True
    setTblinfo_F_FA
    
    Set rangecells = sourceSheet.Range("FA_F_NameOfOtherTrustees").Cells
    ReDim NameOfOtherTrusteesF_FA(end_F_FA)
    
    For i = 1 To end_F_FA
        NameOfOtherTrusteesF_FA(i) = rangecells.item(i).value
        If Len(NameOfOtherTrusteesF_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(NameOfOtherTrusteesF_FA(i)) Then
'            msgbox_F_FA ("* Please enter the Name of  Trustees  at Sr. No  " & i & "  in Table F of Schedule FA")
'            ValidateNameOfOtherTrustees_F_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(NameOfOtherTrusteesF_FA(i)) Then
            msgbox_F_FA ("* Name of  Trustees at Sr. No  " & i & "characters < > & ' " & Chr(34) & " are not allowed in Table F of Schedule FA")
            ValidateNameOfOtherTrustees_F_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateAddressOfOtherTrustees_F_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateAddressOfOtherTrustees_F_FA = True
    setTblinfo_F_FA
    
    Set rangecells = sourceSheet.Range("FA_F_AddressOfOtherTrustees").Cells
    ReDim AddressOfOtherTrusteesF_FA(end_F_FA)
    For i = 1 To end_F_FA
        AddressOfOtherTrusteesF_FA(i) = rangecells.item(i).value
        If Len(AddressOfOtherTrusteesF_FA(i)) = 0 Then
        'Exit For
        End If
'
'        If Not chkCompulsory(AddressOfOtherTrusteesF_FA(i)) Then
'            msgbox_F_FA ("* Please enter the Address of  Trustees  at Sr. No  " & i & "  in Table F of Schedule FA")
'            ValidateAddressOfOtherTrustees_F_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(AddressOfOtherTrusteesF_FA(i)) Then
            msgbox_F_FA ("* Address of  Trustees at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Table F of Schedule FA")
            ValidateAddressOfOtherTrustees_F_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateNameOfSettlor_F_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateNameOfSettlor_F_FA = True
    setTblinfo_F_FA
    
    Set rangecells = sourceSheet.Range("FA_F_NameOfSettlor").Cells
    ReDim NameOfSettlorF_FA(end_F_FA)
    For i = 1 To end_F_FA
        NameOfSettlorF_FA(i) = rangecells.item(i).value
        If Len(NameOfSettlorF_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(NameOfSettlorF_FA(i)) Then
'            msgbox_F_FA ("* Please enter the Name of Settlor at Sr. No  " & i & " in Table F of Schedule FA")
'            ValidateNameOfSettlor_F_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(NameOfSettlorF_FA(i)) Then
            msgbox_F_FA ("* Name of Settlor at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table F of Schedule FA")
            ValidateNameOfSettlor_F_FA = False
            Exit Function
        End If
     Next
End Function

Function ValidateAddressOfSettlor_F_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateAddressOfSettlor_F_FA = True
    setTblinfo_F_FA
    
    Set rangecells = sourceSheet.Range("FA_F_AddressOfSettlor").Cells
    ReDim AddressOfSettlorF_FA(end_F_FA)
    For i = 1 To end_F_FA
        AddressOfSettlorF_FA(i) = rangecells.item(i).value
        If Len(AddressOfSettlorF_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(AddressOfSettlorF_FA(i)) Then
'            msgbox_F_FA ("* Please enter the Address of Settlor at Sr. No  " & i & " in Table F of Schedule FA")
'            ValidateAddressOfSettlor_F_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(AddressOfSettlorF_FA(i)) Then
            msgbox_F_FA ("* Address of Settlor at Sr. No  " & i & "   characters < > & ' " & Chr(34) & " are not allowed in Table F of Schedule FA")
            ValidateAddressOfSettlor_F_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateNameOfBeneficiaries_F_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateNameOfBeneficiaries_F_FA = True
    setTblinfo_F_FA
    
    Set rangecells = sourceSheet.Range("FA_F_NameOfBeneficiaries").Cells
    ReDim NameOfBeneficiariesF_FA(end_F_FA)
    For i = 1 To end_F_FA
        NameOfBeneficiariesF_FA(i) = rangecells.item(i).value
        If Len(NameOfBeneficiariesF_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(NameOfBeneficiariesF_FA(i)) Then
'            msgbox_F_FA ("* Please enter the Name of Beneficiaries at Sr. No  " & i & " in Table F of Schedule FA")
'            ValidateNameOfBeneficiaries_F_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(NameOfBeneficiariesF_FA(i)) Then
            msgbox_F_FA ("* Name of Beneficiares at Sr. No  " & i & "   characters < > & ' " & Chr(34) & " are not allowed in Table F of Schedule FA")
            ValidateNameOfBeneficiaries_F_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateAddressOfBeneficiaries_F_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateAddressOfBeneficiaries_F_FA = True
    setTblinfo_F_FA
    
    Set rangecells = sourceSheet.Range("FA_F_AddressOfBeneficiaries").Cells
    ReDim AddressOfBeneficiariesF_FA(end_F_FA)
    For i = 1 To end_F_FA
        AddressOfBeneficiariesF_FA(i) = rangecells.item(i).value
        If Len(AddressOfBeneficiariesF_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(AddressOfBeneficiariesF_FA(i)) Then
'            msgbox_F_FA ("* Please enter the Address of Beneficiaries at Sr. No  " & i & "  in Table F of Schedule FA")
'            ValidateAddressOfBeneficiaries_F_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(AddressOfBeneficiariesF_FA(i)) Then
            msgbox_F_FA ("* Address of Beneficiares at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table F of Schedule FA")
            ValidateAddressOfBeneficiaries_F_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateDateSincePositionHeld_F_FA() As Boolean
 Dim rangecells As Range
    Dim i As Long
 
    ValidateDateSincePositionHeld_F_FA = True
    setTblinfo_F_FA
    
    Set rangecells = sourceSheet.Range("FA_F_DateSinceHeld").Cells
    ReDim DateSinceHeldF_FA(end_F_FA)
    For i = 1 To end_F_FA
        DateSinceHeldF_FA(i) = rangecells.item(i).value
'        If Not chkCompulsory(DateSinceHeldF_FA(i)) Then
'            'msgbox_F_FA ("* Please enter the Date since held at Sr. No  " & i & " in Table F of Schedule FA")
'
'            ''Added by Aavula SIT-69088
'            msgbox_F_FA ("* Please enter the date of account opening at Sr. No  " & i & " in Table F of Schedule FA.")
'            ValidateDateSincePositionHeld_F_FA = False
'            Exit Function
'        End If
        If Not CheckDateddmmyyyy(DateSinceHeldF_FA(i)) Then
            msgbox_F_FA ("* Date since held at Sr. No  " & i & " is invalid in Table F of Schedule FA")
            ValidateDateSincePositionHeld_F_FA = False
            Exit Function
        End If
        Next
End Function

Function ValidateWhetherincomeTaxable_F_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateWhetherincomeTaxable_F_FA = True
    setTblinfo_F_FA
    
    Set rangecells = sourceSheet.Range("FA_F_IncomeDerivedIsTaxable").Cells
    ReDim WheteherTaxableIncomeF_FA(end_F_FA)
    For i = 1 To end_F_FA
        WheteherTaxableIncomeF_FA(i) = Mid(rangecells.item(i).value, 1, 1)
        
'        If isdropdownblank(WheteherTaxableIncomeF_FA(i)) Then
'            'msgbox_F_FA ("* Please select the Whether income taxable  Sr. No  " & i & " in Table F of Schedule FA")
'
'            ''Added by Aavula SIT-69102
'           msgbox_F_FA ("* Please select the appropriate option at Sr. No " & i & " in Table F of Schedule FA.")
'            ValidateWhetherincomeTaxable_F_FA = False
'            Exit Function
'        End If
        Next
End Function

Function ValidateIncomeDerived_F_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateIncomeDerived_F_FA = True
    setTblinfo_F_FA
    
    Set rangecells = sourceSheet.Range("FA_F_IncomeDerivedAmount").Cells
    ReDim IncomeDerivedF_FA(end_F_FA)
    For i = 1 To end_F_FA
        IncomeDerivedF_FA(i) = rangecells.item(i).value
'      If WheteherTaxableIncomeF_FA(i) = "Y" Then
''        If Not chkCompulsory(IncomeDerivedF_FA(i)) Then
''            'msgbox_F_FA ("* Please enter the Income Derived from trust  Sr. No  " & i & "  in Table F of Schedule FA")
''
''            'Added by Aavula SIT-69101
''             msgbox_F_FA ("* Please enter the details of the income derived in the account at Sr. No " & i & " in Table F of Schedule FA.")
''            ValidateIncomeDerived_F_FA = False
''            Exit Function
''        End If
'      End If
    Next
End Function

Function ValidateAmount_F_FA() As Boolean
Dim rangecells As Range
Dim i As Long
 
    ValidateAmount_F_FA = True
    setTblinfo_F_FA
    
    Set rangecells = sourceSheet.Range("FA_F_Amount").Cells
    ReDim AmountF_FA(end_F_FA)
    For i = 1 To end_F_FA
        AmountF_FA(i) = rangecells.item(i).value
'        If WheteherTaxableIncomeF_FA(i) = "Y" Then
''        If Not chkCompulsory(AmountF_FA(i)) Then
''            'msgbox_F_FA ("* Please enter the Amount at Sr. No  " & i & "  in Table F of Schedule FA")
''
''             'Added by Aavula SIT-69099
''             msgbox_F_FA ("* Please enter the amount of the income taxable at Sr. No " & i & "  in Table F of Schedule FA.")
''            ValidateAmount_F_FA = False
''            Exit Function
''         End If
'        End If
        Next
End Function

Function ValidateScheduleoffered_F_FA() As Boolean
Dim rangecells As Range
Dim i As Long
 
    ValidateScheduleoffered_F_FA = True
    setTblinfo_F_FA
    
    Set rangecells = sourceSheet.Range("FA_F_Scheduleoffered").Cells
    ReDim ScheduleOfferedF_FA(end_F_FA)
    For i = 1 To end_F_FA
        ScheduleOfferedF_FA(i) = rangecells.item(i).value
       If WheteherTaxableIncomeF_FA(i) = "Y" Then
'        If isdropdownblank(ScheduleOfferedF_FA(i)) Then
'            'msgbox_F_FA ("* Please enter the Schedule offered at Sr. No  " & i & " in Table F of Schedule FA")
'
'            'Added by Aavula SIT-69098
'             msgbox_F_FA ("* Please enter the Schedule where income was offered at Sr. No  " & i & " in Table F of Schedule FA.")
'
'            ValidateScheduleoffered_F_FA = False
'            Exit Function
'        End If
       
'        If Not checkfieldspecialcharacter(ScheduleOfferedF_FA(i)) Then
'            msgbox_F_FA ("* Schedule Offered at Sr. No  " & i & " in Sheet TR_FA  characters < > & ' " & Chr(34) & " are not allowed")
'            ValidateScheduleoffered_F_FA = False
'            Exit Function
'        End If
        End If
        Next
End Function

Function ValidateItemNumSchedule_F_FA() As Boolean
Dim rangecells As Range
Dim i As Long
 
    ValidateItemNumSchedule_F_FA = True
    setTblinfo_F_FA
    
    Set rangecells = sourceSheet.Range("FA_F_ItemNumberSchedule").Cells
    ReDim ItemNumOfScheduleF_FA(end_F_FA)
    For i = 1 To end_F_FA
        ItemNumOfScheduleF_FA(i) = rangecells.item(i).value
      If WheteherTaxableIncomeF_FA(i) = "Y" Then
'        If Not chkCompulsory(ItemNumOfScheduleF_FA(i)) Then
'            msgbox_F_FA ("* Please enter the Item No of Schedule at Sr. No  " & i & " in Table F of Schedule FA")
'            ValidateItemNumSchedule_F_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(ItemNumOfScheduleF_FA(i)) Then
            msgbox_F_FA ("* Item No of Schedule at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table F of Schedule FA")
            ValidateItemNumSchedule_F_FA = False
            Exit Function
        End If
        End If
        Next
End Function

Function ValidatesheetEFA() As Boolean
    ValidatesheetEFA = True
    
    Dim i As Long
    
    If Not ValidateInstName_E_FA() Then ValidatesheetEFA = False
    'If Not isdropdownblank(sourceSheet.Range("FA_E_InstName")(1).value) And (Len(sourceSheet.Range("FA_E_InstName").item(1).value) > 0) Then
        If Not ValidateInstAdd_E_FA() Then ValidatesheetEFA = False
        If Not ValidateCountry_E_FA() Then ValidatesheetEFA = False
        If Not ValidateZipCode_E_FA() Then ValidatesheetEFA = False
        If Not ValidateAccountName_E_FA() Then ValidatesheetEFA = False
        If Not ValidateInstitutionAccountNumber_E_FA() Then ValidatesheetEFA = False
        If Not ValidateWhetherIncomeTaxable_E_FA() Then ValidatesheetEFA = False
        If Not ValidateIncomeIncurred_E_FA() Then ValidatesheetEFA = False
        If Not ValidateAmount_E_FA() Then ValidatesheetEFA = False
        If Not ValidateScheduleOffered_E_FA() Then ValidatesheetEFA = False
        If Not ValidateItemNumOfSch_E_FA() Then ValidatesheetEFA = False
        If Not ValidatePeakBalInv_E_FA() Then ValidatesheetEFA = False
        If Not ValidateInvPBalTot_E_FA() Then ValidatesheetEFA = False
'   End If
    
    setTblinfo_E_FA2
'    If (end_E_FA <> end_E_FA2) Then
'        msgbox_E_FA ("* Details of Inst. Name in Table E of Schedule FA are Mandatory")
'        ValidatesheetEFA = False
'        Exit Function
'    End If
    
    
    'Added by Aavula 16/07/2024

    For i = 1 To Sheet27.Range("FA_E_InstName").count

    If Sheet27.Range("FA_E_InstName").Cells(i, 1).value <> "" Or Sheet27.Range("FA_E_InstAdd").Cells(i, 1).value <> "" Or Sheet27.Range("FA_E_Country").Cells(i, 1).value <> "" And Sheet27.Range("FA_E_Country").Cells(i, 1).value <> "(Select)" Or Sheet27.Range("FA_E_ZipCode").Cells(i, 1).value <> "" Or _
        (Sheet27.Range("FA_E_AccountName").Cells(i, 1).value <> "" And Sheet27.Range("FA_E_InstitutionAccountNumber").Cells(i, 1).value <> "") Or _
        Sheet27.Range("FA_E_PeakBalInv").Cells(i, 1).value <> "" Or Sheet27.Range("FA_E_IncomeAccured").Cells(i, 1).value <> "" And Sheet27.Range("FA_E_IncomeAccured").Cells(i, 1).value <> "(Select)" Or Sheet27.Range("FA_E_IncomeAccuredAmount").Cells(i, 1).value <> "" Or Sheet27.Range("FA_E_Amount").Cells(i, 1).value <> "" Or Sheet27.Range("FA_E_Scheduleoffered").Cells(i, 1).value <> "" And Sheet27.Range("FA_E_Scheduleoffered").Cells(i, 1).value <> "(Select)" Or Sheet27.Range("FA_E_ItemNumberSchedule").Cells(i, 1).value <> "" Then
   
    
        If Sheet27.Range("FA_E_InstName").Cells(i, 1).value = "" Then
            ValidatesheetEFA = False
            msgbox_E_FA ("Please enter the Name of the Institution in Table E of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_E_InstAdd").Cells(i, 1).value = "" Then
            ValidatesheetEFA = False
            msgbox_E_FA ("Please enter the Address of the Institution in Table E of Schedule FA at Sr. No " & i & ".")
        End If
    
       If Sheet27.Range("FA_E_Country").Cells(i, 1).value = "" Or Sheet27.Range("FA_E_Country").Cells(i, 1).value = "(Select)" Then
            ValidatesheetEFA = False
            msgbox_E_FA ("Please select the Country Name and Code in Table E of Schedule FA at Sr. No " & i & ".")
        End If
 

        If Sheet27.Range("FA_E_ZipCode").Cells(i, 1).value = "" Then
            ValidatesheetEFA = False
            msgbox_E_FA ("Please enter ZIP code, if  ZIP code is not available, then enter XXXXXX in Table E of Schedule FA at Sr. No " & i & ".")
        End If
 
        If Sheet27.Range("FA_E_AccountName").Cells(i, 1).value = "" Then
            ValidatesheetEFA = False
            msgbox_E_FA ("Please enter the Name of the account holder in Table E of Schedule FA at Sr. No " & i & ".")
        End If

    
        If Sheet27.Range("FA_E_InstitutionAccountNumber").Cells(i, 1).value = "" Then
            ValidatesheetEFA = False
            msgbox_E_FA ("Please enter the Account Number in Table E of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_E_PeakBalInv").Cells(i, 1).value = "" Then
            ValidatesheetEFA = False
            msgbox_E_FA ("Please enter Peak Balance in Table E of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_E_IncomeAccured").Cells(i, 1).value = "" Then
            ValidatesheetEFA = False
            msgbox_E_FA ("Please select the appropriate option in Table E of Schedule FA at Sr. No " & i & ".")
        End If
        
'sadineni (locked false condition developed on 01/12/2025)

 If Sheet27.Range("FA_E_IncomeAccured").Cells(i, 1).value = "Yes" Then

     If Sheet27.Range("FA_E_IncomeAccuredAmount").Cells(i, 1).Locked = False Then
        If Sheet27.Range("FA_E_IncomeAccuredAmount").Cells(i, 1).value = "" Then
            ValidatesheetEFA = False
            msgbox_E_FA ("Please enter the details of the income accrued in the account in Table E of Schedule FA at Sr. No " & i & ".")
        End If
     End If
     
     If Sheet27.Range("FA_E_Amount").Cells(i, 1).Locked = False Then
        If Sheet27.Range("FA_E_Amount").Cells(i, 1).value = "" Then
            ValidatesheetEFA = False
            msgbox_E_FA ("Please enter the amount of the income taxable in Table E of Schedule FA at Sr. No " & i & ".")
        End If
    End If
    
    If Sheet27.Range("FA_E_Scheduleoffered").Cells(i, 1).Locked = False Then
        If Sheet27.Range("FA_E_Scheduleoffered").Cells(i, 1).value = "" Or Sheet27.Range("FA_E_Scheduleoffered").Cells(i, 1).value = "(Select)" Then
            ValidatesheetEFA = False
            msgbox_E_FA ("Please enter the Schedule where income was offered in Table E of Schedule FA at Sr. No " & i & ".")
        End If
     End If
     
     If Sheet27.Range("FA_E_ItemNumberSchedule").Cells(i, 1).Locked = False Then
        If Sheet27.Range("FA_E_ItemNumberSchedule").Cells(i, 1).value = "" Then
            ValidatesheetEFA = False
            msgbox_E_FA ("Please enter the Item number of schedule in Table E of Schedule FA at Sr. No " & i & ".")
        End If
        End If
End If
        End If
Next i
     
    '''''''''''''''
'    If (CountE1 <> CountE11) Then
'        msgbox_E_FA ("* Enter all mandatory field in Table E of Schedule FA")
'        ValidatesheetEFA = False
'        Exit Function
'    End If
End Function



Function ValidateInstName_E_FA() As Boolean
    Dim i As Long
    Dim rangecells As Range
 
    ValidateInstName_E_FA = True
    setTblinfo_E_FA
    
    Set rangecells = sourceSheet.Range("FA_E_InstName").Cells
    ReDim InstNameE_FA(end_E_FA)
    noOfProcessSub = end_E_FA
    For i = 1 To end_E_FA
        InstNameE_FA(i) = rangecells.item(i).value
'        If Len(InstNameE_FA(i)) = 0 Then
'            Exit For
'        End If
        
'        If Not chkCompulsory(InstNameE_FA(I)) Then
'            msgbox_E_FA ("* Please enter the Institution Name at Sr. No  " & I & "  in Table E of Schedule FA")
'            ValidateInstName_E_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(InstNameE_FA(i)) Then
            msgbox_E_FA ("* Institution Name at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Table E of Schedule FA")
            ValidateInstName_E_FA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function

Function ValidateInstAdd_E_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateInstAdd_E_FA = True
    setTblinfo_E_FA
    
    Set rangecells = sourceSheet.Range("FA_E_InstAdd").Cells
    ReDim InstAddE_FA(end_E_FA)
    For i = 1 To end_E_FA
        InstAddE_FA(i) = rangecells.item(i).value
        If Len(InstAddE_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(InstAddE_FA(I)) Then
'            msgbox_E_FA ("* Please enter the Institution Address at Sr. No  " & I & "  in Table E of Schedule FA")
'            ValidateInstAdd_E_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(InstAddE_FA(i)) Then
            msgbox_E_FA ("* Institution Address at Sr. No  " & i & " in Sheet TR_FA  characters < > & ' " & Chr(34) & " are not allowed")
            ValidateInstAdd_E_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateCountry_E_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim CountryName As String
    ValidateCountry_E_FA = True
    setTblinfo_E_FA
    
    Set rangecells = sourceSheet.Range("FA_E_Country").Cells
    ReDim CountryCodeE_FA(end_E_FA)
    ReDim CountryNameE_FA(end_E_FA)
    
    For i = 1 To end_E_FA
        CountryName = rangecells.item(i).value
        
'        If isdropdownblank(CountryName) Then
'            msgbox_E_FA ("* Please select the CountryName and Code at Sr. No  " & I & "  in Table E of Schedule FA")
'            ValidateCountry_E_FA = False
'            Exit Function
'        End If
        
        If Not isdropdownblank(CountryName) Then
            CountryNameE_FA(i) = Mid(CountryName, WorksheetFunction.Search("-", CountryName) + 1, Len(CountryName))
            CountryCodeE_FA(i) = Mid(CountryName, 1, WorksheetFunction.Search("-", CountryName) - 1)
        End If
        
         If Len(CountryCodeE_FA(i)) = 0 Then
            Exit For
        End If
        
        
        If Not checkfieldspecialcharacter1(CountryCodeE_FA(i)) Then
            msgbox_E_FA ("* CountryName and Code at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table E of Schedule FA")
            ValidateCountry_E_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateZipCode_E_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateZipCode_E_FA = True
    setTblinfo_E_FA
    
    Set rangecells = sourceSheet.Range("FA_E_ZipCode").Cells
    ReDim ZipCodeE_FA(end_E_FA)
    For i = 1 To end_E_FA
        ZipCodeE_FA(i) = rangecells.item(i).value
        If Len(ZipCodeE_FA(i)) = 0 Then
            'Exit For
        End If
        
'        If Not chkCompulsory(ZipCodeE_FA(I)) Then
''            msgbox_E_FA ("* Please enter the ZipCode at Sr. No  " & i & "    is Mandatory.If  ZIP code is not available, then enter XXXXXX in Table E of Schedule FA")
'            msgbox_E_FA ("* Please enter ZIP code, if  ZIP code is not available, then enter XXXXXX at Sr. No " & I & " in Table E of Schedule FA.")
'
'            ValidateZipCode_E_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(ZipCodeE_FA(i)) Then
            msgbox_E_FA ("* ZipCode at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table E of Schedule FA")
            ValidateZipCode_E_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateAccountName_E_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateAccountName_E_FA = True
    setTblinfo_E_FA
    
    Set rangecells = sourceSheet.Range("FA_E_AccountName").Cells
    ReDim AccountNameE_FA(end_E_FA)
    For i = 1 To end_E_FA
        AccountNameE_FA(i) = rangecells.item(i).value
        If Len(AccountNameE_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(AccountNameE_FA(I)) Then
'            msgbox_E_FA ("* Please enter the Name of the Account Holder at Sr. No  " & I & " in Table E of Schedule FA")
'            ValidateAccountName_E_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(AccountNameE_FA(i)) Then
            msgbox_E_FA ("* Name of the Account Holder at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Table E of Schedule FA")
            ValidateAccountName_E_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateInstitutionAccountNumber_E_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateInstitutionAccountNumber_E_FA = True
    setTblinfo_E_FA
    
    Set rangecells = sourceSheet.Range("FA_E_InstitutionAccountNumber").Cells
    ReDim InstitutionAccountNumberE_FA(end_E_FA)
    For i = 1 To end_E_FA
        InstitutionAccountNumberE_FA(i) = rangecells.item(i).value
        
'        If Not chkCompulsory(InstitutionAccountNumberE_FA(I)) Then
'            msgbox_E_FA ("* Please enter the Account Number at Sr. No  " & I & " in Table E of Schedule FA")
'            ValidateInstitutionAccountNumber_E_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(InstitutionAccountNumberE_FA(i)) Then
            msgbox_E_FA ("* Account Number at Sr. No  " & i & "   characters < > & ' " & Chr(34) & " are not allowed in Table E of Schedule FA")
            ValidateInstitutionAccountNumber_E_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateWhetherIncomeTaxable_E_FA() As Boolean
 Dim rangecells As Range
    Dim i As Long
 
    ValidateWhetherIncomeTaxable_E_FA = True
    setTblinfo_E_FA
    
    Set rangecells = sourceSheet.Range("FA_E_IncomeAccured").Cells
    ReDim AccountTaxableE_FA(end_E_FA)
    For i = 1 To end_E_FA
        AccountTaxableE_FA(i) = Mid(rangecells.item(i).value, 1, 1)
'        If isdropdownblank(AccountTaxableE_FA(I)) Then
''            msgbox_E_FA ("* Please select the Whether income is taxable at Sr. No  " & i & " in Table E of Schedule FA")
'             msgbox_E_FA ("*  Please select the appropriate option at Sr. No " & I & " in Table E of Schedule FA.")
'
'           '
'            ValidateWhetherIncomeTaxable_E_FA = False
'            Exit Function
'        End If
        
        
    Next
End Function

Function ValidateIncomeIncurred_E_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateIncomeIncurred_E_FA = True
    setTblinfo_E_FA
    
    Set rangecells = sourceSheet.Range("FA_E_IncomeAccuredAmount").Cells
    ReDim IncomeAccuredE_FA(end_E_FA)
    For i = 1 To end_E_FA
        IncomeAccuredE_FA(i) = rangecells.item(i).value
'        If AccountTaxableE_FA(I) = "Y" Then
'        If Not chkCompulsory(IncomeAccuredE_FA(I)) Then
''            msgbox_E_FA ("* Please enter the Income accured at Sr. No  " & i & "  in Table E of Schedule FA")
'
'            'Added by Aavula
'             msgbox_E_FA ("*Please enter the details of the income accrued in the account at Sr. No " & I & " in Table E of Schedule FA.")
'
'            ValidateIncomeIncurred_E_FA = False
'            Exit Function
'        End If
'        End If
    Next
End Function

Function ValidateAmount_E_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateAmount_E_FA = True
    setTblinfo_E_FA
    
    Set rangecells = sourceSheet.Range("FA_E_Amount").Cells
    ReDim AmountE_FA(end_E_FA)
    For i = 1 To end_E_FA
        AmountE_FA(i) = rangecells.item(i).value
'         If AccountTaxableE_FA(I) = "Y" Then
'        If Not chkCompulsory(AmountE_FA(I)) Then
''            msgbox_E_FA ("* Please enter the Amount at Sr. No  " & i & " in Table E of Schedule FA")
'
'           'Added by Aavula
'            msgbox_E_FA ("* Please enter the amount of the income taxable at Sr. No  " & I & " in Table E of Schedule FA.")
'
'            ValidateAmount_E_FA = False
'            Exit Function
'        End If
'        End If
    Next
End Function

Function ValidateScheduleOffered_E_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateScheduleOffered_E_FA = True
    setTblinfo_E_FA
    
    Set rangecells = sourceSheet.Range("FA_E_Scheduleoffered").Cells
    ReDim ScheduleOfferedE_FA(end_E_FA)
    For i = 1 To end_E_FA
        ScheduleOfferedE_FA(i) = rangecells.item(i).value
         If AccountTaxableE_FA(i) = "Y" Then
'        If isdropdownblank(ScheduleOfferedE_FA(I)) Then
'           ' msgbox_E_FA ("* Please select the Schedule Offered at Sr. No  " & i & "  in Table E of Schedule FA")
'
'        'Added by Aavula
'           msgbox_E_FA ("* Please enter the Schedule where income was offered at Sr. No " & I & " in Table E of Schedule FA.")
'
'            ValidateScheduleOffered_E_FA = False
'            Exit Function
'        End If
        
'        If Not checkfieldspecialcharacter(ScheduleOfferedE_FA(i)) Then
'            msgbox_E_FA ("* Schedule Offered at Sr. No  " & i & " in Sheet TR_FA  characters < > & ' " & Chr(34) & " are not allowed")
'            ValidateScheduleOffered_E_FA = False
'            Exit Function
'        End If
        End If
    Next
End Function

Function ValidateItemNumOfSch_E_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateItemNumOfSch_E_FA = True
    setTblinfo_E_FA
    
    Set rangecells = sourceSheet.Range("FA_E_ItemNumberSchedule").Cells
    ReDim ItemNumScheduleE_FA(end_E_FA)
    For i = 1 To end_E_FA
        ItemNumScheduleE_FA(i) = rangecells.item(i).value
         If AccountTaxableE_FA(i) = "Y" Then
'        If Not chkCompulsory(ItemNumScheduleE_FA(I)) Then
'            msgbox_E_FA ("* Pelase enter the  Item No. Of Schedule at Sr. No  " & I & " in Table E of Schedule FA")
'            ValidateItemNumOfSch_E_FA = False
'            Exit Function
'        End If
'
        If Not checkfieldspecialcharacter(ItemNumScheduleE_FA(i)) Then
            msgbox_E_FA ("* Item No. Of Schedule at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Table E of Schedule FA")
            ValidateItemNumOfSch_E_FA = False
            Exit Function
        End If
        End If
    Next
End Function

Function ValidatePeakBalInv_E_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    ValidatePeakBalInv_E_FA = True
    setTblinfo_E_FA
    
    Set rangecells = sourceSheet.Range("FA_E_PeakBalInv").Cells
    ReDim PeakBalInvE_FA(end_E_FA)
    For i = 1 To end_E_FA
        PeakBalInvE_FA(i) = rangecells.item(i).value
        If Len(PeakBalInvE_FA(i)) = 0 Then
        'Exit For
        End If
        
'        If Not chkCompulsory(PeakBalInvE_FA(I)) Then
'            msgbox_E_FA ("* Please enter then Peak Balance/Investment at Sr. No  " & I & "  in Table E of Schedule FA")
'            ValidatePeakBalInv_E_FA = False
'            Exit Function
'        End If
        
        
     '   If (PeakBalInvE_FA(i) < 0) Then
     '       msgbox_E_FA ("* Peak Balance/Investment at Sr. No  " & i & "  in Sheet TR_FA  should be Non negative, no decimal, upto 99,999,999,999,999")
     '       ValidatePeakBalInv_E_FA = False
      '      Exit Function
      '  End If
        If Not checkfieldspecialcharacter(PeakBalInvE_FA(i)) Then
            msgbox_E_FA ("* Peak Balance/Investment at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table E of Schedule FA")
            ValidatePeakBalInv_E_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateInvPBalTot_E_FA() As Boolean
    ValidateInvPBalTot_E_FA = True
    InvPBalTotE_FA = sourceSheet.Range("FA_E_InvPBalTot").value
End Function

Function ValidatesheetGFA() As Boolean
    ValidatesheetGFA = True
    Dim i As Long
    
    If Not ValidateCountry_G_FA() Then ValidatesheetGFA = False
    'If Not isdropdownblank(sourceSheet.Range("FA_G_Country").item(1).value) And (Len(sourceSheet.Range("FA_G_Country").item(1).value) > 0) Then
        If Not ValidateZipCode_G_FA() Then ValidatesheetGFA = False
        If Not ValidatePersonName_G_FA() Then ValidatesheetGFA = False
        If Not ValidateAddressPerson_G_FA() Then ValidatesheetGFA = False
        If Not ValidateIncomeDerived_G_FA() Then ValidatesheetGFA = False
        If Not ValidateNatureOfIncome_G_FA() Then ValidatesheetGFA = False
        If Not ValidateWhetherIncomeTaxable_G_FA() Then ValidatesheetGFA = False
        If Not ValidateAmount_G_FA() Then ValidatesheetGFA = False
        If Not ValidateScheduleOffered_G_FA() Then ValidatesheetGFA = False
        If Not ValidateItemNumOfSch_G_FA() Then ValidatesheetGFA = False
        
   ' End If
    
    setTblinfo_G_FA2
'    If (end_G_FA <> end_G_FA2) Then
'        msgbox_G_FA ("* Details of Inst. Name in Table G of Schedule FA are Mandatory")
'        ValidatesheetGFA = False
'        Exit Function
'    End If
    
    'Added by Aavula 16/07/2024
    
     For i = 1 To Sheet27.Range("FA_G_Country").count

    If Sheet27.Range("FA_G_Country").Cells(i, 1).value <> "" And Sheet27.Range("FA_G_Country").Cells(i, 1).value <> "(Select)" Or Sheet27.Range("FA_G_ZipCode").Cells(i, 1).value <> "" Or _
        (Sheet27.Range("FA_G_NameOfPerson").Cells(i, 1).value <> "" And Sheet27.Range("FA_G_AddressOfPerson").Cells(i, 1).value <> "") Or _
        Sheet27.Range("FA_G_IncomeDerived").Cells(i, 1).value <> "" Or Sheet27.Range("FA_G_NatureOfIncome").Cells(i, 1).value <> "" Or Sheet27.Range("FA_G_WhetherTaxable").Cells(i, 1).value <> "" And Sheet27.Range("FA_G_WhetherTaxable").Cells(i, 1).value <> "(Select)" Or Sheet27.Range("FA_G_Amount").Cells(i, 1).value <> "" Or Sheet27.Range("FA_G_Scheduleoffered").Cells(i, 1).value <> "" And Sheet27.Range("FA_G_Scheduleoffered").Cells(i, 1).value <> "(Select)" Or Sheet27.Range("FA_G_ItemNumberSchedule").Cells(i, 1).value <> "" Then
   
        
        If Sheet27.Range("FA_G_Country").Cells(i, 1).value = "" Or Sheet27.Range("FA_G_Country").Cells(i, 1).value = "(Select)" Then
            ValidatesheetGFA = False
            msgbox_G_FA ("Please select the Country name and Code in Table G of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_G_ZipCode").Cells(i, 1).value = "" Then
            ValidatesheetGFA = False
            msgbox_G_FA ("Please enter ZIP code, if  ZIP code is not available, then enter XXXXXX in Table G of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_G_NameOfPerson").Cells(i, 1).value = "" Then
            ValidatesheetGFA = False
            msgbox_G_FA ("Please enter the Name of the person in Table G of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_G_AddressOfPerson").Cells(i, 1).value = "" Then
            ValidatesheetGFA = False
            msgbox_G_FA ("Please enter the Address of the person in Table G of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_G_IncomeDerived").Cells(i, 1).value = "" Then
            ValidatesheetGFA = False
            msgbox_G_FA ("Please enter Income derived in Table G of Schedule FA at Sr. No " & i & ".")
        End If

        If Sheet27.Range("FA_G_NatureOfIncome").Cells(i, 1).value = "" Then
            ValidatesheetGFA = False
            msgbox_G_FA ("Please enter the Nature of income in Table G of Schedule FA at Sr. No " & i & ".")
        End If
        
        If Sheet27.Range("FA_G_WhetherTaxable").Cells(i, 1).value = "" Or Sheet27.Range("FA_G_WhetherTaxable").Cells(i, 1).value = "(Select)" Then
            ValidatesheetGFA = False
            msgbox_G_FA ("Please select the appropriate option in Table G of Schedule FA at Sr. No " & i & ".")
        End If

'sadineni (locked false condition developed on 01/12/2025)

If Sheet27.Range("FA_G_WhetherTaxable").Cells(i, 1).value = "Yes" Then

        If Sheet27.Range("FA_G_Amount").Cells(i, 1).Locked = False Then
        If Sheet27.Range("FA_G_Amount").Cells(i, 1).value = "" Then
            ValidatesheetGFA = False
            msgbox_G_FA ("Please enter the amount of the income taxable in Table G of Schedule FA at Sr. No " & i & ".")
        End If
     End If
     
     If Sheet27.Range("FA_G_Scheduleoffered").Cells(i, 1).Locked = False Then
        If Sheet27.Range("FA_G_Scheduleoffered").Cells(i, 1).value = "" Or Sheet27.Range("FA_G_Scheduleoffered").Cells(i, 1).value = "(Select)" Then
            ValidatesheetGFA = False
            msgbox_G_FA ("Please enter the Schedule where income was offered in Table G of Schedule FA at Sr. No " & i & ".")
        End If
    End If
    
    If Sheet27.Range("FA_G_ItemNumberSchedule").Cells(i, 1).Locked = False Then
            If Sheet27.Range("FA_G_ItemNumberSchedule").Cells(i, 1).value = "" Then
             ValidatesheetGFA = False
             msgbox_G_FA ("Please enter the Item number of schedule in Table G of Schedule FA at Sr. No " & i & ".")
        End If
        End If
 End If
        End If
Next i
    
    '''''''''
    
    
    
'    If (CountG1 <> CountG11) Then
'        msgbox_G_FA ("* Enter all mandatory field in Table G of Schedule FA")
'        ValidatesheetGFA = False
'        Exit Function
'    End If
End Function

Function ValidateCountry_G_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim CountryName As String
    
    ValidateCountry_G_FA = True
    setTblinfo_G_FA
    
    Set rangecells = sourceSheet.Range("FA_G_Country").Cells
    ReDim CountryCodeG_FA(end_G_FA)
    ReDim CountryNameG_FA(end_G_FA)
    noOfProcessSub = end_G_FA
    For i = 1 To end_G_FA
    
        CountryName = rangecells.item(i).value
        
'        If isdropdownblank(CountryName) Then
'            msgbox_G_FA ("* Please select the CountryName and Code at Sr. No  " & i & "  in Table G of Schedule FA")
'            ValidateCountry_G_FA = False
'            Exit Function
'        End If
        
        
        If CountryName <> "" Then
        CountryNameG_FA(i) = Mid(CountryName, WorksheetFunction.Search("-", CountryName) + 1, Len(CountryName))
        CountryCodeG_FA(i) = Mid(CountryName, 1, WorksheetFunction.Search("-", CountryName) - 1)
        End If
        
'
'        If Not chkCompulsory(CountryCodeG_FA(i)) Then
'            msgbox_G_FA ("* CountryName and Code at Sr. No  " & i & "  in Table G of Schedule FA")
'            ValidateCountry_G_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(CountryCodeG_FA(i)) Then
            msgbox_G_FA ("* CountryName and Code at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Table G of Schedule FA")
            ValidateCountry_G_FA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function


Function ValidateZipCode_G_FA() As Boolean
    Dim rangecells As Range
    Dim i As Long
 
    ValidateZipCode_G_FA = True
    setTblinfo_G_FA
    
    Set rangecells = sourceSheet.Range("FA_G_ZipCode").Cells
    ReDim ZipCodeG_FA(end_G_FA)
    For i = 1 To end_G_FA
        ZipCodeG_FA(i) = rangecells.item(i).value
        If Len(ZipCodeG_FA(i)) = 0 Then
            'Exit For
        End If
        
'        If Not chkCompulsory(ZipCodeG_FA(i)) Then
'            msgbox_G_FA ("* Please enter the ZipCode at Sr. No  " & i & " If  ZIP code is not available, then enter XXXXXX in Table G of Schedule FA")
'            ValidateZipCode_G_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(ZipCodeG_FA(i)) Then
            msgbox_G_FA ("* ZipCode at Sr. No  " & i & "characters < > & ' " & Chr(34) & " are not allowed in Table G of Schedule FA")
            ValidateZipCode_G_FA = False
            Exit Function
        End If
    Next
End Function


Function ValidatePersonName_G_FA() As Boolean
    Dim i As Long
    Dim rangecells As Range
 
    ValidatePersonName_G_FA = True
    setTblinfo_G_FA
    
    Set rangecells = sourceSheet.Range("FA_G_NameOfPerson").Cells
    ReDim PersonNameG_FA(end_G_FA)
    For i = 1 To end_G_FA
        PersonNameG_FA(i) = rangecells.item(i).value
       
        
'        If Not chkCompulsory(PersonNameG_FA(i)) Then
'            msgbox_G_FA ("* Please enter the Name of the Person at Sr. No  " & i & " in Table G of Schedule FA")
'            ValidatePersonName_G_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(PersonNameG_FA(i)) Then
            msgbox_G_FA ("* Name of the Person at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table G of Schedule FA")
            ValidatePersonName_G_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateAddressPerson_G_FA() As Boolean
    Dim i As Long
    Dim rangecells As Range
 
    ValidateAddressPerson_G_FA = True
    setTblinfo_G_FA
    
    Set rangecells = sourceSheet.Range("FA_G_AddressOfPerson").Cells
    ReDim PersonAddressG_FA(end_G_FA)
    For i = 1 To end_G_FA
        PersonAddressG_FA(i) = rangecells.item(i).value
       
'
'        If Not chkCompulsory(PersonAddressG_FA(i)) Then
'            msgbox_G_FA ("* Please enter the Address of the person at Sr. No  " & i & " in Table G of Schedule FA ")
'            ValidateAddressPerson_G_FA = False
'            Exit Function
'        End If
        
        If Not checkfieldspecialcharacter(PersonAddressG_FA(i)) Then
            msgbox_G_FA ("* Address of the person at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table G of Schedule FA")
            ValidateAddressPerson_G_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncomeDerived_G_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateIncomeDerived_G_FA = True
    setTblinfo_G_FA
    
    Set rangecells = sourceSheet.Range("FA_G_IncomeDerived").Cells
    ReDim IncomeDerivedG_FA(end_G_FA)
    For i = 1 To end_G_FA
        IncomeDerivedG_FA(i) = rangecells.item(i).value
'        If Not chkCompulsory(IncomeDerivedG_FA(i)) Then
'            'msgbox_G_FA ("* Income Derived at Sr. No  " & i & "  in Sheet TR_FA  is Mandatory")
'
'            'Added by Aavula SIT-69096
'            msgbox_G_FA ("* Please enter Income derived at Sr. No " & i & " in Table G of Schedule FA.")
'
'            ValidateIncomeDerived_G_FA = False
'            Exit Function
'        End If
    Next
End Function

Function ValidateNatureOfIncome_G_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateNatureOfIncome_G_FA = True
    setTblinfo_G_FA
    
    Set rangecells = sourceSheet.Range("FA_G_NatureOfIncome").Cells
    ReDim NatureIncomeG_FA(end_G_FA)
    For i = 1 To end_G_FA
        NatureIncomeG_FA(i) = rangecells.item(i).value
'        If Not chkCompulsory(NatureIncomeG_FA(i)) Then
'            msgbox_G_FA ("* Please enter the Nature Of Income at Sr. No  " & i & " in Table G of Schedule FA")
'            ValidateNatureOfIncome_G_FA = False
'            Exit Function
'        End If
        
         If Not checkfieldspecialcharacter(NatureIncomeG_FA(i)) Then
            msgbox_G_FA ("* Nature Of Income at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Table G of Schedule FA")
            ValidateNatureOfIncome_G_FA = False
            Exit Function
        End If
    Next
End Function

Function ValidateWhetherIncomeTaxable_G_FA() As Boolean
 Dim rangecells As Range
    Dim i As Long
 
    ValidateWhetherIncomeTaxable_G_FA = True
    setTblinfo_G_FA
    
    Set rangecells = sourceSheet.Range("FA_G_WhetherTaxable").Cells
    ReDim WhetherTaxableIncomeG_FA(end_G_FA)
    For i = 1 To end_G_FA
        WhetherTaxableIncomeG_FA(i) = Mid(rangecells.item(i).value, 1, 1)
'        If isdropdownblank(WhetherTaxableIncomeG_FA(i)) Then
'            'msgbox_G_FA ("* Please select the Whether income is taxable at Sr. No  " & i & " in Table G of Schedule FA")
'
'           'Added by Aavula
'            msgbox_G_FA ("* Please select the appropriate option at Sr. No " & i & " in Table G of Schedule FA.")
'
'            ValidateWhetherIncomeTaxable_G_FA = False
'            Exit Function
'        End If
    Next
End Function

Function ValidateAmount_G_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateAmount_G_FA = True
    setTblinfo_G_FA
    
    Set rangecells = sourceSheet.Range("FA_G_Amount").Cells
    ReDim AmountG_FA(end_G_FA)
    For i = 1 To end_G_FA
        AmountG_FA(i) = rangecells.item(i).value
        If WhetherTaxableIncomeG_FA(i) = "Y" Then
'            If Not chkCompulsory(AmountG_FA(i)) Then
''                msgbox_G_FA ("* Please enter the Amount at Sr. No  " & i & " in Table G of Schedule FA")
'
'              'Added by Aavula
'                msgbox_G_FA ("* Please enter the amount of the income taxable at Sr. No " & i & " in Table G of Schedule FA.")
'
'                ValidateAmount_G_FA = False
'                Exit Function
'            End If
          End If
    Next
End Function

Function ValidateScheduleOffered_G_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateScheduleOffered_G_FA = True
    setTblinfo_G_FA
    
    Set rangecells = sourceSheet.Range("FA_G_Scheduleoffered").Cells
    ReDim ScheduleOfferedG_FA(end_G_FA)
    For i = 1 To end_G_FA
        ScheduleOfferedG_FA(i) = rangecells.item(i).value
        If WhetherTaxableIncomeG_FA(i) = "Y" Then
'            If isdropdownblank(ScheduleOfferedG_FA(i)) Then
''                msgbox_G_FA ("* Please select the Schedule Offered at Sr. No  " & i & "  in Table G of Schedule FA")
'
'                'Added by Aavula
'                msgbox_G_FA ("* Please enter the Schedule where income was offered at Sr. No " & i & " in Table G of Schedule FA.")
'
'                ValidateScheduleOffered_G_FA = False
'                Exit Function
'            End If
            
'             If Not checkfieldspecialcharacter(ScheduleOfferedG_FA(i)) Then
'                msgbox_G_FA ("* Schedule Offered at Sr. No  " & i & " in Sheet TR_FA  characters < > & ' " & Chr(34) & " are not allowed")
'                ValidateScheduleOffered_G_FA = False
'                Exit Function
'            End If
        End If
    Next
End Function

Function ValidateItemNumOfSch_G_FA() As Boolean
Dim rangecells As Range
    Dim i As Long
 
    ValidateItemNumOfSch_G_FA = True
    setTblinfo_G_FA
    
    Set rangecells = sourceSheet.Range("FA_G_ItemNumberSchedule").Cells
    ReDim ItemNumSchG_FA(end_G_FA)
    For i = 1 To end_G_FA
        ItemNumSchG_FA(i) = rangecells.item(i).value
     If WhetherTaxableIncomeG_FA(i) = "Y" Then
'            If Not chkCompulsory(ItemNumSchG_FA(i)) Then
'                msgbox_G_FA ("* Please enter the Item No. Of Schedule at Sr. No  " & i & "  in Table G of Schedule FA")
'                ValidateItemNumOfSch_G_FA = False
'                Exit Function
'            End If
            
            If Not checkfieldspecialcharacter(ItemNumSchG_FA(i)) Then
                msgbox_G_FA ("* Item No. Of Schedule at Sr. No  " & i & " characters < > & ' " & Chr(34) & " are not allowed in Table G of Schedule FA")
                ValidateItemNumOfSch_G_FA = False
                Exit Function
            End If
     End If
    Next
End Function




Function msgbox_TR(strmsg As String) As String
     msgValidateSheetTR = msgValidateSheetTR & strmsg & Chr(13)
End Function

Function msgbox_A_FA(strmsg As String) As String
     msgValidateSheetAFA = msgValidateSheetAFA & strmsg & Chr(13)
End Function

Function msgbox_A1_FA(strmsg As String) As String
     msgValidateSheetA1FA = msgValidateSheetA1FA & strmsg & Chr(13)
End Function

Function msgbox_A2_FA(strmsg As String) As String
     msgValidateSheetA2FA = msgValidateSheetA2FA & strmsg & Chr(13)
End Function

Function msgbox_A3_FA(strmsg As String) As String
     msgValidateSheetA3FA = msgValidateSheetA3FA & strmsg & Chr(13)
End Function

Function msgbox_A4_FA(strmsg As String) As String
     msgValidateSheetA4FA = msgValidateSheetA4FA & strmsg & Chr(13)
End Function

Function msgbox_B_FA(strmsg As String) As String
     msgValidateSheetBFA = msgValidateSheetBFA & strmsg & Chr(13)
End Function

Function msgbox_C_FA(strmsg As String) As String
     msgValidateSheetCFA = msgValidateSheetCFA & strmsg & Chr(13)
End Function

Function msgbox_D_FA(strmsg As String) As String
     msgValidateSheetDFA = msgValidateSheetDFA & strmsg & Chr(13)
End Function

Function msgbox_F_FA(strmsg As String) As String
     msgValidateSheetFFA = msgValidateSheetFFA & strmsg & Chr(13)
End Function

Function msgbox_E_FA(strmsg As String) As String
     msgValidateSheetEFA = msgValidateSheetEFA & strmsg & Chr(13)
End Function

Function msgbox_G_FA(strmsg As String) As String
     msgValidateSheetGFA = msgValidateSheetGFA & strmsg & Chr(13)
End Function







