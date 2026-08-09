Attribute VB_Name = "mdOI"
Option Explicit

Public MethodOfAcct_5 As Variant
Public ChangeInAcctMethFlg_5 As Variant
Public ProfDeviatDueAcctMeth_5 As Variant
Public ProfDeviatDueAcctMeth_6 As Variant
Public ValRawMaterial_5 As Variant
Public ValFinishedGoods_5 As Variant
Public msgValidateSheet5 As Variant

Public ChngStockValMetFlg_6 As Variant
Public EffectOnPL_6 As Variant
Public EffectOnPL_6_4e As Variant
Public Section28Items_6 As Variant
Public ProformaCreditsDue_6 As Variant
Public PrevYrEscalClaim_6 As Variant
Public OthItemInc_6 As Variant
Public CapReceipt_6 As Variant
Public TotNoCredToPLAmt_6 As Variant
Public StkInsurPrem_6 As Variant
Public EmpHealthInsurPrem_6 As Variant
Public EmpBonusCommSum_6 As Variant
Public IntOnBorrCap_6 As Variant
Public ZeroCoupBondDisc_6 As Variant
Public RecogPFContribAmt_6 As Variant
Public AppSuperAnnFundAmt_6 As Variant
Public PensionSchemeSec80CCD_6 As Variant

Public AppGratFundAmt_6 As Variant
Public OthFundAmt_6 As Variant
Public BadDebtDoubtAmt_6 As Variant
Public BadDebtDoubtProvn_6 As Variant
Public SpecResrvTranfr_6 As Variant

Public SecuritiesPaidAmt_6 As Variant
Public CapitalNatureExp_6 As Variant
Public NonComp40aiiChapXVIIBAmt_6 As Variant
Public NonComp40aiiiChapXVIIBAmt_6 As Variant
Public RolyatyOrServiceFee_6 As Variant
Public FamPlanPromoExp_6 As Variant
Public EmpContributionCredits_6 As Variant
'Public DeployedInIndia_6 As Variant
'Public DeployedOutSideIndia_6 As Variant
'Public Total_6 As Variant
Public OthDisallowances_6 As Variant
Public TotAmtDisallUs36_6 As Variant
Public PersonalExp_6 As Variant
Public PoliticPartyExp_6 As Variant
Public BusOrProfessnExp As Variant


Public LawVoilatPenalExp_6 As Variant
Public OthPenalFineExp_6 As Variant
Public OffenceExp_6 As Variant
Public ContigentLiability_6 As Variant
Public AmtNotPartTotInc_6 As Variant
Public OthAmtNotAllowUs37_6 As Variant
Public TotAmtDisallUs37_6 As Variant
Public NonCompChapXVIIBAmt_6 As Variant
Public STTAmt_6 As Variant
Public FBTAmt_6 As Variant
Public TaxAmtOnProfits_6 As Variant
Dim msgValidateSheet6 As String


'----------------
Public WTAmt_7 As Variant
Public NonComp40aibChapXVIIBAmt_7 As Variant
Public IntSalBonPartner_7 As Variant
Public OthDisallow_7 As Variant
Public TotAmtDisallUs40_7 As Variant
Public AmtDisallUs40PyNowAll_7 As Variant
Public AmtPaidUs40A2b_7 As Variant
Public AmtGT20kCash_7 As Variant
Public ProvPmtGrat_7 As Variant
Public ContToSetupTrust_7 As Variant
Public OthDisallow1_7 As Variant
Public TotAmtDisallUs40A_7 As Variant
Public TaxDutyCesAmt_7 As Variant
Public ContToEmpPFSFGF_7 As Variant
Public EmpBonusComm_7 As Variant
Public IntPayaleToFI_7 As Variant
Public IntPayaleToFISchBank_7 As Variant
Public LeaveEncashPayable_7 As Variant
Public RailwayAssetsPayable_7 As Variant
Public TotAmtUs43b_7 As Variant
Public TaxDutyCesAmt1_7 As Variant
Public ContToEmpPFSFGF1_7 As Variant
Public EmpBonusComm1_7 As Variant
Public IntPayaleToFI1_7 As Variant
Public IntPayaleToFISchBank1_7 As Variant
Public LeaveEncashPayable1_7 As Variant
Public TotAmtUs43b1_7 As Variant
Public UnionExciseDuty_7 As Variant
Public ServiceTax_7 As Variant
Public VATorSaleTax_7 As Variant
Public OthDutyTaxCess_7 As Variant
Public TotExciseCustomsVAT_7 As Variant
Public DeemedProfUs33ABs_7 As Variant
Public DeemedProfUs33AB_7 As Variant
Public DeemedProfUs33ABA_7 As Variant
Public ProfTaxAmtUs41_7 As Variant
Public PriorAmtIncCrDrPL_7 As Variant
Public RailwayAssets_10_g As Variant

 Public CGST_7 As Variant
Public SGST_7 As Variant
Public IGST_7 As Variant
Public UTGST_7 As Variant
 
   
   
   
   

Public CustomDuty_7 As Variant  'added new
Public CounterVailDuty_7 As Variant  'added new
Public SplAddDuty_7 As Variant  'added new
Public Cess_7 As Variant  'added new



Dim msgValidateSheet7 As String
Sub ValidateOI_Click()
ValidateSheetPARTA_OI
fmsgboxStatus "Sheet Other Income (OI) is OK" ', vbOKOnly, "ITR-3"
End Sub

Sub ValidateSheetPARTA_OI()
    msgValidateSheet5 = ""
    msgValidateSheet6 = ""
    msgValidateSheet7 = ""
    
    If Not ValidateSheet6 Then
        Sheet5.Activate
        fmsgboxsmall msgValidateSheet5 & msgValidateSheet6 & msgValidateSheet7 ', vbOKOnly, "Error(s)!"
        msgValidateSheet5 = ""
        msgValidateSheet6 = ""
        msgValidateSheet7 = ""
        CloseMsg
    Else
        'MsgBox "Sheet is ok", vbOKOnly, "ITR-3"
    End If
End Sub

Function ValidateSheet6() As Boolean
    ValidateSheet6 = True
'rpcomments
    If Not ValidateMethodOfAcct_5() Then ValidateSheet6 = False
    If Not ValidateChangeInAcctMethFlg_5() Then ValidateSheet6 = False
    If Not ValidateProfDeviatDueAcctMeth_5() Then ValidateSheet6 = False
    If Not ValidateProfDeviatDueAcctMeth_6() Then ValidateSheet6 = False
    If Not ValidateValRawMaterial_5() Then ValidateSheet6 = False
    If Not ValidateValFinishedGoods_5() Then ValidateSheet6 = False

'rpcomments

    If Not ValidateChngStockValMetFlg_6() Then ValidateSheet6 = False
    If Not ValidateEffectOnPL_6() Then ValidateSheet6 = False
    If Not ValidateSection28Items_6() Then ValidateSheet6 = False
    If Not ValidateProformaCreditsDue_6() Then ValidateSheet6 = False
    If Not ValidatePrevYrEscalClaim_6() Then ValidateSheet6 = False
    If Not ValidateOthItemInc_6() Then ValidateSheet6 = False
    If Not ValidateCapReceipt_6() Then ValidateSheet6 = False
    If Not ValidateTotNoCredToPLAmt_6() Then ValidateSheet6 = False
    If Not ValidateStkInsurPrem_6() Then ValidateSheet6 = False
    If Not ValidateEmpHealthInsurPrem_6() Then ValidateSheet6 = False
    If Not ValidateEmpBonusCommSum_6() Then ValidateSheet6 = False
    If Not ValidateIntOnBorrCap_6() Then ValidateSheet6 = False
    If Not ValidateZeroCoupBondDisc_6() Then ValidateSheet6 = False
    If Not ValidateRecogPFContribAmt_6() Then ValidateSheet6 = False
    If Not ValidateAppSuperAnnFundAmt_6() Then ValidateSheet6 = False
    If Not ValidateAppGratFundAmt_6() Then ValidateSheet6 = False
    If Not ValidateOthFundAmt_6() Then ValidateSheet6 = False
    If Not ValidateBadDebtDoubtAmt_6() Then ValidateSheet6 = False
    If Not ValidateBadDebtDoubtProvn_6() Then ValidateSheet6 = False
    If Not ValidateSpecResrvTranfr_6() Then ValidateSheet6 = False
    If Not ValidateFamPlanPromoExp_6() Then ValidateSheet6 = False
    If Not ValidateEmpContributionCredits_6() Then ValidateSheet6 = False
    If Not ValidateOthDisallowances_6() Then ValidateSheet6 = False
    If Not ValidateTotAmtDisallUs36_6() Then ValidateSheet6 = False
    If Not ValidatePersonalExp_6() Then ValidateSheet6 = False
    If Not ValidatePoliticPartyExp_6() Then ValidateSheet6 = False
    If Not ValidateLawVoilatPenalExp_6() Then ValidateSheet6 = False
    If Not ValidateOthPenalFineExp_6() Then ValidateSheet6 = False
    If Not ValidateOffenceExp_6() Then ValidateSheet6 = False
    If Not ValidateContigentLiability_6() Then ValidateSheet6 = False
    'If Not ValidateAmtNotPartTotInc_6() Then ValidateSheet6 = False
    If Not ValidateOthAmtNotAllowUs37_6() Then ValidateSheet6 = False
    If Not ValidateTotAmtDisallUs37_6() Then ValidateSheet6 = False
    If Not ValidateNonCompChapXVIIBAmt_6() Then ValidateSheet6 = False
    'If Not ValidateSTTAmt_6() Then ValidateSheet6 = False
    'If Not ValidateFBTAmt_6() Then ValidateSheet6 = False
    If Not ValidateTaxAmtOnProfits_6() Then ValidateSheet6 = False

    If Not ValidateNonComp40aibChapXVIIBAmt_7() Then ValidateSheet6 = False
    If Not ValidateWTAmt_7() Then ValidateSheet6 = False
    If Not ValidateIntSalBonPartner_7() Then ValidateSheet6 = False
    If Not ValidateOthDisallow_7() Then ValidateSheet6 = False
    If Not ValidateTotAmtDisallUs40_7() Then ValidateSheet6 = False
    If Not ValidateAmtDisallUs40PyNowAll_7() Then ValidateSheet6 = False
    If Not ValidateAmtPaidUs40A2b_7() Then ValidateSheet6 = False
    If Not ValidateAmtGT20kCash_7() Then ValidateSheet6 = False
    If Not ValidateProvPmtGrat_7() Then ValidateSheet6 = False
    If Not ValidateContToSetupTrust_7() Then ValidateSheet6 = False
    If Not ValidateOthDisallow1_7() Then ValidateSheet6 = False
    If Not ValidateTotAmtDisallUs40A_7() Then ValidateSheet6 = False
    If Not ValidateTaxDutyCesAmt_7() Then ValidateSheet6 = False
    If Not ValidateContToEmpPFSFGF_7() Then ValidateSheet6 = False
    If Not ValidateEmpBonusComm_7() Then ValidateSheet6 = False
    If Not ValidateIntPayaleToFI_7() Then ValidateSheet6 = False
    If Not ValidateIntPayaleToFISchBank_7() Then ValidateSheet6 = False
    If Not ValidateLeaveEncashPayable_7() Then ValidateSheet6 = False
    If Not ValidateRailwayAssetsPayable_7() Then ValidateSheet6 = False
    If Not ValidateTotAmtUs43b_7() Then ValidateSheet6 = False
    If Not ValidateTaxDutyCesAmt1_7() Then ValidateSheet6 = False
    If Not ValidateContToEmpPFSFGF1_7() Then ValidateSheet6 = False
    If Not ValidateEmpBonusComm1_7() Then ValidateSheet6 = False
    If Not ValidateIntPayaleToFI1_7() Then ValidateSheet6 = False
    If Not ValidateIntPayaleToFISchBank1_7() Then ValidateSheet6 = False
    If Not ValidateLeaveEncashPayable1_7() Then ValidateSheet6 = False
    If Not ValidateTotAmtUs43b1_7() Then ValidateSheet6 = False
    If Not ValidateUnionExciseDuty_7() Then ValidateSheet6 = False
    If Not ValidateServiceTax_7() Then ValidateSheet6 = False
    If Not ValidateVATorSaleTax_7() Then ValidateSheet6 = False
    If Not ValidateOthDutyTaxCess_7() Then ValidateSheet6 = False
    If Not ValidateTotExciseCustomsVAT_7() Then ValidateSheet6 = False
    If Not ValidateDeemedProfUs33ABs_7() Then ValidateSheet6 = False
    If Not ValidateProfTaxAmtUs41_7() Then ValidateSheet6 = False
    If Not ValidatePriorAmtIncCrDrPL_7() Then ValidateSheet6 = False

    

End Function

Function ValidateMethodOfAcct_5() As Boolean
    ValidateMethodOfAcct_5 = True
    MethodOfAcct_5 = Sheet5.Range("sheet5.MethodOfAcct")
    
    If MethodOfAcct_5 = "Mercantile" Then
        MethodOfAcct_5 = "MERC"
    End If
    
    If MethodOfAcct_5 = "Cash" Then
        MethodOfAcct_5 = "CASH"
    End If
    
    If MethodOfAcct_5 = "None" Then
        MethodOfAcct_5 = ""
    End If
    
    If (MethodOfAcct_5) = "(Select)" Or MethodOfAcct_5 = "" Then
      msgbox5 ("* MethodOfAcct in Sheet :  PART A-OI   is Mandatory ")
       ValidateMethodOfAcct_5 = False
        Exit Function
    End If
    
    If Len(MethodOfAcct_5) > 4 Then
        msgbox5 ("* MethodOfAcct in Sheet :  PART A-OI   should not greater than 4 charactars ")
        ValidateMethodOfAcct_5 = False
        Exit Function
    End If
    
'     If Sheet5.Range("sheet7.Section92CE_Flag").value = "(Select)" Or Sheet5.Range("sheet7.Section92CE_Flag").value = "" Then
'        msgbox5 ("*Please select Whether assessee is exercising option under subsection 2A of section 92CE in Schedule Part A-OI ")
'        ValidateMethodOfAcct_5 = False
'        Exit Function
'End If
End Function
 
Function ValidateChangeInAcctMethFlg_5() As Boolean
    ValidateChangeInAcctMethFlg_5 = True
    ChangeInAcctMethFlg_5 = Sheet5.Range("sheet5.ChangeInAcctMethFlg")
    
    If ChangeInAcctMethFlg_5 = "Yes" Then
        ChangeInAcctMethFlg_5 = "Y"
    End If
    
    If ChangeInAcctMethFlg_5 = "No" Then
        ChangeInAcctMethFlg_5 = "N"
    End If
    
    If (ChangeInAcctMethFlg_5) = "(Select)" Or ChangeInAcctMethFlg_5 = "" Then
'Change-7, ANK, 27.05.2022
        msgbox5 ("* Change in method of accounting in Sheet :  PART A-OI   is Mandatory ")
'        msgbox5 ("* ChangeInAcctMethFlg in Sheet :  PART A-OI   is Mandatory ")
'---
        ValidateChangeInAcctMethFlg_5 = False
        Exit Function
    End If
    
    If Len(ChangeInAcctMethFlg_5) > 1 Then
'Change-8, ANK, 27.05.2022
        msgbox5 ("* Change in method of accounting in Sheet :  PART A-OI   should not greater than 1 charactars ")
'        msgbox5 ("* ChangeInAcctMethFlg in Sheet :  PART A-OI   should not greater than 1 charactars ")
'---
        ValidateChangeInAcctMethFlg_5 = False
        Exit Function
    End If
End Function
 
Function ValidateProfDeviatDueAcctMeth_5() As Boolean
    Dim i As Long
    ValidateProfDeviatDueAcctMeth_5 = True
    ProfDeviatDueAcctMeth_5 = Sheet5.Range("sheet5.ProfDeviatDueAcctMeth").value
    
    For i = 1 To Len(ProfDeviatDueAcctMeth_5)
        If Not IsNumeric(ProfDeviatDueAcctMeth_5) Then
            msgbox5 ("* ProfDeviatDueAcctMeth in Sheet :  PART A-OI   must contain only digits from 0 to 9")
            ValidateProfDeviatDueAcctMeth_5 = False
            Exit Function
        End If
    Next
End Function
Function ValidateProfDeviatDueAcctMeth_6() As Boolean
    Dim i As Long
    ValidateProfDeviatDueAcctMeth_6 = True
    ProfDeviatDueAcctMeth_6 = Sheet5.Range("sheet5.ProfDeviatDueAcctMethb").value

    For i = 1 To Len(ProfDeviatDueAcctMeth_6)
        If Not IsNumeric(ProfDeviatDueAcctMeth_6) Then
            msgbox5 ("* ProfDeviatDueAcctMethb in Sheet :  PART A-OI   must contain only digits from 0 to 9")
            ValidateProfDeviatDueAcctMeth_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateValRawMaterial_5() As Boolean
    ValidateValRawMaterial_5 = True
    ValRawMaterial_5 = Sheet5.Range("sheet5.ValRawMaterial")
    

End Function
 
Function ValidateValFinishedGoods_5() As Boolean
    ValidateValFinishedGoods_5 = True
    ValFinishedGoods_5 = Sheet5.Range("sheet5.ValFinishedGoods")

End Function

Function ValidateChngStockValMetFlg_6() As Boolean
    Dim i As Long
    
    ValidateChngStockValMetFlg_6 = True
    ChngStockValMetFlg_6 = Sheet5.Range("sheet6.ChngStockValMetFlg").value
    ChngStockValMetFlg_6 = Mid(ChngStockValMetFlg_6, 1, 1)

End Function

Function ValidateEffectOnPL_6() As Boolean
    Dim i As Long
    
    ValidateEffectOnPL_6 = True
    EffectOnPL_6 = Sheet5.Range("sheet6.EffectOnPL").value
    EffectOnPL_6_4e = Sheet5.Range("sheet6.EffectOnPL4e").value
    If Range("sheet6.ChngStockValMetFlg") = "Yes" Then
    If (EffectOnPL_6 <> "" Or EffectOnPL_6_4e <> "") Then
    Else
    msgbox6 ("* Please enter the effect for either Point 4d or 4e is mandatory.")
    ValidateEffectOnPL_6 = False
    End If
    End If

    If Len(EffectOnPL_6) > 14 Then
        msgbox6 ("* EffectOnPL in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateEffectOnPL_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(EffectOnPL_6)
        If Not IsNumeric(EffectOnPL_6) Then
            msgbox6 ("* EffectOnPL in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateEffectOnPL_6 = False
            Exit Function
        End If
    
    Next

    If Len(EffectOnPL_6_4e) > 14 Then
        msgbox6 ("* EffectOnPL of Row 4e in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateEffectOnPL_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(EffectOnPL_6_4e)
        If Not IsNumeric(EffectOnPL_6_4e) Then
            msgbox6 ("* EffectOnPL of Row 4e in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateEffectOnPL_6 = False
            Exit Function
        End If
    
    Next
End Function

Function ValidateSection28Items_6() As Boolean
    Dim i As Long
    
    ValidateSection28Items_6 = True
    Section28Items_6 = Sheet5.Range("sheet6.Section28Items").value
    
    If Len(Section28Items_6) > 14 Then
        msgbox6 ("* Section28Items in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateSection28Items_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(Section28Items_6)
        If Not IsNumeric(Section28Items_6) Then
            msgbox6 ("* Section28Items in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateSection28Items_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateProformaCreditsDue_6() As Boolean
    Dim i As Long
    
    ValidateProformaCreditsDue_6 = True
    ProformaCreditsDue_6 = Sheet5.Range("sheet6.ProformaCreditsDue").value
    
    If Len(ProformaCreditsDue_6) > 14 Then
        msgbox6 ("* ProformaCreditsDue in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateProformaCreditsDue_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(ProformaCreditsDue_6)
        If Not IsNumeric(ProformaCreditsDue_6) Then
            msgbox6 ("* ProformaCreditsDue in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateProformaCreditsDue_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidatePrevYrEscalClaim_6() As Boolean
    Dim i As Long
    
    ValidatePrevYrEscalClaim_6 = True
    PrevYrEscalClaim_6 = Sheet5.Range("sheet6.PrevYrEscalClaim").value
    
    If Len(PrevYrEscalClaim_6) > 14 Then
        msgbox6 ("* PrevYrEscalClaim in Sheet : PART A-OI    should be at most 14 digits ")
        ValidatePrevYrEscalClaim_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(PrevYrEscalClaim_6)
        If Not IsNumeric(PrevYrEscalClaim_6) Then
            msgbox6 ("* PrevYrEscalClaim in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidatePrevYrEscalClaim_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateOthItemInc_6() As Boolean
    Dim i As Long
    
    ValidateOthItemInc_6 = True
    OthItemInc_6 = Sheet5.Range("sheet6.OthItemInc").value
    
    If Len(OthItemInc_6) > 14 Then
        msgbox6 ("* OthItemInc in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateOthItemInc_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(OthItemInc_6)
        If Not IsNumeric(OthItemInc_6) Then
            msgbox6 ("* OthItemInc in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateOthItemInc_6 = False
            Exit Function
        End If
    Next
End Function


Function ValidateCapReceipt_6() As Boolean
    Dim i As Long
    
    ValidateCapReceipt_6 = True
    CapReceipt_6 = Sheet5.Range("sheet6.CapReceipt").value
    
    If Len(CapReceipt_6) > 14 Then
        msgbox6 ("* CapReceipt in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateCapReceipt_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(CapReceipt_6)
        If Not IsNumeric(CapReceipt_6) Then
            msgbox6 ("* CapReceipt in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateCapReceipt_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotNoCredToPLAmt_6() As Boolean
    Dim i As Long
    
    ValidateTotNoCredToPLAmt_6 = True
    TotNoCredToPLAmt_6 = Sheet5.Range("sheet6.TotNoCredToPLAmt").value
    
    If Len(TotNoCredToPLAmt_6) > 14 Then
        msgbox6 ("* TotNoCredToPLAmt in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateTotNoCredToPLAmt_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(TotNoCredToPLAmt_6)
        If Not IsNumeric(TotNoCredToPLAmt_6) Then
            msgbox6 ("* TotNoCredToPLAmt in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateTotNoCredToPLAmt_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateStkInsurPrem_6() As Boolean
    Dim i As Long
    
    ValidateStkInsurPrem_6 = True
    StkInsurPrem_6 = Sheet5.Range("sheet6.StkInsurPrem").value
    
    If Len(StkInsurPrem_6) > 14 Then
        msgbox6 ("* StkInsurPrem in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateStkInsurPrem_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(StkInsurPrem_6)
        If Not IsNumeric(StkInsurPrem_6) Then
            msgbox6 ("* StkInsurPrem in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateStkInsurPrem_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateEmpHealthInsurPrem_6() As Boolean
    Dim i As Long
    
    ValidateEmpHealthInsurPrem_6 = True
    EmpHealthInsurPrem_6 = Sheet5.Range("sheet6.EmpHealthInsurPrem").value
    
    If Len(EmpHealthInsurPrem_6) > 14 Then
        msgbox6 ("* EmpHealthInsurPrem in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateEmpHealthInsurPrem_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(EmpHealthInsurPrem_6)
        If Not IsNumeric(EmpHealthInsurPrem_6) Then
            msgbox6 ("* EmpHealthInsurPrem in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateEmpHealthInsurPrem_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateEmpBonusCommSum_6() As Boolean
    Dim i As Long
    
    ValidateEmpBonusCommSum_6 = True
    EmpBonusCommSum_6 = Sheet5.Range("sheet6.EmpBonusCommSum").value
    
    If Len(EmpBonusCommSum_6) > 14 Then
        msgbox6 ("* EmpBonusCommSum in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateEmpBonusCommSum_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(EmpBonusCommSum_6)
        If Not IsNumeric(EmpBonusCommSum_6) Then
            msgbox6 ("* EmpBonusCommSum in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateEmpBonusCommSum_6 = False
            Exit Function
        End If
    Next
End Function


Function ValidateIntOnBorrCap_6() As Boolean
    Dim i As Long
    
    ValidateIntOnBorrCap_6 = True
    IntOnBorrCap_6 = Sheet5.Range("sheet6.IntOnBorrCap").value
    
    If Len(IntOnBorrCap_6) > 14 Then
        msgbox6 ("* IntOnBorrCap in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateIntOnBorrCap_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(IntOnBorrCap_6)
        If Not IsNumeric(IntOnBorrCap_6) Then
            msgbox6 ("* IntOnBorrCap in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateIntOnBorrCap_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateZeroCoupBondDisc_6() As Boolean
    Dim i As Long
    
    ValidateZeroCoupBondDisc_6 = True
    ZeroCoupBondDisc_6 = Sheet5.Range("sheet6.ZeroCoupBondDisc").value
    
    If Len(ZeroCoupBondDisc_6) > 14 Then
        msgbox6 ("* ZeroCoupBondDisc in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateZeroCoupBondDisc_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(ZeroCoupBondDisc_6)
        If Not IsNumeric(ZeroCoupBondDisc_6) Then
            msgbox6 ("* ZeroCoupBondDisc in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateZeroCoupBondDisc_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateRecogPFContribAmt_6() As Boolean
    Dim i As Long
    
    ValidateRecogPFContribAmt_6 = True
    RecogPFContribAmt_6 = Sheet5.Range("sheet6.RecogPFContribAmt").value
    
    If Len(RecogPFContribAmt_6) > 14 Then
        msgbox6 ("* RecogPFContribAmt in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateRecogPFContribAmt_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(RecogPFContribAmt_6)
        If Not IsNumeric(RecogPFContribAmt_6) Then
            msgbox6 ("* RecogPFContribAmt in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateRecogPFContribAmt_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateAppSuperAnnFundAmt_6() As Boolean
    Dim i As Long
    
    ValidateAppSuperAnnFundAmt_6 = True
    AppSuperAnnFundAmt_6 = Sheet5.Range("sheet6.AppSuperAnnFundAmt").value
    PensionSchemeSec80CCD_6 = Sheet5.Range("Sheet6.PensionSchemeSec80CCD").value
    
    If Len(AppSuperAnnFundAmt_6) > 14 Then
        msgbox6 ("* AppSuperAnnFundAmt in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateAppSuperAnnFundAmt_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(AppSuperAnnFundAmt_6)
        If Not IsNumeric(AppSuperAnnFundAmt_6) Then
            msgbox6 ("* AppSuperAnnFundAmt in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateAppSuperAnnFundAmt_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateAppGratFundAmt_6() As Boolean
    Dim i As Long
    
    ValidateAppGratFundAmt_6 = True
    AppGratFundAmt_6 = Sheet5.Range("sheet6.AppGratFundAmt").value
    
    If Len(AppGratFundAmt_6) > 14 Then
        msgbox6 ("* AppGratFundAmt in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateAppGratFundAmt_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(AppGratFundAmt_6)
        If Not IsNumeric(AppGratFundAmt_6) Then
            msgbox6 ("* AppGratFundAmt in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateAppGratFundAmt_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateOthFundAmt_6() As Boolean
    Dim i As Long
    
    ValidateOthFundAmt_6 = True
    OthFundAmt_6 = Sheet5.Range("sheet6.OthFundAmt").value
    
    If Len(OthFundAmt_6) > 14 Then
        msgbox6 ("* OthFundAmt in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateOthFundAmt_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(OthFundAmt_6)
        If Not IsNumeric(OthFundAmt_6) Then
            msgbox6 ("* OthFundAmt in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateOthFundAmt_6 = False
            Exit Function
        End If
    Next
End Function


Function ValidateBadDebtDoubtAmt_6() As Boolean
    Dim i As Long
    
    ValidateBadDebtDoubtAmt_6 = True
    BadDebtDoubtAmt_6 = Sheet5.Range("sheet6.BadDebtDoubtAmt").value
    
    If Len(BadDebtDoubtAmt_6) > 14 Then
        msgbox6 ("* BadDebtDoubtAmt in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateBadDebtDoubtAmt_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(BadDebtDoubtAmt_6)
        If Not IsNumeric(BadDebtDoubtAmt_6) Then
            msgbox6 ("* BadDebtDoubtAmt in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateBadDebtDoubtAmt_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBadDebtDoubtProvn_6() As Boolean
    Dim i As Long
    
    ValidateBadDebtDoubtProvn_6 = True
    BadDebtDoubtProvn_6 = Sheet5.Range("sheet6.BadDebtDoubtProvn").value
    
    If Len(BadDebtDoubtProvn_6) > 14 Then
        msgbox6 ("* BadDebtDoubtProvn in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateBadDebtDoubtProvn_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(BadDebtDoubtProvn_6)
        If Not IsNumeric(BadDebtDoubtProvn_6) Then
            msgbox6 ("* BadDebtDoubtProvn in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateBadDebtDoubtProvn_6 = False
            Exit Function
        End If
    Next
End Function


Function ValidateSpecResrvTranfr_6() As Boolean
    Dim i As Long
    
    ValidateSpecResrvTranfr_6 = True
    SpecResrvTranfr_6 = Sheet5.Range("sheet6.SpecResrvTranfr").value
    
    If Len(SpecResrvTranfr_6) > 14 Then
        msgbox6 ("* SpecResrvTranfr in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateSpecResrvTranfr_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(SpecResrvTranfr_6)
        If Not IsNumeric(SpecResrvTranfr_6) Then
            msgbox6 ("* SpecResrvTranfr in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateSpecResrvTranfr_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateFamPlanPromoExp_6() As Boolean
    Dim i As Long
    
    ValidateFamPlanPromoExp_6 = True
    FamPlanPromoExp_6 = Sheet5.Range("sheet6.FamPlanPromoExp").value
    SecuritiesPaidAmt_6 = Sheet5.Range("sheet6.SecuritiesPaidAmt").value
    CapitalNatureExp_6 = Sheet5.Range("sheet6.CapitalNatureExp").value
    NonComp40aiiChapXVIIBAmt_6 = Sheet5.Range("shee6.NonComp40aiiChapXVIIBAmt").value
    NonComp40aiiiChapXVIIBAmt_6 = Sheet5.Range("sheet6.NonComp40aiiiChapXVIIBAmt").value
    RolyatyOrServiceFee_6 = Sheet5.Range("sheet7.RolyatyOrServiceFee").value
    
    If Len(FamPlanPromoExp_6) > 14 Then
        msgbox6 ("* FamPlanPromoExp in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateFamPlanPromoExp_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(FamPlanPromoExp_6)
        If Not IsNumeric(FamPlanPromoExp_6) Then
            msgbox6 ("* FamPlanPromoExp in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateFamPlanPromoExp_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateEmpContributionCredits_6() As Boolean
    Dim i As Long
    
    ValidateEmpContributionCredits_6 = True
    EmpContributionCredits_6 = Sheet5.Range("sheet6.EmpContributionCredits").value
    
    If Len(EmpContributionCredits_6) > 14 Then
        msgbox6 ("* EmpContributionCredits in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateEmpContributionCredits_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(EmpContributionCredits_6)
        If Not IsNumeric(EmpContributionCredits_6) Then
            msgbox6 ("* EmpContributionCredits in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateEmpContributionCredits_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateOthDisallowances_6() As Boolean
    Dim i As Long
    
    ValidateOthDisallowances_6 = True
    OthDisallowances_6 = Sheet5.Range("sheet6.OthDisallowances").value

    If Len(OthDisallowances_6) > 14 Then
        msgbox6 ("* OthDisallowances in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateOthDisallowances_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(OthDisallowances_6)
        If Not IsNumeric(OthDisallowances_6) Then
            msgbox6 ("* OthDisallowances in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateOthDisallowances_6 = False
            Exit Function
        End If
    Next
End Function


Function ValidateTotAmtDisallUs36_6() As Boolean
    Dim i As Long
    
    ValidateTotAmtDisallUs36_6 = True
    TotAmtDisallUs36_6 = Sheet5.Range("sheet6.TotAmtDisallUs36").value
    
    If Len(TotAmtDisallUs36_6) > 14 Then
        msgbox6 ("* TotAmtDisallUs36 in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateTotAmtDisallUs36_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(TotAmtDisallUs36_6)
        If Not IsNumeric(TotAmtDisallUs36_6) Then
            msgbox6 ("* TotAmtDisallUs36 in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateTotAmtDisallUs36_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidatePersonalExp_6() As Boolean
    Dim i As Long
    
    ValidatePersonalExp_6 = True
    PersonalExp_6 = Sheet5.Range("sheet6.PersonalExp").value
    
    If Len(PersonalExp_6) > 14 Then
        msgbox6 ("* PersonalExp in Sheet : PART A-OI    should be at most 14 digits ")
        ValidatePersonalExp_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(PersonalExp_6)
        If Not IsNumeric(PersonalExp_6) Then
            msgbox6 ("* PersonalExp in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidatePersonalExp_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidatePoliticPartyExp_6() As Boolean
    Dim i As Long
    
    ValidatePoliticPartyExp_6 = True
    PoliticPartyExp_6 = Sheet5.Range("sheet6.PoliticPartyExp").value
    BusOrProfessnExp = Sheet5.Range("sheet6.BusOrProfessnExp").value
    
    If Len(PoliticPartyExp_6) > 14 Then
        msgbox6 ("* PoliticPartyExp in Sheet : PART A-OI    should be at most 14 digits ")
        ValidatePoliticPartyExp_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(PoliticPartyExp_6)
        If Not IsNumeric(PoliticPartyExp_6) Then
            msgbox6 ("* PoliticPartyExp in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidatePoliticPartyExp_6 = False
            Exit Function
        End If
    Next
End Function


Function ValidateLawVoilatPenalExp_6() As Boolean
    Dim i As Long
    
    ValidateLawVoilatPenalExp_6 = True
    LawVoilatPenalExp_6 = Sheet5.Range("sheet6.LawVoilatPenalExp").value
    
    If Len(LawVoilatPenalExp_6) > 14 Then
        msgbox6 ("* LawVoilatPenalExp in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateLawVoilatPenalExp_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(LawVoilatPenalExp_6)
        If Not IsNumeric(LawVoilatPenalExp_6) Then
            msgbox6 ("* LawVoilatPenalExp in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateLawVoilatPenalExp_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateOthPenalFineExp_6() As Boolean
    Dim i As Long
    
    ValidateOthPenalFineExp_6 = True
    OthPenalFineExp_6 = Sheet5.Range("sheet6.OthPenalFineExp").value
    
    If Len(OthPenalFineExp_6) > 14 Then
        msgbox6 ("* OthPenalFineExp in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateOthPenalFineExp_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(OthPenalFineExp_6)
        If Not IsNumeric(OthPenalFineExp_6) Then
            msgbox6 ("* OthPenalFineExp in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateOthPenalFineExp_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateOffenceExp_6() As Boolean
    Dim i As Long
    
    ValidateOffenceExp_6 = True
    OffenceExp_6 = Sheet5.Range("sheet6.OffenceExp").value
    
    If Len(OffenceExp_6) > 14 Then
        msgbox6 ("* OffenceExp in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateOffenceExp_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(OffenceExp_6)
        If Not IsNumeric(OffenceExp_6) Then
            msgbox6 ("* OffenceExp in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateOffenceExp_6 = False
            Exit Function
        End If
    Next
End Function


Function ValidateContigentLiability_6() As Boolean
    Dim i As Long
    
    ValidateContigentLiability_6 = True
    ContigentLiability_6 = Sheet5.Range("sheet6.ContigentLiability").value
    
    If Len(ContigentLiability_6) > 14 Then
        msgbox6 ("* ContigentLiability in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateContigentLiability_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(ContigentLiability_6)
        If Not IsNumeric(ContigentLiability_6) Then
            msgbox6 ("* ContigentLiability in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateContigentLiability_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateAmtNotPartTotInc_6() As Boolean
    Dim i As Long
    
    ValidateAmtNotPartTotInc_6 = True
    AmtNotPartTotInc_6 = Sheet5.Range("sheet6.AmtNotPartTotInc").value
    
    If Len(AmtNotPartTotInc_6) > 14 Then
        msgbox6 ("* AmtNotPartTotInc in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateAmtNotPartTotInc_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(AmtNotPartTotInc_6)
        If Not IsNumeric(AmtNotPartTotInc_6) Then
            msgbox6 ("* AmtNotPartTotInc in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateAmtNotPartTotInc_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateOthAmtNotAllowUs37_6() As Boolean
    Dim i As Long
    
    ValidateOthAmtNotAllowUs37_6 = True
    OthAmtNotAllowUs37_6 = Sheet5.Range("sheet6.OthAmtNotAllowUs37").value
    If Len(OthAmtNotAllowUs37_6) > 14 Then
        msgbox6 ("* OthAmtNotAllowUs37 in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateOthAmtNotAllowUs37_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(OthAmtNotAllowUs37_6)
        If Not IsNumeric(OthAmtNotAllowUs37_6) Then
            msgbox6 ("* OthAmtNotAllowUs37 in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateOthAmtNotAllowUs37_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotAmtDisallUs37_6() As Boolean
    Dim i As Long
    
    ValidateTotAmtDisallUs37_6 = True
    TotAmtDisallUs37_6 = Sheet5.Range("sheet6.TotAmtDisallUs37").value
    
    If Len(TotAmtDisallUs37_6) > 14 Then
        msgbox6 ("* TotAmtDisallUs37 in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateTotAmtDisallUs37_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(TotAmtDisallUs37_6)
        If Not IsNumeric(TotAmtDisallUs37_6) Then
            msgbox6 ("* TotAmtDisallUs37 in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateTotAmtDisallUs37_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateNonCompChapXVIIBAmt_6() As Boolean
    Dim i As Long
    
    ValidateNonCompChapXVIIBAmt_6 = True
    NonCompChapXVIIBAmt_6 = Sheet5.Range("sheet6.NonCompChapXVIIBAmt").value
    
    If Len(NonCompChapXVIIBAmt_6) > 14 Then
        msgbox6 ("* NonCompChapXVIIBAmt in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateNonCompChapXVIIBAmt_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(NonCompChapXVIIBAmt_6)
        If Not IsNumeric(NonCompChapXVIIBAmt_6) Then
            msgbox6 ("* NonCompChapXVIIBAmt in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateNonCompChapXVIIBAmt_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateSTTAmt_6() As Boolean
    Dim i As Long
    
    ValidateSTTAmt_6 = True
    STTAmt_6 = Sheet5.Range("sheet6.STTAmt").value
    
    If Len(STTAmt_6) > 14 Then
        msgbox6 ("* STTAmt in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateSTTAmt_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(STTAmt_6)
        If Not IsNumeric(STTAmt_6) Then
            msgbox6 ("* STTAmt in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateSTTAmt_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateFBTAmt_6() As Boolean
    Dim i As Long
    
    ValidateFBTAmt_6 = True
    FBTAmt_6 = Sheet5.Range("sheet6.FBTAmt").value
    If Len(FBTAmt_6) > 14 Then
        msgbox6 ("* FBTAmt in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateFBTAmt_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(FBTAmt_6)
        If Not IsNumeric(FBTAmt_6) Then
            msgbox6 ("* FBTAmt in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateFBTAmt_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateTaxAmtOnProfits_6() As Boolean
    Dim i As Long
    
    ValidateTaxAmtOnProfits_6 = True
    TaxAmtOnProfits_6 = Sheet5.Range("sheet6.TaxAmtOnProfits").value
    If Len(TaxAmtOnProfits_6) > 14 Then
        msgbox6 ("* TaxAmtOnProfits in Sheet : PART A-OI    should be at most 14 digits ")
        ValidateTaxAmtOnProfits_6 = False
        Exit Function
    End If
    
    For i = 1 To Len(TaxAmtOnProfits_6)
        If Not IsNumeric(TaxAmtOnProfits_6) Then
            msgbox6 ("* TaxAmtOnProfits in Sheet : PART A-OI    must contain only digits from 0 to 9")
            ValidateTaxAmtOnProfits_6 = False
            Exit Function
        End If
    Next
End Function

Function ValidateWTAmt_7() As Boolean
    Dim i As Long
    
    ValidateWTAmt_7 = True
    WTAmt_7 = Sheet5.Range("sheet7.WTAmt").value
    
    If Len(WTAmt_7) > 14 Then
        msgbox7 ("* WTAmt in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateWTAmt_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(WTAmt_7)
        If Not IsNumeric(WTAmt_7) Then
            msgbox7 ("* WTAmt in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateWTAmt_7 = False
            Exit Function
        End If
    Next
End Function


Function ValidateNonComp40aibChapXVIIBAmt_7() As Boolean
    Dim i As Long
    
    ValidateNonComp40aibChapXVIIBAmt_7 = True
    NonComp40aibChapXVIIBAmt_7 = Sheet5.Range("sheet6.NonComp40aibChapXVIIBAmt").value
    
    If Len(NonComp40aibChapXVIIBAmt_7) > 14 Then
        msgbox7 ("* NonComp40aibChapXVIIBAmt in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateNonComp40aibChapXVIIBAmt_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(NonComp40aibChapXVIIBAmt_7)
        If Not IsNumeric(NonComp40aibChapXVIIBAmt_7) Then
            msgbox7 ("* NonComp40aibChapXVIIBAmt in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateNonComp40aibChapXVIIBAmt_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIntSalBonPartner_7() As Boolean
    Dim i As Long
    
    ValidateIntSalBonPartner_7 = True
    IntSalBonPartner_7 = Sheet5.Range("sheet7.IntSalBonPartner").value
    
    If Len(IntSalBonPartner_7) > 14 Then
        msgbox7 ("* IntSalBonPartner in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateIntSalBonPartner_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(IntSalBonPartner_7)
        If Not IsNumeric(IntSalBonPartner_7) Then
            msgbox7 ("* IntSalBonPartner in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateIntSalBonPartner_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateOthDisallow_7() As Boolean
    Dim i As Long
    
    ValidateOthDisallow_7 = True
    OthDisallow_7 = Sheet5.Range("sheet7.OthDisallow").value
    
    If Len(OthDisallow_7) > 14 Then
        msgbox7 ("* OthDisallow in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateOthDisallow_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(OthDisallow_7)
        If Not IsNumeric(OthDisallow_7) Then
            msgbox7 ("* OthDisallow in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateOthDisallow_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotAmtDisallUs40_7() As Boolean
    Dim i As Long
    
    ValidateTotAmtDisallUs40_7 = True
    TotAmtDisallUs40_7 = Sheet5.Range("sheet7.TotAmtDisallUs40").value
    
    If Len(TotAmtDisallUs40_7) > 14 Then
        msgbox7 ("* TotAmtDisallUs40 in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateTotAmtDisallUs40_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(TotAmtDisallUs40_7)
        If Not IsNumeric(TotAmtDisallUs40_7) Then
            msgbox7 ("* TotAmtDisallUs40 in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateTotAmtDisallUs40_7 = False
            Exit Function
        End If
    Next
End Function


Function ValidateAmtDisallUs40PyNowAll_7() As Boolean
    Dim i As Long
    
    ValidateAmtDisallUs40PyNowAll_7 = True
    AmtDisallUs40PyNowAll_7 = Sheet5.Range("sheet7.AmtDisallUs40PyNowAll").value
    
    If Len(AmtDisallUs40PyNowAll_7) > 14 Then
        msgbox7 ("* AmtDisallUs40PyNowAll in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateAmtDisallUs40PyNowAll_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(AmtDisallUs40PyNowAll_7)
        If Not IsNumeric(AmtDisallUs40PyNowAll_7) Then
            msgbox7 ("* AmtDisallUs40PyNowAll in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateAmtDisallUs40PyNowAll_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateAmtPaidUs40A2b_7() As Boolean
    Dim i As Long
    
    ValidateAmtPaidUs40A2b_7 = True
    AmtPaidUs40A2b_7 = Sheet5.Range("sheet7.AmtPaidUs40A2b").value
    
    If Len(AmtPaidUs40A2b_7) > 14 Then
        msgbox7 ("* AmtPaidUs40A2b in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateAmtPaidUs40A2b_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(AmtPaidUs40A2b_7)
        If Not IsNumeric(AmtPaidUs40A2b_7) Then
            msgbox7 ("* AmtPaidUs40A2b in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateAmtPaidUs40A2b_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateAmtGT20kCash_7() As Boolean
    Dim i As Long
    
    ValidateAmtGT20kCash_7 = True
    AmtGT20kCash_7 = Sheet5.Range("sheet7.AmtGT20kCash").value
    
    If Len(AmtGT20kCash_7) > 14 Then
        msgbox7 ("* AmtGT20kCash in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateAmtGT20kCash_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(AmtGT20kCash_7)
        If Not IsNumeric(AmtGT20kCash_7) Then
            msgbox7 ("* AmtGT20kCash in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateAmtGT20kCash_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateProvPmtGrat_7() As Boolean
    Dim i As Long
    
    ValidateProvPmtGrat_7 = True
    ProvPmtGrat_7 = Sheet5.Range("sheet7.ProvPmtGrat").value
    
    If Len(ProvPmtGrat_7) > 14 Then
        msgbox7 ("* ProvPmtGrat in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateProvPmtGrat_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(ProvPmtGrat_7)
        If Not IsNumeric(ProvPmtGrat_7) Then
            msgbox7 ("* ProvPmtGrat in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateProvPmtGrat_7 = False
            Exit Function
        End If
    Next
End Function


Function ValidateContToSetupTrust_7() As Boolean
    Dim i As Long
    
    ValidateContToSetupTrust_7 = True
    ContToSetupTrust_7 = Sheet5.Range("sheet7.ContToSetupTrust").value
    
    If Len(ContToSetupTrust_7) > 14 Then
        msgbox7 ("* ContToSetupTrust in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateContToSetupTrust_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(ContToSetupTrust_7)
        If Not IsNumeric(ContToSetupTrust_7) Then
            msgbox7 ("* ContToSetupTrust in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateContToSetupTrust_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateOthDisallow1_7() As Boolean
    Dim i As Long
    
    ValidateOthDisallow1_7 = True
    OthDisallow1_7 = Sheet5.Range("sheet7.OthDisallow1").value
    
    If Len(OthDisallow1_7) > 14 Then
        msgbox7 ("* OthDisallow1 in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateOthDisallow1_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(OthDisallow1_7)
        If Not IsNumeric(OthDisallow1_7) Then
            msgbox7 ("* OthDisallow1 in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateOthDisallow1_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotAmtDisallUs40A_7() As Boolean
    Dim i As Long
    
    ValidateTotAmtDisallUs40A_7 = True
    TotAmtDisallUs40A_7 = Sheet5.Range("sheet7.TotAmtDisallUs40A").value
    
    If Len(TotAmtDisallUs40A_7) > 14 Then
        msgbox7 ("* TotAmtDisallUs40A in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateTotAmtDisallUs40A_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(TotAmtDisallUs40A_7)
        If Not IsNumeric(TotAmtDisallUs40A_7) Then
            msgbox7 ("* TotAmtDisallUs40A in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateTotAmtDisallUs40A_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateTaxDutyCesAmt_7() As Boolean
    Dim i As Long
    
    ValidateTaxDutyCesAmt_7 = True
    TaxDutyCesAmt_7 = Sheet5.Range("sheet7.TaxDutyCesAmt").value
    
    If Len(TaxDutyCesAmt_7) > 14 Then
        msgbox7 ("* TaxDutyCesAmt in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateTaxDutyCesAmt_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(TaxDutyCesAmt_7)
        If Not IsNumeric(TaxDutyCesAmt_7) Then
            msgbox7 ("* TaxDutyCesAmt in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateTaxDutyCesAmt_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateContToEmpPFSFGF_7() As Boolean
    Dim i As Long
    
    ValidateContToEmpPFSFGF_7 = True
    ContToEmpPFSFGF_7 = Sheet5.Range("sheet7.ContToEmpPFSFGF").value
    
    If Len(ContToEmpPFSFGF_7) > 14 Then
        msgbox7 ("* ContToEmpPFSFGF in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateContToEmpPFSFGF_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(ContToEmpPFSFGF_7)
        If Not IsNumeric(ContToEmpPFSFGF_7) Then
            msgbox7 ("* ContToEmpPFSFGF in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateContToEmpPFSFGF_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateEmpBonusComm_7() As Boolean
    Dim i As Long
    
    ValidateEmpBonusComm_7 = True
    EmpBonusComm_7 = Sheet5.Range("sheet7.EmpBonusComm").value
    
    If Len(EmpBonusComm_7) > 14 Then
        msgbox7 ("* EmpBonusComm in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateEmpBonusComm_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(EmpBonusComm_7)
        If Not IsNumeric(EmpBonusComm_7) Then
            msgbox7 ("* EmpBonusComm in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateEmpBonusComm_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIntPayaleToFI_7() As Boolean
    Dim i As Long
    
    ValidateIntPayaleToFI_7 = True
    IntPayaleToFI_7 = Sheet5.Range("sheet7.IntPayaleToFI").value
    
    If Len(IntPayaleToFI_7) > 14 Then
        msgbox7 ("* IntPayaleToFI in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateIntPayaleToFI_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(IntPayaleToFI_7)
        If Not IsNumeric(IntPayaleToFI_7) Then
            msgbox7 ("* IntPayaleToFI in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateIntPayaleToFI_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIntPayaleToFISchBank_7() As Boolean
    Dim i As Long
    
    ValidateIntPayaleToFISchBank_7 = True
    IntPayaleToFISchBank_7 = Sheet5.Range("sheet7.IntPayaleToFISchBank").value
    
    If Len(IntPayaleToFISchBank_7) > 14 Then
        msgbox7 ("* IntPayaleToFISchBank in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateIntPayaleToFISchBank_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(IntPayaleToFISchBank_7)
        If Not IsNumeric(IntPayaleToFISchBank_7) Then
            msgbox7 ("* IntPayaleToFISchBank in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateIntPayaleToFISchBank_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateLeaveEncashPayable_7() As Boolean
    Dim i As Long
    
    ValidateLeaveEncashPayable_7 = True
    LeaveEncashPayable_7 = Sheet5.Range("sheet7.LeaveEncashPayable").value
    RailwayAssets_10_g = Sheet5.Range("sheet7.RailwayAssets").value
    If Len(LeaveEncashPayable_7) > 14 Then
        msgbox7 ("* LeaveEncashPayable in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateLeaveEncashPayable_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(LeaveEncashPayable_7)
        If Not IsNumeric(LeaveEncashPayable_7) Then
            msgbox7 ("* LeaveEncashPayable in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateLeaveEncashPayable_7 = False
            Exit Function
        End If
    Next
    
    
    If Len(RailwayAssets_10_g) > 14 Then
        msgbox7 ("* Railway Assets of Row 10gin Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateLeaveEncashPayable_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(RailwayAssets_10_g)
        If Not IsNumeric(RailwayAssets_10_g) Then
            msgbox7 ("* Railway Assets of Row 10g in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateLeaveEncashPayable_7 = False
            Exit Function
        End If
    Next
    
    
    
    
End Function

Function ValidateRailwayAssetsPayable_7() As Boolean
    Dim i As Long
    
    ValidateRailwayAssetsPayable_7 = True
    RailwayAssetsPayable_7 = Sheet5.Range("sheet7.RailwayAssetsPayable").value
    
    If Len(RailwayAssetsPayable_7) > 14 Then
        msgbox7 ("* RailwayAssetsPayable in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateRailwayAssetsPayable_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(RailwayAssetsPayable_7)
        If Not IsNumeric(RailwayAssetsPayable_7) Then
            msgbox7 ("* RailwayAssetsPayable in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateRailwayAssetsPayable_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotAmtUs43b_7() As Boolean
    Dim i As Long
    
    ValidateTotAmtUs43b_7 = True
    TotAmtUs43b_7 = Sheet5.Range("sheet7.TotAmtUs43b").value
    
    If Len(TotAmtUs43b_7) > 14 Then
        msgbox7 ("* TotAmtUs43b in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateTotAmtUs43b_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(TotAmtUs43b_7)
        If Not IsNumeric(TotAmtUs43b_7) Then
            msgbox7 ("* TotAmtUs43b in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateTotAmtUs43b_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateTaxDutyCesAmt1_7() As Boolean
    Dim i As Long
    
    ValidateTaxDutyCesAmt1_7 = True
    TaxDutyCesAmt1_7 = Sheet5.Range("sheet7.TaxDutyCesAmt1").value
    
    If Len(TaxDutyCesAmt1_7) > 14 Then
        msgbox7 ("* TaxDutyCesAmt1 in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateTaxDutyCesAmt1_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(TaxDutyCesAmt1_7)
        If Not IsNumeric(TaxDutyCesAmt1_7) Then
            msgbox7 ("* TaxDutyCesAmt1 in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateTaxDutyCesAmt1_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateContToEmpPFSFGF1_7() As Boolean
    Dim i As Long
    
    ValidateContToEmpPFSFGF1_7 = True
    ContToEmpPFSFGF1_7 = Sheet5.Range("sheet7.ContToEmpPFSFGF1").value
    
    If Len(ContToEmpPFSFGF1_7) > 14 Then
        msgbox7 ("* ContToEmpPFSFGF1 in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateContToEmpPFSFGF1_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(ContToEmpPFSFGF1_7)
        If Not IsNumeric(ContToEmpPFSFGF1_7) Then
            msgbox7 ("* ContToEmpPFSFGF1 in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateContToEmpPFSFGF1_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateEmpBonusComm1_7() As Boolean
    Dim i As Long
    
    ValidateEmpBonusComm1_7 = True
    EmpBonusComm1_7 = Sheet5.Range("sheet7.EmpBonusComm1").value
    
    If Len(EmpBonusComm1_7) > 14 Then
        msgbox7 ("* EmpBonusComm1 in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateEmpBonusComm1_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(EmpBonusComm1_7)
        If Not IsNumeric(EmpBonusComm1_7) Then
            msgbox7 ("* EmpBonusComm1 in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateEmpBonusComm1_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIntPayaleToFI1_7() As Boolean
    Dim i As Long
    
    ValidateIntPayaleToFI1_7 = True
    IntPayaleToFI1_7 = Sheet5.Range("sheet7.IntPayaleToFI1").value
    
    If Len(IntPayaleToFI1_7) > 14 Then
        msgbox7 ("* IntPayaleToFI1 in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateIntPayaleToFI1_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(IntPayaleToFI1_7)
        If Not IsNumeric(IntPayaleToFI1_7) Then
            msgbox7 ("* IntPayaleToFI1 in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateIntPayaleToFI1_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIntPayaleToFISchBank1_7() As Boolean
    Dim i As Long
    
    ValidateIntPayaleToFISchBank1_7 = True
    IntPayaleToFISchBank1_7 = Sheet5.Range("sheet7.IntPayaleToFISchBank1").value
    
    If Len(IntPayaleToFISchBank1_7) > 14 Then
        msgbox7 ("* IntPayaleToFISchBank1 in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateIntPayaleToFISchBank1_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(IntPayaleToFISchBank1_7)
        If Not IsNumeric(IntPayaleToFISchBank1_7) Then
            msgbox7 ("* IntPayaleToFISchBank1 in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateIntPayaleToFISchBank1_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateLeaveEncashPayable1_7() As Boolean
    Dim i As Long
    
    ValidateLeaveEncashPayable1_7 = True
    LeaveEncashPayable1_7 = Sheet5.Range("sheet7.LeaveEncashPayable1").value
    
    If Len(LeaveEncashPayable1_7) > 14 Then
        msgbox7 ("* LeaveEncashPayable1 in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateLeaveEncashPayable1_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(LeaveEncashPayable1_7)
        If Not IsNumeric(LeaveEncashPayable1_7) Then
            msgbox7 ("* LeaveEncashPayable1 in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateLeaveEncashPayable1_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotAmtUs43b1_7() As Boolean
    Dim i As Long
    
    ValidateTotAmtUs43b1_7 = True
    TotAmtUs43b1_7 = Sheet5.Range("sheet7.TotAmtUs43b1").value
    
    If Len(TotAmtUs43b1_7) > 14 Then
        msgbox7 ("* TotAmtUs43b1 in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateTotAmtUs43b1_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(TotAmtUs43b1_7)
        If Not IsNumeric(TotAmtUs43b1_7) Then
            msgbox7 ("* TotAmtUs43b1 in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateTotAmtUs43b1_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateUnionExciseDuty_7() As Boolean
    Dim i As Long
    
    ValidateUnionExciseDuty_7 = True
    UnionExciseDuty_7 = Sheet5.Range("sheet7.UnionExciseDuty").value
    
    If Len(UnionExciseDuty_7) > 14 Then
        msgbox7 ("* UnionExciseDuty in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateUnionExciseDuty_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(UnionExciseDuty_7)
        If Not IsNumeric(UnionExciseDuty_7) Then
            msgbox7 ("* UnionExciseDuty in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateUnionExciseDuty_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateServiceTax_7() As Boolean
    Dim i As Long
    
    ValidateServiceTax_7 = True
    ServiceTax_7 = Sheet5.Range("sheet7.ServiceTax").value
    
    If Len(ServiceTax_7) > 14 Then
        msgbox7 ("* ServiceTax in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateServiceTax_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(ServiceTax_7)
        If Not IsNumeric(ServiceTax_7) Then
            msgbox7 ("* ServiceTax in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateServiceTax_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateVATorSaleTax_7() As Boolean
    Dim i As Long
    
    ValidateVATorSaleTax_7 = True
    VATorSaleTax_7 = Sheet5.Range("sheet7.VATorSaleTax").value
    
    
     If Len(VATorSaleTax_7) > 14 Then
        msgbox7 ("* VATorSaleTax in Sheet :PART A  OI  should not greater than 14 digits ")
        ValidateVATorSaleTax_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(VATorSaleTax_7)
        If Not IsNumeric(VATorSaleTax_7) Then
            msgbox7 ("* VATorSaleTax in Sheet :PART A  OI   must contain only digits from 0 to 9")
            ValidateVATorSaleTax_7 = False
            Exit Function
        End If
    Next
    
   CGST_7 = Sheet5.Range("Sheet7.CGST").value
   SGST_7 = Sheet5.Range("Sheet7.SGST").value
   IGST_7 = Sheet5.Range("Sheet7.IGST").value
   UTGST_7 = Sheet5.Range("Sheet7.UTGST").value
    
    If Len(CGST_7) > 14 Then
        msgbox7 ("* CGST in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateVATorSaleTax_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(CGST_7)
        If Not IsNumeric(CGST_7) Then
            msgbox7 ("* CGST in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateVATorSaleTax_7 = False
            Exit Function
        End If
    Next
    
        If Len(SGST_7) > 14 Then
        msgbox7 ("* SGST in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateVATorSaleTax_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(SGST_7)
        If Not IsNumeric(SGST_7) Then
            msgbox7 ("* SGST in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateVATorSaleTax_7 = False
            Exit Function
        End If
    Next
    
    If Len(IGST_7) > 14 Then
        msgbox7 ("* IGST in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateVATorSaleTax_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(IGST_7)
        If Not IsNumeric(IGST_7) Then
            msgbox7 ("* IGST in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateVATorSaleTax_7 = False
            Exit Function
        End If
    Next
    
    If Len(UTGST_7) > 14 Then
        msgbox7 ("* UTGST in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateVATorSaleTax_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(UTGST_7)
        If Not IsNumeric(UTGST_7) Then
            msgbox7 ("* UTGST in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateVATorSaleTax_7 = False
            Exit Function
        End If
    Next
    
    
    
End Function

Function ValidateOthDutyTaxCess_7() As Boolean
    Dim i As Long
    
    ValidateOthDutyTaxCess_7 = True
    OthDutyTaxCess_7 = Sheet5.Range("sheet7.OthDutyTaxCess").value
    
    If Len(OthDutyTaxCess_7) > 14 Then
        msgbox7 ("* OthDutyTaxCess in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateOthDutyTaxCess_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(OthDutyTaxCess_7)
        If Not IsNumeric(OthDutyTaxCess_7) Then
            msgbox7 ("* OthDutyTaxCess in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateOthDutyTaxCess_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotExciseCustomsVAT_7() As Boolean
    Dim i As Long
    
    ValidateTotExciseCustomsVAT_7 = True
    TotExciseCustomsVAT_7 = Sheet5.Range("sheet7.TotExciseCustomsVAT").value
    
    If Len(TotExciseCustomsVAT_7) > 14 Then
        msgbox7 ("* TotExciseCustomsVAT in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateTotExciseCustomsVAT_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(TotExciseCustomsVAT_7)
        If Not IsNumeric(TotExciseCustomsVAT_7) Then
            msgbox7 ("* TotExciseCustomsVAT in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateTotExciseCustomsVAT_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidateDeemedProfUs33ABs_7() As Boolean
    Dim i As Long
    
    ValidateDeemedProfUs33ABs_7 = True
    DeemedProfUs33ABs_7 = Sheet5.Range("sheet7.DeemedProfUs33ABs").value
    DeemedProfUs33AB_7 = Sheet5.Range("sheet7.DeemedProfUs33AB").value
    DeemedProfUs33ABA_7 = Sheet5.Range("sheet7.DeemedProfUs33ABA").value
    
    If Len(DeemedProfUs33ABs_7) > 14 Then
        msgbox7 ("* DeemedProfUs33ABs in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateDeemedProfUs33ABs_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(DeemedProfUs33ABs_7)
        If Not IsNumeric(DeemedProfUs33ABs_7) Then
            msgbox7 ("* DeemedProfUs33ABs in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateDeemedProfUs33ABs_7 = False
            Exit Function
        End If
    Next


    If Len(DeemedProfUs33AB_7) > 14 Then
        msgbox7 ("* DeemedProfUs33AB in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateDeemedProfUs33ABs_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(DeemedProfUs33AB_7)
        If Not IsNumeric(DeemedProfUs33AB_7) Then
            msgbox7 ("* DeemedProfUs33AB in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateDeemedProfUs33ABs_7 = False
            Exit Function
        End If
    Next


    If Len(DeemedProfUs33ABA_7) > 14 Then
        msgbox7 ("* DeemedProfUs33ABA in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateDeemedProfUs33ABs_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(DeemedProfUs33ABA_7)
        If Not IsNumeric(DeemedProfUs33ABA_7) Then
            msgbox7 ("* DeemedProfUs33ABA in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateDeemedProfUs33ABs_7 = False
            Exit Function
        End If
    Next
End Function


Function ValidateProfTaxAmtUs41_7() As Boolean
    Dim i As Long
    
    ValidateProfTaxAmtUs41_7 = True
    ProfTaxAmtUs41_7 = Sheet5.Range("sheet7.ProfTaxAmtUs41").value
    
    If Len(ProfTaxAmtUs41_7) > 14 Then
        msgbox7 ("* ProfTaxAmtUs41 in Sheet :PART A  OI    should not greater than 14 digits ")
        ValidateProfTaxAmtUs41_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(ProfTaxAmtUs41_7)
        If Not IsNumeric(ProfTaxAmtUs41_7) Then
            msgbox7 ("* ProfTaxAmtUs41 in Sheet :PART A  OI    must contain only digits from 0 to 9")
            ValidateProfTaxAmtUs41_7 = False
            Exit Function
        End If
    Next
End Function

Function ValidatePriorAmtIncCrDrPL_7() As Boolean
    Dim i As Long
    
    ValidatePriorAmtIncCrDrPL_7 = True
    PriorAmtIncCrDrPL_7 = Sheet5.Range("sheet7.PriorAmtIncCrDrPL").value
    
    If Len(PriorAmtIncCrDrPL_7) > 14 Then
        msgbox7 ("* PriorAmtIncCrDrPL in Sheet :PART A  OI should not greater than 14 digits ")
        ValidatePriorAmtIncCrDrPL_7 = False
        Exit Function
    End If
    
    For i = 1 To Len(PriorAmtIncCrDrPL_7)
        If Not IsNumeric(PriorAmtIncCrDrPL_7) Then
            msgbox7 ("* PriorAmtIncCrDrPL in Sheet :PART A  OI must contain only digits from 0 to 9")
            ValidatePriorAmtIncCrDrPL_7 = False
            Exit Function
        End If
    Next
End Function


Function msgbox5(strmsg As String) As String
    msgValidateSheet5 = msgValidateSheet5 & strmsg & Chr(13)
End Function

Function msgbox6(strmsg As String) As String
    msgValidateSheet6 = msgValidateSheet6 & strmsg & Chr(13)
End Function

Function msgbox7(strmsg As String) As String
    msgValidateSheet7 = msgValidateSheet7 & strmsg & Chr(13)
End Function

Function DefaultMethodOfAcct_5() As String
    DefaultMethodOfAcct_5 = "MERC"
End Function

Function DefaultChangeInAcctMethFlg_5() As String
    DefaultChangeInAcctMethFlg_5 = "N"
End Function

Function DefaultProfDeviatDueAcctMeth_5() As String
    DefaultProfDeviatDueAcctMeth_5 = "0"
End Function
Function DefaultProfDeviatDueAcctMeth_6() As String
    DefaultProfDeviatDueAcctMeth_6 = "0"
End Function

Function DefaultValRawMaterial_5() As String
    DefaultValRawMaterial_5 = "1"
End Function

Function DefaultValFinishedGoods_5() As String
    DefaultValFinishedGoods_5 = "1"
End Function

Function DefaultChngStockValMetFlg_6() As String
    DefaultChngStockValMetFlg_6 = "N"
End Function

Function DefaultEffectOnPL_6() As String
    DefaultEffectOnPL_6 = "0"
End Function
Function DefaultEffectOnPL_6_4e() As String
    DefaultEffectOnPL_6_4e = "0"
End Function
Function DefaultSection28Items_6() As String
    DefaultSection28Items_6 = "0"
End Function

Function DefaultProformaCreditsDue_6() As String
    DefaultProformaCreditsDue_6 = "0"
End Function

Function DefaultPrevYrEscalClaim_6() As String
    DefaultPrevYrEscalClaim_6 = "0"
End Function

Function DefaultOthItemInc_6() As String
    DefaultOthItemInc_6 = "0"
End Function

Function DefaultCapReceipt_6() As String
    DefaultCapReceipt_6 = "0"
End Function

Function DefaultTotNoCredToPLAmt_6() As String
    DefaultTotNoCredToPLAmt_6 = "0"
End Function

Function DefaultStkInsurPrem_6() As String
    DefaultStkInsurPrem_6 = "0"
End Function

Function DefaultEmpHealthInsurPrem_6() As String
    DefaultEmpHealthInsurPrem_6 = "0"
End Function

Function DefaultEmpBonusCommSum_6() As String
    DefaultEmpBonusCommSum_6 = "0"
End Function

Function DefaultIntOnBorrCap_6() As String
    DefaultIntOnBorrCap_6 = "0"
End Function

Function DefaultZeroCoupBondDisc_6() As String
    DefaultZeroCoupBondDisc_6 = "0"
End Function

Function DefaultRecogPFContribAmt_6() As String
    DefaultRecogPFContribAmt_6 = "0"
End Function

Function DefaultAppSuperAnnFundAmt_6() As String
    DefaultAppSuperAnnFundAmt_6 = "0"
End Function

Function DefaultAppGratFundAmt_6() As String
    DefaultAppGratFundAmt_6 = "0"
End Function

Function DefaultOthFundAmt_6() As String
    DefaultOthFundAmt_6 = "0"
End Function

Function DefaultBadDebtDoubtAmt_6() As String
    DefaultBadDebtDoubtAmt_6 = "0"
End Function

Function DefaultBadDebtDoubtProvn_6() As String
    DefaultBadDebtDoubtProvn_6 = "0"
End Function

Function DefaultSpecResrvTranfr_6() As String
    DefaultSpecResrvTranfr_6 = "0"
End Function

Function DefaultFamPlanPromoExp_6() As String
    DefaultFamPlanPromoExp_6 = "0"
End Function

Function DefaultEmpContributionCredits_6() As String
    DefaultEmpContributionCredits_6 = "0"
End Function

Function DefaultOthDisallowances_6() As String
    DefaultOthDisallowances_6 = "0"
End Function

Function DefaultTotAmtDisallUs36_6() As String
    DefaultTotAmtDisallUs36_6 = "0"
End Function

Function DefaultPersonalExp_6() As String
    DefaultPersonalExp_6 = "0"
End Function

Function DefaultPoliticPartyExp_6() As String
    DefaultPoliticPartyExp_6 = "0"
End Function

Function DefaultLawVoilatPenalExp_6() As String
    DefaultLawVoilatPenalExp_6 = "0"
End Function

Function DefaultOthPenalFineExp_6() As String
    DefaultOthPenalFineExp_6 = "0"
End Function

Function DefaultOffenceExp_6() As String
    DefaultOffenceExp_6 = "0"
End Function

Function DefaultContigentLiability_6() As String
    DefaultContigentLiability_6 = "0"
End Function

Function DefaultAmtNotPartTotInc_6() As String
    DefaultAmtNotPartTotInc_6 = "0"
End Function

Function DefaultOthAmtNotAllowUs37_6() As String
    DefaultOthAmtNotAllowUs37_6 = "0"
End Function

Function DefaultTotAmtDisallUs37_6() As String
    DefaultTotAmtDisallUs37_6 = "0"
End Function

Function DefaultNonCompChapXVIIBAmt_6() As String
    DefaultNonCompChapXVIIBAmt_6 = "0"
End Function

Function DefaultSTTAmt_6() As String
    DefaultSTTAmt_6 = "0"
End Function

Function DefaultFBTAmt_6() As String
    DefaultFBTAmt_6 = "0"
End Function

Function DefaultTaxAmtOnProfits_6() As String
    DefaultTaxAmtOnProfits_6 = "0"
End Function


