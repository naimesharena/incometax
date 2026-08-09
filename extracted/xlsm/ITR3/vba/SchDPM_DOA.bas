Attribute VB_Name = "SchDPM_DOA"
Option Explicit

Public WDVFirstDay_DPM15 As Variant
Public AdditionsGrThan180Days_DPM15 As Variant
Public RealizationTotalPeriod_DPM15 As Variant
Public FullRateDeprAmt_DPM15 As Variant
Public AdditionsLessThan180Days_DPM15 As Variant
Public RealizationPeriodDuringYear_DPM15 As Variant
Public HalfRateDeprAmt_DPM15 As Variant
Public DepreciationAtFullRate_DPM15 As Variant
Public DepreciationAtHalfRate_DPM15 As Variant
Public AddlnDeprOnGT180DayAdditions_DPM15 As Variant
Public AddlnDeprDuringYearAdditions_DPM15 As Variant
Public AddlnDeprlessthan180days_DPM15 As Variant
Public TotalDepreciation_DPM15 As Variant
Public ExpdrOnTrforSaleAsset_DPM15 As Variant
Public CapGainUs50_DPM15 As Variant
Public WDVLastDay_DPM15 As Variant
Public WDVFirstDay_DPM30 As Variant
Public AdditionsGrThan180Days_DPM30 As Variant
Public RealizationTotalPeriod_DPM30 As Variant
Public FullRateDeprAmt_DPM30 As Variant
Public AdditionsLessThan180Days_DPM30 As Variant
Public RealizationPeriodDuringYear_DPM30 As Variant
Public HalfRateDeprAmt_DPM30 As Variant
Public DepreciationAtFullRate_DPM30 As Variant
Public DepreciationAtHalfRate_DPM30 As Variant
Public AddlnDeprOnGT180DayAdditions_DPM30 As Variant
Public AddlnDeprDuringYearAdditions_DPM30 As Variant
Public AddlnDeprlessthan180days_DPM30 As Variant
Public TotalDepreciation_DPM30 As Variant
Public ExpdrOnTrforSaleAsset_DPM30 As Variant
Public CapGainUs50_DPM30 As Variant
Public WDVLastDay_DPM30 As Variant
Public WDVFirstDay_DPM40 As Variant
'Public WDVFirstDay2_DPM40 As Variant
Public AdditionsGrThan180Days_DPM40 As Variant
Public RealizationTotalPeriod_DPM40 As Variant
Public FullRateDeprAmt_DPM40 As Variant
Public AdditionsLessThan180Days_DPM40 As Variant
Public RealizationPeriodDuringYear_DPM40 As Variant
Public HalfRateDeprAmt_DPM40 As Variant
Public DepreciationAtFullRate_DPM40 As Variant
Public DepreciationAtHalfRate_DPM40 As Variant
Public AddlnDeprOnGT180DayAdditions_DPM40 As Variant
Public AddlnDeprDuringYearAdditions_DPM40 As Variant
Public AddlnDeprlessthan180days_DPM40 As Variant
Public TotalDepreciation_DPM40 As Variant
Public ExpdrOnTrforSaleAsset_DPM40 As Variant
Public CapGainUs50_DPM40 As Variant
Public CapGainUs50_DPM45 As Variant
Public WDVLastDay_DPM40 As Variant


Public WDVFirstDay_DPM45 As Variant
Public AdditionsGrThan180Days_DPM45 As Variant
Public RealizationTotalPeriod_DPM45 As Variant
Public FullRateDeprAmt_DPM45 As Variant
Public AdditionsLessThan180Days_DPM45 As Variant
Public RealizationPeriodDuringYear_DPM45 As Variant
Public HalfRateDeprAmt_DPM45 As Variant
Public DepreciationAtFullRate_DPM45 As Variant
Public DepreciationAtHalfRate_DPM45 As Variant
Public AddlnDeprOnGT180DayAdditions_DPM45 As Variant
Public AddlnDeprDuringYearAdditions_DPM45 As Variant
Public AddlnDeprlessthan180days_DPM45 As Variant
Public TotalDepreciation_DPM45 As Variant
Public ExpdrOnTrforSaleAsset_DPM45 As Variant
'Public CapGainUs50_DPM45 As Variant
Public WDVLastDay_DPM45 As Variant
Public NetaggregateDep_DPM45 As Variant
Public PropaggregateDep_DPM45 As Variant
Public Depreciation_Sec38_DPM45 As Variant

Public WDVFirstDay_DPM50 As Variant
Public AdditionsGrThan180Days_DPM50 As Variant
Public RealizationTotalPeriod_DPM50 As Variant
Public FullRateDeprAmt_DPM50 As Variant
Public AdditionsLessThan180Days_DPM50 As Variant
Public RealizationPeriodDuringYear_DPM50 As Variant
Public HalfRateDeprAmt_DPM50 As Variant
Public DepreciationAtFullRate_DPM50 As Variant
Public DepreciationAtHalfRate_DPM50 As Variant
Public AddlnDeprOnGT180DayAdditions_DPM50 As Variant
Public AddlnDeprDuringYearAdditions_DPM50 As Variant
Public AddlnDeprlessthan180days_DPM50 As Variant
Public TotalDepreciation_DPM50 As Variant
Public ExpdrOnTrforSaleAsset_DPM50 As Variant
Public CapGainUs50_DPM50 As Variant
Public WDVLastDay_DPM50 As Variant
Public WDVFirstDay_DPM60 As Variant
Public AdditionsGrThan180Days_DPM60 As Variant
Public RealizationTotalPeriod_DPM60 As Variant
Public FullRateDeprAmt_DPM60 As Variant
Public AdditionsLessThan180Days_DPM60 As Variant
Public RealizationPeriodDuringYear_DPM60 As Variant
Public HalfRateDeprAmt_DPM60 As Variant
Public DepreciationAtFullRate_DPM60 As Variant
Public DepreciationAtHalfRate_DPM60 As Variant
Public AddlnDeprOnGT180DayAdditions_DPM60 As Variant
Public AddlnDeprDuringYearAdditions_DPM60 As Variant
Public AddlnDeprlessthan180days_DPM60 As Variant
Public TotalDepreciation_DPM60 As Variant
Public ExpdrOnTrforSaleAsset_DPM60 As Variant
Public CapGainUs50_DPM60 As Variant
Public WDVLastDay_DPM60 As Variant
Public WDVFirstDay_DPM80 As Variant
Public AdditionsGrThan180Days_DPM80 As Variant
Public RealizationTotalPeriod_DPM80 As Variant
Public FullRateDeprAmt_DPM80 As Variant
Public AdditionsLessThan180Days_DPM80 As Variant
Public RealizationPeriodDuringYear_DPM80 As Variant
Public HalfRateDeprAmt_DPM80 As Variant
Public DepreciationAtFullRate_DPM80 As Variant
Public DepreciationAtHalfRate_DPM80 As Variant
Public AddlnDeprOnGT180DayAdditions_DPM80 As Variant
Public AddlnDeprDuringYearAdditions_DPM80 As Variant
Public AddlnDeprlessthan180days_DPM80 As Variant
Public TotalDepreciation_DPM80 As Variant
Public ExpdrOnTrforSaleAsset_DPM80 As Variant
Public CapGainUs50_DPM80 As Variant
Public WDVLastDay_DPM80 As Variant
Public WDVFirstDay_DPM100 As Variant
Public AdditionsGrThan180Days_DPM100 As Variant
Public RealizationTotalPeriod_DPM100 As Variant
Public FullRateDeprAmt_DPM100 As Variant
Public AdditionsLessThan180Days_DPM100 As Variant
Public RealizationPeriodDuringYear_DPM100 As Variant
Public HalfRateDeprAmt_DPM100 As Variant

Public Depreciation_Sec38_DPM30 As Variant
Public NetaggregateDep_DPM30  As Variant
Public PropaggregateDep_DPM30 As Variant
Public Depreciation_Sec38_DPM15 As Variant
Public NetaggregateDep_DPM15  As Variant
Public PropaggregateDep_DPM15 As Variant
Public Depreciation_Sec38_DPM40 As Variant
Public NetaggregateDep_DPM40  As Variant
Public PropaggregateDep_DPM40 As Variant


Public DepreciationAtFullRate_DPM100 As Variant
Public DepreciationAtHalfRate_DPM100 As Variant
Public AddlnDeprOnGT180DayAdditions_DPM100 As Variant
Public AddlnDeprDuringYearAdditions_DPM100 As Variant
Public AddlnDeprlessthan180days_DPM100 As Variant
Public TotalDepreciation_DPM100 As Variant
Public ExpdrOnTrforSaleAsset_DPM100 As Variant
Public CapGainUs50_DPM100 As Variant
Public WDVLastDay_DPM100 As Variant
Public WDVFirstDay_DAOL0 As Variant
Public WDVLastDay_DAOL0 As Variant
Public WDVFirstDay_DAOB5 As Variant
Public AdditionsGrThan180Days_DAOB5 As Variant
Public RealizationTotalPeriod_DAOB5 As Variant
Public FullRateDeprAmt_DAOB5 As Variant
Public AdditionsLessThan180Days_DAOB5 As Variant
Public RealizationPeriodDuringYear_DAOB5 As Variant
Public HalfRateDeprAmt_DAOB5 As Variant
Public DepreciationAtFullRate_DAOB5 As Variant
Public DepreciationAtHalfRate_DAOB5 As Variant
Public AddlnDeprOnGT180DayAdditions_DAOB5 As Variant
Public AddlnDeprDuringYearAdditions_DAOB5 As Variant
Public TotalDepreciation_DAOB5 As Variant
Public ExpdrOnTrforSaleAsset_DAOB5 As Variant
Public CapGainUs50_DAOB5 As Variant
Public WDVLastDay_DAOB5 As Variant
Public WDVFirstDay_DAOB10 As Variant
Public AdditionsGrThan180Days_DAOB10 As Variant
Public RealizationTotalPeriod_DAOB10 As Variant
Public FullRateDeprAmt_DAOB10 As Variant
Public AdditionsLessThan180Days_DAOB10 As Variant
Public RealizationPeriodDuringYear_DAOB10 As Variant
Public HalfRateDeprAmt_DAOB10 As Variant
Public DepreciationAtFullRate_DAOB10 As Variant
Public DepreciationAtHalfRate_DAOB10 As Variant
Public AddlnDeprOnGT180DayAdditions_DAOB10 As Variant
Public AddlnDeprDuringYearAdditions_DAOB10 As Variant
Public TotalDepreciation_DAOB10 As Variant
Public ExpdrOnTrforSaleAsset_DAOB10 As Variant
Public CapGainUs50_DAOB10 As Variant
Public WDVLastDay_DAOB10 As Variant
Public WDVFirstDay_DAOB100 As Variant
Public AdditionsGrThan180Days_DAOB100 As Variant
Public RealizationTotalPeriod_DAOB100 As Variant
Public FullRateDeprAmt_DAOB100 As Variant
Public AdditionsLessThan180Days_DAOB100 As Variant
Public RealizationPeriodDuringYear_DAOB100 As Variant
Public HalfRateDeprAmt_DAOB100 As Variant
Public DepreciationAtFullRate_DAOB100 As Variant
Public DepreciationAtHalfRate_DAOB100 As Variant
Public AddlnDeprOnGT180DayAdditions_DAOB100 As Variant
Public AddlnDeprDuringYearAdditions_DAOB100 As Variant
Public TotalDepreciation_DAOB100 As Variant
Public ExpdrOnTrforSaleAsset_DAOB100 As Variant
Public CapGainUs50_DAOB100 As Variant
Public WDVLastDay_DAOB100 As Variant
Public WDVFirstDay_DAOF10 As Variant
Public AdditionsGrThan180Days_DAOF10 As Variant
Public RealizationTotalPeriod_DAOF10 As Variant
Public FullRateDeprAmt_DAOF10 As Variant
Public AdditionsLessThan180Days_DAOF10 As Variant
Public RealizationPeriodDuringYear_DAOF10 As Variant
Public HalfRateDeprAmt_DAOF10 As Variant
Public DepreciationAtFullRate_DAOF10 As Variant
Public DepreciationAtHalfRate_DAOF10 As Variant
Public AddlnDeprOnGT180DayAdditions_DAOF10 As Variant
Public AddlnDeprDuringYearAdditions_DAOF10 As Variant
Public TotalDepreciation_DAOF10 As Variant
Public ExpdrOnTrforSaleAsset_DAOF10 As Variant
Public CapGainUs50_DAOF10 As Variant
Public WDVLastDay_DAOF10 As Variant
Public WDVFirstDay_DAOI25 As Variant
Public AdditionsGrThan180Days_DAOI25 As Variant
Public RealizationTotalPeriod_DAOI25 As Variant
Public FullRateDeprAmt_DAOI25 As Variant
Public AdditionsLessThan180Days_DAOI25 As Variant
Public RealizationPeriodDuringYear_DAOI25 As Variant
Public HalfRateDeprAmt_DAOI25 As Variant
Public DepreciationAtFullRate_DAOI25 As Variant
Public DepreciationAtHalfRate_DAOI25 As Variant
Public AddlnDeprOnGT180DayAdditions_DAOI25 As Variant
Public AddlnDeprDuringYearAdditions_DAOI25 As Variant
Public TotalDepreciation_DAOI25 As Variant
Public ExpdrOnTrforSaleAsset_DAOI25 As Variant
Public CapGainUs50_DAOI25 As Variant
Public WDVLastDay_DAOI25 As Variant
Public WDVFirstDay_DAOS20 As Variant
Public AdditionsGrThan180Days_DAOS20 As Variant
Public RealizationTotalPeriod_DAOS20 As Variant
Public FullRateDeprAmt_DAOS20 As Variant
Public AdditionsLessThan180Days_DAOS20 As Variant
Public RealizationPeriodDuringYear_DAOS20 As Variant
Public HalfRateDeprAmt_DAOS20 As Variant
Public DepreciationAtFullRate_DAOS20 As Variant
Public DepreciationAtHalfRate_DAOS20 As Variant
Public AddlnDeprOnGT180DayAdditions_DAOS20 As Variant
Public AddlnDeprDuringYearAdditions_DAOS20 As Variant
Public TotalDepreciation_DAOS20 As Variant
Public ExpdrOnTrforSaleAsset_DAOS20 As Variant
Public CapGainUs50_DAOS20 As Variant
Public WDVLastDay_DAOS20 As Variant

Public Depreciation_Sec38_DAOB10  As Variant
Public NetaggregateDep_DAOB10  As Variant
Public PropaggregateDep_DAOB10  As Variant
Public Depreciation_Sec38_DAOB05  As Variant
Public NetaggregateDep_DAOB05  As Variant
Public PropaggregateDep_DAOB05  As Variant
Public Depreciation_Sec38_DAOB40  As Variant
Public NetaggregateDep_DAOB40  As Variant
Public PropaggregateDep_DAOB40  As Variant
Public Depreciation_Sec38_DAOF10  As Variant
Public NetaggregateDep_DAOF10  As Variant
Public PropaggregateDep_DAOF10  As Variant
Public Depreciation_Sec38_DAOI25  As Variant
Public NetaggregateDep_DAOI25  As Variant
Public PropaggregateDep_DAOI25 As Variant
Public Depreciation_Sec38_DAOS20  As Variant
Public NetaggregateDep_DAOS20  As Variant
Public PropaggregateDep_DAOS20 As Variant
'Malli
Public Adjmt2NDprov3section115BAC_DPM15 As Variant
Public Adjmt2NDprov3section115BAC_DPM30 As Variant
Public Adjmt2NDprov3section115BAC_DPM40 As Variant
Public Adjmt2NDprov3section115BAC_DPM45 As Variant

'------
Public msgValidateSheetDPM15 As Variant
Dim msgValidateSheetDAOB5 As Variant


Sub validateDPMDOAClick()
ValidateSheetDPM_DOA
fmsgboxStatus "sheet DPM-DOA is OK" ', vbOKOnly
End Sub

Sub ValidateSheetDPM_DOA()
    If Not ValidatesheetDPM15 Then
        Sheet10.Activate
        fmsgboxsmall msgValidateSheetDPM15 ', vbOKOnly, "Error(s)!"
        CloseMsg
    'Else
        'MsgBox "Sheet is ok", vbOKOnly, "ITR-3"
    End If
    
    If Not ValidatesheetDAOB5 Then
        Sheet10.Activate
        fmsgboxsmall msgValidateSheetDAOB5 ', vbOKOnly, "Error(s)!"
        CloseMsg
    Else
        'MsgBox "Sheet is ok", vbOKOnly, "ITR-3"
    End If
End Sub

Function ValidatesheetDPM15() As Boolean
     Dim strpassword As String
     
     ValidatesheetDPM15 = True
     strpassword = EfilingCommon.getmsgstate
     Sheet10.Unprotect strpassword
    
    If Not ValidateWDVFirstDay_DPM15() Then ValidatesheetDPM15 = False
    'Malli
    If Not ValidateAdjmt2NDprov3section115BAC() Then ValidatesheetDPM15 = False
    '-------
    If Not ValidateAdditionsGrThan180Days_DPM15() Then ValidatesheetDPM15 = False
    If Not ValidateRealizationTotalPeriod_DPM15() Then ValidatesheetDPM15 = False
    If Not ValidateFullRateDeprAmt_DPM15() Then ValidatesheetDPM15 = False
    If Not ValidateAdditionsLessThan180Days_DPM15() Then ValidatesheetDPM15 = False
    If Not ValidateRealizationPeriodDuringYear_DPM15() Then ValidatesheetDPM15 = False
    If Not ValidateHalfRateDeprAmt_DPM15() Then ValidatesheetDPM15 = False
    If Not ValidateDepreciationAtFullRate_DPM15() Then ValidatesheetDPM15 = False
    If Not ValidateDepreciationAtHalfRate_DPM15() Then ValidatesheetDPM15 = False
    If Not ValidateAddlnDeprOnGT180DayAdditions_DPM15() Then ValidatesheetDPM15 = False
    If Not ValidateAddlnDeprDuringYearAdditions_DPM15() Then ValidatesheetDPM15 = False
    If Not ValidateTotalDepreciation_DPM15() Then ValidatesheetDPM15 = False
     If Not ValidateDepreciation_Sec38_DPM15() Then ValidatesheetDPM15 = False
    If Not ValidateExpdrOnTrforSaleAsset_DPM15() Then ValidatesheetDPM15 = False
    If Not ValidateCapGainUs50_DPM15() Then ValidatesheetDPM15 = False
    If Not ValidateWDVLastDay_DPM15() Then ValidatesheetDPM15 = False
    If Not ValidateWDVFirstDay_DPM30() Then ValidatesheetDPM15 = False
    If Not ValidateAdditionsGrThan180Days_DPM30() Then ValidatesheetDPM15 = False
    If Not ValidateRealizationTotalPeriod_DPM30() Then ValidatesheetDPM15 = False
    If Not ValidateFullRateDeprAmt_DPM30() Then ValidatesheetDPM15 = False
    If Not ValidateAdditionsLessThan180Days_DPM30() Then ValidatesheetDPM15 = False
    If Not ValidateRealizationPeriodDuringYear_DPM30() Then ValidatesheetDPM15 = False
    If Not ValidateHalfRateDeprAmt_DPM30() Then ValidatesheetDPM15 = False
    If Not ValidateDepreciationAtFullRate_DPM30() Then ValidatesheetDPM15 = False
    If Not ValidateDepreciationAtHalfRate_DPM30() Then ValidatesheetDPM15 = False
    If Not ValidateAddlnDeprOnGT180DayAdditions_DPM30() Then ValidatesheetDPM15 = False
    If Not ValidateAddlnDeprDuringYearAdditions_DPM30() Then ValidatesheetDPM15 = False
    If Not ValidateDepreciation_Sec38_DPM30() Then ValidatesheetDPM15 = False
    If Not ValidateTotalDepreciation_DPM30() Then ValidatesheetDPM15 = False
    If Not ValidateExpdrOnTrforSaleAsset_DPM30() Then ValidatesheetDPM15 = False
    If Not ValidateCapGainUs50_DPM30() Then ValidatesheetDPM15 = False
    If Not ValidateWDVLastDay_DPM30() Then ValidatesheetDPM15 = False
    If Not ValidateWDVFirstDay_DPM40() Then ValidatesheetDPM15 = False
    'If Not ValidateWDVFirstDay2_DPM40() Then ValidatesheetDPM15 = False
    If Not ValidateAdditionsGrThan180Days_DPM40() Then ValidatesheetDPM15 = False
    If Not ValidateRealizationTotalPeriod_DPM40() Then ValidatesheetDPM15 = False
    If Not ValidateFullRateDeprAmt_DPM40() Then ValidatesheetDPM15 = False
    If Not ValidateAdditionsLessThan180Days_DPM40() Then ValidatesheetDPM15 = False
    If Not ValidateRealizationPeriodDuringYear_DPM40() Then ValidatesheetDPM15 = False
    If Not ValidateHalfRateDeprAmt_DPM40() Then ValidatesheetDPM15 = False
    If Not ValidateDepreciationAtFullRate_DPM40() Then ValidatesheetDPM15 = False
    If Not ValidateDepreciationAtHalfRate_DPM40() Then ValidatesheetDPM15 = False
    If Not ValidateAddlnDeprOnGT180DayAdditions_DPM40() Then ValidatesheetDPM15 = False
    If Not ValidateAddlnDeprDuringYearAdditions_DPM40() Then ValidatesheetDPM15 = False
    If Not ValidateTotalDepreciation_DPM40() Then ValidatesheetDPM15 = False
    If Not ValidateDepreciation_Sec38_DPM40() Then ValidatesheetDPM15 = False
    If Not ValidateExpdrOnTrforSaleAsset_DPM40() Then ValidatesheetDPM15 = False
    If Not ValidateCapGainUs50_DPM40() Then ValidatesheetDPM15 = False
    If Not ValidateWDVLastDay_DPM40() Then ValidatesheetDPM15 = False
    
    If Not ValidateWDVFirstDay_DPM45() Then ValidatesheetDPM15 = False
    If Not ValidateAdditionsGrThan180Days_DPM45() Then ValidatesheetDPM15 = False
    If Not ValidateRealizationTotalPeriod_DPM45() Then ValidatesheetDPM15 = False
    If Not ValidateFullRateDeprAmt_DPM45() Then ValidatesheetDPM15 = False
    If Not ValidateAdditionsLessThan180Days_DPM45() Then ValidatesheetDPM15 = False
    If Not ValidateRealizationPeriodDuringYear_DPM45() Then ValidatesheetDPM15 = False
    If Not ValidateHalfRateDeprAmt_DPM45() Then ValidatesheetDPM15 = False
    If Not ValidateDepreciationAtFullRate_DPM45() Then ValidatesheetDPM15 = False
    If Not ValidateDepreciationAtHalfRate_DPM45() Then ValidatesheetDPM15 = False
    If Not ValidateAddlnDeprOnGT180DayAdditions_DPM45() Then ValidatesheetDPM15 = False
    If Not ValidateAddlnDeprDuringYearAdditions_DPM45() Then ValidatesheetDPM15 = False
    If Not ValidateTotalDepreciation_DPM45() Then ValidatesheetDPM15 = False
    If Not ValidateDepreciation_Sec38_DPM45() Then ValidatesheetDPM15 = False
    If Not ValidateExpdrOnTrforSaleAsset_DPM45() Then ValidatesheetDPM15 = False
    If Not ValidateCapGainUs50_DPM45() Then ValidatesheetDPM15 = False
    If Not ValidateWDVLastDay_DPM45() Then ValidatesheetDPM15 = False

If (Sheet10.Range("DPM15.DepUnderSec38").value) > (Sheet10.Range("DPM15.TotalDepreciation").value) Then
       msgValidateSheetDPM15 = ("* Value at Row no 16 should not be greater than Value at Row no 15")
        ValidatesheetDPM15 = False
End If

If (Sheet10.Range("DPM30.DepUnderSec38").value) > (Sheet10.Range("DPM30.TotalDepreciation").value) Then
        msgValidateSheetDPM15 = ("* Value at Row no 16 should not be greater than Value at Row no 15")
        ValidatesheetDPM15 = False
End If

If (Sheet10.Range("DPM40.DepUnderSec38").value) > (Sheet10.Range("DPM40.TotalDepreciation").value) Then
        msgValidateSheetDPM15 = ("* Value at Row no 16 should not be greater than Value at Row no 15")
        ValidatesheetDPM15 = False
End If


If (Sheet10.Range("DPM15.ProAgrdep").value) > (Sheet10.Range("DPM15.NetAgrdep").value) Then
       msgValidateSheetDPM15 = ("* Value at Row no 18 should not be greater than Value at Row no 17")
        ValidatesheetDPM15 = False
End If

If (Sheet10.Range("DPM30.ProAgrdep").value) > (Sheet10.Range("DPM30.NetAgrdep").value) Then
       msgValidateSheetDPM15 = ("* Value at Row no 18 should not be greater than Value at Row no 17")
        ValidatesheetDPM15 = False
End If

If (Sheet10.Range("DPM40.ProAgrdep").value) > (Sheet10.Range("DPM40.NetAgrdep").value) Then
       msgValidateSheetDPM15 = ("* Value at Row no 18 should not be greater than Value at Row no 17")
        ValidatesheetDPM15 = False
End If

If (Sheet10.Range("DPM45.ProAgrdep").value) > (Sheet10.Range("DPM45.NetAgrdep").value) Then
       msgValidateSheetDPM15 = ("* Value at Row no 18 should not be greater than Value at Row no 17")
        ValidatesheetDPM15 = False
End If

If Len(Sheet10.Range("DPM15.FullRateDeprAmt").value) > 14 Or Len(Sheet10.Range("DPM30.FullRateDeprAmt").value) > 14 Or Len(Sheet10.Range("DPM40.FullRateDeprAmt").value) > 14 Or Len(Sheet10.Range("DPM45.FullRateDeprAmt").value) > 14 Then
        msgValidateSheetDPM15 = ("* Value at Amount on which depreciation at full rate to be allowed (3 + 4 -5)  (Enter 0, if result is negative) cannot be greater than 14 digits")
        ValidatesheetDPM15 = False
End If

If Len(Sheet10.Range("DPM15.TotalDepreciation").value) > 14 Or Len(Sheet10.Range("DPM30.TotalDepreciation").value) > 14 Or Len(Sheet10.Range("DPM40.TotalDepreciation").value) > 14 Or Len(Sheet10.Range("DPM45.TotalDepreciation").value) > 14 Then
        msgValidateSheetDPM15 = ("* Value at Total depreciation (10+11+12+13+14)  cannot be greater than 14 digits")
        ValidatesheetDPM15 = False
End If

'Malli
If Len(Sheet10.Range("DPM15.tot3A3B").value) > 14 Or Len(Sheet10.Range("DPM30.tot3A3B").value) > 14 Or Len(Sheet10.Range("DPM40.tot3A3B").value) > 14 Or Len(Sheet10.Range("DPM45.tot3A3B").value) > 14 Then
        msgValidateSheetDPM15 = ("* Value at Total(3a+3b) cannot be greater than 14 digits")
        ValidatesheetDPM15 = False
End If


'---------


    
    Sheet10.Protect strpassword
End Function

Function ValidatesheetDAOB5() As Boolean
     Dim strpassword As String
     
     ValidatesheetDAOB5 = True

     strpassword = EfilingCommon.getmsgstate
     Sheet10.Unprotect strpassword
    
    If Not ValidateWDVFirstDay_DAOL0() Then ValidatesheetDAOB5 = False
    If Not ValidateWDVLastDay_DAOL0() Then ValidatesheetDAOB5 = False
    If Not ValidateWDVFirstDay_DAOB5() Then ValidatesheetDAOB5 = False
    If Not ValidateAdditionsGrThan180Days_DAOB5() Then ValidatesheetDAOB5 = False
    If Not ValidateRealizationTotalPeriod_DAOB5() Then ValidatesheetDAOB5 = False
    If Not ValidateFullRateDeprAmt_DAOB5() Then ValidatesheetDAOB5 = False
    If Not ValidateAdditionsLessThan180Days_DAOB5() Then ValidatesheetDAOB5 = False
    If Not ValidateRealizationPeriodDuringYear_DAOB5() Then ValidatesheetDAOB5 = False
    If Not ValidateHalfRateDeprAmt_DAOB5() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciationAtFullRate_DAOB5() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciationAtHalfRate_DAOB5() Then ValidatesheetDAOB5 = False
    If Not ValidateAddlnDeprOnGT180DayAdditions_DAOB5() Then ValidatesheetDAOB5 = False
    If Not ValidateAddlnDeprDuringYearAdditions_DAOB5() Then ValidatesheetDAOB5 = False
    If Not ValidateTotalDepreciation_DAOB5() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciation_Sec38_DAOB05() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciation_Sec38_DAOB10() Then ValidatesheetDAOB5 = False
    If Not ValidateExpdrOnTrforSaleAsset_DAOB5() Then ValidatesheetDAOB5 = False
    If Not ValidateCapGainUs50_DAOB5() Then ValidatesheetDAOB5 = False
    If Not ValidateWDVLastDay_DAOB5() Then ValidatesheetDAOB5 = False
    If Not ValidateWDVFirstDay_DAOB10() Then ValidatesheetDAOB5 = False
    If Not ValidateAdditionsGrThan180Days_DAOB10() Then ValidatesheetDAOB5 = False
    If Not ValidateRealizationTotalPeriod_DAOB10() Then ValidatesheetDAOB5 = False
    If Not ValidateFullRateDeprAmt_DAOB10() Then ValidatesheetDAOB5 = False
    If Not ValidateAdditionsLessThan180Days_DAOB10() Then ValidatesheetDAOB5 = False
    If Not ValidateRealizationPeriodDuringYear_DAOB10() Then ValidatesheetDAOB5 = False
    If Not ValidateHalfRateDeprAmt_DAOB10() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciationAtFullRate_DAOB10() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciationAtHalfRate_DAOB10() Then ValidatesheetDAOB5 = False
    If Not ValidateAddlnDeprOnGT180DayAdditions_DAOB10() Then ValidatesheetDAOB5 = False
    If Not ValidateAddlnDeprDuringYearAdditions_DAOB10() Then ValidatesheetDAOB5 = False
    If Not ValidateTotalDepreciation_DAOB10() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciation_Sec38_DAOB05() Then ValidatesheetDAOB5 = False
    If Not ValidateExpdrOnTrforSaleAsset_DAOB10() Then ValidatesheetDAOB5 = False
    If Not ValidateCapGainUs50_DAOB10() Then ValidatesheetDAOB5 = False
    If Not ValidateWDVLastDay_DAOB10() Then ValidatesheetDAOB5 = False
    If Not ValidateWDVFirstDay_DAOB100() Then ValidatesheetDAOB5 = False
    If Not ValidateAdditionsGrThan180Days_DAOB100() Then ValidatesheetDAOB5 = False
    If Not ValidateRealizationTotalPeriod_DAOB100() Then ValidatesheetDAOB5 = False
    If Not ValidateFullRateDeprAmt_DAOB100() Then ValidatesheetDAOB5 = False
    If Not ValidateAdditionsLessThan180Days_DAOB100() Then ValidatesheetDAOB5 = False
    If Not ValidateRealizationPeriodDuringYear_DAOB100() Then ValidatesheetDAOB5 = False
    If Not ValidateHalfRateDeprAmt_DAOB100() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciationAtFullRate_DAOB100() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciationAtHalfRate_DAOB100() Then ValidatesheetDAOB5 = False
    If Not ValidateAddlnDeprOnGT180DayAdditions_DAOB100() Then ValidatesheetDAOB5 = False
    If Not ValidateAddlnDeprDuringYearAdditions_DAOB100() Then ValidatesheetDAOB5 = False
    If Not ValidateTotalDepreciation_DAOB100() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciation_Sec38_DAOB40() Then ValidatesheetDAOB5 = False
    If Not ValidateExpdrOnTrforSaleAsset_DAOB100() Then ValidatesheetDAOB5 = False
    If Not ValidateCapGainUs50_DAOB100() Then ValidatesheetDAOB5 = False
    If Not ValidateWDVLastDay_DAOB100() Then ValidatesheetDAOB5 = False
    If Not ValidateWDVFirstDay_DAOF10() Then ValidatesheetDAOB5 = False
    If Not ValidateAdditionsGrThan180Days_DAOF10() Then ValidatesheetDAOB5 = False
    If Not ValidateRealizationTotalPeriod_DAOF10() Then ValidatesheetDAOB5 = False
    If Not ValidateFullRateDeprAmt_DAOF10() Then ValidatesheetDAOB5 = False
    If Not ValidateAdditionsLessThan180Days_DAOF10() Then ValidatesheetDAOB5 = False
    If Not ValidateRealizationPeriodDuringYear_DAOF10() Then ValidatesheetDAOB5 = False
    If Not ValidateHalfRateDeprAmt_DAOF10() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciationAtFullRate_DAOF10() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciationAtHalfRate_DAOF10() Then ValidatesheetDAOB5 = False
    If Not ValidateAddlnDeprOnGT180DayAdditions_DAOF10() Then ValidatesheetDAOB5 = False
    If Not ValidateAddlnDeprDuringYearAdditions_DAOF10() Then ValidatesheetDAOB5 = False
    If Not ValidateTotalDepreciation_DAOF10() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciation_Sec38_DAOF10() Then ValidatesheetDAOB5 = False
    If Not ValidateExpdrOnTrforSaleAsset_DAOF10() Then ValidatesheetDAOB5 = False
    If Not ValidateCapGainUs50_DAOF10() Then ValidatesheetDAOB5 = False
    If Not ValidateWDVLastDay_DAOF10() Then ValidatesheetDAOB5 = False
    If Not ValidateWDVFirstDay_DAOI25() Then ValidatesheetDAOB5 = False
    If Not ValidateAdditionsGrThan180Days_DAOI25() Then ValidatesheetDAOB5 = False
    If Not ValidateRealizationTotalPeriod_DAOI25() Then ValidatesheetDAOB5 = False
    If Not ValidateFullRateDeprAmt_DAOI25() Then ValidatesheetDAOB5 = False
    If Not ValidateAdditionsLessThan180Days_DAOI25() Then ValidatesheetDAOB5 = False
    If Not ValidateRealizationPeriodDuringYear_DAOI25() Then ValidatesheetDAOB5 = False
    If Not ValidateHalfRateDeprAmt_DAOI25() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciationAtFullRate_DAOI25() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciationAtHalfRate_DAOI25() Then ValidatesheetDAOB5 = False
    If Not ValidateAddlnDeprOnGT180DayAdditions_DAOI25() Then ValidatesheetDAOB5 = False
    If Not ValidateAddlnDeprDuringYearAdditions_DAOI25() Then ValidatesheetDAOB5 = False
    If Not ValidateTotalDepreciation_DAOI25() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciation_Sec38_DAOI25() Then ValidatesheetDAOB5 = False
    If Not ValidateExpdrOnTrforSaleAsset_DAOI25() Then ValidatesheetDAOB5 = False
    If Not ValidateCapGainUs50_DAOI25() Then ValidatesheetDAOB5 = False
    If Not ValidateWDVLastDay_DAOI25() Then ValidatesheetDAOB5 = False
    If Not ValidateWDVFirstDay_DAOS20() Then ValidatesheetDAOB5 = False
    If Not ValidateAdditionsGrThan180Days_DAOS20() Then ValidatesheetDAOB5 = False
    If Not ValidateRealizationTotalPeriod_DAOS20() Then ValidatesheetDAOB5 = False
    If Not ValidateFullRateDeprAmt_DAOS20() Then ValidatesheetDAOB5 = False
    If Not ValidateAdditionsLessThan180Days_DAOS20() Then ValidatesheetDAOB5 = False
    If Not ValidateRealizationPeriodDuringYear_DAOS20() Then ValidatesheetDAOB5 = False
    If Not ValidateHalfRateDeprAmt_DAOS20() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciationAtFullRate_DAOS20() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciationAtHalfRate_DAOS20() Then ValidatesheetDAOB5 = False
    If Not ValidateAddlnDeprOnGT180DayAdditions_DAOS20() Then ValidatesheetDAOB5 = False
    If Not ValidateAddlnDeprDuringYearAdditions_DAOS20() Then ValidatesheetDAOB5 = False
    If Not ValidateTotalDepreciation_DAOS20() Then ValidatesheetDAOB5 = False
    If Not ValidateDepreciation_Sec38_DAOS20() Then ValidatesheetDAOB5 = False
    If Not ValidateExpdrOnTrforSaleAsset_DAOS20() Then ValidatesheetDAOB5 = False
    If Not ValidateCapGainUs50_DAOS20() Then ValidatesheetDAOB5 = False
    If Not ValidateWDVLastDay_DAOS20() Then ValidatesheetDAOB5 = False
    
    If (Sheet10.Range("DAOB5.DepUnderSec38").value) > (Sheet10.Range("DAOB5.TotalDepreciation").value) Then
           msgValidateSheetDAOB5 = ("* Value at Row no 13 should not be greater than Value at Row no 12")
            ValidatesheetDAOB5 = False
    End If
    
    If (Sheet10.Range("DAOB10.DepUnderSec38").value) > (Sheet10.Range("DAOB10.TotalDepreciation").value) Then
            msgValidateSheetDAOB5 = ("* Value at Row no 13 should not be greater than Value at Row no 12")
            ValidatesheetDAOB5 = False
    End If
    
    If (Sheet10.Range("DAOB100.DepUnderSec38").value) > (Sheet10.Range("DAOB100.TotalDepreciation").value) Then
            msgValidateSheetDAOB5 = ("* Value at Row no 13 should not be greater than Value at Row no 12")
            ValidatesheetDAOB5 = False
    End If
    
    If (Sheet10.Range("DAOF10.DepUnderSec38").value) > (Sheet10.Range("DAOF10.TotalDepreciation").value) Then
            msgValidateSheetDAOB5 = ("* Value at Row no 13 should not be greater than Value at Row no 12")
            ValidatesheetDAOB5 = False
    End If
    
    If (Sheet10.Range("DAOI25.DepUnderSec38").value) > (Sheet10.Range("DAOI25.TotalDepreciation").value) Then
            msgValidateSheetDAOB5 = ("* Value at Row no 13 should not be greater than Value at Row no 12")
            ValidatesheetDAOB5 = False
    End If
    
    If (Sheet10.Range("DAOS20.DepUnderSec38").value) > (Sheet10.Range("DAOS20.TotalDepreciation").value) Then
            msgValidateSheetDAOB5 = ("* Value at Row no 13 should not be greater than Value at Row no 12")
            ValidatesheetDAOB5 = False
    End If
    
    If (Sheet10.Range("DAOB5.ProAgrdep").value) > (Sheet10.Range("DAOB5.NetAgrdep").value) Then
           msgValidateSheetDAOB5 = ("* Value at Row no 15 should not be greater than Value at Row no 14")
            ValidatesheetDAOB5 = False
    End If
    
    If (Sheet10.Range("DAOB10.ProAgrdep").value) > (Sheet10.Range("DAOB10.NetAgrdep").value) Then
           msgValidateSheetDAOB5 = ("* Value at Row no 15 should not be greater than Value at Row no 14")
            ValidatesheetDAOB5 = False
    End If
    
    If (Sheet10.Range("DAOB100.ProAgrdep").value) > (Sheet10.Range("DAOB100.NetAgrdep").value) Then
           msgValidateSheetDAOB5 = ("* Value at Row no 15 should not be greater than Value at Row no 14")
            ValidatesheetDAOB5 = False
    End If
    
    If (Sheet10.Range("DAOF10.ProAgrdep").value) > (Sheet10.Range("DAOF10.NetAgrdep").value) Then
           msgValidateSheetDAOB5 = ("* Value at Row no 15 should not be greater than Value at Row no 14")
            ValidatesheetDAOB5 = False
    End If
    
    If (Sheet10.Range("DAOI25.ProAgrdep").value) > (Sheet10.Range("DAOI25.NetAgrdep").value) Then
           msgValidateSheetDAOB5 = ("* Value at Row no 15 should not be greater than Value at Row no 14")
            ValidatesheetDAOB5 = False
    End If
    
    If (Sheet10.Range("DAOS20.ProAgrdep").value) > (Sheet10.Range("DAOS20.NetAgrdep").value) Then
           msgValidateSheetDAOB5 = ("* Value at Row no 15 should not be greater than Value at Row no 14")
            ValidatesheetDAOB5 = False
    End If


'Added by sudha as per SIT-95263

If Len(Sheet10.Range("DAOB5.FullRateDeprAmt").value) > 14 Or Len(Sheet10.Range("DAOB10.FullRateDeprAmt").value) > 14 Or Len(Sheet10.Range("DAOB100.FullRateDeprAmt").value) > 14 Or Len(Sheet10.Range("DAOF10.FullRateDeprAmt").value) > 14 Or Len(Sheet10.Range("DAOI25.FullRateDeprAmt").value) > 14 Or Len(Sheet10.Range("DAOS20.FullRateDeprAmt").value) > 14 Then
        msgValidateSheetDAOB5 = ("* Value at Amount on which depreciation at full rate to be allowed (3 + 4 -5)  (Enter 0, if result is negative) cannot be greater than 14 digits")
        ValidatesheetDAOB5 = False
End If





'===================================



    Sheet10.Protect strpassword
End Function

Function ValidateWDVFirstDay_DPM15() As Boolean
    ValidateWDVFirstDay_DPM15 = True
    WDVFirstDay_DPM15 = Sheet10.Range("DPM15.WDVFirstDay").text
End Function

'Malli
Function ValidateAdjmt2NDprov3section115BAC() As Boolean
    ValidateAdjmt2NDprov3section115BAC = True
    Adjmt2NDprov3section115BAC_DPM15 = Sheet10.Range("DPM15.Adjmt2NDprov3section115BAC").text
    Adjmt2NDprov3section115BAC_DPM30 = Sheet10.Range("DPM30.Adjmt2NDprov3section115BAC").text
    Adjmt2NDprov3section115BAC_DPM40 = Sheet10.Range("DPM40.Adjmt2NDprov3section115BAC").text
    Adjmt2NDprov3section115BAC_DPM45 = Sheet10.Range("DPM45.Adjmt2NDprov3section115BAC").text
    
    
End Function
'-------

Function ValidateAdditionsGrThan180Days_DPM15() As Boolean
    ValidateAdditionsGrThan180Days_DPM15 = True
    AdditionsGrThan180Days_DPM15 = Sheet10.Range("DPM15.AdditionsGrThan180Days").text
End Function

Function ValidateRealizationTotalPeriod_DPM15() As Boolean
    ValidateRealizationTotalPeriod_DPM15 = True
    RealizationTotalPeriod_DPM15 = Sheet10.Range("DPM15.RealizationTotalPeriod").text
End Function

Function ValidateFullRateDeprAmt_DPM15() As Boolean
    ValidateFullRateDeprAmt_DPM15 = True
    FullRateDeprAmt_DPM15 = Sheet10.Range("DPM15.FullRateDeprAmt").text
End Function

Function ValidateAdditionsLessThan180Days_DPM15() As Boolean
    ValidateAdditionsLessThan180Days_DPM15 = True
    AdditionsLessThan180Days_DPM15 = Sheet10.Range("DPM15.AdditionsLessThan180Days").text
End Function

Function ValidateRealizationPeriodDuringYear_DPM15() As Boolean
    ValidateRealizationPeriodDuringYear_DPM15 = True
    RealizationPeriodDuringYear_DPM15 = Sheet10.Range("DPM15.RealizationPeriodDuringYear").text
End Function

Function ValidateHalfRateDeprAmt_DPM15() As Boolean
    ValidateHalfRateDeprAmt_DPM15 = True
    HalfRateDeprAmt_DPM15 = Sheet10.Range("DPM15.HalfRateDeprAmt").text
End Function

Function ValidateDepreciationAtFullRate_DPM15() As Boolean
    ValidateDepreciationAtFullRate_DPM15 = True
    DepreciationAtFullRate_DPM15 = Sheet10.Range("DPM15.DepreciationAtFullRate").text
End Function

Function ValidateDepreciationAtHalfRate_DPM15() As Boolean
    ValidateDepreciationAtHalfRate_DPM15 = True
    DepreciationAtHalfRate_DPM15 = Sheet10.Range("DPM15.DepreciationAtHalfRate").text
End Function

Function ValidateAddlnDeprOnGT180DayAdditions_DPM15() As Boolean
    ValidateAddlnDeprOnGT180DayAdditions_DPM15 = True
    AddlnDeprOnGT180DayAdditions_DPM15 = Sheet10.Range("DPM15.AddlnDeprOnGT180DayAdditions").text
End Function

Function ValidateAddlnDeprDuringYearAdditions_DPM15() As Boolean
    ValidateAddlnDeprDuringYearAdditions_DPM15 = True
    AddlnDeprDuringYearAdditions_DPM15 = Sheet10.Range("DPM15.AddlnDeprDuringYearAdditions").text
    AddlnDeprlessthan180days_DPM15 = Sheet10.Range("DPM15.AddlnDeprlessthan180days").text
End Function

Function ValidateTotalDepreciation_DPM15() As Boolean
    ValidateTotalDepreciation_DPM15 = True
    TotalDepreciation_DPM15 = Sheet10.Range("DPM15.TotalDepreciation").text
End Function
Function ValidateDepreciation_Sec38_DPM15() As Boolean
    ValidateDepreciation_Sec38_DPM15 = True
    Depreciation_Sec38_DPM15 = Sheet10.Range("DPM15.DepUnderSec38").text
    NetaggregateDep_DPM15 = Sheet10.Range("DPM15.NetAgrdep").text
    PropaggregateDep_DPM15 = Sheet10.Range("DPM15.ProAgrdep").text
    
    
End Function

Function ValidateExpdrOnTrforSaleAsset_DPM15() As Boolean
    ValidateExpdrOnTrforSaleAsset_DPM15 = True
    ExpdrOnTrforSaleAsset_DPM15 = Sheet10.Range("DPM15.ExpdrOnTrforSaleAsset").text
End Function
Function ValidateCapGainUs50_DPM15() As Boolean
 On Error Resume Next
 ValidateCapGainUs50_DPM15 = True
 CapGainUs50_DPM15 = Sheet10.Range("DPM15.CapGainUs50").value
 
 'Change-29.04.2023.101.DPM -> SIT-48795 A
' If Sheet10.Range("DPM15.check20A").value > 0 Then
'    If Sheet10.Range("DPM15.CapGainUs50") = "" Or Sheet10.Range("DPM15.CapGainUs50") = 0 Then
'    ValidateCapGainUs50_DPM15 = False
'    msgValidateSheetDPM15 = msgValidateSheetDPM15 & Chr(13) & "DPM15:Capital gains/ loss under section 50 can not be Blank"
'    End If
' End If
 
    If Sheet10.Range("DPM15.WDVLastDay").value = 0 And Sheet10.Range("DPM15.CapGainUs50").value = "" Then
        ValidateCapGainUs50_DPM15 = False
        msgValidateSheetDPM15 = msgValidateSheetDPM15 & Chr(13) & "DPM15:Capital gains/ loss under section 50 can not be Blank"
    End If
 '----
 If CapGainUs50_DPM15 <> 0 Then
    Sheet10.Activate
    Sheet10.cg_gt_lt_0 (CapGainUs50_DPM15), "DPM15.CapGainUs50"
    fmsgboxStatus "Block Depreciation - Plant and Machinery @ 15% ceases to exist"
    'ValidateCapGainUs50_DPM15 = False
    End If
End Function

Function ValidateCapGainUs50_DPM15old() As Boolean
On Error Resume Next
    ValidateCapGainUs50_DPM15 = True
    CapGainUs50_DPM15 = Sheet10.Range("DPM15.CapGainUs50").value
    
    If CapGainUs50_DPM15 <> 0 Then
       Sheet10.Activate
       Sheet10.cg_gt_lt_0 (CapGainUs50_DPM15), "DPM15.CapGainUs50"
       fmsgboxStatus "Block Depreciation - Plant and Machinery @ 15% ceases to exist"
    End If
End Function


Function ValidateWDVLastDay_DPM15() As Boolean
    ValidateWDVLastDay_DPM15 = True
    WDVLastDay_DPM15 = Sheet10.Range("DPM15.WDVLastDay").text
End Function

Function ValidateWDVFirstDay_DPM30() As Boolean
    ValidateWDVFirstDay_DPM30 = True
    WDVFirstDay_DPM30 = Sheet10.Range("DPM30.WDVFirstDay").text
End Function

Function ValidateAdditionsGrThan180Days_DPM30() As Boolean
    ValidateAdditionsGrThan180Days_DPM30 = True
    AdditionsGrThan180Days_DPM30 = Sheet10.Range("DPM30.AdditionsGrThan180Days").text
End Function

Function ValidateRealizationTotalPeriod_DPM30() As Boolean
    ValidateRealizationTotalPeriod_DPM30 = True
    RealizationTotalPeriod_DPM30 = Sheet10.Range("DPM30.RealizationTotalPeriod").text
End Function

Function ValidateFullRateDeprAmt_DPM30() As Boolean
    ValidateFullRateDeprAmt_DPM30 = True
    FullRateDeprAmt_DPM30 = Sheet10.Range("DPM30.FullRateDeprAmt").text
End Function

Function ValidateAdditionsLessThan180Days_DPM30() As Boolean
    ValidateAdditionsLessThan180Days_DPM30 = True
    AdditionsLessThan180Days_DPM30 = Sheet10.Range("DPM30.AdditionsLessThan180Days").text
End Function

Function ValidateRealizationPeriodDuringYear_DPM30() As Boolean
    ValidateRealizationPeriodDuringYear_DPM30 = True
    RealizationPeriodDuringYear_DPM30 = Sheet10.Range("DPM30.RealizationPeriodDuringYear").text
End Function

Function ValidateHalfRateDeprAmt_DPM30() As Boolean
    ValidateHalfRateDeprAmt_DPM30 = True
    HalfRateDeprAmt_DPM30 = Sheet10.Range("DPM30.HalfRateDeprAmt").text
End Function

Function ValidateDepreciationAtFullRate_DPM30() As Boolean
    ValidateDepreciationAtFullRate_DPM30 = True
    DepreciationAtFullRate_DPM30 = Sheet10.Range("DPM30.DepreciationAtFullRate").text
End Function

Function ValidateDepreciationAtHalfRate_DPM30() As Boolean
    ValidateDepreciationAtHalfRate_DPM30 = True
    DepreciationAtHalfRate_DPM30 = Sheet10.Range("DPM30.DepreciationAtHalfRate").text
End Function

Function ValidateAddlnDeprOnGT180DayAdditions_DPM30() As Boolean
    ValidateAddlnDeprOnGT180DayAdditions_DPM30 = True
    AddlnDeprOnGT180DayAdditions_DPM30 = Sheet10.Range("DPM30.AddlnDeprOnGT180DayAdditions").text
End Function

Function ValidateAddlnDeprDuringYearAdditions_DPM30() As Boolean
    ValidateAddlnDeprDuringYearAdditions_DPM30 = True
    AddlnDeprDuringYearAdditions_DPM30 = Sheet10.Range("DPM30.AddlnDeprDuringYearAdditions").text
    AddlnDeprlessthan180days_DPM30 = Sheet10.Range("DPM30.AddlnDeprlessthan180days").text
End Function

Function ValidateTotalDepreciation_DPM30() As Boolean
    ValidateTotalDepreciation_DPM30 = True
    TotalDepreciation_DPM30 = Sheet10.Range("DPM30.TotalDepreciation").text
End Function
Function ValidateDepreciation_Sec38_DPM30() As Boolean
    ValidateDepreciation_Sec38_DPM30 = True
    Depreciation_Sec38_DPM30 = Sheet10.Range("DPM30.DepUnderSec38").text
    NetaggregateDep_DPM30 = Sheet10.Range("DPM30.NetAgrdep").text
    PropaggregateDep_DPM30 = Sheet10.Range("DPM30.ProAgrdep").text
End Function

Function ValidateExpdrOnTrforSaleAsset_DPM30() As Boolean
    ValidateExpdrOnTrforSaleAsset_DPM30 = True
    ExpdrOnTrforSaleAsset_DPM30 = Sheet10.Range("DPM30.ExpdrOnTrforSaleAsset").text
End Function
Function ValidateCapGainUs50_DPM30() As Boolean
    ValidateCapGainUs50_DPM30 = True
    CapGainUs50_DPM30 = Sheet10.Range("DPM30.CapGainUs50").value
    'Change-29.04.2023.101.DPM -> SIT-48795 B
'    If Sheet10.Range("DPM30.check20A").value > 0 Then
'        If Sheet10.Range("DPM30.CapGainUs50") = "" Or Sheet10.Range("DPM30.CapGainUs50") = 0 Then
'        ValidateCapGainUs50_DPM30 = False
'        msgValidateSheetDPM15 = msgValidateSheetDPM15 & Chr(13) & "DPM30:Capital gains/ loss under section 50 can not be Blank"
'        End If
'    End If

If Sheet10.Range("DPM30.WDVLastDay").value = 0 And Sheet10.Range("DPM30.CapGainUs50").value = "" Then
    ValidateCapGainUs50_DPM30 = False
    msgValidateSheetDPM15 = msgValidateSheetDPM15 & Chr(13) & "DPM30:Capital gains/ loss under section 50 can not be Blank"
End If
'-----
    If CapGainUs50_DPM30 <> 0 Then
       Sheet10.Activate
       Sheet10.cg_gt_lt_0 (CapGainUs50_DPM30), "DPM30.CapGainUs50"
       fmsgboxStatus "Block Depreciation - Plant and Machinery @ 30% ceases to exist"
    End If
End Function
Function ValidateCapGainUs50_DPM30old() As Boolean
    ValidateCapGainUs50_DPM30 = True
    CapGainUs50_DPM30 = Sheet10.Range("DPM30.CapGainUs50").value
    
    If CapGainUs50_DPM30 <> 0 Then
       Sheet10.Activate
       Sheet10.cg_gt_lt_0 (CapGainUs50_DPM30), "DPM30.CapGainUs50"
       fmsgboxStatus "Block Depreciation - Plant and Machinery @ 30% ceases to exist"
    End If
End Function

Function ValidateWDVLastDay_DPM30() As Boolean
    ValidateWDVLastDay_DPM30 = True
    WDVLastDay_DPM30 = Sheet10.Range("DPM30.WDVLastDay").text
End Function

Function ValidateWDVFirstDay_DPM40() As Boolean
    ValidateWDVFirstDay_DPM40 = True
    WDVFirstDay_DPM40 = Sheet10.Range("DPM40.WDVFirstDay").text
End Function


Function ValidateAdditionsGrThan180Days_DPM40() As Boolean
    ValidateAdditionsGrThan180Days_DPM40 = True
    AdditionsGrThan180Days_DPM40 = Sheet10.Range("DPM40.AdditionsGrThan180Days").text
End Function

Function ValidateRealizationTotalPeriod_DPM40() As Boolean
    ValidateRealizationTotalPeriod_DPM40 = True
    RealizationTotalPeriod_DPM40 = Sheet10.Range("DPM40.RealizationTotalPeriod").text
End Function

Function ValidateFullRateDeprAmt_DPM40() As Boolean
    ValidateFullRateDeprAmt_DPM40 = True
    FullRateDeprAmt_DPM40 = Sheet10.Range("DPM40.FullRateDeprAmt").text
End Function

Function ValidateAdditionsLessThan180Days_DPM40() As Boolean
    ValidateAdditionsLessThan180Days_DPM40 = True
    AdditionsLessThan180Days_DPM40 = Sheet10.Range("DPM40.AdditionsLessThan180Days").text
End Function

Function ValidateRealizationPeriodDuringYear_DPM40() As Boolean
    ValidateRealizationPeriodDuringYear_DPM40 = True
    RealizationPeriodDuringYear_DPM40 = Sheet10.Range("DPM40.RealizationPeriodDuringYear").text
End Function

Function ValidateHalfRateDeprAmt_DPM40() As Boolean
    ValidateHalfRateDeprAmt_DPM40 = True
    HalfRateDeprAmt_DPM40 = Sheet10.Range("DPM40.HalfRateDeprAmt").text
End Function

Function ValidateDepreciationAtFullRate_DPM40() As Boolean
    ValidateDepreciationAtFullRate_DPM40 = True
    DepreciationAtFullRate_DPM40 = Sheet10.Range("DPM40.DepreciationAtFullRate").text
End Function

Function ValidateDepreciationAtHalfRate_DPM40() As Boolean
    ValidateDepreciationAtHalfRate_DPM40 = True
    DepreciationAtHalfRate_DPM40 = Sheet10.Range("DPM40.DepreciationAtHalfRate").text
End Function

Function ValidateAddlnDeprOnGT180DayAdditions_DPM40() As Boolean
    ValidateAddlnDeprOnGT180DayAdditions_DPM40 = True
    AddlnDeprOnGT180DayAdditions_DPM40 = Sheet10.Range("DPM40.AddlnDeprOnGT180DayAdditions").text
End Function

Function ValidateAddlnDeprDuringYearAdditions_DPM40() As Boolean
    ValidateAddlnDeprDuringYearAdditions_DPM40 = True
    AddlnDeprDuringYearAdditions_DPM40 = Sheet10.Range("DPM40.AddlnDeprDuringYearAdditions").text
    AddlnDeprlessthan180days_DPM40 = Sheet10.Range("DPM40.AddlnDeprlessthan180days").text
End Function

Function ValidateTotalDepreciation_DPM40() As Boolean
    ValidateTotalDepreciation_DPM40 = True
    TotalDepreciation_DPM40 = Sheet10.Range("DPM40.TotalDepreciation").text
End Function
Function ValidateDepreciation_Sec38_DPM40() As Boolean
    ValidateDepreciation_Sec38_DPM40 = True
    Depreciation_Sec38_DPM40 = Sheet10.Range("DPM40.DepUnderSec38").text
    NetaggregateDep_DPM40 = Sheet10.Range("DPM40.NetAgrdep").text
    PropaggregateDep_DPM40 = Sheet10.Range("DPM40.ProAgrdep").text
End Function

Function ValidateExpdrOnTrforSaleAsset_DPM40() As Boolean
    ValidateExpdrOnTrforSaleAsset_DPM40 = True
    ExpdrOnTrforSaleAsset_DPM40 = Sheet10.Range("DPM40.ExpdrOnTrforSaleAsset").text
End Function
Function ValidateCapGainUs50_DPM40() As Boolean
    ValidateCapGainUs50_DPM40 = True
    CapGainUs50_DPM40 = Sheet10.Range("DPM40.CapGainUs50").value
    'Change-29.04.2023.101.DPM -> SIT-48795 C
'    If Sheet10.Range("DPM40.check20A").value > 0 Then
'        If Sheet10.Range("DPM40.CapGainUs50") = "" Or Sheet10.Range("DPM40.CapGainUs50") = 0 Then
'        ValidateCapGainUs50_DPM40 = False
'        msgValidateSheetDPM15 = msgValidateSheetDPM15 & Chr(13) & "DPM40:Capital gains/ loss under section 50 can not be Blank"
'        End If
'    End If
    
    If Sheet10.Range("DPM40.WDVLastDay") = 0 And Sheet10.Range("DPM40.CapGainUs50") = "" Then
        ValidateCapGainUs50_DPM40 = False
        msgValidateSheetDPM15 = msgValidateSheetDPM15 & Chr(13) & "DPM40:Capital gains/ loss under section 50 can not be Blank"
    End If
'----
    If CapGainUs50_DPM40 <> 0 Then
        Sheet10.Activate
        Sheet10.cg_gt_lt_0 (CapGainUs50_DPM40), "DPM40.CapGainUs50"
        fmsgboxStatus "Block Depreciation - Plant and Machinery @ 40% ceases to exist"
    End If
End Function
Function ValidateCapGainUs50_DPM40old() As Boolean
    ValidateCapGainUs50_DPM40 = True
    CapGainUs50_DPM40 = Sheet10.Range("DPM40.CapGainUs50").value
    
    If CapGainUs50_DPM40 <> 0 Then
        Sheet10.Activate
        Sheet10.cg_gt_lt_0 (CapGainUs50_DPM40), "DPM40.CapGainUs50"
        fmsgboxStatus "Block Depreciation - Plant and Machinery @ 40% ceases to exist"
    End If
End Function

Function ValidateWDVLastDay_DPM40() As Boolean
    ValidateWDVLastDay_DPM40 = True
    WDVLastDay_DPM40 = Sheet10.Range("DPM40.WDVLastDay").text
End Function
Function ValidateCapGainUs50_DPM45() As Boolean
    ValidateCapGainUs50_DPM45 = True
    CapGainUs50_DPM45 = Sheet10.Range("DPM45.CapGainUs50").value
    'Change-29.04.2023.101.DPM -> SIT-48795 D
'    If Sheet10.Range("DPM45.check20A").value > 0 Then
'        If Sheet10.Range("DPM45.CapGainUs50") = "" Or Sheet10.Range("DPM45.CapGainUs50") = 0 Then
'        ValidateCapGainUs50_DPM45 = False
'        msgValidateSheetDPM15 = msgValidateSheetDPM15 & Chr(13) & "DPM45:Capital gains/ loss under section 50 can not be Blank"
'        End If
'    End If
'SIT -97174 need to confirm
'If Sheet10.Range("DPM45.WDVLastDay") = 0 And Sheet10.Range("DPM45.CapGainUs50") = "" And Sheet10.Range("DPM45.CapGainUs50").Locked = False Then
'    ValidateCapGainUs50_DPM45 = False
'    msgValidateSheetDPM15 = msgValidateSheetDPM15 & Chr(13) & "DPM45:Capital gains/ loss under section 50 can not be Blank"
'End If
'----
    If CapGainUs50_DPM45 <> 0 Then
        Sheet10.Activate
        Sheet10.cg_gt_lt_0 (CapGainUs50_DPM45), "DPM45.CapGainUs50"
        fmsgboxStatus "Block Depreciation - Plant and Machinery @ 45% ceases to exist"
    End If
End Function

Function ValidateCapGainUs50_DPM45old() As Boolean
    ValidateCapGainUs50_DPM45 = True
    CapGainUs50_DPM45 = Sheet10.Range("DPM45.CapGainUs50").value
    
    If CapGainUs50_DPM45 <> 0 Then
        Sheet10.Activate
        Sheet10.cg_gt_lt_0 (CapGainUs50_DPM45), "DPM45.CapGainUs50"
        fmsgboxStatus "Block Depreciation - Plant and Machinery @ 45% ceases to exist"
    End If
End Function
Function ValidateWDVFirstDay_DPM45() As Boolean
    ValidateWDVFirstDay_DPM45 = True
    WDVFirstDay_DPM45 = Sheet10.Range("DPM45.WDVFirstDay").text
End Function


Function ValidateAdditionsGrThan180Days_DPM45() As Boolean
    ValidateAdditionsGrThan180Days_DPM45 = True
    AdditionsGrThan180Days_DPM45 = Sheet10.Range("DPM45.AdditionsGrThan180Days").text
End Function

Function ValidateRealizationTotalPeriod_DPM45() As Boolean
    ValidateRealizationTotalPeriod_DPM45 = True
    RealizationTotalPeriod_DPM45 = Sheet10.Range("DPM45.RealizationTotalPeriod").text
End Function

Function ValidateFullRateDeprAmt_DPM45() As Boolean
    ValidateFullRateDeprAmt_DPM45 = True
    FullRateDeprAmt_DPM45 = Sheet10.Range("DPM45.FullRateDeprAmt").text
End Function

Function ValidateAdditionsLessThan180Days_DPM45() As Boolean
    ValidateAdditionsLessThan180Days_DPM45 = True
    AdditionsLessThan180Days_DPM45 = Sheet10.Range("DPM45.AdditionsLessThan180Days").text
End Function

Function ValidateRealizationPeriodDuringYear_DPM45() As Boolean
    ValidateRealizationPeriodDuringYear_DPM45 = True
    RealizationPeriodDuringYear_DPM45 = Sheet10.Range("DPM45.RealizationPeriodDuringYear").text
End Function

Function ValidateHalfRateDeprAmt_DPM45() As Boolean
    ValidateHalfRateDeprAmt_DPM45 = True
    HalfRateDeprAmt_DPM45 = Sheet10.Range("DPM45.HalfRateDeprAmt").text
End Function

Function ValidateDepreciationAtFullRate_DPM45() As Boolean
    ValidateDepreciationAtFullRate_DPM45 = True
    DepreciationAtFullRate_DPM45 = Sheet10.Range("DPM45.DepreciationAtFullRate").text
End Function

Function ValidateDepreciationAtHalfRate_DPM45() As Boolean
    ValidateDepreciationAtHalfRate_DPM45 = True
    DepreciationAtHalfRate_DPM45 = Sheet10.Range("DPM45.DepreciationAtHalfRate").text
End Function

Function ValidateAddlnDeprOnGT180DayAdditions_DPM45() As Boolean
    ValidateAddlnDeprOnGT180DayAdditions_DPM45 = True
    AddlnDeprOnGT180DayAdditions_DPM45 = Sheet10.Range("DPM45.AddlnDeprOnGT180DayAdditions").text
End Function

Function ValidateAddlnDeprDuringYearAdditions_DPM45() As Boolean
    ValidateAddlnDeprDuringYearAdditions_DPM45 = True
    AddlnDeprDuringYearAdditions_DPM45 = Sheet10.Range("DPM45.AddlnDeprDuringYearAdditions").text
    AddlnDeprlessthan180days_DPM45 = Sheet10.Range("DPM45.AddlnDeprlessthan180days").text
End Function

Function ValidateTotalDepreciation_DPM45() As Boolean
    ValidateTotalDepreciation_DPM45 = True
    TotalDepreciation_DPM45 = Sheet10.Range("DPM45.TotalDepreciation").text
End Function
Function ValidateDepreciation_Sec38_DPM45() As Boolean
    ValidateDepreciation_Sec38_DPM45 = True
    Depreciation_Sec38_DPM45 = Sheet10.Range("DPM45.DepUnderSec38").text
    NetaggregateDep_DPM45 = Sheet10.Range("DPM45.NetAgrdep").text
    PropaggregateDep_DPM45 = Sheet10.Range("DPM45.ProAgrdep").text
End Function

Function ValidateExpdrOnTrforSaleAsset_DPM45() As Boolean
    ValidateExpdrOnTrforSaleAsset_DPM45 = True
    ExpdrOnTrforSaleAsset_DPM45 = Sheet10.Range("DPM45.ExpdrOnTrforSaleAsset").text
End Function

'Function ValidateCapGainUs50_DPM45() As Boolean
'    ValidateCapGainUs50_DPM45 = True
'    CapGainUs50_DPM45 = Sheet10.Range("DPM40.CapGainUs50").value
'
'    If CapGainUs50_DPM45 <> 0 Then
'        Sheet10.Activate
'        Sheet10.cg_gt_lt_0 (CapGainUs50_DPM45), "DPM45.CapGainUs50"
'        fmsgboxStatus "Block Depreciation - Plant and Machinery @ 45% ceases to exist"
'    End If
'End Function

Function ValidateWDVLastDay_DPM45() As Boolean
    ValidateWDVLastDay_DPM45 = True
    WDVLastDay_DPM45 = Sheet10.Range("DPM45.WDVLastDay").text
End Function

'DOA____________________________________________________________________________________________________

Function ValidateWDVFirstDay_DAOL0() As Boolean
    ValidateWDVFirstDay_DAOL0 = True
    WDVFirstDay_DAOL0 = Sheet10.Range("DAOL0.WDVFirstDay").text
End Function

Function ValidateWDVLastDay_DAOL0() As Boolean
    ValidateWDVLastDay_DAOL0 = True
    WDVLastDay_DAOL0 = Sheet10.Range("DAOL0.WDVLastDay").text
End Function

Function ValidateWDVFirstDay_DAOB5() As Boolean
    ValidateWDVFirstDay_DAOB5 = True
    WDVFirstDay_DAOB5 = Sheet10.Range("DAOB5.WDVFirstDay").text
End Function

Function ValidateAdditionsGrThan180Days_DAOB5() As Boolean
    ValidateAdditionsGrThan180Days_DAOB5 = True
    AdditionsGrThan180Days_DAOB5 = Sheet10.Range("DAOB5.AdditionsGrThan180Days").text
End Function

Function ValidateRealizationTotalPeriod_DAOB5() As Boolean
    ValidateRealizationTotalPeriod_DAOB5 = True
    RealizationTotalPeriod_DAOB5 = Sheet10.Range("DAOB5.RealizationTotalPeriod").text
End Function

Function ValidateFullRateDeprAmt_DAOB5() As Boolean
    ValidateFullRateDeprAmt_DAOB5 = True
    FullRateDeprAmt_DAOB5 = Sheet10.Range("DAOB5.FullRateDeprAmt").text
End Function

Function ValidateAdditionsLessThan180Days_DAOB5() As Boolean
    ValidateAdditionsLessThan180Days_DAOB5 = True
    AdditionsLessThan180Days_DAOB5 = Sheet10.Range("DAOB5.AdditionsLessThan180Days").text
End Function

Function ValidateRealizationPeriodDuringYear_DAOB5() As Boolean
    ValidateRealizationPeriodDuringYear_DAOB5 = True
    RealizationPeriodDuringYear_DAOB5 = Sheet10.Range("DAOB5.RealizationPeriodDuringYear").text
End Function

Function ValidateHalfRateDeprAmt_DAOB5() As Boolean
    ValidateHalfRateDeprAmt_DAOB5 = True
    HalfRateDeprAmt_DAOB5 = Sheet10.Range("DAOB5.HalfRateDeprAmt").text
End Function

Function ValidateDepreciationAtFullRate_DAOB5() As Boolean
    ValidateDepreciationAtFullRate_DAOB5 = True
    DepreciationAtFullRate_DAOB5 = Sheet10.Range("DAOB5.DepreciationAtFullRate").text
End Function

Function ValidateDepreciationAtHalfRate_DAOB5() As Boolean
    ValidateDepreciationAtHalfRate_DAOB5 = True
    DepreciationAtHalfRate_DAOB5 = Sheet10.Range("DAOB5.DepreciationAtHalfRate").text
End Function

Function ValidateAddlnDeprOnGT180DayAdditions_DAOB5() As Boolean
    ValidateAddlnDeprOnGT180DayAdditions_DAOB5 = True
    AddlnDeprOnGT180DayAdditions_DAOB5 = Sheet10.Range("DAOB5.AddlnDeprOnGT180DayAdditions").text
End Function

Function ValidateAddlnDeprDuringYearAdditions_DAOB5() As Boolean
    ValidateAddlnDeprDuringYearAdditions_DAOB5 = True
    AddlnDeprDuringYearAdditions_DAOB5 = Sheet10.Range("DAOB5.AddlnDeprDuringYearAdditions").text
End Function

Function ValidateTotalDepreciation_DAOB5() As Boolean
    ValidateTotalDepreciation_DAOB5 = True
    TotalDepreciation_DAOB5 = Sheet10.Range("DAOB5.TotalDepreciation").text
End Function

Function ValidateExpdrOnTrforSaleAsset_DAOB5() As Boolean
    ValidateExpdrOnTrforSaleAsset_DAOB5 = True
    ExpdrOnTrforSaleAsset_DAOB5 = Sheet10.Range("DAOB5.ExpdrOnTrforSaleAsset").text
End Function
Function ValidateCapGainUs50_DAOB5() As Boolean
    ValidateCapGainUs50_DAOB5 = True
    CapGainUs50_DAOB5 = Sheet10.Range("DAOB5.CapGainUs50").value
    If Sheet10.Range("DAOB5.Check17A").value > 0 Then
    If Sheet10.Range("DAOB5.CapGainUs50") = "" Or Sheet10.Range("DAOB5.CapGainUs50") = 0 Then
    Sheet9.Activate
    ValidateCapGainUs50_DAOB5 = False
    msgValidateSheetDAOB5 = msgValidateSheetDAOB5 & Chr(13) & "DOA5:Capital gains/ loss under section 50 can not be Blank"
    End If
    End If
    If CapGainUs50_DAOB5 <> 0 Then
       Sheet10.Activate
       Sheet10.cg_gt_lt_0 (CapGainUs50_DAOB5), "DAOB5.CapGainUs50"
       fmsgboxStatus "Block Depreciation - Building @ 5% ceases to exist"
    End If
End Function

Function ValidateCapGainUs50_DAOB5old() As Boolean
    ValidateCapGainUs50_DAOB5 = True
    CapGainUs50_DAOB5 = Sheet10.Range("DAOB5.CapGainUs50").value
    
    If CapGainUs50_DAOB5 <> 0 Then
       Sheet10.Activate
       Sheet10.cg_gt_lt_0 (CapGainUs50_DAOB5), "DAOB5.CapGainUs50"
       fmsgboxStatus "Block Depreciation - Building @ 5% ceases to exist"
    End If
End Function

Function ValidateWDVLastDay_DAOB5() As Boolean
    ValidateWDVLastDay_DAOB5 = True
    WDVLastDay_DAOB5 = Sheet10.Range("DAOB5.WDVLastDay").text
End Function

Function ValidateWDVFirstDay_DAOB10() As Boolean
    ValidateWDVFirstDay_DAOB10 = True
    WDVFirstDay_DAOB10 = Sheet10.Range("DAOB10.WDVFirstDay").text
End Function

Function ValidateAdditionsGrThan180Days_DAOB10() As Boolean
    ValidateAdditionsGrThan180Days_DAOB10 = True
    AdditionsGrThan180Days_DAOB10 = Sheet10.Range("DAOB10.AdditionsGrThan180Days").text
End Function

Function ValidateRealizationTotalPeriod_DAOB10() As Boolean
    ValidateRealizationTotalPeriod_DAOB10 = True
    RealizationTotalPeriod_DAOB10 = Sheet10.Range("DAOB10.RealizationTotalPeriod").text
End Function

Function ValidateFullRateDeprAmt_DAOB10() As Boolean
    ValidateFullRateDeprAmt_DAOB10 = True
    FullRateDeprAmt_DAOB10 = Sheet10.Range("DAOB10.FullRateDeprAmt").text
End Function

Function ValidateAdditionsLessThan180Days_DAOB10() As Boolean
    ValidateAdditionsLessThan180Days_DAOB10 = True
    AdditionsLessThan180Days_DAOB10 = Sheet10.Range("DAOB10.AdditionsLessThan180Days").text
End Function

Function ValidateRealizationPeriodDuringYear_DAOB10() As Boolean
    ValidateRealizationPeriodDuringYear_DAOB10 = True
    RealizationPeriodDuringYear_DAOB10 = Sheet10.Range("DAOB10.RealizationPeriodDuringYear").text
End Function

Function ValidateHalfRateDeprAmt_DAOB10() As Boolean
    ValidateHalfRateDeprAmt_DAOB10 = True
    HalfRateDeprAmt_DAOB10 = Sheet10.Range("DAOB10.HalfRateDeprAmt").text
End Function

Function ValidateDepreciationAtFullRate_DAOB10() As Boolean
    ValidateDepreciationAtFullRate_DAOB10 = True
    DepreciationAtFullRate_DAOB10 = Sheet10.Range("DAOB10.DepreciationAtFullRate").text
End Function

Function ValidateDepreciationAtHalfRate_DAOB10() As Boolean
    ValidateDepreciationAtHalfRate_DAOB10 = True
    DepreciationAtHalfRate_DAOB10 = Sheet10.Range("DAOB10.DepreciationAtHalfRate").text
End Function

Function ValidateAddlnDeprOnGT180DayAdditions_DAOB10() As Boolean
    ValidateAddlnDeprOnGT180DayAdditions_DAOB10 = True
    AddlnDeprOnGT180DayAdditions_DAOB10 = Sheet10.Range("DAOB10.AddlnDeprOnGT180DayAdditions").text
End Function

Function ValidateAddlnDeprDuringYearAdditions_DAOB10() As Boolean
    ValidateAddlnDeprDuringYearAdditions_DAOB10 = True
    AddlnDeprDuringYearAdditions_DAOB10 = Sheet10.Range("DAOB10.AddlnDeprDuringYearAdditions").text
End Function

Function ValidateTotalDepreciation_DAOB10() As Boolean
    ValidateTotalDepreciation_DAOB10 = True
    TotalDepreciation_DAOB10 = Sheet10.Range("DAOB10.TotalDepreciation").text
End Function
Function ValidateDepreciation_Sec38_DAOB05() As Boolean
    ValidateDepreciation_Sec38_DAOB05 = True
    Depreciation_Sec38_DAOB05 = Sheet10.Range("DAOB5.DepUnderSec38").text
    NetaggregateDep_DAOB05 = Sheet10.Range("DAOB5.NetAgrdep").text
    PropaggregateDep_DAOB05 = Sheet10.Range("DAOB5.ProAgrdep").text
End Function
Function ValidateDepreciation_Sec38_DAOB10() As Boolean
    ValidateDepreciation_Sec38_DAOB10 = True
    Depreciation_Sec38_DAOB10 = Sheet10.Range("DAOB10.DepUnderSec38").text
    NetaggregateDep_DAOB10 = Sheet10.Range("DAOB10.NetAgrdep").text
    PropaggregateDep_DAOB10 = Sheet10.Range("DAOB10.ProAgrdep").text
End Function
Function ValidateDepreciation_Sec38_DAOB40() As Boolean
    ValidateDepreciation_Sec38_DAOB40 = True
    Depreciation_Sec38_DAOB40 = Sheet10.Range("DAOB100.DepUnderSec38").text
    NetaggregateDep_DAOB40 = Sheet10.Range("DAOB100.NetAgrdep").text
    PropaggregateDep_DAOB40 = Sheet10.Range("DAOB100.ProAgrdep").text
End Function
Function ValidateDepreciation_Sec38_DAOF10() As Boolean
    ValidateDepreciation_Sec38_DAOF10 = True
    Depreciation_Sec38_DAOF10 = Sheet10.Range("DAOF10.DepUnderSec38").text
    NetaggregateDep_DAOF10 = Sheet10.Range("DAOF10.NetAgrdep").text
    PropaggregateDep_DAOF10 = Sheet10.Range("DAOF10.ProAgrdep").text
End Function
Function ValidateDepreciation_Sec38_DAOI25() As Boolean
    ValidateDepreciation_Sec38_DAOI25 = True
    Depreciation_Sec38_DAOI25 = Sheet10.Range("DAOI25.DepUnderSec38").text
    NetaggregateDep_DAOI25 = Sheet10.Range("DAOI25.NetAgrdep").text
    PropaggregateDep_DAOI25 = Sheet10.Range("DAOI25.ProAgrdep").text
End Function
Function ValidateDepreciation_Sec38_DAOS20() As Boolean
    ValidateDepreciation_Sec38_DAOS20 = True
    Depreciation_Sec38_DAOS20 = Sheet10.Range("DAOS20.DepUnderSec38").text
    NetaggregateDep_DAOS20 = Sheet10.Range("DAOS20.NetAgrdep").text
    PropaggregateDep_DAOS20 = Sheet10.Range("DAOS20.ProAgrdep").text
End Function
Function ValidateExpdrOnTrforSaleAsset_DAOB10() As Boolean
    ValidateExpdrOnTrforSaleAsset_DAOB10 = True
    ExpdrOnTrforSaleAsset_DAOB10 = Sheet10.Range("DAOB10.ExpdrOnTrforSaleAsset").text
End Function
Function ValidateCapGainUs50_DAOB10() As Boolean
    ValidateCapGainUs50_DAOB10 = True
    CapGainUs50_DAOB10 = Sheet10.Range("DAOB10.CapGainUs50").value
    If Sheet10.Range("DAOB10.Check17A").value > 0 Then
    If Sheet10.Range("DAOB10.CapGainUs50") = "" Or Sheet10.Range("DAOB10.CapGainUs50") = 0 Then
    Sheet9.Activate
    ValidateCapGainUs50_DAOB10 = False
    msgValidateSheetDAOB5 = msgValidateSheetDAOB5 & Chr(13) & "DOA10:Capital gains/ loss under section 50 can not be Blank"
    End If
    End If
    If CapGainUs50_DAOB10 <> 0 Then
        Sheet10.Activate
        Sheet10.cg_gt_lt_0 (CapGainUs50_DAOB10), "DAOB10.CapGainUs50"
        fmsgboxStatus "Block Depreciation - Building @ 10% ceases to exist"
    End If
End Function
Function ValidateCapGainUs50_DAOB10old() As Boolean
    ValidateCapGainUs50_DAOB10 = True
    CapGainUs50_DAOB10 = Sheet10.Range("DAOB10.CapGainUs50").value
    
    If CapGainUs50_DAOB10 <> 0 Then
        Sheet10.Activate
        Sheet10.cg_gt_lt_0 (CapGainUs50_DAOB10), "DAOB10.CapGainUs50"
        fmsgboxStatus "Block Depreciation - Building @ 10% ceases to exist"
    End If
End Function

Function ValidateWDVLastDay_DAOB10() As Boolean
    ValidateWDVLastDay_DAOB10 = True
    WDVLastDay_DAOB10 = Sheet10.Range("DAOB10.WDVLastDay").text
End Function

Function ValidateWDVFirstDay_DAOB100() As Boolean
    ValidateWDVFirstDay_DAOB100 = True
    WDVFirstDay_DAOB100 = Sheet10.Range("DAOB100.WDVFirstDay").text
End Function

Function ValidateAdditionsGrThan180Days_DAOB100() As Boolean
    ValidateAdditionsGrThan180Days_DAOB100 = True
    AdditionsGrThan180Days_DAOB100 = Sheet10.Range("DAOB100.AdditionsGrThan180Days").text
End Function

Function ValidateRealizationTotalPeriod_DAOB100() As Boolean
    ValidateRealizationTotalPeriod_DAOB100 = True
    RealizationTotalPeriod_DAOB100 = Sheet10.Range("DAOB100.RealizationTotalPeriod").text
End Function

Function ValidateFullRateDeprAmt_DAOB100() As Boolean
    ValidateFullRateDeprAmt_DAOB100 = True
    FullRateDeprAmt_DAOB100 = Sheet10.Range("DAOB100.FullRateDeprAmt").text
End Function

Function ValidateAdditionsLessThan180Days_DAOB100() As Boolean
    ValidateAdditionsLessThan180Days_DAOB100 = True
    AdditionsLessThan180Days_DAOB100 = Sheet10.Range("DAOB100.AdditionsLessThan180Days").text
End Function

Function ValidateRealizationPeriodDuringYear_DAOB100() As Boolean
    ValidateRealizationPeriodDuringYear_DAOB100 = True
    RealizationPeriodDuringYear_DAOB100 = Sheet10.Range("DAOB100.RealizationPeriodDuringYear").text
End Function

Function ValidateHalfRateDeprAmt_DAOB100() As Boolean
    ValidateHalfRateDeprAmt_DAOB100 = True
    HalfRateDeprAmt_DAOB100 = Sheet10.Range("DAOB100.HalfRateDeprAmt").text
End Function

Function ValidateDepreciationAtFullRate_DAOB100() As Boolean
    ValidateDepreciationAtFullRate_DAOB100 = True
    DepreciationAtFullRate_DAOB100 = Sheet10.Range("DAOB100.DepreciationAtFullRate").text
End Function

Function ValidateDepreciationAtHalfRate_DAOB100() As Boolean
    ValidateDepreciationAtHalfRate_DAOB100 = True
    DepreciationAtHalfRate_DAOB100 = Sheet10.Range("DAOB100.DepreciationAtHalfRate").text
End Function

Function ValidateAddlnDeprOnGT180DayAdditions_DAOB100() As Boolean
    ValidateAddlnDeprOnGT180DayAdditions_DAOB100 = True
    AddlnDeprOnGT180DayAdditions_DAOB100 = Sheet10.Range("DAOB100.AddlnDeprOnGT180DayAdditions").text
End Function

Function ValidateAddlnDeprDuringYearAdditions_DAOB100() As Boolean
    ValidateAddlnDeprDuringYearAdditions_DAOB100 = True
    AddlnDeprDuringYearAdditions_DAOB100 = Sheet10.Range("DAOB100.AddlnDeprDuringYearAdditions").value
End Function

Function ValidateTotalDepreciation_DAOB100() As Boolean
    ValidateTotalDepreciation_DAOB100 = True
    TotalDepreciation_DAOB100 = Sheet10.Range("DAOB100.TotalDepreciation").text
End Function

Function ValidateExpdrOnTrforSaleAsset_DAOB100() As Boolean
    ValidateExpdrOnTrforSaleAsset_DAOB100 = True
    ExpdrOnTrforSaleAsset_DAOB100 = Sheet10.Range("DAOB100.ExpdrOnTrforSaleAsset").text
End Function
Function ValidateCapGainUs50_DAOB100() As Boolean
 ValidateCapGainUs50_DAOB100 = True
 CapGainUs50_DAOB100 = Sheet10.Range("DAOB100.CapGainUs50").value
    If Sheet10.Range("DAOB40.Check17A").value > 0 Then
    If Sheet10.Range("DAOB100.CapGainUs50") = "" Or Sheet10.Range("DAOB100.CapGainUs50") = 0 Then
    Sheet9.Activate
    ValidateCapGainUs50_DAOB100 = False
    msgValidateSheetDAOB5 = msgValidateSheetDAOB5 & Chr(13) & "DOA40:Capital gains/ loss under section 50 can not be Blank"
    End If
    End If
 If CapGainUs50_DAOB100 <> 0 Then
    Sheet10.Activate
    Sheet10.cg_gt_lt_0 (CapGainUs50_DAOB100), "DAOB100.CapGainUs50"
    fmsgboxStatus "Block Depreciation - Building @ 40% ceases to exist"
 End If
End Function

Function ValidateCapGainUs50_DAOB100old() As Boolean
 ValidateCapGainUs50_DAOB100 = True
 CapGainUs50_DAOB100 = Sheet10.Range("DAOB100.CapGainUs50").value

 If CapGainUs50_DAOB100 <> 0 Then
    Sheet10.Activate
    Sheet10.cg_gt_lt_0 (CapGainUs50_DAOB100), "DAOB100.CapGainUs50"
    fmsgboxStatus "Block Depreciation - Building @ 40% ceases to exist"
 End If
End Function

Function ValidateWDVLastDay_DAOB100() As Boolean
    ValidateWDVLastDay_DAOB100 = True
    WDVLastDay_DAOB100 = Sheet10.Range("DAOB100.WDVLastDay").text
End Function

Function ValidateWDVFirstDay_DAOF10() As Boolean
    ValidateWDVFirstDay_DAOF10 = True
    WDVFirstDay_DAOF10 = Sheet10.Range("DAOF10.WDVFirstDay").text
End Function

Function ValidateAdditionsGrThan180Days_DAOF10() As Boolean
    ValidateAdditionsGrThan180Days_DAOF10 = True
    AdditionsGrThan180Days_DAOF10 = Sheet10.Range("DAOF10.AdditionsGrThan180Days").text
End Function

Function ValidateRealizationTotalPeriod_DAOF10() As Boolean
    ValidateRealizationTotalPeriod_DAOF10 = True
    RealizationTotalPeriod_DAOF10 = Sheet10.Range("DAOF10.RealizationTotalPeriod").text
End Function

Function ValidateFullRateDeprAmt_DAOF10() As Boolean
    ValidateFullRateDeprAmt_DAOF10 = True
    FullRateDeprAmt_DAOF10 = Sheet10.Range("DAOF10.FullRateDeprAmt").text
End Function

Function ValidateAdditionsLessThan180Days_DAOF10() As Boolean
    ValidateAdditionsLessThan180Days_DAOF10 = True
    AdditionsLessThan180Days_DAOF10 = Sheet10.Range("DAOF10.AdditionsLessThan180Days").text
End Function

Function ValidateRealizationPeriodDuringYear_DAOF10() As Boolean
    ValidateRealizationPeriodDuringYear_DAOF10 = True
    RealizationPeriodDuringYear_DAOF10 = Sheet10.Range("DAOF10.RealizationPeriodDuringYear").text
End Function

Function ValidateHalfRateDeprAmt_DAOF10() As Boolean
    ValidateHalfRateDeprAmt_DAOF10 = True
    HalfRateDeprAmt_DAOF10 = Sheet10.Range("DAOF10.HalfRateDeprAmt").text
End Function

Function ValidateDepreciationAtFullRate_DAOF10() As Boolean
    ValidateDepreciationAtFullRate_DAOF10 = True
    DepreciationAtFullRate_DAOF10 = Sheet10.Range("DAOF10.DepreciationAtFullRate").text
End Function

Function ValidateDepreciationAtHalfRate_DAOF10() As Boolean
    ValidateDepreciationAtHalfRate_DAOF10 = True
    DepreciationAtHalfRate_DAOF10 = Sheet10.Range("DAOF10.DepreciationAtHalfRate").text
End Function

Function ValidateAddlnDeprOnGT180DayAdditions_DAOF10() As Boolean
    ValidateAddlnDeprOnGT180DayAdditions_DAOF10 = True
    AddlnDeprOnGT180DayAdditions_DAOF10 = Sheet10.Range("DAOF10.AddlnDeprOnGT180DayAdditions").text
End Function

Function ValidateAddlnDeprDuringYearAdditions_DAOF10() As Boolean
    ValidateAddlnDeprDuringYearAdditions_DAOF10 = True
    AddlnDeprDuringYearAdditions_DAOF10 = Sheet10.Range("DAOF10.AddlnDeprDuringYearAdditions").text
End Function

Function ValidateTotalDepreciation_DAOF10() As Boolean
    ValidateTotalDepreciation_DAOF10 = True
    TotalDepreciation_DAOF10 = Sheet10.Range("DAOF10.TotalDepreciation").text
End Function

Function ValidateExpdrOnTrforSaleAsset_DAOF10() As Boolean
    ValidateExpdrOnTrforSaleAsset_DAOF10 = True
    ExpdrOnTrforSaleAsset_DAOF10 = Sheet10.Range("DAOF10.ExpdrOnTrforSaleAsset").text
End Function
Function ValidateCapGainUs50_DAOF10() As Boolean
    ValidateCapGainUs50_DAOF10 = True
    CapGainUs50_DAOF10 = Sheet10.Range("DAOF10.CapGainUs50").value
    If Sheet10.Range("DAOB10.Check17AFurniture").value > 0 Then
    If Sheet10.Range("DAOF10.CapGainUs50") = "" Or Sheet10.Range("DAOF10.CapGainUs50") = 0 Then
    Sheet9.Activate
    ValidateCapGainUs50_DAOF10 = False
    msgValidateSheetDAOB5 = msgValidateSheetDAOB5 & Chr(13) & "DOA10Furniture:Capital gains/ loss under section 50 can not be Blank"
    End If
    End If
    If CapGainUs50_DAOF10 <> 0 Then
       Sheet10.Activate
       Sheet10.cg_gt_lt_0 (CapGainUs50_DAOF10), "DAOF10.CapGainUs50"
       fmsgboxStatus "Block Depreciation - Furniture & Fittings ceases to exist"
    End If
End Function
Function ValidateCapGainUs50_DAOF10old() As Boolean
    ValidateCapGainUs50_DAOF10 = True
    CapGainUs50_DAOF10 = Sheet10.Range("DAOF10.CapGainUs50").value
    If CapGainUs50_DAOF10 <> 0 Then
       Sheet10.Activate
       Sheet10.cg_gt_lt_0 (CapGainUs50_DAOF10), "DAOF10.CapGainUs50"
       fmsgboxStatus "Block Depreciation - Furniture & Fittings ceases to exist"
    End If
End Function

Function ValidateWDVLastDay_DAOF10() As Boolean
    ValidateWDVLastDay_DAOF10 = True
    WDVLastDay_DAOF10 = Sheet10.Range("DAOF10.WDVLastDay").text
End Function

Function ValidateWDVFirstDay_DAOI25() As Boolean
    ValidateWDVFirstDay_DAOI25 = True
    WDVFirstDay_DAOI25 = Sheet10.Range("DAOI25.WDVFirstDay").text
End Function

Function ValidateAdditionsGrThan180Days_DAOI25() As Boolean
    ValidateAdditionsGrThan180Days_DAOI25 = True
    AdditionsGrThan180Days_DAOI25 = Sheet10.Range("DAOI25.AdditionsGrThan180Days").text
End Function

Function ValidateRealizationTotalPeriod_DAOI25() As Boolean
    ValidateRealizationTotalPeriod_DAOI25 = True
    RealizationTotalPeriod_DAOI25 = Sheet10.Range("DAOI25.RealizationTotalPeriod").text
End Function

Function ValidateFullRateDeprAmt_DAOI25() As Boolean
    ValidateFullRateDeprAmt_DAOI25 = True
    FullRateDeprAmt_DAOI25 = Sheet10.Range("DAOI25.FullRateDeprAmt").text
End Function

Function ValidateAdditionsLessThan180Days_DAOI25() As Boolean
    ValidateAdditionsLessThan180Days_DAOI25 = True
    AdditionsLessThan180Days_DAOI25 = Sheet10.Range("DAOI25.AdditionsLessThan180Days").text
End Function

Function ValidateRealizationPeriodDuringYear_DAOI25() As Boolean
    ValidateRealizationPeriodDuringYear_DAOI25 = True
    RealizationPeriodDuringYear_DAOI25 = Sheet10.Range("DAOI25.RealizationPeriodDuringYear").text
End Function

Function ValidateHalfRateDeprAmt_DAOI25() As Boolean
    ValidateHalfRateDeprAmt_DAOI25 = True
    HalfRateDeprAmt_DAOI25 = Sheet10.Range("DAOI25.HalfRateDeprAmt").text
End Function

Function ValidateDepreciationAtFullRate_DAOI25() As Boolean
    ValidateDepreciationAtFullRate_DAOI25 = True
    DepreciationAtFullRate_DAOI25 = Sheet10.Range("DAOI25.DepreciationAtFullRate").text
End Function

Function ValidateDepreciationAtHalfRate_DAOI25() As Boolean
    ValidateDepreciationAtHalfRate_DAOI25 = True
    DepreciationAtHalfRate_DAOI25 = Sheet10.Range("DAOI25.DepreciationAtHalfRate").text
End Function

Function ValidateAddlnDeprOnGT180DayAdditions_DAOI25() As Boolean
    ValidateAddlnDeprOnGT180DayAdditions_DAOI25 = True
    AddlnDeprOnGT180DayAdditions_DAOI25 = Sheet10.Range("DAOI25.AddlnDeprOnGT180DayAdditions").text
End Function

Function ValidateAddlnDeprDuringYearAdditions_DAOI25() As Boolean
    ValidateAddlnDeprDuringYearAdditions_DAOI25 = True
    AddlnDeprDuringYearAdditions_DAOI25 = Sheet10.Range("DAOI25.AddlnDeprDuringYearAdditions").text
End Function

Function ValidateTotalDepreciation_DAOI25() As Boolean
    ValidateTotalDepreciation_DAOI25 = True
    TotalDepreciation_DAOI25 = Sheet10.Range("DAOI25.TotalDepreciation").text
End Function

Function ValidateExpdrOnTrforSaleAsset_DAOI25() As Boolean
    ValidateExpdrOnTrforSaleAsset_DAOI25 = True
    ExpdrOnTrforSaleAsset_DAOI25 = Sheet10.Range("DAOI25.ExpdrOnTrforSaleAsset").text
End Function
Function ValidateCapGainUs50_DAOI25() As Boolean
 ValidateCapGainUs50_DAOI25 = True
 CapGainUs50_DAOI25 = Sheet10.Range("DAOI25.CapGainUs50").value
    If Sheet10.Range("DAOB25.Check17A").value > 0 Then
    If Sheet10.Range("DAOI25.CapGainUs50") = "" Or Sheet10.Range("DAOI25.CapGainUs50") = 0 Then
    Sheet9.Activate
    ValidateCapGainUs50_DAOI25 = False
    msgValidateSheetDAOB5 = msgValidateSheetDAOB5 & Chr(13) & "DOA25:Capital gains/ loss under section 50 can not be Blank"
    End If
    End If
 If CapGainUs50_DAOI25 <> 0 Then
    Sheet10.Activate
    Sheet10.cg_gt_lt_0 (CapGainUs50_DAOI25), "DAOI25.CapGainUs50"
    fmsgboxStatus "Block Depreciation - Intangible ceases to exist"
 End If
End Function

Function ValidateCapGainUs50_DAOI25old() As Boolean
 ValidateCapGainUs50_DAOI25 = True
 CapGainUs50_DAOI25 = Sheet10.Range("DAOI25.CapGainUs50").value

 If CapGainUs50_DAOI25 <> 0 Then
    Sheet10.Activate
    Sheet10.cg_gt_lt_0 (CapGainUs50_DAOI25), "DAOI25.CapGainUs50"
    fmsgboxStatus "Block Depreciation - Intangible ceases to exist"
 End If
End Function

Function ValidateWDVLastDay_DAOI25() As Boolean
    ValidateWDVLastDay_DAOI25 = True
    WDVLastDay_DAOI25 = Sheet10.Range("DAOI25.WDVLastDay").text
End Function

Function ValidateWDVFirstDay_DAOS20() As Boolean
    ValidateWDVFirstDay_DAOS20 = True
    WDVFirstDay_DAOS20 = Sheet10.Range("DAOS20.WDVFirstDay").text
End Function

Function ValidateAdditionsGrThan180Days_DAOS20() As Boolean
    ValidateAdditionsGrThan180Days_DAOS20 = True
    AdditionsGrThan180Days_DAOS20 = Sheet10.Range("DAOS20.AdditionsGrThan180Days").text
End Function

Function ValidateRealizationTotalPeriod_DAOS20() As Boolean
    ValidateRealizationTotalPeriod_DAOS20 = True
    RealizationTotalPeriod_DAOS20 = Sheet10.Range("DAOS20.RealizationTotalPeriod").text
End Function

Function ValidateFullRateDeprAmt_DAOS20() As Boolean
    ValidateFullRateDeprAmt_DAOS20 = True
    FullRateDeprAmt_DAOS20 = Sheet10.Range("DAOS20.FullRateDeprAmt").text
End Function

Function ValidateAdditionsLessThan180Days_DAOS20() As Boolean
    ValidateAdditionsLessThan180Days_DAOS20 = True
    AdditionsLessThan180Days_DAOS20 = Sheet10.Range("DAOS20.AdditionsLessThan180Days").text
End Function

Function ValidateRealizationPeriodDuringYear_DAOS20() As Boolean
    ValidateRealizationPeriodDuringYear_DAOS20 = True
    RealizationPeriodDuringYear_DAOS20 = Sheet10.Range("DAOS20.RealizationPeriodDuringYear").text
End Function

Function ValidateHalfRateDeprAmt_DAOS20() As Boolean
    ValidateHalfRateDeprAmt_DAOS20 = True
    HalfRateDeprAmt_DAOS20 = Sheet10.Range("DAOS20.HalfRateDeprAmt").text
End Function

Function ValidateDepreciationAtFullRate_DAOS20() As Boolean
    ValidateDepreciationAtFullRate_DAOS20 = True
    DepreciationAtFullRate_DAOS20 = Sheet10.Range("DAOS20.DepreciationAtFullRate").text
End Function

Function ValidateDepreciationAtHalfRate_DAOS20() As Boolean
    ValidateDepreciationAtHalfRate_DAOS20 = True
    DepreciationAtHalfRate_DAOS20 = Sheet10.Range("DAOS20.DepreciationAtHalfRate").text
End Function

Function ValidateAddlnDeprOnGT180DayAdditions_DAOS20() As Boolean
    ValidateAddlnDeprOnGT180DayAdditions_DAOS20 = True
    AddlnDeprOnGT180DayAdditions_DAOS20 = Sheet10.Range("DAOS20.AddlnDeprOnGT180DayAdditions").text
End Function

Function ValidateAddlnDeprDuringYearAdditions_DAOS20() As Boolean
    ValidateAddlnDeprDuringYearAdditions_DAOS20 = True
    AddlnDeprDuringYearAdditions_DAOS20 = Sheet10.Range("DAOS20.AddlnDeprDuringYearAdditions").text
End Function

Function ValidateTotalDepreciation_DAOS20() As Boolean
    ValidateTotalDepreciation_DAOS20 = True
    TotalDepreciation_DAOS20 = Sheet10.Range("DAOS20.TotalDepreciation").text
End Function

Function ValidateExpdrOnTrforSaleAsset_DAOS20() As Boolean
    ValidateExpdrOnTrforSaleAsset_DAOS20 = True
    ExpdrOnTrforSaleAsset_DAOS20 = Sheet10.Range("DAOS20.ExpdrOnTrforSaleAsset").text
End Function
Function ValidateCapGainUs50_DAOS20() As Boolean
    ValidateCapGainUs50_DAOS20 = True
    CapGainUs50_DAOS20 = Sheet10.Range("DAOS20.CapGainUs50").value
    If Sheet10.Range("DAOB20.Check17A").value > 0 Then
    If Sheet10.Range("DAOS20.CapGainUs50") = "" Or Sheet10.Range("DAOS20.CapGainUs50") = 0 Then
    Sheet9.Activate
    ValidateCapGainUs50_DAOS20 = False
    msgValidateSheetDAOB5 = msgValidateSheetDAOB5 & Chr(13) & "DOA20:Capital gains/ loss under section 50 can not be Blank"
    End If
    End If
    If CapGainUs50_DAOS20 <> 0 Then
        Sheet10.Activate
        Sheet10.cg_gt_lt_0 (CapGainUs50_DAOS20), "DAOS20.CapGainUs50"
        fmsgboxStatus "Block Depreciation - Ships ceases to exist"
    End If
End Function
Function ValidateCapGainUs50_DAOS20old() As Boolean
    ValidateCapGainUs50_DAOS20 = True
    CapGainUs50_DAOS20 = Sheet10.Range("DAOS20.CapGainUs50").value
    
    If CapGainUs50_DAOS20 <> 0 Then
        Sheet10.Activate
        Sheet10.cg_gt_lt_0 (CapGainUs50_DAOS20), "DAOS20.CapGainUs50"
        fmsgboxStatus "Block Depreciation - Ships ceases to exist"
    End If
End Function

Function ValidateWDVLastDay_DAOS20() As Boolean
    ValidateWDVLastDay_DAOS20 = True
    WDVLastDay_DAOS20 = Sheet10.Range("DAOS20.WDVLastDay").text
End Function

Function DefaultWDVFirstDay_DPM15() As String
DefaultWDVFirstDay_DPM15 = "0"
End Function
Function DefaultAdditionsGrThan180Days_DPM15() As String
DefaultAdditionsGrThan180Days_DPM15 = "0"
End Function
Function DefaultRealizationTotalPeriod_DPM15() As String
DefaultRealizationTotalPeriod_DPM15 = "0"
End Function
Function DefaultFullRateDeprAmt_DPM15() As String
DefaultFullRateDeprAmt_DPM15 = "0"
End Function
Function DefaultAdditionsLessThan180Days_DPM15() As String
DefaultAdditionsLessThan180Days_DPM15 = "0"
End Function
Function DefaultRealizationPeriodDuringYear_DPM15() As String
DefaultRealizationPeriodDuringYear_DPM15 = "0"
End Function
Function DefaultHalfRateDeprAmt_DPM15() As String
DefaultHalfRateDeprAmt_DPM15 = "0"
End Function
Function DefaultDepreciationAtFullRate_DPM15() As String
DefaultDepreciationAtFullRate_DPM15 = "0"
End Function
Function DefaultDepreciationAtHalfRate_DPM15() As String
DefaultDepreciationAtHalfRate_DPM15 = "0"
End Function
Function DefaultAddlnDeprOnGT180DayAdditions_DPM15() As String
DefaultAddlnDeprOnGT180DayAdditions_DPM15 = "0"
End Function
Function DefaultAddlnDeprDuringYearAdditions_DPM15() As String
DefaultAddlnDeprDuringYearAdditions_DPM15 = "0"
End Function
Function DefaultTotalDepreciation_DPM15() As String
DefaultTotalDepreciation_DPM15 = "0"
End Function
Function DefaultExpdrOnTrforSaleAsset_DPM15() As String
DefaultExpdrOnTrforSaleAsset_DPM15 = "0"
End Function
Function DefaultCapGainUs50_DPM15() As String
DefaultCapGainUs50_DPM15 = "0"
End Function
Function DefaultWDVLastDay_DPM15() As String
DefaultWDVLastDay_DPM15 = "0"
End Function
Function DefaultWDVFirstDay_DPM30() As String
DefaultWDVFirstDay_DPM30 = "0"
End Function
Function DefaultAdditionsGrThan180Days_DPM30() As String
DefaultAdditionsGrThan180Days_DPM30 = "0"
End Function
Function DefaultRealizationTotalPeriod_DPM30() As String
DefaultRealizationTotalPeriod_DPM30 = "0"
End Function
Function DefaultFullRateDeprAmt_DPM30() As String
DefaultFullRateDeprAmt_DPM30 = "0"
End Function
Function DefaultAdditionsLessThan180Days_DPM30() As String
DefaultAdditionsLessThan180Days_DPM30 = "0"
End Function
Function DefaultRealizationPeriodDuringYear_DPM30() As String
DefaultRealizationPeriodDuringYear_DPM30 = "0"
End Function
Function DefaultHalfRateDeprAmt_DPM30() As String
DefaultHalfRateDeprAmt_DPM30 = "0"
End Function
Function DefaultDepreciationAtFullRate_DPM30() As String
DefaultDepreciationAtFullRate_DPM30 = "0"
End Function
Function DefaultDepreciationAtHalfRate_DPM30() As String
DefaultDepreciationAtHalfRate_DPM30 = "0"
End Function
Function DefaultAddlnDeprOnGT180DayAdditions_DPM30() As String
DefaultAddlnDeprOnGT180DayAdditions_DPM30 = "0"
End Function
Function DefaultAddlnDeprDuringYearAdditions_DPM30() As String
DefaultAddlnDeprDuringYearAdditions_DPM30 = "0"
End Function
Function DefaultTotalDepreciation_DPM30() As String
DefaultTotalDepreciation_DPM30 = "0"
End Function
Function DefaultExpdrOnTrforSaleAsset_DPM30() As String
DefaultExpdrOnTrforSaleAsset_DPM30 = "0"
End Function
Function DefaultCapGainUs50_DPM30() As String
DefaultCapGainUs50_DPM30 = "0"
End Function
Function DefaultWDVLastDay_DPM30() As String
DefaultWDVLastDay_DPM30 = "0"
End Function
Function DefaultWDVFirstDay_DPM40() As String
DefaultWDVFirstDay_DPM40 = "0"
End Function
Function DefaultAdditionsGrThan180Days_DPM40() As String
DefaultAdditionsGrThan180Days_DPM40 = "0"
End Function
Function DefaultRealizationTotalPeriod_DPM40() As String
DefaultRealizationTotalPeriod_DPM40 = "0"
End Function
Function DefaultFullRateDeprAmt_DPM40() As String
DefaultFullRateDeprAmt_DPM40 = "0"
End Function
Function DefaultAdditionsLessThan180Days_DPM40() As String
DefaultAdditionsLessThan180Days_DPM40 = "0"
End Function
Function DefaultRealizationPeriodDuringYear_DPM40() As String
DefaultRealizationPeriodDuringYear_DPM40 = "0"
End Function
Function DefaultHalfRateDeprAmt_DPM40() As String
DefaultHalfRateDeprAmt_DPM40 = "0"
End Function
Function DefaultDepreciationAtFullRate_DPM40() As String
DefaultDepreciationAtFullRate_DPM40 = "0"
End Function
Function DefaultDepreciationAtHalfRate_DPM40() As String
DefaultDepreciationAtHalfRate_DPM40 = "0"
End Function
Function DefaultAddlnDeprOnGT180DayAdditions_DPM40() As String
DefaultAddlnDeprOnGT180DayAdditions_DPM40 = "0"
End Function
Function DefaultAddlnDeprDuringYearAdditions_DPM40() As String
DefaultAddlnDeprDuringYearAdditions_DPM40 = "0"
End Function
Function DefaultTotalDepreciation_DPM40() As String
DefaultTotalDepreciation_DPM40 = "0"
End Function
Function DefaultExpdrOnTrforSaleAsset_DPM40() As String
DefaultExpdrOnTrforSaleAsset_DPM40 = "0"
End Function
Function DefaultCapGainUs50_DPM40() As String
DefaultCapGainUs50_DPM40 = "0"
End Function
Function DefaultWDVLastDay_DPM40() As String
DefaultWDVLastDay_DPM40 = "0"
End Function

Function DefaultWDVFirstDay_DAOL0() As String
DefaultWDVFirstDay_DAOL0 = "0"
End Function
Function DefaultWDVLastDay_DAOL0() As String
DefaultWDVLastDay_DAOL0 = "0"
End Function
Function DefaultWDVFirstDay_DAOB5() As String
DefaultWDVFirstDay_DAOB5 = "0"
End Function
Function DefaultAdditionsGrThan180Days_DAOB5() As String
DefaultAdditionsGrThan180Days_DAOB5 = "0"
End Function
Function DefaultRealizationTotalPeriod_DAOB5() As String
DefaultRealizationTotalPeriod_DAOB5 = "0"
End Function
Function DefaultFullRateDeprAmt_DAOB5() As String
DefaultFullRateDeprAmt_DAOB5 = "0"
End Function
Function DefaultAdditionsLessThan180Days_DAOB5() As String
DefaultAdditionsLessThan180Days_DAOB5 = "0"
End Function
Function DefaultRealizationPeriodDuringYear_DAOB5() As String
DefaultRealizationPeriodDuringYear_DAOB5 = "0"
End Function
Function DefaultHalfRateDeprAmt_DAOB5() As String
DefaultHalfRateDeprAmt_DAOB5 = "0"
End Function
Function DefaultDepreciationAtFullRate_DAOB5() As String
DefaultDepreciationAtFullRate_DAOB5 = "0"
End Function
Function DefaultDepreciationAtHalfRate_DAOB5() As String
DefaultDepreciationAtHalfRate_DAOB5 = "0"
End Function
Function DefaultAddlnDeprOnGT180DayAdditions_DAOB5() As String
DefaultAddlnDeprOnGT180DayAdditions_DAOB5 = "0"
End Function
Function DefaultAddlnDeprDuringYearAdditions_DAOB5() As String
DefaultAddlnDeprDuringYearAdditions_DAOB5 = "0"
End Function
Function DefaultTotalDepreciation_DAOB5() As String
DefaultTotalDepreciation_DAOB5 = "0"
End Function
Function DefaultExpdrOnTrforSaleAsset_DAOB5() As String
DefaultExpdrOnTrforSaleAsset_DAOB5 = "0"
End Function
Function DefaultCapGainUs50_DAOB5() As String
DefaultCapGainUs50_DAOB5 = "0"
End Function
Function DefaultWDVLastDay_DAOB5() As String
DefaultWDVLastDay_DAOB5 = "0"
End Function
Function DefaultWDVFirstDay_DAOB10() As String
DefaultWDVFirstDay_DAOB10 = "0"
End Function
Function DefaultAdditionsGrThan180Days_DAOB10() As String
DefaultAdditionsGrThan180Days_DAOB10 = "0"
End Function
Function DefaultRealizationTotalPeriod_DAOB10() As String
DefaultRealizationTotalPeriod_DAOB10 = "0"
End Function
Function DefaultFullRateDeprAmt_DAOB10() As String
DefaultFullRateDeprAmt_DAOB10 = "0"
End Function
Function DefaultAdditionsLessThan180Days_DAOB10() As String
DefaultAdditionsLessThan180Days_DAOB10 = "0"
End Function
Function DefaultRealizationPeriodDuringYear_DAOB10() As String
DefaultRealizationPeriodDuringYear_DAOB10 = "0"
End Function
Function DefaultHalfRateDeprAmt_DAOB10() As String
DefaultHalfRateDeprAmt_DAOB10 = "0"
End Function
Function DefaultDepreciationAtFullRate_DAOB10() As String
DefaultDepreciationAtFullRate_DAOB10 = "0"
End Function
Function DefaultDepreciationAtHalfRate_DAOB10() As String
DefaultDepreciationAtHalfRate_DAOB10 = "0"
End Function
Function DefaultAddlnDeprOnGT180DayAdditions_DAOB10() As String
DefaultAddlnDeprOnGT180DayAdditions_DAOB10 = "0"
End Function
Function DefaultAddlnDeprDuringYearAdditions_DAOB10() As String
DefaultAddlnDeprDuringYearAdditions_DAOB10 = "0"
End Function
Function DefaultTotalDepreciation_DAOB10() As String
DefaultTotalDepreciation_DAOB10 = "0"
End Function
Function DefaultExpdrOnTrforSaleAsset_DAOB10() As String
DefaultExpdrOnTrforSaleAsset_DAOB10 = "0"
End Function
Function DefaultCapGainUs50_DAOB10() As String
DefaultCapGainUs50_DAOB10 = "0"
End Function
Function DefaultWDVLastDay_DAOB10() As String
DefaultWDVLastDay_DAOB10 = "0"
End Function
Function DefaultWDVFirstDay_DAOB100() As String
DefaultWDVFirstDay_DAOB100 = "0"
End Function
Function DefaultAdditionsGrThan180Days_DAOB100() As String
DefaultAdditionsGrThan180Days_DAOB100 = "0"
End Function
Function DefaultRealizationTotalPeriod_DAOB100() As String
DefaultRealizationTotalPeriod_DAOB100 = "0"
End Function
Function DefaultFullRateDeprAmt_DAOB100() As String
DefaultFullRateDeprAmt_DAOB100 = "0"
End Function
Function DefaultAdditionsLessThan180Days_DAOB100() As String
DefaultAdditionsLessThan180Days_DAOB100 = "0"
End Function
Function DefaultRealizationPeriodDuringYear_DAOB100() As String
DefaultRealizationPeriodDuringYear_DAOB100 = "0"
End Function
Function DefaultHalfRateDeprAmt_DAOB100() As String
DefaultHalfRateDeprAmt_DAOB100 = "0"
End Function
Function DefaultDepreciationAtFullRate_DAOB100() As String
DefaultDepreciationAtFullRate_DAOB100 = "0"
End Function
Function DefaultDepreciationAtHalfRate_DAOB100() As String
DefaultDepreciationAtHalfRate_DAOB100 = "0"
End Function
Function DefaultAddlnDeprOnGT180DayAdditions_DAOB100() As String
DefaultAddlnDeprOnGT180DayAdditions_DAOB100 = "0"
End Function
Function DefaultAddlnDeprDuringYearAdditions_DAOB100() As String
DefaultAddlnDeprDuringYearAdditions_DAOB100 = "0"
End Function
Function DefaultTotalDepreciation_DAOB100() As String
DefaultTotalDepreciation_DAOB100 = "0"
End Function
Function DefaultExpdrOnTrforSaleAsset_DAOB100() As String
DefaultExpdrOnTrforSaleAsset_DAOB100 = "0"
End Function
Function DefaultCapGainUs50_DAOB100() As String
DefaultCapGainUs50_DAOB100 = "0"
End Function
Function DefaultWDVLastDay_DAOB100() As String
DefaultWDVLastDay_DAOB100 = "0"
End Function
Function DefaultWDVFirstDay_DAOF10() As String
DefaultWDVFirstDay_DAOF10 = "0"
End Function
Function DefaultAdditionsGrThan180Days_DAOF10() As String
DefaultAdditionsGrThan180Days_DAOF10 = "0"
End Function
Function DefaultRealizationTotalPeriod_DAOF10() As String
DefaultRealizationTotalPeriod_DAOF10 = "0"
End Function
Function DefaultFullRateDeprAmt_DAOF10() As String
DefaultFullRateDeprAmt_DAOF10 = "0"
End Function
Function DefaultAdditionsLessThan180Days_DAOF10() As String
DefaultAdditionsLessThan180Days_DAOF10 = "0"
End Function
Function DefaultRealizationPeriodDuringYear_DAOF10() As String
DefaultRealizationPeriodDuringYear_DAOF10 = "0"
End Function
Function DefaultHalfRateDeprAmt_DAOF10() As String
DefaultHalfRateDeprAmt_DAOF10 = "0"
End Function
Function DefaultDepreciationAtFullRate_DAOF10() As String
DefaultDepreciationAtFullRate_DAOF10 = "0"
End Function
Function DefaultDepreciationAtHalfRate_DAOF10() As String
DefaultDepreciationAtHalfRate_DAOF10 = "0"
End Function
Function DefaultAddlnDeprOnGT180DayAdditions_DAOF10() As String
DefaultAddlnDeprOnGT180DayAdditions_DAOF10 = "0"
End Function
Function DefaultAddlnDeprDuringYearAdditions_DAOF10() As String
DefaultAddlnDeprDuringYearAdditions_DAOF10 = "0"
End Function
Function DefaultTotalDepreciation_DAOF10() As String
DefaultTotalDepreciation_DAOF10 = "0"
End Function
Function DefaultExpdrOnTrforSaleAsset_DAOF10() As String
DefaultExpdrOnTrforSaleAsset_DAOF10 = "0"
End Function
Function DefaultCapGainUs50_DAOF10() As String
DefaultCapGainUs50_DAOF10 = "0"
End Function
Function DefaultWDVLastDay_DAOF10() As String
DefaultWDVLastDay_DAOF10 = "0"
End Function
Function DefaultWDVFirstDay_DAOI25() As String
DefaultWDVFirstDay_DAOI25 = "0"
End Function
Function DefaultAdditionsGrThan180Days_DAOI25() As String
DefaultAdditionsGrThan180Days_DAOI25 = "0"
End Function
Function DefaultRealizationTotalPeriod_DAOI25() As String
DefaultRealizationTotalPeriod_DAOI25 = "0"
End Function
Function DefaultFullRateDeprAmt_DAOI25() As String
DefaultFullRateDeprAmt_DAOI25 = "0"
End Function
Function DefaultAdditionsLessThan180Days_DAOI25() As String
DefaultAdditionsLessThan180Days_DAOI25 = "0"
End Function
Function DefaultRealizationPeriodDuringYear_DAOI25() As String
DefaultRealizationPeriodDuringYear_DAOI25 = "0"
End Function
Function DefaultHalfRateDeprAmt_DAOI25() As String
DefaultHalfRateDeprAmt_DAOI25 = "0"
End Function
Function DefaultDepreciationAtFullRate_DAOI25() As String
DefaultDepreciationAtFullRate_DAOI25 = "0"
End Function
Function DefaultDepreciationAtHalfRate_DAOI25() As String
DefaultDepreciationAtHalfRate_DAOI25 = "0"
End Function
Function DefaultAddlnDeprOnGT180DayAdditions_DAOI25() As String
DefaultAddlnDeprOnGT180DayAdditions_DAOI25 = "0"
End Function
Function DefaultAddlnDeprDuringYearAdditions_DAOI25() As String
DefaultAddlnDeprDuringYearAdditions_DAOI25 = "0"
End Function
Function DefaultTotalDepreciation_DAOI25() As String
DefaultTotalDepreciation_DAOI25 = "0"
End Function
Function DefaultExpdrOnTrforSaleAsset_DAOI25() As String
DefaultExpdrOnTrforSaleAsset_DAOI25 = "0"
End Function
Function DefaultCapGainUs50_DAOI25() As String
DefaultCapGainUs50_DAOI25 = "0"
End Function
Function DefaultWDVLastDay_DAOI25() As String
DefaultWDVLastDay_DAOI25 = "0"
End Function
Function DefaultWDVFirstDay_DAOS20() As String
DefaultWDVFirstDay_DAOS20 = "0"
End Function
Function DefaultAdditionsGrThan180Days_DAOS20() As String
DefaultAdditionsGrThan180Days_DAOS20 = "0"
End Function
Function DefaultRealizationTotalPeriod_DAOS20() As String
DefaultRealizationTotalPeriod_DAOS20 = "0"
End Function
Function DefaultFullRateDeprAmt_DAOS20() As String
DefaultFullRateDeprAmt_DAOS20 = "0"
End Function
Function DefaultAdditionsLessThan180Days_DAOS20() As String
DefaultAdditionsLessThan180Days_DAOS20 = "0"
End Function
Function DefaultRealizationPeriodDuringYear_DAOS20() As String
DefaultRealizationPeriodDuringYear_DAOS20 = "0"
End Function
Function DefaultHalfRateDeprAmt_DAOS20() As String
DefaultHalfRateDeprAmt_DAOS20 = "0"
End Function
Function DefaultDepreciationAtFullRate_DAOS20() As String
DefaultDepreciationAtFullRate_DAOS20 = "0"
End Function
Function DefaultDepreciationAtHalfRate_DAOS20() As String
DefaultDepreciationAtHalfRate_DAOS20 = "0"
End Function
Function DefaultAddlnDeprOnGT180DayAdditions_DAOS20() As String
DefaultAddlnDeprOnGT180DayAdditions_DAOS20 = "0"
End Function
Function DefaultAddlnDeprDuringYearAdditions_DAOS20() As String
DefaultAddlnDeprDuringYearAdditions_DAOS20 = "0"
End Function
Function DefaultTotalDepreciation_DAOS20() As String
DefaultTotalDepreciation_DAOS20 = "0"
End Function
Function DefaultExpdrOnTrforSaleAsset_DAOS20() As String
DefaultExpdrOnTrforSaleAsset_DAOS20 = "0"
End Function
Function DefaultCapGainUs50_DAOS20() As String
DefaultCapGainUs50_DAOS20 = "0"
End Function
Function DefaultWDVLastDay_DAOS20() As String
DefaultWDVLastDay_DAOS20 = "0"
End Function

