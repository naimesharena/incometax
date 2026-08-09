Attribute VB_Name = "CYLACalculations"
Option Explicit
Dim sPassword As String

Dim adjustSpecu As Variant
Dim adjustSpeci As Variant
Dim IncOfCurYrUndHeadFromCYLA2_16_1 As Variant
Dim BusLossOthThanSpecLossCF8_16_1 As Variant
Dim IncOfCurYrUndHeadFromCYLA2a_16_1 As Variant
Dim LossFrmSpecBusCF8_16_1 As Variant
Dim IncOfCurYrUndHeadFromCYLA2b_16_1 As Variant
Dim LossFrmSpecifiedBusCF8_16_1 As Variant
Dim BFlossPrevYrUndSameHeadSetoff2_16_1 As Variant
Dim BFlossPrevYrUndSameHeadSetoff2a_16_1 As Variant
Dim BFlossPrevYrUndSameHeadSetoff2b_16_1 As Variant
Dim BpPendingLossAmt As Variant


Public LossOfUnderThatHead_HP As Variant
Public LossOfUnderThatHead_BP As Variant
Public LossOfUnderThatHead_OS As Variant


Public salary_IncOfCurYrUndHeadFromCYLA As Variant
Public salary_IncOfCurYrAfterSetOffBFLosses As Variant


Public salary_IncOfCurYrUnderThatHead1 As Variant
Public salary_HPlossCurYrSetoff1 As Variant
Public salary_OthSrcLossNoRaceHorseSetoff1 As Variant
Public salary_IncOfCurYrAfterSetOff1 As Variant

Public IncOfCurYrUnderThatHead1_15 As Variant
Public HPlossCurYrSetoff1_15 As Variant
Public OthSrcLossNoRaceHorseSetoff1_15 As Variant
Public IncOfCurYrAfterSetOff1_15 As Variant
Public IncOfCurYrUnderThatHead2_15 As Variant
Public BusLossSetoff2_15 As Variant
Public OthSrcLossNoRaceHorseSetoff2_15 As Variant
Public IncOfCurYrAfterSetOff2_15 As Variant
Public IncOfCurYrUnderThatHead_16 As Variant

Public IncOfCurYrUnderThatHead0a_16 As Variant
Public IncOfCurYrUnderThatHead0b_16 As Variant
Public IncOfCurYrUnderThatHead4_16 As Variant

Public HPlossCurYrSetoff0a_16 As Variant
Public BusLossSetoff0a_16 As Variant
Public OthSrcLossNoRaceHorseSetoff0a_16 As Variant
Public IncOfCurYrAfterSetOff0a_16 As Variant
Public HPlossCurYrSetoff0b_16 As Variant
Public BusLossSetoff0b_16 As Variant
Public OthSrcLossNoRaceHorseSetoff0b_16 As Variant
Public IncOfCurYrAfterSetOff0b_16 As Variant
Public HPlossCurYrSetoff4_16 As Variant
Public BusLossSetoff4_16 As Variant
Public OthSrcLossNoRaceHorseSetoff4_16 As Variant
Public IncOfCurYrAfterSetOff4_16 As Variant


Public IncOfCurYrUndHeadFromCYLA2a_16 As Variant
Public BFlossPrevYrUndSameHeadSetoff2a_16 As Variant
Public BFUnabsorbedDeprSetoff2a_16 As Variant
Public BFAllUs35Cl4Setoff2a_16 As Variant
Public IncOfCurYrAfterSetOffBFLosses2a_16 As Variant
Public IncOfCurYrUndHeadFromCYLA2b_16 As Variant
Public BFlossPrevYrUndSameHeadSetoff2b_16 As Variant
Public BFUnabsorbedDeprSetoff2b_16 As Variant
Public BFAllUs35Cl4Setoff2b_16 As Variant
Public IncOfCurYrAfterSetOffBFLosses2b_16 As Variant
Public IncOfCurYrUndHeadFromCYLA6_16 As Variant
Public BFlossPrevYrUndSameHeadSetoff6_16 As Variant
Public BFUnabsorbedDeprSetoff6_16 As Variant
Public BFAllUs35Cl4Setoff6_16 As Variant
Public IncOfCurYrAfterSetOffBFLosses6_16 As Variant


Public HPlossCurYrSetoff_16 As Variant
Public BusLossSetoff_16 As Variant
Public OthSrcLossNoRaceHorseSetoff_16 As Variant
Public IncOfCurYrAfterSetOff_16 As Variant
Public IncOfCurYrUnderThatHead1_16 As Variant
Public IncOfCurYrUnderThatHead1a_16 As Variant
Public IncOfCurYrUnderThatHead1b_16 As Variant

Public HPlossCurYrSetoff1_16 As Variant
Public BusLossSetoff1_16 As Variant
Public OthSrcLossNoRaceHorseSetoff1_16 As Variant
Public IncOfCurYrAfterSetOff1_16 As Variant

Public HPlossCurYrSetoff1a_16 As Variant
Public BusLossSetoff1a_16 As Variant
Public OthSrcLossNoRaceHorseSetoff1a_16 As Variant
Public IncOfCurYrAfterSetOff1a_16 As Variant

Public HPlossCurYrSetoff1b_16 As Variant
Public BusLossSetoff1b_16 As Variant
Public OthSrcLossNoRaceHorseSetoff1b_16 As Variant
Public IncOfCurYrAfterSetOff1b_16 As Variant


Public IncOfCurYrUnderThatHead2_16 As Variant
Public HPlossCurYrSetoff2_16 As Variant
Public BusLossSetoff2_16 As Variant
Public OthSrcLossNoRaceHorseSetoff2_16 As Variant
Public IncOfCurYrAfterSetOff2_16 As Variant

Public IncOfCurYrUnderThatHead2a_16 As Variant
Public HPlossCurYrSetoff2a_16 As Variant
Public BusLossSetoff2a_16 As Variant
Public OthSrcLossNoRaceHorseSetoff2a_16 As Variant
Public IncOfCurYrAfterSetOff2a_16 As Variant

Public IncOfCurYrUnderThatHead3_16 As Variant
Public HPlossCurYrSetoff3_16 As Variant
Public BusLossSetoff3_16 As Variant
Public OthSrcLossNoRaceHorseSetoff3_16 As Variant
Public IncOfCurYrAfterSetOff3_16 As Variant
Public TotHPlossCurYrSetoff_16 As Variant
Public TotBusLossSetoff_16 As Variant
Public TotOthSrcLossNoRaceHorseSetoff_16 As Variant
Public BalHPlossCurYrAftSetoff_16 As Variant
Public BalBusLossAftSetoff_16 As Variant
Public BalOthSrcLossNoRaceHorseAftSetoff_16 As Variant
Public IncOfCurYrUndHeadFromCYLA_16 As Variant
Public BFlossPrevYrUndSameHeadSetoff_16 As Variant
Public BFUnabsorbedDeprSetoff_16 As Variant
Public BFAllUs35Cl4Setoff_16 As Variant
Public IncOfCurYrAfterSetOffBFLosses_16 As Variant
Public IncOfCurYrUndHeadFromCYLA1_16 As Variant
Public BFlossPrevYrUndSameHeadSetoff1_16 As Variant
Public BFUnabsorbedDeprSetoff1_16 As Variant
Public BFAllUs35Cl4Setoff1_16 As Variant
Public IncOfCurYrAfterSetOffBFLosses1_16 As Variant
Public IncOfCurYrUndHeadFromCYLA2_16 As Variant
Public BFlossPrevYrUndSameHeadSetoff2_16 As Variant
Public BFUnabsorbedDeprSetoff2_16 As Variant
Public BFAllUs35Cl4Setoff2_16 As Variant
Public IncOfCurYrAfterSetOffBFLosses2_16 As Variant

Public IncOfCurYrUndHeadFromCYLA3_16 As Variant
Public BFlossPrevYrUndSameHeadSetoff3_16 As Variant
Public BFUnabsorbedDeprSetoff3_16 As Variant
Public BFAllUs35Cl4Setoff3_16 As Variant
Public IncOfCurYrAfterSetOffBFLosses3_16 As Variant

Public IncOfCurYrUndHeadFromCYLA3a_16 As Variant
Public BFlossPrevYrUndSameHeadSetoff3a_16 As Variant
Public BFUnabsorbedDeprSetoff3a_16 As Variant
Public BFAllUs35Cl4Setoff3a_16 As Variant
Public IncOfCurYrAfterSetOffBFLosses3a_16 As Variant

Public IncOfCurYrUndHeadFromCYLA3b_16 As Variant
Public BFlossPrevYrUndSameHeadSetoff3b_16 As Variant
Public BFUnabsorbedDeprSetoff3b_16 As Variant
Public BFAllUs35Cl4Setoff3b_16 As Variant
Public IncOfCurYrAfterSetOffBFLosses3b_16 As Variant


Public IncOfCurYrUndHeadFromCYLA4_16 As Variant
Public BFlossPrevYrUndSameHeadSetoff4_16 As Variant
Public BFUnabsorbedDeprSetoff4_16 As Variant
Public BFAllUs35Cl4Setoff4_16 As Variant
Public IncOfCurYrAfterSetOffBFLosses4_16 As Variant

Public IncOfCurYrUndHeadFromCYLA4a_16 As Variant
Public BFlossPrevYrUndSameHeadSetoff4a_16 As Variant
Public BFUnabsorbedDeprSetoff4a_16 As Variant
Public BFAllUs35Cl4Setoff4a_16 As Variant
Public IncOfCurYrAfterSetOffBFLosses4a_16 As Variant

Public IncOfCurYrUndHeadFromCYLA5_16 As Variant
Public BFlossPrevYrUndSameHeadSetoff5_16 As Variant
Public BFUnabsorbedDeprSetoff5_16 As Variant
Public BFAllUs35Cl4Setoff5_16 As Variant
Public IncOfCurYrAfterSetOffBFLosses5_16 As Variant
Public TotBFLossSetoff_16 As Variant
Public TotUnabsorbedDeprSetoff_16 As Variant
Public TotAllUs35cl4Setoff_16 As Variant
Public IncomeOfCurrYrAftCYLABFLA_16 As Variant
Dim msgValidateSheet16 As String
Dim msgValidateSheet15 As String

Sub ValidateCFLABFLAClick()
ValidateCFLABFLA
fmsgboxStatus "Sheet CYLA-BFLA is OK" ', vbOKOnly, "ITR-3"
End Sub

Sub ValidateCFLABFLA()
'CYLACalculations.CYLABFLASetOff ' Newly added by Bindu on 7th March 25 as it's creating issue in Validating

'Updated by Bindu
If (UCase(Sheet16.Range("sheet16.CYLA_Edit").value) = UCase("No") And UCase(Sheet16.Range("sheet16.BFLA_Edit").value) = UCase("No")) Then
    CYLACalculations.CYLABFLASetOff
ElseIf (UCase(Sheet16.Range("sheet16.CYLA_Edit").value) = UCase("Yes") Or UCase(Sheet16.Range("sheet16.BFLA_Edit").value) = UCase("Yes")) Then
    CYLA_BFLA_Editable
End If
'--


Sheet16.Activate ' Added by Shrutika(26/05/2025)
If Not Validatesheet16 Then
Sheet16.Activate
If msgValidateSheet16 <> "" Then fmsgboxsmall msgValidateSheet16 ', vbOKOnly, "Error(s)!"
If msgValidateSheet15 <> "" Then fmsgboxsmall msgValidateSheet15 ', vbOKOnly, "Error(s)!"
CloseMsg
End If
End Sub

Function Validatesheet16() As Boolean
    Validatesheet16 = True

    If Not validatesalary() Then Validatesheet16 = False
    If Not ValidateLossOfCurYrUnder() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrUnderThatHead2_15() Then Validatesheet16 = False
    If Not ValidateBusLossSetoff2_15() Then Validatesheet16 = False
    If Not ValidateOthSrcLossNoRaceHorseSetoff2_15() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrAfterSetOff2_15() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrUnderThatHead_16() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrUnderThatHead0a_16() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrUnderThatHead0b_16() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrUnderThatHead4_16() Then Validatesheet16 = False
    If Not ValidateHPlossCurYrSetoff0a_16() Then Validatesheet16 = False
    If Not ValidateOthSrcLossNoRaceHorseSetoff0a_16() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrAfterSetOff0a_16() Then Validatesheet16 = False
    If Not ValidateHPlossCurYrSetoff0b_16() Then Validatesheet16 = False
    If Not ValidateOthSrcLossNoRaceHorseSetoff0b_16() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrAfterSetOff0b_16() Then Validatesheet16 = False
    If Not ValidateHPlossCurYrSetoff4_16() Then Validatesheet16 = False
    If Not ValidateBusLossSetoff4_16() Then Validatesheet16 = False
    If Not ValidateOthSrcLossNoRaceHorseSetoff4_16() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrAfterSetOff4_16() Then Validatesheet16 = False

    If Not ValidateHPlossCurYrSetoff_16() Then Validatesheet16 = False
        'chandru
    If Not ValidateHPlossSetoff_Upto2Lakhs() Then Validatesheet16 = False
    If Not ValidateBroughtForwardLossSetoff_BFLA_iv() Then Validatesheet16 = False
    If Not ValidateBroughtForwardLossSetoff_BFLA_v() Then Validatesheet16 = False
    If Not ValidateBroughtForwardLossSetoff_BFLA_viii() Then Validatesheet16 = False
    If Not ValidateBroughtForwardLossSetoff_BFLA_iii() Then Validatesheet16 = False
    If Not ValidateBroughtForwardLossSetoff_BFLA_vii() Then Validatesheet16 = False
    If Not ValidateBroughtForwardLossSetoff_BFLA_vi() Then Validatesheet16 = False
    If Not ValidateBroughtForwardLossSetoff_BFLA_ix() Then Validatesheet16 = False
    If Not ValidateBFLA_LongtermShortterm_Limit_CFL() Then Validatesheet16 = False
    If Not ValidateEntireLossSetOff_CYLA() Then Validatesheet16 = False
    If Not ValidateBFLA_LTCG_Limit() Then Validatesheet16 = False
    If Not ValidateBFLA_STCG_Limit() Then Validatesheet16 = False
    If Not ValidateTotalLoss_CYLA_New_Approach() Then Validatesheet16 = False
   ' If Not ValidateBlankData_BFLA() Then Validatesheet16 = False
    

    If Not ValidateBFLA_LTCG_LessthanCFL_LTCG() Then Validatesheet16 = False
    If Not ValidateBFLA_STCG_LessthanCFL_STCG() Then Validatesheet16 = False
    'If Not ValidateBFLA_LTCG_MorethanCFL_LTCG() Then Validatesheet16 = False
    'If Not ValidateBFLA_STCG_MorethanCFL_STCG() Then Validatesheet16 = False
    'If Not ValidateBFLA_STCG_Limit_CFL() Then Validatesheet16 = False
    
    'If Not ValidateBusLossSetoff_16() Then Validatesheet16 = False
    If Not ValidateOthSrcLossNoRaceHorseSetoff_16() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrAfterSetOff_16() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrUnderThatHead1_16() Then Validatesheet16 = False
    If Not ValidateHPlossCurYrSetoff1_16() Then Validatesheet16 = False
    If Not ValidateBusLossSetoff1_16() Then Validatesheet16 = False
    If Not ValidateOthSrcLossNoRaceHorseSetoff1_16() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrAfterSetOff1_16() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrUnderThatHead2_16() Then Validatesheet16 = False
    If Not ValidateHPlossCurYrSetoff2_16() Then Validatesheet16 = False
    If Not ValidateBusLossSetoff2_16() Then Validatesheet16 = False
    If Not ValidateOthSrcLossNoRaceHorseSetoff2_16() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrAfterSetOff2_16() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrUnderThatHead3_16() Then Validatesheet16 = False
    If Not ValidateHPlossCurYrSetoff3_16() Then Validatesheet16 = False
    If Not ValidateBusLossSetoff3_16() Then Validatesheet16 = False
   ' If Not ValidateOthSrcLossNoRaceHorseSetoff3_16() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrAfterSetOff3_16() Then Validatesheet16 = False
    If Not ValidateTotHPlossCurYrSetoff_16() Then Validatesheet16 = False
    If Not ValidateTotBusLossSetoff_16() Then Validatesheet16 = False
    If Not ValidateTotOthSrcLossNoRaceHorseSetoff_16() Then Validatesheet16 = False
    If Not ValidateBalHPlossCurYrAftSetoff_16() Then Validatesheet16 = False
    If Not ValidateBalBusLossAftSetoff_16() Then Validatesheet16 = False
    If Not ValidateBalOthSrcLossNoRaceHorseAftSetoff_16() Then Validatesheet16 = False
 
    If Not validateSalBFLA() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrUndHeadFromCYLA1_16() Then Validatesheet16 = False
    If Not ValidateBFlossPrevYrUndSameHeadSetoff1_16() Then Validatesheet16 = False
    If Not ValidateBFUnabsorbedDeprSetoff1_16() Then Validatesheet16 = False
    If Not ValidateBFAllUs35Cl4Setoff1_16() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrAfterSetOffBFLosses1_16() Then Validatesheet16 = False
    
    If Not ValidateIncOfCurYrUndHeadFromCYLA2_16() Then Validatesheet16 = False
    If Not ValidateBFlossPrevYrUndSameHeadSetoff2_16() Then Validatesheet16 = False
    If Not ValidateBFUnabsorbedDeprSetoff2_16() Then Validatesheet16 = False
    If Not ValidateBFAllUs35Cl4Setoff2_16() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrAfterSetOffBFLosses2_16() Then Validatesheet16 = False
    
    If Not ValidateIncOfCurYrUndHeadFromCYLA3_16() Then Validatesheet16 = False
    If Not ValidateBFlossPrevYrUndSameHeadSetoff3_16() Then Validatesheet16 = False
    If Not ValidateBFUnabsorbedDeprSetoff3_16() Then Validatesheet16 = False
    If Not ValidateBFAllUs35Cl4Setoff3_16() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrAfterSetOffBFLosses3_16() Then Validatesheet16 = False
    
    If Not ValidateIncOfCurYrUndHeadFromCYLA4_16() Then Validatesheet16 = False
    If Not ValidateBFlossPrevYrUndSameHeadSetoff4_16() Then Validatesheet16 = False
    If Not ValidateBFUnabsorbedDeprSetoff4_16() Then Validatesheet16 = False
    If Not ValidateBFAllUs35Cl4Setoff4_16() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrAfterSetOffBFLosses4_16() Then Validatesheet16 = False
    
    If Not ValidateIncOfCurYrUndHeadFromCYLA5_16() Then Validatesheet16 = False
    If Not ValidateBFlossPrevYrUndSameHeadSetoff5_16() Then Validatesheet16 = False
    If Not ValidateBFUnabsorbedDeprSetoff5_16() Then Validatesheet16 = False
    If Not ValidateBFAllUs35Cl4Setoff5_16() Then Validatesheet16 = False
    If Not ValidateIncOfCurYrAfterSetOffBFLosses5_16() Then Validatesheet16 = False

    If Not ValidateTotBFLossSetoff_16() Then Validatesheet16 = False
    If Not ValidateTotUnabsorbedDeprSetoff_16() Then Validatesheet16 = False
    If Not ValidateTotAllUs35cl4Setoff_16() Then Validatesheet16 = False
    If Not ValidateIncomeOfCurrYrAftCYLABFLA_16() Then Validatesheet16 = False
    
'Chandru
'  If Sheet16.Range("sheet16.BalHPlossCurYrAftSetoff").value > 0 Or Sheet16.Range("sheet16.BalOthSrcLossNoRaceHorseAftSetoff").value > 0 And Sheet16.Range("Total_Column4").value > 0 Then
'fmsgboxStatus """entire loss is not set off """
'End If
    
End Function


Function ValidateIncOfCurYrUnderThatHead2_15() As Boolean
    Dim i As Long
    
    ValidateIncOfCurYrUnderThatHead2_15 = True
    IncOfCurYrUnderThatHead2_15 = Sheet16.Range("hp.IncOfCurYrUnderThatHead2").value
    
    If Len(IncOfCurYrUnderThatHead2_15) > 14 Then
        msgbox15 ("IncOfCurYrUnderThatHead2 in Sheet : CYLA - BFLA  cannot be greater than 14 digits ")
        ValidateIncOfCurYrUnderThatHead2_15 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrUnderThatHead2_15)
        If Not IsNumeric(IncOfCurYrUnderThatHead2_15) Then
            msgbox15 ("IncOfCurYrUnderThatHead2 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrUnderThatHead2_15 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBusLossSetoff2_15() As Boolean
    Dim i As Long
    
    ValidateBusLossSetoff2_15 = True
    BusLossSetoff2_15 = Sheet16.Range("hp.BusLossSetoff2").value
    
    If Len(BusLossSetoff2_15) > 14 Then
        msgbox15 ("BusLossSetoff2 in Sheet : CYLA - BFLA  cannot be greater than 14 digits ")
        ValidateBusLossSetoff2_15 = False
        Exit Function
    End If
    
    For i = 1 To Len(BusLossSetoff2_15)
        If Not IsNumeric(BusLossSetoff2_15) Then
            msgbox15 ("BusLossSetoff2 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBusLossSetoff2_15 = False
            Exit Function
        End If
    Next
End Function

Function ValidateOthSrcLossNoRaceHorseSetoff2_15() As Boolean
    Dim i As Long

    ValidateOthSrcLossNoRaceHorseSetoff2_15 = True
    OthSrcLossNoRaceHorseSetoff2_15 = Sheet16.Range("hp.OthSrcLossNoRaceHorseSetoff2").value
    
    If Len(OthSrcLossNoRaceHorseSetoff2_15) > 14 Then
        msgbox15 ("OthSrcLossNoRaceHorseSetoff2 in Sheet : CYLA - BFLA  cannot be greater than 14 digits ")
        ValidateOthSrcLossNoRaceHorseSetoff2_15 = False
        Exit Function
    End If
    
    For i = 1 To Len(OthSrcLossNoRaceHorseSetoff2_15)
        If Not IsNumeric(OthSrcLossNoRaceHorseSetoff2_15) Then
            msgbox15 ("OthSrcLossNoRaceHorseSetoff2 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateOthSrcLossNoRaceHorseSetoff2_15 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncOfCurYrAfterSetOff2_15() As Boolean
    Dim i As Long

    ValidateIncOfCurYrAfterSetOff2_15 = True
    IncOfCurYrAfterSetOff2_15 = Sheet16.Range("hp.IncOfCurYrAfterSetOff2").value
    
    If Len(IncOfCurYrAfterSetOff2_15) > 14 Then
        msgbox15 ("IncOfCurYrAfterSetOff2 in Sheet : CYLA - BFLA  cannot be greater than 14 digits ")
        ValidateIncOfCurYrAfterSetOff2_15 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrAfterSetOff2_15)
        If Not IsNumeric(IncOfCurYrAfterSetOff2_15) Then
            msgbox15 ("IncOfCurYrAfterSetOff2 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrAfterSetOff2_15 = False
            Exit Function
        End If
    Next
End Function

Function DefaultIncOfCurYrUnderThatHead1_15() As String
    IncOfCurYrUnderThatHead1_15 = "0"
End Function

Function DefaultHPlossCurYrSetoff1_15() As String
    HPlossCurYrSetoff1_15 = "0"
End Function

Function DefaultOthSrcLossNoRaceHorseSetoff1_15() As String
    OthSrcLossNoRaceHorseSetoff1_15 = "0"
End Function

Function DefaultIncOfCurYrAfterSetOff1_15() As String
    IncOfCurYrAfterSetOff1_15 = "0"
End Function

Function DefaultIncOfCurYrUnderThatHead2_15() As String
    IncOfCurYrUnderThatHead2_15 = "0"
End Function

Function DefaultBusLossSetoff2_15() As String
    BusLossSetoff2_15 = "0"
End Function

Function DefaultOthSrcLossNoRaceHorseSetoff2_15() As String
    OthSrcLossNoRaceHorseSetoff2_15 = "0"
End Function

Function DefaultIncOfCurYrAfterSetOff2_15() As String
    IncOfCurYrAfterSetOff2_15 = "0"
End Function

Function ValidateIncOfCurYrUnderThatHead_16() As Boolean
    Dim i As Long
    
    ValidateIncOfCurYrUnderThatHead_16 = True
    IncOfCurYrUnderThatHead_16 = Sheet16.Range("busprof.IncOfCurYrUnderThatHead").value
    
    If Len(IncOfCurYrUnderThatHead_16) > 14 Then
        msgbox16 ("IncOfCurYrUnderThatHead in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncOfCurYrUnderThatHead_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrUnderThatHead_16)
        If Not IsNumeric(IncOfCurYrUnderThatHead_16) Then
            msgbox16 ("IncOfCurYrUnderThatHead in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrUnderThatHead_16 = False
            Exit Function
        End If
    Next
End Function

'''''
Function ValidateIncOfCurYrUnderThatHead0a_16() As Boolean
    ValidateIncOfCurYrUnderThatHead0a_16 = True

    If (Sheet16.Range("busprofspec.IncOfCurYrUnderThatHead0a").value <> "") Then
        IncOfCurYrUnderThatHead0a_16 = Sheet16.Range("busprofspec.IncOfCurYrUnderThatHead0a").value
    Else
        IncOfCurYrUnderThatHead0a_16 = 0
    End If
End Function

Function ValidateIncOfCurYrUnderThatHead0b_16() As Boolean
    ValidateIncOfCurYrUnderThatHead0b_16 = True

    If (Sheet16.Range("busprofspecified.IncOfCurYrUnderThatHead0b").value <> "") Then
        IncOfCurYrUnderThatHead0b_16 = Sheet16.Range("busprofspecified.IncOfCurYrUnderThatHead0b").value
    Else
        IncOfCurYrUnderThatHead0b_16 = 0
    End If
End Function

Function ValidateIncOfCurYrUnderThatHead4_16() As Boolean
    ValidateIncOfCurYrUnderThatHead4_16 = True
    
    If (Sheet16.Range("rh.IncOfCurYrUnderThatHead4").value <> "") Then
        IncOfCurYrUnderThatHead4_16 = Sheet16.Range("rh.IncOfCurYrUnderThatHead4").value
    Else
        IncOfCurYrUnderThatHead4_16 = 0
    End If
End Function

Function ValidateHPlossCurYrSetoff0a_16() As Boolean
    ValidateHPlossCurYrSetoff0a_16 = True
    
    If (Sheet16.Range("busprofspec.HPlossCurYrSetoff0a").value <> "") Then
        HPlossCurYrSetoff0a_16 = Sheet16.Range("busprofspec.HPlossCurYrSetoff0a").value
    Else
        HPlossCurYrSetoff0a_16 = 0
    End If
End Function

Function ValidateOthSrcLossNoRaceHorseSetoff0a_16() As Boolean
    ValidateOthSrcLossNoRaceHorseSetoff0a_16 = True
    
    If (Sheet16.Range("busprofspec.OthSrcLossNoRaceHorseSetoff0a").value <> "") Then
        OthSrcLossNoRaceHorseSetoff0a_16 = Sheet16.Range("busprofspec.OthSrcLossNoRaceHorseSetoff0a").value
    Else
        OthSrcLossNoRaceHorseSetoff0a_16 = 0
    End If
End Function

Function ValidateIncOfCurYrAfterSetOff0a_16() As Boolean
    ValidateIncOfCurYrAfterSetOff0a_16 = True
    
    If (Sheet16.Range("busprofspec.IncOfCurYrAfterSetOff0a").value <> "") Then
        IncOfCurYrAfterSetOff0a_16 = Sheet16.Range("busprofspec.IncOfCurYrAfterSetOff0a").value
    Else
        IncOfCurYrAfterSetOff0a_16 = 0
    End If
End Function

Function ValidateHPlossCurYrSetoff0b_16() As Boolean
    ValidateHPlossCurYrSetoff0b_16 = True
    
    If (Sheet16.Range("busprof0b.HPlossCurYrSetoff0b").value <> "") Then
        HPlossCurYrSetoff0b_16 = Sheet16.Range("busprof0b.HPlossCurYrSetoff0b").value
    Else
        HPlossCurYrSetoff0b_16 = 0
    End If
End Function

Function ValidateOthSrcLossNoRaceHorseSetoff0b_16() As Boolean
    ValidateOthSrcLossNoRaceHorseSetoff0b_16 = True
    
    If (Sheet16.Range("busprofspecified.OthSrcLossNoRaceHorseSetoff0b").value <> "") Then
        OthSrcLossNoRaceHorseSetoff0b_16 = Sheet16.Range("busprofspecified.OthSrcLossNoRaceHorseSetoff0b").value
    Else
        OthSrcLossNoRaceHorseSetoff0b_16 = 0
    End If
End Function

Function ValidateIncOfCurYrAfterSetOff0b_16() As Boolean
    ValidateIncOfCurYrAfterSetOff0b_16 = True
    
    If (Sheet16.Range("busprofspecified.IncOfCurYrAfterSetOff0b").value <> "") Then
        IncOfCurYrAfterSetOff0b_16 = Sheet16.Range("busprofspecified.IncOfCurYrAfterSetOff0b").value
    Else
        IncOfCurYrAfterSetOff0b_16 = 0
    End If
End Function

Function ValidateHPlossCurYrSetoff4_16() As Boolean
    ValidateHPlossCurYrSetoff4_16 = True
    
    If (Sheet16.Range("rh.HPlossCurYrSetoff4").value <> "") Then
        HPlossCurYrSetoff4_16 = Sheet16.Range("rh.HPlossCurYrSetoff4").value
    Else
        HPlossCurYrSetoff4_16 = 0
    End If
End Function

Function ValidateBusLossSetoff4_16() As Boolean
    ValidateBusLossSetoff4_16 = True
    
    If (Sheet16.Range("rh.BusLossSetoff4").value <> "") Then
        BusLossSetoff4_16 = Sheet16.Range("rh.BusLossSetoff4").value
    Else
        BusLossSetoff4_16 = 0
    End If
End Function

Function ValidateOthSrcLossNoRaceHorseSetoff4_16() As Boolean
    ValidateOthSrcLossNoRaceHorseSetoff4_16 = True
    
    If (Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4").value <> "") Then
        OthSrcLossNoRaceHorseSetoff4_16 = Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4").value
    Else
        OthSrcLossNoRaceHorseSetoff4_16 = 0
    End If
End Function

Function ValidateIncOfCurYrAfterSetOff4_16() As Boolean
    ValidateIncOfCurYrAfterSetOff4_16 = True
    
    If (Sheet16.Range("rh.IncOfCurYrAfterSetOff4").value <> "") Then
        IncOfCurYrAfterSetOff4_16 = Sheet16.Range("rh.IncOfCurYrAfterSetOff4").value
    Else
        IncOfCurYrAfterSetOff4_16 = 0
    End If
    
    If (Sheet16.Range("busipofinclspec.IncOfCurYrUndHeadFromCYLA2a").value <> "") Then
        IncOfCurYrUndHeadFromCYLA2a_16 = Sheet16.Range("busipofinclspec.IncOfCurYrUndHeadFromCYLA2a").value
    Else
        IncOfCurYrUndHeadFromCYLA2a_16 = 0
    End If
    
    If (Sheet16.Range("busipofinclspec.BFlossPrevYrUndSameHeadSetoff2a").value <> "") Then
        BFlossPrevYrUndSameHeadSetoff2a_16 = Sheet16.Range("busipofinclspec.BFlossPrevYrUndSameHeadSetoff2a").value
    Else
        BFlossPrevYrUndSameHeadSetoff2a_16 = 0
    End If
    
    If (Sheet16.Range("busipofinclspec.BFUnabsorbedDeprSetoff2a").value <> "") Then
        BFUnabsorbedDeprSetoff2a_16 = Sheet16.Range("busipofinclspec.BFUnabsorbedDeprSetoff2a").value
    Else
        BFUnabsorbedDeprSetoff2a_16 = 0
    End If
    
    If (Sheet16.Range("busipofinclspec.BFAllUs35Cl4Setoff2a").value <> "") Then
        BFAllUs35Cl4Setoff2a_16 = Sheet16.Range("busipofinclspec.BFAllUs35Cl4Setoff2a").value
    Else
        BFAllUs35Cl4Setoff2a_16 = 0
    End If
    
    If (Sheet16.Range("busipofinclspec.IncOfCurYrAfterSetOffBFLosses2a").value <> "") Then
        IncOfCurYrAfterSetOffBFLosses2a_16 = Sheet16.Range("busipofinclspec.IncOfCurYrAfterSetOffBFLosses2a").value
    Else
        IncOfCurYrAfterSetOffBFLosses2a_16 = 0
    End If
    
    If (Sheet16.Range("busipofinclspecified.IncOfCurYrUndHeadFromCYLA2b").value <> "") Then
        IncOfCurYrUndHeadFromCYLA2b_16 = Sheet16.Range("busipofinclspecified.IncOfCurYrUndHeadFromCYLA2b").value
    Else
        IncOfCurYrUndHeadFromCYLA2b_16 = 0
    End If
    
    If (Sheet16.Range("busipofinclspecified.BFlossPrevYrUndSameHeadSetoff2b").value <> "") Then
        BFlossPrevYrUndSameHeadSetoff2b_16 = Sheet16.Range("busipofinclspecified.BFlossPrevYrUndSameHeadSetoff2b").value
    Else
        BFlossPrevYrUndSameHeadSetoff2b_16 = 0
    End If
    
    If (Sheet16.Range("busipofinclspecified.BFUnabsorbedDeprSetoff2b").value <> "") Then
        BFUnabsorbedDeprSetoff2b_16 = Sheet16.Range("busipofinclspecified.BFUnabsorbedDeprSetoff2b").value
    Else
        BFUnabsorbedDeprSetoff2b_16 = 0
    End If
    
    If (Sheet16.Range("busipofinclspecified.BFAllUs35Cl4Setoff2b").value <> "") Then
        BFAllUs35Cl4Setoff2b_16 = Sheet16.Range("busipofinclspecified.BFAllUs35Cl4Setoff2b").value
    Else
        BFAllUs35Cl4Setoff2b_16 = 0
    End If
    
    If (Sheet16.Range("busipofinclspecified.IncOfCurYrAfterSetOffBFLosses2b").value <> "") Then
        IncOfCurYrAfterSetOffBFLosses2b_16 = Sheet16.Range("busipofinclspecified.IncOfCurYrAfterSetOffBFLosses2b").value
    Else
        IncOfCurYrAfterSetOffBFLosses2b_16 = 0
    End If
    
    If (Sheet16.Range("rh.IncOfCurYrUndHeadFromCYLA6").value <> "") Then
        IncOfCurYrUndHeadFromCYLA6_16 = Sheet16.Range("rh.IncOfCurYrUndHeadFromCYLA6").value
    Else
        IncOfCurYrUndHeadFromCYLA6_16 = 0
    End If
    
    If (Sheet16.Range("rh.IncOfCurYrUndHeadFromCYLA6").value <> "") Then
        IncOfCurYrUndHeadFromCYLA6_16 = Sheet16.Range("rh.IncOfCurYrUndHeadFromCYLA6").value
    Else
        IncOfCurYrUndHeadFromCYLA6_16 = 0
    End If
    
    If (Sheet16.Range("rh.BFlossPrevYrUndSameHeadSetoff6").value <> "") Then
        BFlossPrevYrUndSameHeadSetoff6_16 = Sheet16.Range("rh.BFlossPrevYrUndSameHeadSetoff6").value
    Else
        BFlossPrevYrUndSameHeadSetoff6_16 = 0
    End If
    
    If (Sheet16.Range("rh.BFUnabsorbedDeprSetoff6").value <> "") Then
        BFUnabsorbedDeprSetoff6_16 = Sheet16.Range("rh.BFUnabsorbedDeprSetoff6").value
    Else
        BFUnabsorbedDeprSetoff6_16 = 0
    End If
    
    If (Sheet16.Range("rh.BFAllUs35Cl4Setoff6").value <> "") Then
        BFAllUs35Cl4Setoff6_16 = Sheet16.Range("rh.BFAllUs35Cl4Setoff6").value
    Else
        BFAllUs35Cl4Setoff6_16 = 0
    End If
    
    If (Sheet16.Range("rh.IncOfCurYrAfterSetOffBFLosses6").value <> "") Then
        IncOfCurYrAfterSetOffBFLosses6_16 = Sheet16.Range("rh.IncOfCurYrAfterSetOffBFLosses6").value
    Else
        IncOfCurYrAfterSetOffBFLosses6_16 = 0
    End If
End Function


'Chandru
Function ValidateHPlossSetoff_Upto2Lakhs() As Boolean
       
    ValidateHPlossSetoff_Upto2Lakhs = True
   ' HPlossCurYrSetoff_Upto2Lakhs = Sheet16.Range("Total_2XIII").value
    
    If Sheet16.Range("Total_2XIII").value > 200000 Then
        msgbox16 ("* Maximum loss that can be set off for house property is Rs 200000.")
        ValidateHPlossSetoff_Upto2Lakhs = False
        Exit Function
    End If
    

End Function

'Chandru
Function ValidateBroughtForwardLossSetoff_BFLA_iv() As Boolean

      
        ValidateBroughtForwardLossSetoff_BFLA_iv = True
    
        If Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3a").value > Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3a").value Then
             Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3a").value = 0
             msgbox16 ("* Value of column no 2 cannot be more than amount as referred in column 1 for Sl no iv.")
        ValidateBroughtForwardLossSetoff_BFLA_iv = False
Exit Function
End If
End Function

'Chandru
Function ValidateBroughtForwardLossSetoff_BFLA_v() As Boolean
        
        ValidateBroughtForwardLossSetoff_BFLA_v = True
        If Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b").value > Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3b").value Then
            Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b").value = 0
            msgbox16 ("* Value of column no 2 cannot be more than amount as referred in column 1 for Sl no v.")

        ValidateBroughtForwardLossSetoff_BFLA_v = False
Exit Function
End If
End Function

'Chandru
Function ValidateBroughtForwardLossSetoff_BFLA_viii() As Boolean
        ValidateBroughtForwardLossSetoff_BFLA_viii = True
        If Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4a").value > Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4a").value Then
            Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4a").value = 0
            msgbox16 ("* Value of column no 2 cannot be more than amount as referred in column 1 for Sl no viii.")
        
        ValidateBroughtForwardLossSetoff_BFLA_viii = False
Exit Function
End If
End Function

'Chandru
Function ValidateBroughtForwardLossSetoff_BFLA_iii() As Boolean
        ValidateBroughtForwardLossSetoff_BFLA_iii = True
        If Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3").value > Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3").value Then
            Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3").value = 0
            msgbox16 ("* Value of column no 2 cannot be more than amount as referred in column 1 for Sl no iii.")
        
         ValidateBroughtForwardLossSetoff_BFLA_iii = False
Exit Function
End If
End Function

'Chandru
Function ValidateBroughtForwardLossSetoff_BFLA_vii() As Boolean

        ValidateBroughtForwardLossSetoff_BFLA_vii = True
        If Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4").value > Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4").value Then
            Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4").value = 0
            msgbox16 ("* Value of column no 2 cannot be more than amount as referred in column 1 for Sl no vii.")
        ValidateBroughtForwardLossSetoff_BFLA_vii = False
        
Exit Function
End If
 End Function

'Chandru
Function ValidateBroughtForwardLossSetoff_BFLA_vi() As Boolean

        ValidateBroughtForwardLossSetoff_BFLA_vi = True
        
        If Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b_DTAARates").value > Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3b_DTAARates").value Then
            Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b_DTAARates").value = 0
            msgbox16 ("* Value of column no 2 cannot be more than amount as referred in column 1 for Sl no vi.")
        ValidateBroughtForwardLossSetoff_BFLA_vi = False
Exit Function
End If
End Function

'Chandru
Function ValidateBroughtForwardLossSetoff_BFLA_ix() As Boolean

        ValidateBroughtForwardLossSetoff_BFLA_ix = True
        
        If Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff_DTAARates").value > Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4a_DTAARates").value Then
            Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff_DTAARates").value = 0
            msgbox16 ("* Value of column no 2 cannot be more than amount as referred in column 1 for Sl no ix.")
        
        ValidateBroughtForwardLossSetoff_BFLA_ix = False
Exit Function
End If
End Function


'Chandru
Function ValidateBFLA_LongtermShortterm_Limit_CFL() As Boolean

        ValidateBFLA_LongtermShortterm_Limit_CFL = True

        If (Sheet16.Range("BF_Total_Loss").value > Sheet17.Range("totofbfloss.STCGLossCF8").value + Sheet17.Range("totofbfloss.LTCGLossCF8").value) Or _
        Sheet16.Range("LTCG_TOTAL_LOSS").value > Sheet17.Range("totofbfloss.STCGLossCF8").value + Sheet17.Range("totofbfloss.LTCGLossCF8").value Or _
        Sheet16.Range("STCG_TOTAL_LOSS").value > Sheet17.Range("totofbfloss.STCGLossCF8").value Then
             msgbox16 ("* ""Losses set off cannot be more than the losses brought forward from previous years in Schedule CFL."" ")

        ValidateBFLA_LongtermShortterm_Limit_CFL = False
Exit Function
End If
End Function


'''''
'Chandru
Function ValidateBFLA_STCG_Limit() As Boolean

        ValidateBFLA_STCG_Limit = True

        If Sheet17.Range("totcfloss.STCGLossCF11").value > 0 Then
         
         If Sheet16.Range("BFLA_Col3_Total").value > 0 Then
        
            msgbox16 ("* ""Maximum STCG losses has not been setoff"" ")

        ValidateBFLA_STCG_Limit = False
Exit Function
End If
End If
End Function

'Chandru
Function ValidateBFLA_LTCG_Limit() As Boolean

        ValidateBFLA_LTCG_Limit = True

        If Sheet17.Range("totcfloss.LTCGLossCF11").value > 0 Then
         
         If Sheet16.Range("LTCG_Col3_Total").value > 0 Then
        
            msgbox16 ("* ""Maximum LTCG losses has not been setoff"" ")

        ValidateBFLA_LTCG_Limit = False
Exit Function
End If
End If
End Function

'Chandru Less than Validation

Function ValidateBFLA_LTCG_LessthanCFL_LTCG() As Boolean

        ValidateBFLA_LTCG_LessthanCFL_LTCG = True
        
        If Sheet16.Range("LTCG_TOTAL_LOSS").value < Sheet16.Range("Longterm_Income").value Then
        
        If Sheet16.Range("LTCG_TOTAL_LOSS").value < Sheet17.Range("CFL_LTCG_Normal_2023").value Then
                     msgbox16 ("* ""Maximum LTCG losses has not been setoff"" ")
                ValidateBFLA_LTCG_LessthanCFL_LTCG = False
        Exit Function

End If
End If
End Function

'Chandru (Set off STCG by removing LTCG in total)

Function ValidateBFLA_STCG_LessthanCFL_STCG() As Boolean

        ValidateBFLA_STCG_LessthanCFL_STCG = True
        
    If Sheet16.Range("STLT_Total").value <> Sheet16.Range("BF_Total_Loss").value Then
    
    If Sheet16.Range("STCG_TOTAL_LOSS").value <> Sheet17.Range("CFL_STCG_Normal_2023").value Then
    If Sheet16.Range("BF_Total_Loss").value > Sheet17.Range("CFL_STCG_Normal_2023").value _
     Or Sheet16.Range("STCG_TOTAL_LOSS").value < Sheet17.Range("CFL_STCG_Normal_2023").value Then
    

        If (Sheet16.Range("BF_Total_Loss").value - Sheet17.Range("CFL_LTCG_Normal_2023").value) < Sheet17.Range("CFL_STCG_Normal_2023").value Then
                     msgbox16 ("* ""Maximum STCG losses has not been setoff"" ")
                ValidateBFLA_STCG_LessthanCFL_STCG = False
        Exit Function
End If

End If
End If
End If
End Function


Function ValidateEntireLossSetOff_CYLA() As Boolean


        ValidateEntireLossSetOff_CYLA = True
        
        If Sheet16.Range("sheet16.TotHPlossCurYrSetoff").value < 200000 Or Sheet16.Range("sheet16.BalOthSrcLossNoRaceHorseAftSetoff").value > 0 Then
        
        
                If Sheet16.Range("sheet16.BalHPlossCurYrAftSetoff").value > 0 Or Sheet16.Range("sheet16.BalOthSrcLossNoRaceHorseAftSetoff").value > 0 Then
                If Sheet16.Range("Total_Column4").value > 0 Then
                msgbox16 """entire loss is not set off """

         ValidateEntireLossSetOff_CYLA = False
        Exit Function
End If
End If
End If

End Function

'Chandru
'Function ValidateBFLA_STCG_Limit_CFL() As Boolean
'
'        ValidateBFLA_STCG_Limit_CFL = True
'
'        If Sheet16.Range("STCG_TOTAL_LOSS").value > Sheet17.Range("totofbfloss.STCGLossCF8").value Then
'             msgbox16 ("* ""11Losses set off cannot be more than the losses brought forward from previous years in Schedule CFL."" ")
'
'        ValidateBFLA_STCG_Limit_CFL = False
'Exit Function
'End If
'End Function

'Function ValidateBFLA_STCG_MorethanCFL_STCG() As Boolean
'
'        ValidateBFLA_STCG_MorethanCFL_STCG = True
'
'        If Sheet16.Range("STCG_TOTAL_LOSS").value > Sheet17.Range("CFL_STCG_Normal_2023").value Then
'                     msgbox16 ("* ""11Losses set off cannot be more than the losses brought forward from previous years in Schedule CFL"" ")
'                ValidateBFLA_STCG_MorethanCFL_STCG = False
'        Exit Function
'
'End If
'End Function

'Chandru


'Chandru
'Total LTCL setoff

'Function ValidateBFLA_LTCG_MorethanCFL_LTCG() As Boolean
'
'        ValidateBFLA_LTCG_MorethanCFL_LTCG = True
'
'        If Sheet16.Range("LTCG_TOTAL_LOSS").value > Sheet17.Range("CFL_LTCG_Normal_2023").value Then
'                     msgbox16 ("* ""11Losses set off cannot be more than the losses brought forward from previous years in Schedule CFL"" ")
'                ValidateBFLA_LTCG_MorethanCFL_LTCG = False
'        Exit Function
'
'End If
'End Function



'Function ValidateBFLA_Maximum_losses_Notset_Total() As Boolean
'
'        ValidateBFLA_Maximum_losses_Notset_Total = True
'If Sheet16.Range("LTCG_TOTAL_LOSS").value > Sheet16.Range("CFL_Longterm").value Then
'        If Sheet16.Range("BF_Total_Loss").value < Sheet16.Range("STLT_Total").value And Sheet16.Range("BF_Total_Loss").value < Sheet16.Range("CFL_STLT_Total").value Then
'               ' If Sheet16.Range("STCG_TOTAL_LOSS").value < Sheet16.Range("CFL_Shortterm").value Then
'                     msgbox16 ("* ""Maximum losses has not been setoff"" ")
'
'
'                ValidateBFLA_Maximum_losses_Notset_Total = False
'        Exit Function
'End If
'End If
'End Function

'Function ValidateBFLA_Maximum_losses_Notset_STCG() As Boolean
'
'        ValidateBFLA_Maximum_losses_Notset_STCG = True
'        If Sheet16.Range("STCG_TOTAL_LOSS").value > Sheet16.Range("CFL_Shortterm").value Then
'        If Sheet16.Range("BF_Total_Loss").value < Sheet16.Range("STLT_Total").value And Sheet16.Range("STCG_TOTAL_LOSS").value < Sheet16.Range("CFL_STLT_Total").value Then
'                If Sheet16.Range("BF_Total_Loss").value < Sheet16.Range("CFL_STLT_Total").value Then
'                     msgbox16 ("* ""Maximum losses has not been setoff"" ")
'                ValidateBFLA_Maximum_losses_Notset_STCG = False
'        Exit Function
'End If
'End If
'End If
'End Function


'Chandru
Function ValidateTotalLoss_CYLA_New_Approach() As Boolean


        ValidateTotalLoss_CYLA_New_Approach = True

        If Sheet16.Range("OS_TotalLossCYLA").value > Sheet16.Range("cyla.TotOthSrcLossNoRaceHorse").value Then
                 msgbox16 ("* ""Total Losses set off cannot exceed losses that can be set-off and incomes available for set off."" ")
                ValidateTotalLoss_CYLA_New_Approach = False
        Exit Function
End If
'End If
End Function


'Chandru
'Function ValidateBlankData_BFLA() As Boolean
'
'
'        ValidateBlankData_BFLA = True
'                If Sheet16.Range("STCG_TOTAL_LOSS").value <= 0 Then
'                If Sheet16.Range("LTCG_TOTAL_LOSS").value = 0 Then
'
'                msgbox16 """Maximum losses has not been setoff """
'
'         ValidateBlankData_BFLA = False
'        Exit Function
'End If
'End If
'End Function

'  If Sheet16.Range("sheet16.BalHPlossCurYrAftSetoff").value > 0 Or Sheet16.Range("sheet16.BalOthSrcLossNoRaceHorseAftSetoff").value > 0 And Sheet16.Range("Total_Column4").value > 0 Then
'fmsgboxStatus """entire loss is not set off """
'End If



'Chandru
'Function ValidateBFLA_STCG_Limit_CFL() As Boolean
'
'        ValidateBFLA_STCG_Limit_CFL = True
'
'        If Sheet16.Range("STCG_TOTAL_LOSS").value > Sheet17.Range("totofbfloss.STCGLossCF8").value Then
'             msgbox16 ("* ""Losses set off cannot be more than the losses brought forward from previous years in Schedule CFL."" ")
'
'        ValidateBFLA_STCG_Limit_CFL = False
'Exit Function
'End If
'End Function
'Chandru
'Function ValidateBFLA_LTCG_Limit_CFL() As Boolean
'
'        ValidateBFLA_LTCG_Limit_CFL = True
'
'        If Sheet16.Range("LTCG_TOTAL_LOSS").value > Sheet17.Range("totofbfloss.STCGLossCF8").value + Sheet17.Range("totofbfloss.LTCGLossCF8").value Then
'             msgbox16 ("* ""Losses set off cannot be more than the losses brought forward from previous years in Schedule CFL."" ")
'
'        ValidateBFLA_LTCG_Limit_CFL = False
'Exit Function
'End If
'End Function

Function ValidateHPlossCurYrSetoff_16() As Boolean
    Dim i As Long
    
    ValidateHPlossCurYrSetoff_16 = True
    HPlossCurYrSetoff_16 = Sheet16.Range("busprof.HPlossCurYrSetoff").value
    
    If Len(HPlossCurYrSetoff_16) > 14 Then
        msgbox16 ("HPlossCurYrSetoff in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateHPlossCurYrSetoff_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(HPlossCurYrSetoff_16)
        If Not IsNumeric(HPlossCurYrSetoff_16) Then
            msgbox16 ("HPlossCurYrSetoff in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateHPlossCurYrSetoff_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBusLossSetoff_16() As Boolean
    Dim i As Long
    
    ValidateBusLossSetoff_16 = True
    BusLossSetoff_16 = Sheet16.Range("busprof.BusLossSetoff").value
    
    If Len(BusLossSetoff_16) > 14 Then
        msgbox16 ("BusLossSetoff in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBusLossSetoff_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BusLossSetoff_16)
        If Not IsNumeric(BusLossSetoff_16) Then
            msgbox16 ("BusLossSetoff in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBusLossSetoff_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateOthSrcLossNoRaceHorseSetoff_16() As Boolean
    Dim i As Long
    
    ValidateOthSrcLossNoRaceHorseSetoff_16 = True
    OthSrcLossNoRaceHorseSetoff_16 = Sheet16.Range("busprof.OthSrcLossNoRaceHorseSetoff").value
    
    If Len(OthSrcLossNoRaceHorseSetoff_16) > 14 Then
        msgbox16 ("OthSrcLossNoRaceHorseSetoff in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateOthSrcLossNoRaceHorseSetoff_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(OthSrcLossNoRaceHorseSetoff_16)
        If Not IsNumeric(OthSrcLossNoRaceHorseSetoff_16) Then
            msgbox16 ("OthSrcLossNoRaceHorseSetoff in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateOthSrcLossNoRaceHorseSetoff_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncOfCurYrAfterSetOff_16() As Boolean
    Dim i As Long
    
    ValidateIncOfCurYrAfterSetOff_16 = True
    IncOfCurYrAfterSetOff_16 = Sheet16.Range("busprof.IncOfCurYrAfterSetOff").value
    
    If Len(IncOfCurYrAfterSetOff_16) > 14 Then
        msgbox16 ("IncOfCurYrAfterSetOff in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncOfCurYrAfterSetOff_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrAfterSetOff_16)
        If Not IsNumeric(IncOfCurYrAfterSetOff_16) Then
            msgbox16 ("IncOfCurYrAfterSetOff in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrAfterSetOff_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncOfCurYrUnderThatHead1_16() As Boolean
    Dim i As Long
    
    ValidateIncOfCurYrUnderThatHead1_16 = True
    IncOfCurYrUnderThatHead1_16 = Sheet16.Range("stcg.IncOfCurYrUnderThatHead1").value
    IncOfCurYrUnderThatHead1a_16 = Sheet16.Range("stcg.IncOfCurYrUnderThatHead1a").value
    IncOfCurYrUnderThatHead1b_16 = Sheet16.Range("stcg.IncOfCurYrUnderThatHead1b").value
    
    If Len(IncOfCurYrUnderThatHead1_16) > 14 Then
        msgbox16 ("IncOfCurYrUnderThatHead1 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncOfCurYrUnderThatHead1_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrUnderThatHead1_16)
        If Not IsNumeric(IncOfCurYrUnderThatHead1_16) Then
            msgbox16 ("IncOfCurYrUnderThatHead1 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrUnderThatHead1_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateHPlossCurYrSetoff1_16() As Boolean
    Dim i As Long
    
    ValidateHPlossCurYrSetoff1_16 = True
    HPlossCurYrSetoff1_16 = Sheet16.Range("stcg.HPlossCurYrSetoff1").value
    HPlossCurYrSetoff1a_16 = Sheet16.Range("stcg.HPlossCurYrSetoff1a").value
    HPlossCurYrSetoff1b_16 = Sheet16.Range("stcg.HPlossCurYrSetoff1b").value
    
    If Len(HPlossCurYrSetoff1_16) > 14 Then
        msgbox16 ("HPlossCurYrSetoff1 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateHPlossCurYrSetoff1_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(HPlossCurYrSetoff1_16)
        If Not IsNumeric(HPlossCurYrSetoff1_16) Then
            msgbox16 ("HPlossCurYrSetoff1 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateHPlossCurYrSetoff1_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBusLossSetoff1_16() As Boolean
    Dim i As Long
    
    ValidateBusLossSetoff1_16 = True
    BusLossSetoff1_16 = Sheet16.Range("stcg.BusLossSetoff1").value
    BusLossSetoff1a_16 = Sheet16.Range("stcg.BusLossSetoff1a").value
    BusLossSetoff1b_16 = Sheet16.Range("stcg.BusLossSetoff1b").value
    
    If Len(BusLossSetoff1_16) > 14 Then
        msgbox16 ("BusLossSetoff1 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBusLossSetoff1_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BusLossSetoff1_16)
        If Not IsNumeric(BusLossSetoff1_16) Then
            msgbox16 ("BusLossSetoff1 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBusLossSetoff1_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateOthSrcLossNoRaceHorseSetoff1_16() As Boolean
    Dim i As Long

    ValidateOthSrcLossNoRaceHorseSetoff1_16 = True
    OthSrcLossNoRaceHorseSetoff1_16 = Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1").value
    OthSrcLossNoRaceHorseSetoff1a_16 = Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1a").value
    OthSrcLossNoRaceHorseSetoff1b_16 = Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1b").value
    
    If Len(OthSrcLossNoRaceHorseSetoff1_16) > 14 Then
        msgbox16 ("OthSrcLossNoRaceHorseSetoff1 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateOthSrcLossNoRaceHorseSetoff1_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(OthSrcLossNoRaceHorseSetoff1_16)
        If Not IsNumeric(OthSrcLossNoRaceHorseSetoff1_16) Then
            msgbox16 ("OthSrcLossNoRaceHorseSetoff1 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateOthSrcLossNoRaceHorseSetoff1_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncOfCurYrAfterSetOff1_16() As Boolean
    Dim i As Long
    
    ValidateIncOfCurYrAfterSetOff1_16 = True
    IncOfCurYrAfterSetOff1_16 = Sheet16.Range("stcg.IncOfCurYrAfterSetOff1").value
    IncOfCurYrAfterSetOff1a_16 = Sheet16.Range("stcg.IncOfCurYrAfterSetOff1a").value
    IncOfCurYrAfterSetOff1b_16 = Sheet16.Range("stcg.IncOfCurYrAfterSetOff1b").value
    
    If Len(IncOfCurYrAfterSetOff1_16) > 14 Then
        msgbox16 ("IncOfCurYrAfterSetOff1 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncOfCurYrAfterSetOff1_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrAfterSetOff1_16)
        If Not IsNumeric(IncOfCurYrAfterSetOff1_16) Then
            msgbox16 ("IncOfCurYrAfterSetOff1 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrAfterSetOff1_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncOfCurYrUnderThatHead2_16() As Boolean
    Dim i As Long

    ValidateIncOfCurYrUnderThatHead2_16 = True
    IncOfCurYrUnderThatHead2_16 = Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2").value
    IncOfCurYrUnderThatHead2a_16 = Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2a").value
    
    If Len(IncOfCurYrUnderThatHead2_16) > 14 Then
        msgbox16 ("IncOfCurYrUnderThatHead2 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncOfCurYrUnderThatHead2_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrUnderThatHead2_16)
        If Not IsNumeric(IncOfCurYrUnderThatHead2_16) Then
            msgbox16 ("IncOfCurYrUnderThatHead2 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrUnderThatHead2_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateHPlossCurYrSetoff2_16() As Boolean
    Dim i As Long
    
    ValidateHPlossCurYrSetoff2_16 = True
    HPlossCurYrSetoff2_16 = Sheet16.Range("ltcg.HPlossCurYrSetoff2").value
    HPlossCurYrSetoff2a_16 = Sheet16.Range("ltcg.HPlossCurYrSetoff2a").value
    
    If Len(HPlossCurYrSetoff2_16) > 14 Then
        msgbox16 ("HPlossCurYrSetoff2 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateHPlossCurYrSetoff2_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(HPlossCurYrSetoff2_16)
        If Not IsNumeric(HPlossCurYrSetoff2_16) Then
            msgbox16 ("HPlossCurYrSetoff2 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateHPlossCurYrSetoff2_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBusLossSetoff2_16() As Boolean
    Dim i As Long

    ValidateBusLossSetoff2_16 = True
    BusLossSetoff2_16 = Sheet16.Range("ltcg.BusLossSetoff2").value
    BusLossSetoff2a_16 = Sheet16.Range("ltcg.BusLossSetoff2a").value
    
    If Len(BusLossSetoff2_16) > 14 Then
        msgbox16 ("BusLossSetoff2 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBusLossSetoff2_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BusLossSetoff2_16)
        If Not IsNumeric(BusLossSetoff2_16) Then
            msgbox16 ("BusLossSetoff2 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBusLossSetoff2_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateOthSrcLossNoRaceHorseSetoff2_16() As Boolean
    Dim i As Long
    
    ValidateOthSrcLossNoRaceHorseSetoff2_16 = True
    OthSrcLossNoRaceHorseSetoff2_16 = Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2").value
    OthSrcLossNoRaceHorseSetoff2a_16 = Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").value
    
    If Len(OthSrcLossNoRaceHorseSetoff2_16) > 14 Then
        msgbox16 ("OthSrcLossNoRaceHorseSetoff2 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateOthSrcLossNoRaceHorseSetoff2_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(OthSrcLossNoRaceHorseSetoff2_16)
        If Not IsNumeric(OthSrcLossNoRaceHorseSetoff2_16) Then
            msgbox16 ("OthSrcLossNoRaceHorseSetoff2 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateOthSrcLossNoRaceHorseSetoff2_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncOfCurYrAfterSetOff2_16() As Boolean
    Dim i As Long
    
    ValidateIncOfCurYrAfterSetOff2_16 = True
    IncOfCurYrAfterSetOff2_16 = Sheet16.Range("ltcg.IncOfCurYrAfterSetOff2").value
    IncOfCurYrAfterSetOff2a_16 = Sheet16.Range("ltcg.IncOfCurYrAfterSetOff2a").value
    
    If Len(IncOfCurYrAfterSetOff2_16) > 14 Then
        msgbox16 ("IncOfCurYrAfterSetOff2 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncOfCurYrAfterSetOff2_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrAfterSetOff2_16)
        If Not IsNumeric(IncOfCurYrAfterSetOff2_16) Then
            msgbox16 ("IncOfCurYrAfterSetOff2 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrAfterSetOff2_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncOfCurYrUnderThatHead3_16() As Boolean
    Dim i As Long
    
    ValidateIncOfCurYrUnderThatHead3_16 = True
    IncOfCurYrUnderThatHead3_16 = Sheet16.Range("othSecinclnlhrs.IncOfCurYrUnderThatHead3").value
    
    If Len(IncOfCurYrUnderThatHead3_16) > 14 Then
        msgbox16 ("IncOfCurYrUnderThatHead3 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncOfCurYrUnderThatHead3_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrUnderThatHead3_16)
        If Not IsNumeric(IncOfCurYrUnderThatHead3_16) Then
            msgbox16 ("IncOfCurYrUnderThatHead3 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrUnderThatHead3_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateHPlossCurYrSetoff3_16() As Boolean
    Dim i As Long
    
    ValidateHPlossCurYrSetoff3_16 = True
    HPlossCurYrSetoff3_16 = Sheet16.Range("othSecinclnlhrs.HPlossCurYrSetoff3").value
    
    If Len(HPlossCurYrSetoff3_16) > 14 Then
        msgbox16 ("HPlossCurYrSetoff3 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateHPlossCurYrSetoff3_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(HPlossCurYrSetoff3_16)
        If Not IsNumeric(HPlossCurYrSetoff3_16) Then
            msgbox16 ("HPlossCurYrSetoff3 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateHPlossCurYrSetoff3_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBusLossSetoff3_16() As Boolean
    Dim i As Long
    
    ValidateBusLossSetoff3_16 = True
    BusLossSetoff3_16 = Sheet16.Range("othSecinclnlhrs.BusLossSetoff3").value
    
    If Len(BusLossSetoff3_16) > 14 Then
        msgbox16 ("BusLossSetoff3 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBusLossSetoff3_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BusLossSetoff3_16)
        If Not IsNumeric(BusLossSetoff3_16) Then
            msgbox16 ("BusLossSetoff3 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBusLossSetoff3_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateOthSrcLossNoRaceHorseSetoff3_16() As Boolean
    Dim i As Long
    
    ValidateOthSrcLossNoRaceHorseSetoff3_16 = True
    OthSrcLossNoRaceHorseSetoff3_16 = Sheet16.Range("othSecinclnlhrs.OthSrcLossNoRaceHorseSetoff3").value
    
    If Len(OthSrcLossNoRaceHorseSetoff3_16) > 14 Then
        msgbox16 ("OthSrcLossNoRaceHorseSetoff3 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateOthSrcLossNoRaceHorseSetoff3_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(OthSrcLossNoRaceHorseSetoff3_16)
        If Not IsNumeric(OthSrcLossNoRaceHorseSetoff3_16) Then
            msgbox16 ("OthSrcLossNoRaceHorseSetoff3 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateOthSrcLossNoRaceHorseSetoff3_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncOfCurYrAfterSetOff3_16() As Boolean
    Dim i As Long
    
    ValidateIncOfCurYrAfterSetOff3_16 = True
    IncOfCurYrAfterSetOff3_16 = Sheet16.Range("othSecinclnlhrs.IncOfCurYrAfterSetOff3").value
    
    If Len(IncOfCurYrAfterSetOff3_16) > 14 Then
        msgbox16 ("IncOfCurYrAfterSetOff3 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncOfCurYrAfterSetOff3_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrAfterSetOff3_16)
        If Not IsNumeric(IncOfCurYrAfterSetOff3_16) Then
            msgbox16 ("IncOfCurYrAfterSetOff3 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrAfterSetOff3_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotHPlossCurYrSetoff_16() As Boolean
    Dim i As Long
    
    ValidateTotHPlossCurYrSetoff_16 = True
    TotHPlossCurYrSetoff_16 = Sheet16.Range("sheet16.TotHPlossCurYrSetoff").value
    
    If Len(TotHPlossCurYrSetoff_16) > 14 Then
        msgbox16 ("TotHPlossCurYrSetoff in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateTotHPlossCurYrSetoff_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(TotHPlossCurYrSetoff_16)
        If Not IsNumeric(TotHPlossCurYrSetoff_16) Then
            msgbox16 ("TotHPlossCurYrSetoff in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateTotHPlossCurYrSetoff_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotBusLossSetoff_16() As Boolean
    Dim i As Long
    
    ValidateTotBusLossSetoff_16 = True
    TotBusLossSetoff_16 = Sheet16.Range("sheet16.TotBusLossSetoff").value
    
    If Len(TotBusLossSetoff_16) > 14 Then
        msgbox16 ("TotBusLossSetoff in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateTotBusLossSetoff_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(TotBusLossSetoff_16)
        If Not IsNumeric(TotBusLossSetoff_16) Then
            msgbox16 ("TotBusLossSetoff in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateTotBusLossSetoff_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotOthSrcLossNoRaceHorseSetoff_16() As Boolean
    Dim i As Long
    
    ValidateTotOthSrcLossNoRaceHorseSetoff_16 = True
    TotOthSrcLossNoRaceHorseSetoff_16 = Sheet16.Range("sheet16.TotOthSrcLossNoRaceHorseSetoff").value
    
    If Len(TotOthSrcLossNoRaceHorseSetoff_16) > 14 Then
        msgbox16 ("TotOthSrcLossNoRaceHorseSetoff in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateTotOthSrcLossNoRaceHorseSetoff_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(TotOthSrcLossNoRaceHorseSetoff_16)
        If Not IsNumeric(TotOthSrcLossNoRaceHorseSetoff_16) Then
            msgbox16 ("TotOthSrcLossNoRaceHorseSetoff in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateTotOthSrcLossNoRaceHorseSetoff_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBalHPlossCurYrAftSetoff_16() As Boolean
    Dim i As Long
    
    ValidateBalHPlossCurYrAftSetoff_16 = True
    BalHPlossCurYrAftSetoff_16 = Sheet16.Range("sheet16.BalHPlossCurYrAftSetoff").value
    
    If Len(BalHPlossCurYrAftSetoff_16) > 14 Then
        msgbox16 ("BalHPlossCurYrAftSetoff in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBalHPlossCurYrAftSetoff_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BalHPlossCurYrAftSetoff_16)
        If Not IsNumeric(BalHPlossCurYrAftSetoff_16) Then
            msgbox16 ("BalHPlossCurYrAftSetoff in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBalHPlossCurYrAftSetoff_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBalBusLossAftSetoff_16() As Boolean
    Dim i As Long
    
    ValidateBalBusLossAftSetoff_16 = True
    BalBusLossAftSetoff_16 = Sheet16.Range("sheet16.BalBusLossAftSetoff").value
    
    If Len(BalBusLossAftSetoff_16) > 14 Then
        msgbox16 ("BalBusLossAftSetoff in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBalBusLossAftSetoff_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BalBusLossAftSetoff_16)
        If Not IsNumeric(BalBusLossAftSetoff_16) Then
            msgbox16 ("BalBusLossAftSetoff in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBalBusLossAftSetoff_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBalOthSrcLossNoRaceHorseAftSetoff_16() As Boolean
    Dim i As Long
    
    ValidateBalOthSrcLossNoRaceHorseAftSetoff_16 = True
    BalOthSrcLossNoRaceHorseAftSetoff_16 = Sheet16.Range("sheet16.BalOthSrcLossNoRaceHorseAftSetoff").value
    
    If Len(BalOthSrcLossNoRaceHorseAftSetoff_16) > 14 Then
        msgbox16 ("BalOthSrcLossNoRaceHorseAftSetoff in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBalOthSrcLossNoRaceHorseAftSetoff_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BalOthSrcLossNoRaceHorseAftSetoff_16)
        If Not IsNumeric(BalOthSrcLossNoRaceHorseAftSetoff_16) Then
            msgbox16 ("BalOthSrcLossNoRaceHorseAftSetoff in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBalOthSrcLossNoRaceHorseAftSetoff_16 = False
            Exit Function
        End If
    Next
End Function

Function validateSalBFLA() As Boolean
Dim i As Variant
validateSalBFLA = True
salary_IncOfCurYrUndHeadFromCYLA = Sheet16.Range("salary.IncOfCurYrUndHeadFromCYLA").value
salary_IncOfCurYrAfterSetOffBFLosses = Sheet16.Range("salary.IncOfCurYrAfterSetOffBFLosses").value

If Len(salary_IncOfCurYrUndHeadFromCYLA) > 14 Then
        msgbox16 ("salary_IncOfCurYrUndHeadFromCYLA in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        validateSalBFLA = False
        Exit Function
    End If
    
    For i = 1 To Len(salary_IncOfCurYrUndHeadFromCYLA)
        If Not IsNumeric(salary_IncOfCurYrUndHeadFromCYLA) Then
            msgbox16 ("salary_IncOfCurYrUndHeadFromCYLA in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            validateSalBFLA = False
            Exit Function
        End If
    Next
    
    If Len(salary_IncOfCurYrAfterSetOffBFLosses) > 14 Then
        msgbox16 ("salary_IncOfCurYrAfterSetOffBFLosses in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        validateSalBFLA = False
        Exit Function
    End If
    
    For i = 1 To Len(salary_IncOfCurYrAfterSetOffBFLosses)
        If Not IsNumeric(salary_IncOfCurYrAfterSetOffBFLosses) Then
            msgbox16 ("salary_IncOfCurYrAfterSetOffBFLosses in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            validateSalBFLA = False
            Exit Function
        End If
    Next

End Function

Function ValidateIncOfCurYrUndHeadFromCYLA1_16() As Boolean
    Dim i As Long
    
    ValidateIncOfCurYrUndHeadFromCYLA1_16 = True
    IncOfCurYrUndHeadFromCYLA1_16 = Sheet16.Range("hp.IncOfCurYrUndHeadFromCYLA1").value
    
    If Len(IncOfCurYrUndHeadFromCYLA1_16) > 14 Then
        msgbox16 ("IncOfCurYrUndHeadFromCYLA1 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncOfCurYrUndHeadFromCYLA1_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrUndHeadFromCYLA1_16)
        If Not IsNumeric(IncOfCurYrUndHeadFromCYLA1_16) Then
            msgbox16 ("IncOfCurYrUndHeadFromCYLA1 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrUndHeadFromCYLA1_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBFlossPrevYrUndSameHeadSetoff1_16() As Boolean
    Dim i As Long
    
    ValidateBFlossPrevYrUndSameHeadSetoff1_16 = True
    BFlossPrevYrUndSameHeadSetoff1_16 = Sheet16.Range("hp.BFlossPrevYrUndSameHeadSetoff1").value
    
    If Len(BFlossPrevYrUndSameHeadSetoff1_16) > 14 Then
        msgbox16 ("BFlossPrevYrUndSameHeadSetoff1 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBFlossPrevYrUndSameHeadSetoff1_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BFlossPrevYrUndSameHeadSetoff1_16)
        If Not IsNumeric(BFlossPrevYrUndSameHeadSetoff1_16) Then
            msgbox16 ("BFlossPrevYrUndSameHeadSetoff1 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBFlossPrevYrUndSameHeadSetoff1_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBFUnabsorbedDeprSetoff1_16() As Boolean
    Dim i As Long
    
    ValidateBFUnabsorbedDeprSetoff1_16 = True
    BFUnabsorbedDeprSetoff1_16 = Sheet16.Range("hp.BFUnabsorbedDeprSetoff1").value
    
    If Len(BFUnabsorbedDeprSetoff1_16) > 14 Then
        msgbox16 ("BFUnabsorbedDeprSetoff1 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBFUnabsorbedDeprSetoff1_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BFUnabsorbedDeprSetoff1_16)
        If Not IsNumeric(BFUnabsorbedDeprSetoff1_16) Then
            msgbox16 ("BFUnabsorbedDeprSetoff1 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBFUnabsorbedDeprSetoff1_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBFAllUs35Cl4Setoff1_16() As Boolean
    Dim i As Long
    
    ValidateBFAllUs35Cl4Setoff1_16 = True
    BFAllUs35Cl4Setoff1_16 = Sheet16.Range("hp.BFAllUs35Cl4Setoff1").value
    
    If Len(BFAllUs35Cl4Setoff1_16) > 14 Then
        msgbox16 ("BFAllUs35Cl4Setoff1 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBFAllUs35Cl4Setoff1_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BFAllUs35Cl4Setoff1_16)
        If Not IsNumeric(BFAllUs35Cl4Setoff1_16) Then
            msgbox16 ("BFAllUs35Cl4Setoff1 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBFAllUs35Cl4Setoff1_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncOfCurYrAfterSetOffBFLosses1_16() As Boolean
    Dim i As Long
    
    ValidateIncOfCurYrAfterSetOffBFLosses1_16 = True
    IncOfCurYrAfterSetOffBFLosses1_16 = Sheet16.Range("hp.IncOfCurYrAfterSetOffBFLosses1").value
    
    If Len(IncOfCurYrAfterSetOffBFLosses1_16) > 14 Then
        msgbox16 ("IncOfCurYrAfterSetOffBFLosses1 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncOfCurYrAfterSetOffBFLosses1_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrAfterSetOffBFLosses1_16)
        If Not IsNumeric(IncOfCurYrAfterSetOffBFLosses1_16) Then
            msgbox16 ("IncOfCurYrAfterSetOffBFLosses1 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrAfterSetOffBFLosses1_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncOfCurYrUndHeadFromCYLA2_16() As Boolean
    Dim i As Long
    
    ValidateIncOfCurYrUndHeadFromCYLA2_16 = True
    IncOfCurYrUndHeadFromCYLA2_16 = Sheet16.Range("busipofincl.IncOfCurYrUndHeadFromCYLA2").value
    
    If Len(IncOfCurYrUndHeadFromCYLA2_16) > 14 Then
        msgbox16 ("IncOfCurYrUndHeadFromCYLA2 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncOfCurYrUndHeadFromCYLA2_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrUndHeadFromCYLA2_16)
        If Not IsNumeric(IncOfCurYrUndHeadFromCYLA2_16) Then
            msgbox16 ("IncOfCurYrUndHeadFromCYLA2 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrUndHeadFromCYLA2_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBFlossPrevYrUndSameHeadSetoff2_16() As Boolean
    Dim i As Long
    
    ValidateBFlossPrevYrUndSameHeadSetoff2_16 = True
    BFlossPrevYrUndSameHeadSetoff2_16 = Sheet16.Range("busipofincl.BFlossPrevYrUndSameHeadSetoff2").value
    
    If Len(BFlossPrevYrUndSameHeadSetoff2_16) > 14 Then
        msgbox16 ("BFlossPrevYrUndSameHeadSetoff2 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBFlossPrevYrUndSameHeadSetoff2_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BFlossPrevYrUndSameHeadSetoff2_16)
        If Not IsNumeric(BFlossPrevYrUndSameHeadSetoff2_16) Then
            msgbox16 ("BFlossPrevYrUndSameHeadSetoff2 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBFlossPrevYrUndSameHeadSetoff2_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBFUnabsorbedDeprSetoff2_16() As Boolean
    Dim i As Long
    
    ValidateBFUnabsorbedDeprSetoff2_16 = True
    BFUnabsorbedDeprSetoff2_16 = Sheet16.Range("busipofincl.BFUnabsorbedDeprSetoff2").value
    
    If Len(BFUnabsorbedDeprSetoff2_16) > 14 Then
        msgbox16 ("BFUnabsorbedDeprSetoff2 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBFUnabsorbedDeprSetoff2_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BFUnabsorbedDeprSetoff2_16)
        If Not IsNumeric(BFUnabsorbedDeprSetoff2_16) Then
            msgbox16 ("BFUnabsorbedDeprSetoff2 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBFUnabsorbedDeprSetoff2_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBFAllUs35Cl4Setoff2_16() As Boolean
    Dim i As Long
    
    ValidateBFAllUs35Cl4Setoff2_16 = True
    BFAllUs35Cl4Setoff2_16 = Sheet16.Range("busipofincl.BFAllUs35Cl4Setoff2").value
    
    If Len(BFAllUs35Cl4Setoff2_16) > 14 Then
        msgbox16 ("BFAllUs35Cl4Setoff2 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBFAllUs35Cl4Setoff2_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BFAllUs35Cl4Setoff2_16)
        If Not IsNumeric(BFAllUs35Cl4Setoff2_16) Then
            msgbox16 ("BFAllUs35Cl4Setoff2 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBFAllUs35Cl4Setoff2_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncOfCurYrAfterSetOffBFLosses2_16() As Boolean
    Dim i As Long
    
    ValidateIncOfCurYrAfterSetOffBFLosses2_16 = True
    IncOfCurYrAfterSetOffBFLosses2_16 = Sheet16.Range("busipofincl.IncOfCurYrAfterSetOffBFLosses2").value
    
    If Len(IncOfCurYrAfterSetOffBFLosses2_16) > 14 Then
        msgbox16 ("IncOfCurYrAfterSetOffBFLosses2 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncOfCurYrAfterSetOffBFLosses2_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrAfterSetOffBFLosses2_16)
        If Not IsNumeric(IncOfCurYrAfterSetOffBFLosses2_16) Then
            msgbox16 ("IncOfCurYrAfterSetOffBFLosses2 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrAfterSetOffBFLosses2_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncOfCurYrUndHeadFromCYLA3_16() As Boolean
    Dim i As Long
    
    ValidateIncOfCurYrUndHeadFromCYLA3_16 = True
    IncOfCurYrUndHeadFromCYLA3_16 = Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3").value
    IncOfCurYrUndHeadFromCYLA3a_16 = Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3a").value
    IncOfCurYrUndHeadFromCYLA3b_16 = Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3b").value
    
    If Len(IncOfCurYrUndHeadFromCYLA3_16) > 14 Then
        msgbox16 ("IncOfCurYrUndHeadFromCYLA3 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncOfCurYrUndHeadFromCYLA3_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrUndHeadFromCYLA3_16)
        If Not IsNumeric(IncOfCurYrUndHeadFromCYLA3_16) Then
            msgbox16 ("IncOfCurYrUndHeadFromCYLA3 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrUndHeadFromCYLA3_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBFlossPrevYrUndSameHeadSetoff3_16() As Boolean
    Dim i As Long
    
    ValidateBFlossPrevYrUndSameHeadSetoff3_16 = True
    BFlossPrevYrUndSameHeadSetoff3_16 = Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3").value
    BFlossPrevYrUndSameHeadSetoff3a_16 = Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3a").value
    BFlossPrevYrUndSameHeadSetoff3b_16 = Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b").value
    
    
    If Len(BFlossPrevYrUndSameHeadSetoff3_16) > 14 Then
        msgbox16 ("BFlossPrevYrUndSameHeadSetoff3 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBFlossPrevYrUndSameHeadSetoff3_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BFlossPrevYrUndSameHeadSetoff3_16)
        If Not IsNumeric(BFlossPrevYrUndSameHeadSetoff3_16) Then
            msgbox16 ("BFlossPrevYrUndSameHeadSetoff3 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBFlossPrevYrUndSameHeadSetoff3_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBFUnabsorbedDeprSetoff3_16() As Boolean
    Dim i As Long
    
    ValidateBFUnabsorbedDeprSetoff3_16 = True
    BFUnabsorbedDeprSetoff3_16 = Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3").value
    BFUnabsorbedDeprSetoff3a_16 = Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3a").value
    BFUnabsorbedDeprSetoff3b_16 = Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3b").value
    
    If Len(BFUnabsorbedDeprSetoff3_16) > 14 Then
        msgbox16 ("BFUnabsorbedDeprSetoff3 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBFUnabsorbedDeprSetoff3_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BFUnabsorbedDeprSetoff3_16)
        If Not IsNumeric(BFUnabsorbedDeprSetoff3_16) Then
            msgbox16 ("BFUnabsorbedDeprSetoff3 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBFUnabsorbedDeprSetoff3_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBFAllUs35Cl4Setoff3_16() As Boolean
    Dim i As Long
    
    ValidateBFAllUs35Cl4Setoff3_16 = True
    BFAllUs35Cl4Setoff3_16 = Sheet16.Range("stcg.BFAllUs35Cl4Setoff3").value
    BFAllUs35Cl4Setoff3a_16 = Sheet16.Range("stcg.BFAllUs35Cl4Setoff3a").value
    BFAllUs35Cl4Setoff3b_16 = Sheet16.Range("stcg.BFAllUs35Cl4Setoff3b").value
    
    If Len(BFAllUs35Cl4Setoff3_16) > 14 Then
        msgbox16 ("BFAllUs35Cl4Setoff3 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBFAllUs35Cl4Setoff3_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BFAllUs35Cl4Setoff3_16)
        If Not IsNumeric(BFAllUs35Cl4Setoff3_16) Then
            msgbox16 ("BFAllUs35Cl4Setoff3 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBFAllUs35Cl4Setoff3_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncOfCurYrAfterSetOffBFLosses3_16() As Boolean
    Dim i As Long
    
    ValidateIncOfCurYrAfterSetOffBFLosses3_16 = True
    IncOfCurYrAfterSetOffBFLosses3_16 = Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3").value
    IncOfCurYrAfterSetOffBFLosses3a_16 = Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3a").value
    IncOfCurYrAfterSetOffBFLosses3b_16 = Sheet16.Range("stcg.IncOfCurYrAfterSetOffBFLosses3b").value
    
    If Len(IncOfCurYrAfterSetOffBFLosses3_16) > 14 Then
        msgbox16 ("IncOfCurYrAfterSetOffBFLosses3 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncOfCurYrAfterSetOffBFLosses3_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrAfterSetOffBFLosses3_16)
        If Not IsNumeric(IncOfCurYrAfterSetOffBFLosses3_16) Then
            msgbox16 ("IncOfCurYrAfterSetOffBFLosses3 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrAfterSetOffBFLosses3_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncOfCurYrUndHeadFromCYLA4_16() As Boolean
    Dim i As Long
    
    ValidateIncOfCurYrUndHeadFromCYLA4_16 = True
    IncOfCurYrUndHeadFromCYLA4_16 = Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4").value
    IncOfCurYrUndHeadFromCYLA4a_16 = Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4a").value
    
    If Len(IncOfCurYrUndHeadFromCYLA4_16) > 14 Then
        msgbox16 ("IncOfCurYrUndHeadFromCYLA4 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncOfCurYrUndHeadFromCYLA4_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrUndHeadFromCYLA4_16)
        If Not IsNumeric(IncOfCurYrUndHeadFromCYLA4_16) Then
            msgbox16 ("IncOfCurYrUndHeadFromCYLA4 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrUndHeadFromCYLA4_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBFlossPrevYrUndSameHeadSetoff4_16() As Boolean
    Dim i As Long
    
    ValidateBFlossPrevYrUndSameHeadSetoff4_16 = True
    BFlossPrevYrUndSameHeadSetoff4_16 = Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4").value
    BFlossPrevYrUndSameHeadSetoff4a_16 = Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4a").value
    
    If Len(BFlossPrevYrUndSameHeadSetoff4_16) > 14 Then
        msgbox16 ("BFlossPrevYrUndSameHeadSetoff4 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBFlossPrevYrUndSameHeadSetoff4_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BFlossPrevYrUndSameHeadSetoff4_16)
        If Not IsNumeric(BFlossPrevYrUndSameHeadSetoff4_16) Then
            msgbox16 ("BFlossPrevYrUndSameHeadSetoff4 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBFlossPrevYrUndSameHeadSetoff4_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBFUnabsorbedDeprSetoff4_16() As Boolean
    Dim i As Long
    
    ValidateBFUnabsorbedDeprSetoff4_16 = True
    BFUnabsorbedDeprSetoff4_16 = Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4").value
    BFUnabsorbedDeprSetoff4a_16 = Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4a").value
    
    If Len(BFUnabsorbedDeprSetoff4_16) > 14 Then
        msgbox16 ("BFUnabsorbedDeprSetoff4 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBFUnabsorbedDeprSetoff4_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BFUnabsorbedDeprSetoff4_16)
        If Not IsNumeric(BFUnabsorbedDeprSetoff4_16) Then
            msgbox16 ("BFUnabsorbedDeprSetoff4 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBFUnabsorbedDeprSetoff4_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBFAllUs35Cl4Setoff4_16() As Boolean
    Dim i As Long
    
    ValidateBFAllUs35Cl4Setoff4_16 = True
    BFAllUs35Cl4Setoff4_16 = Sheet16.Range("ltcg.BFAllUs35Cl4Setoff4").value
    BFAllUs35Cl4Setoff4a_16 = Sheet16.Range("ltcg.BFAllUs35Cl4Setoff4a").value
    
    If Len(BFAllUs35Cl4Setoff4_16) > 14 Then
        msgbox16 ("BFAllUs35Cl4Setoff4 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBFAllUs35Cl4Setoff4_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BFAllUs35Cl4Setoff4_16)
        If Not IsNumeric(BFAllUs35Cl4Setoff4_16) Then
            msgbox16 ("BFAllUs35Cl4Setoff4 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBFAllUs35Cl4Setoff4_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncOfCurYrAfterSetOffBFLosses4_16() As Boolean
    Dim i As Long
    
    ValidateIncOfCurYrAfterSetOffBFLosses4_16 = True
    IncOfCurYrAfterSetOffBFLosses4_16 = Sheet16.Range("ltcg.IncOfCurYrAfterSetOffBFLosses4").value
    IncOfCurYrAfterSetOffBFLosses4a_16 = Sheet16.Range("ltcg.IncOfCurYrAfterSetOffBFLosses4a").value
    
    If Len(IncOfCurYrAfterSetOffBFLosses4_16) > 14 Then
        msgbox16 ("IncOfCurYrAfterSetOffBFLosses4 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncOfCurYrAfterSetOffBFLosses4_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrAfterSetOffBFLosses4_16)
        If Not IsNumeric(IncOfCurYrAfterSetOffBFLosses4_16) Then
            msgbox16 ("IncOfCurYrAfterSetOffBFLosses4 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrAfterSetOffBFLosses4_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncOfCurYrUndHeadFromCYLA5_16() As Boolean
    Dim i As Long
    
    ValidateIncOfCurYrUndHeadFromCYLA5_16 = True
    IncOfCurYrUndHeadFromCYLA5_16 = Sheet16.Range("othsrcincl.IncOfCurYrUndHeadFromCYLA5").value
    
    If Len(IncOfCurYrUndHeadFromCYLA5_16) > 14 Then
        msgbox16 ("IncOfCurYrUndHeadFromCYLA5 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncOfCurYrUndHeadFromCYLA5_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrUndHeadFromCYLA5_16)
        If Not IsNumeric(IncOfCurYrUndHeadFromCYLA5_16) Then
            msgbox16 ("IncOfCurYrUndHeadFromCYLA5 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrUndHeadFromCYLA5_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBFlossPrevYrUndSameHeadSetoff5_16() As Boolean
    Dim i As Long
    
    ValidateBFlossPrevYrUndSameHeadSetoff5_16 = True
    BFlossPrevYrUndSameHeadSetoff5_16 = Sheet16.Range("othsrcincl.BFlossPrevYrUndSameHeadSetoff5").value
    
    If Len(BFlossPrevYrUndSameHeadSetoff5_16) > 14 Then
        msgbox16 ("BFlossPrevYrUndSameHeadSetoff5 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBFlossPrevYrUndSameHeadSetoff5_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BFlossPrevYrUndSameHeadSetoff5_16)
        If Not IsNumeric(BFlossPrevYrUndSameHeadSetoff5_16) Then
            msgbox16 ("BFlossPrevYrUndSameHeadSetoff5 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBFlossPrevYrUndSameHeadSetoff5_16 = False
            Exit Function
        End If
    Next
End Function


Function ValidateBFUnabsorbedDeprSetoff5_16() As Boolean
    Dim i As Long
    
    ValidateBFUnabsorbedDeprSetoff5_16 = True
    BFUnabsorbedDeprSetoff5_16 = Sheet16.Range("othsrcincl.BFUnabsorbedDeprSetoff5").value
    
    If Len(BFUnabsorbedDeprSetoff5_16) > 14 Then
        msgbox16 ("BFUnabsorbedDeprSetoff5 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBFUnabsorbedDeprSetoff5_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BFUnabsorbedDeprSetoff5_16)
        If Not IsNumeric(BFUnabsorbedDeprSetoff5_16) Then
            msgbox16 ("BFUnabsorbedDeprSetoff5 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBFUnabsorbedDeprSetoff5_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateBFAllUs35Cl4Setoff5_16() As Boolean
    Dim i As Long
    
    ValidateBFAllUs35Cl4Setoff5_16 = True
    BFAllUs35Cl4Setoff5_16 = Sheet16.Range("othsrcincl.BFAllUs35Cl4Setoff5").value
    
    If Len(BFAllUs35Cl4Setoff5_16) > 14 Then
        msgbox16 ("BFAllUs35Cl4Setoff5 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateBFAllUs35Cl4Setoff5_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(BFAllUs35Cl4Setoff5_16)
        If Not IsNumeric(BFAllUs35Cl4Setoff5_16) Then
            msgbox16 ("BFAllUs35Cl4Setoff5 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateBFAllUs35Cl4Setoff5_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncOfCurYrAfterSetOffBFLosses5_16() As Boolean
    Dim i As Long
    
    ValidateIncOfCurYrAfterSetOffBFLosses5_16 = True
    IncOfCurYrAfterSetOffBFLosses5_16 = Sheet16.Range("othsrcincl.IncOfCurYrAfterSetOffBFLosses5").value
    
    If Len(IncOfCurYrAfterSetOffBFLosses5_16) > 14 Then
        msgbox16 ("IncOfCurYrAfterSetOffBFLosses5 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncOfCurYrAfterSetOffBFLosses5_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncOfCurYrAfterSetOffBFLosses5_16)
        If Not IsNumeric(IncOfCurYrAfterSetOffBFLosses5_16) Then
            msgbox16 ("IncOfCurYrAfterSetOffBFLosses5 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncOfCurYrAfterSetOffBFLosses5_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotBFLossSetoff_16() As Boolean
    Dim i As Long
    
    ValidateTotBFLossSetoff_16 = True
    TotBFLossSetoff_16 = Sheet16.Range("sheet16.TotBFLossSetoff").value
    
    If Len(TotBFLossSetoff_16) > 14 Then
        msgbox16 ("TotBFLossSetoff in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateTotBFLossSetoff_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(TotBFLossSetoff_16)
        If Not IsNumeric(TotBFLossSetoff_16) Then
            msgbox16 ("TotBFLossSetoff in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateTotBFLossSetoff_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotUnabsorbedDeprSetoff_16() As Boolean
    Dim i As Long
    
    ValidateTotUnabsorbedDeprSetoff_16 = True
    TotUnabsorbedDeprSetoff_16 = Sheet16.Range("sheet16.TotUnabsorbedDeprSetoff").value
    
    If Len(TotUnabsorbedDeprSetoff_16) > 14 Then
        msgbox16 ("TotUnabsorbedDeprSetoff in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateTotUnabsorbedDeprSetoff_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(TotUnabsorbedDeprSetoff_16)
        If Not IsNumeric(TotUnabsorbedDeprSetoff_16) Then
            msgbox16 ("TotUnabsorbedDeprSetoff in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateTotUnabsorbedDeprSetoff_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotAllUs35cl4Setoff_16() As Boolean
    Dim i As Long
    
    ValidateTotAllUs35cl4Setoff_16 = True
    TotAllUs35cl4Setoff_16 = Sheet16.Range("sheet16.TotAllUs35cl4Setoff").value
    
    If Len(TotAllUs35cl4Setoff_16) > 14 Then
        msgbox16 ("TotAllUs35cl4Setoff in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateTotAllUs35cl4Setoff_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(TotAllUs35cl4Setoff_16)
        If Not IsNumeric(TotAllUs35cl4Setoff_16) Then
            msgbox16 ("TotAllUs35cl4Setoff in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateTotAllUs35cl4Setoff_16 = False
            Exit Function
        End If
    Next
End Function

Function ValidateIncomeOfCurrYrAftCYLABFLA_16() As Boolean
    Dim i As Long
    
    ValidateIncomeOfCurrYrAftCYLABFLA_16 = True
    IncomeOfCurrYrAftCYLABFLA_16 = Sheet16.Range("sheet16.IncomeOfCurrYrAftCYLABFLA").value
    
    If Len(IncomeOfCurrYrAftCYLABFLA_16) > 14 Then
        msgbox16 ("IncomeOfCurrYrAftCYLABFLA in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        ValidateIncomeOfCurrYrAftCYLABFLA_16 = False
        Exit Function
    End If
    
    For i = 1 To Len(IncomeOfCurrYrAftCYLABFLA_16)
        If Not IsNumeric(IncomeOfCurrYrAftCYLABFLA_16) Then
            msgbox16 ("IncomeOfCurrYrAftCYLABFLA in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            ValidateIncomeOfCurrYrAftCYLABFLA_16 = False
            Exit Function
        End If
    Next
End Function

Function DefaultIncOfCurYrUnderThatHead_16() As String
    IncOfCurYrUnderThatHead_16 = "0"
End Function

Function DefaultHPlossCurYrSetoff_16() As String
    HPlossCurYrSetoff_16 = "0"
End Function

Function DefaultBusLossSetoff_16() As String
    BusLossSetoff_16 = "0"
End Function

Function DefaultOthSrcLossNoRaceHorseSetoff_16() As String
    OthSrcLossNoRaceHorseSetoff_16 = "0"
End Function

Function DefaultIncOfCurYrAfterSetOff_16() As String
    IncOfCurYrAfterSetOff_16 = "0"
End Function

Function DefaultIncOfCurYrUnderThatHead1_16() As String
    IncOfCurYrUnderThatHead1_16 = "0"
End Function

Function DefaultHPlossCurYrSetoff1_16() As String
    HPlossCurYrSetoff1_16 = "0"
End Function

Function DefaultBusLossSetoff1_16() As String
    BusLossSetoff1_16 = "0"
End Function

Function DefaultOthSrcLossNoRaceHorseSetoff1_16() As String
    OthSrcLossNoRaceHorseSetoff1_16 = "0"
End Function

Function DefaultIncOfCurYrAfterSetOff1_16() As String
    IncOfCurYrAfterSetOff1_16 = "0"
End Function

Function DefaultIncOfCurYrUnderThatHead2_16() As String
    IncOfCurYrUnderThatHead2_16 = "0"
End Function

Function DefaultHPlossCurYrSetoff2_16() As String
    HPlossCurYrSetoff2_16 = "0"
End Function

Function DefaultBusLossSetoff2_16() As String
    BusLossSetoff2_16 = "0"
End Function

Function DefaultOthSrcLossNoRaceHorseSetoff2_16() As String
    OthSrcLossNoRaceHorseSetoff2_16 = "0"
End Function

Function DefaultIncOfCurYrAfterSetOff2_16() As String
    IncOfCurYrAfterSetOff2_16 = "0"
End Function

Function DefaultIncOfCurYrUnderThatHead3_16() As String
    IncOfCurYrUnderThatHead3_16 = "0"
End Function

Function DefaultHPlossCurYrSetoff3_16() As String
    HPlossCurYrSetoff3_16 = "0"
End Function

Function DefaultBusLossSetoff3_16() As String
    BusLossSetoff3_16 = "0"
End Function

Function DefaultOthSrcLossNoRaceHorseSetoff3_16() As String
    OthSrcLossNoRaceHorseSetoff3_16 = "0"
End Function

Function DefaultIncOfCurYrAfterSetOff3_16() As String
    IncOfCurYrAfterSetOff3_16 = "0"
End Function

Function DefaultTotHPlossCurYrSetoff_16() As String
    TotHPlossCurYrSetoff_16 = "0"
End Function

Function DefaultTotBusLossSetoff_16() As String
    TotBusLossSetoff_16 = "0"
End Function

Function DefaultTotOthSrcLossNoRaceHorseSetoff_16() As String
    TotOthSrcLossNoRaceHorseSetoff_16 = "0"
End Function

Function DefaultBalHPlossCurYrAftSetoff_16() As String
    BalHPlossCurYrAftSetoff_16 = "0"
End Function

Function DefaultBalBusLossAftSetoff_16() As String
    BalBusLossAftSetoff_16 = "0"
End Function

Function DefaultBalOthSrcLossNoRaceHorseAftSetoff_16() As String
    BalOthSrcLossNoRaceHorseAftSetoff_16 = "0"
End Function

Function DefaultIncOfCurYrUndHeadFromCYLA_16() As String
    IncOfCurYrUndHeadFromCYLA_16 = "0"
End Function

Function DefaultBFlossPrevYrUndSameHeadSetoff_16() As String
    BFlossPrevYrUndSameHeadSetoff_16 = "0"
End Function

Function DefaultBFUnabsorbedDeprSetoff_16() As String
    BFUnabsorbedDeprSetoff_16 = "0"
End Function

Function DefaultBFAllUs35Cl4Setoff_16() As String
    BFAllUs35Cl4Setoff_16 = "0"
End Function

Function DefaultIncOfCurYrAfterSetOffBFLosses_16() As String
    IncOfCurYrAfterSetOffBFLosses_16 = "0"
End Function

Function DefaultIncOfCurYrUndHeadFromCYLA1_16() As String
    IncOfCurYrUndHeadFromCYLA1_16 = "0"
End Function

Function DefaultBFlossPrevYrUndSameHeadSetoff1_16() As String
    BFlossPrevYrUndSameHeadSetoff1_16 = "0"
End Function

Function DefaultBFUnabsorbedDeprSetoff1_16() As String
    BFUnabsorbedDeprSetoff1_16 = "0"
End Function

Function DefaultBFAllUs35Cl4Setoff1_16() As String
    BFAllUs35Cl4Setoff1_16 = "0"
End Function

Function DefaultIncOfCurYrAfterSetOffBFLosses1_16() As String
    IncOfCurYrAfterSetOffBFLosses1_16 = "0"
End Function

Function DefaultIncOfCurYrUndHeadFromCYLA2_16() As String
    IncOfCurYrUndHeadFromCYLA2_16 = "0"
End Function

Function DefaultBFlossPrevYrUndSameHeadSetoff2_16() As String
    BFlossPrevYrUndSameHeadSetoff2_16 = "0"
End Function

Function DefaultBFUnabsorbedDeprSetoff2_16() As String
    BFUnabsorbedDeprSetoff2_16 = "0"
End Function

Function DefaultBFAllUs35Cl4Setoff2_16() As String
    BFAllUs35Cl4Setoff2_16 = "0"
End Function

Function DefaultIncOfCurYrAfterSetOffBFLosses2_16() As String
    IncOfCurYrAfterSetOffBFLosses2_16 = "0"
End Function

Function DefaultIncOfCurYrUndHeadFromCYLA3_16() As String
    IncOfCurYrUndHeadFromCYLA3_16 = "0"
End Function

Function DefaultBFlossPrevYrUndSameHeadSetoff3_16() As String
    BFlossPrevYrUndSameHeadSetoff3_16 = "0"
End Function

Function DefaultBFUnabsorbedDeprSetoff3_16() As String
    BFUnabsorbedDeprSetoff3_16 = "0"
End Function

Function DefaultBFAllUs35Cl4Setoff3_16() As String
    BFAllUs35Cl4Setoff3_16 = "0"
End Function

Function DefaultIncOfCurYrAfterSetOffBFLosses3_16() As String
    IncOfCurYrAfterSetOffBFLosses3_16 = "0"
End Function

Function DefaultIncOfCurYrUndHeadFromCYLA4_16() As String
    IncOfCurYrUndHeadFromCYLA4_16 = "0"
End Function

Function DefaultBFlossPrevYrUndSameHeadSetoff4_16() As String
    BFlossPrevYrUndSameHeadSetoff4_16 = "0"
End Function

Function DefaultBFUnabsorbedDeprSetoff4_16() As String
    BFUnabsorbedDeprSetoff4_16 = "0"
End Function

Function DefaultBFAllUs35Cl4Setoff4_16() As String
    BFAllUs35Cl4Setoff4_16 = "0"
End Function

Function DefaultIncOfCurYrAfterSetOffBFLosses4_16() As String
    IncOfCurYrAfterSetOffBFLosses4_16 = "0"
End Function

Function DefaultIncOfCurYrUndHeadFromCYLA5_16() As String
    IncOfCurYrUndHeadFromCYLA5_16 = "0"
End Function

Function DefaultBFlossPrevYrUndSameHeadSetoff5_16() As String
    BFlossPrevYrUndSameHeadSetoff5_16 = "0"
End Function

Function DefaultBFUnabsorbedDeprSetoff5_16() As String
    BFUnabsorbedDeprSetoff5_16 = "0"
End Function

Function DefaultBFAllUs35Cl4Setoff5_16() As String
    BFAllUs35Cl4Setoff5_16 = "0"
End Function

Function DefaultIncOfCurYrAfterSetOffBFLosses5_16() As String
    IncOfCurYrAfterSetOffBFLosses5_16 = "0"
End Function

Function DefaultTotBFLossSetoff_16() As String
    TotBFLossSetoff_16 = "0"
End Function

Function DefaultTotUnabsorbedDeprSetoff_16() As String
    TotUnabsorbedDeprSetoff_16 = "0"
End Function

Function DefaultTotAllUs35cl4Setoff_16() As String
    TotAllUs35cl4Setoff_16 = "0"
End Function

Function DefaultIncomeOfCurrYrAftCYLABFLA_16() As String
    IncomeOfCurrYrAftCYLABFLA_16 = "0"
End Function

Function msgbox15(strmsg As String) As String
    msgValidateSheet15 = msgValidateSheet15 & strmsg & Chr(13)
End Function
Function msgbox16(strmsg As String) As String
    msgValidateSheet16 = msgValidateSheet16 & strmsg & Chr(13)
End Function

Function validatesalary() As Boolean
Dim i As Long
validatesalary = True
salary_IncOfCurYrUnderThatHead1 = Sheet16.Range("salary.IncOfCurYrUnderThatHead1").value
salary_HPlossCurYrSetoff1 = Sheet16.Range("salary.HPlossCurYrSetoff1").value
salary_OthSrcLossNoRaceHorseSetoff1 = Sheet16.Range("salary.OthSrcLossNoRaceHorseSetoff1").value
salary_IncOfCurYrAfterSetOff1 = Sheet16.Range("salary.IncOfCurYrAfterSetOff1").value

If Len(salary_IncOfCurYrUnderThatHead1) > 14 Then
        msgbox16 ("salary_IncOfCurYrUnderThatHead1 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        validatesalary = False
        Exit Function
    End If
    
For i = 1 To Len(salary_IncOfCurYrUnderThatHead1)
    If Not IsNumeric(salary_IncOfCurYrUnderThatHead1) Then
            msgbox16 ("salary_IncOfCurYrUnderThatHead1 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            validatesalary = False
            Exit Function
    End If
Next

If Len(salary_HPlossCurYrSetoff1) > 14 Then
        msgbox16 ("salary_HPlossCurYrSetoff1 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        validatesalary = False
        Exit Function
    End If
    
For i = 1 To Len(salary_HPlossCurYrSetoff1)
    If Not IsNumeric(salary_HPlossCurYrSetoff1) Then
            msgbox16 ("salary_HPlossCurYrSetoff1 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            validatesalary = False
            Exit Function
    End If
Next

If Len(salary_OthSrcLossNoRaceHorseSetoff1) > 14 Then
        msgbox16 ("salary_OthSrcLossNoRaceHorseSetoff1 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        validatesalary = False
        Exit Function
    End If
    
For i = 1 To Len(salary_OthSrcLossNoRaceHorseSetoff1)
    If Not IsNumeric(salary_OthSrcLossNoRaceHorseSetoff1) Then
            msgbox16 ("salary_OthSrcLossNoRaceHorseSetoff1 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            validatesalary = False
            Exit Function
    End If
Next

If Len(salary_IncOfCurYrAfterSetOff1) > 14 Then
        msgbox16 ("salary_IncOfCurYrAfterSetOff1 in Sheet : CYLA - BFLA  should not greater than 14 digits ")
        validatesalary = False
        Exit Function
    End If
    
For i = 1 To Len(salary_IncOfCurYrAfterSetOff1)
    If Not IsNumeric(salary_IncOfCurYrAfterSetOff1) Then
            msgbox16 ("salary_IncOfCurYrAfterSetOff1 in Sheet : CYLA - BFLA  must contain only digits from 0 to 9")
            validatesalary = False
            Exit Function
    End If
Next

End Function

Function ValidateLossOfCurYrUnder() As Boolean
    ValidateLossOfCurYrUnder = True
    LossOfUnderThatHead_HP = Sheet16.Range("cyla.TotHPlossCurYr").value
    LossOfUnderThatHead_BP = Sheet16.Range("cyla.TotBusLoss").value
    LossOfUnderThatHead_OS = Sheet16.Range("cyla.TotOthSrcLossNoRaceHorse").value
End Function

Sub CYLABFLAsetoffclick()
Dim name As Variant
name = ActiveSheet.name
CYLABFLASetOff
Sheets(name).Activate
MsgBox "Please give quarterly break of LTCG and STCG according to values in Schedule BFLA in schedule CG table F"
End Sub

Function CYLABFLASetOff()
'On Error GoTo endline
 Application.EnableEvents = False
 Application.ScreenUpdating = False


'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   
   Sheet31.Unprotect Password:=sPassword   'CG
   Sheet17.Unprotect Password:=sPassword    'CFL
   Sheet16.Unprotect Password:=sPassword  'CYLA-BFLA
'----------------Unlock Password-------------------END-----

setOffOthSrcLossCYLA '--STCG 20% DONE
setOffHPLossCYLA '--STCG 20% DONE
setOffBussLossCYLA '--STCG 20% DONE
calcBFLA
calcCFL
      
'----------------Lock Password-------------------START---
  Sheet31.Protect Password:=sPassword
  Sheet17.Protect Password:=sPassword
  Sheet16.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
endline:
Application.EnableEvents = True
Application.ScreenUpdating = True
End Function

Function setOffBussLossCYLA()

Application.EnableEvents = False

Dim hpSetOffRem As Double
Dim prevSetOffRem As Double
Dim setOffRem As Double
Dim tempAk As Double

'Commented by Riyaz on 29/01/2026
'Dim stcgPercent15Sec111a As Double
'Dim stcgPercent15Sec115ad_1_b_ii As Double
'--STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
Dim stcgPercent20Sec111a As Double
Dim stcgPercent20Sec115ad_1_b_ii As Double

'Commented by Riyaz on 29/01/2026
'Dim ltcgPrctg20Sec112 As Double
'Dim ltcgPrctg20Sec11EA As Double
'Newly added by Bindu as per DE V4 on 15th March 25
'Dim ltcgPrctg20SecProviso As Double

'Commented by Riyaz on 29/01/2026
'Dim ltcgPrctg10SecProviso As Double
'Dim ltcgPrctg10Sec112c2 As Double
'Dim ltcgPrctg10Sec115AC1 As Double
'Dim ltcgPrctg10Sec115ACA1 As Double
'Dim ltcgPrctg10Sec115AD3 As Double
'Dim ltcgPrctg10Sec115AD3_Proviso As Double
'Dim ltcgPrctg10Sec112A As Double
'Dim ltcgPrctg10Sec115Eb As Double

'--LTCG 12.5% NEWLY ADDED BY BINDU ON 25TH FEN 25
Dim ltcgPrctg125Sec115ACA1 As Double
Dim ltcgPrctg125Sec115Eb As Double
Dim ltcgPrctg125Sec1121c2 As Double
Dim ltcgPrctg125Sec115AC1 As Double
Dim ltcgPrctg125Sec115Ea As Double
Dim ltcgPrctg125Sec115AD3_Proviso As Double
Dim ltcgPrctg125Sec112A As Double
Dim ltcgPTI112Sec48 As Double
Dim ltcgPTI125112B9 As Double
Dim ltcgPrctg125SecProviso As Double '15th March
Dim ltcgPrctg125Sec115AD3 As Double

'Commented by Riyaz on 29/01/2026
'Dim ltcgPTI10 As Double
'Dim ltcgPTI10_112 As Double
Dim stcgPTI30 As Double
Dim stcgPrctg30Sec115AD As Double
'Dim stcgPTI15 As Double
'---STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
Dim stcgPTI20 As Double

Dim ltcgPTI20 As Double
'--LTCG 12.5% NEWLY ADDED BY BINDU ON 25TH FEB 25
Dim ltcgPTI125  As Double
Dim ltcgPTI125_112  As Double

'Commented by Riyaz on 29/01/2026
'stcgPercent15Sec111a = cgosIncome_cgInc_stcg_prctg15_sec111a
'stcgPercent15Sec115ad_1_b_ii = cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii
'---STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
stcgPercent20Sec111a = cgosIncome_cgInc_stcg_prctg20_sec111a
stcgPercent20Sec115ad_1_b_ii = cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii

'Commented by Riyaz on 29/01/2026
'ltcgPrctg20Sec112 = cgosIncome_cgInc_ltcg_prctg20_sec112
'ltcgPrctg20Sec11EA = cgosIncome_cgInc_ltcg_prctg20_sec11EA
'Newly added by Bindu as per DE V4 on 15th March 25
'ltcgPrctg20SecProviso = cgosIncome_cgInc_ltcg_prctg20_secProviso

'Commented by Riyaz on 29/01/2026
'ltcgPrctg10SecProviso = cgosIncome_cgInc_ltcg_prctg10_secProviso ''NOt using  commented  by Bindu as this sec changed from 10 to 20
'ltcgPrctg10Sec112c2 = cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2
'ltcgPrctg10Sec115AC1 = cgosIncome_cgInc_ltcg_prctg10_sec115AC_1
'ltcgPrctg10Sec115ACA1 = cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1
'ltcgPrctg10Sec115AD3 = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3
'ltcgPrctg10Sec115AD3_Proviso = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso
'ltcgPrctg10Sec112A = cgosIncome_cgInc_ltcg_prctg10_sec112A
'ltcgPrctg10Sec115Eb = cgosIncome_cgInc_ltcg_prctg10_sec115E_b

'---LTCG 12.5% NEWLY ADDED BY BINDU 26TH FEB 25
 ltcgPrctg125Sec115ACA1 = cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1
 ltcgPrctg125Sec115Eb = cgosIncome_cgInc_ltcg_prctg125_sec115E_b
 ltcgPrctg125Sec1121c2 = cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2
 ltcgPrctg125Sec115AC1 = cgosIncome_cgInc_ltcg_prctg125_sec115AC_1
 ltcgPrctg125Sec115Ea = cgosIncome_cgInc_ltcg_prctg125_sec115E_a
 ltcgPrctg125Sec115AD3_Proviso = cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso
 ltcgPrctg125Sec112A = cgosIncome_cgInc_ltcg_prctg125_sec112A
 ltcgPTI125 = cgosIncome_cgInc_ltcg_PTI125
 ltcgPTI125_112 = cgosIncome_cgInc_ltcg_PTI125_112
 ltcgPTI112Sec48 = cgosIncome_cgInc_ltcg_PTI112_sec48
 ltcgPTI125112B9 = cgosIncome_cgInc_ltcg_PTI125_112_B9
 ltcgPrctg125SecProviso = cgosIncome_cgInc_ltcg_prctg125_secProviso '15th March
 ltcgPrctg125Sec115AD3 = cgosIncome_cgInc_ltcg_prctg125_sec115AD_3
'--END


 stcgPrctg30Sec115AD = cgosIncome_cgInc_stcg_prctg30_sec115AD
 stcgPTI30 = cgosIncome_cgInc_stcg_PTI30
'Commented by Riyaz on 29/01/2026
' stcgPTI15 = cgosIncome_cgInc_stcg_PTI15



stcgPrctg30Sec115AD = cgosIncome_cgInc_stcg_prctg30_sec115AD
 stcgPTI30 = cgosIncome_cgInc_stcg_PTI30
 'Commented by Riyaz on 29/01/2026
' stcgPTI15 = cgosIncome_cgInc_stcg_PTI15
 '---STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
 stcgPTI20 = cgosIncome_cgInc_stcg_PTI20
 
'Commented by Riyaz on 29/01/2026
' ltcgPTI10 = cgosIncome_cgInc_ltcg_PTI10
'  ltcgPTI10_112 = cgosIncome_cgInc_ltcg_PTI10_112
' ltcgPTI20 = cgosIncome_cgInc_ltcg_PTI20


' 1. HP Income
hpSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("hp.IncOfCurYrUnderThatHead2").value - Sheet16.Range("hp.OthSrcLossNoRaceHorseSetoff2").value)
prevSetOffRem = Sheet16.Range("cyla.TotBusLoss").value
setOffRem = prevSetOffRem - hpSetOffRem
        
        If (setOffRem >= 0) Then
            If (Sheet16.Range("cyla.TotBusLoss").value = 0) Then
                Sheet16.Range("hp.BusLossSetoff2").value = 0
            Else
                    Sheet16.Range("hp.BusLossSetoff2").value = hpSetOffRem
            End If
        Else
                Sheet16.Range("hp.BusLossSetoff2").value = Sheet16.Range("cyla.TotBusLoss").value
            setOffRem = 0
        End If
        
'2. Other sources
Dim osIncSetOffRem As Double
         osIncSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("othSecinclnlhrs.IncOfCurYrUnderThatHead3").value - Sheet16.Range("othSecinclnlhrs.HPlossCurYrSetoff3").value)
        prevSetOffRem = setOffRem
        setOffRem = prevSetOffRem - osIncSetOffRem
        
        If (setOffRem >= 0) Then
            If (prevSetOffRem = 0) Then
                Sheet16.Range("othSecinclnlhrs.BusLossSetoff3").value = 0
            Else
                    Sheet16.Range("othSecinclnlhrs.BusLossSetoff3").value = osIncSetOffRem
            End If
        Else
            Sheet16.Range("othSecinclnlhrs.BusLossSetoff3").value = prevSetOffRem
            setOffRem = 0
        End If
        
'3) Profit from owning and maintaining race horse
Dim raceHrsSetOffRem As Double

 raceHrsSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("rh.IncOfCurYrUnderThatHead4").value - Sheet16.Range("rh.HPlossCurYrSetoff4").value - Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4").value)
        
        prevSetOffRem = setOffRem
        setOffRem = setOffRem - raceHrsSetOffRem
        If (setOffRem >= 0) Then
            If (prevSetOffRem = 0) Then
                Sheet16.Range("rh.BusLossSetoff4").value = 0
            Else
                Sheet16.Range("rh.BusLossSetoff4").value = raceHrsSetOffRem
            End If
        Else
                Sheet16.Range("rh.BusLossSetoff4").value = prevSetOffRem
            setOffRem = 0
        End If
        
 '4 OS DTAA
   Dim OSDTAASetOffRem As Double

 OSDTAASetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("rh.IncOfCurYrUnderThatHead4_DTAARates").value - Sheet16.Range("rh.HPlossCurYrSetoff4_DTAARates").value - Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4_DTAARates").value)

        prevSetOffRem = setOffRem
        setOffRem = setOffRem - OSDTAASetOffRem
        If (setOffRem >= 0) Then
            If (prevSetOffRem = 0) Then
                Sheet16.Range("rh.BusLossSetoff4_DTAARates").value = 0
            Else
                Sheet16.Range("rh.BusLossSetoff4_DTAARates").value = OSDTAASetOffRem
            End If
        Else
                Sheet16.Range("rh.BusLossSetoff4_DTAARates").value = prevSetOffRem
            setOffRem = 0
        End If

 
 
'5) Short-term capital gain taxable @ 30%
Dim stcg30PercntSetOffRem As Double
        


stcg30PercntSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("stcg.IncOfCurYrUnderThatHead1a").value - Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1a").value - Sheet16.Range("stcg.HPlossCurYrSetoff1a").value)
        
        prevSetOffRem = setOffRem
        setOffRem = setOffRem - stcg30PercntSetOffRem
        If (setOffRem >= 0) Then
            If (prevSetOffRem = 0) Then
                 Sheet16.Range("stcg.BusLossSetoff1a").value = 0
            Else
                     Sheet16.Range("stcg.BusLossSetoff1a").value = stcg30PercntSetOffRem
                    stcg30PercntSetOffRem = 0
                    stcgPTI30 = 0
                    stcgPrctg30Sec115AD = 0
                    
            End If
        Else
                Sheet16.Range("stcg.BusLossSetoff1a").value = prevSetOffRem
                
                ' sec break up
                If (prevSetOffRem >= 0 And stcgPrctg30Sec115AD > prevSetOffRem) Then
                    stcgPrctg30Sec115AD = stcgPrctg30Sec115AD - prevSetOffRem
                    prevSetOffRem = 0
                Else
                    prevSetOffRem = prevSetOffRem - stcgPrctg30Sec115AD
                    stcgPrctg30Sec115AD = 0
                End If

                If (prevSetOffRem >= 0 And stcgPTI30 > prevSetOffRem) Then
                    stcgPTI30 = stcgPTI30 - prevSetOffRem
                    prevSetOffRem = 0
                Else
                    prevSetOffRem = prevSetOffRem - stcgPTI30
                    stcgPTI30 = 0
                End If
                
            setOffRem = 0
    End If
        
'6) Short-term capital gain taxable at applicable rates
Dim stcgAppRateSetOffRem As Double
        
   stcgAppRateSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("stcg.IncOfCurYrUnderThatHead1b").value - Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1b").value - Sheet16.Range("stcg.HPlossCurYrSetoff1b").value)
        
        prevSetOffRem = setOffRem
        setOffRem = setOffRem - stcgAppRateSetOffRem
        If (setOffRem >= 0) Then
            If (prevSetOffRem = 0) Then
               Sheet16.Range("stcg.BusLossSetoff1b").value = 0
            Else
                Sheet16.Range("stcg.BusLossSetoff1b").value = stcgAppRateSetOffRem
            End If
        Else
               Sheet16.Range("stcg.BusLossSetoff1b").value = prevSetOffRem
                setOffRem = 0
        End If
        
        
'--STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
Dim stcg20PercntSetOffRem As Double
stcg20PercntSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("stcg.IncOfCurYrUnderThatHead1c").value - Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1c").value - Sheet16.Range("stcg.HPlossCurYrSetoff1c").value)
        
        prevSetOffRem = setOffRem
        setOffRem = setOffRem - stcg20PercntSetOffRem
        If (setOffRem >= 0) Then
            If (prevSetOffRem = 0) Then
                 Sheet16.Range("stcg.BusLossSetoff1c").value = 0
            Else
                     Sheet16.Range("stcg.BusLossSetoff1c").value = stcg20PercntSetOffRem
                    stcg20PercntSetOffRem = 0
                    stcgPercent20Sec111a = 0
                    stcgPTI20 = 0
                    stcgPercent20Sec115ad_1_b_ii = 0
                    
            End If
        Else
                Sheet16.Range("stcg.BusLossSetoff1c").value = prevSetOffRem
                
                ' sec break up
                If (prevSetOffRem >= 0 And stcgPercent20Sec115ad_1_b_ii > prevSetOffRem) Then
                    stcgPercent20Sec115ad_1_b_ii = stcgPercent20Sec115ad_1_b_ii - prevSetOffRem
                    prevSetOffRem = 0
                Else
                    prevSetOffRem = prevSetOffRem - stcgPercent20Sec115ad_1_b_ii
                    stcgPercent20Sec115ad_1_b_ii = 0
                End If
                
                If (prevSetOffRem >= 0 And stcgPTI20 > prevSetOffRem) Then
                    stcgPTI20 = stcgPTI20 - prevSetOffRem
                    prevSetOffRem = 0
                Else
                    prevSetOffRem = prevSetOffRem - stcgPTI20
                    stcgPTI20 = 0
                End If


                If (prevSetOffRem >= 0 And stcgPercent20Sec111a > prevSetOffRem) Then
                    stcgPercent20Sec111a = stcgPercent20Sec111a - prevSetOffRem
                    prevSetOffRem = 0
                Else
                    prevSetOffRem = prevSetOffRem - stcgPercent20Sec111a
                    stcgPercent20Sec111a = 0
                End If
                
            setOffRem = 0
    End If

'Commented by Riyaz on 29/01/2026
'7) Long term capital gain taxable @ 20%
'Dim ltcg20PercntSetOffRem As Double
'Dim prevLtcgSetOffRem As Double
'
'        prevSetOffRem = setOffRem
'      ltcg20PercntSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2a").value - Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").value - Sheet16.Range("ltcg.HPlossCurYrSetoff2a").value)
'        If (ltcg20PercntSetOffRem >= 0 And setOffRem >= 0) Then
'            If (ltcg20PercntSetOffRem >= setOffRem) Then
'                tempAk = setOffRem
'                setOffRem = setOffRem - ltcg20PercntSetOffRem
'
'                If (setOffRem >= 0) Then
'                        Sheet16.Range("ltcg.BusLossSetoff2a").value = ltcg20PercntSetOffRem
'                        ltcgPrctg20Sec112 = 0
'                        ltcgPTI20 = 0
'                        ltcgPrctg20Sec11EA = 0
'                        ltcgPrctg20SecProviso = 0 ''Newly added by Bindu as per DE V4 on 15th March 25
'                    ltcg20PercntSetOffRem = Application.WorksheetFunction.Max(0, ltcg20PercntSetOffRem - Sheet16.Range("ltcg.BusLossSetoff2a").value)
'                Else
'                        Sheet16.Range("ltcg.BusLossSetoff2a").value = tempAk
'                        setOffRem = 0
'
'                        ' sec break up
'
'                        If (tempAk >= 0 And ltcgPrctg20Sec11EA > tempAk) Then
'                            ltcgPrctg20Sec11EA = ltcgPrctg20Sec11EA - tempAk
'                            tempAk = 0
'                        Else
'                            tempAk = tempAk - ltcgPrctg20Sec11EA
'                            ltcgPrctg20Sec11EA = 0
'                        End If
'
'
'
'                        If (tempAk >= 0 And ltcgPTI20 > tempAk) Then
'                            ltcgPTI20 = ltcgPTI20 - tempAk
'                            tempAk = 0
'                        Else
'                            tempAk = tempAk - ltcgPTI20
'                            ltcgPTI20 = 0
'                        End If
'
'                         ''Newly added by Bindu as per DE V4 on 15th March 25
'                        If (tempAk >= 0 And ltcgPrctg20SecProviso > tempAk) Then
'                            ltcgPrctg20SecProviso = ltcgPrctg20SecProviso - tempAk
'                            tempAk = 0
'                        Else
'                            tempAk = tempAk - ltcgPrctg20SecProviso
'                            ltcgPrctg20SecProviso = 0
'                        End If
'                        '--
'
'                        If (tempAk >= 0 And ltcgPrctg20Sec112 > tempAk) Then
'                            ltcgPrctg20Sec112 = ltcgPrctg20Sec112 - tempAk
'                            tempAk = 0
'                        Else
'                            tempAk = tempAk - ltcgPrctg20Sec112
'                            ltcgPrctg20Sec112 = 0
'                        End If
'
'                    ltcg20PercntSetOffRem = Application.WorksheetFunction.Max(0, ltcg20PercntSetOffRem - Sheet16.Range("ltcg.BusLossSetoff2a").value)
'            End If
'        Else
'                 prevLtcgSetOffRem = ltcg20PercntSetOffRem
'
'                If (ltcg20PercntSetOffRem >= 0) Then
'                        Sheet16.Range("ltcg.BusLossSetoff2a").value = ltcg20PercntSetOffRem
'                        ltcg20PercntSetOffRem = 0
'                        ltcgPrctg20Sec112 = 0
'                        ltcgPTI20 = 0
'                        ltcgPrctg20Sec11EA = 0
'                        ltcgPrctg20SecProviso = 0 'Newly added by Bindu as per DE V4 on 15th March 25
'                    setOffRem = Application.WorksheetFunction.Max(0, setOffRem - Sheet16.Range("ltcg.BusLossSetoff2a").value)
'                Else
'                        Sheet16.Range("ltcg.BusLossSetoff2a").value = prevLtcgSetOffRem
'                        ltcg20PercntSetOffRem = 0
'                        ltcgPrctg20Sec112 = 0
'                        ltcgPTI20 = 0
'                        ltcgPrctg20Sec11EA = 0
'                        'Newly added by Bindu as per DE V4 on 15th March 25
'                        ltcgPrctg20SecProviso = 0
'                    setOffRem = Application.WorksheetFunction.Max(0, setOffRem - Sheet16.Range("ltcg.BusLossSetoff2a").value)
'                End If
'            End If
'        End If
        
'Commented by Riyaz on 29/01/2026
' 8) Short-term capital gain taxable @ 15%
'Dim stcg15PercntSetOffRem As Double
'stcg15PercntSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("stcg.IncOfCurYrUnderThatHead1").value - Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1").value - Sheet16.Range("stcg.HPlossCurYrSetoff1").value)
'
'        prevSetOffRem = setOffRem
'        setOffRem = setOffRem - stcg15PercntSetOffRem
'        If (setOffRem >= 0) Then
'            If (prevSetOffRem = 0) Then
'                 Sheet16.Range("stcg.BusLossSetoff1").value = 0
'            Else
'                     Sheet16.Range("stcg.BusLossSetoff1").value = stcg15PercntSetOffRem
'                    stcg15PercntSetOffRem = 0
'                    stcgPercent15Sec111a = 0
'                    stcgPTI15 = 0
'                    stcgPercent15Sec115ad_1_b_ii = 0
'
'            End If
'        Else
'                Sheet16.Range("stcg.BusLossSetoff1").value = prevSetOffRem
'
'                ' sec break up
'                If (prevSetOffRem >= 0 And stcgPercent15Sec115ad_1_b_ii > prevSetOffRem) Then
'                    stcgPercent15Sec115ad_1_b_ii = stcgPercent15Sec115ad_1_b_ii - prevSetOffRem
'                    prevSetOffRem = 0
'                Else
'                    prevSetOffRem = prevSetOffRem - stcgPercent15Sec115ad_1_b_ii
'                    stcgPercent15Sec115ad_1_b_ii = 0
'                End If
'
'                If (prevSetOffRem >= 0 And stcgPTI15 > prevSetOffRem) Then
'                    stcgPTI15 = stcgPTI15 - prevSetOffRem
'                    prevSetOffRem = 0
'                Else
'                    prevSetOffRem = prevSetOffRem - stcgPTI15
'                    stcgPTI15 = 0
'                End If
'
'
'                If (prevSetOffRem >= 0 And stcgPercent15Sec111a > prevSetOffRem) Then
'                    stcgPercent15Sec111a = stcgPercent15Sec111a - prevSetOffRem
'                    prevSetOffRem = 0
'                Else
'                    prevSetOffRem = prevSetOffRem - stcgPercent15Sec111a
'                    stcgPercent15Sec111a = 0
'                End If
'
'            setOffRem = 0
'    End If
        
'---LTCG 12.5% NEWLY ADDED BY BINDU 26TH FEB 25


Dim prevLtcgSetOffRem As Double
Dim ltcg125PercntSetOffRem As Double

        prevSetOffRem = setOffRem
     ltcg125PercntSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2b").value - Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2b").value - Sheet16.Range("ltcg.HPlossCurYrSetoff2b").value)
     
        If (ltcg125PercntSetOffRem >= 0 And setOffRem >= 0) Then
            If (ltcg125PercntSetOffRem >= setOffRem) Then
                tempAk = setOffRem
                setOffRem = setOffRem - ltcg125PercntSetOffRem
                
                If (setOffRem >= 0) Then
                        Sheet16.Range("ltcg.BusLossSetoff2b").value = ltcg125PercntSetOffRem
                       ltcgPrctg125Sec115AC1 = 0
                        ltcgPrctg125Sec115Eb = 0
                        ltcgPrctg125Sec1121c2 = 0
                        ltcgPrctg125Sec115ACA1 = 0
                        ltcgPrctg125Sec115Ea = 0
                        ltcgPrctg125Sec115AD3_Proviso = 0
                        ltcgPrctg125Sec112A = 0
                        ltcgPTI125 = 0
                        ltcgPTI125_112 = 0
                        ltcgPTI112Sec48 = 0
                        ltcgPTI125112B9 = 0
                        ltcgPrctg125Sec115AD3 = 0
                        ltcgPrctg125SecProviso = 0 '15th March
                    ltcg125PercntSetOffRem = Application.WorksheetFunction.Max(0, ltcg125PercntSetOffRem - Sheet16.Range("ltcg.BusLossSetoff2b").value)
                Else
                        Sheet16.Range("ltcg.BusLossSetoff2b").value = tempAk
                        setOffRem = 0
                        
                       'Section break up

                      '-1st Not using
                        If (tempAk >= 0 And ltcgPrctg125Sec115Ea > tempAk) Then
                                ltcgPrctg125Sec115Ea = ltcgPrctg125Sec115Ea - tempAk
                                tempAk = 0
                       Else
                                tempAk = tempAk - ltcgPrctg125Sec115Ea
                                ltcgPrctg125Sec115Ea = 0
                       End If
                       '--2nd not using
                        If (tempAk >= 0 And ltcgPTI112Sec48 > tempAk) Then
                                ltcgPTI112Sec48 = ltcgPTI112Sec48 - tempAk
                                tempAk = 0
                        Else
                              tempAk = tempAk - ltcgPTI112Sec48
                                ltcgPTI112Sec48 = 0
                        End If
                        '--10th
                         If (tempAk >= 0 And ltcgPTI125112B9 > tempAk) Then
                                ltcgPTI125112B9 = ltcgPTI125112B9 - tempAk
                                tempAk = 0
                        Else
                              tempAk = tempAk - ltcgPTI125112B9
                                ltcgPTI125112B9 = 0
                        End If
                        '--9th
                        If (tempAk >= 0 And ltcgPrctg125Sec1121c2 > tempAk) Then
                                ltcgPrctg125Sec1121c2 = ltcgPrctg125Sec1121c2 - tempAk
                                tempAk = 0
                        Else
                               tempAk = tempAk - ltcgPrctg125Sec1121c2
                                ltcgPrctg125Sec1121c2 = 0
                        End If
                        '--8th
                        If (tempAk >= 0 And ltcgPrctg125Sec115AC1 > tempAk) Then
                                ltcgPrctg125Sec115AC1 = ltcgPrctg125Sec115AC1 - tempAk
                                tempAk = 0
                        Else
                                tempAk = tempAk - ltcgPrctg125Sec115AC1
                                ltcgPrctg125Sec115AC1 = 0
                        End If
                        '7th
                        If (tempAk >= 0 And ltcgPrctg125Sec115ACA1 > tempAk) Then
                                ltcgPrctg125Sec115ACA1 = ltcgPrctg125Sec115ACA1 - tempAk
                                tempAk = 0
                       Else
                              tempAk = tempAk - ltcgPrctg125Sec115ACA1
                                ltcgPrctg125Sec115ACA1 = 0
                       End If
'Added by Riyaz on 03/03/2026
                       If (tempAk >= 0 And ltcgPrctg125Sec115AD3 > tempAk) Then
                                ltcgPrctg125Sec115AD3 = ltcgPrctg125Sec115AD3 - tempAk
                                tempAk = 0
                       Else
                              tempAk = tempAk - ltcgPrctg125Sec115AD3
                                ltcgPrctg125Sec115AD3 = 0
                       End If
                       
                       '6th
                       If (tempAk >= 0 And ltcgPrctg125Sec115AD3_Proviso > tempAk) Then
                                ltcgPrctg125Sec115AD3_Proviso = ltcgPrctg125Sec115AD3_Proviso - tempAk
                                tempAk = 0
                       Else
                                tempAk = tempAk - ltcgPrctg125Sec115AD3_Proviso
                                ltcgPrctg125Sec115AD3_Proviso = 0
                       End If
                       '5th
                       If (tempAk >= 0 And ltcgPrctg125Sec115Eb > tempAk) Then
                                ltcgPrctg125Sec115Eb = ltcgPrctg125Sec115Eb - tempAk
                                tempAk = 0
                       Else
                                tempAk = tempAk - ltcgPrctg125Sec115Eb
                                ltcgPrctg125Sec115Eb = 0
                       End If
                       '4th
                       If (tempAk >= 0 And ltcgPTI125 > tempAk) Then
                                ltcgPTI125 = ltcgPTI125 - tempAk
                                tempAk = 0
                       Else
                                tempAk = tempAk - ltcgPTI125
                                ltcgPTI125 = 0
                       End If
                       '--3rd
                         '-
                        If (tempAk >= 0 And ltcgPrctg125SecProviso > tempAk) Then
                                ltcgPrctg125SecProviso = ltcgPrctg125SecProviso - tempAk
                                tempAk = 0
                        Else
                              tempAk = tempAk - ltcgPrctg125SecProviso
                                ltcgPrctg125SecProviso = 0
                        End If
                        '--
                        '--2nd
                         If (tempAk >= 0 And ltcgPTI125_112 > tempAk) Then
                                ltcgPTI125_112 = ltcgPTI125_112 - tempAk
                                tempAk = 0
                       Else
                                tempAk = tempAk - ltcgPTI125_112
                                ltcgPTI125_112 = 0
                       End If
                       '1st
                       If (tempAk >= 0 And ltcgPrctg125Sec112A > tempAk) Then
                                ltcgPrctg125Sec112A = ltcgPrctg125Sec112A - tempAk
                                tempAk = 0
                        Else
                              tempAk = tempAk - ltcgPrctg125Sec112A
                                ltcgPrctg125Sec112A = 0
                        End If

                    ltcg125PercntSetOffRem = Application.WorksheetFunction.Max(0, ltcg125PercntSetOffRem - Sheet16.Range("ltcg.BusLossSetoff2b").value)
                End If
            Else
                 prevLtcgSetOffRem = ltcg125PercntSetOffRem
                
                If (ltcg125PercntSetOffRem >= 0) Then
                        Sheet16.Range("ltcg.BusLossSetoff2b").value = ltcg125PercntSetOffRem
                        ltcgPrctg125Sec115AC1 = 0
                        ltcgPrctg125Sec115Eb = 0
                        ltcgPrctg125Sec1121c2 = 0
                        ltcgPrctg125Sec115ACA1 = 0
                        ltcgPrctg125Sec115Ea = 0
                        ltcgPrctg125Sec115AD3_Proviso = 0
                        ltcgPrctg125Sec112A = 0
                        ltcgPTI125 = 0
                        ltcgPTI125_112 = 0
                        ltcgPTI112Sec48 = 0
                        ltcgPTI125112B9 = 0
                        ltcgPrctg125Sec115AD3 = 0
                        ltcgPrctg125SecProviso = 0 '15th March
                    setOffRem = Application.WorksheetFunction.Max(0, setOffRem - Sheet16.Range("ltcg.BusLossSetoff2b").value)
               End If
            End If
       End If
       
  '---END
        
'Commented by Riyaz on 29/01/2026
'9) Long term capital gain taxable @ 10%

'Dim prevLtcgSetOffRem As Double
'Dim ltcg10PercntSetOffRem As Double
'
'        prevSetOffRem = setOffRem
'     ltcg10PercntSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2").value - Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2").value - Sheet16.Range("ltcg.HPlossCurYrSetoff2").value)
'
'        If (ltcg10PercntSetOffRem >= 0 And setOffRem >= 0) Then
'            If (ltcg10PercntSetOffRem >= setOffRem) Then
'                tempAk = setOffRem
'                setOffRem = setOffRem - ltcg10PercntSetOffRem
'
'                If (setOffRem >= 0) Then
'                        Sheet16.Range("ltcg.BusLossSetoff2").value = ltcg10PercntSetOffRem
'                       ltcgPrctg10Sec112c2 = 0
'                        ltcgPrctg10Sec112A = 0
'                        ltcgPrctg10Sec115AC1 = 0
'                        ltcgPrctg10Sec115ACA1 = 0
'                        ltcgPrctg10Sec115AD3 = 0
'                        ltcgPrctg10Sec115AD3_Proviso = 0
'                        ltcgPrctg10Sec115Eb = 0
'                        'ltcgPrctg10SecProviso = 0 ''NOt using  commented  by Bindu as this sec changed from 10 to 20
'                        ltcgPTI10 = 0
'                        ltcgPTI10_112 = 0
'
'                    ltcg10PercntSetOffRem = Application.WorksheetFunction.Max(0, ltcg10PercntSetOffRem - Sheet16.Range("ltcg.BusLossSetoff2").value)
'                Else
'                        Sheet16.Range("ltcg.BusLossSetoff2").value = tempAk
'                        setOffRem = 0
'
'                       'Section break up
'
'                        If (tempAk >= 0 And ltcgPrctg10Sec112c2 > tempAk) Then
'                            ltcgPrctg10Sec112c2 = ltcgPrctg10Sec112c2 - tempAk
'                            tempAk = 0
'                        Else
'                            tempAk = tempAk - ltcgPrctg10Sec112c2
'                            ltcgPrctg10Sec112c2 = 0
'                        End If
'
'                        If (tempAk >= 0 And ltcgPrctg10Sec115AC1 > tempAk) Then
'                            ltcgPrctg10Sec115AC1 = ltcgPrctg10Sec115AC1 - tempAk
'                            tempAk = 0
'                       Else
'                            tempAk = tempAk - ltcgPrctg10Sec115AC1
'                            ltcgPrctg10Sec115AC1 = 0
'                        End If
'
'                        If (tempAk >= 0 And ltcgPrctg10Sec115ACA1 > tempAk) Then
'                            ltcgPrctg10Sec115ACA1 = ltcgPrctg10Sec115ACA1 - tempAk
'                            tempAk = 0
'                       Else
'                            tempAk = tempAk - ltcgPrctg10Sec115ACA1
'                            ltcgPrctg10Sec115ACA1 = 0
'                        End If
'
'                        If (tempAk >= 0 And ltcgPrctg10Sec115AD3 > tempAk) Then
'                            ltcgPrctg10Sec115AD3 = ltcgPrctg10Sec115AD3 - tempAk
'                            tempAk = 0
'                      Else
'                            tempAk = tempAk - ltcgPrctg10Sec115AD3
'                            ltcgPrctg10Sec115AD3 = 0
'                       End If
'
'                       If (tempAk >= 0 And ltcgPrctg10Sec115AD3_Proviso > tempAk) Then
'                            ltcgPrctg10Sec115AD3_Proviso = ltcgPrctg10Sec115AD3_Proviso - tempAk
'                            tempAk = 0
'                      Else
'                            tempAk = tempAk - ltcgPrctg10Sec115AD3_Proviso
'                            ltcgPrctg10Sec115AD3_Proviso = 0
'                       End If
'
'                        If (tempAk >= 0 And ltcgPrctg10Sec115Eb > tempAk) Then
'                            ltcgPrctg10Sec115Eb = ltcgPrctg10Sec115Eb - tempAk
'                            tempAk = 0
'                       Else
'                            tempAk = tempAk - ltcgPrctg10Sec115Eb
'                            ltcgPrctg10Sec115Eb = 0
'                        End If
'
'
'                        If (tempAk >= 0 And ltcgPTI10 > tempAk) Then
'                            ltcgPTI10 = ltcgPTI10 - tempAk
'                            tempAk = 0
'                       Else
'                            tempAk = tempAk - ltcgPTI10
'                            ltcgPTI10 = 0
'                        End If
'
'                        'NOt using  commented  by Bindu as this sec changed from 10 to 20
''                        If (tempAk >= 0 And ltcgPrctg10SecProviso > tempAk) Then
''                            ltcgPrctg10SecProviso = ltcgPrctg10SecProviso - tempAk
''                            tempAk = 0
''                       Else
''                            tempAk = tempAk - ltcgPrctg10SecProviso
''                            ltcgPrctg10SecProviso = 0
''                        End If
'
'                        If (tempAk >= 0 And ltcgPTI10_112 > tempAk) Then
'                            ltcgPTI10_112 = ltcgPTI10_112 - tempAk
'                            tempAk = 0
'                       Else
'                            tempAk = tempAk - ltcgPTI10_112
'                            ltcgPTI10_112 = 0
'                        End If
'
'                        If (tempAk >= 0 And ltcgPrctg10Sec112A > tempAk) Then
'                            ltcgPrctg10Sec112A = ltcgPrctg10Sec112A - tempAk
'                            tempAk = 0
'                        Else
'                            tempAk = tempAk - ltcgPrctg10Sec112A
'                            ltcgPrctg10Sec112A = 0
'                        End If
'
'                    ltcg10PercntSetOffRem = Application.WorksheetFunction.Max(0, ltcg10PercntSetOffRem - Sheet16.Range("ltcg.BusLossSetoff2").value)
'                End If
'            Else
'                 prevLtcgSetOffRem = ltcg10PercntSetOffRem
'
'                If (ltcg10PercntSetOffRem >= 0) Then
'                        Sheet16.Range("ltcg.BusLossSetoff2").value = ltcg10PercntSetOffRem
'                        ltcg10PercntSetOffRem = 0
'                        ltcgPrctg10Sec112c2 = 0
'                        ltcgPrctg10Sec115AC1 = 0
'                        ltcgPrctg10Sec115ACA1 = 0
'                        ltcgPrctg10Sec115AD3 = 0
'                        ltcgPrctg10Sec115AD3_Proviso = 0
'                        ltcgPrctg10Sec112A = 0
'                        ltcgPrctg10Sec115Eb = 0
'                        'ltcgPrctg10SecProviso = 0 'NOt using  commented  by Bindu as this sec changed from 10 to 20
'                        ltcgPTI10 = 0
'                        ltcgPTI10_112 = 0
'                    setOffRem = Application.WorksheetFunction.Max(0, setOffRem - Sheet16.Range("ltcg.BusLossSetoff2").value)
'               End If
'            End If
'       End If
'       'STCG DTAA
'Dim stcgDTAARateSetOffRem As Double
'
'   stcgDTAARateSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("stcg.IncOfCurYrUnderThatHead_DTAARates").value - Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff_DTAARates").value - Sheet16.Range("stcg.HPlossCurYrSetoff_DTAARates").value)
'
'        prevSetOffRem = setOffRem
'        setOffRem = setOffRem - stcgDTAARateSetOffRem
'        If (setOffRem >= 0) Then
'            If (prevSetOffRem = 0) Then
'               Sheet16.Range("stcg.BusLossSetoff_DTAARates").value = 0
'            Else
'                Sheet16.Range("stcg.BusLossSetoff_DTAARates").value = stcgDTAARateSetOffRem
'            End If
'        Else
'               Sheet16.Range("stcg.BusLossSetoff_DTAARates").value = prevSetOffRem
'                setOffRem = 0
'        End If
        
'LTCG DTAA
Dim ltcgDTAARateSetOffRem As Double
        
   ltcgDTAARateSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("ltcg.IncOfCurYrUnderThatHead_DTAARates").value - Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff_DTAARates").value - Sheet16.Range("ltcg.HPlossCurYrSetoff_DTAARates").value)
        
        prevSetOffRem = setOffRem
        setOffRem = setOffRem - ltcgDTAARateSetOffRem
        If (setOffRem >= 0) Then
            If (prevSetOffRem = 0) Then
               Sheet16.Range("ltcg.BusLossSetoff_DTAARates").value = 0
            Else
                Sheet16.Range("ltcg.BusLossSetoff_DTAARates").value = ltcgDTAARateSetOffRem
            End If
        Else
               Sheet16.Range("ltcg.BusLossSetoff_DTAARates").value = prevSetOffRem
                setOffRem = 0
        End If
        
'Commented by Riyaz on 29/01/2026
'        cgosIncome_cgInc_stcg_prctg15_sec111a = stcgPercent15Sec111a
'        cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = stcgPercent15Sec115ad_1_b_ii
        '--STCG 20% NEWLY ADDED BY BINDU
        cgosIncome_cgInc_stcg_prctg20_sec111a = stcgPercent20Sec111a
        cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii = stcgPercent20Sec115ad_1_b_ii
        
'Commented by Riyaz on 29/01/2026
'        cgosIncome_cgInc_ltcg_prctg20_sec112 = ltcgPrctg20Sec112
'        cgosIncome_cgInc_ltcg_prctg20_sec11EA = ltcgPrctg20Sec11EA
        ''Newly added by Bindu as per DE V4 on 15th March 25
'        cgosIncome_cgInc_ltcg_prctg20_secProviso = ltcgPrctg20SecProviso
        
'Commented by Riyaz on 29/01/2026
        'cgosIncome_cgInc_ltcg_prctg10_secProviso = ltcgPrctg10SecProviso ''NOt using  commented  by Bindu as this sec changed from 10 to 20
'        cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = ltcgPrctg10Sec112c2
'        cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 = ltcgPrctg10Sec115AC1
'        cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 = ltcgPrctg10Sec115ACA1
'        cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 = ltcgPrctg10Sec115AD3
'        cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso = ltcgPrctg10Sec115AD3_Proviso
'        cgosIncome_cgInc_ltcg_prctg10_sec112A = ltcgPrctg10Sec112A
'       cgosIncome_cgInc_ltcg_prctg10_sec115E_b = ltcgPrctg10Sec115Eb
'
'
'       cgosIncome_cgInc_ltcg_PTI10 = ltcgPTI10
'       cgosIncome_cgInc_ltcg_PTI10_112 = ltcgPTI10_112
        cgosIncome_cgInc_stcg_PTI30 = stcgPTI30
        cgosIncome_cgInc_stcg_prctg30_sec115AD = stcgPrctg30Sec115AD
'Commented by Riyaz on 29/01/2026
'        cgosIncome_cgInc_stcg_PTI15 = stcgPTI15
        '----LTCG 12.5% NEWLY ADDED BY BINDU 26TH FEB 25
        cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 = ltcgPrctg125Sec115ACA1
        cgosIncome_cgInc_ltcg_prctg125_sec115E_b = ltcgPrctg125Sec115Eb
        cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 = ltcgPrctg125Sec1121c2
        cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 = ltcgPrctg125Sec115AC1
        cgosIncome_cgInc_ltcg_prctg125_sec115E_a = ltcgPrctg125Sec115Ea
        cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso = ltcgPrctg125Sec115AD3_Proviso
        cgosIncome_cgInc_ltcg_prctg125_sec112A = ltcgPrctg125Sec112A
        cgosIncome_cgInc_ltcg_PTI125 = ltcgPTI125
        cgosIncome_cgInc_ltcg_PTI125_112 = ltcgPTI125_112
        cgosIncome_cgInc_ltcg_PTI112_sec48 = ltcgPTI112Sec48
        cgosIncome_cgInc_ltcg_PTI125_112_B9 = ltcgPTI125112B9
        cgosIncome_cgInc_ltcg_prctg125_secProviso = ltcgPrctg125SecProviso '15th March 25
        cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 = ltcgPrctg125Sec115AD3
    '--END
        
        
        
        '---STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
        cgosIncome_cgInc_stcg_PTI20 = stcgPTI20
'Commented by Riyaz on 29/01/2026
'        cgosIncome_cgInc_ltcg_PTI20 = ltcgPTI20
       
          'storing diff section remaining value
          '--LTCG 20%
'Commented by Riyaz on 29/01/2026
'    Sheet31.Range("CYLAtemp20Sec112").value = cgosIncome_cgInc_ltcg_prctg20_sec112
'    Sheet31.Range("CYLAtemp20Sec11EA").value = cgosIncome_cgInc_ltcg_prctg20_sec11EA
    'Newly added by Bindu as per DE V4 on 15th March 25
'    Sheet31.Range("CYLAtemp20SecPro").value = cgosIncome_cgInc_ltcg_prctg20_secProviso
   
'Commented by Riyaz on 29/01/2026
    '---STCG 15%
'    Sheet31.Range("CYLAtemp15Sec111a").value = cgosIncome_cgInc_stcg_prctg15_sec111a
'    Sheet31.Range("CYLAtemp15Sec115ad1bii").value = cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii
    
     '---STCG 20% NEWLY ADDED BY BINDU
     Sheet31.Range("CYLAtemp20Sec111a").value = cgosIncome_cgInc_stcg_prctg20_sec111a
    Sheet31.Range("CYLAtemp20Sec115ad1bii").value = cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii
    '--END
'Commented by Riyaz on 29/01/2026
        'Sheet31.Range("CYLAtemp10SecPro").value = cgosIncome_cgInc_ltcg_prctg10_secProviso ''NOt using  commented  by Bindu as this sec changed from 10 to 20
'        Sheet31.Range("CYLAtemp10Sec115ACA1").value = cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1
'        Sheet31.Range("CYLAtemp10Sec115Eb").value = cgosIncome_cgInc_ltcg_prctg10_sec115E_b
'        Sheet31.Range("CYLAtemp10Sec1121c2").value = cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2
'        Sheet31.Range("CYLAtemp10Sec115AC1").value = cgosIncome_cgInc_ltcg_prctg10_sec115AC_1
'        Sheet31.Range("CYLAtemp10Sec115AD3").value = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3
'        Sheet31.Range("CYLAtemp10Sec115AD3_Proviso").value = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso
'        Sheet31.Range("CYLAtemp10Sec112A").value = cgosIncome_cgInc_ltcg_prctg10_sec112A
'
'        Sheet31.Range("CYLAtempPTI10").value = cgosIncome_cgInc_ltcg_PTI10
'        Sheet31.Range("CYLAtempPTI10_112").value = cgosIncome_cgInc_ltcg_PTI10_112
        
        '--LTCG 12.5% NEWLY ADDED BY BINDU 26TH FEB 25
        
        Sheet31.Range("CYLAtemp125Sec112_c").value = cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2
        Sheet31.Range("CYLAtemp125Sec_115ACA").value = cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1
        Sheet31.Range("CYLAtemp125Sec_115EB").value = cgosIncome_cgInc_ltcg_prctg125_sec115E_b
        Sheet31.Range("CYLAtemp125Sec_115EA").value = cgosIncome_cgInc_ltcg_prctg125_sec115E_a
        Sheet31.Range("CYLAtemp125Sec_115AC").value = cgosIncome_cgInc_ltcg_prctg125_sec115AC_1
        Sheet31.Range("CYLAtemp125Sec_115AD_iii").value = cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso
        Sheet31.Range("CYLAtemp125Sec_PTI").value = cgosIncome_cgInc_ltcg_PTI125
        Sheet31.Range("CYLAtemp125Sec_112_PTI").value = cgosIncome_cgInc_ltcg_PTI125_112
        Sheet31.Range("CYLAtemp125Sec_112").value = cgosIncome_cgInc_ltcg_PTI125_112_B9
        Sheet31.Range("CYLAtemp125Sec_112A").value = cgosIncome_cgInc_ltcg_prctg125_sec112A
        Sheet31.Range("CYLAtemp125Sec_112sec_48").value = cgosIncome_cgInc_ltcg_PTI112_sec48
        Sheet31.Range("CYLAtemp125SecPro").value = cgosIncome_cgInc_ltcg_prctg125_secProviso '15th March
        Sheet31.Range("CYLAtemp115AD125").value = cgosIncome_cgInc_ltcg_prctg125_sec115AD_3
        
        
        '---STCG 30%
        Sheet31.Range("CYLAtempPTI30").value = cgosIncome_cgInc_stcg_PTI30
        Sheet31.Range("CYLAtemp115AD30").value = cgosIncome_cgInc_stcg_prctg30_sec115AD
        Sheet31.Range("CYLAtempPTI15").value = cgosIncome_cgInc_stcg_PTI15
     
     '--STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
        Sheet31.Range("CYLAtempPTI20_STCG").value = cgosIncome_cgInc_stcg_PTI20
        '--LTCG 20%
'Commented by Riyaz on 29/01/2026
'        Sheet31.Range("CYLAtempPTI20").value = cgosIncome_cgInc_ltcg_PTI20
        
Application.EnableEvents = True

End Function

Function setOffHPLossCYLA()

Application.EnableEvents = False

Dim stcgPercent15Sec111a As Double
Dim stcgPercent15Sec115ad_1_b_ii As Double
'--STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
Dim stcgPercent20Sec111a As Double
Dim stcgPercent20Sec115ad_1_b_ii As Double


Dim ltcgPrctg20Sec112 As Double
Dim ltcgPrctg10Sec112A As Double
Dim ltcgPrctg20Sec11EA As Double
'Newly added by Bindu as per DE V4 on 15th March 25
Dim ltcgPrctg20SecProviso As Double



Dim ltcgPrctg10SecProviso As Double
Dim ltcgPrctg10Sec112c2 As Double
Dim ltcgPrctg10Sec115AC1 As Double
Dim ltcgPrctg10Sec115ACA1 As Double
Dim ltcgPrctg10Sec115AD3 As Double
Dim ltcgPrctg10Sec115AD3_Proviso As Double
Dim ltcgPrctg10Sec115Eb As Double

'LTCG 12.5% NEWLY ADDED BY BINDU ON 25TH FEN 25
Dim ltcgPrctg125Sec115ACA1 As Double
Dim ltcgPrctg125Sec115Eb As Double
Dim ltcgPrctg125Sec1121c2 As Double
Dim ltcgPrctg125Sec115AC1 As Double
Dim ltcgPrctg125Sec115Ea As Double
Dim ltcgPrctg125Sec115AD3_Proviso As Double
Dim ltcgPrctg125Sec112A As Double
Dim ltcgPTI112Sec48 As Double
Dim ltcgPTI125112B9 As Double
Dim ltcgPrctg125SecProviso As Double '15th March
Dim ltcgPrctg125Sec115AD3 As Double


Dim stcgPrctg30Sec115AD As Double
Dim stcgPTI30 As Double
Dim stcgPTI15 As Double
'--STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
Dim stcgPTI20 As Double

Dim ltcgPTI10 As Double
Dim ltcgPTI10_112 As Double
Dim ltcgPTI20 As Double


'--LTCG 12.5% NEWLY ADDED BY BINDU ON 25TH FEB 25
Dim ltcgPTI125  As Double
Dim ltcgPTI125_112  As Double


'Commented by Riyaz on 29/01/2026
'stcgPercent15Sec111a = cgosIncome_cgInc_stcg_prctg15_sec111a
'stcgPercent15Sec115ad_1_b_ii = cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii
'--STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
stcgPercent20Sec111a = cgosIncome_cgInc_stcg_prctg20_sec111a
stcgPercent20Sec115ad_1_b_ii = cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii

'Commented by Riyaz on 29/01/2026
'ltcgPrctg20Sec112 = cgosIncome_cgInc_ltcg_prctg20_sec112
'ltcgPrctg10Sec112A = cgosIncome_cgInc_ltcg_prctg10_sec112A
'ltcgPrctg20Sec11EA = cgosIncome_cgInc_ltcg_prctg20_sec11EA
''Newly added by Bindu as per DE V4 on 15th March 25
'ltcgPrctg20SecProviso = cgosIncome_cgInc_ltcg_prctg20_secProviso

'Commented by Riyaz on 29/01/2026
'ltcgPrctg10SecProviso = cgosIncome_cgInc_ltcg_prctg10_secProviso 'NOt using  commented  by Bindu as this sec changed from 10 to 20
'ltcgPrctg10Sec112c2 = cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2
'ltcgPrctg10Sec115AC1 = cgosIncome_cgInc_ltcg_prctg10_sec115AC_1
'ltcgPrctg10Sec115ACA1 = cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1
'ltcgPrctg10Sec115AD3 = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3
'ltcgPrctg10Sec115AD3_Proviso = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso
'ltcgPrctg10Sec115Eb = cgosIncome_cgInc_ltcg_prctg10_sec115E_b

'---LTCG 12.5% NEWLY ADDED BY BINDU 26TH FEB 25
 ltcgPrctg125Sec115ACA1 = cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1
 ltcgPrctg125Sec115Eb = cgosIncome_cgInc_ltcg_prctg125_sec115E_b
 ltcgPrctg125Sec1121c2 = cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2
 ltcgPrctg125Sec115AC1 = cgosIncome_cgInc_ltcg_prctg125_sec115AC_1
 ltcgPrctg125Sec115Ea = cgosIncome_cgInc_ltcg_prctg125_sec115E_a
 ltcgPrctg125Sec115AD3_Proviso = cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso
 ltcgPrctg125Sec112A = cgosIncome_cgInc_ltcg_prctg125_sec112A
 ltcgPTI125 = cgosIncome_cgInc_ltcg_PTI125
 ltcgPTI125_112 = cgosIncome_cgInc_ltcg_PTI125_112
 ltcgPTI112Sec48 = cgosIncome_cgInc_ltcg_PTI112_sec48
 ltcgPTI125112B9 = cgosIncome_cgInc_ltcg_PTI125_112_B9
 ltcgPrctg125SecProviso = cgosIncome_cgInc_ltcg_prctg125_secProviso '15th March
 ltcgPrctg125Sec115AD3 = cgosIncome_cgInc_ltcg_prctg125_sec115AD_3

 stcgPrctg30Sec115AD = cgosIncome_cgInc_stcg_prctg30_sec115AD
 stcgPTI30 = cgosIncome_cgInc_stcg_PTI30
'Commented by Riyaz on 29/01/2026
' stcgPTI15 = cgosIncome_cgInc_stcg_PTI15
 '---STCG 20% NEWLY ADDED BY BINDU
 stcgPTI20 = cgosIncome_cgInc_stcg_PTI20
 
' ltcgPTI10 = cgosIncome_cgInc_ltcg_PTI10
' ltcgPTI10_112 = cgosIncome_cgInc_ltcg_PTI10_112
' ltcgPTI20 = cgosIncome_cgInc_ltcg_PTI20

Dim salarySetOffRem As Double
Dim prevSetOffRem As Double
Dim setOffRem As Double
Dim tempAk As Double


'1. salary Income

'added by Chandru
 If Sheet1.Range("Radio_Button_115").value = 2 Then

 Sheet16.Unprotect Password:=getmsgstate

        Sheet16.Range("Salary.HPlossCurYrSetoff1").Interior.Color = (&HD8D8D8)
        Sheet16.Range("Salary.HPlossCurYrSetoff1").Locked = True
        Sheet16.Range("Salary.HPlossCurYrSetoff1").value = 0
    
    Else

                If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
                Sheet16.Unprotect Password:=getmsgstate
                Sheet16.Range("salary.HPlossCurYrSetoff1").Locked = False
                Sheet16.Range("salary.HPlossCurYrSetoff1").Interior.Color = RGB(255, 255, 204)
                
                Else
                
                Sheet16.Range("salary.HPlossCurYrSetoff1").Locked = True
                Sheet16.Range("salary.HPlossCurYrSetoff1").Interior.ColorIndex = 2
' End If
 'end Change Chandru
 
  salarySetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("salary.IncOfCurYrUnderThatHead1").value - Sheet16.Range("salary.OthSrcLossNoRaceHorseSetoff1").value)
'MsgBox salarySetOffRem
  prevSetOffRem = Sheet16.Range("cyla.TotHPlossCurYr1").value
'MsgBox prevSetOffRem
  setOffRem = prevSetOffRem - salarySetOffRem
'MsgBox setOffRem
  
      If (setOffRem >= 0) Then
      
        If (Sheet16.Range("cyla.TotHPlossCurYr1").value = "0") Then
        
            Sheet16.Range("salary.HPlossCurYrSetoff1").value = "0"
        Else
        
         Sheet16.Range("salary.HPlossCurYrSetoff1").value = salarySetOffRem
        End If
    Else

    Sheet16.Range("salary.HPlossCurYrSetoff1").value = Sheet16.Range("cyla.TotHPlossCurYr1").value
    setOffRem = 0
    End If
    End If
End If
'Added by Chandru
    If Sheet16.Range("salary.HPlossCurYrSetoff1").value > Sheet16.Range("SAL.TotIncUnderHeadSalaries1").value Then
    Sheet16.Range("salary.HPlossCurYrSetoff1").value = 0
    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
    End If

    If Sheet16.Range("Sum_Salaries").value > Sheet16.Range("SAL.TotIncUnderHeadSalaries1").value Then
    Sheet16.Range("salary.HPlossCurYrSetoff1", "salary.OthSrcLossNoRaceHorseSetoff1").value = 0
    MsgBox ("The sum of Column No. 2 + 3 should not exceed amount as referred in Column No. 1")
    End If

'End change

        
' 2.  Business Income
Dim busIncSetOffRem As Double
busIncSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("busprof.IncOfCurYrUnderThatHead").value - Sheet16.Range("busprof.OthSrcLossNoRaceHorseSetoff").value)
        prevSetOffRem = setOffRem
        setOffRem = prevSetOffRem - busIncSetOffRem
        
        If (setOffRem >= 0) Then
            If (prevSetOffRem = 0) Then
                Sheet16.Range("busprof.HPlossCurYrSetoff").value = 0
           Else
             Sheet16.Range("busprof.HPlossCurYrSetoff").value = busIncSetOffRem
           End If
      Else
         Sheet16.Range("busprof.HPlossCurYrSetoff").value = prevSetOffRem
         
            setOffRem = 0
    End If
'3 Speculative Income
Dim specIncSetOffRem As Double
specIncSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("busprofspec.IncOfCurYrUnderThatHead0a").value - Sheet16.Range("busprofspec.OthSrcLossNoRaceHorseSetoff0a").value)
    prevSetOffRem = setOffRem
    setOffRem = prevSetOffRem - specIncSetOffRem
    If (setOffRem >= 0) Then
            If (prevSetOffRem = 0) Then
                Sheet16.Range("busprofspec.HPlossCurYrSetoff0a").value = 0
           Else
             Sheet16.Range("busprofspec.HPlossCurYrSetoff0a").value = specIncSetOffRem
           End If
      Else
         Sheet16.Range("busprofspec.HPlossCurYrSetoff0a").value = prevSetOffRem
         
            setOffRem = 0
    End If
    
'4 Specified Business Income
Dim specBusIncSetOffRem As Double
specBusIncSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("busprofspecified.IncOfCurYrUnderThatHead0b").value - Sheet16.Range("busprofspecified.OthSrcLossNoRaceHorseSetoff0b").value)
    prevSetOffRem = setOffRem
    setOffRem = prevSetOffRem - specBusIncSetOffRem
    If (setOffRem >= 0) Then
           If (prevSetOffRem = 0) Then
             Sheet16.Range("busprof0b.HPlossCurYrSetoff0b").value = 0
           Else
             Sheet16.Range("busprof0b.HPlossCurYrSetoff0b").value = specBusIncSetOffRem
           End If
      Else
            Sheet16.Range("busprof0b.HPlossCurYrSetoff0b").value = prevSetOffRem
            setOffRem = 0
    End If
            
'5) Other sources (excluding profit from owning race horses)
  'added by Chandru
 If Sheet1.Range("Radio_Button_115").value = 2 Then

 Sheet16.Unprotect Password:=getmsgstate

        Sheet16.Range("othSecinclnlhrs.HPlossCurYrSetoff3").Interior.Color = (&HD8D8D8)
        Sheet16.Range("othSecinclnlhrs.HPlossCurYrSetoff3").Locked = True
        Sheet16.Range("othSecinclnlhrs.HPlossCurYrSetoff3").value = 0
    
        Else

        If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
        Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("othSecinclnlhrs.HPlossCurYrSetoff3").Locked = False
        Sheet16.Range("othSecinclnlhrs.HPlossCurYrSetoff3").Interior.Color = RGB(255, 255, 204)
        
        Else
        
        Sheet16.Range("othSecinclnlhrs.HPlossCurYrSetoff3").Locked = True
        Sheet16.Range("othSecinclnlhrs.HPlossCurYrSetoff3").Interior.ColorIndex = 2

 'end Change Chandru
 
 prevSetOffRem = setOffRem
setOffRem = setOffRem - Sheet16.Range("othSecinclnlhrs.IncOfCurYrUnderThatHead3").value
        If (setOffRem >= 0) Then
            If (prevSetOffRem = 0) Then
                Sheet16.Range("othSecinclnlhrs.HPlossCurYrSetoff3").value = 0
            Else
                Sheet16.Range("othSecinclnlhrs.HPlossCurYrSetoff3").value = Sheet16.Range("othSecinclnlhrs.IncOfCurYrUnderThatHead3").value
                
            End If
      Else
            Sheet16.Range("othSecinclnlhrs.HPlossCurYrSetoff3").value = prevSetOffRem
            setOffRem = 0
      End If
      End If
      End If
     'Added by Chandru
    If Sheet16.Range("othSecinclnlhrs.HPlossCurYrSetoff3").value > Sheet16.Range("othSecinclnlhrs.IncOfCurYrUnderThatHead3").value Then
    Sheet16.Range("othSecinclnlhrs.HPlossCurYrSetoff3").value = 0
    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
    End If
    If Sheet16.Range("SUM_NetInc").value > Sheet16.Range("othSecinclnlhrs.IncOfCurYrUnderThatHead3").value Then
    Sheet16.Range("othSecinclnlhrs.HPlossCurYrSetoff3", "othSecinclnlhrs.BusLossSetoff3").value = 0
    MsgBox ("The sum of Column No. 2 + 3 should not exceed amount as referred in Column No. 1")
    End If
    'End change
    
                
'6) Profit from owning and maintaining race horse
Dim raceHrsSetOffRem As Double

'added by Chandru
 If Sheet1.Range("Radio_Button_115").value = 2 Then

 Sheet16.Unprotect Password:=getmsgstate

        Sheet16.Range("rh.HPlossCurYrSetoff4").Interior.Color = (&HD8D8D8)
        Sheet16.Range("rh.HPlossCurYrSetoff4").Locked = True
        Sheet16.Range("rh.HPlossCurYrSetoff4").value = 0
    
        Else

        If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
        Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("rh.HPlossCurYrSetoff4").Locked = False
        Sheet16.Range("rh.HPlossCurYrSetoff4").Interior.Color = RGB(255, 255, 204)
        
        Else
        
        Sheet16.Range("rh.HPlossCurYrSetoff4").Locked = True
        Sheet16.Range("rh.HPlossCurYrSetoff4").Interior.ColorIndex = 2

 'end Change Chandru
 

raceHrsSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("rh.IncOfCurYrUnderThatHead4").value - Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4").value)
        
        prevSetOffRem = setOffRem
        setOffRem = setOffRem - raceHrsSetOffRem
        If (setOffRem >= 0) Then
            If (prevSetOffRem = 0) Then
                Sheet16.Range("rh.HPlossCurYrSetoff4").value = 0
            Else
                Sheet16.Range("rh.HPlossCurYrSetoff4").value = raceHrsSetOffRem
            End If
      Else
            Sheet16.Range("rh.HPlossCurYrSetoff4").value = prevSetOffRem
           setOffRem = 0
      End If
      End If
 End If
 
 
    'Added by Chandru
    If Sheet16.Range("rh.HPlossCurYrSetoff4").value > Sheet16.Range("rh.IncOfCurYrUnderThatHead4").value Then
    Sheet16.Range("rh.HPlossCurYrSetoff4").value = 0
    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
    End If
    If Sheet16.Range("SUM_RaceHorse").value > Sheet16.Range("rh.IncOfCurYrUnderThatHead4").value Then
    Sheet16.Range("rh.HPlossCurYrSetoff4", "rh.BusLossSetoff4").value = 0
    Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4").value = 0
    MsgBox ("The sum of Column No. 2 + 3 should not exceed amount as referred in Column No. 1")
    End If
    'End change
    
      
'7  OS DTAA
      Dim OSDTAASetOffRem As Double

OSDTAASetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("rh.IncOfCurYrUnderThatHead4_DTAARates").value - Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4_DTAARates").value)

        prevSetOffRem = setOffRem
        setOffRem = setOffRem - OSDTAASetOffRem
        If (setOffRem >= 0) Then
            If (prevSetOffRem = 0) Then
                Sheet16.Range("rh.HPlossCurYrSetoff4_DTAARates").value = 0
            Else
                Sheet16.Range("rh.HPlossCurYrSetoff4_DTAARates").value = OSDTAASetOffRem
            End If
      Else
            Sheet16.Range("rh.HPlossCurYrSetoff4_DTAARates").value = prevSetOffRem
           setOffRem = 0
      End If
      
        
'8) Short-term capital gain taxable @ 30%
Dim stcg30PercntSetOffRem As Double
        
        'added by Chandru
 If Sheet1.Range("Radio_Button_115").value = 2 Then

 Sheet16.Unprotect Password:=getmsgstate

        Sheet16.Range("stcg.HPlossCurYrSetoff1a").Interior.Color = (&HD8D8D8)
        Sheet16.Range("stcg.HPlossCurYrSetoff1a").Locked = True
        Sheet16.Range("stcg.HPlossCurYrSetoff1a").value = 0
    
        Else

        If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
        Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("stcg.HPlossCurYrSetoff1a").Locked = False
        Sheet16.Range("stcg.HPlossCurYrSetoff1a").Interior.Color = RGB(255, 255, 204)
        
        Else
        
        Sheet16.Range("stcg.HPlossCurYrSetoff1a").Locked = True
        Sheet16.Range("stcg.HPlossCurYrSetoff1a").Interior.ColorIndex = 2

 'end Change Chandru
 
 
'stcg30PercntSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("stcg.IncOfCurYrUnderThatHead1a").value - Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1a").value)
'
'        prevSetOffRem = setOffRem
'        setOffRem = setOffRem - stcg30PercntSetOffRem
'        If (setOffRem >= 0) Then
'            If (prevSetOffRem = 0) Then
'                Sheet16.Range("stcg.HPlossCurYrSetoff1a").value = 0
'            Else
'               Sheet16.Range("stcg.HPlossCurYrSetoff1a").value = stcg30PercntSetOffRem
'
'            End If
'        Else
'            Sheet16.Range("stcg.HPlossCurYrSetoff1a").value = prevSetOffRem
'            setOffRem = 0
'        End If
 '       Dim stcg15PercntSetOffRem As Double

stcg30PercntSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("stcg.IncOfCurYrUnderThatHead1a").value - Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1a").value)
        
        prevSetOffRem = setOffRem
        setOffRem = setOffRem - stcg30PercntSetOffRem
        If (setOffRem >= 0) Then
            If (prevSetOffRem = 0) Then
                Sheet16.Range("stcg.HPlossCurYrSetoff1a").value = 0
            Else
                    Sheet16.Range("stcg.HPlossCurYrSetoff1a").value = stcg30PercntSetOffRem
                    stcg30PercntSetOffRem = 0
                    
                    stcgPrctg30Sec115AD = 0
                    stcgPTI30 = 0
                    
            End If
        Else
             Sheet16.Range("stcg.HPlossCurYrSetoff1a").value = prevSetOffRem
             
                If (prevSetOffRem >= 0 And stcgPrctg30Sec115AD > prevSetOffRem) Then
                    stcgPrctg30Sec115AD = stcgPrctg30Sec115AD - prevSetOffRem
                    prevSetOffRem = 0
                Else
                    prevSetOffRem = prevSetOffRem - stcgPrctg30Sec115AD
                    stcgPrctg30Sec115AD = 0
                End If

                If (prevSetOffRem >= 0 And stcgPTI30 > prevSetOffRem) Then
                    stcgPTI30 = stcgPTI30 - prevSetOffRem
                    prevSetOffRem = 0
                Else
                    prevSetOffRem = prevSetOffRem - stcgPTI30
                    stcgPTI30 = 0
                End If
           
            setOffRem = 0
        End If
        End If
        End If
        
'Added by Chandru
    If Sheet16.Range("stcg.HPlossCurYrSetoff1a").value > Sheet16.Range("stcg.HPlossCurYrSetoff1a").value Then
    Sheet16.Range("stcg.IncOfCurYrUnderThatHead1a").value = 0
    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
    End If
    If Sheet16.Range("SUM_STCG30").value > Sheet16.Range("stcg.IncOfCurYrUnderThatHead1a").value Then
    Sheet16.Range("stcg.HPlossCurYrSetoff1a", "stcg.BusLossSetoff1a").value = 0
    Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1a").value = 0
    MsgBox ("The sum of Column No. 2 + 3 should not exceed amount as referred in Column No. 1")
    End If
    
'End change

        
' 9) Short-term capital gain taxable at applicable rates
Dim stcgAppRateSetOffRem As Double

'added by Chandru
 If Sheet1.Range("Radio_Button_115").value = 2 Then

 Sheet16.Unprotect Password:=getmsgstate

        Sheet16.Range("stcg.HPlossCurYrSetoff1b").Interior.Color = (&HD8D8D8)
        Sheet16.Range("stcg.HPlossCurYrSetoff1b").Locked = True
        Sheet16.Range("stcg.HPlossCurYrSetoff1b").value = 0
    
        Else

        If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
        Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("stcg.HPlossCurYrSetoff1b").Locked = False
        Sheet16.Range("stcg.HPlossCurYrSetoff1b").Interior.Color = RGB(255, 255, 204)
        
        Else
        
        Sheet16.Range("stcg.HPlossCurYrSetoff1b").Locked = True
        Sheet16.Range("stcg.HPlossCurYrSetoff1b").Interior.ColorIndex = 2

 'end Change Chandru
 
        
stcgAppRateSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("stcg.IncOfCurYrUnderThatHead1b").value - Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1b").value)
        prevSetOffRem = setOffRem
        setOffRem = setOffRem - stcgAppRateSetOffRem
        If (setOffRem >= 0) Then
            If (prevSetOffRem = 0) Then
                Sheet16.Range("stcg.HPlossCurYrSetoff1b").value = 0
            Else
              Sheet16.Range("stcg.HPlossCurYrSetoff1b").value = stcgAppRateSetOffRem
                
            End If
        Else
              Sheet16.Range("stcg.HPlossCurYrSetoff1b").value = prevSetOffRem
              setOffRem = 0
       End If
       End If
       End If
       
'Added by Chandru
    If Sheet16.Range("stcg.HPlossCurYrSetoff1b").value > Sheet16.Range("stcg.IncOfCurYrUnderThatHead1b").value Then
    Sheet16.Range("stcg.HPlossCurYrSetoff1b").value = 0
    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
    End If
    If Sheet16.Range("SUM_STCGapprte").value > Sheet16.Range("stcg.IncOfCurYrUnderThatHead1b").value Then
    Sheet16.Range("stcg.HPlossCurYrSetoff1b", "stcg.BusLossSetoff1b").value = 0
    Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1b").value = 0
    MsgBox ("The sum of Column No. 2 + 3 should not exceed amount as referred in Column No. 1")
    End If
    
'End change

'---STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
' 11) Short-term capital gain taxable @ 15%
Dim stcg20PercntSetOffRem As Double

'added by Chandru
 If Sheet1.Range("Radio_Button_115").value = 2 Then

        Sheet16.Unprotect Password:=getmsgstate

        Sheet16.Range("stcg.HPlossCurYrSetoff1c").Interior.Color = (&HD8D8D8)
        Sheet16.Range("stcg.HPlossCurYrSetoff1c").Locked = True
        Sheet16.Range("stcg.HPlossCurYrSetoff1c").value = 0
    
Else

        If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
            Sheet16.Unprotect Password:=getmsgstate
            Sheet16.Range("stcg.HPlossCurYrSetoff1c").Locked = False
            Sheet16.Range("stcg.HPlossCurYrSetoff1c").Interior.Color = RGB(255, 255, 204)
        
        Else
        
                Sheet16.Range("stcg.HPlossCurYrSetoff1c").Locked = True
                Sheet16.Range("stcg.HPlossCurYrSetoff1c").Interior.ColorIndex = 2
        
         'end Change Chandru
         
                stcg20PercntSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("stcg.IncOfCurYrUnderThatHead1c").value - Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1c").value)
                
                prevSetOffRem = setOffRem
                setOffRem = setOffRem - stcg20PercntSetOffRem
                If (setOffRem >= 0) Then
                    If (prevSetOffRem = 0) Then
                        Sheet16.Range("stcg.HPlossCurYrSetoff1c").value = 0
                    Else
                            Sheet16.Range("stcg.HPlossCurYrSetoff1c").value = stcg20PercntSetOffRem
                            stcg20PercntSetOffRem = 0
                            stcgPercent20Sec111a = 0
                            stcgPTI20 = 0
                            stcgPercent20Sec115ad_1_b_ii = 0
                            
                    End If
                Else
                       Sheet16.Range("stcg.HPlossCurYrSetoff1c").value = prevSetOffRem
                     
                        If (prevSetOffRem >= 0 And stcgPercent20Sec115ad_1_b_ii > prevSetOffRem) Then
                            stcgPercent20Sec115ad_1_b_ii = stcgPercent20Sec115ad_1_b_ii - prevSetOffRem
                            prevSetOffRem = 0
                        Else
                            prevSetOffRem = prevSetOffRem - stcgPercent20Sec115ad_1_b_ii
                            stcgPercent20Sec115ad_1_b_ii = 0
                        End If
                        
                         If (prevSetOffRem >= 0 And stcgPTI20 > prevSetOffRem) Then
                            stcgPTI20 = stcgPTI20 - prevSetOffRem
                            prevSetOffRem = 0
                        Else
                            prevSetOffRem = prevSetOffRem - stcgPTI20
                            stcgPTI20 = 0
                        End If
        
                        If (prevSetOffRem >= 0 And stcgPercent20Sec111a > prevSetOffRem) Then
                            stcgPercent20Sec111a = stcgPercent20Sec111a - prevSetOffRem
                            prevSetOffRem = 0
                        Else
                            prevSetOffRem = prevSetOffRem - stcgPercent20Sec111a
                            stcgPercent20Sec111a = 0
                        End If
                   
                    setOffRem = 0
                End If
        End If
End If
        
        
'Added by Chandru
If Sheet16.Range("stcg.HPlossCurYrSetoff1c").value > Sheet16.Range("stcg.IncOfCurYrUnderThatHead1c").value Then
    Sheet16.Range("stcg.HPlossCurYrSetoff1c").value = 0
    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
End If
If Sheet16.Range("SUM_STCG20").value > Sheet16.Range("stcg.IncOfCurYrUnderThatHead1c").value Then
    Sheet16.Range("stcg.HPlossCurYrSetoff1c", "stcg.BusLossSetoff1c").value = 0
    Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1c").value = 0
    MsgBox ("The sum of Column No. 2 + 3 should not exceed amount as referred in Column No. 1")
End If
    
    'End change
    
'Commented by Riyaz on 29/01/2026
'10) Long term capital gain taxable @ 20%
Dim ltcg20PercntSetOffRem As Double
Dim prevLtcgSetOffRem As Double

'added by Chandru
' If Sheet1.Range("Radio_Button_115").value = 2 Then
'
' Sheet16.Unprotect Password:=getmsgstate
'
'        Sheet16.Range("ltcg.HPlossCurYrSetoff2a").Interior.Color = (&HD8D8D8)
'        Sheet16.Range("ltcg.HPlossCurYrSetoff2a").Locked = True
'        Sheet16.Range("ltcg.HPlossCurYrSetoff2a").value = 0
'
'        Else
'
'        If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
'        Sheet16.Unprotect Password:=getmsgstate
'        Sheet16.Range("ltcg.HPlossCurYrSetoff2a").Locked = False
'        Sheet16.Range("ltcg.HPlossCurYrSetoff2a").Interior.Color = RGB(255, 255, 204)
'
'        Else
'
'        Sheet16.Range("ltcg.HPlossCurYrSetoff2a").Locked = True
'        Sheet16.Range("ltcg.HPlossCurYrSetoff2a").Interior.ColorIndex = 2
'
' 'end Change Chandru
'
'
'
'        prevSetOffRem = setOffRem
'  ltcg20PercntSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2a").value - Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").value)
'        If (ltcg20PercntSetOffRem >= 0 And setOffRem >= 0) Then
'            If (ltcg20PercntSetOffRem >= setOffRem) Then
'               tempAk = setOffRem
'                setOffRem = setOffRem - ltcg20PercntSetOffRem
'
'                If (setOffRem >= 0) Then
'                       Sheet16.Range("ltcg.HPlossCurYrSetoff2a").value = ltcg20PercntSetOffRem
'                        ltcgPrctg20Sec112 = 0
'                        ltcgPTI20 = 0
'                        ltcgPrctg20Sec11EA = 0
'                        ltcgPrctg20SecProviso = 0 ''Newly added by Bindu as per DE V4 on 15th March 25
'                       ltcg20PercntSetOffRem = Application.WorksheetFunction.Max(0, ltcg20PercntSetOffRem - Sheet16.Range("ltcg.HPlossCurYrSetoff2a").value)
'                Else
'                      Sheet16.Range("ltcg.HPlossCurYrSetoff2a").value = tempAk
'                        setOffRem = 0
'                     If (tempAk >= 0 And ltcgPrctg20Sec11EA > tempAk) Then
'                            ltcgPrctg20Sec11EA = ltcgPrctg20Sec11EA - tempAk
'                            tempAk = 0
'                    Else
'                            tempAk = tempAk - ltcgPrctg20Sec11EA
'                            ltcgPrctg20Sec11EA = 0
'                    End If
'
'
'                    If (tempAk >= 0 And ltcgPTI20 > tempAk) Then
'                            ltcgPTI20 = ltcgPTI20 - tempAk
'                            tempAk = 0
'                    Else
'                            tempAk = tempAk - ltcgPTI20
'                            ltcgPTI20 = 0
'                    End If
'
'                    'Newly added by Bindu as per DE V4 on 15th March 25
'                    If (tempAk >= 0 And ltcgPrctg20SecProviso > tempAk) Then
'                            ltcgPrctg20SecProviso = ltcgPrctg20SecProviso - tempAk
'                            tempAk = 0
'                    Else
'                            tempAk = tempAk - ltcgPrctg20SecProviso
'                            ltcgPrctg20SecProviso = 0
'                    End If
'                    '--
'
'                     If (tempAk >= 0 And ltcgPrctg20Sec112 > tempAk) Then
'                            ltcgPrctg20Sec112 = ltcgPrctg20Sec112 - tempAk
'                            tempAk = 0
'                     Else
'                            tempAk = tempAk - ltcgPrctg20Sec112
'                            ltcgPrctg20Sec112 = 0
'                     End If
'                    ltcg20PercntSetOffRem = Application.WorksheetFunction.Max(0, ltcg20PercntSetOffRem - Sheet16.Range("ltcg.HPlossCurYrSetoff2a").value)
'                End If
'         Else
'            prevLtcgSetOffRem = ltcg20PercntSetOffRem
'
'                If (ltcg20PercntSetOffRem >= 0) Then
'                         Sheet16.Range("ltcg.HPlossCurYrSetoff2a").value = ltcg20PercntSetOffRem
'                        ltcg20PercntSetOffRem = 0
'                        ltcgPrctg20Sec112 = 0
'                        ltcgPTI20 = 0
'                        ltcgPrctg20Sec11EA = 0
'                        ltcgPrctg20SecProviso = 0 'Newly added by Bindu as per DE V4 on 15th March 25
'                       setOffRem = Application.WorksheetFunction.Max(0, setOffRem - Sheet16.Range("ltcg.HPlossCurYrSetoff2a").value)
'                Else
'                       Sheet16.Range("ltcg.HPlossCurYrSetoff2a").value = prevLtcgSetOffRem
'                        ltcg20PercntSetOffRem = 0
'                        ltcgPrctg20Sec112 = 0
'                        ltcgPTI20 = 0
'                        ltcgPrctg20Sec11EA = 0
'                        'Newly added by Bindu as per DE V4 on 15th March 25
'                        ltcgPrctg20SecProviso = 0
'                        setOffRem = Application.WorksheetFunction.Max(0, setOffRem - Sheet16.Range("ltcg.HPlossCurYrSetoff2a").value)
'                End If
'            End If
'        End If
'        End If
'        End If
'
''Added by Chandru
'    If Sheet16.Range("ltcg.HPlossCurYrSetoff2a").value > Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2a").value Then
'    Sheet16.Range("ltcg.HPlossCurYrSetoff2a").value = 0
'    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
'    End If
'    If Sheet16.Range("SUM_LTCG20").value > Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2a").value Then
'    Sheet16.Range("ltcg.HPlossCurYrSetoff2a", "ltcg.BusLossSetoff2a").value = 0
'    Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").value = 0
'    MsgBox ("The sum of Column No. 2 + 3 should not exceed amount as referred in Column No. 1")
'    End If
'End change


'Commented by Riyaz on 29/01/2026
' 11) Short-term capital gain taxable @ 15%
Dim stcg15PercntSetOffRem As Double

'added by Chandru
' If Sheet1.Range("Radio_Button_115").value = 2 Then
'
' Sheet16.Unprotect Password:=getmsgstate
'
'        Sheet16.Range("stcg.HPlossCurYrSetoff1").Interior.Color = (&HD8D8D8)
'        Sheet16.Range("stcg.HPlossCurYrSetoff1").Locked = True
'        Sheet16.Range("stcg.HPlossCurYrSetoff1").value = 0
'
'        Else
'
'        If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
'        Sheet16.Unprotect Password:=getmsgstate
'        Sheet16.Range("stcg.HPlossCurYrSetoff1").Locked = False
'        Sheet16.Range("stcg.HPlossCurYrSetoff1").Interior.Color = RGB(255, 255, 204)
'
'        Else
'
'        Sheet16.Range("stcg.HPlossCurYrSetoff1").Locked = True
'        Sheet16.Range("stcg.HPlossCurYrSetoff1").Interior.ColorIndex = 2
'
' 'end Change Chandru
'
'stcg15PercntSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("stcg.IncOfCurYrUnderThatHead1").value - Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1").value)
'
'        prevSetOffRem = setOffRem
'        setOffRem = setOffRem - stcg15PercntSetOffRem
'        If (setOffRem >= 0) Then
'            If (prevSetOffRem = 0) Then
'                Sheet16.Range("stcg.HPlossCurYrSetoff1").value = 0
'            Else
'                    Sheet16.Range("stcg.HPlossCurYrSetoff1").value = stcg15PercntSetOffRem
'                    stcg15PercntSetOffRem = 0
'                    stcgPercent15Sec111a = 0
'                    stcgPTI15 = 0
'                    stcgPercent15Sec115ad_1_b_ii = 0
'
'            End If
'        Else
'             Sheet16.Range("stcg.HPlossCurYrSetoff1").value = prevSetOffRem
'
'                If (prevSetOffRem >= 0 And stcgPercent15Sec115ad_1_b_ii > prevSetOffRem) Then
'                    stcgPercent15Sec115ad_1_b_ii = stcgPercent15Sec115ad_1_b_ii - prevSetOffRem
'                    prevSetOffRem = 0
'                Else
'                    prevSetOffRem = prevSetOffRem - stcgPercent15Sec115ad_1_b_ii
'                    stcgPercent15Sec115ad_1_b_ii = 0
'                End If
'
'                 If (prevSetOffRem >= 0 And stcgPTI15 > prevSetOffRem) Then
'                    stcgPTI15 = stcgPTI15 - prevSetOffRem
'                    prevSetOffRem = 0
'                Else
'                    prevSetOffRem = prevSetOffRem - stcgPTI15
'                    stcgPTI15 = 0
'                End If
'
'                If (prevSetOffRem >= 0 And stcgPercent15Sec111a > prevSetOffRem) Then
'                    stcgPercent15Sec111a = stcgPercent15Sec111a - prevSetOffRem
'                    prevSetOffRem = 0
'                Else
'                    prevSetOffRem = prevSetOffRem - stcgPercent15Sec111a
'                    stcgPercent15Sec111a = 0
'                End If
'
'            setOffRem = 0
'        End If
'        End If
'        End If
'
'
''Added by Chandru
'    If Sheet16.Range("stcg.HPlossCurYrSetoff1").value > Sheet16.Range("stcg.IncOfCurYrUnderThatHead1").value Then
'    Sheet16.Range("stcg.HPlossCurYrSetoff1").value = 0
'    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
'    End If
'    If Sheet16.Range("SUM_STCG15").value > Sheet16.Range("stcg.IncOfCurYrUnderThatHead1").value Then
'    Sheet16.Range("stcg.HPlossCurYrSetoff1", "stcg.BusLossSetoff1").value = 0
'    Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1").value = 0
'    MsgBox ("The sum of Column No. 2 + 3 should not exceed amount as referred in Column No. 1")
'    End If
    
    'End change
    
'---LTCG 12.5% NEWLY ADDED BY BINDU 26TH FEB 25
Dim ltcg125PercntSetOffRem As Double

'added by Chandru
 If Sheet1.Range("Radio_Button_115").value = 2 Then

 Sheet16.Unprotect Password:=getmsgstate

        Sheet16.Range("ltcg.HPlossCurYrSetoff2b").Interior.Color = (&HD8D8D8)
        Sheet16.Range("ltcg.HPlossCurYrSetoff2b").Locked = True
        Sheet16.Range("ltcg.HPlossCurYrSetoff2b").value = 0
    
        Else

        If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
        Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("ltcg.HPlossCurYrSetoff2b").Locked = False
        Sheet16.Range("ltcg.HPlossCurYrSetoff2b").Interior.Color = RGB(255, 255, 204)
        
        Else
        
        Sheet16.Range("ltcg.HPlossCurYrSetoff2b").Locked = True
        Sheet16.Range("ltcg.HPlossCurYrSetoff2b").Interior.ColorIndex = 2

 'end Change Chandru
 
 

        prevSetOffRem = setOffRem
     ltcg125PercntSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2b").value - Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2b").value)
        If (ltcg125PercntSetOffRem >= 0 And setOffRem >= 0) Then
            If (ltcg125PercntSetOffRem >= setOffRem) Then
                tempAk = setOffRem
                setOffRem = setOffRem - ltcg125PercntSetOffRem
                
                If (setOffRem >= 0) Then
                        Sheet16.Range("ltcg.HPlossCurYrSetoff2b").value = ltcg125PercntSetOffRem
                        ltcgPrctg125Sec115AC1 = 0
                        ltcgPrctg125Sec115Eb = 0
                        ltcgPrctg125Sec1121c2 = 0
                        ltcgPrctg125Sec115ACA1 = 0
                        ltcgPrctg125Sec115Ea = 0
                        ltcgPrctg125Sec115AD3_Proviso = 0
                        ltcgPrctg125Sec112A = 0
                        ltcgPTI125 = 0
                        ltcgPTI125_112 = 0
                        ltcgPTI112Sec48 = 0
                        ltcgPTI125112B9 = 0
                        ltcgPrctg125Sec115AD3 = 0
                        ltcgPrctg125SecProviso = 0 '15th March
                    ltcg125PercntSetOffRem = Application.WorksheetFunction.Max(0, ltcg125PercntSetOffRem - Sheet16.Range("ltcg.HPlossCurYrSetoff2b").value)
                Else
                       Sheet16.Range("ltcg.HPlossCurYrSetoff2b").value = tempAk
                        setOffRem = 0
                        '-1st Not using
If (tempAk >= 0 And ltcgPrctg125Sec115Ea > tempAk) Then
                                ltcgPrctg125Sec115Ea = ltcgPrctg125Sec115Ea - tempAk
                                tempAk = 0
                       Else
                                tempAk = tempAk - ltcgPrctg125Sec115Ea
                                ltcgPrctg125Sec115Ea = 0
                       End If
                       '--2nd not using
                        If (tempAk >= 0 And ltcgPTI112Sec48 > tempAk) Then
                                ltcgPTI112Sec48 = ltcgPTI112Sec48 - tempAk
                                tempAk = 0
                        Else
                              tempAk = tempAk - ltcgPTI112Sec48
                                ltcgPTI112Sec48 = 0
                        End If
                        '--10th
                         If (tempAk >= 0 And ltcgPTI125112B9 > tempAk) Then
                                ltcgPTI125112B9 = ltcgPTI125112B9 - tempAk
                                tempAk = 0
                        Else
                              tempAk = tempAk - ltcgPTI125112B9
                                ltcgPTI125112B9 = 0
                        End If
                        '--9th
                        If (tempAk >= 0 And ltcgPrctg125Sec1121c2 > tempAk) Then
                                ltcgPrctg125Sec1121c2 = ltcgPrctg125Sec1121c2 - tempAk
                                tempAk = 0
                        Else
                               tempAk = tempAk - ltcgPrctg125Sec1121c2
                                ltcgPrctg125Sec1121c2 = 0
                        End If
                        '--8th
                        If (tempAk >= 0 And ltcgPrctg125Sec115AC1 > tempAk) Then
                                ltcgPrctg125Sec115AC1 = ltcgPrctg125Sec115AC1 - tempAk
                                tempAk = 0
                        Else
                                tempAk = tempAk - ltcgPrctg125Sec115AC1
                                ltcgPrctg125Sec115AC1 = 0
                        End If
                        '7th
                        If (tempAk >= 0 And ltcgPrctg125Sec115ACA1 > tempAk) Then
                                ltcgPrctg125Sec115ACA1 = ltcgPrctg125Sec115ACA1 - tempAk
                                tempAk = 0
                       Else
                              tempAk = tempAk - ltcgPrctg125Sec115ACA1
                                ltcgPrctg125Sec115ACA1 = 0
                       End If
'Added by Riyaz 03/03/2026
                       If (tempAk >= 0 And ltcgPrctg125Sec115AD3 > tempAk) Then
                                ltcgPrctg125Sec115AD3 = ltcgPrctg125Sec115AD3 - tempAk
                                tempAk = 0
                       Else
                              tempAk = tempAk - ltcgPrctg125Sec115AD3
                                ltcgPrctg125Sec115AD3 = 0
                       End If
                       
                       '6th
                       If (tempAk >= 0 And ltcgPrctg125Sec115AD3_Proviso > tempAk) Then
                                ltcgPrctg125Sec115AD3_Proviso = ltcgPrctg125Sec115AD3_Proviso - tempAk
                                tempAk = 0
                       Else
                                tempAk = tempAk - ltcgPrctg125Sec115AD3_Proviso
                                ltcgPrctg125Sec115AD3_Proviso = 0
                       End If
                       '5th
                       If (tempAk >= 0 And ltcgPrctg125Sec115Eb > tempAk) Then
                                ltcgPrctg125Sec115Eb = ltcgPrctg125Sec115Eb - tempAk
                                tempAk = 0
                       Else
                                tempAk = tempAk - ltcgPrctg125Sec115Eb
                                ltcgPrctg125Sec115Eb = 0
                       End If
                       '4th
                       If (tempAk >= 0 And ltcgPTI125 > tempAk) Then
                                ltcgPTI125 = ltcgPTI125 - tempAk
                                tempAk = 0
                       Else
                                tempAk = tempAk - ltcgPTI125
                                ltcgPTI125 = 0
                       End If
                       '--3rd
                         '-
                        If (tempAk >= 0 And ltcgPrctg125SecProviso > tempAk) Then
                                ltcgPrctg125SecProviso = ltcgPrctg125SecProviso - tempAk
                                tempAk = 0
                        Else
                              tempAk = tempAk - ltcgPrctg125SecProviso
                                ltcgPrctg125SecProviso = 0
                        End If
                        '--
                        '--2nd
                         If (tempAk >= 0 And ltcgPTI125_112 > tempAk) Then
                                ltcgPTI125_112 = ltcgPTI125_112 - tempAk
                                tempAk = 0
                       Else
                                tempAk = tempAk - ltcgPTI125_112
                                ltcgPTI125_112 = 0
                       End If
                       '1st
                       If (tempAk >= 0 And ltcgPrctg125Sec112A > tempAk) Then
                                ltcgPrctg125Sec112A = ltcgPrctg125Sec112A - tempAk
                                tempAk = 0
                        Else
                              tempAk = tempAk - ltcgPrctg125Sec112A
                                ltcgPrctg125Sec112A = 0
                        End If
                   
                    ltcg125PercntSetOffRem = Application.WorksheetFunction.Max(0, ltcg125PercntSetOffRem - Sheet16.Range("ltcg.HPlossCurYrSetoff2b").value)
                End If
            Else
                prevLtcgSetOffRem = ltcg125PercntSetOffRem
                            
                If (ltcg125PercntSetOffRem >= 0) Then 'Correctly updated by Bindu on 24th March 25
                        Sheet16.Range("ltcg.HPlossCurYrSetoff2b").value = ltcg125PercntSetOffRem
                       ltcgPrctg125Sec115AC1 = 0
                        ltcgPrctg125Sec115Eb = 0
                        ltcgPrctg125Sec1121c2 = 0
                        ltcgPrctg125Sec115ACA1 = 0
                        ltcgPrctg125Sec115Ea = 0
                        ltcgPrctg125Sec115AD3_Proviso = 0
                        ltcgPrctg125Sec112A = 0
                        ltcgPTI125 = 0
                        ltcgPTI125_112 = 0
                        ltcgPTI112Sec48 = 0
                        ltcgPTI125112B9 = 0
                        ltcgPrctg125Sec115AD3 = 0
                        ltcgPrctg125SecProviso = 0 '15th March
                    setOffRem = Application.WorksheetFunction.Max(0, setOffRem - Sheet16.Range("ltcg.HPlossCurYrSetoff2b").value)
                End If
           End If
        End If
        End If
  End If
  
 'Added by Chandru
    If Sheet16.Range("ltcg.HPlossCurYrSetoff2b").value > Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2b").value Then
        Sheet16.Range("ltcg.HPlossCurYrSetoff2b").value = 0
       MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
    End If
    If Sheet16.Range("SUM_LTCG125").value > Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2b").value Then
        Sheet16.Range("ltcg.HPlossCurYrSetoff2b", "ltcg.BusLossSetoff2").value = 0
        Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2b").value = 0
        MsgBox ("The sum of Column No. 2 + 3 should not exceed amount as referred in Column No. 1")
    End If
    
'Commented by Riyaz on 29/01/2026
' 12) Long term capital gain taxable @ 10%
Dim ltcg10PercntSetOffRem As Double

'added by Chandru
' If Sheet1.Range("Radio_Button_115").value = 2 Then
'
' Sheet16.Unprotect Password:=getmsgstate
'
'        Sheet16.Range("ltcg.HPlossCurYrSetoff2").Interior.Color = (&HD8D8D8)
'        Sheet16.Range("ltcg.HPlossCurYrSetoff2").Locked = True
'        Sheet16.Range("ltcg.HPlossCurYrSetoff2").value = 0
'
'        Else
'
'        If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
'        Sheet16.Unprotect Password:=getmsgstate
'        Sheet16.Range("ltcg.HPlossCurYrSetoff2").Locked = False
'        Sheet16.Range("ltcg.HPlossCurYrSetoff2").Interior.Color = RGB(255, 255, 204)
'
'        Else
'
'        Sheet16.Range("ltcg.HPlossCurYrSetoff2").Locked = True
'        Sheet16.Range("ltcg.HPlossCurYrSetoff2").Interior.ColorIndex = 2
'
' 'end Change Chandru
'
'
'
'        prevSetOffRem = setOffRem
'     ltcg10PercntSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2").value - Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2").value)
'        If (ltcg10PercntSetOffRem >= 0 And setOffRem >= 0) Then
'            If (ltcg10PercntSetOffRem >= setOffRem) Then
'                tempAk = setOffRem
'                setOffRem = setOffRem - ltcg10PercntSetOffRem
'
'                If (setOffRem >= 0) Then
'                        Sheet16.Range("ltcg.HPlossCurYrSetoff2").value = ltcg10PercntSetOffRem
'                        ltcgPrctg10Sec112c2 = 0
'                        ltcgPrctg10Sec112A = 0
'                        ltcgPrctg10Sec115AC1 = 0
'                        ltcgPrctg10Sec115ACA1 = 0
'                        ltcgPrctg10Sec115AD3 = 0
'                        ltcgPrctg10Sec115AD3_Proviso = 0
'                        ltcgPrctg10Sec115Eb = 0
'                       ' ltcgPrctg10SecProviso = 0 ''NOt using  commented  by Bindu as this sec changed from 10 to 20
'                        ltcgPTI10 = 0
'                        ltcgPTI10_112 = 0
'                    ltcg10PercntSetOffRem = Application.WorksheetFunction.Max(0, ltcg10PercntSetOffRem - Sheet16.Range("ltcg.HPlossCurYrSetoff2").value)
'                Else
'                       Sheet16.Range("ltcg.HPlossCurYrSetoff2").value = tempAk
'                        setOffRem = 0
'                        If (tempAk >= 0 And ltcgPrctg10Sec112c2 > tempAk) Then
'                            ltcgPrctg10Sec112c2 = ltcgPrctg10Sec112c2 - tempAk
'                            tempAk = 0
'                        Else
'                            tempAk = tempAk - ltcgPrctg10Sec112c2
'                            ltcgPrctg10Sec112c2 = 0
'                        End If
'
'                        If (tempAk >= 0 And ltcgPrctg10Sec115AC1 > tempAk) Then
'                            ltcgPrctg10Sec115AC1 = ltcgPrctg10Sec115AC1 - tempAk
'                            tempAk = 0
'                       Else
'                            tempAk = tempAk - ltcgPrctg10Sec115AC1
'                            ltcgPrctg10Sec115AC1 = 0
'                        End If
'
'                        If (tempAk >= 0 And ltcgPrctg10Sec115ACA1 > tempAk) Then
'                            ltcgPrctg10Sec115ACA1 = ltcgPrctg10Sec115ACA1 - tempAk
'                            tempAk = 0
'                       Else
'                            tempAk = tempAk - ltcgPrctg10Sec115ACA1
'                            ltcgPrctg10Sec115ACA1 = 0
'                        End If
'
'                        If (tempAk >= 0 And ltcgPrctg10Sec115AD3 > tempAk) Then
'                            ltcgPrctg10Sec115AD3 = ltcgPrctg10Sec115AD3 - tempAk
'                            tempAk = 0
'                       Else
'                            tempAk = tempAk - ltcgPrctg10Sec115AD3
'                            ltcgPrctg10Sec115AD3 = 0
'                        End If
'
'                        If (tempAk >= 0 And ltcgPrctg10Sec115AD3_Proviso > tempAk) Then
'                            ltcgPrctg10Sec115AD3_Proviso = ltcgPrctg10Sec115AD3_Proviso - tempAk
'                            tempAk = 0
'                       Else
'                            tempAk = tempAk - ltcgPrctg10Sec115AD3_Proviso
'                            ltcgPrctg10Sec115AD3_Proviso = 0
'                        End If
'
'                        If (tempAk >= 0 And ltcgPrctg10Sec115Eb > tempAk) Then
'                            ltcgPrctg10Sec115Eb = ltcgPrctg10Sec115Eb - tempAk
'                            tempAk = 0
'                        Else
'                            tempAk = tempAk - ltcgPrctg10Sec115Eb
'                            ltcgPrctg10Sec115Eb = 0
'                        End If
'
'                         If (tempAk >= 0 And ltcgPTI10 > tempAk) Then
'                            ltcgPTI10 = ltcgPTI10 - tempAk
'                            tempAk = 0
'                        Else
'                            tempAk = tempAk - ltcgPTI10
'                            ltcgPTI10 = 0
'                        End If
'                        'NOt using  commented  by Bindu as this sec changed from 10 to 20
'
''                        If (tempAk >= 0 And ltcgPrctg10SecProviso > tempAk) Then
''                            ltcgPrctg10SecProviso = ltcgPrctg10SecProviso - tempAk
''                            tempAk = 0
''                       Else
''                            tempAk = tempAk - ltcgPrctg10SecProviso
''                            ltcgPrctg10SecProviso = 0
''                        End If
'
'                        If (tempAk >= 0 And ltcgPTI10_112 > tempAk) Then
'                            ltcgPTI10_112 = ltcgPTI10_112 - tempAk
'                            tempAk = 0
'                        Else
'                            tempAk = tempAk - ltcgPTI10_112
'                            ltcgPTI10_112 = 0
'                        End If
'
'                        If (tempAk >= 0 And ltcgPrctg10Sec112A > tempAk) Then
'                            ltcgPrctg10Sec112A = ltcgPrctg10Sec112A - tempAk
'                            tempAk = 0
'                        Else
'                            tempAk = tempAk - ltcgPrctg10Sec112A
'                            ltcgPrctg10Sec112A = 0
'                        End If
'
'                    ltcg10PercntSetOffRem = Application.WorksheetFunction.Max(0, ltcg10PercntSetOffRem - Sheet16.Range("ltcg.HPlossCurYrSetoff2").value)
'                End If
'            Else
'                prevLtcgSetOffRem = ltcg10PercntSetOffRem
'
'                If (ltcg10PercntSetOffRem >= 0) Then
'                        Sheet16.Range("ltcg.HPlossCurYrSetoff2").value = ltcg10PercntSetOffRem
'                        ltcg10PercntSetOffRem = 0
'                        ltcgPrctg10Sec112c2 = 0
'                        ltcgPrctg10Sec112A = 0
'                        ltcgPrctg10Sec115AC1 = 0
'                        ltcgPrctg10Sec115ACA1 = 0
'                        ltcgPrctg10Sec115AD3 = 0
'                        ltcgPrctg10Sec115AD3_Proviso = 0
'                        ltcgPrctg10Sec115Eb = 0
'                      '  ltcgPrctg10SecProviso = 0 ''NOt using  commented  by Bindu as this sec changed from 10 to 20
'                        ltcgPTI10 = 0
'                        ltcgPTI10_112 = 0
'                    setOffRem = Application.WorksheetFunction.Max(0, setOffRem - Sheet16.Range("ltcg.HPlossCurYrSetoff2").value)
'                End If
'           End If
'        End If
'        End If
'  End If
'
' 'Added by Chandru
'    If Sheet16.Range("ltcg.HPlossCurYrSetoff2").value > Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2").value Then
'    Sheet16.Range("ltcg.HPlossCurYrSetoff2").value = 0
'    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
'    End If
'    If Sheet16.Range("SUM_LTCG10").value > Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2").value Then
'    Sheet16.Range("ltcg.HPlossCurYrSetoff2", "ltcg.BusLossSetoff2").value = 0
'    Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2").value = 0
'    MsgBox ("The sum of Column No. 2 + 3 should not exceed amount as referred in Column No. 1")
'    End If
'End change


'Commented by Riyaz on 29/01/2026
'        cgosIncome_cgInc_stcg_prctg15_sec111a = stcgPercent15Sec111a
'        cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = stcgPercent15Sec115ad_1_b_ii
        '---STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
        cgosIncome_cgInc_stcg_prctg20_sec111a = stcgPercent20Sec111a
        cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii = stcgPercent20Sec115ad_1_b_ii
        
        
'        cgosIncome_cgInc_ltcg_prctg20_sec112 = ltcgPrctg20Sec112
'        cgosIncome_cgInc_ltcg_prctg20_sec11EA = ltcgPrctg20Sec11EA
        'Newly added by Bindu as per DE V4 on 15th March 25
'        cgosIncome_cgInc_ltcg_prctg20_secProviso = ltcgPrctg20SecProviso
        
'Commented by Riyaz on 29/01/2026
        'cgosIncome_cgInc_ltcg_prctg10_secProviso = ltcgPrctg10SecProviso ''NOt using  commented  by Bindu as this sec changed from 10 to 20
'        cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = ltcgPrctg10Sec112c2
'        cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 = ltcgPrctg10Sec115AC1
'        cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 = ltcgPrctg10Sec115ACA1
'        cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 = ltcgPrctg10Sec115AD3
'        cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso = ltcgPrctg10Sec115AD3_Proviso
'        cgosIncome_cgInc_ltcg_prctg10_sec112A = ltcgPrctg10Sec112A
'        cgosIncome_cgInc_ltcg_prctg10_sec115E_b = ltcgPrctg10Sec115Eb
'
'        cgosIncome_cgInc_ltcg_PTI10 = ltcgPTI10
'        cgosIncome_cgInc_ltcg_PTI10_112 = ltcgPTI10_112
        cgosIncome_cgInc_stcg_PTI30 = stcgPTI30
        cgosIncome_cgInc_stcg_prctg30_sec115AD = stcgPrctg30Sec115AD
'        cgosIncome_cgInc_stcg_PTI15 = stcgPTI15
        '----LTCG 12.5% NEWLY ADDED BY BINDU ON 26TH FEB 25
        cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 = ltcgPrctg125Sec115ACA1
        cgosIncome_cgInc_ltcg_prctg125_sec115E_b = ltcgPrctg125Sec115Eb
        cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 = ltcgPrctg125Sec1121c2
        cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 = ltcgPrctg125Sec115AC1
        cgosIncome_cgInc_ltcg_prctg125_sec115E_a = ltcgPrctg125Sec115Ea
        cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso = ltcgPrctg125Sec115AD3_Proviso
        cgosIncome_cgInc_ltcg_prctg125_sec112A = ltcgPrctg125Sec112A
        cgosIncome_cgInc_ltcg_PTI125 = ltcgPTI125
        cgosIncome_cgInc_ltcg_PTI125_112 = ltcgPTI125_112
        cgosIncome_cgInc_ltcg_PTI112_sec48 = ltcgPTI112Sec48
        cgosIncome_cgInc_ltcg_PTI125_112_B9 = ltcgPTI125112B9
        cgosIncome_cgInc_ltcg_prctg125_secProviso = ltcgPrctg125SecProviso  '15th March
        cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 = ltcgPrctg125Sec115AD3
        
        
        
        
        
        '---STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
        cgosIncome_cgInc_stcg_PTI20 = stcgPTI20
        
'        cgosIncome_cgInc_ltcg_PTI20 = ltcgPTI20
      
'13 STCG  HP DTAA Rate

'added by Chandru
 If Sheet1.Range("Radio_Button_115").value = 2 Then

 Sheet16.Unprotect Password:=getmsgstate

        Sheet16.Range("stcg.HPlossCurYrSetoff_DTAARates").Interior.Color = (&HD8D8D8)
        Sheet16.Range("stcg.HPlossCurYrSetoff_DTAARates").Locked = True
        Sheet16.Range("stcg.HPlossCurYrSetoff_DTAARates").value = 0
    
        Else

        If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
        Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("stcg.HPlossCurYrSetoff_DTAARates").Locked = False
        Sheet16.Range("stcg.HPlossCurYrSetoff_DTAARates").Interior.Color = RGB(255, 255, 204)
        
        Else
        
        Sheet16.Range("stcg.HPlossCurYrSetoff_DTAARates").Locked = True
        Sheet16.Range("stcg.HPlossCurYrSetoff_DTAARates").Interior.ColorIndex = 2

 'end Change Chandru
 
 
    Dim stcgDTAARateSetOffRem As Double
        
stcgDTAARateSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("stcg.IncOfCurYrUnderThatHead_DTAARates").value - Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff_DTAARates").value)
        prevSetOffRem = setOffRem
        setOffRem = setOffRem - stcgDTAARateSetOffRem
        If (setOffRem >= 0) Then
            If (prevSetOffRem = 0) Then
                Sheet16.Range("stcg.HPlossCurYrSetoff_DTAARates").value = 0
            Else
              Sheet16.Range("stcg.HPlossCurYrSetoff_DTAARates").value = stcgDTAARateSetOffRem
                
            End If
        Else
              Sheet16.Range("stcg.HPlossCurYrSetoff_DTAARates").value = prevSetOffRem
              setOffRem = 0
       End If
       End If
       End If
'Added by Chandru

    If Sheet16.Range("stcg.HPlossCurYrSetoff_DTAARates").value > Sheet16.Range("stcg.IncOfCurYrUnderThatHead_DTAARates").value Then
    Sheet16.Range("stcg.HPlossCurYrSetoff_DTAARates").value = 0
    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
    End If
    If Sheet16.Range("SUM_STCGDTAA").value > Sheet16.Range("stcg.IncOfCurYrUnderThatHead_DTAARates").value Then
    Sheet16.Range("stcg.HPlossCurYrSetoff_DTAARates", "stcg.BusLossSetoff_DTAARates").value = 0
    Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff_DTAARates").value = 0
    MsgBox ("The sum of Column No. 2 + 3 should not exceed amount as referred in Column No. 1")
    End If
    
'End change


'14 LTCG HP DTAA Rate
      Dim ltcgDTAARateSetOffRem As Double
      'added by Chandru
 If Sheet1.Range("Radio_Button_115").value = 2 Then

 Sheet16.Unprotect Password:=getmsgstate

        Sheet16.Range("ltcg.HPlossCurYrSetoff_DTAARates").Interior.Color = (&HD8D8D8)
        Sheet16.Range("ltcg.HPlossCurYrSetoff_DTAARates").Locked = True
        Sheet16.Range("ltcg.HPlossCurYrSetoff_DTAARates").value = 0
    
        Else

        If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
        Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("ltcg.HPlossCurYrSetoff_DTAARates").Locked = False
        Sheet16.Range("ltcg.HPlossCurYrSetoff_DTAARates").Interior.Color = RGB(255, 255, 204)
        
        Else
        
        Sheet16.Range("ltcg.HPlossCurYrSetoff_DTAARates").Locked = True
        Sheet16.Range("ltcg.HPlossCurYrSetoff_DTAARates").Interior.ColorIndex = 2

 'end Change Chandru
 
        
ltcgDTAARateSetOffRem = Application.WorksheetFunction.Max(0, Sheet16.Range("ltcg.IncOfCurYrUnderThatHead_DTAARates").value - Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff_DTAARates").value)
        prevSetOffRem = setOffRem
        setOffRem = setOffRem - ltcgDTAARateSetOffRem
        If (setOffRem >= 0) Then
            If (prevSetOffRem = 0) Then
                Sheet16.Range("ltcg.HPlossCurYrSetoff_DTAARates").value = 0
            Else
              Sheet16.Range("ltcg.HPlossCurYrSetoff_DTAARates").value = ltcgDTAARateSetOffRem
                
            End If
        Else
              Sheet16.Range("ltcg.HPlossCurYrSetoff_DTAARates").value = prevSetOffRem
              setOffRem = 0
       End If
        End If
  End If
 'Added by Chandru
    If Sheet16.Range("ltcg.HPlossCurYrSetoff_DTAARates").value > Sheet16.Range("ltcg.IncOfCurYrUnderThatHead_DTAARates").value Then
    Sheet16.Range("ltcg.HPlossCurYrSetoff_DTAARates").value = 0
    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
    End If
    If Sheet16.Range("SUM_LTCGDTAA").value > Sheet16.Range("ltcg.IncOfCurYrUnderThatHead_DTAARates").value Then
    Sheet16.Range("ltcg.HPlossCurYrSetoff_DTAARates", "ltcg.BusLossSetoff_DTAARates").value = 0
    Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff_DTAARates").value = 0
    MsgBox ("The sum of Column No. 2 + 3 should not exceed amount as referred in Column No. 1")
    End If

'End change

      
Application.EnableEvents = True
End Function

Function setOffOthSrcLossCYLA()

Application.EnableEvents = False

Dim stcgPercent15Sec111a As Double
Dim stcgPercent15Sec115ad_1_b_ii As Double

'STCG 20% NEWLY ADDED BY BINDU ON 19TH FEB 25
Dim stcgPercent20Sec111a As Double
Dim stcgPercent20Sec115ad_1_b_ii As Double


Dim ltcgPrctg20Sec112 As Double
Dim ltcgPrctg20Sec11EA As Double
'Newly added by Bindu as per DE V4 on 15th March 25
Dim ltcgPrctg20SecProviso As Double

Dim ltcgPrctg10SecProviso As Double
Dim ltcgPrctg10Sec112c2 As Double
Dim ltcgPrctg10Sec115AC1 As Double
Dim ltcgPrctg10Sec115ACA1 As Double
Dim ltcgPrctg10Sec115AD3 As Double
Dim ltcgPrctg10Sec115AD3_Proviso As Double
Dim ltcgPrctg10Sec115Eb As Double
Dim ltcgPrctg10Sec112A As Double
Dim stcgPrctg30Sec115AD  As Double
Dim stcgPTI30  As Double
Dim stcgPTI15  As Double

'LTCG 12.5% NEWLY ADDED BY BINDU ON 25TH FEN 25


Dim ltcgPrctg125Sec115ACA1 As Double
Dim ltcgPrctg125Sec115Eb As Double
Dim ltcgPrctg125Sec1121c2 As Double
Dim ltcgPrctg125Sec115AC1 As Double
Dim ltcgPrctg125Sec115Ea As Double
Dim ltcgPrctg125Sec115AD3_Proviso As Double
Dim ltcgPrctg125Sec112A As Double
Dim ltcgPTI112Sec48 As Double
Dim ltcgPTI125112B9 As Double
Dim ltcgPrctg125SecProviso As Double '15th March

Dim ltcgPrctg125Sec115AD3 As Double

'--STCG 20% NEWLY ADDED BY BINDU ON 19TH FEB 25
Dim stcgPTI20 As Double

Dim ltcgPTI10  As Double
Dim ltcgPTI10_112  As Double
Dim ltcgPTI20  As Double

'--LTCG 12.5% NEWLY ADDED BY BINDU ON 25TH FEB 25
Dim ltcgPTI125  As Double
Dim ltcgPTI125_112  As Double


'Commented by Riyaz on 29/01/2026
'stcgPercent15Sec111a = Application.WorksheetFunction.Max(0, Sheet31.Range("temp15Sec111a").value)
'stcgPercent15Sec115ad_1_b_ii = Application.WorksheetFunction.Max(0, Sheet31.Range("temp15Sec115ad1bii").value)
'--STCG 20% NEWLY ADDED BY BINDU ON 19TH FEB 25
stcgPercent20Sec111a = Application.WorksheetFunction.Max(0, Sheet31.Range("temp20Sec111a").value)
stcgPercent20Sec115ad_1_b_ii = Application.WorksheetFunction.Max(0, Sheet31.Range("temp20Sec115ad1bii").value)

'Commented by Riyaz on 29/01/2026
'ltcgPrctg20Sec112 = Application.WorksheetFunction.Max(0, Sheet31.Range("temp20Sec112").value)
'ltcgPrctg20Sec11EA = Application.WorksheetFunction.Max(0, Sheet31.Range("temp20Sec11EA").value)
'''Newly added by Bindu as per DE V4 on 15th March 25
'ltcgPrctg20SecProviso = Application.WorksheetFunction.Max(0, Sheet31.Range("temp20SecPro").value)

'Commented by Riyaz on 29/01/2026
'ltcgPrctg10SecProviso = Application.WorksheetFunction.Max(0, Sheet31.Range("temp10SecPro").value)'NOt using  commented  by Bindu as this sec changed from 10 to 20
'ltcgPrctg10Sec112c2 = Application.WorksheetFunction.Max(0, Sheet31.Range("temp10Sec1121c2").value)
'ltcgPrctg10Sec115AC1 = Application.WorksheetFunction.Max(0, Sheet31.Range("temp10Sec115AC1").value)
'ltcgPrctg10Sec115ACA1 = Application.WorksheetFunction.Max(0, Sheet31.Range("temp10Sec115ACA1").value)
'ltcgPrctg10Sec115AD3 = Application.WorksheetFunction.Max(0, Sheet31.Range("temp10Sec115AD3").value)
'ltcgPrctg10Sec115AD3_Proviso = Application.WorksheetFunction.Max(0, Sheet31.Range("temp10Sec115AD3_Proviso").value)
'ltcgPrctg10Sec112A = Application.WorksheetFunction.Max(0, Sheet31.Range("temp10Sec112A").value)
'ltcgPrctg10Sec115Eb = Application.WorksheetFunction.Max(0, Sheet31.Range("temp10Sec115Eb").value)


'----LTCG 12.5%
ltcgPrctg125Sec1121c2 = Application.WorksheetFunction.Max(0, Sheet31.Range("temp125Sec112_c").value)
ltcgPrctg125Sec115AC1 = Application.WorksheetFunction.Max(0, Sheet31.Range("temp125Sec_115AC").value)
ltcgPrctg125Sec115ACA1 = Application.WorksheetFunction.Max(0, Sheet31.Range("temp125Sec_115ACA").value)
ltcgPrctg125Sec115AD3_Proviso = Application.WorksheetFunction.Max(0, Sheet31.Range("temp125Sec_115AD_iii").value)
ltcgPrctg125Sec115Eb = Application.WorksheetFunction.Max(0, Sheet31.Range("temp125Sec_115EB").value)
ltcgPrctg125Sec115Ea = Application.WorksheetFunction.Max(0, Sheet31.Range("temp125Sec_115Ea").value)
ltcgPrctg125Sec112A = Application.WorksheetFunction.Max(0, Sheet31.Range("temp125Sec_112A").value)
ltcgPTI112Sec48 = Application.WorksheetFunction.Max(0, Sheet31.Range("temp125Sec_112sec_48").value)
ltcgPTI125112B9 = Application.WorksheetFunction.Max(0, Sheet31.Range("temp125Sec_112").value)
ltcgPrctg125SecProviso = Application.WorksheetFunction.Max(0, Sheet31.Range("temp125SecPro").value) '15th March
'Added by Riyaz on 03/03/2026
ltcgPrctg125Sec115AD3 = Application.WorksheetFunction.Max(0, Sheet31.Range("temp115AD125").value)

'new
 stcgPrctg30Sec115AD = Application.WorksheetFunction.Max(0, Sheet31.Range("temp115AD30").value)
stcgPTI30 = Application.WorksheetFunction.Max(0, Sheet31.Range("tempPTI30").value)
'Commented by Riyaz on 29/01/2026
'stcgPTI15 = Application.WorksheetFunction.Max(0, Sheet31.Range("tempPTI15").value)
'--STCG 20% NEWLY ADDED BY BINDU ON 19TH FEB 25
stcgPTI20 = Application.WorksheetFunction.Max(0, Sheet31.Range("tempPTI20_STCG").value)
'Commented by Riyaz on 29/01/2026
'ltcgPTI10 = Application.WorksheetFunction.Max(0, Sheet31.Range("tempPTI10").value)
'ltcgPTI10_112 = Application.WorksheetFunction.Max(0, Sheet31.Range("tempPTI10_112").value)
' ltcgPTI20 = Application.WorksheetFunction.Max(0, Sheet31.Range("tempPTI20").value)
'---LTCG 12.5% NEWLY ADDED BY BINDU 25TH FEB 25

ltcgPTI125_112 = Application.WorksheetFunction.Max(0, Sheet31.Range("temp125Sec_112_PTI").value)
ltcgPTI125 = Application.WorksheetFunction.Max(0, Sheet31.Range("temp125Sec_PTI").value)


'end

 

Dim setOffRem As Double
Dim prevSetOffRem As Double


setOffRem = Sheet16.Range("cyla.TotOthSrcLossNoRaceHorse").value - Sheet16.Range("rh.IncOfCurYrUnderThatHead4").value
                        
' 1. Race Horse
'added by Chandru
'If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
'Sheet16.Unprotect Password:=getmsgstate
'        Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4").Locked = False
'        Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4").Interior.Color = RGB(255, 255, 204)
'
'    Else
'
'        Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4").Locked = True
'        Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4").Interior.ColorIndex = 2
' 'end Change Chandru

        If (setOffRem >= 0) Then
            Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4").value = Sheet16.Range("rh.IncOfCurYrUnderThatHead4").value
            
        Else
           Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4").value = Sheet16.Range("cyla.TotOthSrcLossNoRaceHorse").value
            setOffRem = 0
        End If
   'End If
   
   'Added by Chandru
If Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4").value > Sheet16.Range("rh.IncOfCurYrUnderThatHead4").value Then
Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4").value = 0
MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")

End If
'End change





' 2 DTAA OS Income

'added by Chandru
If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4_DTAARates").Locked = False
        Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4_DTAARates").Interior.Color = RGB(255, 255, 204)

    Else

        Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4_DTAARates").Locked = True
        Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4_DTAARates").Interior.ColorIndex = 2
 'end Change Chandru
 
 
      prevSetOffRem = setOffRem
      setOffRem = setOffRem - Sheet16.Range("rh.IncOfCurYrUnderThatHead4_DTAARates").value
      If (setOffRem >= 0) Then
            Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4_DTAARates").value = Sheet16.Range("rh.IncOfCurYrUnderThatHead4_DTAARates").value

        Else
           Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4_DTAARates").value = prevSetOffRem
            setOffRem = 0
        End If
        End If
        
   'Added by Chandru
If Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4_DTAARates").value > Sheet16.Range("rh.IncOfCurYrUnderThatHead4_DTAARates").value Then
Sheet16.Range("rh.OthSrcLossNoRaceHorseSetoff4_DTAARates").value = 0
MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")

End If
'End change
        
        
'3.salary

'added by Chandru
If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("salary.OthSrcLossNoRaceHorseSetoff1").Locked = False
        Sheet16.Range("salary.OthSrcLossNoRaceHorseSetoff1").Interior.Color = RGB(255, 255, 204)

    Else

        Sheet16.Range("salary.OthSrcLossNoRaceHorseSetoff1").Locked = True
        Sheet16.Range("salary.OthSrcLossNoRaceHorseSetoff1").Interior.ColorIndex = 2
 'end Change Chandru
 
 
prevSetOffRem = setOffRem
        setOffRem = setOffRem - Sheet16.Range("salary.IncOfCurYrUnderThatHead1").value
        
        If setOffRem >= 0 Then
           Sheet16.Range("salary.OthSrcLossNoRaceHorseSetoff1").value = Sheet16.Range("salary.IncOfCurYrUnderThatHead1").value
              
        Else
           Sheet16.Range("salary.OthSrcLossNoRaceHorseSetoff1").value = prevSetOffRem
            setOffRem = 0
       End If
       End If
       
'Added by Chandru
If Sheet16.Range("salary.OthSrcLossNoRaceHorseSetoff1").value > Sheet16.Range("SAL.TotIncUnderHeadSalaries1").value Then
Sheet16.Range("salary.OthSrcLossNoRaceHorseSetoff1").value = 0
MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")

End If
'End change
        
'4.House property

'added by Chandru
If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("hp.OthSrcLossNoRaceHorseSetoff2").Locked = False
        Sheet16.Range("hp.OthSrcLossNoRaceHorseSetoff2").Interior.Color = RGB(255, 255, 204)

    Else

        Sheet16.Range("hp.OthSrcLossNoRaceHorseSetoff2").Locked = True
        Sheet16.Range("hp.OthSrcLossNoRaceHorseSetoff2").Interior.ColorIndex = 2

'end Change Chandru
 
 
        prevSetOffRem = setOffRem
        setOffRem = setOffRem - Sheet16.Range("hp.IncOfCurYrUnderThatHead2").value
        If setOffRem >= 0 Then
           Sheet16.Range("hp.OthSrcLossNoRaceHorseSetoff2").value = Sheet16.Range("hp.IncOfCurYrUnderThatHead2").value
        Else
            Sheet16.Range("hp.OthSrcLossNoRaceHorseSetoff2").value = prevSetOffRem
            setOffRem = 0
        End If
        
        End If
        
'Added by Chandru

    If Sheet16.Range("hp.OthSrcLossNoRaceHorseSetoff2").value > Sheet16.Range("hp.IncOfCurYrUnderThatHead2").value Then
    Sheet16.Range("hp.OthSrcLossNoRaceHorseSetoff2").value = 0
    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
    End If

'End change

        
'5. Business Income

'added by Chandru
        If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
        Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("busprof.OthSrcLossNoRaceHorseSetoff").Locked = False
        Sheet16.Range("busprof.OthSrcLossNoRaceHorseSetoff").Interior.Color = RGB(255, 255, 204)

         Else

        Sheet16.Range("busprof.OthSrcLossNoRaceHorseSetoff").Locked = True
        Sheet16.Range("busprof.OthSrcLossNoRaceHorseSetoff").Interior.ColorIndex = 2
 'end Change Chandru
 
 
        prevSetOffRem = setOffRem
        setOffRem = setOffRem - Sheet16.Range("busprof.IncOfCurYrUnderThatHead").value
        If setOffRem >= 0 Then
            Sheet16.Range("busprof.OthSrcLossNoRaceHorseSetoff").value = Sheet16.Range("busprof.IncOfCurYrUnderThatHead").value
        Else
             Sheet16.Range("busprof.OthSrcLossNoRaceHorseSetoff").value = prevSetOffRem
             setOffRem = 0
        End If
        
        End If
        
        
'Added by Chandru
If Sheet16.Range("busprof.OthSrcLossNoRaceHorseSetoff").value > Sheet16.Range("busprof.IncOfCurYrUnderThatHead").value Then
Sheet16.Range("busprof.OthSrcLossNoRaceHorseSetoff").value = 0
MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")

End If
'End change
        
'Speculative Income

'added by Chandru
        If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
        Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("busprofspec.OthSrcLossNoRaceHorseSetoff0a").Locked = False
        Sheet16.Range("busprofspec.OthSrcLossNoRaceHorseSetoff0a").Interior.Color = RGB(255, 255, 204)

         Else

        Sheet16.Range("busprofspec.OthSrcLossNoRaceHorseSetoff0a").Locked = True
        Sheet16.Range("busprofspec.OthSrcLossNoRaceHorseSetoff0a").Interior.ColorIndex = 2
 'end Change Chandru
 
        prevSetOffRem = setOffRem
        setOffRem = setOffRem - Sheet16.Range("busprofspec.IncOfCurYrUnderThatHead0a").value
        If setOffRem >= 0 Then
            Sheet16.Range("busprofspec.OthSrcLossNoRaceHorseSetoff0a").value = Sheet16.Range("busprofspec.IncOfCurYrUnderThatHead0a").value
        Else
             Sheet16.Range("busprofspec.OthSrcLossNoRaceHorseSetoff0a").value = prevSetOffRem
             setOffRem = 0
        End If
        End If

'Added by Chandru

    If Sheet16.Range("busprofspec.OthSrcLossNoRaceHorseSetoff0a").value > Sheet16.Range("busprofspec.IncOfCurYrUnderThatHead0a").value Then
    Sheet16.Range("busprofspec.OthSrcLossNoRaceHorseSetoff0a").value = 0
    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
    End If

'End change


'Specified Income

'added by Chandru
        If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
        Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("busprofspecified.OthSrcLossNoRaceHorseSetoff0b").Locked = False
        Sheet16.Range("busprofspecified.OthSrcLossNoRaceHorseSetoff0b").Interior.Color = RGB(255, 255, 204)

         Else

        Sheet16.Range("busprofspecified.OthSrcLossNoRaceHorseSetoff0b").Locked = True
        Sheet16.Range("busprofspecified.OthSrcLossNoRaceHorseSetoff0b").Interior.ColorIndex = 2
 'end Change Chandru
 
 
        prevSetOffRem = setOffRem
        setOffRem = setOffRem - Sheet16.Range("busprofspecified.IncOfCurYrUnderThatHead0b").value
        If setOffRem >= 0 Then
            Sheet16.Range("busprofspecified.OthSrcLossNoRaceHorseSetoff0b").value = Sheet16.Range("busprofspecified.IncOfCurYrUnderThatHead0b").value
        Else
             Sheet16.Range("busprofspecified.OthSrcLossNoRaceHorseSetoff0b").value = prevSetOffRem
             setOffRem = 0
        End If
        End If
  
 'Added by Chandru

    If Sheet16.Range("busprofspecified.OthSrcLossNoRaceHorseSetoff0b").value > Sheet16.Range("busprofspecified.IncOfCurYrUnderThatHead0b").value Then
    Sheet16.Range("busprofspecified.OthSrcLossNoRaceHorseSetoff0b").value = 0
    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
    End If

'End change


'6. Short-term capital gain taxable @ 30%

 'added by Chandru
        If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
        Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1a").Locked = False
        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1a").Interior.Color = RGB(255, 255, 204)

    Else

        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1a").Locked = True
        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1a").Interior.ColorIndex = 2
 'end Change Chandru
        
        
         prevSetOffRem = setOffRem
        setOffRem = setOffRem - Sheet16.Range("stcg.IncOfCurYrUnderThatHead1a").value
        If (setOffRem >= 0) Then
           Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1a").value = Sheet16.Range("stcg.IncOfCurYrUnderThatHead1a").value
                stcgPrctg30Sec115AD = 0
                stcgPTI30 = 0
       Else
        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1a").value = prevSetOffRem
             If (prevSetOffRem >= 0 And stcgPrctg30Sec115AD > prevSetOffRem) Then
                    stcgPrctg30Sec115AD = stcgPrctg30Sec115AD - prevSetOffRem
                    prevSetOffRem = 0
            Else
                    prevSetOffRem = prevSetOffRem - stcgPrctg30Sec115AD
                    stcgPrctg30Sec115AD = 0
            End If

            If (prevSetOffRem >= 0 And stcgPTI30 > prevSetOffRem) Then
                    stcgPTI30 = stcgPTI30 - prevSetOffRem
                    prevSetOffRem = 0
            Else
                  prevSetOffRem = prevSetOffRem - stcgPTI30
                    stcgPTI30 = 0
            End If
        setOffRem = 0
     End If
      End If
 
 'Added by Chandru

    If Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1a").value > Sheet16.Range("stcg.IncOfCurYrUnderThatHead1a").value Then
    Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1a").value = 0
    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
    End If

'End change

'7.Short-term capital gain taxable at applicable rates
'added by Chandru
If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1b").Locked = False
        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1b").Interior.Color = RGB(255, 255, 204)

    Else

        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1b").Locked = True
        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1b").Interior.ColorIndex = 2
'end Change Chandru

         prevSetOffRem = setOffRem
        setOffRem = setOffRem - Sheet16.Range("stcg.IncOfCurYrUnderThatHead1b").value
        If setOffRem >= 0 Then
            Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1b").value = Sheet16.Range("stcg.IncOfCurYrUnderThatHead1b").value
        Else
             Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1b").value = prevSetOffRem
             setOffRem = 0
        End If
        End If
  'Added by Chandru

    If Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1b").value > Sheet16.Range("stcg.IncOfCurYrUnderThatHead1b").value Then
    Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1b").value = 0
    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
    End If

'End change


'8) STCG 20% NEWLY ADDED BY BINDU ON 19TH FEB 25

'added by Chandru IN AY 24-25 LOGIC
If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1c").Locked = False
        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1c").Interior.Color = RGB(255, 255, 204)

    Else

        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1c").Locked = True
        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1c").Interior.ColorIndex = 2
'end Change Chandru


        prevSetOffRem = setOffRem
        setOffRem = setOffRem - Sheet16.Range("stcg.IncOfCurYrUnderThatHead1c").value
        If (setOffRem >= 0) Then
           Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1c").value = Sheet16.Range("stcg.IncOfCurYrUnderThatHead1c").value
                stcgPercent20Sec111a = 0
                stcgPTI20 = 0
                stcgPercent20Sec115ad_1_b_ii = 0
       Else
        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1c").value = prevSetOffRem
             If (prevSetOffRem >= 0 And stcgPercent20Sec115ad_1_b_ii > prevSetOffRem) Then
                    stcgPercent20Sec115ad_1_b_ii = stcgPercent20Sec115ad_1_b_ii - prevSetOffRem
                    prevSetOffRem = 0
            Else
                    prevSetOffRem = prevSetOffRem - stcgPercent20Sec115ad_1_b_ii
                    stcgPercent20Sec115ad_1_b_ii = 0
            End If
            
            If (prevSetOffRem >= 0 And stcgPTI20 > prevSetOffRem) Then
                    stcgPTI20 = stcgPTI20 - prevSetOffRem
                    prevSetOffRem = 0
            Else
                    prevSetOffRem = prevSetOffRem - stcgPTI20
                    stcgPTI20 = 0
            End If
            

            If (prevSetOffRem >= 0 And stcgPercent20Sec111a > prevSetOffRem) Then
                    stcgPercent20Sec111a = stcgPercent20Sec111a - prevSetOffRem
                    prevSetOffRem = 0
            Else
                  prevSetOffRem = prevSetOffRem - stcgPercent20Sec111a
                    stcgPercent20Sec111a = 0
            End If
        setOffRem = 0
     End If
     End If
'Added by Chandru AY 24-25 LOGIC

    If Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1c").value > Sheet16.Range("stcg.IncOfCurYrUnderThatHead1c").value Then
    Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1c").value = 0
    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
    End If
'End Change

'Commented by Riyaz on 29/01/2026
'' 8)Long term capital gain taxable @ 20%
Dim ltcgSetOffRem As Double
Dim tempAk As Double
Dim prevLtcgSetOffRem As Double
Dim remaLTCGcomplete As Double
'
''added by Chandru
'If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
'Sheet16.Unprotect Password:=getmsgstate
'        Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").Locked = False
'        Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").Interior.Color = RGB(255, 255, 204)
'
'    Else
'
'        Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").Locked = True
'        Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").Interior.ColorIndex = 2
''end Change Chandru
'
'
'    prevSetOffRem = setOffRem
'    ltcgSetOffRem = Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2a").value
'    If (ltcgSetOffRem >= 0 And setOffRem >= 0) Then
'            If (ltcgSetOffRem >= setOffRem) Then
'                 tempAk = setOffRem
'                 setOffRem = setOffRem - ltcgSetOffRem
'                If (setOffRem >= 0) Then
'                    Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").value = ltcgSetOffRem
'                        ltcgPrctg20Sec11EA = 0
'                        ltcgPTI20 = 0
'                        ltcgPrctg20Sec112 = 0
'                        ''Newly added by Bindu as per DE V4 on 15th March 25
'                        ltcgPrctg20SecProviso = 0
'
'                    ltcgSetOffRem = Application.WorksheetFunction.Max(0, ltcgSetOffRem - Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").value)
'                Else
'                    Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").value = tempAk
'                    setOffRem = 0
'
'                    If (tempAk >= 0 And ltcgPrctg20Sec11EA > tempAk) Then
'                            ltcgPrctg20Sec11EA = ltcgPrctg20Sec11EA - tempAk
'                            tempAk = 0
'                    Else
'                            tempAk = tempAk - ltcgPrctg20Sec11EA
'                            ltcgPrctg20Sec11EA = 0
'                    End If
'
'
'
'                    If (tempAk >= 0 And ltcgPTI20 > tempAk) Then
'                            ltcgPTI20 = ltcgPTI20 - tempAk
'                            tempAk = 0
'                    Else
'                            tempAk = tempAk - ltcgPTI20
'                            ltcgPTI20 = 0
'                    End If
'
'                    '- 'Newly added by Bindu as per DE V4 on 15th March 25
'                    If (tempAk >= 0 And ltcgPrctg20SecProviso > tempAk) Then
'                            ltcgPrctg20SecProviso = ltcgPrctg20SecProviso - tempAk
'                            tempAk = 0
'                    Else
'                            tempAk = tempAk - ltcgPrctg20SecProviso
'                            ltcgPrctg20SecProviso = 0
'                    End If
'                    '--
'
'
'                    If (tempAk >= 0 And ltcgPrctg20Sec112 > tempAk) Then
'                        ltcgPrctg20Sec112 = ltcgPrctg20Sec112 - tempAk
'                            tempAk = 0
'                    Else
'                        tempAk = tempAk - ltcgPrctg20Sec112
'                            ltcgPrctg20Sec112 = 0
'                    End If
'                   ltcgSetOffRem = Application.WorksheetFunction.Max(0, ltcgSetOffRem - Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").value)
'                End If
'            Else
'              prevLtcgSetOffRem = ltcgSetOffRem
'                 If (ltcgSetOffRem >= 0) Then
'                       Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").value = ltcgSetOffRem
'                        ltcgSetOffRem = 0
'                        ltcgPrctg20Sec112 = 0
'                        ltcgPTI20 = 0
'                        ltcgPrctg20Sec11EA = 0
'                        ''Newly added by Bindu as per DE V4 on 15th March 25
'                        ltcgPrctg20SecProviso = 0
'                    setOffRem = Application.WorksheetFunction.Max(0, setOffRem - Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").value)
'                Else
'                   Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").value = prevLtcgSetOffRem
'                        ltcgSetOffRem = 0
'                        ltcgPrctg20Sec112 = 0
'                        ltcgPTI20 = 0
'                        ltcgPrctg20Sec11EA = 0
'                        'Newly added by Bindu as per DE V4 on 15th March 25
'                        ltcgPrctg20SecProviso = 0
'                    setOffRem = Application.WorksheetFunction.Max(0, setOffRem - Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").value)
'                End If
'            End If
'        End If
'       End If
'
''Added by Chandru
'
'    If Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").value > Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2a").value Then
'    Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2a").value = 0
'    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
'    End If
'
''End change

'Commented by Riyaz on 29/01/2026
'8.Short-term capital gain taxable @ 15%

'added by Chandru
'If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
'Sheet16.Unprotect Password:=getmsgstate
'        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1").Locked = False
'        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1").Interior.Color = RGB(255, 255, 204)
'
'    Else
'
'        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1").Locked = True
'        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1").Interior.ColorIndex = 2
''end Change Chandru
'
'
'        prevSetOffRem = setOffRem
'        setOffRem = setOffRem - Sheet16.Range("stcg.IncOfCurYrUnderThatHead1").value
'        If (setOffRem >= 0) Then
'           Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1").value = Sheet16.Range("stcg.IncOfCurYrUnderThatHead1").value
'                stcgPercent15Sec111a = 0
'                stcgPTI15 = 0
'                stcgPercent15Sec115ad_1_b_ii = 0
'       Else
'        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1").value = prevSetOffRem
'             If (prevSetOffRem >= 0 And stcgPercent15Sec115ad_1_b_ii > prevSetOffRem) Then
'                    stcgPercent15Sec115ad_1_b_ii = stcgPercent15Sec115ad_1_b_ii - prevSetOffRem
'                    prevSetOffRem = 0
'            Else
'                    prevSetOffRem = prevSetOffRem - stcgPercent15Sec115ad_1_b_ii
'                    stcgPercent15Sec115ad_1_b_ii = 0
'            End If
'
'            If (prevSetOffRem >= 0 And stcgPTI15 > prevSetOffRem) Then
'                    stcgPTI15 = stcgPTI15 - prevSetOffRem
'                    prevSetOffRem = 0
'            Else
'                    prevSetOffRem = prevSetOffRem - stcgPTI15
'                    stcgPTI15 = 0
'            End If
'
'
'            If (prevSetOffRem >= 0 And stcgPercent15Sec111a > prevSetOffRem) Then
'                    stcgPercent15Sec111a = stcgPercent15Sec111a - prevSetOffRem
'                    prevSetOffRem = 0
'            Else
'                  prevSetOffRem = prevSetOffRem - stcgPercent15Sec111a
'                    stcgPercent15Sec111a = 0
'            End If
'        setOffRem = 0
'     End If
'     End If
''Added by Chandru
'
'    If Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1").value > Sheet16.Range("stcg.IncOfCurYrUnderThatHead1").value Then
'    Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff1").value = 0
'    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
'    End If

'End change
     
'---LTCG 12.5% NEWLY ADDED BY BINDU 25TH FEB 25
If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2b").Locked = False
        Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2b").Interior.Color = RGB(255, 255, 204)
    Else
        Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2b").Locked = True
        Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2b").Interior.ColorIndex = 2
'end Change Chandru

 prevSetOffRem = setOffRem
ltcgSetOffRem = Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2b").value
        If (ltcgSetOffRem >= 0 And setOffRem >= 0) Then
            If (ltcgSetOffRem >= setOffRem) Then
                tempAk = setOffRem
                setOffRem = setOffRem - ltcgSetOffRem
                   If (setOffRem >= 0) Then
                       Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2b").value = ltcgSetOffRem
                        ltcgPrctg125Sec115AC1 = 0
                        ltcgPrctg125Sec115Eb = 0
                        ltcgPrctg125Sec1121c2 = 0
                        ltcgPrctg125Sec115ACA1 = 0
                        ltcgPrctg125Sec115Ea = 0
                        ltcgPrctg125Sec115AD3_Proviso = 0
                        ltcgPrctg125Sec112A = 0
                        ltcgPTI125 = 0
                        ltcgPTI125_112 = 0
                        ltcgPTI112Sec48 = 0
                        ltcgPTI125112B9 = 0
                        ltcgPrctg125Sec115AD3 = 0
                        ltcgPrctg125SecProviso = 0 '15th March
                      ltcgSetOffRem = Application.WorksheetFunction.Max(0, ltcgSetOffRem - Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2b").value)
                  Else
                    Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2b").value = tempAk
                     setOffRem = 0
                         '-1st Not using
If (tempAk >= 0 And ltcgPrctg125Sec115Ea > tempAk) Then
                                ltcgPrctg125Sec115Ea = ltcgPrctg125Sec115Ea - tempAk
                                tempAk = 0
                       Else
                                tempAk = tempAk - ltcgPrctg125Sec115Ea
                                ltcgPrctg125Sec115Ea = 0
                       End If
                       '--2nd not using
                        If (tempAk >= 0 And ltcgPTI112Sec48 > tempAk) Then
                                ltcgPTI112Sec48 = ltcgPTI112Sec48 - tempAk
                                tempAk = 0
                        Else
                              tempAk = tempAk - ltcgPTI112Sec48
                                ltcgPTI112Sec48 = 0
                        End If
                        '--10th
                         If (tempAk >= 0 And ltcgPTI125112B9 > tempAk) Then
                                ltcgPTI125112B9 = ltcgPTI125112B9 - tempAk
                                tempAk = 0
                        Else
                              tempAk = tempAk - ltcgPTI125112B9
                                ltcgPTI125112B9 = 0
                        End If
                        '--9th
                        If (tempAk >= 0 And ltcgPrctg125Sec1121c2 > tempAk) Then
                                ltcgPrctg125Sec1121c2 = ltcgPrctg125Sec1121c2 - tempAk
                                tempAk = 0
                        Else
                               tempAk = tempAk - ltcgPrctg125Sec1121c2
                                ltcgPrctg125Sec1121c2 = 0
                        End If
                        '--8th
                        If (tempAk >= 0 And ltcgPrctg125Sec115AC1 > tempAk) Then
                                ltcgPrctg125Sec115AC1 = ltcgPrctg125Sec115AC1 - tempAk
                                tempAk = 0
                        Else
                                tempAk = tempAk - ltcgPrctg125Sec115AC1
                                ltcgPrctg125Sec115AC1 = 0
                        End If
                        '7th
                        If (tempAk >= 0 And ltcgPrctg125Sec115ACA1 > tempAk) Then
                                ltcgPrctg125Sec115ACA1 = ltcgPrctg125Sec115ACA1 - tempAk
                                tempAk = 0
                       Else
                              tempAk = tempAk - ltcgPrctg125Sec115ACA1
                                ltcgPrctg125Sec115ACA1 = 0
                       End If
'Added by Riyaz on 03/03/2026
                       If (tempAk >= 0 And ltcgPrctg125Sec115AD3 > tempAk) Then
                                ltcgPrctg125Sec115AD3 = ltcgPrctg125Sec115AD3 - tempAk
                                tempAk = 0
                       Else
                              tempAk = tempAk - ltcgPrctg125Sec115AD3
                                ltcgPrctg125Sec115AD3 = 0
                       End If
                                               
                       '6th
                       If (tempAk >= 0 And ltcgPrctg125Sec115AD3_Proviso > tempAk) Then
                                ltcgPrctg125Sec115AD3_Proviso = ltcgPrctg125Sec115AD3_Proviso - tempAk
                                tempAk = 0
                       Else
                                tempAk = tempAk - ltcgPrctg125Sec115AD3_Proviso
                                ltcgPrctg125Sec115AD3_Proviso = 0
                       End If
                       '5th
                       If (tempAk >= 0 And ltcgPrctg125Sec115Eb > tempAk) Then
                                ltcgPrctg125Sec115Eb = ltcgPrctg125Sec115Eb - tempAk
                                tempAk = 0
                       Else
                                tempAk = tempAk - ltcgPrctg125Sec115Eb
                                ltcgPrctg125Sec115Eb = 0
                       End If
                       '4th
                       If (tempAk >= 0 And ltcgPTI125 > tempAk) Then
                                ltcgPTI125 = ltcgPTI125 - tempAk
                                tempAk = 0
                       Else
                                tempAk = tempAk - ltcgPTI125
                                ltcgPTI125 = 0
                       End If
                       '--3rd
                         '-
                        If (tempAk >= 0 And ltcgPrctg125SecProviso > tempAk) Then
                                ltcgPrctg125SecProviso = ltcgPrctg125SecProviso - tempAk
                                tempAk = 0
                        Else
                              tempAk = tempAk - ltcgPrctg125SecProviso
                                ltcgPrctg125SecProviso = 0
                        End If
                        '--
                        '--2nd
                         If (tempAk >= 0 And ltcgPTI125_112 > tempAk) Then
                                ltcgPTI125_112 = ltcgPTI125_112 - tempAk
                                tempAk = 0
                       Else
                                tempAk = tempAk - ltcgPTI125_112
                                ltcgPTI125_112 = 0
                       End If
                       '1st
                       If (tempAk >= 0 And ltcgPrctg125Sec112A > tempAk) Then
                                ltcgPrctg125Sec112A = ltcgPrctg125Sec112A - tempAk
                                tempAk = 0
                        Else
                              tempAk = tempAk - ltcgPrctg125Sec112A
                                ltcgPrctg125Sec112A = 0
                        End If
                        
                    ltcgSetOffRem = Application.WorksheetFunction.Max(0, ltcgSetOffRem - Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2b").value)
                End If
            Else
                prevLtcgSetOffRem = ltcgSetOffRem
                If (ltcgSetOffRem >= 0) Then
                       Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2b").value = ltcgSetOffRem
                        ltcgSetOffRem = 0
                        ltcgPrctg125Sec115AC1 = 0
                        ltcgPrctg125Sec115Eb = 0
                        ltcgPrctg125Sec1121c2 = 0
                        ltcgPrctg125Sec115ACA1 = 0
                        ltcgPrctg125Sec115Ea = 0
                        ltcgPrctg125Sec115AD3_Proviso = 0
                        ltcgPrctg125Sec112A = 0
                        ltcgPTI125 = 0
                        ltcgPTI125_112 = 0
                        ltcgPTI112Sec48 = 0
                        ltcgPTI125112B9 = 0
                        ltcgPrctg125Sec115AD3 = 0
                        ltcgPrctg125SecProviso = 0 '15th March
                    setOffRem = Application.WorksheetFunction.Max(0, setOffRem - Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2b").value)
              End If
            End If
        End If
        End If

'Added by Chandru

    If Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2b").value > Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2b").value Then
      Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2b").value = 0
      MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
    End If
'---

'Commented by Riyaz on 29/01/2026
' 9)Long term capital gain taxable @ 10%
'added by Chandru
'If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
'Sheet16.Unprotect Password:=getmsgstate
'        Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2").Locked = False
'        Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2").Interior.Color = RGB(255, 255, 204)
'
'    Else
'
'        Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2").Locked = True
'        Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2").Interior.ColorIndex = 2
''end Change Chandru
'
' prevSetOffRem = setOffRem
'ltcgSetOffRem = Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2").value
'        If (ltcgSetOffRem >= 0 And setOffRem >= 0) Then
'            If (ltcgSetOffRem >= setOffRem) Then
'                tempAk = setOffRem
'                setOffRem = setOffRem - ltcgSetOffRem
'                   If (setOffRem >= 0) Then
'                       Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2").value = ltcgSetOffRem
'                         ltcgPrctg10Sec112c2 = 0
'                        ltcgPrctg10Sec115AC1 = 0
'                        ltcgPrctg10Sec115ACA1 = 0
'                        ltcgPrctg10Sec115AD3 = 0
'                        ltcgPrctg10Sec115AD3_Proviso = 0
'                        ltcgPrctg10Sec115Eb = 0
'                        ltcgPTI10 = 0
'                        ltcgPTI10_112 = 0
'                        'ltcgPrctg10SecProviso = 0 ''NOt using  commented  by Bindu as this sec changed from 10 to 20
'                        ltcgPrctg10Sec112A = 0
'                      ltcgSetOffRem = Application.WorksheetFunction.Max(0, ltcgSetOffRem - Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2").value)
'                  Else
'                    Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2").value = tempAk
'                     setOffRem = 0
'                         If (tempAk >= 0 And ltcgPrctg10Sec112c2 > tempAk) Then
'                                ltcgPrctg10Sec112c2 = ltcgPrctg10Sec112c2 - tempAk
'                                tempAk = 0
'                        Else
'                               tempAk = tempAk - ltcgPrctg10Sec112c2
'                                ltcgPrctg10Sec112c2 = 0
'                        End If
'
'                        If (tempAk >= 0 And ltcgPrctg10Sec115AC1 > tempAk) Then
'                                ltcgPrctg10Sec115AC1 = ltcgPrctg10Sec115AC1 - tempAk
'                                tempAk = 0
'                        Else
'                                tempAk = tempAk - ltcgPrctg10Sec115AC1
'                                ltcgPrctg10Sec115AC1 = 0
'                        End If
'
'                        If (tempAk >= 0 And ltcgPrctg10Sec115ACA1 > tempAk) Then
'                                ltcgPrctg10Sec115ACA1 = ltcgPrctg10Sec115ACA1 - tempAk
'                                tempAk = 0
'                       Else
'                              tempAk = tempAk - ltcgPrctg10Sec115ACA1
'                                ltcgPrctg10Sec115ACA1 = 0
'                       End If
'
'                       If (tempAk >= 0 And ltcgPrctg10Sec115AD3 > tempAk) Then
'                                ltcgPrctg10Sec115AD3 = ltcgPrctg10Sec115AD3 - tempAk
'                                tempAk = 0
'                       Else
'                                tempAk = tempAk - ltcgPrctg10Sec115AD3
'                                ltcgPrctg10Sec115AD3 = 0
'                       End If
'
'                       If (tempAk >= 0 And ltcgPrctg10Sec115AD3_Proviso > tempAk) Then
'                                ltcgPrctg10Sec115AD3_Proviso = ltcgPrctg10Sec115AD3_Proviso - tempAk
'                                tempAk = 0
'                       Else
'                                tempAk = tempAk - ltcgPrctg10Sec115AD3_Proviso
'                                ltcgPrctg10Sec115AD3_Proviso = 0
'                       End If
'
'
'                       If (tempAk >= 0 And ltcgPrctg10Sec115Eb > tempAk) Then
'                                ltcgPrctg10Sec115Eb = ltcgPrctg10Sec115Eb - tempAk
'                                tempAk = 0
'                       Else
'                                tempAk = tempAk - ltcgPrctg10Sec115Eb
'                                ltcgPrctg10Sec115Eb = 0
'                       End If
'
'
'
'                       If (tempAk >= 0 And ltcgPTI10 > tempAk) Then
'                                ltcgPTI10 = ltcgPTI10 - tempAk
'                                tempAk = 0
'                       Else
'                                tempAk = tempAk - ltcgPTI10
'                                ltcgPTI10 = 0
'                       End If
'                            'NOt using  commented  by Bindu as this sec changed from 10 to 20
'
''                        If (tempAk >= 0 And ltcgPrctg10SecProviso > tempAk) Then
''                                ltcgPrctg10SecProviso = ltcgPrctg10SecProviso - tempAk
''                                tempAk = 0
''                        Else
''                              tempAk = tempAk - ltcgPrctg10SecProviso
''                                ltcgPrctg10SecProviso = 0
''                        End If
'
'                        If (tempAk >= 0 And ltcgPTI10_112 > tempAk) Then
'                                ltcgPTI10_112 = ltcgPTI10_112 - tempAk
'                                tempAk = 0
'                       Else
'                                tempAk = tempAk - ltcgPTI10_112
'                                ltcgPTI10_112 = 0
'                       End If
'
'                        If (tempAk >= 0 And ltcgPrctg10Sec112A > tempAk) Then
'                                ltcgPrctg10Sec112A = ltcgPrctg10Sec112A - tempAk
'                                tempAk = 0
'                        Else
'                              tempAk = tempAk - ltcgPrctg10Sec112A
'                                ltcgPrctg10Sec112A = 0
'                        End If
'                    ltcgSetOffRem = Application.WorksheetFunction.Max(0, ltcgSetOffRem - Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2").value)
'                End If
'            Else
'                prevLtcgSetOffRem = ltcgSetOffRem
'                If (ltcgSetOffRem >= 0) Then
'                       Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2").value = ltcgSetOffRem
'                        ltcgSetOffRem = 0
'                        ltcgPrctg10Sec112c2 = 0
'                        ltcgPrctg10Sec115AC1 = 0
'                        ltcgPrctg10Sec115ACA1 = 0
'                        ltcgPrctg10Sec115AD3 = 0
'                        ltcgPrctg10Sec115AD3_Proviso = 0
'                        ltcgPrctg10Sec115Eb = 0
'                        'ltcgPrctg10SecProviso = 0 'NOt using  commented  by Bindu as this sec changed from 10 to 20
'                        ltcgPrctg10Sec112A = 0
'                        ltcgPTI10 = 0
'                        ltcgPTI10_112 = 0
'                    setOffRem = Application.WorksheetFunction.Max(0, setOffRem - Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2").value)
'              End If
'            End If
'        End If
'        End If
'
''Added by Chandru
'
'    If Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2").value > Sheet16.Range("ltcg.IncOfCurYrUnderThatHead2").value Then
'    Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff2").value = 0
'    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
'    End If

'End change
'Commented by Riyaz on 29/01/2026
'        cgosIncome_cgInc_stcg_prctg15_sec111a = stcgPercent15Sec111a
'        cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = stcgPercent15Sec115ad_1_b_ii
        '--STCG 20% NEWLY ADDED BY BINDU ON 19TH FEB 25
        cgosIncome_cgInc_stcg_prctg20_sec111a = stcgPercent20Sec111a
        cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii = stcgPercent20Sec115ad_1_b_ii
        
'Commented by Riyaz on 29/01/2026
'        cgosIncome_cgInc_ltcg_prctg20_sec112 = ltcgPrctg20Sec112
'        cgosIncome_cgInc_ltcg_prctg20_sec11EA = ltcgPrctg20Sec11EA
        'Newly added by Bindu as per DE V4 on 15th March 25
'        cgosIncome_cgInc_ltcg_prctg20_secProviso = ltcgPrctg20SecProviso
        
'Commented by Riyaz on 29/01/2026
       ' cgosIncome_cgInc_ltcg_prctg10_secProviso = ltcgPrctg10SecProviso ''NOt using  commented  by Bindu as this sec changed from 10 to 20
'        cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = ltcgPrctg10Sec112c2
'        cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 = ltcgPrctg10Sec115AC1
'        cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 = ltcgPrctg10Sec115ACA1
'        cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 = ltcgPrctg10Sec115AD3
'        cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso = ltcgPrctg10Sec115AD3_Proviso
'        cgosIncome_cgInc_ltcg_prctg10_sec115E_b = ltcgPrctg10Sec115Eb
'        cgosIncome_cgInc_ltcg_prctg10_sec112A = ltcgPrctg10Sec112A
        
        'LTCG 12.5% NEWLY ADDED BY BINDU 25TH FEB 25
        cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 = ltcgPrctg125Sec115ACA1
        cgosIncome_cgInc_ltcg_prctg125_sec115E_b = ltcgPrctg125Sec115Eb
        cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 = ltcgPrctg125Sec1121c2
        cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 = ltcgPrctg125Sec115AC1
        cgosIncome_cgInc_ltcg_prctg125_sec115E_a = ltcgPrctg125Sec115Ea
        cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso = ltcgPrctg125Sec115AD3_Proviso
        cgosIncome_cgInc_ltcg_prctg125_sec112A = ltcgPrctg125Sec112A
        cgosIncome_cgInc_ltcg_PTI125 = ltcgPTI125
        cgosIncome_cgInc_ltcg_PTI125_112 = ltcgPTI125_112
        cgosIncome_cgInc_ltcg_PTI112_sec48 = ltcgPTI112Sec48
        cgosIncome_cgInc_ltcg_PTI125_112_B9 = ltcgPTI125112B9
        cgosIncome_cgInc_ltcg_prctg125_secProviso = ltcgPrctg125SecProviso '15th March
        cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 = ltcgPrctg125Sec115AD3

        '--STCG 20% NEWLY ADDED BY BINDU
        cgosIncome_cgInc_stcg_PTI20 = stcgPTI20
        
'        cgosIncome_cgInc_ltcg_PTI20 = ltcgPTI20
        
        
        '--
'        cgosIncome_cgInc_ltcg_PTI10 = ltcgPTI10

'        cgosIncome_cgInc_ltcg_PTI10_112 = ltcgPTI10_112

        cgosIncome_cgInc_stcg_PTI30 = stcgPTI30

        cgosIncome_cgInc_stcg_prctg30_sec115AD = stcgPrctg30Sec115AD

'        cgosIncome_cgInc_stcg_PTI15 = stcgPTI15

'        cgosIncome_cgInc_ltcg_PTI20 = ltcgPTI20
 
        
'10.Short-term capital gain taxable at DTAA Rate
'added by Chandru
If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff_DTAARates").Locked = False
        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff_DTAARates").Interior.Color = RGB(255, 255, 204)

    Else

        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff_DTAARates").Locked = True
        Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff_DTAARates").Interior.ColorIndex = 2
'end Change Chandru

         prevSetOffRem = setOffRem
        setOffRem = setOffRem - Sheet16.Range("stcg.IncOfCurYrUnderThatHead_DTAARates").value
        If setOffRem >= 0 Then
            Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff_DTAARates").value = Sheet16.Range("stcg.IncOfCurYrUnderThatHead_DTAARates").value
        Else
             Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff_DTAARates").value = prevSetOffRem
             setOffRem = 0
        End If
          End If
          
'Added by Chandru

    If Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff_DTAARates").value > Sheet16.Range("stcg.IncOfCurYrUnderThatHead_DTAARates").value Then
    Sheet16.Range("stcg.OthSrcLossNoRaceHorseSetoff_DTAARates").value = 0
    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
    End If

'End change

        
'11.Long-term capital gain taxable at DTAA Rate

'added by Chandru
If (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
Sheet16.Unprotect Password:=getmsgstate
        Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff_DTAARates").Locked = False
        Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff_DTAARates").Interior.Color = RGB(255, 255, 204)

    Else

        Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff_DTAARates").Locked = True
        Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff_DTAARates").Interior.ColorIndex = 2
'end Change Chandru

         prevSetOffRem = setOffRem
        setOffRem = setOffRem - Sheet16.Range("ltcg.IncOfCurYrUnderThatHead_DTAARates").value
        If setOffRem >= 0 Then
            Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff_DTAARates").value = Sheet16.Range("ltcg.IncOfCurYrUnderThatHead_DTAARates").value
        Else
             Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff_DTAARates").value = prevSetOffRem
             setOffRem = 0
        End If
            End If
'Added by Chandru

    If Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff_DTAARates").value > Sheet16.Range("ltcg.IncOfCurYrUnderThatHead_DTAARates").value Then
    Sheet16.Range("ltcg.OthSrcLossNoRaceHorseSetoff_DTAARates").value = 0
    MsgBox ("Total Losses set off cannot exceed losses that can be set-off and incomes available for set off.")
    End If

'End change

        
  Application.EnableEvents = True
End Function

Function calcBFLA()
Application.EnableEvents = False
Application.ScreenUpdating = False

Sheet17.Unprotect Password:=getmsgstate

Dim stcgBreakUp1 As Double
Dim stcgBreakUp2 As Double
'Commented by Riyaz on 30/01/2026
'Dim stcgBreakUp3 As Double
'Dim stcgBreakUp4 As Double
Dim stcgBreakUp5 As Double
Dim stcgBreakUp6 As Double
'Commented by Riyaz on 30/01/2026
'Dim stcgBreakUp7 As Double

'---stcg 20% newly added by Bindu 19th Feb 25
Dim stcgBreakUp3_New As Double
Dim stcgBreakUp7_New As Double
Dim stcgBreakUp4_New As Double

'Commented by Riyaz on 30/01/2026
'Dim ltcgBreakUp1 As Double
'Dim ltcgBreakUp2 As Double

'Newly added by Bindu as per DE V4 on 15th March 25
'Dim ltcgBreakUp2_New As Double

'Commented by Riyaz on 30/01/2026
'Dim ltcgBreakUp3 As Double
'Dim ltcgBreakUp4 As Double
'Dim ltcgBreakUp5 As Double
'Dim ltcgBreakUp6 As Double
'Dim ltcgBreakUp7 As Double
'Dim ltcgBreakUp8 As Double
Dim ltcgBreakUp9 As Double
'Dim ltcgBreakUp7_1 As Double
'Dim ltcgBreakUp7_2 As Double
'Dim ltcgBreakUp10 As Double
'Dim ltcgBreakUp11 As Double
'Dim ltcgBreakUp12 As Double

'--LTCG 12.5% NEWLY ADDED BY BINDU
Dim ltcgBreakUp4_New  As Double
Dim ltcgBreakUp5_New  As Double
Dim ltcgBreakUp6_New  As Double
Dim ltcgBreakUp8_New  As Double
Dim ltcgBreakUp7_1_New  As Double
Dim ltcgBreakUp7_2_New  As Double
Dim ltcgBreakUp11_New  As Double
Dim ltcgBreakUp12_New  As Double
Dim ltcgBreakUp13_New  As Double
Dim ltcgBreakUp14_New  As Double
Dim ltcgBreakUp15_New  As Double
Dim ltcgBreakUp16_New As Double '15th March
Dim ltcgBreakUp17_New As Double


Dim tempstcgBreakUp1 As Double
Dim tempstcgBreakUp2 As Double
'Commented by Riyaz on 30/01/2026
'Dim tempstcgBreakUp3 As Double
'Dim tempstcgBreakUp4 As Double

Dim tempstcgBreakUp5 As Double
Dim tempstcgBreakUp6 As Double
'Commented by Riyaz on 30/01/2026
'Dim tempstcgBreakUp7 As Double

'---stcg 20% newly added by bindu 19th FEB 25
Dim tempstcgBreakUp3_New As Double
Dim tempstcgBreakUp4_New As Double
Dim tempstcgBreakUp7_New As Double


'Commented by Riyaz on 30/01/2026
'Dim templtcgBreakUp1 As Double
'Dim templtcgBreakUp2 As Double

'Newly added by Bindu as per DE V4 on 15th March 25
'Dim templtcgBreakUp2_New As Double

'Commented by Riyaz on 30/01/2026
'Dim templtcgBreakUp3 As Double
'Dim templtcgBreakUp4 As Double
'Dim templtcgBreakUp5 As Double
'Dim templtcgBreakUp6 As Double
'Dim templtcgBreakUp7 As Double
'Dim templtcgBreakUp7_1 As Double
'Dim templtcgBreakUp7_2 As Double
'Dim templtcgBreakUp8 As Double
Dim templtcgBreakUp9 As Double
'Dim templtcgBreakUp10 As Double
'Dim templtcgBreakUp11 As Double
'Dim templtcgBreakUp12 As Double

'--LTCG 12.5% NEWLY ADDED BY BINDU 26TH FEB 25
Dim templtcgBreakUp4_New  As Double
Dim templtcgBreakUp5_New  As Double
Dim templtcgBreakUp6_New  As Double
Dim templtcgBreakUp8_New As Double
Dim templtcgBreakUp7_1_New  As Double
Dim templtcgBreakUp7_2_New  As Double
Dim templtcgBreakUp11_New  As Double
Dim templtcgBreakUp12_New  As Double
Dim templtcgBreakUp13_New  As Double
Dim templtcgBreakUp14_New  As Double
Dim templtcgBreakUp15_New  As Double
Dim templtcgBreakUp16_New As Double '15th March
Dim templtcgBreakUp17_New As Double

Dim tempLtcgCFL As Double
Dim tempstcgCFL As Double

'Commented by Riyaz on 30/01/2026
'cgosIncome_cgInc_stcg_prctg15_sec111a = Sheet31.Range("CYLAtemp15Sec111a").value
'cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = Sheet31.Range("CYLAtemp15Sec115ad1bii").value

'---STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
cgosIncome_cgInc_stcg_prctg20_sec111a = Sheet31.Range("CYLAtemp20Sec111a").value
cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii = Sheet31.Range("CYLAtemp20Sec115ad1bii").value


'Commented by Riyaz on 30/01/2026
'cgosIncome_cgInc_ltcg_prctg20_sec112 = Sheet31.Range("CYLAtemp20Sec112").value
'cgosIncome_cgInc_ltcg_prctg20_sec11EA = Sheet31.Range("CYLAtemp20Sec11EA").value
'Newly added by Bindu as per DE V4 on 15th March 25
'cgosIncome_cgInc_ltcg_prctg20_secProviso = Sheet31.Range("CYLAtemp20SecPro").value

'Commented by Riyaz on 30/01/2026
'cgosIncome_cgInc_ltcg_prctg10_secProviso = Sheet31.Range("CYLAtemp10SecPro").value ''NOt using  commented  by Bindu as this sec changed from 10 to 20
'cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = Sheet31.Range("CYLAtemp10Sec1121c2").value
'cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 = Sheet31.Range("CYLAtemp10Sec115AC1").value
'cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 = Sheet31.Range("CYLAtemp10Sec115ACA1").value
'cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 = Sheet31.Range("CYLAtemp10Sec115AD3").value
'cgosIncome_cgInc_ltcg_prctg10_sec115E_b = Sheet31.Range("CYLAtemp10Sec115Eb").value
'cgosIncome_cgInc_ltcg_prctg10_sec112A = Sheet31.Range("CYLAtemp10Sec112A").value
'cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso = Sheet31.Range("CYLAtemp10Sec115AD3_Proviso").value

'Commented by Riyaz on 30/01/2026
'cgosIncome_cgInc_ltcg_PTI10 = Sheet31.Range("CYLAtempPTI10").value
'cgosIncome_cgInc_ltcg_PTI10_112 = Sheet31.Range("CYLAtempPTI10_112").value
cgosIncome_cgInc_stcg_PTI30 = Sheet31.Range("CYLAtempPTI30").value
cgosIncome_cgInc_stcg_prctg30_sec115AD = Sheet31.Range("CYLAtemp115AD30").value
'cgosIncome_cgInc_stcg_PTI15 = Sheet31.Range("CYLAtempPTI15").value
'----LTCG 12.5% NEWLY ADDED BY BINDU 26TH FEB 25
 cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 = Sheet31.Range("CYLAtemp125Sec112_c").value
 cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 = Sheet31.Range("CYLAtemp125Sec_115ACA").value
 cgosIncome_cgInc_ltcg_prctg125_sec115E_b = Sheet31.Range("CYLAtemp125Sec_115EB").value
 cgosIncome_cgInc_ltcg_prctg125_sec115E_a = Sheet31.Range("CYLAtemp125Sec_115EA").value
 cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 = Sheet31.Range("CYLAtemp125Sec_115AC").value
 cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso = Sheet31.Range("CYLAtemp125Sec_115AD_iii").value
 cgosIncome_cgInc_ltcg_PTI125 = Sheet31.Range("CYLAtemp125Sec_PTI").value
 cgosIncome_cgInc_ltcg_PTI125_112 = Sheet31.Range("CYLAtemp125Sec_112_PTI").value
 cgosIncome_cgInc_ltcg_PTI125_112_B9 = Sheet31.Range("CYLAtemp125Sec_112").value
 cgosIncome_cgInc_ltcg_prctg125_sec112A = Sheet31.Range("CYLAtemp125Sec_112A").value
 cgosIncome_cgInc_ltcg_PTI112_sec48 = Sheet31.Range("CYLAtemp125Sec_112sec_48").value
cgosIncome_cgInc_ltcg_prctg125_secProviso = Sheet31.Range("CYLAtemp125SecPro").value '15th March
cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 = Sheet31.Range("CYLAtemp115AD125").value




'---STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
cgosIncome_cgInc_stcg_PTI20 = Sheet31.Range("CYLAtempPTI20_STCG").value
'Commented by Riyaz on 30/01/2026
'cgosIncome_cgInc_ltcg_PTI20 = Sheet31.Range("CYLAtempPTI20").value

stcgBreakUp1 = cgosIncome_cgInc_stcg_prctg30_sec115AD 'Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3a").value  'cgosIncome.cgInc.stcg.prctg30
stcgBreakUp6 = cgosIncome_cgInc_stcg_PTI30
stcgBreakUp2 = Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3b").value 'cgosIncome.cgInc.stcg.prctgAr
'Commented by Riyaz on 30/01/2026
'stcgBreakUp3 = cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii
'stcgBreakUp7 = cgosIncome_cgInc_stcg_PTI15
'stcgBreakUp4 = cgosIncome_cgInc_stcg_prctg15_sec111a

'---STCG 20% NEWLY ADDED BY BINDU
stcgBreakUp3_New = cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii
stcgBreakUp7_New = cgosIncome_cgInc_stcg_PTI20
stcgBreakUp4_New = cgosIncome_cgInc_stcg_prctg20_sec111a
'--END

stcgBreakUp5 = Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3b_DTAARates").value
        
'Commented by Riyaz on 30/01/2026
'ltcgBreakUp2 = cgosIncome_cgInc_ltcg_prctg20_sec11EA
''Newly added by Bindu as per DE V4 on 15th March 25
'ltcgBreakUp2_New = cgosIncome_cgInc_ltcg_prctg20_secProviso

'Commented by Riyaz on 30/01/2026
'ltcgBreakUp10 = cgosIncome_cgInc_ltcg_PTI20
'ltcgBreakUp1 = cgosIncome_cgInc_ltcg_prctg20_sec112
'ltcgBreakUp4 = cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2
'ltcgBreakUp5 = cgosIncome_cgInc_ltcg_prctg10_sec115AC_1
'ltcgBreakUp6 = cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1
'ltcgBreakUp7 = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3
'ltcgBreakUp8 = cgosIncome_cgInc_ltcg_prctg10_sec115E_b
'ltcgBreakUp3 = cgosIncome_cgInc_ltcg_prctg10_secProviso ''NOt using  commented  by Bindu as this sec changed from 10 to 20
'ltcgBreakUp7_1 = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso
'ltcgBreakUp7_2 = cgosIncome_cgInc_ltcg_prctg10_sec112A
'ltcgBreakUp11 = cgosIncome_cgInc_ltcg_PTI10
'ltcgBreakUp12 = cgosIncome_cgInc_ltcg_PTI10_112
ltcgBreakUp9 = Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4a_DTAARates").value
'---LTCG 12.5% NEWLY ADDED BY BINDU
ltcgBreakUp4_New = cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2
ltcgBreakUp5_New = cgosIncome_cgInc_ltcg_prctg125_sec115AC_1
ltcgBreakUp6_New = cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1
ltcgBreakUp8_New = cgosIncome_cgInc_ltcg_prctg125_sec115E_b
ltcgBreakUp7_1_New = cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso
ltcgBreakUp7_2_New = cgosIncome_cgInc_ltcg_prctg125_sec112A
ltcgBreakUp11_New = cgosIncome_cgInc_ltcg_PTI125
ltcgBreakUp12_New = cgosIncome_cgInc_ltcg_PTI125_112
ltcgBreakUp13_New = cgosIncome_cgInc_ltcg_PTI112_sec48
ltcgBreakUp14_New = cgosIncome_cgInc_ltcg_PTI125_112_B9
ltcgBreakUp15_New = cgosIncome_cgInc_ltcg_prctg125_sec115E_a
ltcgBreakUp16_New = cgosIncome_cgInc_ltcg_prctg125_secProviso '15th March
ltcgBreakUp17_New = cgosIncome_cgInc_ltcg_prctg125_sec115AD_3

 
tempstcgBreakUp1 = stcgBreakUp1
tempstcgBreakUp2 = stcgBreakUp2
'Commented by Riyaz on 30/01/2026
'tempstcgBreakUp3 = stcgBreakUp3
'tempstcgBreakUp4 = stcgBreakUp4
tempstcgBreakUp5 = stcgBreakUp5
tempstcgBreakUp6 = stcgBreakUp6
'Commented by Riyaz on 30/01/2026
'tempstcgBreakUp7 = stcgBreakUp7

'---STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
tempstcgBreakUp3_New = stcgBreakUp3_New
tempstcgBreakUp7_New = stcgBreakUp7_New
tempstcgBreakUp4_New = stcgBreakUp4_New


'Commented by Riyaz on 30/01/2026
'templtcgBreakUp1 = ltcgBreakUp1
'templtcgBreakUp2 = ltcgBreakUp2
''Newly added by Bindu as per DE V4 on 15th March 25
'templtcgBreakUp2_New = ltcgBreakUp2_New

'Commented by Riyaz on 30/01/2026
'templtcgBreakUp3 = ltcgBreakUp3 ''NOt using  commented  by Bindu as this sec changed from 10 to 20
'templtcgBreakUp4 = ltcgBreakUp4
'templtcgBreakUp5 = ltcgBreakUp5
'templtcgBreakUp6 = ltcgBreakUp6
'templtcgBreakUp7 = ltcgBreakUp7
'templtcgBreakUp7_1 = ltcgBreakUp7_1
'templtcgBreakUp7_2 = ltcgBreakUp7_2
'templtcgBreakUp8 = ltcgBreakUp8
templtcgBreakUp9 = ltcgBreakUp9
'templtcgBreakUp10 = ltcgBreakUp10
'templtcgBreakUp11 = ltcgBreakUp11
'templtcgBreakUp12 = ltcgBreakUp12
'---LTCG 12.5% NEWLY ADDED BY BINDU 26TH FEB 25
templtcgBreakUp4_New = ltcgBreakUp4_New
templtcgBreakUp5_New = ltcgBreakUp5_New
templtcgBreakUp6_New = ltcgBreakUp6_New
templtcgBreakUp8_New = ltcgBreakUp8_New
templtcgBreakUp7_1_New = ltcgBreakUp7_1_New
templtcgBreakUp7_2_New = ltcgBreakUp7_2_New
templtcgBreakUp8_New = ltcgBreakUp8_New
templtcgBreakUp11_New = ltcgBreakUp11_New
templtcgBreakUp12_New = ltcgBreakUp12_New
templtcgBreakUp13_New = ltcgBreakUp13_New
templtcgBreakUp14_New = ltcgBreakUp14_New
templtcgBreakUp15_New = ltcgBreakUp15_New
templtcgBreakUp16_New = ltcgBreakUp16_New '15th march
templtcgBreakUp17_New = ltcgBreakUp17_New
 
 tempLtcgCFL = Sheet17.Range("totofbfloss.LTCGLossCF8").value
 tempstcgCFL = Sheet17.Range("totofbfloss.STCGLossCF8").value

'stage-1 setting off LTCG_CFL
'setting with templtcgBreakUp1
        
'--20% LTCG
'Commented by Riyaz on 30/01/2026

'If (tempLtcgCFL > templtcgBreakUp2) Then 'cgosIncome_cgInc_ltcg_prctg20_sec11EA'Not using
'      tempLtcgCFL = tempLtcgCFL - templtcgBreakUp2
'      templtcgBreakUp2 = 0
'Else
'    templtcgBreakUp2 = Application.WorksheetFunction.Max(0, templtcgBreakUp2 - tempLtcgCFL)
'    tempLtcgCFL = 0
'End If
'
'If (tempLtcgCFL > templtcgBreakUp10) Then 'cgosIncome_cgInc_ltcg_PTI20
'      tempLtcgCFL = tempLtcgCFL - templtcgBreakUp10
'      templtcgBreakUp10 = 0
'Else
'    templtcgBreakUp10 = Application.WorksheetFunction.Max(0, templtcgBreakUp10 - tempLtcgCFL)
'    tempLtcgCFL = 0
'End If
'
''Newly added by Bindu as per DE V4 on 15th March 25
'If (tempLtcgCFL > templtcgBreakUp2_New) Then 'cgosIncome_cgInc_ltcg_prctg20_secProviso
'      tempLtcgCFL = tempLtcgCFL - templtcgBreakUp2_New
'      templtcgBreakUp2_New = 0
'Else
'    templtcgBreakUp2_New = Application.WorksheetFunction.Max(0, templtcgBreakUp2_New - tempLtcgCFL)
'    tempLtcgCFL = 0
'End If
''--
'
'If (tempLtcgCFL > templtcgBreakUp1) Then
'       tempLtcgCFL = tempLtcgCFL - templtcgBreakUp1 'cgosIncome_cgInc_ltcg_prctg20_sec112
'       templtcgBreakUp1 = 0
'Else
'       templtcgBreakUp1 = Application.WorksheetFunction.Max(0, templtcgBreakUp1 - tempLtcgCFL)
'            tempLtcgCFL = 0
'End If

'--LTCG 12.5% NEWLY ADDED BY BINDU
'--Not using
 If (tempLtcgCFL > templtcgBreakUp15_New) Then
        tempLtcgCFL = tempLtcgCFL - templtcgBreakUp15_New
        templtcgBreakUp15_New = 0
Else
        templtcgBreakUp15_New = Application.WorksheetFunction.Max(0, templtcgBreakUp15_New - tempLtcgCFL)
        tempLtcgCFL = 0
End If

'Not using
If (tempLtcgCFL > templtcgBreakUp13_New) Then
        tempLtcgCFL = tempLtcgCFL - templtcgBreakUp13_New
        templtcgBreakUp13_New = 0
Else
        templtcgBreakUp13_New = Application.WorksheetFunction.Max(0, templtcgBreakUp13_New - tempLtcgCFL)
        tempLtcgCFL = 0
End If

'--10th
 If (tempLtcgCFL > templtcgBreakUp14_New) Then
        tempLtcgCFL = tempLtcgCFL - templtcgBreakUp14_New
        templtcgBreakUp14_New = 0
Else
        templtcgBreakUp14_New = Application.WorksheetFunction.Max(0, templtcgBreakUp14_New - tempLtcgCFL)
        tempLtcgCFL = 0
End If
'--9th
If (tempLtcgCFL > templtcgBreakUp4_New) Then
     tempLtcgCFL = tempLtcgCFL - templtcgBreakUp4_New
     templtcgBreakUp4_New = 0
Else
     templtcgBreakUp4_New = Application.WorksheetFunction.Max(0, templtcgBreakUp4_New - tempLtcgCFL)
     tempLtcgCFL = 0
End If
'--8th
If (tempLtcgCFL > templtcgBreakUp5_New) Then
      tempLtcgCFL = tempLtcgCFL - templtcgBreakUp5_New
     templtcgBreakUp5_New = 0
Else
      templtcgBreakUp5_New = Application.WorksheetFunction.Max(0, templtcgBreakUp5_New - tempLtcgCFL)
      tempLtcgCFL = 0
End If
'--7th
If (tempLtcgCFL > templtcgBreakUp6_New) Then
       tempLtcgCFL = tempLtcgCFL - templtcgBreakUp6_New
       templtcgBreakUp6_New = 0
Else
     templtcgBreakUp6_New = Application.WorksheetFunction.Max(0, templtcgBreakUp6_New - tempLtcgCFL)
     tempLtcgCFL = 0
End If

'-6th
If (tempLtcgCFL > templtcgBreakUp7_1_New) Then
            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp7_1_New
            templtcgBreakUp7_1_New = 0
Else
      templtcgBreakUp7_1_New = Application.WorksheetFunction.Max(0, templtcgBreakUp7_1_New - tempLtcgCFL)
            tempLtcgCFL = 0
End If
'Added by Riyaz on 03/03/2026
If (tempLtcgCFL > templtcgBreakUp17_New) Then
            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp17_New
            templtcgBreakUp17_New = 0
Else
      templtcgBreakUp17_New = Application.WorksheetFunction.Max(0, templtcgBreakUp17_New - tempLtcgCFL)
            tempLtcgCFL = 0
End If

'--5th
If (tempLtcgCFL > templtcgBreakUp8_New) Then
            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp8_New
            templtcgBreakUp8_New = 0
Else
      templtcgBreakUp8_New = Application.WorksheetFunction.Max(0, templtcgBreakUp8_New - tempLtcgCFL)
            tempLtcgCFL = 0
End If

'--4th
If (tempLtcgCFL > templtcgBreakUp11_New) Then
         tempLtcgCFL = tempLtcgCFL - templtcgBreakUp11_New
         templtcgBreakUp11_New = 0
Else
         templtcgBreakUp11_New = Application.WorksheetFunction.Max(0, templtcgBreakUp11_New - tempLtcgCFL)
         tempLtcgCFL = 0
End If
'--3rd '--15th March
 If (tempLtcgCFL > templtcgBreakUp16_New) Then
        tempLtcgCFL = tempLtcgCFL - templtcgBreakUp16_New
        templtcgBreakUp16_New = 0
Else
        templtcgBreakUp16_New = Application.WorksheetFunction.Max(0, templtcgBreakUp16_New - tempLtcgCFL)
        tempLtcgCFL = 0
End If
'--2nd
If (tempLtcgCFL > templtcgBreakUp12_New) Then
        tempLtcgCFL = tempLtcgCFL - templtcgBreakUp12_New
        templtcgBreakUp12_New = 0
Else
        templtcgBreakUp12_New = Application.WorksheetFunction.Max(0, templtcgBreakUp12_New - tempLtcgCFL)
        tempLtcgCFL = 0
End If
 '--1st
If (tempLtcgCFL > templtcgBreakUp7_2_New) Then
         tempLtcgCFL = tempLtcgCFL - templtcgBreakUp7_2_New
         templtcgBreakUp7_2_New = 0
Else
         templtcgBreakUp7_2_New = Application.WorksheetFunction.Max(0, templtcgBreakUp7_2_New - tempLtcgCFL)
         tempLtcgCFL = 0
End If
'--END

'Commented by Riyaz on 30/01/2026
'
'If (tempLtcgCFL > templtcgBreakUp4) Then
'     tempLtcgCFL = tempLtcgCFL - templtcgBreakUp4
'     templtcgBreakUp4 = 0
'Else
'     templtcgBreakUp4 = Application.WorksheetFunction.Max(0, templtcgBreakUp4 - tempLtcgCFL)
'     tempLtcgCFL = 0
'End If
'
'If (tempLtcgCFL > templtcgBreakUp5) Then
'      tempLtcgCFL = tempLtcgCFL - templtcgBreakUp5
'     templtcgBreakUp5 = 0
'Else
'      templtcgBreakUp5 = Application.WorksheetFunction.Max(0, templtcgBreakUp5 - tempLtcgCFL)
'      tempLtcgCFL = 0
'End If
'
'If (tempLtcgCFL > templtcgBreakUp6) Then
'       tempLtcgCFL = tempLtcgCFL - templtcgBreakUp6
'       templtcgBreakUp6 = 0
'Else
'     templtcgBreakUp6 = Application.WorksheetFunction.Max(0, templtcgBreakUp6 - tempLtcgCFL)
'     tempLtcgCFL = 0
'End If
'
'If (tempLtcgCFL > templtcgBreakUp7) Then
'            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp7
'            templtcgBreakUp7 = 0
'Else
'      templtcgBreakUp7 = Application.WorksheetFunction.Max(0, templtcgBreakUp7 - tempLtcgCFL)
'            tempLtcgCFL = 0
'End If
'
''new
'If (tempLtcgCFL > templtcgBreakUp7_1) Then
'            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp7_1
'            templtcgBreakUp7_1 = 0
'Else
'      templtcgBreakUp7_1 = Application.WorksheetFunction.Max(0, templtcgBreakUp7_1 - tempLtcgCFL)
'            tempLtcgCFL = 0
'End If
'
''end
'
'If (tempLtcgCFL > templtcgBreakUp8) Then
'         tempLtcgCFL = tempLtcgCFL - templtcgBreakUp8
'         templtcgBreakUp8 = 0
'Else
'         templtcgBreakUp8 = Application.WorksheetFunction.Max(0, templtcgBreakUp8 - tempLtcgCFL)
'         tempLtcgCFL = 0
'End If
'
'If (tempLtcgCFL > templtcgBreakUp11) Then
'         tempLtcgCFL = tempLtcgCFL - templtcgBreakUp11
'         templtcgBreakUp11 = 0
'Else
'         templtcgBreakUp11 = Application.WorksheetFunction.Max(0, templtcgBreakUp11 - tempLtcgCFL)
'         tempLtcgCFL = 0
'End If
'
''NOt using  commented  by Bindu as this sec changed from 10 to 20
'' If (tempLtcgCFL > templtcgBreakUp3) Then
''        tempLtcgCFL = tempLtcgCFL - templtcgBreakUp3
''        templtcgBreakUp3 = 0
''Else
''        templtcgBreakUp3 = Application.WorksheetFunction.Max(0, templtcgBreakUp3 - tempLtcgCFL)
''        tempLtcgCFL = 0
''End If
'
'If (tempLtcgCFL > templtcgBreakUp12) Then
'        tempLtcgCFL = tempLtcgCFL - templtcgBreakUp12
'        templtcgBreakUp12 = 0
'Else
'        templtcgBreakUp12 = Application.WorksheetFunction.Max(0, templtcgBreakUp12 - tempLtcgCFL)
'        tempLtcgCFL = 0
'End If
'
''new
'If (tempLtcgCFL > templtcgBreakUp7_2) Then
'        tempLtcgCFL = tempLtcgCFL - templtcgBreakUp7_2
'        templtcgBreakUp7_2 = 0
'Else
'        templtcgBreakUp7_2 = Application.WorksheetFunction.Max(0, templtcgBreakUp7_2 - tempLtcgCFL)
'        tempLtcgCFL = 0
'End If
''end
'

'SIT-115946 uncommented by sadineni on 31/03/2026
If (tempLtcgCFL > templtcgBreakUp9) Then
        tempLtcgCFL = tempLtcgCFL - templtcgBreakUp9
        templtcgBreakUp9 = 0
Else
        templtcgBreakUp9 = Application.WorksheetFunction.Max(0, templtcgBreakUp9 - tempLtcgCFL)
        tempLtcgCFL = 0
End If
'******************************************************************************

'Sheet17.Unprotect Password:=getmsgstate
'Application.WorksheetFunction.Max(0, ltcgBreakUp3 - templtcgBreakUp3) ''Deleted  by Bindu as this sec changed from 10 to 20

'ltcgBreakUp2_New/ltcgBreakUp16_New 'Newly added by Bindu as per DE V4 on 15th March 25
'Sheet17.Range("adjtotloss.LTCGLossCF9").value = Application.WorksheetFunction.Max(0, ltcgBreakUp1 - templtcgBreakUp1) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp2 - templtcgBreakUp2) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp2_New - templtcgBreakUp2_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp4 - templtcgBreakUp4) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp5 - templtcgBreakUp5) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp6 - templtcgBreakUp6) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp7 - templtcgBreakUp7) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp7_1 - templtcgBreakUp7_1) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp7_2 - templtcgBreakUp7_2) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp8 - templtcgBreakUp8) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp9 - templtcgBreakUp9) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp10 - templtcgBreakUp10) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp11 - templtcgBreakUp11) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp12 - templtcgBreakUp12) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp4_New - templtcgBreakUp4_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp5_New - templtcgBreakUp5_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp6_New - templtcgBreakUp6_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp8_New - templtcgBreakUp8_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp7_1_New - templtcgBreakUp7_1_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp7_2_New - templtcgBreakUp7_2_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp11_New - templtcgBreakUp11_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp12_New - templtcgBreakUp12_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp13_New - templtcgBreakUp13_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp14_New - templtcgBreakUp14_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp15_New - templtcgBreakUp15_New) + Application.WorksheetFunction.Max(0, ltcgBreakUp16_New - templtcgBreakUp16_New)

Sheet17.Range("adjtotloss.LTCGLossCF9").value = Application.WorksheetFunction.Max(0, ltcgBreakUp9 - templtcgBreakUp9) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp4_New - templtcgBreakUp4_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp5_New - templtcgBreakUp5_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp6_New - templtcgBreakUp6_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp8_New - templtcgBreakUp8_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp7_1_New - templtcgBreakUp7_1_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp7_2_New - templtcgBreakUp7_2_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp11_New - templtcgBreakUp11_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp12_New - templtcgBreakUp12_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp13_New - templtcgBreakUp13_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp14_New - templtcgBreakUp14_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp15_New - templtcgBreakUp15_New) + _
                                               Application.WorksheetFunction.Max(0, ltcgBreakUp16_New - templtcgBreakUp16_New) + Application.WorksheetFunction.Max(0, ltcgBreakUp17_New - templtcgBreakUp17_New)

Dim storeSTCG1, storeSTCG2, storeSTCG3, storeSTCG4, storeSTCG5, storeSTCG6, storeSTCG7 As Double
'---STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
Dim storeSTCG3_New, storeSTCG7_New, storeSTCG4_New As Double

'storeLTCG2_New Newly added by Bindu
Dim storeLTCG1, storeLTCG2, storeLTCG2_New, storeLTCG3, storeLTCG4, storeLTCG5, storeLTCG6, storeLTCG7, storeLTCG7_1, storeLTCG7_2, storeLTCG8, storeLTCG9, storeLTCG10, storeLTCG11, storeLTCG12 As Double
'---LTCG 12.5% NEWLY ADDED BY BINDU 26TH FEB 25
Dim storeLTCG4_New, storeLTCG5_New, storeLTCG6_New, storeLTCG8_New, storeLTCG7_1_New, storeLTCG7_2_New, storeLTCG11_New, storeLTCG12_New, storeLTCG13_New, storeLTCG14_New, storeLTCG15_New, storeLTCG16_New, storeLTCG17_New As Double



storeSTCG1 = tempstcgBreakUp1
storeSTCG2 = tempstcgBreakUp2
'Commented by Riyaz on 30/01/2026
'storeSTCG3 = tempstcgBreakUp3
'storeSTCG4 = tempstcgBreakUp4
storeSTCG5 = tempstcgBreakUp5
storeSTCG6 = tempstcgBreakUp6
'Commented by Riyaz on 30/01/2026
'storeSTCG7 = tempstcgBreakUp7

'---STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
storeSTCG3_New = tempstcgBreakUp3_New
storeSTCG7_New = tempstcgBreakUp7_New
storeSTCG4_New = tempstcgBreakUp4_New



'Commented by Riyaz on 30/01/2026
'storeLTCG1 = templtcgBreakUp1
'storeLTCG2 = templtcgBreakUp2

'Newly added by Bindu as per DE V4 on 15th March 25
'storeLTCG2_New = templtcgBreakUp2_New


'storeLTCG3 = templtcgBreakUp3 ''NOt using  commented  by Bindu as this sec changed from 10 to 20
'storeLTCG4 = templtcgBreakUp4
'storeLTCG5 = templtcgBreakUp5
'storeLTCG6 = templtcgBreakUp6
'storeLTCG7 = templtcgBreakUp7
'storeLTCG8 = templtcgBreakUp8
'storeLTCG9 = templtcgBreakUp9
'storeLTCG10 = templtcgBreakUp10
'storeLTCG11 = templtcgBreakUp11
'storeLTCG12 = templtcgBreakUp12
        
'storeLTCG7_1 = templtcgBreakUp7_1
'storeLTCG7_2 = templtcgBreakUp7_2
'---LTCG 12.5% NEWLY ADDED BY BINDU 26TH FEB 25

storeLTCG4_New = templtcgBreakUp4_New
storeLTCG5_New = templtcgBreakUp5_New
storeLTCG6_New = templtcgBreakUp6_New
storeLTCG8_New = templtcgBreakUp8_New
storeLTCG7_1_New = templtcgBreakUp7_1_New
storeLTCG7_2_New = templtcgBreakUp7_2_New
storeLTCG11_New = templtcgBreakUp11_New
storeLTCG12_New = templtcgBreakUp12_New
storeLTCG13_New = templtcgBreakUp13_New
storeLTCG14_New = templtcgBreakUp14_New
storeLTCG15_New = templtcgBreakUp15_New
storeLTCG16_New = templtcgBreakUp16_New '15th March
storeLTCG17_New = templtcgBreakUp17_New

'SCTCG SetOff
'--115AD 30% STCG
If (tempstcgCFL > tempstcgBreakUp1) Then
            tempstcgCFL = tempstcgCFL - tempstcgBreakUp1
            tempstcgBreakUp1 = 0
Else
            tempstcgBreakUp1 = Application.WorksheetFunction.Max(0, tempstcgBreakUp1 - tempstcgCFL)
            tempstcgCFL = 0
End If

'--PTI 30% STCG
If (tempstcgCFL > tempstcgBreakUp6) Then
            tempstcgCFL = tempstcgCFL - tempstcgBreakUp6
            tempstcgBreakUp6 = 0
Else
            tempstcgBreakUp6 = Application.WorksheetFunction.Max(0, tempstcgBreakUp6 - tempstcgCFL)
            tempstcgCFL = 0
End If
        
'--APP STCG
If (tempstcgCFL > tempstcgBreakUp2) Then
            tempstcgCFL = tempstcgCFL - tempstcgBreakUp2
            tempstcgBreakUp2 = 0
Else
            tempstcgBreakUp2 = Application.WorksheetFunction.Max(0, tempstcgBreakUp2 - tempstcgCFL)
            tempstcgCFL = 0
End If

'---STCG 20% NEWLY ADDED BY BINDU 19TH FEB 25
'--115AD STCG 20%
If (tempstcgCFL > tempstcgBreakUp3_New) Then
            tempstcgCFL = tempstcgCFL - tempstcgBreakUp3_New
            tempstcgBreakUp3_New = 0
Else
            tempstcgBreakUp3_New = Application.WorksheetFunction.Max(0, tempstcgBreakUp3_New - tempstcgCFL)
            tempstcgCFL = 0
End If
'--PTI STCG 20%
If (tempstcgCFL > tempstcgBreakUp7_New) Then
            tempstcgCFL = tempstcgCFL - tempstcgBreakUp7_New
            tempstcgBreakUp7_New = 0
Else
            tempstcgBreakUp7_New = Application.WorksheetFunction.Max(0, tempstcgBreakUp7_New - tempstcgCFL)
            tempstcgCFL = 0
End If
'--111A STCG 20%
If (tempstcgCFL > tempstcgBreakUp4_New) Then
            tempstcgCFL = tempstcgCFL - tempstcgBreakUp4_New
            tempstcgBreakUp4_New = 0
Else
            tempstcgBreakUp4_New = Application.WorksheetFunction.Max(0, tempstcgBreakUp4_New - tempstcgCFL)
            tempstcgCFL = 0
End If
'---ENDS

'Commented by Riyaz on 30/01/2026
'If (tempstcgCFL > templtcgBreakUp2) Then
'            tempstcgCFL = tempstcgCFL - templtcgBreakUp2
'            templtcgBreakUp2 = 0
'Else
'            templtcgBreakUp2 = Application.WorksheetFunction.Max(0, templtcgBreakUp2 - tempstcgCFL)
'            tempstcgCFL = 0
'End If
'
'
'If (tempstcgCFL > templtcgBreakUp10) Then
'            tempstcgCFL = tempstcgCFL - templtcgBreakUp10
'            templtcgBreakUp10 = 0
'Else
'            templtcgBreakUp10 = Application.WorksheetFunction.Max(0, templtcgBreakUp10 - tempstcgCFL)
'            tempstcgCFL = 0
'End If
'
''--'Newly added by Bindu as per DE V4 on 15th March 25
'If (tempstcgCFL > templtcgBreakUp2_New) Then
'            tempstcgCFL = tempstcgCFL - templtcgBreakUp2_New
'            templtcgBreakUp2_New = 0
'Else
'            templtcgBreakUp2_New = Application.WorksheetFunction.Max(0, templtcgBreakUp2_New - tempstcgCFL)
'            tempstcgCFL = 0
'End If
''--
'
'If (tempstcgCFL > templtcgBreakUp1) Then
'            tempstcgCFL = tempstcgCFL - templtcgBreakUp1
'            templtcgBreakUp1 = 0
'Else
'            templtcgBreakUp1 = Application.WorksheetFunction.Max(0, templtcgBreakUp1 - tempstcgCFL)
'            tempstcgCFL = 0
'End If
        
'Commented by Riyaz on 30/01/2026
'If (tempstcgCFL > tempstcgBreakUp3) Then
'            tempstcgCFL = tempstcgCFL - tempstcgBreakUp3
'            tempstcgBreakUp3 = 0
'Else
'            tempstcgBreakUp3 = Application.WorksheetFunction.Max(0, tempstcgBreakUp3 - tempstcgCFL)
'            tempstcgCFL = 0
'End If
'
'If (tempstcgCFL > tempstcgBreakUp7) Then
'            tempstcgCFL = tempstcgCFL - tempstcgBreakUp7
'            tempstcgBreakUp7 = 0
'Else
'            tempstcgBreakUp7 = Application.WorksheetFunction.Max(0, tempstcgBreakUp7 - tempstcgCFL)
'            tempstcgCFL = 0
'End If
'
'If (tempstcgCFL > tempstcgBreakUp4) Then
'            tempstcgCFL = tempstcgCFL - tempstcgBreakUp4
'            tempstcgBreakUp4 = 0
'Else
'            tempstcgBreakUp4 = Application.WorksheetFunction.Max(0, tempstcgBreakUp4 - tempstcgCFL)
'            tempstcgCFL = 0
'End If

'--LTCG 12.5% NEWLY ADDED BY BINDU
'--Not using
 If (tempstcgCFL > templtcgBreakUp15_New) Then
        tempstcgCFL = tempstcgCFL - templtcgBreakUp15_New
        templtcgBreakUp15_New = 0
Else
        templtcgBreakUp15_New = Application.WorksheetFunction.Max(0, templtcgBreakUp15_New - tempstcgCFL)
        tempstcgCFL = 0
End If

'Not using
If (tempstcgCFL > templtcgBreakUp13_New) Then
        tempstcgCFL = tempstcgCFL - templtcgBreakUp13_New
        templtcgBreakUp13_New = 0
Else
        templtcgBreakUp13_New = Application.WorksheetFunction.Max(0, templtcgBreakUp13_New - tempstcgCFL)
        tempstcgCFL = 0
End If

'--10th
 If (tempstcgCFL > templtcgBreakUp14_New) Then
        tempstcgCFL = tempstcgCFL - templtcgBreakUp14_New
        templtcgBreakUp14_New = 0
Else
        templtcgBreakUp14_New = Application.WorksheetFunction.Max(0, templtcgBreakUp14_New - tempstcgCFL)
        tempstcgCFL = 0
End If
'--9th
If (tempstcgCFL > templtcgBreakUp4_New) Then
     tempstcgCFL = tempstcgCFL - templtcgBreakUp4_New
     templtcgBreakUp4_New = 0
Else
     templtcgBreakUp4_New = Application.WorksheetFunction.Max(0, templtcgBreakUp4_New - tempstcgCFL)
     tempstcgCFL = 0
End If
'--8th
If (tempstcgCFL > templtcgBreakUp5_New) Then
      tempstcgCFL = tempstcgCFL - templtcgBreakUp5_New
     templtcgBreakUp5_New = 0
Else
      templtcgBreakUp5_New = Application.WorksheetFunction.Max(0, templtcgBreakUp5_New - tempstcgCFL)
      tempstcgCFL = 0
End If
'--7th
If (tempstcgCFL > templtcgBreakUp6_New) Then
       tempstcgCFL = tempstcgCFL - templtcgBreakUp6_New
       templtcgBreakUp6_New = 0
Else
     templtcgBreakUp6_New = Application.WorksheetFunction.Max(0, templtcgBreakUp6_New - tempstcgCFL)
     tempstcgCFL = 0
End If
'Added by Riyaz on 03/03/2026
If (tempstcgCFL > templtcgBreakUp17_New) Then
       tempstcgCFL = tempstcgCFL - templtcgBreakUp17_New
       templtcgBreakUp17_New = 0
Else
     templtcgBreakUp17_New = Application.WorksheetFunction.Max(0, templtcgBreakUp17_New - tempstcgCFL)
     tempstcgCFL = 0
End If

'-6th
If (tempstcgCFL > templtcgBreakUp7_1_New) Then
            tempstcgCFL = tempstcgCFL - templtcgBreakUp7_1_New
            templtcgBreakUp7_1_New = 0
Else
      templtcgBreakUp7_1_New = Application.WorksheetFunction.Max(0, templtcgBreakUp7_1_New - tempstcgCFL)
            tempstcgCFL = 0
End If
'--5th
If (tempstcgCFL > templtcgBreakUp8_New) Then
            tempstcgCFL = tempstcgCFL - templtcgBreakUp8_New
            templtcgBreakUp8_New = 0
Else
      templtcgBreakUp8_New = Application.WorksheetFunction.Max(0, templtcgBreakUp8_New - tempstcgCFL)
            tempstcgCFL = 0
End If

'--4th
If (tempstcgCFL > templtcgBreakUp11_New) Then
         tempstcgCFL = tempstcgCFL - templtcgBreakUp11_New
         templtcgBreakUp11_New = 0
Else
         templtcgBreakUp11_New = Application.WorksheetFunction.Max(0, templtcgBreakUp11_New - tempstcgCFL)
         tempstcgCFL = 0
End If
'--3rd '--15th March
 If (tempstcgCFL > templtcgBreakUp16_New) Then
        tempstcgCFL = tempstcgCFL - templtcgBreakUp16_New
        templtcgBreakUp16_New = 0
Else
        templtcgBreakUp16_New = Application.WorksheetFunction.Max(0, templtcgBreakUp16_New - tempstcgCFL)
        tempstcgCFL = 0
End If
'--2nd
If (tempstcgCFL > templtcgBreakUp12_New) Then
        tempstcgCFL = tempstcgCFL - templtcgBreakUp12_New
        templtcgBreakUp12_New = 0
Else
        templtcgBreakUp12_New = Application.WorksheetFunction.Max(0, templtcgBreakUp12_New - tempstcgCFL)
        tempstcgCFL = 0
End If
 '--1st
If (tempstcgCFL > templtcgBreakUp7_2_New) Then
         tempstcgCFL = tempstcgCFL - templtcgBreakUp7_2_New
         templtcgBreakUp7_2_New = 0
Else
         templtcgBreakUp7_2_New = Application.WorksheetFunction.Max(0, templtcgBreakUp7_2_New - tempstcgCFL)
         tempstcgCFL = 0
End If
'--END

'Commented by Riyaz on 30/01/2026

' If (tempstcgCFL > templtcgBreakUp4) Then
'            tempstcgCFL = tempstcgCFL - templtcgBreakUp4
'            templtcgBreakUp4 = 0
'Else
'            templtcgBreakUp4 = Application.WorksheetFunction.Max(0, templtcgBreakUp4 - tempstcgCFL)
'            tempstcgCFL = 0
'End If
'If (tempstcgCFL > templtcgBreakUp5) Then
'            tempstcgCFL = tempstcgCFL - templtcgBreakUp5
'            templtcgBreakUp5 = 0
'Else
'            templtcgBreakUp5 = Application.WorksheetFunction.Max(0, templtcgBreakUp5 - tempstcgCFL)
'            tempstcgCFL = 0
'End If
'
'If (tempstcgCFL > templtcgBreakUp6) Then
'            tempstcgCFL = tempstcgCFL - templtcgBreakUp6
'            templtcgBreakUp6 = 0
'Else
'            templtcgBreakUp6 = Application.WorksheetFunction.Max(0, templtcgBreakUp6 - tempstcgCFL)
'            tempstcgCFL = 0
'End If
'
'If (tempstcgCFL > templtcgBreakUp7) Then
'            tempstcgCFL = tempstcgCFL - templtcgBreakUp7
'            templtcgBreakUp7 = 0
'Else
'            templtcgBreakUp7 = Application.WorksheetFunction.Max(0, templtcgBreakUp7 - tempstcgCFL)
'            tempstcgCFL = 0
'End If
'
''new
'If (tempstcgCFL > templtcgBreakUp7_1) Then
'            tempstcgCFL = tempstcgCFL - templtcgBreakUp7_1
'            templtcgBreakUp7_1 = 0
'Else
'            templtcgBreakUp7_1 = Application.WorksheetFunction.Max(0, templtcgBreakUp7_1 - tempstcgCFL)
'            tempstcgCFL = 0
'End If
''end
'
'If (tempstcgCFL > templtcgBreakUp8) Then
'            tempstcgCFL = tempstcgCFL - templtcgBreakUp8
'            templtcgBreakUp8 = 0
'Else
'            templtcgBreakUp8 = Application.WorksheetFunction.Max(0, templtcgBreakUp8 - tempstcgCFL)
'            tempstcgCFL = 0
'End If
'
'If (tempstcgCFL > templtcgBreakUp11) Then
'            tempstcgCFL = tempstcgCFL - templtcgBreakUp11
'            templtcgBreakUp11 = 0
'Else
'            templtcgBreakUp11 = Application.WorksheetFunction.Max(0, templtcgBreakUp11 - tempstcgCFL)
'            tempstcgCFL = 0
'End If
'
        'NOt using  commented  by Bindu as this sec changed from 10 to 20
'If (tempstcgCFL > templtcgBreakUp3) Then
'            tempstcgCFL = tempstcgCFL - templtcgBreakUp3
'            templtcgBreakUp3 = 0
'Else
'            templtcgBreakUp3 = Application.WorksheetFunction.Max(0, templtcgBreakUp3 - tempstcgCFL)
'            tempstcgCFL = 0
'End If
'Commented by Riyaz on 30/01/2026
'If (tempstcgCFL > templtcgBreakUp12) Then
'            tempstcgCFL = tempstcgCFL - templtcgBreakUp12
'            templtcgBreakUp12 = 0
'Else
'            templtcgBreakUp12 = Application.WorksheetFunction.Max(0, templtcgBreakUp12 - tempstcgCFL)
'            tempstcgCFL = 0
'End If
'
''new
'If (tempstcgCFL > templtcgBreakUp7_2) Then
'            tempstcgCFL = tempstcgCFL - templtcgBreakUp7_2
'            templtcgBreakUp7_2 = 0
'Else
'            templtcgBreakUp7_2 = Application.WorksheetFunction.Max(0, templtcgBreakUp7_2 - tempstcgCFL)
'            tempstcgCFL = 0
'
' End If

'SIT-115946 uncommented by sadineni on 31/03/2026
'********************************************************************************************************
If (tempstcgCFL > tempstcgBreakUp5) Then
            tempstcgCFL = tempstcgCFL - tempstcgBreakUp5
            tempstcgBreakUp5 = 0
Else
            tempstcgBreakUp5 = Application.WorksheetFunction.Max(0, tempstcgBreakUp5 - tempstcgCFL)
            tempstcgCFL = 0
End If
'
If (tempstcgCFL > templtcgBreakUp9) Then
            tempstcgCFL = tempstcgCFL - templtcgBreakUp9
            templtcgBreakUp9 = 0
Else
            templtcgBreakUp9 = Application.WorksheetFunction.Max(0, templtcgBreakUp9 - tempstcgCFL)
            tempstcgCFL = 0

End If
'***********************************************************************************************************

''end
'
'
'
' If (tempLtcgCFL > templtcgBreakUp2) Then
'            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp2
'            templtcgBreakUp2 = 0
'Else
'            templtcgBreakUp2 = Application.WorksheetFunction.Max(0, templtcgBreakUp2 - tempLtcgCFL)
'            tempLtcgCFL = 0
'End If
'
' If (tempLtcgCFL > templtcgBreakUp10) Then
'            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp10
'            templtcgBreakUp10 = 0
'Else
'            templtcgBreakUp10 = Application.WorksheetFunction.Max(0, templtcgBreakUp10 - tempLtcgCFL)
'            tempLtcgCFL = 0
'End If
'
''--'Newly added by Bindu as per DE V4 on 15th March 25
' If (tempLtcgCFL > templtcgBreakUp2_New) Then
'            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp2_New
'            templtcgBreakUp2_New = 0
'Else
'            templtcgBreakUp2_New = Application.WorksheetFunction.Max(0, templtcgBreakUp2_New - tempLtcgCFL)
'            tempLtcgCFL = 0
'End If
''--
'
'
' If (tempLtcgCFL > templtcgBreakUp1) Then
'            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp1
'            templtcgBreakUp1 = 0
'Else
'            templtcgBreakUp1 = Application.WorksheetFunction.Max(0, templtcgBreakUp1 - tempLtcgCFL)
'            tempLtcgCFL = 0
'End If

'--LTCG 12.5% NEWLY ADDED BY BINDU
'--Not using
 If (tempLtcgCFL > templtcgBreakUp15_New) Then
        tempLtcgCFL = tempLtcgCFL - templtcgBreakUp15_New
        templtcgBreakUp15_New = 0
Else
        templtcgBreakUp15_New = Application.WorksheetFunction.Max(0, templtcgBreakUp15_New - tempLtcgCFL)
        tempLtcgCFL = 0
End If

'Not using
If (tempLtcgCFL > templtcgBreakUp13_New) Then
        tempLtcgCFL = tempLtcgCFL - templtcgBreakUp13_New
        templtcgBreakUp13_New = 0
Else
        templtcgBreakUp13_New = Application.WorksheetFunction.Max(0, templtcgBreakUp13_New - tempLtcgCFL)
        tempLtcgCFL = 0
End If

'--10th
 If (tempLtcgCFL > templtcgBreakUp14_New) Then
        tempLtcgCFL = tempLtcgCFL - templtcgBreakUp14_New
        templtcgBreakUp14_New = 0
Else
        templtcgBreakUp14_New = Application.WorksheetFunction.Max(0, templtcgBreakUp14_New - tempLtcgCFL)
        tempLtcgCFL = 0
End If
'--9th
If (tempLtcgCFL > templtcgBreakUp4_New) Then
     tempLtcgCFL = tempLtcgCFL - templtcgBreakUp4_New
     templtcgBreakUp4_New = 0
Else
     templtcgBreakUp4_New = Application.WorksheetFunction.Max(0, templtcgBreakUp4_New - tempLtcgCFL)
     tempLtcgCFL = 0
End If
'--8th
If (tempLtcgCFL > templtcgBreakUp5_New) Then
      tempLtcgCFL = tempLtcgCFL - templtcgBreakUp5_New
     templtcgBreakUp5_New = 0
Else
      templtcgBreakUp5_New = Application.WorksheetFunction.Max(0, templtcgBreakUp5_New - tempLtcgCFL)
      tempLtcgCFL = 0
End If
'--7th
If (tempLtcgCFL > templtcgBreakUp6_New) Then
       tempLtcgCFL = tempLtcgCFL - templtcgBreakUp6_New
       templtcgBreakUp6_New = 0
Else
     templtcgBreakUp6_New = Application.WorksheetFunction.Max(0, templtcgBreakUp6_New - tempLtcgCFL)
     tempLtcgCFL = 0
End If

'Added by Riyaz on 03/03/2026
If (tempLtcgCFL > templtcgBreakUp17_New) Then
       tempLtcgCFL = tempLtcgCFL - templtcgBreakUp17_New
       templtcgBreakUp17_New = 0
Else
     templtcgBreakUp17_New = Application.WorksheetFunction.Max(0, templtcgBreakUp17_New - tempLtcgCFL)
     tempLtcgCFL = 0
End If

'-6th
If (tempLtcgCFL > templtcgBreakUp7_1_New) Then
            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp7_1_New
            templtcgBreakUp7_1_New = 0
Else
      templtcgBreakUp7_1_New = Application.WorksheetFunction.Max(0, templtcgBreakUp7_1_New - tempLtcgCFL)
            tempLtcgCFL = 0
End If
'--5th
If (tempLtcgCFL > templtcgBreakUp8_New) Then
            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp8_New
            templtcgBreakUp8_New = 0
Else
      templtcgBreakUp8_New = Application.WorksheetFunction.Max(0, templtcgBreakUp8_New - tempLtcgCFL)
            tempLtcgCFL = 0
End If

'--4th
If (tempLtcgCFL > templtcgBreakUp11_New) Then
         tempLtcgCFL = tempLtcgCFL - templtcgBreakUp11_New
         templtcgBreakUp11_New = 0
Else
         templtcgBreakUp11_New = Application.WorksheetFunction.Max(0, templtcgBreakUp11_New - tempLtcgCFL)
         tempLtcgCFL = 0
End If
'--3rd '--15th March
 If (tempLtcgCFL > templtcgBreakUp16_New) Then
        tempLtcgCFL = tempLtcgCFL - templtcgBreakUp16_New
        templtcgBreakUp16_New = 0
Else
        templtcgBreakUp16_New = Application.WorksheetFunction.Max(0, templtcgBreakUp16_New - tempLtcgCFL)
        tempLtcgCFL = 0
End If
'--2nd
If (tempLtcgCFL > templtcgBreakUp12_New) Then
        tempLtcgCFL = tempLtcgCFL - templtcgBreakUp12_New
        templtcgBreakUp12_New = 0
Else
        templtcgBreakUp12_New = Application.WorksheetFunction.Max(0, templtcgBreakUp12_New - tempLtcgCFL)
        tempLtcgCFL = 0
End If
 '--1st
If (tempLtcgCFL > templtcgBreakUp7_2_New) Then
         tempLtcgCFL = tempLtcgCFL - templtcgBreakUp7_2_New
         templtcgBreakUp7_2_New = 0
Else
         templtcgBreakUp7_2_New = Application.WorksheetFunction.Max(0, templtcgBreakUp7_2_New - tempLtcgCFL)
         tempLtcgCFL = 0
End If
'--END

'Commented by Riyaz on 30/01/2026
'If (tempLtcgCFL > templtcgBreakUp4) Then
'            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp4
'            templtcgBreakUp4 = 0
'Else
'            templtcgBreakUp4 = Application.WorksheetFunction.Max(0, templtcgBreakUp4 - tempLtcgCFL)
'            tempLtcgCFL = 0
'End If
'
'If (tempLtcgCFL > templtcgBreakUp5) Then
'            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp5
'            templtcgBreakUp5 = 0
'Else
'            templtcgBreakUp5 = Application.WorksheetFunction.Max(0, templtcgBreakUp5 - tempLtcgCFL)
'            tempLtcgCFL = 0
'End If
'
'If (tempLtcgCFL > templtcgBreakUp6) Then
'            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp6
'            templtcgBreakUp6 = 0
'Else
'            templtcgBreakUp6 = Application.WorksheetFunction.Max(0, templtcgBreakUp6 - tempLtcgCFL)
'            tempLtcgCFL = 0
'End If
'
'If (tempLtcgCFL > templtcgBreakUp7) Then
'            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp7
'            templtcgBreakUp7 = 0
'Else
'            templtcgBreakUp7 = Application.WorksheetFunction.Max(0, templtcgBreakUp7 - tempLtcgCFL)
'            tempLtcgCFL = 0
'End If
'
''new
'If (tempLtcgCFL > templtcgBreakUp7_1) Then
'            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp7_1
'            templtcgBreakUp7_1 = 0
'Else
'            templtcgBreakUp7_1 = Application.WorksheetFunction.Max(0, templtcgBreakUp7_1 - tempLtcgCFL)
'            tempLtcgCFL = 0
'End If
''end
'
'
'If (tempLtcgCFL > templtcgBreakUp8) Then
'            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp8
'            templtcgBreakUp8 = 0
'Else
'            templtcgBreakUp8 = Application.WorksheetFunction.Max(0, templtcgBreakUp8 - tempLtcgCFL)
'            tempLtcgCFL = 0
'End If
'
'If (tempLtcgCFL > templtcgBreakUp11) Then
'            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp11
'            templtcgBreakUp11 = 0
'Else
'            templtcgBreakUp11 = Application.WorksheetFunction.Max(0, templtcgBreakUp11 - tempLtcgCFL)
'            tempLtcgCFL = 0
'End If
        
 'NOt using  commented  by Bindu as this sec changed from 10 to 20
'If (tempLtcgCFL > templtcgBreakUp3) Then
'            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp3
'            templtcgBreakUp3 = 0
'Else
'            templtcgBreakUp3 = Application.WorksheetFunction.Max(0, templtcgBreakUp3 - tempLtcgCFL)
'            tempLtcgCFL = 0
'End If

'If (tempLtcgCFL > templtcgBreakUp12) Then
'            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp12
'            templtcgBreakUp12 = 0
'Else
'            templtcgBreakUp12 = Application.WorksheetFunction.Max(0, templtcgBreakUp12 - tempLtcgCFL)
'            tempLtcgCFL = 0
'End If
'
''new
'If (tempLtcgCFL > templtcgBreakUp7_2) Then
'            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp7_2
'            templtcgBreakUp7_2 = 0
'Else
'            templtcgBreakUp7_2 = Application.WorksheetFunction.Max(0, templtcgBreakUp7_2 - tempLtcgCFL)
'            tempLtcgCFL = 0
'End If
'
'
'
'SIT-115946 uncommented by sadineni on 31/03/2026
If (tempLtcgCFL > templtcgBreakUp9) Then
            tempLtcgCFL = tempLtcgCFL - templtcgBreakUp9
            templtcgBreakUp9 = 0
Else
            templtcgBreakUp9 = Application.WorksheetFunction.Max(0, templtcgBreakUp9 - tempLtcgCFL)
            tempLtcgCFL = 0
End If

'end
'-----------------
Sheet16.Range("hp.BFlossPrevYrUndSameHeadSetoff1").value = Application.WorksheetFunction.Min(Sheet16.Range("hp.IncOfCurYrUndHeadFromCYLA1").value, Sheet17.Range("totofbfloss.HPLossCF8").value)

Sheet16.Range("rh.BFlossPrevYrUndSameHeadSetoff6").value = Application.WorksheetFunction.Min(Sheet16.Range("rh.IncOfCurYrUndHeadFromCYLA6").value, Sheet17.Range("totofbfloss.OthSrcLossRaceHorseCF8").value)

'Setting off business loss with spec and specified business loss

    IncOfCurYrUndHeadFromCYLA2_16_1 = 0
    BusLossOthThanSpecLossCF8_16_1 = 0
    IncOfCurYrUndHeadFromCYLA2a_16_1 = 0
    LossFrmSpecBusCF8_16_1 = 0
    IncOfCurYrUndHeadFromCYLA2b_16_1 = 0
    LossFrmSpecifiedBusCF8_16_1 = 0

    BFlossPrevYrUndSameHeadSetoff2_16_1 = 0
    BFlossPrevYrUndSameHeadSetoff2a_16_1 = 0
    BFlossPrevYrUndSameHeadSetoff2b_16_1 = 0
    BpPendingLossAmt = 0
    adjustSpecu = 0
    adjustSpeci = 0

    IncOfCurYrUndHeadFromCYLA2_16_1 = Sheet16.Range("busipofincl.IncOfCurYrUndHeadFromCYLA2").value
    BusLossOthThanSpecLossCF8_16_1 = Sheet17.Range("totofbfloss.BusLossOthThanSpecLossCF8").value
    IncOfCurYrUndHeadFromCYLA2a_16_1 = Sheet16.Range("busipofinclspec.IncOfCurYrUndHeadFromCYLA2a").value
    LossFrmSpecBusCF8_16_1 = Sheet17.Range("totofbfloss.LossFrmSpecBusCF8").value
    IncOfCurYrUndHeadFromCYLA2b_16_1 = Sheet16.Range("busipofinclspecified.IncOfCurYrUndHeadFromCYLA2b").value
    LossFrmSpecifiedBusCF8_16_1 = Sheet17.Range("totofbfloss.LossFrmSpecifiedBusCF8").value


    'BP
    BFlossPrevYrUndSameHeadSetoff2_16_1 = WorksheetFunction.Min(IncOfCurYrUndHeadFromCYLA2_16_1 _
                                                                , BusLossOthThanSpecLossCF8_16_1)
    'BP-Specu
    BFlossPrevYrUndSameHeadSetoff2a_16_1 = WorksheetFunction.Min(IncOfCurYrUndHeadFromCYLA2a_16_1 _
                                                                , LossFrmSpecBusCF8_16_1)
    'BP-Speci
    BFlossPrevYrUndSameHeadSetoff2b_16_1 = WorksheetFunction.Min(IncOfCurYrUndHeadFromCYLA2b_16_1 _
                                                                , LossFrmSpecifiedBusCF8_16_1)
    
    If (IIf(IncOfCurYrUndHeadFromCYLA2_16_1 = "", 0, IncOfCurYrUndHeadFromCYLA2_16_1) < _
         IIf(BusLossOthThanSpecLossCF8_16_1 = "", 0, BusLossOthThanSpecLossCF8_16_1) _
        ) Then

        
        BpPendingLossAmt = BusLossOthThanSpecLossCF8_16_1 - IncOfCurYrUndHeadFromCYLA2_16_1
            
        If IncOfCurYrUndHeadFromCYLA2a_16_1 - BFlossPrevYrUndSameHeadSetoff2a_16_1 > 0 Then
            adjustSpecu = WorksheetFunction.Min(BpPendingLossAmt, IncOfCurYrUndHeadFromCYLA2a_16_1 - BFlossPrevYrUndSameHeadSetoff2a_16_1)
            BFlossPrevYrUndSameHeadSetoff2a_16_1 = BFlossPrevYrUndSameHeadSetoff2a_16_1 + adjustSpecu
        End If
            
        If BpPendingLossAmt - adjustSpecu > 0 Then
            If IncOfCurYrUndHeadFromCYLA2b_16_1 - BFlossPrevYrUndSameHeadSetoff2b_16_1 > 0 Then
                adjustSpeci = WorksheetFunction.Min(BpPendingLossAmt - adjustSpecu, IncOfCurYrUndHeadFromCYLA2b_16_1 - BFlossPrevYrUndSameHeadSetoff2b_16_1)
                BFlossPrevYrUndSameHeadSetoff2b_16_1 = BFlossPrevYrUndSameHeadSetoff2b_16_1 + adjustSpeci
            End If
        End If
    End If
    
    
    Sheet16.Range("busipofincl.BFlossPrevYrUndSameHeadSetoff2").value = BFlossPrevYrUndSameHeadSetoff2_16_1
    Sheet16.Range("busipofinclspecified.BFlossPrevYrUndSameHeadSetoff2b").value = BFlossPrevYrUndSameHeadSetoff2b_16_1
    Sheet16.Range("busipofinclspec.BFlossPrevYrUndSameHeadSetoff2a").value = BFlossPrevYrUndSameHeadSetoff2a_16_1

'Setting off business loss with spec and specified business loss finished

'*****
'added by Chandru
If (Mid(Range("sheet16.BFLA_Edit").value, 1, 1) = "Y") Or (Mid(Range("sheet16.CYLA_Edit").value, 1, 1) = "Y") Then
    Sheet16.Unprotect Password:=getmsgstate
'        Sheet16.Range("BFlossPrevYrUndSameHeadSetoff3a").Locked = False
'        Sheet16.Range("busipofinclspec.BFlossPrevYrUndSameHeadSetoff2a").Locked = False
'        Sheet16.Range("busipofinclspecified.BFlossPrevYrUndSameHeadSetoff2b").Locked = False
       ' Sheet16.Range("rh.BFlossPrevYrUndSameHeadSetoff6").Locked = False
       ' Sheet16.Range("HP.BFlossPrevYrUndSameHeadSetoff1").Locked = False
'Commented by Riyaz on 30/01/2026
'        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3").Locked = False
        '---STCG 20% NEWLY ADDED BY BINDU ON 20TH FEB 25
        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3c").Locked = False
        
        
        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3a").Locked = False
        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b").Locked = False
        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b_DTAARates").Locked = False
'Commented by Riyaz on 30/01/2026
'        Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4").Locked = False
        '---LTCG 12.5% NEWLY ADDED BY BINDU ON 20TH FEB 25
        Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4b").Locked = False
'Commented by Riyaz on 30/01/2026
'        Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4a").Locked = False

        Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff_DTAARates").Locked = False
               
      '  Sheet16.Range("HP.BFlossPrevYrUndSameHeadSetoff1").Interior.Color = RGB(255, 255, 204)
'Commented by Riyaz on 30/01/2026
'        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3").Interior.Color = RGB(255, 255, 204)
        '--STCG 20% NEWLY ADDED BY BINDU 20TH FEB 25
        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3c").Interior.Color = RGB(255, 255, 204)
        
        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3a").Interior.Color = RGB(255, 255, 204)
        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b").Interior.Color = RGB(255, 255, 204)
        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b_DTAARates").Interior.Color = RGB(255, 255, 204)
'Commented by Riyaz on 30/01/2026
'        Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4").Interior.Color = RGB(255, 255, 204)
        '---LTCG 12.5% NEWLY ADDED BY BINDU ON 20TH FEB 25
        Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4b").Interior.Color = RGB(255, 255, 204)
        
'Commented by Riyaz on 30/01/2026
'        Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4a").Interior.Color = RGB(255, 255, 204)
        Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff_DTAARates").Interior.Color = RGB(255, 255, 204)
'        Sheet16.Range("BFlossPrevYrUndSameHeadSetoff3a").Interior.Color = RGB(255, 255, 204)
'        Sheet16.Range("busipofinclspec.BFlossPrevYrUndSameHeadSetoff2a").Interior.Color = RGB(255, 255, 204)
'        Sheet16.Range("busipofinclspecified.BFlossPrevYrUndSameHeadSetoff2b").Interior.Color = RGB(255, 255, 204)
        ' Sheet16.Range("rh.BFlossPrevYrUndSameHeadSetoff6").Interior.Color = RGB(255, 255, 204)
        

    Else
'        Sheet16.Range("BFlossPrevYrUndSameHeadSetoff3a").Locked = True
'        Sheet16.Range("busipofinclspec.BFlossPrevYrUndSameHeadSetoff2a").Locked = True
'        Sheet16.Range("busipofinclspecified.BFlossPrevYrUndSameHeadSetoff2b").Locked = True

     '  Sheet16.Range("HP.BFlossPrevYrUndSameHeadSetoff1").Locked = True
       ' Sheet16.Range("rh.BFlossPrevYrUndSameHeadSetoff6").Locked = True
'Commented by Riyaz on 30/01/2026
'        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3").Locked = True
        '---STCG 20% NEWLY ADDED BY BINDU ON 20TH FEB 25
        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3c").Locked = True
        
        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3a").Locked = True
        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b").Locked = True
        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b_DTAARates").Locked = True
'Commented by Riyaz on 30/01/2026
'        Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4").Locked = True
        '---LTCG 12.5% NEWLY ADDED BY BINDU ON 20TH FEB 25
        Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4b").Locked = True
        
'Commented by Riyaz on 30/01/2026
'        Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4a").Locked = True
        Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff_DTAARates").Locked = True
               
       ' Sheet16.Range("HP.BFlossPrevYrUndSameHeadSetoff1").Interior.ColorIndex = 2
'Commented by Riyaz on 30/01/2026
'        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3").Interior.ColorIndex = 2
        '---STCG 20% NEWLY ADDED BY BINDU ON 20TH FEB 25
        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3c").Interior.ColorIndex = 2
        
        
        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3a").Interior.ColorIndex = 2
        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b").Interior.ColorIndex = 2
        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b_DTAARates").Interior.ColorIndex = 2
'Commented by Riyaz on 30/01/2026
'        Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4").Interior.ColorIndex = 2
        '---LTCG 12.5% NEWLY ADDED BY BINDU ON 20TH FEB 25
        Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4b").Interior.ColorIndex = 2
        
'Commented by Riyaz on 30/01/2026
'        Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4a").Interior.ColorIndex = 2
        Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff_DTAARates").Interior.ColorIndex = 2
        'Sheet16.Range("BFlossPrevYrUndSameHeadSetoff3a").Interior.ColorIndex = 2
'        Sheet16.Range("busipofinclspec.BFlossPrevYrUndSameHeadSetoff2a").Interior.ColorIndex = 2
'        Sheet16.Range("busipofinclspecified.BFlossPrevYrUndSameHeadSetoff2b").Interior.ColorIndex = 2
       ' Sheet16.Range("rh.BFlossPrevYrUndSameHeadSetoff6").Interior.ColorIndex = 2
        
 'end Change Chandru
'------------
'---30% STCG
Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3a").value = Application.WorksheetFunction.Max(0, stcgBreakUp1 - tempstcgBreakUp1) + _
                                                             Application.WorksheetFunction.Max(0, stcgBreakUp6 - tempstcgBreakUp6)

'---STCG APP
Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b").value = Application.WorksheetFunction.Max(0, stcgBreakUp2 - tempstcgBreakUp2)

'---STCG 15%
'Commented by Riyaz on 30/01/2026
'Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3").value = Application.WorksheetFunction.Max(0, stcgBreakUp3 - tempstcgBreakUp3) + _
                                                            Application.WorksheetFunction.Max(0, stcgBreakUp4 - tempstcgBreakUp4) + _
                                                             Application.WorksheetFunction.Max(0, stcgBreakUp7 - tempstcgBreakUp7)
                                                             
 '---STCG 20%  NEWLY ADDED BY BINDU 20TH FEB 25
Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3c").value = Application.WorksheetFunction.Max(0, stcgBreakUp3_New - tempstcgBreakUp3_New) + _
                                                              Application.WorksheetFunction.Max(0, stcgBreakUp4_New - tempstcgBreakUp4_New) + _
                                                              Application.WorksheetFunction.Max(0, stcgBreakUp7_New - tempstcgBreakUp7_New)
                                                            
 Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b_DTAARates").value = Application.WorksheetFunction.Max(0, stcgBreakUp5 - tempstcgBreakUp5)
                                                            
'ltcgBreakUp2_New Newly added by Bindu on 15th March as per DE V4
'Commented by Riyaz on 30/01/2026
'Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4a").value = Application.WorksheetFunction.Max(0, ltcgBreakUp1 - templtcgBreakUp1) + _
                                                              Application.WorksheetFunction.Max(0, ltcgBreakUp10 - templtcgBreakUp10) + _
                                                            Application.WorksheetFunction.Max(0, ltcgBreakUp2 - templtcgBreakUp2) + _
                                                            Application.WorksheetFunction.Max(0, ltcgBreakUp2_New - templtcgBreakUp2_New)
                                                            
'---LTCG 12.5% NEWLY ADDED BY BINDU 26th FEB 25
Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4b").value = Application.WorksheetFunction.Max(0, ltcgBreakUp15_New - templtcgBreakUp15_New) + _
                                                               Application.WorksheetFunction.Max(0, ltcgBreakUp13_New - templtcgBreakUp13_New) + _
                                                               Application.WorksheetFunction.Max(0, ltcgBreakUp14_New - templtcgBreakUp14_New) + _
                                                               Application.WorksheetFunction.Max(0, ltcgBreakUp4_New - templtcgBreakUp4_New) + _
                                                               Application.WorksheetFunction.Max(0, ltcgBreakUp5_New - templtcgBreakUp5_New) + _
                                                               Application.WorksheetFunction.Max(0, ltcgBreakUp6_New - templtcgBreakUp6_New) + _
                                                               Application.WorksheetFunction.Max(0, ltcgBreakUp7_1_New - templtcgBreakUp7_1_New) + _
                                                               Application.WorksheetFunction.Max(0, ltcgBreakUp8_New - templtcgBreakUp8_New) + _
                                                               Application.WorksheetFunction.Max(0, ltcgBreakUp11_New - templtcgBreakUp11_New) + _
                                                               Application.WorksheetFunction.Max(0, ltcgBreakUp16_New - templtcgBreakUp16_New) + _
                                                               Application.WorksheetFunction.Max(0, ltcgBreakUp12_New - templtcgBreakUp12_New) + _
                                                                Application.WorksheetFunction.Max(0, ltcgBreakUp7_2_New - templtcgBreakUp7_2_New) + Application.WorksheetFunction.Max(0, ltcgBreakUp17_New - templtcgBreakUp17_New)



                                                            
'Application.WorksheetFunction.Max(0, ltcgBreakUp3 - templtcgBreakUp3) ''NOt using  deleted  by Bindu as this sec changed from 10 to 20
'Commented by Riyaz on 30/01/2026
'Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4").value = Application.WorksheetFunction.Max(0, ltcgBreakUp7_1 - templtcgBreakUp7_1) + _
                                                             Application.WorksheetFunction.Max(0, ltcgBreakUp7_2 - templtcgBreakUp7_2) + _
                                                            Application.WorksheetFunction.Max(0, ltcgBreakUp4 - templtcgBreakUp4) + _
                                                            Application.WorksheetFunction.Max(0, ltcgBreakUp5 - templtcgBreakUp5) + _
                                                            Application.WorksheetFunction.Max(0, ltcgBreakUp6 - templtcgBreakUp6) + _
                                                            Application.WorksheetFunction.Max(0, ltcgBreakUp7 - templtcgBreakUp7) + _
                                                           Application.WorksheetFunction.Max(0, ltcgBreakUp8 - templtcgBreakUp8) + _
                                                           Application.WorksheetFunction.Max(0, ltcgBreakUp11 - templtcgBreakUp11) + _
                                                           Application.WorksheetFunction.Max(0, ltcgBreakUp12 - templtcgBreakUp12)
                                                           
                                                           
Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff_DTAARates").value = Application.WorksheetFunction.Max(0, ltcgBreakUp9 - templtcgBreakUp9)
                                                            
      End If
      
      
      
'Chandru

 If (Mid(Range("sheet16.BFLA_Edit").value, 1, 1) = "Y") And Sheet17.Range("totofbfloss.STCGLossCF8").value > 0 And Sheet17.Range("totofbfloss.LTCGLossCF8").value = 0 Then
        Sheet17.Range("adjtotloss.STCGLossCF9").value = Sheet16.Range("sheet16.TotBFLossSetoff").value
 Else
 
 '---LTCG 12.5% NEWLY ADDED BY BINDU AS UNABLE TO ADD IN BELOW LINE BCZ OF TO MANY LINE CONTINUES ERROR

Dim LTCG_125_DUMMY
LTCG_125_DUMMY = Application.WorksheetFunction.Max(0, storeLTCG15_New - templtcgBreakUp15_New) + _
                 Application.WorksheetFunction.Max(0, storeLTCG13_New - templtcgBreakUp13_New) + _
                 Application.WorksheetFunction.Max(0, storeLTCG14_New - templtcgBreakUp14_New) + _
                 Application.WorksheetFunction.Max(0, storeLTCG4_New - templtcgBreakUp4_New) + _
                 Application.WorksheetFunction.Max(0, storeLTCG5_New - templtcgBreakUp5_New) + _
                 Application.WorksheetFunction.Max(0, storeLTCG6_New - templtcgBreakUp6_New) + _
                 Application.WorksheetFunction.Max(0, storeLTCG7_1_New - templtcgBreakUp7_1_New) + _
                 Application.WorksheetFunction.Max(0, storeLTCG8_New - templtcgBreakUp8_New) + _
                 Application.WorksheetFunction.Max(0, storeLTCG11_New - templtcgBreakUp11_New) + _
                 Application.WorksheetFunction.Max(0, storeLTCG16_New - templtcgBreakUp16_New) + _
                 Application.WorksheetFunction.Max(0, storeLTCG12_New - templtcgBreakUp12_New) + _
                 Application.WorksheetFunction.Max(0, storeLTCG7_2_New - templtcgBreakUp7_2_New) + Application.WorksheetFunction.Max(0, storeLTCG17_New - templtcgBreakUp17_New)
 
 'tempstcgBreakUp3_New, tempstcgBreakUp4_New,tempstcgBreakUp7_New - STCG 20% NEWLY ADDED BY BINDU
 ' Application.WorksheetFunction.Max(0, storeLTCG3 - templtcgBreakUp3)''NOt using  deleted  by Bindu as this sec changed from 10 to 20
 'Application.WorksheetFunction.Max(0, storeLTCG2_New - templtcgBreakUp2_New)'Newly added  by Bindu
 
'Commented by Riyaz on 30/01/2026
'     Sheet17.Range("adjtotloss.STCGLossCF9").value = Application.WorksheetFunction.Max(0, storeSTCG1 - tempstcgBreakUp1) + _
                                               Application.WorksheetFunction.Max(0, storeSTCG2 - tempstcgBreakUp2) + _
                                               Application.WorksheetFunction.Max(0, storeSTCG3 - tempstcgBreakUp3) + _
                                               Application.WorksheetFunction.Max(0, storeSTCG3_New - tempstcgBreakUp3_New) + _
                                               Application.WorksheetFunction.Max(0, storeSTCG4 - tempstcgBreakUp4) + _
                                               Application.WorksheetFunction.Max(0, storeSTCG4_New - tempstcgBreakUp4_New) + _
                                               Application.WorksheetFunction.Max(0, storeSTCG5 - tempstcgBreakUp5) + _
                                               Application.WorksheetFunction.Max(0, storeSTCG6 - tempstcgBreakUp6) + _
                                               Application.WorksheetFunction.Max(0, storeSTCG7 - tempstcgBreakUp7) + _
                                               Application.WorksheetFunction.Max(0, storeSTCG7_New - tempstcgBreakUp7_New) + _
                                               Application.WorksheetFunction.Max(0, storeLTCG1 - templtcgBreakUp1) + _
                                               Application.WorksheetFunction.Max(0, storeLTCG2 - templtcgBreakUp2) + _
                                               Application.WorksheetFunction.Max(0, storeLTCG2_New - templtcgBreakUp2_New) + _
                                               Application.WorksheetFunction.Max(0, storeLTCG4 - templtcgBreakUp4) + _
                                               Application.WorksheetFunction.Max(0, storeLTCG5 - templtcgBreakUp5) + _
                                               Application.WorksheetFunction.Max(0, storeLTCG6 - templtcgBreakUp6) + _
                                               Application.WorksheetFunction.Max(0, storeLTCG7 - templtcgBreakUp7) + _
                                               Application.WorksheetFunction.Max(0, storeLTCG7_1 - templtcgBreakUp7_1) + _
                                               Application.WorksheetFunction.Max(0, storeLTCG7_2 - templtcgBreakUp7_2) + _
                                                Application.WorksheetFunction.Max(0, storeLTCG8 - templtcgBreakUp8) + _
                                               Application.WorksheetFunction.Max(0, storeLTCG9 - templtcgBreakUp9) + _
                                                Application.WorksheetFunction.Max(0, storeLTCG10 - templtcgBreakUp10) + _
                                                 Application.WorksheetFunction.Max(0, storeLTCG11 - templtcgBreakUp11) + _
                                                 Application.WorksheetFunction.Max(0, storeLTCG12 - templtcgBreakUp12) + LTCG_125_DUMMY
                                                 
                                              
Sheet17.Range("adjtotloss.STCGLossCF9").value = Application.WorksheetFunction.Max(0, storeSTCG1 - tempstcgBreakUp1) + _
                                               Application.WorksheetFunction.Max(0, storeSTCG2 - tempstcgBreakUp2) + _
                                               Application.WorksheetFunction.Max(0, storeSTCG3_New - tempstcgBreakUp3_New) + _
                                               Application.WorksheetFunction.Max(0, storeSTCG4_New - tempstcgBreakUp4_New) + _
                                               Application.WorksheetFunction.Max(0, storeSTCG5 - tempstcgBreakUp5) + _
                                               Application.WorksheetFunction.Max(0, storeSTCG6 - tempstcgBreakUp6) + _
                                               Application.WorksheetFunction.Max(0, storeSTCG7_New - tempstcgBreakUp7_New) + _
                                                LTCG_125_DUMMY
     

                                                 

End If
'  '<------------UD Depriciation-------------->'
                                            
  Dim totalUDDeprctn As Double, busProfSpecProfset As Double
  
  
  
  
 totalUDDeprctn = Sheet15.Range("UD.TotSetoff").value
 busProfSpecProfset = Sheet16.Range("busipofincl.IncOfCurYrUndHeadFromCYLA2").value - Sheet16.Range("busipofincl.BFlossPrevYrUndSameHeadSetoff2").value
 
 If (busProfSpecProfset > 0) Then
      If (busProfSpecProfset <= totalUDDeprctn) Then
        Sheet16.Range("busipofincl.BFUnabsorbedDeprSetoff2").value = busProfSpecProfset
        totalUDDeprctn = totalUDDeprctn - busProfSpecProfset
           
      Else:
      Sheet16.Range("busipofincl.BFUnabsorbedDeprSetoff2").value = totalUDDeprctn
      totalUDDeprctn = 0
    End If
Else:
        Sheet16.Range("busipofincl.BFUnabsorbedDeprSetoff2").value = 0
End If
    

'Speculative
Dim Speculativeset As Double

Speculativeset = Sheet16.Range("busipofinclspec.IncOfCurYrUndHeadFromCYLA2a").value - Sheet16.Range("busipofinclspec.BFlossPrevYrUndSameHeadSetoff2a").value

If (Speculativeset > 0) Then
    If (Speculativeset <= totalUDDeprctn) Then
        Sheet16.Range("busipofinclspec.BFUnabsorbedDeprSetoff2a").value = Speculativeset
        totalUDDeprctn = totalUDDeprctn - Speculativeset
    Else:
        Sheet16.Range("busipofinclspec.BFUnabsorbedDeprSetoff2a").value = totalUDDeprctn
        totalUDDeprctn = 0
    End If
Else:
    Sheet16.Range("busipofinclspec.BFUnabsorbedDeprSetoff2a").value = 0
End If
    

'Specified
Dim specifiedSet As Double

specifiedSet = Sheet16.Range("busipofinclspecified.IncOfCurYrUndHeadFromCYLA2b").value - Sheet16.Range("busipofinclspecified.BFlossPrevYrUndSameHeadSetoff2b").value

If (specifiedSet > 0) Then
    If (specifiedSet <= totalUDDeprctn) Then
       Sheet16.Range("busipofinclspecified.BFUnabsorbedDeprSetoff2b").value = specifiedSet
        totalUDDeprctn = totalUDDeprctn - specifiedSet
        
    Else:
        Sheet16.Range("busipofinclspecified.BFUnabsorbedDeprSetoff2b").value = totalUDDeprctn
        totalUDDeprctn = 0
    End If
Else:
        Sheet16.Range("busipofinclspecified.BFUnabsorbedDeprSetoff2b").value = 0
End If
    
    
'House Prperty
Dim hpset As Double

hpset = Sheet16.Range("hp.IncOfCurYrUndHeadFromCYLA1").value - Sheet16.Range("hp.BFlossPrevYrUndSameHeadSetoff1").value

    If (hpset > 0) Then
        If (hpset <= totalUDDeprctn) Then
            Sheet16.Range("hp.BFUnabsorbedDeprSetoff1").value = hpset
            totalUDDeprctn = totalUDDeprctn - hpset
        
     Else:
        Sheet16.Range("hp.BFUnabsorbedDeprSetoff1").value = totalUDDeprctn
        totalUDDeprctn = 0
    End If
        Else:
        Sheet16.Range("hp.BFUnabsorbedDeprSetoff1").value = 0
        
    End If
    
'Other Sources
Dim otherSrcset As Double
otherSrcset = Sheet16.Range("othsrcincl.IncOfCurYrUndHeadFromCYLA5").value

    If (otherSrcset > 0) Then
    If (otherSrcset <= totalUDDeprctn) Then
        Sheet16.Range("othsrcincl.BFUnabsorbedDeprSetoff5").value = otherSrcset
        totalUDDeprctn = totalUDDeprctn - otherSrcset
    Else:
        Sheet16.Range("othsrcincl.BFUnabsorbedDeprSetoff5").value = totalUDDeprctn
        totalUDDeprctn = 0
    End If
    Else:
        Sheet16.Range("othsrcincl.BFUnabsorbedDeprSetoff5").value = 0
    End If
    
'    'Added by Chandru
'    If Sheet16.Range("SUM_NetInc2").value > Sheet16.Range("othsrcincl.IncOfCurYrUndHeadFromCYLA5").value Then
'    Sheet16.Range("othsrcincl.BFlossPrevYrUndSameHeadSetoff5", "othsrcincl.BFUnabsorbedDeprSetoff5").value = 0
'    Sheet16.Range("othsrcincl.BFAllUs35Cl4Setoff5").value = 0
'    MsgBox ("The sum of Column No. 2 + 3 should not exceed amount as referred in Column No. 1")
'    End If
''End change
        
'horse
Dim profitOwnset As Double

profitOwnset = Sheet16.Range("rh.IncOfCurYrUndHeadFromCYLA6").value - Sheet16.Range("rh.BFlossPrevYrUndSameHeadSetoff6").value

    If (profitOwnset > 0) Then
    If (profitOwnset <= totalUDDeprctn) Then
        Sheet16.Range("rh.BFUnabsorbedDeprSetoff6").value = profitOwnset
        totalUDDeprctn = totalUDDeprctn - profitOwnset
    Else:
        Sheet16.Range("rh.BFUnabsorbedDeprSetoff6").value = totalUDDeprctn
        totalUDDeprctn = 0
    End If
        
    Else:
        Sheet16.Range("rh.BFUnabsorbedDeprSetoff6").value = 0
    End If
    
'     'Added by Chandru
'    If Sheet16.Range("SUM_RaceHorse2").value > Sheet16.Range("rh.IncOfCurYrUndHeadFromCYLA6").value Then
'    Sheet16.Range("rh.BFlossPrevYrUndSameHeadSetoff6", "rh.IncOfCurYrAfterSetOffBFLosses6").value = 0
'    'Sheet16.Range("rh.BFAllUs35Cl4Setoff6").value = 0
'    MsgBox ("The sum of Column No. 2 + 3 should not exceed amount as referred in Column No. 1")
'    End If
''End change


 'OS DTAA
Dim OSDTAAOwnset As Double

OSDTAAOwnset = Sheet16.Range("rh.IncOfCurYrUndHeadFromCYLA_DTAARates").value
    If (OSDTAAOwnset > 0) Then
    If (OSDTAAOwnset <= totalUDDeprctn) Then
        Sheet16.Range("rh.BFUnabsorbedDeprSetoff_DTAARates").value = OSDTAAOwnset
        totalUDDeprctn = totalUDDeprctn - OSDTAAOwnset
    Else:
        Sheet16.Range("rh.BFUnabsorbedDeprSetoff_DTAARates").value = totalUDDeprctn
        totalUDDeprctn = 0
    End If

    Else:
        Sheet16.Range("rh.BFUnabsorbedDeprSetoff_DTAARates").value = 0
    End If
        
        
'UD at 30% setoff
  
Dim stcg30Perset As Double

    stcg30Perset = Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3a").value - Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3a").value
    If (stcg30Perset > 0) Then
    If (stcg30Perset <= totalUDDeprctn) Then
        Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3a").value = stcg30Perset
        totalUDDeprctn = totalUDDeprctn - stcg30Perset
        
    Else:
        Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3a").value = totalUDDeprctn
        totalUDDeprctn = 0
    End If
    Else:
        Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3a").value = 0
  End If

        
        
 'UD at Applicable Rate setoff
 Dim stcgAppset As Double
 
 stcgAppset = Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3b").value - Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b").value
 
    If (stcgAppset > 0) Then
    If (stcgAppset <= totalUDDeprctn) Then
        Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3b").value = stcgAppset
        totalUDDeprctn = totalUDDeprctn - stcgAppset
    
    Else:
        Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3b").value = totalUDDeprctn
        totalUDDeprctn = 0
    End If
    Else:
        Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3b").value = 0
    End If
    
    
'---UD STCG 20% NEWLY ADDED BY BINDU 20TH FEB 25
Dim stcg20Perset As Double
stcg20Perset = Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3c").value - Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3c").value

    If (stcg20Perset > 0) Then
    If (stcg20Perset <= totalUDDeprctn) Then
        Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3c").value = stcg20Perset
        totalUDDeprctn = totalUDDeprctn - stcg20Perset
    Else:
        Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3c").value = totalUDDeprctn
        totalUDDeprctn = 0
    End If
    Else:
        Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3c").value = 0
    End If
'---
    
 'UD at LTCG20%
'Commented by Riyaz on 30/01/2026
' Dim ltcg20Perset As Double
' ltcg20Perset = Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4a").value - Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4a").value
'
'
'    If (ltcg20Perset > 0) Then
'    If (ltcg20Perset <= totalUDDeprctn) Then
'        Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4a").value = ltcg20Perset
'        totalUDDeprctn = totalUDDeprctn - ltcg20Perset
'    Else:
'        Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4a").value = totalUDDeprctn
'        totalUDDeprctn = 0
'   End If
'   Else:
'        Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4a").value = 0
'    End If

'UD at STCG15%
'Commented by Riyaz on 30/01/2026
'Dim stcg15Perset As Double
'stcg15Perset = Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3").value - Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3").value
'
'    If (stcg15Perset > 0) Then
'    If (stcg15Perset <= totalUDDeprctn) Then
'        Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3").value = stcg15Perset
'        totalUDDeprctn = totalUDDeprctn - stcg15Perset
'    Else:
'        Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3").value = totalUDDeprctn
'        totalUDDeprctn = 0
'    End If
'    Else:
'        Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3").value = 0
'    End If
    
'--UD LTCG 12.5% NEWLY ADDED BY BINDU 26TH FEB 25
Dim ltcg125Perset As Double
ltcg125Perset = Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4b").value - Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4b").value
    
    If (ltcg125Perset > 0) Then
    If (ltcg125Perset <= totalUDDeprctn) Then
        Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4b").value = ltcg125Perset
        totalUDDeprctn = totalUDDeprctn - ltcg125Perset
    Else:
 
        Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4b").value = totalUDDeprctn
        totalUDDeprctn = 0
    End If
    Else:
        Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4b").value = 0
    End If
'--END
        
'UD at LTCG10%
'Commented by Riyaz on 30/01/2026
'Dim ltcg10Perset As Double
'ltcg10Perset = Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4").value - Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4").value
'
'    If (ltcg10Perset > 0) Then
'    If (ltcg10Perset <= totalUDDeprctn) Then
'        Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4").value = ltcg10Perset
'        totalUDDeprctn = totalUDDeprctn - ltcg10Perset
'    Else:
'
'        Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4").value = totalUDDeprctn
'        totalUDDeprctn = 0
'    End If
'    Else:
'        Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4").value = 0
'    End If
        
        
  'Stcg DTAA rates
Dim StcgDTAAOwnset As Double

StcgDTAAOwnset = Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3b_DTAARates").value - Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b_DTAARates").value
    If (StcgDTAAOwnset > 0) Then
    If (StcgDTAAOwnset <= totalUDDeprctn) Then
        Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3b_DTAARates").value = StcgDTAAOwnset
        totalUDDeprctn = totalUDDeprctn - StcgDTAAOwnset
    Else:
        Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3b_DTAARates").value = totalUDDeprctn
        totalUDDeprctn = 0
    End If
        
    Else:
        Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3b_DTAARates").value = 0
    End If
        
  'Ltcg DTAA rates
Dim LtcgDTAAOwnset As Double

LtcgDTAAOwnset = Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4a_DTAARates").value - Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff_DTAARates").value
    If (LtcgDTAAOwnset > 0) Then
    If (LtcgDTAAOwnset <= totalUDDeprctn) Then
        Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff_DTAARates").value = LtcgDTAAOwnset
        totalUDDeprctn = totalUDDeprctn - LtcgDTAAOwnset
    Else:
        Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff_DTAARates").value = totalUDDeprctn
        totalUDDeprctn = 0
    End If
        
    Else:
        Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff_DTAARates").value = 0
    End If
 '// calculating sum of column 3 BFLA
 
 '---Calculated at front end-------------
 
 '//
                                               
                                               
'<-----------End of Depreciation set off----


'//checking max limits of "Brought forward depreciation set off" --- column 4
Dim busProfSpecAllwnceSet As Double, totalUDAllwnce As Double
busProfSpecAllwnceSet = Sheet16.Range("busipofincl.IncOfCurYrUndHeadFromCYLA2").value - Sheet16.Range("busipofincl.BFlossPrevYrUndSameHeadSetoff2").value - Sheet16.Range("busipofincl.BFUnabsorbedDeprSetoff2").value

totalUDAllwnce = Sheet15.Range("UD2.TotSetoff").value



If (busProfSpecAllwnceSet > 0) Then
    If (busProfSpecAllwnceSet <= totalUDAllwnce) Then
        Sheet16.Range("busipofincl.BFAllUs35Cl4Setoff2").value = busProfSpecAllwnceSet
        totalUDAllwnce = totalUDAllwnce - busProfSpecAllwnceSet
    Else:
        Sheet16.Range("busipofincl.BFAllUs35Cl4Setoff2").value = totalUDAllwnce
        totalUDAllwnce = 0
    End If
    Else:
        Sheet16.Range("busipofincl.BFAllUs35Cl4Setoff2").value = 0
    End If

'Allowances at Speculative
Dim SpeculatvAllwnceSet As Double
SpeculatvAllwnceSet = Sheet16.Range("busipofinclspec.IncOfCurYrUndHeadFromCYLA2a").value - Sheet16.Range("busipofinclspec.BFlossPrevYrUndSameHeadSetoff2a").value - Sheet16.Range("busipofinclspec.BFUnabsorbedDeprSetoff2a").value
     
    If (SpeculatvAllwnceSet > 0) Then
    If (SpeculatvAllwnceSet <= totalUDAllwnce) Then
        Sheet16.Range("busipofinclspec.BFAllUs35Cl4Setoff2a").value = SpeculatvAllwnceSet
        totalUDAllwnce = totalUDAllwnce - SpeculatvAllwnceSet
    Else:
   
        Sheet16.Range("busipofinclspec.BFAllUs35Cl4Setoff2a").value = totalUDAllwnce
        totalUDAllwnce = 0
    End If
   Else:
        Sheet16.Range("busipofinclspec.BFAllUs35Cl4Setoff2a").value = 0
    End If
    
'Allowances at Specified
  Dim SpecfiedAllwnceSet As Double
  
 SpecfiedAllwnceSet = Sheet16.Range("busipofinclspecified.IncOfCurYrUndHeadFromCYLA2b").value - Sheet16.Range("busipofinclspecified.BFlossPrevYrUndSameHeadSetoff2b").value - Sheet16.Range("busipofinclspecified.BFUnabsorbedDeprSetoff2b").value
     
    If (SpecfiedAllwnceSet > 0) Then
    If (SpecfiedAllwnceSet <= totalUDAllwnce) Then
        Sheet16.Range("busipofinclspecified.BFAllUs35Cl4Setoff2b").value = SpecfiedAllwnceSet
        totalUDAllwnce = totalUDAllwnce - SpecfiedAllwnceSet
   Else:
        Sheet16.Range("busipofinclspecified.BFAllUs35Cl4Setoff2b").value = totalUDAllwnce
        totalUDAllwnce = 0
    End If
    Else:
        Sheet16.Range("busipofinclspecified.BFAllUs35Cl4Setoff2b").value = 0
    End If
    
'Allowances at hp
  Dim hpAllwnceSet As Double
  
    
hpAllwnceSet = Sheet16.Range("hp.IncOfCurYrUndHeadFromCYLA1").value - Sheet16.Range("hp.BFlossPrevYrUndSameHeadSetoff1").value - Sheet16.Range("hp.BFUnabsorbedDeprSetoff1").value
    If (hpAllwnceSet > 0) Then
    If (hpAllwnceSet <= totalUDAllwnce) Then
        Sheet16.Range("hp.BFAllUs35Cl4Setoff1").value = hpAllwnceSet
        totalUDAllwnce = totalUDAllwnce - hpAllwnceSet
    Else:
        Sheet16.Range("hp.BFAllUs35Cl4Setoff1").value = totalUDAllwnce
        totalUDAllwnce = 0
    End If
    Else:
        Sheet16.Range("hp.BFAllUs35Cl4Setoff1").value = 0
    End If
    
'Allowances at others
Dim othrSrcAllwnceSet As Double
othrSrcAllwnceSet = Sheet16.Range("othsrcincl.IncOfCurYrUndHeadFromCYLA5").value - Sheet16.Range("othsrcincl.BFUnabsorbedDeprSetoff5").value
           
    If (othrSrcAllwnceSet > 0) Then
    If (othrSrcAllwnceSet <= totalUDAllwnce) Then
        Sheet16.Range("othsrcincl.BFAllUs35Cl4Setoff5").value = othrSrcAllwnceSet
        totalUDAllwnce = totalUDAllwnce - othrSrcAllwnceSet
    Else:
        Sheet16.Range("othsrcincl.BFAllUs35Cl4Setoff5").value = totalUDAllwnce
        totalUDAllwnce = 0
    End If
    Else:
        Sheet16.Range("othsrcincl.BFAllUs35Cl4Setoff5").value = 0
    End If
 
 'Allowances at Horse
Dim OSDTAAAllwnceSet As Double
OSDTAAAllwnceSet = Sheet16.Range("rh.IncOfCurYrUndHeadFromCYLA6").value - Sheet16.Range("rh.BFlossPrevYrUndSameHeadSetoff6").value - Sheet16.Range("rh.BFUnabsorbedDeprSetoff6").value
       
    If (OSDTAAAllwnceSet > 0) Then
    If (OSDTAAAllwnceSet <= totalUDAllwnce) Then
        Sheet16.Range("rh.BFAllUs35Cl4Setoff6").value = OSDTAAAllwnceSet
        totalUDAllwnce = totalUDAllwnce - OSDTAAAllwnceSet
     Else:
        Sheet16.Range("rh.BFAllUs35Cl4Setoff6").value = totalUDAllwnce
        totalUDAllwnce = 0
    End If
    Else:
        Sheet16.Range("rh.BFAllUs35Cl4Setoff6").value = 0
    End If
    
    'OS DTAA
Dim profitAllwnceSet As Double
profitAllwnceSet = Sheet16.Range("rh.IncOfCurYrUndHeadFromCYLA_DTAARates").value - Sheet16.Range("rh.BFUnabsorbedDeprSetoff_DTAARates").value
If (profitAllwnceSet > 0) Then
    If (profitAllwnceSet <= totalUDAllwnce) Then
        Sheet16.Range("rh.BFAllUs35Cl4Setoff_DTAARates").value = profitAllwnceSet
        totalUDAllwnce = totalUDAllwnce - profitAllwnceSet
     Else:
        Sheet16.Range("rh.BFAllUs35Cl4Setoff_DTAARates").value = totalUDAllwnce
        totalUDAllwnce = 0
    End If
    Else:
        Sheet16.Range("rh.BFAllUs35Cl4Setoff_DTAARates").value = 0
    End If
    
    

'Alowances at STCG30%
Dim stcg30PerAllwnceSet As Double
stcg30PerAllwnceSet = Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3a").value - Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3a").value - Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3a").value
        
    If (stcg30PerAllwnceSet > 0) Then
    If (stcg30PerAllwnceSet <= totalUDAllwnce) Then
        Sheet16.Range("stcg.BFAllUs35Cl4Setoff3a").value = stcg30PerAllwnceSet
        totalUDAllwnce = totalUDAllwnce - stcg30PerAllwnceSet
    Else:
        Sheet16.Range("stcg.BFAllUs35Cl4Setoff3a").value = totalUDAllwnce
        totalUDAllwnce = 0
    End If
    Else:
        Sheet16.Range("stcg.BFAllUs35Cl4Setoff3a").value = 0
    End If
    
''Added by Chandru
'    If Sheet16.Range("SUM_STCG30_2").value > Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3a").value Then
'    Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3a").value = 0
'    'Sheet16.Range("stcg.BFAllUs35Cl4Setoff3a").value = 0
'    MsgBox ("The sum of Column No. 2 + 3 should not exceed amount as referred in Column No. 1"), vbOKOnly, "Alert!"
'
'    End If
    
'Added by Chandru
'        If Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3a").value > Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3a").value Then
'        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3a").value = 0
'        MsgBox ("Value of column no 2 cannot be more than amount as referred in column 1."), vbOKOnly, "Alert!"
'
'        End If
'End change

'ChandruD
   
'Allwances at ApplicableRate
Dim stcgAppAllwnceSet As Double
stcgAppAllwnceSet = Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3b").value - Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b").value - Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3b").value
       
    If (stcgAppAllwnceSet > 0) Then
    If (stcgAppAllwnceSet <= totalUDAllwnce) Then
        Sheet16.Range("stcg.BFAllUs35Cl4Setoff3b").value = stcgAppAllwnceSet
        totalUDAllwnce = totalUDAllwnce - stcgAppAllwnceSet
     Else:
        Sheet16.Range("stcg.BFAllUs35Cl4Setoff3b").value = totalUDAllwnce
        totalUDAllwnce = 0
    End If
    Else:
            Sheet16.Range("stcg.BFAllUs35Cl4Setoff3b").value = 0
    End If
    
'Added by Chandru
'        If Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b").value > Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3b").value Then
'        Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b").value = 0
'        MsgBox ("Value of column no 2 cannot be more than amount as referred in column 1."), vbOKOnly, "Alert!"
'
'        End If
'End change


'Allowances at STCG 20% NEWLY ADDED BY BINDU 20TH FEB 25

Dim stcg20PerAllwnceSet As Double
stcg20PerAllwnceSet = Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3c").value - Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3c").value - Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3c").value
       
    If (stcg20PerAllwnceSet > 0) Then
    If (stcg20PerAllwnceSet <= totalUDAllwnce) Then
        Sheet16.Range("stcg.BFAllUs35Cl4Setoff3C").value = stcg20PerAllwnceSet
        totalUDAllwnce = totalUDAllwnce - stcg20PerAllwnceSet
    Else:
        Sheet16.Range("stcg.BFAllUs35Cl4Setoff3C").value = totalUDAllwnce
        totalUDAllwnce = 0
    End If
    Else:
        Sheet16.Range("stcg.BFAllUs35Cl4Setoff3C").value = 0
    End If


'Allowances at LTCG 20%
'Commented by Riyaz on 30/01/2026
'Dim ltcg20PerAllwnceSet As Double
'ltcg20PerAllwnceSet = Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4a").value - Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4a").value - Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4a").value
'
'
'    If (ltcg20PerAllwnceSet > 0) Then
'    If (ltcg20PerAllwnceSet <= totalUDAllwnce) Then
'        Sheet16.Range("ltcg.BFAllUs35Cl4Setoff4a").value = ltcg20PerAllwnceSet
'        totalUDAllwnce = totalUDAllwnce - ltcg20PerAllwnceSet
'    Else:
'        Sheet16.Range("ltcg.BFAllUs35Cl4Setoff4a").value = totalUDAllwnce
'        totalUDAllwnce = 0
'    End If
'    Else:
'        Sheet16.Range("ltcg.BFAllUs35Cl4Setoff4a").value = 0
'    End If
'
    'Added by Chandru
'            If Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4a").value > Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4a").value Then
'            Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4a").value = 0
'            MsgBox ("Value of column no 2 cannot be more than amount as referred in column 1."), vbOKOnly, "Alert!"
'
'            End If
'End change
    
 'Allowances at STCG 15%
'Commented by Riyaz on 30/01/2026
'Dim stcg15PerAllwnceSet As Double
'stcg15PerAllwnceSet = Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3").value - Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3").value - Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3").value
'
'    If (stcg15PerAllwnceSet > 0) Then
'    If (stcg15PerAllwnceSet <= totalUDAllwnce) Then
'        Sheet16.Range("stcg.BFAllUs35Cl4Setoff3").value = stcg15PerAllwnceSet
'        totalUDAllwnce = totalUDAllwnce - stcg15PerAllwnceSet
'    Else:
'        Sheet16.Range("stcg.BFAllUs35Cl4Setoff3").value = totalUDAllwnce
'        totalUDAllwnce = 0
'    End If
'    Else:
'        Sheet16.Range("stcg.BFAllUs35Cl4Setoff3").value = 0
'    End If
    
'Added by Chandru
        'If Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3").value > Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3").value Then
        'Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3").value = 0
        'MsgBox ("Value of column no 2 cannot be more than amount as referred in column 1."), vbOKOnly, "Alert!"
        '
        'End If
'End change

'Allowances at LTCG 12.5% Newly added by Bindu 26th Feb 25
Dim ltcg125PerAllwnceSet As Double
ltcg125PerAllwnceSet = Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4b").value - Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4b").value - Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4b").value
      
    If (ltcg125PerAllwnceSet > 0) Then
    If (ltcg125PerAllwnceSet <= totalUDAllwnce) Then
        Sheet16.Range("ltcg.BFAllUs35Cl4Setoff4b").value = ltcg125PerAllwnceSet
        totalUDAllwnce = totalUDAllwnce - ltcg125PerAllwnceSet
    Else:
        Sheet16.Range("ltcg.BFAllUs35Cl4Setoff4b").value = totalUDAllwnce
        totalUDAllwnce = 0
    End If
    Else:
        Sheet16.Range("ltcg.BFAllUs35Cl4Setoff4b").value = 0
    End If
    
    'Added by Chandru
'            If Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4b").value > Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4b").value Then
'            Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4b").value = 0
'            MsgBox ("Value of column no 2 cannot be more than amount as referred in column 1."), vbOKOnly, "Alert!"
'
'            End If
'End change


'Allowances at LTCG 10%
'Commented by Riyaz on 30/01/2026
'Dim ltcg10PerAllwnceSet As Double
'ltcg10PerAllwnceSet = Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4").value - Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4").value - Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4").value
'
'    If (ltcg10PerAllwnceSet > 0) Then
'    If (ltcg10PerAllwnceSet <= totalUDAllwnce) Then
'        Sheet16.Range("ltcg.BFAllUs35Cl4Setoff4").value = ltcg10PerAllwnceSet
'        totalUDAllwnce = totalUDAllwnce - ltcg10PerAllwnceSet
'    Else:
'        Sheet16.Range("ltcg.BFAllUs35Cl4Setoff4").value = totalUDAllwnce
'        totalUDAllwnce = 0
'    End If
'    Else:
'        Sheet16.Range("ltcg.BFAllUs35Cl4Setoff4").value = 0
'    End If
    
    'Added by Chandru
'            If Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4").value > Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4").value Then
'            Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff4").value = 0
'            MsgBox ("Value of column no 2 cannot be more than amount as referred in column 1."), vbOKOnly, "Alert!"
'
'            End If
'End change


    
  'stcg DTAA
Dim stcgDTAAAllwnceSet As Double
stcgDTAAAllwnceSet = Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3b_DTAARates").value - Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b_DTAARates").value - Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3b_DTAARates").value
If (stcgDTAAAllwnceSet > 0) Then
    If (stcgDTAAAllwnceSet <= totalUDAllwnce) Then
        Sheet16.Range("stcg.BFAllUs35Cl4Setoff3b_DTAARates").value = stcgDTAAAllwnceSet
        totalUDAllwnce = totalUDAllwnce - stcgDTAAAllwnceSet
     Else:
        Sheet16.Range("stcg.BFAllUs35Cl4Setoff3b_DTAARates").value = totalUDAllwnce
        totalUDAllwnce = 0
    End If
    Else:
        Sheet16.Range("stcg.BFAllUs35Cl4Setoff3b_DTAARates").value = 0
    End If
    
    
'Added by Chandru
'            If Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b_DTAARates").value > Sheet16.Range("stcg.IncOfCurYrUndHeadFromCYLA3b_DTAARates").value Then
'            Sheet16.Range("stcg.BFlossPrevYrUndSameHeadSetoff3b_DTAARates").value = 0
'            MsgBox ("Value of column no 2 cannot be more than amount as referred in column 1."), vbOKOnly, "Alert!"
'
'            End If
'End change


    
'ltcg DTAA
Dim ltcgDTAAAllwnceSet As Double
ltcgDTAAAllwnceSet = Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4a_DTAARates").value - Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff_DTAARates").value - Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff_DTAARates").value
If (ltcgDTAAAllwnceSet > 0) Then
    If (ltcgDTAAAllwnceSet <= totalUDAllwnce) Then
        Sheet16.Range("ltcg.BFAllUs35Cl4Setoff_DTAARates").value = ltcgDTAAAllwnceSet
        totalUDAllwnce = totalUDAllwnce - ltcgDTAAAllwnceSet
     Else:
        Sheet16.Range("ltcg.BFAllUs35Cl4Setoff_DTAARates").value = totalUDAllwnce
        totalUDAllwnce = 0
    End If
    Else:
        Sheet16.Range("ltcg.BFAllUs35Cl4Setoff_DTAARates").value = 0
    End If
    
'Added by Chandru
            'If Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff_DTAARates").value > Sheet16.Range("ltcg.IncOfCurYrUndHeadFromCYLA4a_DTAARates").value Then
            'Sheet16.Range("ltcg.BFlossPrevYrUndSameHeadSetoff_DTAARates").value = 0
            'MsgBox ("Value of column no 2 cannot be more than amount as referred in column 1."), vbOKOnly, "Alert!"
            '
            'End If
'End change


    
    
'//calculating the income after setOff

'--System Calculation last Column of BFLA---'


'----
Dim sumUD35_4_stcg15Per As Double, sumUD35_4_stcg30Per As Double, sumUD35_4_stcgAppPer, sumUD35_4_stcgDTAA As Double
Dim sumUD35_4_ltcg20Per As Double, sumUD35_4_ltcg10Per, sumUD35_4_ltcgDTAA As Double

'---NEWLY ADDED BY BINDU 20TH FEB 25
Dim sumUD35_4_stcg20Per, sumUD35_4_ltcg12_5Per As Double

'Commented by Riyaz on 30/01/2026
'sumUD35_4_stcg15Per = Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3").value + Sheet16.Range("stcg.BFAllUs35Cl4Setoff3").value
'---STCG 20% NEWLY ADDED BY BINDU 20TH FEB 25
sumUD35_4_stcg20Per = Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3c").value + Sheet16.Range("stcg.BFAllUs35Cl4Setoff3C").value

sumUD35_4_stcg30Per = Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3a").value + Sheet16.Range("stcg.BFAllUs35Cl4Setoff3a").value
sumUD35_4_stcgAppPer = Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3b").value + Sheet16.Range("stcg.BFAllUs35Cl4Setoff3b").value
'Commented by Riyaz on 30/01/2026
'sumUD35_4_ltcg20Per = Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4a").value + Sheet16.Range("ltcg.BFAllUs35Cl4Setoff4a").value
'sumUD35_4_ltcg10Per = Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4").value + Sheet16.Range("ltcg.BFAllUs35Cl4Setoff4").value
'----LTCG 12.5% NEWLY ADDED BY BINDU 20TH FEB 25
sumUD35_4_ltcg12_5Per = Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff4b").value + Sheet16.Range("ltcg.BFAllUs35Cl4Setoff4b").value

sumUD35_4_stcgDTAA = Sheet16.Range("stcg.BFUnabsorbedDeprSetoff3b_DTAARates").value + Sheet16.Range("stcg.BFAllUs35Cl4Setoff3b_DTAARates").value
sumUD35_4_ltcgDTAA = Sheet16.Range("ltcg.BFUnabsorbedDeprSetoff_DTAARates").value + Sheet16.Range("ltcg.BFAllUs35Cl4Setoff_DTAARates").value

        If (sumUD35_4_stcg30Per > tempstcgBreakUp1) Then
            sumUD35_4_stcg30Per = sumUD35_4_stcg30Per - tempstcgBreakUp1
            tempstcgBreakUp1 = 0
        Else:
            tempstcgBreakUp1 = Application.WorksheetFunction.Max(0, tempstcgBreakUp1 - sumUD35_4_stcg30Per)
            sumUD35_4_stcg30Per = 0
        End If
        
        If (sumUD35_4_stcg30Per > tempstcgBreakUp6) Then
            sumUD35_4_stcg30Per = sumUD35_4_stcg30Per - tempstcgBreakUp6
            tempstcgBreakUp6 = 0
        Else:
            tempstcgBreakUp6 = Application.WorksheetFunction.Max(0, tempstcgBreakUp6 - sumUD35_4_stcg30Per)
            sumUD35_4_stcg30Per = 0
        End If
        
        If (sumUD35_4_stcgAppPer > tempstcgBreakUp2) Then
            sumUD35_4_stcgAppPer = sumUD35_4_stcgAppPer - tempstcgBreakUp2
            tempstcgBreakUp2 = 0
        Else:
            tempstcgBreakUp2 = Application.WorksheetFunction.Max(0, tempstcgBreakUp2 - sumUD35_4_stcgAppPer)
            sumUD35_4_stcgAppPer = 0
        End If
        
        '-----STCG 20% NEWLY ADDED BY BINDU 20TH FEB 25
        
        If (sumUD35_4_stcg20Per > tempstcgBreakUp3_New) Then
            sumUD35_4_stcg20Per = sumUD35_4_stcg20Per - tempstcgBreakUp3_New
            tempstcgBreakUp3_New = 0
        Else:
            tempstcgBreakUp3_New = Application.WorksheetFunction.Max(0, tempstcgBreakUp3_New - sumUD35_4_stcg20Per)
            sumUD35_4_stcg20Per = 0
        End If
        
         If (sumUD35_4_stcg20Per > tempstcgBreakUp7_New) Then
            sumUD35_4_stcg20Per = sumUD35_4_stcg20Per - tempstcgBreakUp7_New
            tempstcgBreakUp7_New = 0
        Else:
            tempstcgBreakUp7_New = Application.WorksheetFunction.Max(0, tempstcgBreakUp7_New - sumUD35_4_stcg20Per)
            sumUD35_4_stcg20Per = 0
        End If
        
        If (sumUD35_4_stcg20Per > tempstcgBreakUp4_New) Then
            sumUD35_4_stcg20Per = sumUD35_4_stcg20Per - tempstcgBreakUp4_New
            tempstcgBreakUp4_New = 0
        Else:
            tempstcgBreakUp4_New = Application.WorksheetFunction.Max(0, tempstcgBreakUp4_New - sumUD35_4_stcg20Per)
            sumUD35_4_stcg20Per = 0
        End If
        '----
        
'Commented by Riyaz on 30/01/2026
'        If (sumUD35_4_ltcg20Per > templtcgBreakUp2) Then
'            sumUD35_4_ltcg20Per = sumUD35_4_ltcg20Per - templtcgBreakUp2
'            templtcgBreakUp2 = 0
'        Else:
'            templtcgBreakUp2 = Application.WorksheetFunction.Max(0, templtcgBreakUp2 - sumUD35_4_ltcg20Per)
'            sumUD35_4_ltcg20Per = 0
'        End If
'
'        If (sumUD35_4_ltcg20Per > templtcgBreakUp10) Then
'            sumUD35_4_ltcg20Per = sumUD35_4_ltcg20Per - templtcgBreakUp10
'            templtcgBreakUp10 = 0
'        Else:
'            templtcgBreakUp10 = Application.WorksheetFunction.Max(0, templtcgBreakUp10 - sumUD35_4_ltcg20Per)
'            sumUD35_4_ltcg20Per = 0
'        End If
'
'         '-'Newly added by Bindu as per DE V4 on 15th March 25
'         If (sumUD35_4_ltcg20Per > templtcgBreakUp2_New) Then
'            sumUD35_4_ltcg20Per = sumUD35_4_ltcg20Per - templtcgBreakUp2_New
'            templtcgBreakUp2_New = 0
'        Else:
'            templtcgBreakUp2_New = Application.WorksheetFunction.Max(0, templtcgBreakUp2_New - sumUD35_4_ltcg20Per)
'            sumUD35_4_ltcg20Per = 0
'        End If
        '--Ends

'        If (sumUD35_4_ltcg20Per > templtcgBreakUp1) Then
'            sumUD35_4_ltcg20Per = sumUD35_4_ltcg20Per - templtcgBreakUp1
'            templtcgBreakUp1 = 0
'        Else:
'            templtcgBreakUp1 = Application.WorksheetFunction.Max(0, templtcgBreakUp1 - sumUD35_4_ltcg20Per)
'            sumUD35_4_ltcg20Per = 0
'        End If
        
'Commented by Riyaz on 30/01/2026
'        If (sumUD35_4_stcg15Per > tempstcgBreakUp3) Then
'            sumUD35_4_stcg15Per = sumUD35_4_stcg15Per - tempstcgBreakUp3
'            tempstcgBreakUp3 = 0
'        Else:
'            tempstcgBreakUp3 = Application.WorksheetFunction.Max(0, tempstcgBreakUp3 - sumUD35_4_stcg15Per)
'            sumUD35_4_stcg15Per = 0
'        End If
'
'         If (sumUD35_4_stcg15Per > tempstcgBreakUp7) Then
'            sumUD35_4_stcg15Per = sumUD35_4_stcg15Per - tempstcgBreakUp7
'            tempstcgBreakUp7 = 0
'        Else:
'            tempstcgBreakUp7 = Application.WorksheetFunction.Max(0, tempstcgBreakUp7 - sumUD35_4_stcg15Per)
'            sumUD35_4_stcg15Per = 0
'        End If
'
'        If (sumUD35_4_stcg15Per > tempstcgBreakUp4) Then
'            sumUD35_4_stcg15Per = sumUD35_4_stcg15Per - tempstcgBreakUp4
'            tempstcgBreakUp4 = 0
'        Else:
'            tempstcgBreakUp4 = Application.WorksheetFunction.Max(0, tempstcgBreakUp4 - sumUD35_4_stcg15Per)
'            sumUD35_4_stcg15Per = 0
'        End If
        
        
        
' // adjusting STCG first
        
'Commented by Riyaz on 30/01/2026
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp4) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp4
'            templtcgBreakUp4 = 0
'        Else:
'            templtcgBreakUp4 = Application.WorksheetFunction.Max(0, templtcgBreakUp4 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
'
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp5) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp5
'            templtcgBreakUp5 = 0
'        Else:
'            templtcgBreakUp5 = Application.WorksheetFunction.Max(0, templtcgBreakUp5 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
'
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp6) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp6
'            templtcgBreakUp6 = 0
'        Else:
'            templtcgBreakUp6 = Application.WorksheetFunction.Max(0, templtcgBreakUp6 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
'
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp7) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp7
'            templtcgBreakUp7 = 0
'        Else:
'            templtcgBreakUp7 = Application.WorksheetFunction.Max(0, templtcgBreakUp7 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
'
'       'new
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp7_1) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp7_1
'            templtcgBreakUp7_1 = 0
'        Else:
'            templtcgBreakUp7_1 = Application.WorksheetFunction.Max(0, templtcgBreakUp7_1 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
'
'        'end
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp8) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp8
'            templtcgBreakUp8 = 0
'        Else:
'            templtcgBreakUp8 = Application.WorksheetFunction.Max(0, templtcgBreakUp8 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
'
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp11) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp11
'            templtcgBreakUp11 = 0
'        Else:
'            templtcgBreakUp11 = Application.WorksheetFunction.Max(0, templtcgBreakUp11 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
'
'        'NOt using  commented  by Bindu as this sec changed from 10 to 20
''        If (sumUD35_4_ltcg10Per > templtcgBreakUp3) Then
''            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp3
''            templtcgBreakUp3 = 0
''        Else:
''            templtcgBreakUp3 = Application.WorksheetFunction.Max(0, templtcgBreakUp3 - sumUD35_4_ltcg10Per)
''            sumUD35_4_ltcg10Per = 0
''        End If
'
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp12) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp12
'            templtcgBreakUp12 = 0
'        Else:
'            templtcgBreakUp12 = Application.WorksheetFunction.Max(0, templtcgBreakUp12 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
'
'        'new
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp7_2) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp7_2
'            templtcgBreakUp7_2 = 0
'        Else:
'            templtcgBreakUp7_2 = Application.WorksheetFunction.Max(0, templtcgBreakUp7_2 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
        'end
        
        'new
        If (sumUD35_4_stcgDTAA > tempstcgBreakUp5) Then
            sumUD35_4_stcgDTAA = sumUD35_4_stcgDTAA - tempstcgBreakUp5
            tempstcgBreakUp5 = 0
        Else:
            tempstcgBreakUp5 = Application.WorksheetFunction.Max(0, tempstcgBreakUp5 - sumUD35_4_stcgDTAA)
            sumUD35_4_stcgDTAA = 0
        End If
        'end
        
        'new
        If (sumUD35_4_ltcgDTAA > templtcgBreakUp9) Then
            sumUD35_4_ltcgDTAA = sumUD35_4_ltcgDTAA - templtcgBreakUp9
            templtcgBreakUp9 = 0
        Else:
            templtcgBreakUp9 = Application.WorksheetFunction.Max(0, templtcgBreakUp9 - sumUD35_4_ltcgDTAA)
            sumUD35_4_ltcgDTAA = 0
        End If
        'end
        
        
        
'Commented by Riyaz on 30/01/2026
'        If (sumUD35_4_ltcg20Per > templtcgBreakUp2) Then
'            sumUD35_4_ltcg20Per = sumUD35_4_ltcg20Per - templtcgBreakUp2
'            templtcgBreakUp2 = 0
'        Else:
'            templtcgBreakUp2 = Application.WorksheetFunction.Max(0, templtcgBreakUp2 - sumUD35_4_ltcg20Per)
'            sumUD35_4_ltcg20Per = 0
'        End If
'
'
'        If (sumUD35_4_ltcg20Per > templtcgBreakUp10) Then
'            sumUD35_4_ltcg20Per = sumUD35_4_ltcg20Per - templtcgBreakUp10
'            templtcgBreakUp10 = 0
'        Else:
'            templtcgBreakUp10 = Application.WorksheetFunction.Max(0, templtcgBreakUp10 - sumUD35_4_ltcg20Per)
'            sumUD35_4_ltcg20Per = 0
'        End If
'
'        '-'Newly added by Bindu as per DE V4 on 15th March 25
'        If (sumUD35_4_ltcg20Per > templtcgBreakUp2_New) Then
'            sumUD35_4_ltcg20Per = sumUD35_4_ltcg20Per - templtcgBreakUp2_New
'            templtcgBreakUp2_New = 0
'        Else:
'            templtcgBreakUp2_New = Application.WorksheetFunction.Max(0, templtcgBreakUp2_New - sumUD35_4_ltcg20Per)
'            sumUD35_4_ltcg20Per = 0
'        End If
'        '---
'
'        If (sumUD35_4_ltcg20Per > templtcgBreakUp1) Then
'            sumUD35_4_ltcg20Per = sumUD35_4_ltcg20Per - templtcgBreakUp1
'            templtcgBreakUp1 = 0
'        Else:
'            templtcgBreakUp1 = Application.WorksheetFunction.Max(0, templtcgBreakUp1 - sumUD35_4_ltcg20Per)
'            sumUD35_4_ltcg20Per = 0
'        End If
        
'Commented by Riyaz on 30/01/2026
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp4) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp4
'            templtcgBreakUp4 = 0
'        Else:
'            templtcgBreakUp4 = Application.WorksheetFunction.Max(0, templtcgBreakUp4 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
'
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp5) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp5
'            templtcgBreakUp5 = 0
'        Else:
'            templtcgBreakUp5 = Application.WorksheetFunction.Max(0, templtcgBreakUp5 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
'
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp6) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp6
'            templtcgBreakUp6 = 0
'        Else:
'            templtcgBreakUp6 = Application.WorksheetFunction.Max(0, templtcgBreakUp6 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
'
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp7) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp7
'            templtcgBreakUp7 = 0
'        Else:
'            templtcgBreakUp7 = Application.WorksheetFunction.Max(0, templtcgBreakUp7 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
'
'        'new
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp7_1) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp7_1
'            templtcgBreakUp7_1 = 0
'        Else:
'            templtcgBreakUp7_1 = Application.WorksheetFunction.Max(0, templtcgBreakUp7_1 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
'        'end
'
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp8) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp8
'            templtcgBreakUp8 = 0
'        Else:
'            templtcgBreakUp8 = Application.WorksheetFunction.Max(0, templtcgBreakUp8 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
'
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp11) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp11
'            templtcgBreakUp11 = 0
'        Else:
'            templtcgBreakUp11 = Application.WorksheetFunction.Max(0, templtcgBreakUp11 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
        
        'NOt using  commented  by Bindu as this sec changed from 10 to 20
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp3) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp3
'            templtcgBreakUp3 = 0
'        Else:
'            templtcgBreakUp3 = Application.WorksheetFunction.Max(0, templtcgBreakUp3 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
        
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp12) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp12
'            templtcgBreakUp12 = 0
'        Else:
'            templtcgBreakUp12 = Application.WorksheetFunction.Max(0, templtcgBreakUp12 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
'
'        'new
'        If (sumUD35_4_ltcg10Per > templtcgBreakUp7_2) Then
'            sumUD35_4_ltcg10Per = sumUD35_4_ltcg10Per - templtcgBreakUp7_2
'            templtcgBreakUp7_2 = 0
'        Else:
'            templtcgBreakUp7_2 = Application.WorksheetFunction.Max(0, templtcgBreakUp7_2 - sumUD35_4_ltcg10Per)
'            sumUD35_4_ltcg10Per = 0
'        End If
        'end
        
'        If (sumUD35_4_ltcgDTAA > templtcgBreakUp9) Then
'            sumUD35_4_ltcgDTAA = sumUD35_4_ltcgDTAA - templtcgBreakUp9
'            templtcgBreakUp9 = 0
'        Else:
'            templtcgBreakUp9 = Application.WorksheetFunction.Max(0, templtcgBreakUp9 - sumUD35_4_ltcgDTAA)
'            sumUD35_4_ltcgDTAA = 0
'        End If
        
'---LTCG 12.5% NEWLY ADDED BY BINDU 26TH FEB 25
        If (sumUD35_4_ltcg12_5Per > templtcgBreakUp4_New) Then
            sumUD35_4_ltcg12_5Per = sumUD35_4_ltcg12_5Per - templtcgBreakUp4_New
            templtcgBreakUp4_New = 0
        Else:
            templtcgBreakUp4_New = Application.WorksheetFunction.Max(0, templtcgBreakUp4_New - sumUD35_4_ltcg12_5Per)
            sumUD35_4_ltcg12_5Per = 0
        End If
        
        If (sumUD35_4_ltcg12_5Per > templtcgBreakUp5_New) Then
            sumUD35_4_ltcg12_5Per = sumUD35_4_ltcg12_5Per - templtcgBreakUp5_New
            templtcgBreakUp5_New = 0
        Else:
            templtcgBreakUp5_New = Application.WorksheetFunction.Max(0, templtcgBreakUp5_New - sumUD35_4_ltcg12_5Per)
            sumUD35_4_ltcg12_5Per = 0
        End If
        
        If (sumUD35_4_ltcg12_5Per > templtcgBreakUp6_New) Then
            sumUD35_4_ltcg12_5Per = sumUD35_4_ltcg12_5Per - templtcgBreakUp6_New
            templtcgBreakUp6_New = 0
        Else:
            templtcgBreakUp6_New = Application.WorksheetFunction.Max(0, templtcgBreakUp6_New - sumUD35_4_ltcg12_5Per)
            sumUD35_4_ltcg12_5Per = 0
        End If
        
        If (sumUD35_4_ltcg12_5Per > templtcgBreakUp8_New) Then
            sumUD35_4_ltcg12_5Per = sumUD35_4_ltcg12_5Per - templtcgBreakUp8_New
            templtcgBreakUp8_New = 0
        Else:
            templtcgBreakUp8_New = Application.WorksheetFunction.Max(0, templtcgBreakUp8_New - sumUD35_4_ltcg12_5Per)
            sumUD35_4_ltcg12_5Per = 0
        End If
        
       'new
        If (sumUD35_4_ltcg12_5Per > templtcgBreakUp7_1_New) Then
            sumUD35_4_ltcg12_5Per = sumUD35_4_ltcg12_5Per - templtcgBreakUp7_1_New
            templtcgBreakUp7_1_New = 0
        Else:
            templtcgBreakUp7_1_New = Application.WorksheetFunction.Max(0, templtcgBreakUp7_1_New - sumUD35_4_ltcg12_5Per)
            sumUD35_4_ltcg12_5Per = 0
        End If
        
        'end
'Added by Riyaz on 03/03/2026
        If (sumUD35_4_ltcg12_5Per > templtcgBreakUp17_New) Then
            sumUD35_4_ltcg12_5Per = sumUD35_4_ltcg12_5Per - templtcgBreakUp17_New
            templtcgBreakUp17_New = 0
        Else:
            templtcgBreakUp17_New = Application.WorksheetFunction.Max(0, templtcgBreakUp17_New - sumUD35_4_ltcg12_5Per)
            sumUD35_4_ltcg12_5Per = 0
        End If
        

        If (sumUD35_4_ltcg12_5Per > templtcgBreakUp11_New) Then
            sumUD35_4_ltcg12_5Per = sumUD35_4_ltcg12_5Per - templtcgBreakUp11_New
            templtcgBreakUp11_New = 0
        Else:
            templtcgBreakUp11_New = Application.WorksheetFunction.Max(0, templtcgBreakUp11_New - sumUD35_4_ltcg12_5Per)
            sumUD35_4_ltcg12_5Per = 0
        End If
        
        If (sumUD35_4_ltcg12_5Per > templtcgBreakUp12_New) Then
            sumUD35_4_ltcg12_5Per = sumUD35_4_ltcg12_5Per - templtcgBreakUp12_New
            templtcgBreakUp12_New = 0
        Else:
            templtcgBreakUp12_New = Application.WorksheetFunction.Max(0, templtcgBreakUp12_New - sumUD35_4_ltcg12_5Per)
            sumUD35_4_ltcg12_5Per = 0
        End If
        
        If (sumUD35_4_ltcg12_5Per > templtcgBreakUp13_New) Then
            sumUD35_4_ltcg12_5Per = sumUD35_4_ltcg12_5Per - templtcgBreakUp13_New
            templtcgBreakUp13_New = 0
        Else:
            templtcgBreakUp13_New = Application.WorksheetFunction.Max(0, templtcgBreakUp13_New - sumUD35_4_ltcg12_5Per)
            sumUD35_4_ltcg12_5Per = 0
        End If
        
        If (sumUD35_4_ltcg12_5Per > templtcgBreakUp14_New) Then
            sumUD35_4_ltcg12_5Per = sumUD35_4_ltcg12_5Per - templtcgBreakUp14_New
            templtcgBreakUp14_New = 0
        Else:
            templtcgBreakUp14_New = Application.WorksheetFunction.Max(0, templtcgBreakUp14_New - sumUD35_4_ltcg12_5Per)
            sumUD35_4_ltcg12_5Per = 0
        End If
        
        'new
        If (sumUD35_4_ltcg12_5Per > templtcgBreakUp7_2_New) Then
            sumUD35_4_ltcg12_5Per = sumUD35_4_ltcg12_5Per - templtcgBreakUp7_2_New
            templtcgBreakUp7_2_New = 0
        Else:
            templtcgBreakUp7_2_New = Application.WorksheetFunction.Max(0, templtcgBreakUp7_2_New - sumUD35_4_ltcg12_5Per)
            sumUD35_4_ltcg12_5Per = 0
        End If
        'end
        
        If (sumUD35_4_ltcg12_5Per > templtcgBreakUp15_New) Then
            sumUD35_4_ltcg12_5Per = sumUD35_4_ltcg12_5Per - templtcgBreakUp15_New
            templtcgBreakUp15_New = 0
        Else:
            templtcgBreakUp15_New = Application.WorksheetFunction.Max(0, templtcgBreakUp15_New - sumUD35_4_ltcg12_5Per)
            sumUD35_4_ltcg12_5Per = 0
        End If
        '--15th March
        If (sumUD35_4_ltcg12_5Per > templtcgBreakUp16_New) Then
            sumUD35_4_ltcg12_5Per = sumUD35_4_ltcg12_5Per - templtcgBreakUp16_New
            templtcgBreakUp16_New = 0
        Else:
            templtcgBreakUp16_New = Application.WorksheetFunction.Max(0, templtcgBreakUp16_New - sumUD35_4_ltcg12_5Per)
            sumUD35_4_ltcg12_5Per = 0
        End If
        
'--END


'        cgosIncome.cgInc.stcg.prctg30 = tempstcgBreakUp1;
'        cgosIncome.cgInc.stcg.prctgAr = tempstcgBreakUp2;
'        cgosIncome.cgInc.stcg.prctg15.sec115ad_1_b_ii = tempstcgBreakUp3;
'        cgosIncome.cgInc.stcg.prctg15.sec111a = tempstcgBreakUp4;
'        cgosIncome.cgInc.ltcg.prctg20.sec112 = templtcgBreakUp1;
'        cgosIncome.cgInc.ltcg.prctg20.sec11EA = templtcgBreakUp2;
'        cgosIncome.cgInc.ltcg.prctg10.secProviso = templtcgBreakUp3;
'        cgosIncome.cgInc.ltcg.prctg10.sec112_1_c_2 = templtcgBreakUp4;
'        cgosIncome.cgInc.ltcg.prctg10.sec115AC_1 = templtcgBreakUp5;
'        cgosIncome.cgInc.ltcg.prctg10.sec115ACA_1 = templtcgBreakUp6;
'        cgosIncome.cgInc.ltcg.prctg10.sec115AD_3 = templtcgBreakUp7;
'        cgosIncome.cgInc.ltcg.prctg10.sec115E_b = templtcgBreakUp8;


                                               
                                               
'<-------------End of Allowances set off at BFLA----------------
                                               
                                               
                                               
   '<----------------------End------------------------->
   
        cgosIncome_cgInc_stcg_prctg30_sec115AD = tempstcgBreakUp1
        cgosIncome_cgInc_stcg_prctgAppRate = tempstcgBreakUp2
'Commented by Riyaz on 30/01/2026
'        cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii = tempstcgBreakUp3
'        cgosIncome_cgInc_stcg_prctg15_sec111a = tempstcgBreakUp4

        '----STCG 20% NEWLY ADDED BY BINDU 20TH FEB 25
        cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii = tempstcgBreakUp3_New
        cgosIncome_cgInc_stcg_prctg20_sec111a = tempstcgBreakUp4_New
        
'Commented by Riyaz on 30/01/2026
'        cgosIncome_cgInc_ltcg_prctg20_sec112 = templtcgBreakUp1
'        cgosIncome_cgInc_ltcg_prctg20_sec11EA = templtcgBreakUp2
        'Newly added by Bindu as per DE V4 on 15th March 25
'        cgosIncome_cgInc_ltcg_prctg20_secProviso = templtcgBreakUp2_New
        
        'cgosIncome_cgInc_ltcg_prctg10_secProviso = templtcgBreakUp3 ''NOt using  commented  by Bindu as this sec changed from 10 to 20
'Commented by Riyaz on 30/01/2026
'        cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2 = templtcgBreakUp4
'        cgosIncome_cgInc_ltcg_prctg10_sec115AC_1 = templtcgBreakUp5
'        cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1 = templtcgBreakUp6
'        cgosIncome_cgInc_ltcg_prctg10_sec115AD_3 = templtcgBreakUp7
'        cgosIncome_cgInc_ltcg_prctg10_sec115E_b = templtcgBreakUp8
'        cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso = templtcgBreakUp7_1
'        cgosIncome_cgInc_ltcg_prctg10_sec112A = templtcgBreakUp7_2
        
        cgosIncome_cgInc_stcg_PTI30 = tempstcgBreakUp6
'Commented by Riyaz on 30/01/2026
'        cgosIncome_cgInc_stcg_PTI15 = tempstcgBreakUp7
        
        '---LTCG 12.5% NEWLY ADDED BY BINDU 26TH FEB 25
        cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2 = templtcgBreakUp4_New
        cgosIncome_cgInc_ltcg_prctg125_sec115AC_1 = templtcgBreakUp5_New
        cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1 = templtcgBreakUp6_New
        cgosIncome_cgInc_ltcg_prctg125_sec115E_b = templtcgBreakUp8_New
        cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso = templtcgBreakUp7_1_New
        cgosIncome_cgInc_ltcg_prctg125_sec112A = templtcgBreakUp7_2_New
        cgosIncome_cgInc_ltcg_PTI125 = templtcgBreakUp11_New
        cgosIncome_cgInc_ltcg_PTI125_112 = templtcgBreakUp12_New
        cgosIncome_cgInc_ltcg_PTI112_sec48 = templtcgBreakUp13_New
        cgosIncome_cgInc_ltcg_PTI125_112_B9 = templtcgBreakUp14_New
        cgosIncome_cgInc_ltcg_prctg125_sec115E_a = templtcgBreakUp15_New
        cgosIncome_cgInc_ltcg_prctg125_secProviso = templtcgBreakUp16_New '15th March
        cgosIncome_cgInc_ltcg_prctg125_sec115AD_3 = templtcgBreakUp17_New
        
        
        
        
        '---STCG 20% NEWLY ADDED BY BINDU 20TH FEB 25
        cgosIncome_cgInc_stcg_PTI20 = tempstcgBreakUp7_New
        
'Commented by Riyaz on 30/01/2026
'        cgosIncome_cgInc_ltcg_PTI20 = templtcgBreakUp10
'        cgosIncome_cgInc_ltcg_PTI10 = templtcgBreakUp11
'        cgosIncome_cgInc_ltcg_PTI10_112 = templtcgBreakUp12

 'storing diff section remaining value
 Sheet31.Unprotect Password:=getmsgstate
 'Commented by Riyaz on 30/01/2026
'    Sheet31.Range("BFLAtemp20Sec112").value = cgosIncome_cgInc_ltcg_prctg20_sec112
'    Sheet31.Range("BFLAtemp20Sec11EA").value = cgosIncome_cgInc_ltcg_prctg20_sec11EA
    'Newly added by Bindu as per DE V4 on 15th March 25
'    Sheet31.Range("BFLAtemp20SecPro").value = cgosIncome_cgInc_ltcg_prctg20_secProviso
    
    
'Commented by Riyaz on 30/01/2026
'    Sheet31.Range("BFLAtemp15Sec111a").value = cgosIncome_cgInc_stcg_prctg15_sec111a
'    Sheet31.Range("BFLAtemp15Sec115ad1bii").value = cgosIncome_cgInc_stcg_prctg15_sec115ad_1_b_ii

    '---NEWLY ADDED BY BINDU STCG 20% 20TH FEB 25
     Sheet31.Range("BFLAtemp20Sec111a").value = cgosIncome_cgInc_stcg_prctg20_sec111a
    Sheet31.Range("BFLAtemp20Sec115ad1bii").value = cgosIncome_cgInc_stcg_prctg20_sec115ad_1_b_ii

    'Sheet31.Range("BFLAtemp10SecPro").value = cgosIncome_cgInc_ltcg_prctg10_secProviso ''NOt using  commented  by Bindu as this sec changed from 10 to 20
'Commented by Riyaz on 30/01/2026
'    Sheet31.Range("BFLAtemp10Sec115ACA1").value = cgosIncome_cgInc_ltcg_prctg10_sec115ACA_1
'    Sheet31.Range("BFLAtemp10Sec115Eb").value = cgosIncome_cgInc_ltcg_prctg10_sec115E_b
'    Sheet31.Range("BFLAtemp10Sec1121c2").value = cgosIncome_cgInc_ltcg_prctg10_sec112_1_c_2
'    Sheet31.Range("BFLAtemp10Sec115AC1").value = cgosIncome_cgInc_ltcg_prctg10_sec115AC_1
'    Sheet31.Range("BFLAtemp10Sec115AD3").value = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3
'    Sheet31.Range("BFLAtemp10Sec115AD3_Proviso").value = cgosIncome_cgInc_ltcg_prctg10_sec115AD_3_Proviso
'    Sheet31.Range("BFLAtemp10Sec112A").value = cgosIncome_cgInc_ltcg_prctg10_sec112A
    
     Sheet31.Range("BFLAtempPTI30").value = cgosIncome_cgInc_stcg_PTI30
    Sheet31.Range("BFLAtemp115AD30").value = cgosIncome_cgInc_stcg_prctg30_sec115AD
'Commented by Riyaz on 30/01/2026
'    Sheet31.Range("BFLAtempPTI15").value = cgosIncome_cgInc_stcg_PTI15

    '---STCG 20% NEWLY ADDED BY BINDU 20TH FEB 25
     Sheet31.Range("BFLAtempPTI20_STCG").value = cgosIncome_cgInc_stcg_PTI20
     
 'Commented by Riyaz on 30/01/2026
'    Sheet31.Range("BFLAtempPTI20").value = cgosIncome_cgInc_ltcg_PTI20
'    Sheet31.Range("BFLAtempPTI10").value = cgosIncome_cgInc_ltcg_PTI10
'    Sheet31.Range("BFLAtempPTI10_112").value = cgosIncome_cgInc_ltcg_PTI10_112
    
    
    '--LTCG 12.5% NEWLY ADDED BY BINDU 26TH BINDU 25
    Sheet31.Range("BFLAtemp125Sec112_c").value = cgosIncome_cgInc_ltcg_prctg125_sec112_1_c_2
    Sheet31.Range("BFLAtemp125Sec_115ACA").value = cgosIncome_cgInc_ltcg_prctg125_sec115ACA_1
    Sheet31.Range("BFLAtemp125Sec_115EB").value = cgosIncome_cgInc_ltcg_prctg125_sec115E_b
    Sheet31.Range("BFLAtemp125Sec_115EA").value = cgosIncome_cgInc_ltcg_prctg125_sec115E_a
    Sheet31.Range("BFLAtemp125Sec_115AC").value = cgosIncome_cgInc_ltcg_prctg125_sec115AC_1
    Sheet31.Range("BFLAtemp125Sec_115AD_iii").value = cgosIncome_cgInc_ltcg_prctg125_sec115AD_3_Proviso
    Sheet31.Range("BFLAtemp125Sec_PTI").value = cgosIncome_cgInc_ltcg_PTI125
    Sheet31.Range("BFLAtemp125Sec_112_PTI").value = cgosIncome_cgInc_ltcg_PTI125_112
    Sheet31.Range("BFLAtemp125Sec_112").value = cgosIncome_cgInc_ltcg_PTI125_112_B9
    Sheet31.Range("BFLAtemp125Sec_112A").value = cgosIncome_cgInc_ltcg_prctg125_sec112A
    Sheet31.Range("BFLAtemp125Sec_112sec_48").value = cgosIncome_cgInc_ltcg_PTI112_sec48
    Sheet31.Range("BFLAtemp125SecPro").value = cgosIncome_cgInc_ltcg_prctg125_secProviso '15th March
    Sheet31.Range("BFLAtemp115AD125").value = cgosIncome_cgInc_ltcg_prctg125_sec115AD_3
    

Application.ScreenUpdating = True
Application.EnableEvents = True
End Function

Function setFirstToSecondMax(first As Double, second As Double)

If (first > second) Then
           setFirstToSecondMax = second
Else
           setFirstToSecondMax = first
End If
End Function



Function calcCFL()

Dim specltvIncSetOff As Double, specfdIncSetOff As Double

'CFL-Adjusted losss in BFLA
specltvIncSetOff = Sheet16.Range("busipofinclspec.BFlossPrevYrUndSameHeadSetoff2a").value
specfdIncSetOff = Sheet16.Range("busipofinclspecified.BFlossPrevYrUndSameHeadSetoff2b").value

Sheet17.Range("adjtotloss.HPLossCF9").value = Application.WorksheetFunction.Max(0, Sheet16.Range("hp.BFlossPrevYrUndSameHeadSetoff1").value)

'Sheet17.Range("adjtotloss.BusLossOthThanSpecLossCF9").value = Application.WorksheetFunction.Min(BusLossOthThanSpecLossCF8_16_1, IIf(BusLossOthThanSpecLossCF8_16_1 = 0, 0, BFlossPrevYrUndSameHeadSetoff2_16_1 + _
'                                                              adjustSpecu + adjustSpeci))
'BP-Speculative loss
'Sheet17.Range("adjtotloss.LossFrmSpecBusCF9").value = Application.WorksheetFunction.Min(LossFrmSpecBusCF8_16_1, IIf(LossFrmSpecBusCF8_16_1 = 0, 0, BFlossPrevYrUndSameHeadSetoff2a_16_1))

'BP-Specified loss
'Sheet17.Range("adjtotloss.LossFrmSpecifiedBusCF9").value = Application.WorksheetFunction.Min(LossFrmSpecifiedBusCF8_16_1, IIf(LossFrmSpecifiedBusCF8_16_1 = 0, 0, BFlossPrevYrUndSameHeadSetoff2b_16_1))

Sheet17.Range("adjtotloss.OthSrcLossRaceHorseCF9").value = Application.WorksheetFunction.Max(0, Sheet16.Range("rh.BFlossPrevYrUndSameHeadSetoff6").value)
                
End Function




