Attribute VB_Name = "mdAMTC"
Option Explicit

Public TaxSection115JC_AMTC As Variant
Public TaxOthProvisions_AMTC As Variant
Public AmtTaxCreditAvailable_AMTC As Variant

Public AssYr_AMTC0 As Variant
Public Gross4i0 As Variant
Public SetOff4i0 As Variant
Public Balance4i0 As Variant
Public AMTCredit4i0 As Variant
Public BalAmtCreditCarryFwd10 As Variant

Public AssYr_AMTC2 As Variant
Public Gross4i2 As Variant
Public SetOff4i2 As Variant
Public Balance4i2 As Variant
Public AMTCredit4i2 As Variant
Public BalAmtCreditCarryFwd12 As Variant

Public AssYr_AMTC3 As Variant
Public Gross4i3 As Variant
Public SetOff4i3 As Variant
Public Balance4i3 As Variant
Public AMTCredit4i3 As Variant
Public BalAmtCreditCarryFwd13 As Variant

Public AssYr_AMTC4 As Variant
Public Gross4i4 As Variant
Public SetOff4i4 As Variant
Public Balance4i4 As Variant
Public AMTCredit4i4 As Variant
Public BalAmtCreditCarryFwd14 As Variant

Public AssYr_AMTC6 As Variant
Public Gross4i6 As Variant
Public SetOff4i6 As Variant
Public Balance4i6 As Variant
Public AMTCredit4i6 As Variant
Public BalAmtCreditCarryFwd16 As Variant

Public AssYr_AMTC7 As Variant
Public Gross4i7 As Variant
Public SetOff4i7 As Variant
Public Balance4i7 As Variant
Public AMTCredit4i7 As Variant
Public BalAmtCreditCarryFwd17 As Variant

Public AssYr_AMTC8 As Variant
Public Gross4i8 As Variant
Public SetOff4i8 As Variant
Public Balance4i8 As Variant
Public AMTCredit4i8 As Variant
Public BalAmtCreditCarryFwd18 As Variant


Public AssYr_AMTC5 As Variant
Public Gross4i5 As Variant
Public SetOff4i5 As Variant
Public Balance4i5 As Variant
Public AMTCredit4i5 As Variant
Public BalAmtCreditCarryFwd15 As Variant

Public AssYr_AMTC As Variant
Public Gross4i As Variant
Public SetOff4i As Variant
Public Balance4i As Variant
Public AMTCredit4i As Variant
Public BalAmtCreditCarryFwd1 As Variant

Public AssYr1_AMTC As Variant
Public Gross4ii As Variant
Public Balance4ii As Variant
Public AMTCredit4ii As Variant
Public BalAmtCreditCarryFwd As Variant

Public Gross4iii As Variant
Public Balance4iii As Variant
Public BalAmtCreditCarryFwd0 As Variant


Public Gross As Variant
Public SetOff As Variant
Public Balance As Variant
Public AMTCredit As Variant
Public BalAmtCreditCarryFwd_Total As Variant

Public TaxSection115JD_AMTC As Variant
Public AmtLiabilityAvailable_AMTC As Variant

'AY_2024_25    'Malli
Public AssYr_AMTC6D, AssYr_AMTC7D, AssYr_AMTC8D As Variant
Public SetOff4i6D, SetOff4i7D, SetOff4i8D As Variant

'Public As Variant
Public AssYr_AMTC9D As Variant
Public Gross4i9 As Variant
Public Balance4i9 As Variant
Public AMTCredit4i9 As Variant
Public BalAmtCreditCarryFwd22_23 As Variant
Public SetOff4i9D As Variant

Public AssYr_AMTC10D As Variant
Public Gross4i10 As Variant
Public Balance4i10 As Variant
Public AMTCredit4i10 As Variant
Public BalAmtCreditCarryFwd23_24 As Variant
Public SetOff4i10D As Variant

'---------------------------------------
'Added by Aavula Naresh A.Y 2025-26

Public AssYr_AMTC11D As Variant
Public Balance4i11 As Variant
Public AMTCredit4i11 As Variant
Public BalAmtCreditCarryFwd24_25 As Variant
Public Gross4i11 As Variant
Public SetOff4i11D As Variant

'---------------------------------

'Added by Aavula Naresh A.Y 2026-27

Public AssYr_AMTC12D As Variant
Public Gross4i12 As Variant
Public Balance4i12 As Variant
Public AMTCredit4i12 As Variant
Public BalAmtCreditCarryFwd25_26 As Variant

Public MsgboxAMTC As Variant


Sub ValidateScheduleAMTCCLick()
ValidateScheduleAMTC
fmsgboxStatus "Sheet AMTC is OK" ', vbOKOnly, "ITR-3"
End Sub
Sub ValidateScheduleAMTC()
    MsgboxAMTC = ""
    If (Sheet23.Range("AMTC.AmtTaxCreditAvailable").value < Sheet23.Range("AMTC.AmtCreditUtilized_Total").value) Then
        fmsgboxStatus "* Sum of AMT credit utilized during the current year is " & Chr(13) & "subject to maximum of amount mentioned in 3 above" & Chr(13) & "and cannot exceed the sum of AMT Credit Brought Forward" ', vbOKOnly, "Error(s!)"
        Sheet23.Activate
        CloseMsg
    End If
    
    If Not ValidateSheetAMTC Then
         Sheet23.Activate
         fmsgboxsmall (MsgboxAMTC) ', vbOKOnly, "Error(s!)"
         CloseMsg
    End If
End Sub


Function ValidateSheetAMTC() As Boolean
    ValidateSheetAMTC = True
    If Not ValidateTaxSection115JC() Then ValidateSheetAMTC = False
    If Not ValidateTaxOthProvisions() Then ValidateSheetAMTC = False
    If Not ValidateAmtTaxCreditAvailable() Then ValidateSheetAMTC = False
    
    If Not ValidateAssessmentYearForAMTC() Then ValidateSheetAMTC = False
    If Not ValidateAmtCreditGross4i() Then ValidateSheetAMTC = False
    If Not ValidateAmtCreditSetOff4i() Then ValidateSheetAMTC = False
    If Not ValidateAmtCreditBalance4i() Then ValidateSheetAMTC = False
    If Not ValidateAmtCreditAMTCredit4i() Then ValidateSheetAMTC = False
    If Not ValidateAmtBalAmtCreditCarryFwd1() Then ValidateSheetAMTC = False
    
    If Not ValidateAssessmentYear1ForAMTC() Then ValidateSheetAMTC = False
    If Not ValidateAmtCreditGross4ii() Then ValidateSheetAMTC = False
    If Not ValidateAmtCreditBalance4ii() Then ValidateSheetAMTC = False
    If Not ValidateAmtCreditAMTCredit4ii() Then ValidateSheetAMTC = False
    If Not ValidateAmtBalAmtCreditCarryFwd() Then ValidateSheetAMTC = False

    
    If Not ValidateAmtCreditGross() Then ValidateSheetAMTC = False
    If Not ValidateAmtCreditSetOff() Then ValidateSheetAMTC = False
    If Not ValidateAmtCreditBalance() Then ValidateSheetAMTC = False
    If Not ValidateAmtCreditAMTCredit() Then ValidateSheetAMTC = False
    If Not ValidateAmtBalAmtCreditCarryFwd2() Then ValidateSheetAMTC = False
    
    If Not ValidateTaxSection115JD() Then ValidateSheetAMTC = False
    If Not ValidateAmtLiabilityAvailable() Then ValidateSheetAMTC = False
End Function

Function ValidateAmtCreditGross() As Boolean
    ValidateAmtCreditGross = True
    Dim i As Long
    Gross = Sheet23.Range("AMTC.AmtCreditFwd_Total").value
    If Len(Gross) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Total of AMT Credit Gross in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditGross = False
        Exit Function
    End If
End Function

Function ValidateAmtCreditSetOff() As Boolean
    ValidateAmtCreditSetOff = True
    Dim i As Long
    SetOff = Sheet23.Range("AMTC.AmtCreditSetOfEy_Total").value
    If Len(SetOff) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Total of AMT Credit Set-off in earlier years in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditSetOff = False
        Exit Function
    End If
End Function

Function ValidateAmtCreditBalance() As Boolean
    ValidateAmtCreditBalance = True
    Dim i As Long
    Balance = Sheet23.Range("AMTC.AmtCreditBalBroughtFwd_Total").value
    If Len(Balance) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Total of AMT Credit Balance brought forward in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditBalance = False
        Exit Function
    End If
End Function

Function ValidateAmtCreditAMTCredit() As Boolean
    ValidateAmtCreditAMTCredit = True
    Dim i As Long
    AMTCredit = Sheet23.Range("AMTC.AmtCreditUtilized_Total").value
    If Len(AMTCredit) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Total of AMT Credit Utilised during the Current Year in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditAMTCredit = False
        Exit Function
    End If
End Function

Function ValidateAmtBalAmtCreditCarryFwd2() As Boolean
    ValidateAmtBalAmtCreditCarryFwd2 = True
    Dim i As Long
    BalAmtCreditCarryFwd_Total = Sheet23.Range("AMTC.BalAmtCreditCarryFwd_Total").value
    If Len(BalAmtCreditCarryFwd_Total) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Total of Balance AMT Credit Carried Forward in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtBalAmtCreditCarryFwd2 = False
        Exit Function
    End If
End Function


Function ValidateTaxSection115JC() As Boolean
    ValidateTaxSection115JC = True
    Dim i As Long
    TaxSection115JC_AMTC = Sheet23.Range("AMTC.TaxSection115JC").value
    If Len(TaxSection115JC_AMTC) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Tax under section 115JC in assessment year 2018-19 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateTaxSection115JC = False
        Exit Function
    End If
    For i = 1 To Len(TaxSection115JC_AMTC)
        If Not IsNumeric(TaxSection115JC_AMTC) Then
            MsgboxAMTC = MsgboxAMTC & "* Tax under section 115JC in assessment year 2018-19 in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateTaxSection115JC = False
            Exit Function
        End If
    Next
End Function

Function ValidateTaxOthProvisions() As Boolean
    ValidateTaxOthProvisions = True
    Dim i As Long
    TaxOthProvisions_AMTC = Sheet23.Range("AMTC.TaxOthProvisions").value
    If Len(TaxOthProvisions_AMTC) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Tax under other provisions of the Act in assessment year 2018-19 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateTaxOthProvisions = False
        Exit Function
    End If
    For i = 1 To Len(TaxOthProvisions_AMTC)
        If Not IsNumeric(TaxOthProvisions_AMTC) Then
            MsgboxAMTC = MsgboxAMTC & "* Tax under other provisions of the Act in assessment year 2018-19 in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateTaxOthProvisions = False
            Exit Function
        End If
    Next
End Function

Function ValidateAmtTaxCreditAvailable() As Boolean
    ValidateAmtTaxCreditAvailable = True
    Dim i As Long
    AmtTaxCreditAvailable_AMTC = Sheet23.Range("AMTC.AmtTaxCreditAvailable").value
    If Len(AmtTaxCreditAvailable_AMTC) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Amount of tax against which credit is available in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtTaxCreditAvailable = False
        Exit Function
    End If
    For i = 1 To Len(AmtTaxCreditAvailable_AMTC)
        If Not IsNumeric(AmtTaxCreditAvailable_AMTC) Then
            MsgboxAMTC = MsgboxAMTC & "* Amount of tax against which credit is available in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtTaxCreditAvailable = False
            Exit Function
        End If
    Next
End Function

Function ValidateAssessmentYearForAMTC() As Boolean
    ValidateAssessmentYearForAMTC = True
    Dim i As Long
    AssYr_AMTC = Sheet23.Range("AMTC.AssYr13_14").value
    If Len(AssYr_AMTC) <> 7 Then
        MsgboxAMTC = MsgboxAMTC & "* Assessment Year in Sheet : Schedule AMTC must be exactly 7 characters " & Chr(13)
        ValidateAssessmentYearForAMTC = False
        Exit Function
    End If
    
    AssYr_AMTC0 = Sheet23.Range("AMTC.AssYr13_140").value
    If Len(AssYr_AMTC0) <> 7 Then
        MsgboxAMTC = MsgboxAMTC & "* Assessment Year in Sheet : Schedule AMTC must be exactly 7 characters " & Chr(13)
        ValidateAssessmentYearForAMTC = False
        Exit Function
    End If
    
     AssYr_AMTC2 = Sheet23.Range("AMTC.AssYr13_142").value
    If Len(AssYr_AMTC2) <> 7 Then
        MsgboxAMTC = MsgboxAMTC & "* Assessment Year in Sheet : Schedule AMTC must be exactly 7 characters " & Chr(13)
        ValidateAssessmentYearForAMTC = False
        Exit Function
    End If
    
   AssYr_AMTC3 = Sheet23.Range("AMTC.AssYr13_143").value
    If Len(AssYr_AMTC3) <> 7 Then
        MsgboxAMTC = MsgboxAMTC & "* Assessment Year in Sheet : Schedule AMTC must be exactly 7 characters " & Chr(13)
        ValidateAssessmentYearForAMTC = False
        Exit Function
    End If
    
    AssYr_AMTC4 = Sheet23.Range("AMTC.AssYr13_144").value
    If Len(AssYr_AMTC4) <> 7 Then
        MsgboxAMTC = MsgboxAMTC & "* Assessment Year in Sheet : Schedule AMTC must be exactly 7 characters " & Chr(13)
        ValidateAssessmentYearForAMTC = False
        Exit Function
    End If
        
     AssYr_AMTC5 = Sheet23.Range("AMTC.AssYr17_18").value
    If Len(AssYr_AMTC5) <> 7 Then
        MsgboxAMTC = MsgboxAMTC & "* Assessment Year in Sheet : Schedule AMTC must be exactly 7 characters " & Chr(13)
        ValidateAssessmentYearForAMTC = False
        Exit Function
    End If

'AY_2024_25 ADD  'Malli
AssYr_AMTC6D = Sheet23.Range("AMTC.AssYr20_21").value
    If Len(AssYr_AMTC6D) <> 7 Then
        MsgboxAMTC = MsgboxAMTC & "* Assessment Year in Sheet : Schedule AMTC must be exactly 7 characters " & Chr(13)
        ValidateAssessmentYearForAMTC = False
        Exit Function
    End If

AssYr_AMTC7D = Sheet23.Range("AMTC.AssYr20_21_R").value
    If Len(AssYr_AMTC7D) <> 7 Then
        MsgboxAMTC = MsgboxAMTC & "* Assessment Year in Sheet : Schedule AMTC must be exactly 7 characters " & Chr(13)
        ValidateAssessmentYearForAMTC = False
        Exit Function
    End If

AssYr_AMTC8D = Sheet23.Range("AMTC.AssYr21_22").value
    If Len(AssYr_AMTC8D) <> 7 Then
        MsgboxAMTC = MsgboxAMTC & "* Assessment Year in Sheet : Schedule AMTC must be exactly 7 characters " & Chr(13)
        ValidateAssessmentYearForAMTC = False
        Exit Function
    End If
    
AssYr_AMTC9D = Sheet23.Range("AMTC.AssYr22_23").value
    If Len(AssYr_AMTC9D) <> 7 Then
        MsgboxAMTC = MsgboxAMTC & "* Assessment Year in Sheet : Schedule AMTC must be exactly 7 characters " & Chr(13)
        ValidateAssessmentYearForAMTC = False
        Exit Function
    End If

AssYr_AMTC10D = Sheet23.Range("AMTC.AssYr22_23").value
    If Len(AssYr_AMTC10D) <> 7 Then
        MsgboxAMTC = MsgboxAMTC & "* Assessment Year in Sheet : Schedule AMTC must be exactly 7 characters " & Chr(13)
        ValidateAssessmentYearForAMTC = False
        Exit Function
    End If
'--------------------------

    AssYr_AMTC11D = Sheet23.Range("AMTC.AssYr24_25").value
      If Len(AssYr_AMTC11D) <> 7 Then
        MsgboxAMTC = MsgboxAMTC & "* Assessment Year in Sheet : Schedule AMTC must be exactly 7 characters " & Chr(13)
        ValidateAssessmentYearForAMTC = False
        Exit Function
      End If
      
      
      'Added by Aavula Naresh for AY 2026-27
    AssYr_AMTC12D = Sheet23.Range("AMTC.AssYr25_26").value
      If Len(AssYr_AMTC12D) <> 7 Then
        MsgboxAMTC = MsgboxAMTC & "* Assessment Year in Sheet : Schedule AMTC must be exactly 7 characters " & Chr(13)
        ValidateAssessmentYearForAMTC = False
        Exit Function
      End If
      

End Function


Function ValidateAmtCreditGross4i() As Boolean
    ValidateAmtCreditGross4i = True
    Dim i As Long
    Gross4i = Sheet23.Range("AMTC.AmtCreditFwd").value
    If Len(Gross4i) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross in assessment year 2016-17 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditGross4i = False
        Exit Function
    End If
    For i = 1 To Len(Gross4i)
        If Not IsNumeric(Gross4i) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross  in assessment year 2016-17  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditGross4i = False
            Exit Function
        End If
    Next
    
    Gross4i2 = Sheet23.Range("AMTC.AmtCreditFwd2").value
    If Len(Gross4i2) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross in assessment year 2014-15 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditGross4i = False
        Exit Function
    End If
    For i = 1 To Len(Gross4i2)
        If Not IsNumeric(Gross4i2) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross  in assessment year 2014-15  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditGross4i = False
            Exit Function
        End If
    Next
    
    Gross4i3 = Sheet23.Range("AMTC.AmtCreditFwd3").value
    If Len(Gross4i3) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross in assessment year 2015-16 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditGross4i = False
        Exit Function
    End If
    For i = 1 To Len(Gross4i3)
        If Not IsNumeric(Gross4i3) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross  in assessment year 2015-16  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditGross4i = False
            Exit Function
        End If
    Next
    
    Gross4i4 = Sheet23.Range("AMTC.AmtCreditFwd4").value
    If Len(Gross4i4) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross in assessment year 2018-19 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditGross4i = False
        Exit Function
    End If
    For i = 1 To Len(Gross4i4)
        If Not IsNumeric(Gross4i4) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross  in assessment year 2018-19  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditGross4i = False
            Exit Function
        End If
    Next
    
    Gross4i6 = Sheet23.Range("AMTC.AmtCreditFwd6").value
    If Len(Gross4i6) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross in assessment year 2019-20 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditGross4i = False
        Exit Function
    End If
    For i = 1 To Len(Gross4i6)
        If Not IsNumeric(Gross4i6) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross  in assessment year 2019-20  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditGross4i = False
            Exit Function
        End If
    Next
    
    Gross4i7 = Sheet23.Range("AMTC.AmtCreditFwd7").value
    If Len(Gross4i7) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross in assessment year 2020-21 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditGross4i = False
        Exit Function
    End If
    For i = 1 To Len(Gross4i7)
        If Not IsNumeric(Gross4i7) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross  in assessment year 2020-21  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditGross4i = False
            Exit Function
        End If
    Next
    
    Gross4i8 = Sheet23.Range("AMTC.AmtCreditFwd8").value
    If Len(Gross4i8) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross in assessment year 2021-22 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditGross4i = False
        Exit Function
    End If
    For i = 1 To Len(Gross4i8)
        If Not IsNumeric(Gross4i8) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross  in assessment year 2021-22  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditGross4i = False
            Exit Function
        End If
    Next
    
    Gross4i0 = Sheet23.Range("AMTC.AmtCreditFwd1").value
    If Len(Gross4i0) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross in assessment year 2013-14 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditGross4i = False
        Exit Function
    End If
    For i = 1 To Len(Gross4i0)
        If Not IsNumeric(Gross4i0) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross  in assessment year 2013-14  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditGross4i = False
            Exit Function
        End If
    Next
    
    Gross4i5 = Sheet23.Range("AMTC.AmtCreditFwd5").value
    If Len(Gross4i5) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross in assessment year 2017-18 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditGross4i = False
        Exit Function
    End If
    For i = 1 To Len(Gross4i5)
        If Not IsNumeric(Gross4i5) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross  in assessment year 2017-18  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditGross4i = False
            Exit Function
        End If
    Next
    
    'AY_2024_25 Change 'Malli
    'AMTC.AmtCreditFwd9   'AMTC.AmtCreditFwd10
    Gross4i9 = Sheet23.Range("AMTC.AmtCreditFwd9").value
    If Len(Gross4i9) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross in assessment year 2022-23 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditGross4i = False
        Exit Function
    End If
    For i = 1 To Len(Gross4i9)
        If Not IsNumeric(Gross4i9) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross  in assessment year 2022-23  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditGross4i = False
            Exit Function
        End If
    Next
    
    Gross4i10 = Sheet23.Range("AMTC.AmtCreditFwd10").value
    If Len(Gross4i10) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross in assessment year 2023-24 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditGross4i = False
        Exit Function
    End If
    For i = 1 To Len(Gross4i10)
        If Not IsNumeric(Gross4i10) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross  in assessment year 2023-24  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditGross4i = False
            Exit Function
        End If
    Next
  '----------------------------------------------------------------------
  
   Gross4i11 = Sheet23.Range("AMTC.AmtCreditFwd11").value
    If Len(Gross4i11) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross in assessment year 2024-25 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditGross4i = False
        Exit Function
    End If
    For i = 1 To Len(Gross4i11)
        If Not IsNumeric(Gross4i11) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross  in assessment year 2024-25  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditGross4i = False
            Exit Function
        End If
    Next
  '----------------------------------------------------------------------
  
  'Added by Aavula Naresh for AY 2026-27
  
   Gross4i12 = Sheet23.Range("AMTC.AmtCreditFwd12").value
    If Len(Gross4i12) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross in assessment year 2025-26 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditGross4i = False
        Exit Function
    End If
    For i = 1 To Len(Gross4i12)
        If Not IsNumeric(Gross4i12) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross  in assessment year 2025-26  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditGross4i = False
            Exit Function
        End If
    Next
  '----------------------------------------------------------------------
  
End Function

Function ValidateAmtCreditSetOff4i() As Boolean
    ValidateAmtCreditSetOff4i = True
    Dim i As Long
    SetOff4i = Sheet23.Range("AMTC.AmtCreditSetOfEy").value
    If Len(SetOff4i) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff in assessment year 2016-17 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    If (IIf(Gross4i = "", 0, Gross4i) < IIf(SetOff4i = "", 0, SetOff4i)) Then
        MsgboxAMTC = MsgboxAMTC & "* Set-off in earlier years B2 cannot be greater than Gross B1 in Sheet : AMTC" & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    For i = 1 To Len(SetOff4i)
        If Not IsNumeric(SetOff4i) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff  in assessment year 2016-17  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditSetOff4i = False
            Exit Function
        End If
    Next
    
    SetOff4i2 = Sheet23.Range("AMTC.AmtCreditSetOfEy2").value
    If Len(SetOff4i2) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff in assessment year 2014-15 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    If (IIf(Gross4i2 = "", 0, Gross4i2) < IIf(SetOff4i2 = "", 0, SetOff4i2)) Then
        MsgboxAMTC = MsgboxAMTC & "* Set-off in earlier years B2 cannot be greater than Gross B1 in Sheet : AMTC" & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    For i = 1 To Len(SetOff4i2)
        If Not IsNumeric(SetOff4i2) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff  in assessment year 2014-15  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditSetOff4i = False
            Exit Function
        End If
    Next
    
    
    SetOff4i3 = Sheet23.Range("AMTC.AmtCreditSetOfEy3").value
    If Len(SetOff4i3) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff in assessment year 2015-16 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    If (IIf(Gross4i3 = "", 0, Gross4i3) < IIf(SetOff4i3 = "", 0, SetOff4i3)) Then
        MsgboxAMTC = MsgboxAMTC & "* Set-off in earlier years B2 cannot be greater than Gross B1 in Sheet : AMTC" & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    For i = 1 To Len(SetOff4i3)
        If Not IsNumeric(SetOff4i3) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff  in assessment year 2015-16  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditSetOff4i = False
            Exit Function
        End If
    Next
    
    SetOff4i4 = Sheet23.Range("AMTC.AmtCreditSetOfEy4").value
    If Len(SetOff4i4) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff in assessment year 2018-19 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    If (IIf(Gross4i4 = "", 0, Gross4i4) < IIf(SetOff4i4 = "", 0, SetOff4i4)) Then
        MsgboxAMTC = MsgboxAMTC & "* Set-off in earlier years B2 cannot be greater than Gross B1 in Sheet : AMTC" & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    For i = 1 To Len(SetOff4i4)
        If Not IsNumeric(SetOff4i4) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff  in assessment year 2018-19  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditSetOff4i = False
            Exit Function
        End If
    Next
    
    SetOff4i5 = Sheet23.Range("AMTC.AmtCreditSetOfEy5").value
    If Len(SetOff4i5) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff in assessment year 2017-18 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    If (IIf(Gross4i5 = "", 0, Gross4i5) < IIf(SetOff4i5 = "", 0, SetOff4i5)) Then
        MsgboxAMTC = MsgboxAMTC & "* Set-off in earlier years B2 cannot be greater than Gross B1 in Sheet : AMTC" & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    For i = 1 To Len(SetOff4i5)
        If Not IsNumeric(SetOff4i5) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff  in assessment year 2017-18  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditSetOff4i = False
            Exit Function
        End If
    Next
    
    
     SetOff4i0 = Sheet23.Range("AMTC.AmtCreditSetOfEy1").value
    If Len(SetOff4i0) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff in assessment year 2013-14 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    If (IIf(Gross4i0 = "", 0, Gross4i0) < IIf(SetOff4i0 = "", 0, SetOff4i0)) Then
        MsgboxAMTC = MsgboxAMTC & "* Set-off in earlier years B2 cannot be greater than Gross B1 in Sheet : AMTC" & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    For i = 1 To Len(SetOff4i0)
        If Not IsNumeric(SetOff4i0) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff  in assessment year 2013-14  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditSetOff4i = False
            Exit Function
        End If
    Next
    
    
    
    'AY_2024_25 ADD Malli
    
    '2019-20
    SetOff4i6D = Sheet23.Range("AMTC.AmtCreditSetOfEy6").value
    If Len(SetOff4i6D) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff in assessment year 2019-20 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    If (IIf(Gross4i6 = "", 0, Gross4i6) < IIf(SetOff4i6D = "", 0, SetOff4i6D)) Then
        MsgboxAMTC = MsgboxAMTC & "* Set-off in earlier years B2 cannot be greater than Gross B1 in Sheet : AMTC" & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    For i = 1 To Len(SetOff4i6D)
        If Not IsNumeric(SetOff4i6D) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff  in assessment year 2019-20  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditSetOff4i = False
            Exit Function
        End If
    Next
    '-------------------
    '2020-21
    SetOff4i7D = Sheet23.Range("AMTC.AmtCreditSetOfEy7").value
    If Len(SetOff4i7D) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff in assessment year 2020-21 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    If (IIf(Gross4i7 = "", 0, Gross4i7) < IIf(SetOff4i7D = "", 0, SetOff4i7D)) Then
        MsgboxAMTC = MsgboxAMTC & "* Set-off in earlier years B2 cannot be greater than Gross B1 in Sheet : AMTC" & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    For i = 1 To Len(SetOff4i7D)
        If Not IsNumeric(SetOff4i7D) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff  in assessment year 2020-21  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditSetOff4i = False
            Exit Function
        End If
    Next
    
    '----------------------------------
    '2021-22
    SetOff4i8D = Sheet23.Range("AMTC.AmtCreditSetOfEy8").value
    If Len(SetOff4i8D) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff in assessment year 2021-22 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    If (IIf(Gross4i8 = "", 0, Gross4i8) < IIf(SetOff4i8D = "", 0, SetOff4i8D)) Then
        MsgboxAMTC = MsgboxAMTC & "* Set-off in earlier years B2 cannot be greater than Gross B1 in Sheet : AMTC" & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    For i = 1 To Len(SetOff4i8D)
        If Not IsNumeric(SetOff4i8D) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff  in assessment year 2021-22  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditSetOff4i = False
            Exit Function
        End If
    Next
    
    
    '----------------------------------
    SetOff4i9D = Sheet23.Range("AMTC.AmtCreditSetOfEy9").value
    If Len(SetOff4i9D) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff in assessment year 2022-23 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    If (IIf(Gross4i9 = "", 0, Gross4i9) < IIf(SetOff4i9D = "", 0, SetOff4i9D)) Then
        MsgboxAMTC = MsgboxAMTC & "* Set-off in earlier years B2 cannot be greater than Gross B1 in Sheet : AMTC" & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    For i = 1 To Len(SetOff4i9D)
        If Not IsNumeric(SetOff4i9D) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff  in assessment year 2022-23  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditSetOff4i = False
            Exit Function
        End If
    Next
    
    '----------------------------------
    
        '----------------------------------
    SetOff4i10D = Sheet23.Range("AMTC.AmtCreditSetOfEy10").value
    If Len(SetOff4i10D) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff in assessment year 2023-24 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    If (IIf(Gross4i10 = "", 0, Gross4i10) < IIf(SetOff4i10D = "", 0, SetOff4i10D)) Then
        MsgboxAMTC = MsgboxAMTC & "* Set-off in earlier years B2 cannot be greater than Gross B1 in Sheet : AMTC" & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    For i = 1 To Len(SetOff4i10D)
        If Not IsNumeric(SetOff4i10D) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff  in assessment year 2023-24  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditSetOff4i = False
            Exit Function
        End If
    Next
    
    '----------------------------------
    
    
 'Added by Aavula Naresh for AY 2026-27
      SetOff4i11D = Sheet23.Range("AMTC.AmtCreditSetOfEy11").value
    If Len(SetOff4i11D) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff in assessment year 2024-25 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    If (IIf(Gross4i11 = "", 0, Gross4i11) < IIf(SetOff4i11D = "", 0, SetOff4i11D)) Then
        MsgboxAMTC = MsgboxAMTC & "* Set-off in earlier years B2 cannot be greater than Gross B1 in Sheet : AMTC" & Chr(13)
        ValidateAmtCreditSetOff4i = False
        Exit Function
    End If
    
    For i = 1 To Len(SetOff4i11D)
        If Not IsNumeric(SetOff4i11D) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit SetOff  in assessment year 2024-25  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditSetOff4i = False
            Exit Function
        End If
    Next
    ''''''''''''''''''''
    
    
End Function

Function ValidateAmtCreditBalance4i() As Boolean
    ValidateAmtCreditBalance4i = True
    Dim i As Long
    Balance4i = Sheet23.Range("AMTC.AmtCreditBalBroughtFwd").value
    If Len(Balance4i) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance in assessment year 2016-17 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditBalance4i = False
        Exit Function
    End If
    For i = 1 To Len(Balance4i)
        If Not IsNumeric(Balance4i) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance  in assessment year 2016-17  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditBalance4i = False
            Exit Function
        End If
    Next
    
     Balance4i4 = Sheet23.Range("AMTC.AmtCreditBalBroughtFwd4").value
    If Len(Balance4i4) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance in assessment year 2018-19 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditBalance4i = False
        Exit Function
    End If
    For i = 1 To Len(Balance4i4)
        If Not IsNumeric(Balance4i4) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance  in assessment year 2018-19  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditBalance4i = False
            Exit Function
        End If
    Next
    
     Balance4i6 = Sheet23.Range("AMTC.AmtCreditBalBroughtFwd6").value
    If Len(Balance4i6) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance in assessment year 2019-20 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditBalance4i = False
        Exit Function
    End If
    For i = 1 To Len(Balance4i6)
        If Not IsNumeric(Balance4i6) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance  in assessment year 2019-20  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditBalance4i = False
            Exit Function
        End If
    Next
    
    Balance4i7 = Sheet23.Range("AMTC.AmtCreditBalBroughtFwd7").value
    If Len(Balance4i7) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance in assessment year 2020-21 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditBalance4i = False
        Exit Function
    End If
    For i = 1 To Len(Balance4i7)
        If Not IsNumeric(Balance4i7) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance  in assessment year 2020-21  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditBalance4i = False
            Exit Function
        End If
    Next
    
    Balance4i8 = Sheet23.Range("AMTC.AmtCreditBalBroughtFwd8").value
    If Len(Balance4i8) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance in assessment year 2021-22 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditBalance4i = False
        Exit Function
    End If
    For i = 1 To Len(Balance4i8)
        If Not IsNumeric(Balance4i8) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance  in assessment year 2021-22  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditBalance4i = False
            Exit Function
        End If
    Next
    
    Balance4i5 = Sheet23.Range("AMTC.AmtCreditBalBroughtFwd5").value
    If Len(Balance4i5) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance in assessment year 2017-18 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditBalance4i = False
        Exit Function
    End If
    For i = 1 To Len(Balance4i5)
        If Not IsNumeric(Balance4i5) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance  in assessment year 2017-18  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditBalance4i = False
            Exit Function
        End If
    Next
    
    Balance4i2 = Sheet23.Range("AMTC.AmtCreditBalBroughtFwd2").value
    If Len(Balance4i2) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance in assessment year 2014-15 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditBalance4i = False
        Exit Function
    End If
    For i = 1 To Len(Balance4i2)
        If Not IsNumeric(Balance4i2) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance  in assessment year 2014-15  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditBalance4i = False
            Exit Function
        End If
    Next
    
    Balance4i3 = Sheet23.Range("AMTC.AmtCreditBalBroughtFwd3").value
    If Len(Balance4i3) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance in assessment year 2015-16 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditBalance4i = False
        Exit Function
    End If
    For i = 1 To Len(Balance4i3)
        If Not IsNumeric(Balance4i3) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance  in assessment year 2015-16  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditBalance4i = False
            Exit Function
        End If
    Next
    
    Balance4i0 = Sheet23.Range("AMTC.AmtCreditBalBroughtFwd1").value
    If Len(Balance4i0) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance in assessment year 2013-14 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditBalance4i = False
        Exit Function
    End If
    For i = 1 To Len(Balance4i0)
        If Not IsNumeric(Balance4i0) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance  in assessment year 2013-14  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditBalance4i = False
            Exit Function
        End If
    Next
    
    'AY_2024_25 ADD 'Malli
    Balance4i9 = Sheet23.Range("AMTC.AmtCreditBalBroughtFwd9").value
    If Len(Balance4i9) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance in assessment year 2022-23 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditBalance4i = False
        Exit Function
    End If
    For i = 1 To Len(Balance4i9)
        If Not IsNumeric(Balance4i9) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance  in assessment year 2022-23  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditBalance4i = False
            Exit Function
        End If
    Next
    
    Balance4i10 = Sheet23.Range("AMTC.AmtCreditBalBroughtFwd10").value
    If Len(Balance4i10) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance in assessment year 2023-24 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditBalance4i = False
        Exit Function
    End If
    For i = 1 To Len(Balance4i10)
        If Not IsNumeric(Balance4i10) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance  in assessment year 2023-24  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditBalance4i = False
            Exit Function
        End If
    Next
    
    '-------------------------------------------------------
    
    
      Balance4i11 = Sheet23.Range("AMTC.AmtCreditBalBroughtFwd11").value
    If Len(Balance4i11) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance in assessment year 2024-25 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditBalance4i = False
        Exit Function
    End If
    For i = 1 To Len(Balance4i11)
        If Not IsNumeric(Balance4i11) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance  in assessment year 2024-25  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditBalance4i = False
            Exit Function
        End If
    Next
    
    
      'Added by Aavula Naresh for AY 2026-27
      
  Balance4i12 = Sheet23.Range("AMTC.AmtCreditBalBroughtFwd12").value
    If Len(Balance4i12) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance in assessment year 2025-26 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditBalance4i = False
        Exit Function
    End If
    For i = 1 To Len(Balance4i12)
        If Not IsNumeric(Balance4i12) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance  in assessment year 2025-26  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditBalance4i = False
            Exit Function
        End If
    Next
    
End Function


Function ValidateAmtCreditAMTCredit4i() As Boolean
    ValidateAmtCreditAMTCredit4i = True
    Dim i As Long
    AMTCredit4i = Sheet23.Range("AMTC.AmtCreditUtilized").value
    If Len(AMTCredit4i) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2016-17 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    For i = 1 To Len(AMTCredit4i)
        If Not IsNumeric(AMTCredit4i) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized  in assessment year 2016-17  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditAMTCredit4i = False
            Exit Function
        End If
    Next
    
    AMTCredit4i2 = Sheet23.Range("AMTC.AmtCreditUtilized2").value
    If Len(AMTCredit4i) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2014-15 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    For i = 1 To Len(AMTCredit4i2)
        If Not IsNumeric(AMTCredit4i2) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized  in assessment year 2014-15  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditAMTCredit4i = False
            Exit Function
        End If
    Next
    
    AMTCredit4i3 = Sheet23.Range("AMTC.AmtCreditUtilized3").value
    If Len(AMTCredit4i3) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2015-16 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    For i = 1 To Len(AMTCredit4i3)
        If Not IsNumeric(AMTCredit4i3) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized  in assessment year 2015-16  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditAMTCredit4i = False
            Exit Function
        End If
    Next
    
    AMTCredit4i4 = Sheet23.Range("AMTC.AmtCreditUtilized4").value
    If Len(AMTCredit4i4) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2018-19 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    For i = 1 To Len(AMTCredit4i4)
        If Not IsNumeric(AMTCredit4i4) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized  in assessment year 2018-19  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditAMTCredit4i = False
            Exit Function
        End If
    Next
    
    AMTCredit4i6 = Sheet23.Range("AMTC.AmtCreditUtilized6").value
    If Len(AMTCredit4i6) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2019-20 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    For i = 1 To Len(AMTCredit4i6)
        If Not IsNumeric(AMTCredit4i6) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized  in assessment year 2019-20  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditAMTCredit4i = False
            Exit Function
        End If
    Next
    
    AMTCredit4i7 = Sheet23.Range("AMTC.AmtCreditUtilized7").value
    If Len(AMTCredit4i7) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2020-21 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    For i = 1 To Len(AMTCredit4i7)
        If Not IsNumeric(AMTCredit4i7) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized  in assessment year 2020-21  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditAMTCredit4i = False
            Exit Function
        End If
    Next
    
    AMTCredit4i8 = Sheet23.Range("AMTC.AmtCreditUtilized8").value
    If Len(AMTCredit4i8) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2021-22 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    For i = 1 To Len(AMTCredit4i8)
        If Not IsNumeric(AMTCredit4i8) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized  in assessment year 2021-22  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditAMTCredit4i = False
            Exit Function
        End If
    Next
    
    AMTCredit4i5 = Sheet23.Range("AMTC.AmtCreditUtilized5").value
    If Len(AMTCredit4i5) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2017-18 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    For i = 1 To Len(AMTCredit4i5)
        If Not IsNumeric(AMTCredit4i5) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized  in assessment year 2017-18  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditAMTCredit4i = False
            Exit Function
        End If
    Next
    
     AMTCredit4i0 = Sheet23.Range("AMTC.AmtCreditUtilized1").value
    If Len(AMTCredit4i0) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2013-14 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    For i = 1 To Len(AMTCredit4i0)
        If Not IsNumeric(AMTCredit4i0) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized  in assessment year 2013-14  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditAMTCredit4i = False
            Exit Function
        End If
    Next
    
    
    If AMTCredit4i0 > Balance4i0 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2013-14 in Sheet : Schedule AMTC Cannot be greater than AMT Credit Brought forward. " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    
    If AMTCredit4i > Balance4i Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2016-17 in Sheet : Schedule AMTC Cannot be greater than AMT Credit Brought forward. " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    
    If AMTCredit4i2 > Balance4i2 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2014-15 in Sheet : Schedule AMTC Cannot be greater than AMT Credit Brought forward. " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    
    If AMTCredit4i3 > Balance4i3 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2015-16 in Sheet : Schedule AMTC Cannot be greater than AMT Credit Brought forward. " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    
    If AMTCredit4i4 > Balance4i4 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2018-19 in Sheet : Schedule AMTC Cannot be greater than AMT Credit Brought forward. " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    
    If AMTCredit4i5 > Balance4i5 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2017-18 in Sheet : Schedule AMTC Cannot be greater than AMT Credit Brought forward. " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    
    'AY_2024_25 ADD Malli
    
    AMTCredit4i9 = Sheet23.Range("AMTC.AmtCreditUtilized9").value
    If Len(AMTCredit4i9) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2022-23 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    For i = 1 To Len(AMTCredit4i9)
        If Not IsNumeric(AMTCredit4i9) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized  in assessment year 2022-23  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditAMTCredit4i = False
            Exit Function
        End If
    Next
    
     
    If AMTCredit4i9 > Balance4i9 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2022-23 in Sheet : Schedule AMTC Cannot be greater than AMT Credit Brought forward. " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    
    
    
    AMTCredit4i10 = Sheet23.Range("AMTC.AmtCreditUtilized10").value
    If Len(AMTCredit4i10) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2023-24 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    For i = 1 To Len(AMTCredit4i10)
        If Not IsNumeric(AMTCredit4i10) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized  in assessment year 2023-24  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditAMTCredit4i = False
            Exit Function
        End If
    Next
    If AMTCredit4i10 > Balance4i10 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2023-24 in Sheet : Schedule AMTC Cannot be greater than AMT Credit Brought forward. " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    '---------------------
    
     AMTCredit4i11 = Sheet23.Range("AMTC.AmtCreditUtilized11").value
    If Len(AMTCredit4i11) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2024-25 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    For i = 1 To Len(AMTCredit4i11)
        If Not IsNumeric(AMTCredit4i11) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized  in assessment year 2024-25  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditAMTCredit4i = False
            Exit Function
        End If
    Next
    If AMTCredit4i11 > Balance4i11 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2024-25 in Sheet : Schedule AMTC Cannot be greater than AMT Credit Brought forward. " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    
      'Added by Aavula Naresh for AY 2026-27
    
    AMTCredit4i12 = Sheet23.Range("AMTC.AmtCreditUtilized12").value
    If Len(AMTCredit4i12) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2025-26 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    For i = 1 To Len(AMTCredit4i12)
        If Not IsNumeric(AMTCredit4i12) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized  in assessment year 2025-26  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditAMTCredit4i = False
            Exit Function
        End If
    Next
    If AMTCredit4i12 > Balance4i12 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2025-26 in Sheet : Schedule AMTC Cannot be greater than AMT Credit Brought forward. " & Chr(13)
        ValidateAmtCreditAMTCredit4i = False
        Exit Function
    End If
    
    
    
End Function


Function ValidateAmtBalAmtCreditCarryFwd1() As Boolean
    ValidateAmtBalAmtCreditCarryFwd1 = True
    Dim i As Long
    BalAmtCreditCarryFwd1 = Sheet23.Range("AMTC.BalAmtCreditCarryFwd").value
    If Len(BalAmtCreditCarryFwd1) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2016-17 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtBalAmtCreditCarryFwd1 = False
        Exit Function
    End If
    For i = 1 To Len(BalAmtCreditCarryFwd1)
        If Not IsNumeric(BalAmtCreditCarryFwd1) Then
            MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2016-17  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtBalAmtCreditCarryFwd1 = False
            Exit Function
        End If
    Next
    
    BalAmtCreditCarryFwd12 = Sheet23.Range("AMTC.BalAmtCreditCarryFwd2").value
    If Len(BalAmtCreditCarryFwd12) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2014-15 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtBalAmtCreditCarryFwd1 = False
        Exit Function
    End If
    For i = 1 To Len(BalAmtCreditCarryFwd12)
        If Not IsNumeric(BalAmtCreditCarryFwd12) Then
            MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2014-15  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtBalAmtCreditCarryFwd1 = False
            Exit Function
        End If
    Next
    
    BalAmtCreditCarryFwd13 = Sheet23.Range("AMTC.BalAmtCreditCarryFwd3").value
    If Len(BalAmtCreditCarryFwd13) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2015-16 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtBalAmtCreditCarryFwd1 = False
        Exit Function
    End If
    For i = 1 To Len(BalAmtCreditCarryFwd13)
        If Not IsNumeric(BalAmtCreditCarryFwd13) Then
            MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2015-16  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtBalAmtCreditCarryFwd1 = False
            Exit Function
        End If
    Next
    
    BalAmtCreditCarryFwd14 = Sheet23.Range("AMTC.BalAmtCreditCarryFwd4").value
    If Len(BalAmtCreditCarryFwd14) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2018-19 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtBalAmtCreditCarryFwd1 = False
        Exit Function
    End If
    For i = 1 To Len(BalAmtCreditCarryFwd14)
        If Not IsNumeric(BalAmtCreditCarryFwd14) Then
            MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2018-19  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtBalAmtCreditCarryFwd1 = False
            Exit Function
        End If
    Next
    
    BalAmtCreditCarryFwd16 = Sheet23.Range("AMTC.BalAmtCreditCarryFwd6").value
    If Len(BalAmtCreditCarryFwd16) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2019-20 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtBalAmtCreditCarryFwd1 = False
        Exit Function
    End If
    For i = 1 To Len(BalAmtCreditCarryFwd16)
        If Not IsNumeric(BalAmtCreditCarryFwd16) Then
            MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2019-20  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtBalAmtCreditCarryFwd1 = False
            Exit Function
        End If
    Next
    
    BalAmtCreditCarryFwd17 = Sheet23.Range("AMTC.BalAmtCreditCarryFwd7").value
    If Len(BalAmtCreditCarryFwd17) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2020-21 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtBalAmtCreditCarryFwd1 = False
        Exit Function
    End If
    For i = 1 To Len(BalAmtCreditCarryFwd17)
        If Not IsNumeric(BalAmtCreditCarryFwd17) Then
            MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2020-21  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtBalAmtCreditCarryFwd1 = False
            Exit Function
        End If
    Next
    
    BalAmtCreditCarryFwd18 = Sheet23.Range("AMTC.BalAmtCreditCarryFwd8").value
    If Len(BalAmtCreditCarryFwd18) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2021-22 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtBalAmtCreditCarryFwd1 = False
        Exit Function
    End If
    For i = 1 To Len(BalAmtCreditCarryFwd18)
        If Not IsNumeric(BalAmtCreditCarryFwd18) Then
            MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2021-22  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtBalAmtCreditCarryFwd1 = False
            Exit Function
        End If
    Next
    
    BalAmtCreditCarryFwd15 = Sheet23.Range("AMTC.BalAmtCreditCarryFwd5").value
    If Len(BalAmtCreditCarryFwd15) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2017-18 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtBalAmtCreditCarryFwd1 = False
        Exit Function
    End If
    For i = 1 To Len(BalAmtCreditCarryFwd15)
        If Not IsNumeric(BalAmtCreditCarryFwd15) Then
            MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2017-18  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtBalAmtCreditCarryFwd1 = False
            Exit Function
        End If
    Next
    
    BalAmtCreditCarryFwd10 = Sheet23.Range("AMTC.BalAmtCreditCarryFwd1").value
    If Len(BalAmtCreditCarryFwd10) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2013-14 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtBalAmtCreditCarryFwd1 = False
        Exit Function
    End If
    For i = 1 To Len(BalAmtCreditCarryFwd10)
        If Not IsNumeric(BalAmtCreditCarryFwd10) Then
            MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2013-14  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtBalAmtCreditCarryFwd1 = False
            Exit Function
        End If
    Next
    
    
    'AY_2024_25 ADD  Malli
    
    BalAmtCreditCarryFwd22_23 = Sheet23.Range("AMTC.BalAmtCreditCarryFwd9").value
    If Len(BalAmtCreditCarryFwd22_23) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2022-23 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtBalAmtCreditCarryFwd1 = False
        Exit Function
    End If
    For i = 1 To Len(BalAmtCreditCarryFwd22_23)
        If Not IsNumeric(BalAmtCreditCarryFwd22_23) Then
            MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2022-23  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtBalAmtCreditCarryFwd1 = False
            Exit Function
        End If
    Next
    
    
    
    BalAmtCreditCarryFwd23_24 = Sheet23.Range("AMTC.BalAmtCreditCarryFwd10").value
    If Len(BalAmtCreditCarryFwd23_24) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2023-24 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtBalAmtCreditCarryFwd1 = False
        Exit Function
    End If
    For i = 1 To Len(BalAmtCreditCarryFwd23_24)
        If Not IsNumeric(BalAmtCreditCarryFwd23_24) Then
            MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2023-24  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtBalAmtCreditCarryFwd1 = False
            Exit Function
        End If
    Next
    
    '------------------
    
     BalAmtCreditCarryFwd24_25 = Sheet23.Range("AMTC.BalAmtCreditCarryFwd11").value
    If Len(BalAmtCreditCarryFwd24_25) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2024-25 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtBalAmtCreditCarryFwd1 = False
        Exit Function
    End If
    For i = 1 To Len(BalAmtCreditCarryFwd24_25)
        If Not IsNumeric(BalAmtCreditCarryFwd24_25) Then
            MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2024-25  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtBalAmtCreditCarryFwd1 = False
            Exit Function
        End If
    Next
    
   'Added by Aavula Naresh for AY 2026-27
   
 BalAmtCreditCarryFwd25_26 = Sheet23.Range("AMTC.BalAmtCreditCarryFwd12").value
    If Len(BalAmtCreditCarryFwd25_26) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2025-26 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtBalAmtCreditCarryFwd1 = False
        Exit Function
    End If
    For i = 1 To Len(BalAmtCreditCarryFwd25_26)
        If Not IsNumeric(BalAmtCreditCarryFwd25_26) Then
            MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2025-26  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtBalAmtCreditCarryFwd1 = False
            Exit Function
        End If
    Next
    
    
    
End Function


Function ValidateAssessmentYear1ForAMTC() As Boolean
    ValidateAssessmentYear1ForAMTC = True
    Dim i As Long
    AssYr1_AMTC = Sheet23.Range("AMTC.AssYr14_15").value


End Function

Function ValidateAmtCreditGross4ii() As Boolean
    ValidateAmtCreditGross4ii = True
    Dim i As Long
    Gross4ii = Sheet23.Range("AMTC.AmtCreditFwd0").value
    If Len(Gross4ii) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross in assessment year 2023-24 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditGross4ii = False
        Exit Function
    End If
    For i = 1 To Len(Gross4ii)
        If Not IsNumeric(Gross4ii) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Gross  in assessment year 2023-24  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditGross4ii = False
            Exit Function
        End If
    Next
    
    If Gross4ii > WorksheetFunction.Max(0, TaxSection115JC_AMTC - TaxOthProvisions_AMTC) Then
        MsgboxAMTC = MsgboxAMTC & "* The amount of Gross AMT Credit (B1) 2023-24 can not exceed (1-2)" & Chr(13)
        ValidateAmtCreditGross4ii = False
        Exit Function
    End If
End Function

Function ValidateAmtCreditBalance4ii() As Boolean
    ValidateAmtCreditBalance4ii = True
    Dim i As Long
    Balance4ii = Sheet23.Range("AMTC.AmtCreditBalBroughtFwd0").value
    If Len(Balance4ii) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance in assessment year 2019-20 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditBalance4ii = False
        Exit Function
    End If
    For i = 1 To Len(Balance4ii)
        If Not IsNumeric(Balance4ii) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Balance  in assessment year 2019-20  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditBalance4ii = False
            Exit Function
        End If
    Next
End Function


Function ValidateAmtCreditAMTCredit4ii() As Boolean
    ValidateAmtCreditAMTCredit4ii = True
    Dim i As Long
    AMTCredit4ii = Sheet23.Range("AMTC.AmtCreditUtilized0").value
    If Len(AMTCredit4ii) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized in assessment year 2019-20 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtCreditAMTCredit4ii = False
        Exit Function
    End If
    For i = 1 To Len(AMTCredit4ii)
        If Not IsNumeric(AMTCredit4ii) Then
            MsgboxAMTC = MsgboxAMTC & "* AMT credit Utilized  in assessment year 2019-20  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtCreditAMTCredit4ii = False
            Exit Function
        End If
    Next
End Function


Function ValidateAmtBalAmtCreditCarryFwd() As Boolean
    ValidateAmtBalAmtCreditCarryFwd = True
    Dim i As Long
    BalAmtCreditCarryFwd = Sheet23.Range("AMTC.BalAmtCreditCarryFwd0").value
    If Len(BalAmtCreditCarryFwd1) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2019-20 in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtBalAmtCreditCarryFwd = False
        Exit Function
    End If
    For i = 1 To Len(BalAmtCreditCarryFwd)
        If Not IsNumeric(BalAmtCreditCarryFwd) Then
            MsgboxAMTC = MsgboxAMTC & "* Balance AMT credit carried forward  in assessment year 2019-20  in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateAmtBalAmtCreditCarryFwd = False
            Exit Function
        End If
    Next
End Function


Function ValidateTaxSection115JD() As Boolean
    ValidateTaxSection115JD = True
    Dim i As Long
    TaxSection115JD_AMTC = Sheet23.Range("AMTC.TaxSection115JD").value
    If Len(TaxSection115JD_AMTC) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Amount of tax credit under section 115JD in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateTaxSection115JD = False
        Exit Function
    End If
    For i = 1 To Len(TaxSection115JD_AMTC)
        If Not IsNumeric(TaxSection115JD_AMTC) Then
            MsgboxAMTC = MsgboxAMTC & "* Amount of tax credit under section 115JD in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
            ValidateTaxSection115JD = False
            Exit Function
        End If
    Next
End Function

Function ValidateAmtLiabilityAvailable() As Boolean
    ValidateAmtLiabilityAvailable = True
    Dim i As Long
    AmtLiabilityAvailable_AMTC = Sheet23.Range("AMTC.AmtLiabilityAvailable").value
    If Len(AmtLiabilityAvailable_AMTC) > 14 Then
        MsgboxAMTC = MsgboxAMTC & "* Amount of AMT liability available in Sheet : Schedule AMTC should be at most 14 digits " & Chr(13)
        ValidateAmtLiabilityAvailable = False
        Exit Function
    End If
    For i = 1 To Len(AmtLiabilityAvailable_AMTC)
    If Not IsNumeric(AmtLiabilityAvailable_AMTC) Then
        MsgboxAMTC = MsgboxAMTC & "* Amount of AMT liability available in Sheet : Schedule AMTC must contain only digits from 0 to 9" & Chr(13)
        ValidateAmtLiabilityAvailable = False
        Exit Function
    End If
    Next
End Function




