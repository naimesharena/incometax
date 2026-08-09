Attribute VB_Name = "mdCFL"
Option Explicit

Public DateOfFiling0_16 As Variant
Public DateOfFiling1_16 As Variant
Public DateOfFiling2_16 As Variant
Public DateOfFiling3_16 As Variant
Public DateOfFiling4_16 As Variant
Public DateOfFiling5_16 As Variant
Public DateOfFiling6_16 As Variant
Public DateOfFiling7_16 As Variant
Public DateOfFiling8_16 As Variant
Public DateOfFiling9_16 As Variant
Public DateOfFiling10_16 As Variant
Public DateOfFiling11_16 As Variant
Public DateOfFiling12_16 As Variant

Public DateOfFiling13_16 As Variant

Public DateOfFiling14_16 As Variant

Public DateOfFiling15_16 As Variant


Public msgValidatesheetCFL As String
Public LengthCheckValue As Boolean

Public datefield0 As Variant
Public datefield1 As Variant
Public datefield2 As Variant
Public datefield3 As Variant
Public datefield4 As Variant
Public datefield5 As Variant
Public datefield6 As Variant
Public datefield7 As Variant
Public datefield8 As Variant
Public datefield9 As Variant
Public datefield10 As Variant
Public datefield11 As Variant

Dim HPLossCurrentYr As Variant
Dim HPLossCurrentYrCalculated As Variant
Dim STCGLossCurrentYr As Variant
Dim STCGLossCurrentYrCalculated As Variant
Dim LTCGLossCurrentYr As Variant
Dim LTCGLossCurrentYrCalculated As Variant
Dim OthrLossCurrentYr As Variant
Dim OthrLossCurrentYrCalculated As Variant
Dim SpecifiedBusLossCurrentYr As Variant
Dim SpecifiedBusLossCurrentYrCalculated As Variant
Dim SpecBusLossCurrentYr As Variant
Dim SpecBusLossCurrentYrCalculated As Variant
Dim BusLossCurrentYr As Variant
Dim BusLossCurrentYrCalculated As Variant

Dim STCGLoss1 As Variant
Dim STCGLoss2 As Variant
Dim STCGLoss3 As Variant
Dim STCGLoss4 As Variant
Dim STCGLoss5 As Variant    ''Added by Sai(07-05-2025)
Dim LTCGLoss1 As Variant
Dim LTCGLoss2 As Variant
Dim LTCGLoss3 As Variant
Dim LTCGLoss4 As Variant    ''Added by Sai(07-05-2025)



Sub ValidatecflClick()
ValidateCFL
fmsgboxStatus "Sheet CFL is oK"
End Sub

Sub ValidateCFL()
    If Not ValidatesheetCFL Then
        Sheet17.Activate
        fmsgboxsmall msgValidatesheetCFL
        CloseMsg
    End If
End Sub

Function msgbox16(strmsg As String) As String
    msgValidatesheetCFL = msgValidatesheetCFL & strmsg & Chr(13)
End Function

Function ValidatesheetCFL() As Boolean
    ValidatesheetCFL = True
    
    If Not ValidateDateOfFiling0_16() Then ValidatesheetCFL = False
    If Not ValidateDateOfFiling1_16() Then ValidatesheetCFL = False
    If Not ValidateDateOfFiling2_16() Then ValidatesheetCFL = False
    If Not ValidateDateOfFiling3_16() Then ValidatesheetCFL = False
    If Not ValidateDateOfFiling4_16() Then ValidatesheetCFL = False
    If Not ValidateDateOfFiling5_16() Then ValidatesheetCFL = False
    If Not ValidateDateOfFiling6_16() Then ValidatesheetCFL = False
    If Not ValidateDateOfFiling7_16() Then ValidatesheetCFL = False
    If Not ValidateDateOfFiling8_16() Then ValidatesheetCFL = False
    If Not ValidateDateOfFiling9_16() Then ValidatesheetCFL = False
    If Not ValidateDateOfFiling10_16() Then ValidatesheetCFL = False
    If Not ValidateDateOfFiling11_16() Then ValidatesheetCFL = False
    If Not ValidateDateOfFiling12_16() Then ValidatesheetCFL = False
    'AY_2024_25 ADD 'Malli
    If Not ValidateDateOfFiling13_16() Then ValidatesheetCFL = False   'Malli
    
    If Not ValidateDateOfFiling14_16() Then ValidatesheetCFL = False 'naresh
    
    If Not ValidateDateOfFiling15_16() Then ValidatesheetCFL = False 'naresh 2026-27

    
    If Not ValidateFilingDates() Then ValidatesheetCFL = False
    
    If Not ValidateLTCGLossCurrentYr() Then ValidatesheetCFL = False
    If Not ValidateSTCGLossCurrentYr() Then ValidatesheetCFL = False
    If Not ValidateHPLossCurrentYr() Then ValidatesheetCFL = False
    If Not ValidateBusLossCurrentYr() Then ValidatesheetCFL = False
    If Not ValidateSpecBusLossCurrentYr() Then ValidatesheetCFL = False
    If Not ValidateSpecifiedBusLossCurrentYr() Then ValidatesheetCFL = False
    If Not ValidateOthrLossCurrentYear() Then ValidatesheetCFL = False
    
    If Not ValidateNonEditablefieldsCFL() Then ValidatesheetCFL = False
    
End Function

Function ValidateDateOfFiling0_16() As Boolean
    ValidateDateOfFiling0_16 = True
    
    DateOfFiling0_16 = Sheet17.Range("yr1999.DateOfFiling").value
    If Not CheckDateddmmyyyy(DateOfFiling0_16) Then
        ValidateDateOfFiling0_16 = False
        msgbox16 ("* DateOfFiling1 in Sheet : Schedule CFL  is invalid")
        Exit Function
    Else
        DateOfFiling0_16 = Dformat(Sheet17.Range("yr1999.DateOfFiling"), "")
    End If
    
    If Len(DateOfFiling0_16) = 10 Then
        If Not CheckDateMinDDMMYYYY(DateOfFiling0_16, 1, 4, 2010, "") Then
            ValidateDateOfFiling0_16 = False
            msgbox16 ("* Date of Filing in Sch CFL  must not be less than 01/04/2010")
            Exit Function
        Else
            DateOfFiling0_16 = Dformat(Sheet17.Range("yr1999.DateOfFiling"), "")
        End If
    End If
End Function
                                                                    
Function ValidateDateOfFiling1_16() As Boolean
    Dim i As Long
    
    ValidateDateOfFiling1_16 = True
    DateOfFiling1_16 = Sheet17.Range("yr2000.DateOfFiling1").value
    
    If Not CheckDateddmmyyyy(DateOfFiling1_16) Then
        ValidateDateOfFiling1_16 = False
        msgbox16 ("* DateOfFiling2 in Sheet : Schedule CFL  is invalid")
        Exit Function
    Else
        DateOfFiling1_16 = Dformat(Sheet17.Range("yr2000.DateOfFiling1"), "")
    End If
    
    If Len(DateOfFiling1_16) = 10 Then
        If Not CheckDateMinDDMMYYYY(DateOfFiling1_16, 1, 4, 2011, "") Then
            ValidateDateOfFiling1_16 = False
            msgbox16 ("* Date of Filing (2) in Sch CFL  must not be less than 01/04/2011")
            Exit Function
        Else
            DateOfFiling1_16 = Dformat(Sheet17.Range("yr2000.DateOfFiling1"), "")
        End If
    End If
End Function
                                                              
Function ValidateDateOfFiling2_16() As Boolean
    Dim i As Long
    
    ValidateDateOfFiling2_16 = True
    DateOfFiling2_16 = Sheet17.Range("yr2001.DateOfFiling2").value
    
    If Not CheckDateddmmyyyy(DateOfFiling2_16) Then
        ValidateDateOfFiling2_16 = False
        msgbox16 ("* DateOfFiling3 in Sheet : Schedule CFL  is invalid")
        Exit Function
    Else
        DateOfFiling2_16 = Dformat(Sheet17.Range("yr2001.DateOfFiling2"), "")
    End If
    
    If Len(DateOfFiling2_16) = 10 Then
        If Not CheckDateMinDDMMYYYY(DateOfFiling2_16, 1, 4, 2012, "") Then
            ValidateDateOfFiling2_16 = False
            msgbox16 ("* Date of Filing (3) in Sch CFL  must not be less than 01/04/2012")
            Exit Function
        Else
            DateOfFiling2_16 = Dformat(Sheet17.Range("yr2001.DateOfFiling2"), "")
        End If
    End If
End Function

Function ValidateDateOfFiling3_16() As Boolean
    ValidateDateOfFiling3_16 = True
    DateOfFiling3_16 = Sheet17.Range("yr2002.DateOfFiling3").value

    If Not CheckDateddmmyyyy(DateOfFiling3_16) Then
        ValidateDateOfFiling3_16 = False
        msgbox16 ("* DateOfFiling4 in Sheet : Schedule CFL  is invalid")
        Exit Function
    Else
        DateOfFiling3_16 = Dformat(Sheet17.Range("yr2002.DateOfFiling3"), "")
    End If
    
    If Len(DateOfFiling3_16) = 10 Then
        If Not CheckDateMinDDMMYYYY(DateOfFiling3_16, 1, 4, 2013, "") Then
            ValidateDateOfFiling3_16 = False
            msgbox16 ("* Date of Filing (4) in Sch CFL  must not be less than 01/04/2013")
            Exit Function
        Else
            DateOfFiling3_16 = Dformat(Sheet17.Range("yr2002.DateOfFiling3"), "")
        End If
    End If
End Function

Function ValidateDateOfFiling4_16() As Boolean
    ValidateDateOfFiling4_16 = True
    
    DateOfFiling4_16 = Sheet17.Range("yr2003.DateOfFiling4").value
    
    If Not CheckDateddmmyyyy(DateOfFiling4_16) Then
        ValidateDateOfFiling4_16 = False
        msgbox16 ("* DateOfFiling5 in Sheet : Schedule CFL  is invalid")
        Exit Function
    Else
        DateOfFiling4_16 = Dformat(Sheet17.Range("yr2003.DateOfFiling4"), "")
    End If
    
    If Len(DateOfFiling4_16) = 10 Then
        If Not CheckDateMinDDMMYYYY(DateOfFiling4_16, 1, 4, 2014, "") Then
            ValidateDateOfFiling4_16 = False
            msgbox16 ("* Date of Filing (5) in Sch CFL  must not be less than 01/04/2014")
            Exit Function
        Else
            DateOfFiling4_16 = Dformat(Sheet17.Range("yr2003.DateOfFiling4"), "")
        End If
    End If
End Function
                                                                 
Function ValidateDateOfFiling5_16() As Boolean

    ValidateDateOfFiling5_16 = True
    DateOfFiling5_16 = Sheet17.Range("yr2004.DateOfFiling5").value
    
    If Not CheckDateddmmyyyy(DateOfFiling5_16) Then
        ValidateDateOfFiling5_16 = False
        msgbox16 ("* DateOfFiling6 in Sheet : Schedule CFL  is invalid")
        Exit Function
    Else
        DateOfFiling5_16 = Dformat(Sheet17.Range("yr2004.DateOfFiling5"), "")
    End If
    
    If Len(DateOfFiling5_16) = 10 Then
        If Not CheckDateMinDDMMYYYY(DateOfFiling5_16, 1, 4, 2015, "") Then
            ValidateDateOfFiling5_16 = False
            msgbox16 ("* Date of Filing (6) in Sch CFL  must not be less than 01/04/2015")
            Exit Function
        Else
            DateOfFiling5_16 = Dformat(Sheet17.Range("yr2004.DateOfFiling5"), "")
        End If
    End If
End Function
                                                                   
Function ValidateDateOfFiling6_16() As Boolean
    ValidateDateOfFiling6_16 = True
    DateOfFiling6_16 = Sheet17.Range("yr2005.DateOfFiling6").value
    
    If Not CheckDateddmmyyyy(DateOfFiling6_16) Then
        ValidateDateOfFiling6_16 = False
        
        msgbox16 ("* DateOfFiling7 in Sheet : Schedule CFL  is invalid")
        Exit Function
    Else
        DateOfFiling6_16 = Dformat(Sheet17.Range("yr2005.DateOfFiling6"), "")
    End If
    
    If Len(DateOfFiling6_16) = 10 Then
        If Not CheckDateMinDDMMYYYY(DateOfFiling6_16, 1, 4, 2016, "") Then
            ValidateDateOfFiling6_16 = False
            msgbox16 ("* Date of Filing (7) in Sch CFL  must not be less than 01/04/2016")
            Exit Function
        Else
            DateOfFiling6_16 = Dformat(Sheet17.Range("yr2005.DateOfFiling6"), "")
        End If
    End If
End Function
                                                                 
Function ValidateDateOfFiling7_16() As Boolean
    Dim i As Long
    
    ValidateDateOfFiling7_16 = True
    DateOfFiling7_16 = Sheet17.Range("yr2006.DateOfFiling7").value
    
    If Not CheckDateddmmyyyy(DateOfFiling7_16) Then
        ValidateDateOfFiling7_16 = False
        msgbox16 ("* DateOfFiling8 in Sheet : Schedule CFL  is invalid")
        Exit Function
    Else
        DateOfFiling7_16 = Dformat(Sheet17.Range("yr2006.DateOfFiling7"), "")
    End If
    
    If Len(DateOfFiling7_16) = 10 Then
        If Not CheckDateMinDDMMYYYY(DateOfFiling7_16, 1, 4, 2017, "") Then
            ValidateDateOfFiling7_16 = False
            msgbox16 ("* Date of Filing (8) in Sch CFL  must not be less than 01/04/2017")
            Exit Function
        Else
            DateOfFiling7_16 = Dformat(Sheet17.Range("yr2006.DateOfFiling7"), "")
        End If
    End If
End Function

Function ValidateDateOfFiling8_16() As Boolean
    Dim i As Long
    
    ValidateDateOfFiling8_16 = True
    DateOfFiling8_16 = Sheet17.Range("yr2006.DateOfFiling9").value
    
    If Not CheckDateddmmyyyy(DateOfFiling8_16) Then
        ValidateDateOfFiling8_16 = False
        msgbox16 ("* DateOfFiling9 in Sheet : Schedule CFL  is invalid")
        Exit Function
    Else
        DateOfFiling8_16 = Dformat(Sheet17.Range("yr2006.DateOfFiling9"), "")
    End If
    
    If Len(DateOfFiling8_16) = 10 Then
        If Not CheckDateMinDDMMYYYY(DateOfFiling8_16, 1, 4, 2018, "") Then
            ValidateDateOfFiling8_16 = False
            msgbox16 ("* Date of Filing (9) in Sch CFL  must not be less than 01/04/2018")
            Exit Function
        Else
            DateOfFiling8_16 = Dformat(Sheet17.Range("yr2006.DateOfFiling9"), "")
        End If
    End If
End Function

Function ValidateDateOfFiling9_16() As Boolean
    Dim i As Long
    
    ValidateDateOfFiling9_16 = True
    DateOfFiling9_16 = Sheet17.Range("yr2007.DateOfFiling9").value
    
    If Not CheckDateddmmyyyy(DateOfFiling9_16) Then
        ValidateDateOfFiling9_16 = False
        msgbox16 ("* DateOfFiling10 in Sheet : Schedule CFL  is invalid")
        Exit Function
    Else
        DateOfFiling9_16 = Dformat(Sheet17.Range("yr2007.DateOfFiling9"), "")
    End If
    
    If Len(DateOfFiling9_16) = 10 Then
        If Not CheckDateMinDDMMYYYY(DateOfFiling9_16, 1, 4, 2019, "") Then
            ValidateDateOfFiling9_16 = False
            msgbox16 ("* Date of Filing (10) in Sch CFL  must not be less than 01/04/2019")
            Exit Function
        Else
            DateOfFiling9_16 = Dformat(Sheet17.Range("yr2007.DateOfFiling9"), "")
        End If
    End If
End Function
Function ValidateDateOfFiling10_16() As Boolean
    Dim i As Long
    
    ValidateDateOfFiling10_16 = True
    DateOfFiling10_16 = Sheet17.Range("yr2020.DateOfFiling10").value
    
    If Not CheckDateddmmyyyy(DateOfFiling10_16) Then
        ValidateDateOfFiling10_16 = False
        msgbox16 ("* DateOfFiling11 in Sheet : Schedule CFL  is invalid")
        Exit Function
    Else
        DateOfFiling10_16 = Dformat(Sheet17.Range("yr2020.DateOfFiling10"), "")
    End If
    
    If Len(DateOfFiling10_16) = 10 Then
        If Not CheckDateMinDDMMYYYY(DateOfFiling10_16, 1, 4, 2020, "") Then
            ValidateDateOfFiling10_16 = False
            msgbox16 ("* Date of Filing (11) in Sch CFL  must not be less than 01/04/2020")
            Exit Function
        Else
            DateOfFiling10_16 = Dformat(Sheet17.Range("yr2020.DateOfFiling10"), "")
        End If
    End If
End Function
Function ValidateDateOfFiling11_16() As Boolean
    Dim i As Long
    
    ValidateDateOfFiling11_16 = True
    DateOfFiling11_16 = Sheet17.Range("yr2021.DateOfFiling11").value
    
    If Not CheckDateddmmyyyy(DateOfFiling11_16) Then
        ValidateDateOfFiling11_16 = False
        msgbox16 ("* DateOfFiling12 in Sheet : Schedule CFL  is invalid")
        Exit Function
    Else
        DateOfFiling11_16 = Dformat(Sheet17.Range("yr2021.DateOfFiling11"), "")
    End If
    
    If Len(DateOfFiling11_16) = 10 Then
        If Not CheckDateMinDDMMYYYY(DateOfFiling11_16, 1, 4, 2021, "") Then
            ValidateDateOfFiling11_16 = False
            msgbox16 ("* Date of Filing (12) in Sch CFL  must not be less than 01/04/2021")
            Exit Function
        Else
            DateOfFiling11_16 = Dformat(Sheet17.Range("yr2021.DateOfFiling11"), "")
        End If
    End If
End Function


Function ValidateFilingDates() As Boolean
    ValidateFilingDates = True
    
    Dim amount0 As Variant
    Dim amount1 As Variant
    Dim amount2 As Variant
    Dim amount3 As Variant
    Dim amount4 As Variant
    Dim amount5 As Variant
    Dim amount6 As Variant
    Dim amount7 As Variant
    Dim amount8 As Variant
    Dim amount9 As Variant
    Dim amount10 As Variant
    Dim amount11 As Variant
    
    datefield0 = Sheet17.Range("yr1999.DateOfFiling").value
    amount0 = Sheet17.Range("yr1999.BusLossSpecBus").value
    
    If Len(datefield0) = 0 Then
        If amount0 > 0 Then
            ValidateFilingDates = False
'Change-11.06.2023.101 -> SIT-53956
'            msgbox16 ("* DateOfFiling(1) in Sheet : Schedule CFL  is Mandatory ")
            msgbox16 ("* Date of filing for Sl no (i) in Sheet : Schedule CFL  is Mandatory ")
            Exit Function
        End If
    End If
    
    datefield1 = Sheet17.Range("yr2000.DateOfFiling1").value
    amount1 = Sheet17.Range("yr2003.LossFrmSpecifiedBusCF9").value '+ Sheet17.Range("yr2000.LTCGLossCF1").value '+ Sheet17.Range("yr2000.STCGLossCF1").value
    
    If Len(datefield1) = 0 Then
        If amount1 > 0 Then
            ValidateFilingDates = False
'            msgbox16 ("* DateOfFiling(2) in Sheet : Schedule CFL  is Mandatory ")
            msgbox16 ("* Date of filing for Sl no (ii) in Sheet : Schedule CFL  is Mandatory ")
            Exit Function
        End If
    End If
       
    datefield2 = Sheet17.Range("yr2001.DateOfFiling2").value
    amount2 = Sheet17.Range("yr2001.HPLossCF2").value + Sheet17.Range("yr2001.BusLossOthThanSpecLossCF2").value + Sheet17.Range("yr2003.LossFrmSpecifiedBusCF8").value + _
                Sheet17.Range("yr2001.STCGLossCF2").value + Sheet17.Range("yr2001.LTCGLossCF2").value
    
    If Len(datefield2) = 0 Then
        If amount2 > 0 Then
            ValidateFilingDates = False
'            msgbox16 ("* DateOfFiling(3) in Sheet : Schedule CFL  is Mandatory ")
            msgbox16 ("* Date of filing for Sl no (iii) in Sheet : Schedule CFL  is Mandatory ")
            Exit Function
        End If
    End If
    
    datefield3 = Sheet17.Range("yr2002.DateOfFiling3").value
    amount3 = Sheet17.Range("yr2002.HPLossCF3").value + Sheet17.Range("yr2002.BusLossOthThanSpecLossCF3").value + Sheet17.Range("yr2003.LossFrmSpecifiedBusCF3").value + _
                Sheet17.Range("yr2002.STCGLossCF3").value + Sheet17.Range("yr2002.LTCGLossCF3").value
    
    If Len(datefield3) = 0 Then
        If amount3 > 0 Then
            ValidateFilingDates = False
'            msgbox16 ("* DateOfFiling(4) in Sheet : Schedule CFL  is Mandatory ")
            msgbox16 ("* Date of filing for Sl no (iv) in Sheet : Schedule CFL  is Mandatory ")
            Exit Function
        End If
    End If
    
    datefield4 = Sheet17.Range("yr2003.DateOfFiling4").value
    amount4 = Sheet17.Range("yr2003.HPLossCF4").value + Sheet17.Range("yr2003.BusLossOthThanSpecLossCF4").value + Sheet17.Range("yr2003.LossFrmSpecifiedBusCF4").value + _
                Sheet17.Range("yr2003.STCGLossCF4").value + Sheet17.Range("yr2003.LTCGLossCF4").value
    
    If Len(datefield4) = 0 Then
        If amount4 > 0 Then
            ValidateFilingDates = False
'            msgbox16 ("* DateOfFiling(5) in Sheet : Schedule CFL  is Mandatory ")
            msgbox16 ("* Date of filing for Sl no (v) in Sheet : Schedule CFL  is Mandatory ")
            Exit Function
        End If
    End If
    
    datefield5 = Sheet17.Range("yr2004.DateOfFiling5").value
    amount5 = Sheet17.Range("yr2004.HPLossCF5").value + Sheet17.Range("yr2004.BusLossOthThanSpecLossCF5").value + Sheet17.Range("yr2004.LossFrmSpecifiedBusCF5").value + _
                Sheet17.Range("yr2004.STCGLossCF5").value + Sheet17.Range("yr2004.LTCGLossCF5").value '+ Sheet17.Range("yr2004.OthSrcLossRaceHorseCF5").value
    
    If Len(datefield5) = 0 Then
        If amount5 > 0 Then
            ValidateFilingDates = False
'            msgbox16 ("* DateOfFiling(6) in Sheet : Schedule CFL  is Mandatory ")
            msgbox16 ("* Date of filing for Sl no (vi) in Sheet : Schedule CFL  is Mandatory ")
            Exit Function
        End If
    End If
    
    datefield6 = Sheet17.Range("yr2005.DateOfFiling6").value
    amount6 = Sheet17.Range("yr2005.HPLossCF6").value + Sheet17.Range("yr2005.BusLossOthThanSpecLossCF6").value + Sheet17.Range("yr2005.LossFrmSpecBusCF6").value + _
                Sheet17.Range("yr2005.LossFrmSpecifiedBusCF6").value + Sheet17.Range("yr2005.STCGLossCF6").value + Sheet17.Range("yr2005.LTCGLossCF6").value + Sheet17.Range("yr2005.OthSrcLossRaceHorseCF6").value
    
    If Len(datefield6) = 0 Then
        If amount6 > 0 Then
            ValidateFilingDates = False
'            msgbox16 ("* DateOfFiling(7) in Sheet : Schedule CFL  is Mandatory ")
            msgbox16 ("* Date of filing for Sl no (vii) in Sheet : Schedule CFL  is Mandatory ")
            Exit Function
        End If
    End If
    
    datefield7 = Sheet17.Range("yr2006.DateOfFiling7").value
    amount7 = Sheet17.Range("yr2006.HPLossCF7").value + Sheet17.Range("yr2006.BusLossOthThanSpecLossCF7").value + Sheet17.Range("yr2006.LossFrmSpecBusCF7").value + _
                Sheet17.Range("yr2006.LossFrmSpecifiedBusCF7").value + Sheet17.Range("yr2006.STCGLossCF7").value + Sheet17.Range("yr2006.LTCGLossCF7").value + Sheet17.Range("yr2006.OthSrcLossRaceHorseCF7").value
    
    If Len(datefield7) = 0 Then
        If amount7 > 0 Then
            ValidateFilingDates = False
'            msgbox16 ("* DateOfFiling(8) in Sheet : Schedule CFL  is Mandatory ")
            msgbox16 ("* Date of filing for Sl no (viii) in Sheet : Schedule CFL  is Mandatory ")
            Exit Function
        End If
    End If
    
    datefield8 = Sheet17.Range("yr2006.DateOfFiling9").value
    amount8 = Sheet17.Range("yr2006.HPLossCF9").value + Sheet17.Range("yr2006.BusLossOthThanSpecLossCF9").value + Sheet17.Range("yr2006.LossFrmSpecBusCF9").value + _
                Sheet17.Range("yr2006.LossFrmSpecifiedBusCF9").value + Sheet17.Range("yr2006.STCGLossCF9").value + Sheet17.Range("yr2006.LTCGLossCF9").value + Sheet17.Range("yr2006.OthSrcLossRaceHorseCF9").value
    
    If Len(datefield8) = 0 Then
        If amount8 > 0 Then
            ValidateFilingDates = False
'            msgbox16 ("* DateOfFiling(9) in Sheet : Schedule CFL  is Mandatory ")
            msgbox16 ("* Date of filing for Sl no (ix) in Sheet : Schedule CFL  is Mandatory ")
            Exit Function
        End If
    End If
    
    datefield9 = Sheet17.Range("yr2007.DateOfFiling9").value
    amount9 = Sheet17.Range("yr2007.HPLossCF9").value + Sheet17.Range("yr2007.BusLossOthThanSpecLossCF9").value + Sheet17.Range("yr2007.LossFrmSpecBusCF9").value + _
                Sheet17.Range("yr2007.LossFrmSpecifiedBusCF9").value + Sheet17.Range("yr2007.STCGLossCF9").value + Sheet17.Range("yr2007.LTCGLossCF9").value + Sheet17.Range("yr2007.OthSrcLossRaceHorseCF9").value
    
    If Len(datefield9) = 0 Then
        If amount9 > 0 Then
            ValidateFilingDates = False
'            msgbox16 ("* DateOfFiling(10) in Sheet : Schedule CFL  is Mandatory ")
            msgbox16 ("* Date of filing for Sl no (x) in Sheet : Schedule CFL  is Mandatory ")
            Exit Function
        End If
    End If
    
    datefield10 = Sheet17.Range("yr2020.DateOfFiling10").value
    amount10 = Sheet17.Range("yr2020.HPLossCF10").value + Sheet17.Range("yr2020.BusLossOthThanSpecLossCF10").value + Sheet17.Range("yr2020.LossFrmSpecBusCF10").value + _
                Sheet17.Range("yr2020.LossFrmSpecifiedBusCF10").value + Sheet17.Range("yr2020.STCGLossCF10").value + Sheet17.Range("yr2020.LTCGLossCF10").value + Sheet17.Range("yr2020.OthSrcLossRaceHorseCF10").value
    
    If Len(datefield10) = 0 Then
        If amount10 > 0 Then
            ValidateFilingDates = False
'            msgbox16 ("* DateOfFiling(11) in Sheet : Schedule CFL  is Mandatory ")
            msgbox16 ("* Date of filing for Sl no (xi) in Sheet : Schedule CFL  is Mandatory ")
            Exit Function
        End If
    End If
    
    datefield11 = Sheet17.Range("yr2021.DateOfFiling11").value
    amount11 = Sheet17.Range("yr2021.HPLossCF11").value + Sheet17.Range("yr2021.BusLossOthThanSpecLossCF11").value + Sheet17.Range("yr2021.LossFrmSpecBusCF11").value + _
                Sheet17.Range("yr2021.LossFrmSpecifiedBusCF11").value + Sheet17.Range("yr2021.STCGLossCF11").value + Sheet17.Range("yr2021.LTCGLossCF11").value + Sheet17.Range("yr2021.OthSrcLossRaceHorseCF11").value
    
    If Len(datefield11) = 0 Then
        If amount11 > 0 Then
            ValidateFilingDates = False
'            msgbox16 ("* DateOfFiling(12) in Sheet : Schedule CFL  is Mandatory ")
            msgbox16 ("* Date of filing for Sl no (xii) in Sheet : Schedule CFL  is Mandatory ")
            Exit Function
        End If
    End If
    
'Change-07.06.2023.101 -> Production VF issue on mail
    Dim datefield12, amount12
    datefield12 = Sheet17.Range("yr2022.DateOfFiling12").value
    amount12 = Sheet17.Range("CFL_HP_Normal_2022").value + Sheet17.Range("yr2022.BusLossOthThanSpecLossCF12a").value + Sheet17.Range("yr2022.BusLossOthThanSpecLossCF12b").value + _
                Sheet17.Range("yr2022.BusLossOthThanSpecLossCF12").value + Sheet17.Range("yr2022.LossFrmSpecBusCF12").value + Sheet17.Range("yr2022.LossFrmSpecifiedBusCF12").value + Sheet17.Range("CFL_STCG_Normal_2022").value + _
                Sheet17.Range("CFL_LTCG_Normal_2022").value + Sheet17.Range("yr2022.OthSrcLossRaceHorseCF12").value
                
    If Len(datefield12) = 0 Then
        If amount12 > 0 Then
            ValidateFilingDates = False
'            msgbox16 ("* DateOfFiling(13) in Sheet : Schedule CFL  is Mandatory ")
            msgbox16 ("* Date of filing for Sl no (xiii) in Sheet : Schedule CFL  is Mandatory ")
'----
            Exit Function
        End If
    End If
'----


'AY_2024_25 ADD  'Malli
Dim datefield13, amount13
    datefield13 = Sheet17.Range("yr2023.DateOfFiling13").value
    amount13 = Sheet17.Range("CFL_HP_Normal_2023").value + Sheet17.Range("yr2023.BusLossOthThanSpecLossCF13a").value + Sheet17.Range("yr2023.BusLossOthThanSpecLossCF13b").value + _
                Sheet17.Range("yr2023.BusLossOthThanSpecLossCF13").value + Sheet17.Range("yr2023.LossFrmSpecBusCF13").value + Sheet17.Range("yr2023.LossFrmSpecifiedBusCF13").value + Sheet17.Range("CFL_STCG_Normal_2023").value + _
                Sheet17.Range("CFL_LTCG_Normal_2023").value + Sheet17.Range("yr2023.OthSrcLossRaceHorseCF13").value
                
    If Len(datefield13) = 0 Then
        If amount13 > 0 Then
            ValidateFilingDates = False
 
            msgbox16 ("* Date of filing for Sl no (xiv) in Sheet : Schedule CFL  is Mandatory ")
 
            Exit Function
        End If
    End If
 
'Newly added by sai on 14/02/2025 for AY 2025-26
'=================================================================================

   Dim datefield14, amount14
    datefield14 = Sheet17.Range("yr2024.DateOfFiling14").value
    amount14 = Sheet17.Range("CFL_HP_Normal_2024").value + Sheet17.Range("yr2024.BusLossOthThanSpecLossCF14a").value + Sheet17.Range("yr2024.BusLossOthThanSpecLossCF14b").value + _
                Sheet17.Range("yr2024.BusLossOthThanSpecLossCF14").value + Sheet17.Range("yr2024.LossFrmSpecBusCF14").value + Sheet17.Range("yr2024.LossFrmSpecifiedBusCF14").value + Sheet17.Range("CFL_STCG_Normal_2024").value + _
                Sheet17.Range("CFL_LTCG_Normal_2024").value + Sheet17.Range("yr2024.OthSrcLossRaceHorseCF14").value
                
    If Len(datefield14) = 0 Then
        If amount14 > 0 Then
            ValidateFilingDates = False
 
            msgbox16 ("* Date of filing for Sl no (xv) in Sheet : Schedule CFL  is Mandatory ")
 
            Exit Function
        End If
    End If
 '==================================================================================
 
   'Added by Aavula Naresh for AY 2026-27
   
  Dim datefield15, amount15
    datefield15 = Sheet17.Range("yr2025.DateOfFiling15").value
    amount15 = Sheet17.Range("CFL_HP_Normal_2025").value + Sheet17.Range("yr2025.BusLossOthThanSpecLossCF15a").value + Sheet17.Range("yr2025.BusLossOthThanSpecLossCF15b").value + _
                Sheet17.Range("yr2025.BusLossOthThanSpecLossCF15").value + Sheet17.Range("yr2025.LossFrmSpecBusCF15").value + Sheet17.Range("yr2025.LossFrmSpecifiedBusCF15").value + Sheet17.Range("CFL_STCG_Normal_2025").value + _
                Sheet17.Range("CFL_LTCG_Normal_2025").value + Sheet17.Range("yr2025.OthSrcLossRaceHorseCF15").value
                
    If Len(datefield15) = 0 Then
        If amount15 > 0 Then
            ValidateFilingDates = False
 
            msgbox16 ("* Date of filing for Sl no (xvi) in Sheet : Schedule CFL  is Mandatory ")
 
            Exit Function
        End If
    End If
'-------------------------


End Function

Function ValidateLTCGLossCurrentYr() As Boolean
 ValidateLTCGLossCurrentYr = True
    LTCGLossCurrentYr = Sheet17.Range("yr2007.LTCGLossCF10").value
    LTCGLoss1 = Sheet13.Range("IHLA.Eviii5_LtclSetOff10Per").value
    LTCGLoss2 = Sheet13.Range("IHLA.Eviii6_LtclSetOff20Per").value
    LTCGLoss3 = Sheet13.Range("IHLA.Eviii8_LtclSetOffDTAARate").value
    
    LTCGLoss4 = Sheet13.Range("IHLA.Exii8_LtclSetOff12.5Per").value 'Added by Sai(07-05-2025)
    'LTCGLossCurrentYrCalculated = LTCGLoss1 + LTCGLoss2 + LTCGLoss3
    LTCGLossCurrentYrCalculated = LTCGLoss1 + LTCGLoss2 + LTCGLoss3 + LTCGLoss4 ''Added by Sai(07-05-2025)
    
    If Len(LTCGLossCurrentYr) > 14 Then
       msgbox16 ("* Total of earlier LTCG loss should not be greater than 14 digits in Sheet CFL.")
        ValidateLTCGLossCurrentYr = False
        Exit Function
    End If
    If LTCGLossCurrentYr > LTCGLossCurrentYrCalculated Then
       msgbox16 ("* Current year LTCG loss should not exceed that Total Capital Gain Losses")
        ValidateLTCGLossCurrentYr = False
        Exit Function
    End If
End Function

Function ValidateSTCGLossCurrentYr() As Boolean
 ValidateSTCGLossCurrentYr = True
    STCGLossCurrentYr = Sheet17.Range("yr2007.STCGLossCF10").value
    STCGLoss1 = Sheet13.Range("IHLA.Eviii2_StclSetoff15Per").value
    STCGLoss2 = Sheet13.Range("IHLA.Eviii3_StclSetoff30Per").value
    STCGLoss3 = Sheet13.Range("IHLA.Eviii4_StclSetoffAppRate").value
    STCGLoss4 = Sheet13.Range("IHLA.Eviii5_StclSetoffDTAARate").value
    
    STCGLoss5 = Sheet13.Range("IHLA.Exii2_StclSetoff20Per").value    'Added by Sai(07-05-2025)
    'STCGLossCurrentYrCalculated = STCGLoss1 + STCGLoss2 + STCGLoss3 + STCGLoss4
    STCGLossCurrentYrCalculated = STCGLoss1 + STCGLoss2 + STCGLoss3 + STCGLoss4 + STCGLoss5 'Added by Sai(07-05-2025)
    If Len(STCGLossCurrentYr) > 14 Then
       msgbox16 ("* Total of earlier STCG loss should not be greater than 14 digits in Sheet CFL.")
        ValidateSTCGLossCurrentYr = False
        Exit Function
    End If
    If STCGLossCurrentYr > STCGLossCurrentYrCalculated Then
       msgbox16 ("* Current year STCG loss should not exceed that Total Capital Gain Losses")
        ValidateSTCGLossCurrentYr = False
        Exit Function
    End If
End Function


Function ValidateHPLossCurrentYr() As Boolean
 ValidateHPLossCurrentYr = True
    HPLossCurrentYr = Sheet17.Range("yr2007.HPLossCF10").value
    HPLossCurrentYrCalculated = Sheet16.Range("sheet16.BalHPlossCurYrAftSetoff").value
    If Len(HPLossCurrentYr) > 14 Then
       msgbox16 ("* Current year HP loss should not be greater than 14 digits in Sheet CFL.")
        ValidateHPLossCurrentYr = False
        Exit Function
    End If
    If HPLossCurrentYr > HPLossCurrentYrCalculated Then
       msgbox16 ("* Current year HP loss should not exceed that Total House Property Loss set off")
        ValidateHPLossCurrentYr = False
        Exit Function
    End If
End Function

Function ValidateBusLossCurrentYr() As Boolean
 ValidateBusLossCurrentYr = True
    BusLossCurrentYr = Sheet17.Range("yr2007.BusLossOthThanSpecLossCF10").value
    BusLossCurrentYrCalculated = Sheet16.Range("sheet16.BalBusLossAftSetoff").value
    If Len(BusLossCurrentYr) > 14 Then
       msgbox16 ("* Current year Bussiness loss should not be greater than 14 digits in Sheet CFL.")
        ValidateBusLossCurrentYr = False
        Exit Function
    End If
    If BusLossCurrentYr > BusLossCurrentYrCalculated Then
       msgbox16 ("* Current year Bussiness loss should not exceed that Total Bussiness Loss set off")
        ValidateBusLossCurrentYr = False
        Exit Function
    End If
End Function

Function ValidateSpecBusLossCurrentYr() As Boolean
 ValidateSpecBusLossCurrentYr = True
    SpecBusLossCurrentYr = Sheet17.Range("yr2007.LossFrmSpecBusCF10").value
    SpecBusLossCurrentYrCalculated = Abs(WorksheetFunction.Min(Sheet9.Range("sheet12.AdjustedPLFrmSpecuBus").value, 0))

    If Len(SpecBusLossCurrentYr) > 14 Then
       msgbox16 ("* Current year Speculative Bussiness loss should not be greater than 14 digits in Sheet CFL.")
        ValidateSpecBusLossCurrentYr = False
        Exit Function
    End If
    If SpecBusLossCurrentYr > SpecBusLossCurrentYrCalculated Then
       msgbox16 ("* Current year Speculative Bussiness loss should not exceed that Total Speculative Bussiness Loss set off")
        ValidateSpecBusLossCurrentYr = False
        Exit Function
    End If
End Function

Function ValidateSpecifiedBusLossCurrentYr() As Boolean
 ValidateSpecifiedBusLossCurrentYr = True
    SpecifiedBusLossCurrentYr = Sheet17.Range("yr2007.LossFrmSpecifiedBusCF10").value
    SpecifiedBusLossCurrentYrCalculated = Abs(WorksheetFunction.Min(Sheet9.Range("sheet12.AdjustedPLFrmSpecifiedBus").value, 0))

    If Len(SpecifiedBusLossCurrentYr) > 14 Then
       msgbox16 ("* Current year Specified Bussiness loss should not be greater than 14 digits in Sheet CFL.")
        ValidateSpecifiedBusLossCurrentYr = False
        Exit Function
    End If
    If SpecifiedBusLossCurrentYr > SpecifiedBusLossCurrentYrCalculated Then
       msgbox16 ("* Current year Specified Bussiness loss should not exceed that Specified Bussiness loss set off")
        ValidateSpecifiedBusLossCurrentYr = False
        Exit Function
    End If
End Function

Function ValidateOthrLossCurrentYear() As Boolean
 ValidateOthrLossCurrentYear = True
    OthrLossCurrentYr = Sheet17.Range("yr2007.OthSrcLossRaceHorseCF10").value
    OthrLossCurrentYrCalculated = Abs(WorksheetFunction.Min(Sheet14.Range("os.BalanceOwnRaceHorse").value, 0))
    If Len(OthrLossCurrentYr) > 14 Then
       msgbox16 ("* Loss from owning and maintaining race horses should not be greater than 14 digits in Sheet CFL.")
        ValidateOthrLossCurrentYear = False
        Exit Function
    End If
    If OthrLossCurrentYr > OthrLossCurrentYrCalculated Then
       msgbox16 ("* Loss from owning and maintaining race horses should not exceed Total Other Source Losses.")
        ValidateOthrLossCurrentYear = False
        Exit Function
    End If
 End Function
 
Function ValidateNonEditablefieldsCFL() As Boolean
ValidateNonEditablefieldsCFL = True

If Not LengthCheck(Sheet17.Range("yr2001.HPLossCF2").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2002.HPLossCF3").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2003.HPLossCF4").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2004.HPLossCF5").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2005.HPLossCF6").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2006.HPLossCF7").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2006.HPLossCF9").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2007.HPLossCF9").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2020.HPLossCF10").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2021.HPLossCF11").value) Then ValidateNonEditablefieldsCFL = False

'If Not LengthCheck(Sheet17.Range("yr2000.STCGLossCF1").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2001.STCGLossCF2").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2002.STCGLossCF3").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2003.STCGLossCF4").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2004.STCGLossCF5").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2005.STCGLossCF6").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2006.STCGLossCF7").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2006.STCGLossCF9").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2007.STCGLossCF9").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2020.STCGLossCF10").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2021.STCGLossCF11").value) Then ValidateNonEditablefieldsCFL = False

'If Not LengthCheck(Sheet17.Range("yr2000.LTCGLossCF1").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2001.LTCGLossCF2").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2002.LTCGLossCF3").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2003.LTCGLossCF4").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2004.LTCGLossCF5").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2005.LTCGLossCF6").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2006.LTCGLossCF7").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2006.LTCGLossCF9").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2007.LTCGLossCF9").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2020.LTCGLossCF10").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2021.LTCGLossCF11").value) Then ValidateNonEditablefieldsCFL = False

If Not LengthCheck(Sheet17.Range("totofbfloss.HPLossCF8").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("adjtotloss.HPLossCF9").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2007.HPLossCF10").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("totcfloss.HPLossCF11").value) Then ValidateNonEditablefieldsCFL = False

If Not LengthCheck(Sheet17.Range("totofbfloss.BusLossOthThanSpecLossCF8").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("adjtotloss.BusLossOthThanSpecLossCF9").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2007.BusLossOthThanSpecLossCF10").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("totcfloss.BusLossOthThanSpecLossCF11").value) Then ValidateNonEditablefieldsCFL = False

If Not LengthCheck(Sheet17.Range("totofbfloss.LossFrmSpecBusCF8").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("adjtotloss.LossFrmSpecBusCF9").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2007.LossFrmSpecBusCF10").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("totcfloss.LossFrmSpecBusCF11").value) Then ValidateNonEditablefieldsCFL = False

If Not LengthCheck(Sheet17.Range("totofbfloss.LossFrmSpecifiedBusCF8").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("adjtotloss.LossFrmSpecifiedBusCF9").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2007.LossFrmSpecifiedBusCF10").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("totcfloss.LossFrmSpecifiedBusCF11").value) Then ValidateNonEditablefieldsCFL = False

If Not LengthCheck(Sheet17.Range("totofbfloss.STCGLossCF8").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("adjtotloss.STCGLossCF9").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2007.STCGLossCF10").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("totcfloss.STCGLossCF11").value) Then ValidateNonEditablefieldsCFL = False

If Not LengthCheck(Sheet17.Range("totofbfloss.LTCGLossCF8").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("adjtotloss.LTCGLossCF9").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2007.LTCGLossCF10").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("totcfloss.LTCGLossCF11").value) Then ValidateNonEditablefieldsCFL = False

If Not LengthCheck(Sheet17.Range("totofbfloss.OthSrcLossRaceHorseCF8").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("adjtotloss.OthSrcLossRaceHorseCF9").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("yr2007.OthSrcLossRaceHorseCF10").value) Then ValidateNonEditablefieldsCFL = False
If Not LengthCheck(Sheet17.Range("totcfloss.OthSrcLossRaceHorseCF11").value) Then ValidateNonEditablefieldsCFL = False

If ValidateNonEditablefieldsCFL = False Then
    msgbox16 ("* Amount in any filed in Schedule CFL should not exceed 14 digits. Please check.")
    Exit Function
End If
 
End Function

Function LengthCheck(field As Variant) As Boolean
LengthCheck = True

If Len(field) > 14 Then
LengthCheck = False
End If

End Function

Function addCFLformula()

Sheet17.Unprotect Password:=getmsgstate

'Sheet17.Range("yr2007.HPLossCF10").formula = "=sheet16.BalHPlossCurYrAftSetoff"
'Sheet17.Range("yr2007.BusLossOthThanSpecLossCF10").formula = "=sheet16.BalBusLossAftSetoff"
'Sheet17.Range("yr2007.LossFrmSpecBusCF10").formula = "=ABS(MIN(sheet12.AdjustedPLFrmSpecuBus,0))"
'Sheet17.Range("yr2007.LossFrmSpecifiedBusCF10").formula = "=ABS(MIN(sheet12.AdjustedPLFrmSpecifiedBus,0))"
'Sheet17.Range("yr2007.STCGLossCF10").formula = "=IHLA.Eviii2_StclSetoff15Per+IHLA.Eviii3_StclSetoff30Per+IHLA.Eviii4_StclSetoffAppRate+IHLA.Eviii5_StclSetoffDTAARate"
'Sheet17.Range("yr2007.LTCGLossCF10").formula = "=IHLA.Eviii5_LtclSetOff10Per+IHLA.Eviii6_LtclSetOff20Per+IHLA.Eviii8_LtclSetOffDTAARate"
'Sheet17.Range("yr2007.OthSrcLossRaceHorseCF10").formula = "=ABS(MIN(os.BalanceOwnRaceHorse,0))"

Sheet17.Protect Password:=getmsgstate

End Function

Function ValidateDateOfFiling12_16() As Boolean
    Dim i As Long
    
    ValidateDateOfFiling12_16 = True
    DateOfFiling12_16 = Sheet17.Range("yr2022.DateOfFiling12").value
    
    If Not CheckDateddmmyyyy(DateOfFiling12_16) Then
        ValidateDateOfFiling12_16 = False
        msgbox16 ("* DateOfFiling13 in Sheet : Schedule CFL  is invalid")
        Exit Function
    Else
        DateOfFiling12_16 = Dformat(Sheet17.Range("yr2022.DateOfFiling12"), "")
    End If
    
    If Len(DateOfFiling12_16) = 10 Then
        If Not CheckDateMinDDMMYYYY(DateOfFiling12_16, 1, 4, 2022, "") Then
            ValidateDateOfFiling12_16 = False
            msgbox16 ("* Date of Filing (13) in Sch CFL  must not be less than 01/04/2022")
            Exit Function
        Else
            DateOfFiling12_16 = Dformat(Sheet17.Range("yr2022.DateOfFiling12"), "")
        End If
    End If
End Function

'Malli
Function ValidateDateOfFiling13_16() As Boolean
    Dim i As Long
    
    ValidateDateOfFiling13_16 = True
    DateOfFiling13_16 = Sheet17.Range("yr2023.DateOfFiling13").value
    
    If Not CheckDateddmmyyyy(DateOfFiling13_16) Then
        ValidateDateOfFiling13_16 = False
        msgbox16 ("* DateOfFiling (14) in Sheet : Schedule CFL  is invalid")
        Exit Function
    Else
        DateOfFiling13_16 = Dformat(Sheet17.Range("yr2023.DateOfFiling13"), "")
    End If
    
    If Len(DateOfFiling13_16) = 10 Then
        If Not CheckDateMinDDMMYYYY(DateOfFiling13_16, 1, 4, 2023, "") Then
            ValidateDateOfFiling13_16 = False
            msgbox16 ("* Date of Filing (14) in Sch CFL  must not be less than 01/04/2023")
            Exit Function
        Else
            DateOfFiling13_16 = Dformat(Sheet17.Range("yr2023.DateOfFiling13"), "")
        End If
    End If
End Function

Function ValidateDateOfFiling14_16() As Boolean
    Dim i As Long
    
    ValidateDateOfFiling14_16 = True
    DateOfFiling14_16 = Sheet17.Range("yr2024.DateOfFiling14").value
    
    If Not CheckDateddmmyyyy(DateOfFiling14_16) Then
        ValidateDateOfFiling14_16 = False
        msgbox16 ("* DateOfFiling (15) in Sheet : Schedule CFL  is invalid")
        Exit Function
    Else
        DateOfFiling14_16 = Dformat(Sheet17.Range("yr2024.DateOfFiling14"), "")
    End If
    
    If Len(DateOfFiling14_16) = 10 Then
        If Not CheckDateMinDDMMYYYY(DateOfFiling14_16, 1, 4, 2024, "") Then
            ValidateDateOfFiling14_16 = False
            msgbox16 ("* Date of Filing (15) in Sch CFL  must not be less than 01/04/2024")
            Exit Function
        Else
            DateOfFiling14_16 = Dformat(Sheet17.Range("yr2024.DateOfFiling14"), "")
        End If
    End If
End Function
Function ValidateDateOfFiling15_16() As Boolean
    Dim i As Long
    
    ValidateDateOfFiling15_16 = True
    DateOfFiling15_16 = Sheet17.Range("yr2025.DateOfFiling15").value
    
    If Not CheckDateddmmyyyy(DateOfFiling15_16) Then
        ValidateDateOfFiling15_16 = False
        msgbox16 ("* DateOfFiling (16) in Sheet : Schedule CFL  is invalid")
        Exit Function
    Else
        DateOfFiling15_16 = Dformat(Sheet17.Range("yr2025.DateOfFiling15"), "")
    End If
    
    If Len(DateOfFiling15_16) = 10 Then
        If Not CheckDateMinDDMMYYYY(DateOfFiling15_16, 1, 4, 2025, "") Then
            ValidateDateOfFiling15_16 = False
            msgbox16 ("* Date of Filing (16) in Sch CFL  must not be less than 01/04/2025")
            Exit Function
        Else
            DateOfFiling15_16 = Dformat(Sheet17.Range("yr2025.DateOfFiling15"), "")
        End If
    End If
End Function
