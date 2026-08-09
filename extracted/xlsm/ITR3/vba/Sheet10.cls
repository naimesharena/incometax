Attribute VB_Name = "Sheet10"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Dim sPassword As String

Public ValEntred16_DPM15 As Variant
Public ValEntred16_DPM30 As Variant
Public ValEntred16_DPM40 As Variant
Public ValEntred16_DPM50 As Variant
Public ValEntred16_DPM60 As Variant
Public ValEntred16_DPM80 As Variant
Public ValEntred16_DPM100 As Variant

Function getRangeName(ByVal Target As Range) As String
    Dim start As Long
    start = Len(name)

    If InStr(1, Target.name.name, "'") > 0 Then
        start = start + 3 + 1
    Else
        start = start + 1 + 1
    End If

    If InStr(1, Target.name.name, ActiveSheet.name) = 0 Then
        start = 1
    End If
    getRangeName = Mid(Target.name.name, start)
End Function

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo exit1

CalculateTaxCheckFlag = False

    Dim strpassword As String
    Dim rangenm As String

    Application.EnableEvents = False
    strpassword = getmsgstate
    Sheet10.Unprotect strpassword

    If Target.Validation.Type = 3 Then
        GoTo exit1
    End If



    If (Target.name.name) = "DPM15.DepUnderSec38" Then
        If Sheet10.Range("DPM15.DepUnderSec38").value > Sheet10.Range("DPM15.TotalDepreciation").value Then
            fmsgbox "*Amount cannot be more than amount in ""Total depreciation""." & Chr(13)
                    Sheet10.Range("DPM15.DepUnderSec38").value = 0

        End If
    End If

    If (Target.name.name) = "DPM30.DepUnderSec38" Then
        If Sheet10.Range("DPM30.DepUnderSec38").value > Sheet10.Range("DPM30.TotalDepreciation").value Then
            fmsgbox "*Amount cannot be more than amount in ""Total depreciation""." & Chr(13)
                    Sheet10.Range("DPM30.DepUnderSec38").value = 0

        End If
    End If

    If (Target.name.name) = "DPM40.DepUnderSec38" Then
        If Sheet10.Range("DPM40.DepUnderSec38").value > Sheet10.Range("DPM40.TotalDepreciation").value Then
            fmsgbox "*Amount cannot be more than amount in ""Total depreciation""." & Chr(13)
                    Sheet10.Range("DPM40.DepUnderSec38").value = 0

        End If
    End If

    If (Target.name.name) = "DPM45.DepUnderSec38" Then
        If Sheet10.Range("DPM45.DepUnderSec38").value > Sheet10.Range("DPM45.TotalDepreciation").value Then
            fmsgbox "*Amount cannot be more than amount in ""Total depreciation""." & Chr(13)
                    Sheet10.Range("DPM45.DepUnderSec38").value = 0

        End If
    End If


'    Target.value = UCase(Target.value)
    rangenm = getRangeName(Target)

    If (InStr(1, rangenm, "CapGain") > 0) Then
        cg_gt_lt_0 Target.value, rangenm
    End If
exit1:
     Sheet10.Protect strpassword
     Application.EnableEvents = True
End Sub

Public Sub cg_gt_lt_0(strval As Variant, rangenam As String)
 On Error GoTo endline
    Dim CGUS50 As Variant
    Dim RealizationTotalPeriod As Variant
    Dim RealizationPeriodDuringYear As Variant
    Dim WDVFirstDay As Variant
    Dim WDVFirstDay2 As Variant
    Dim AdditionsGrThan180Days As Variant
    Dim AdditionsLessThan180Days As Variant
    Dim ExpdrOnTrforSaleAsset As Variant
    Dim NEWCGUS50 As Variant
    Dim strpassword As String


    strpassword = getmsgstate
    Sheet10.Unprotect strpassword
    Application.EnableEvents = False

    CGUS50 = 0
    NEWCGUS50 = 0
    RealizationTotalPeriod = 0
    RealizationPeriodDuringYear = 0
    WDVFirstDay = 0
    AdditionsGrThan180Days = 0
    AdditionsLessThan180Days = 0
    ExpdrOnTrforSaleAsset = 0
    WDVFirstDay2 = 0
    
    strval = IIf(strval = "", 0, strval)


    rangenam = Mid(rangenam, 1, Len(rangenam) - 12)
    Dim isDPM As Variant
    isDPM = InStr(1, rangenam, "DPM")
    If (strval = 0) Then
        Sheet10.Range(rangenam & ".DepreciationAtFullRate").formula = "=ROUND(" & rangenam & ".FullRateDeprAmt*" & rangenam & ".RATE/100,0)"
        Sheet10.Range(rangenam & ".DepreciationAtHalfRate").formula = "=ROUND(" & rangenam & ".HalfRateDeprAmt*" & rangenam & ".RATE/200,0)"
        If isDPM > 0 Then
            Sheet10.Range(rangenam & ".TotalDepreciation").formula = "=" & rangenam & ".DepreciationAtHalfRate + " & rangenam & ".DepreciationAtFullRate + " & rangenam & ".AddlnDeprDuringYearAdditions + " & rangenam & ".AddlnDeprOnGT180DayAdditions  + " & rangenam & ".AddlnDeprlessthan180days"
        Else
            Sheet10.Range(rangenam & ".TotalDepreciation").formula = "=" & rangenam & ".DepreciationAtHalfRate + " & rangenam & ".DepreciationAtFullRate + " & rangenam & ".AddlnDeprDuringYearAdditions + " & rangenam & ".AddlnDeprOnGT180DayAdditions"
        End If
        
      'Newly changed by sai on 24/04/2024
        'Sheet10.Range(rangenam & ".WDVLastDay").formula = "=MAX(" & rangenam & ".WDVFirstDay+" & rangenam & ".AdditionsGrThan180Days -" & rangenam & ".RealizationTotalPeriod + " & rangenam & ".AdditionsLessThan180Days - " & rangenam & ".RealizationPeriodDuringYear - " & rangenam & ".TotalDepreciation,0)"
    If isDPM > 0 Then
    
       Sheet10.Range(rangenam & ".WDVLastDay").formula = "=MAX(" & rangenam & ".WDVFirstDay+" & rangenam & ".Adjmt2NDprov3section115BAC+" & rangenam & ".AdditionsGrThan180Days -" & rangenam & ".RealizationTotalPeriod + " & rangenam & ".AdditionsLessThan180Days - " & rangenam & ".RealizationPeriodDuringYear - " & rangenam & ".TotalDepreciation,0)"
    Else
        Sheet10.Range(rangenam & ".WDVLastDay").formula = "=MAX(" & rangenam & ".WDVFirstDay+" & rangenam & ".AdditionsGrThan180Days -" & rangenam & ".RealizationTotalPeriod + " & rangenam & ".AdditionsLessThan180Days - " & rangenam & ".RealizationPeriodDuringYear - " & rangenam & ".TotalDepreciation,0)"
    End If
    
    
    
    'Sheet10.Range(rangenam & ".WDVLastDay").formula = "=MAX(" & rangenam & ".WDVFirstDay+" & rangenam & ".Adjmt2NDprov3section115BAC+" & rangenam & ".AdditionsGrThan180Days -" & rangenam & ".RealizationTotalPeriod + " & rangenam & ".AdditionsLessThan180Days - " & rangenam & ".RealizationPeriodDuringYear - " & rangenam & ".TotalDepreciation,0)"
    'DPM15.Adjmt2NDprov3section115BAC
    
    End If

        If (strval <> 0) Then
        If isDPM > 0 Then
            Sheet10.Range(rangenam & ".AddlnDeprlessthan180days").value = 0
        End If
        Sheet10.Range(rangenam & ".AddlnDeprDuringYearAdditions").value = 0
        Sheet10.Range(rangenam & ".AddlnDeprOnGT180DayAdditions").value = 0
        Sheet10.Range(rangenam & ".DepreciationAtHalfRate").value = 0
        Sheet10.Range(rangenam & ".DepreciationAtFullRate").value = 0
        Sheet10.Range(rangenam & ".TotalDepreciation").value = 0
        Sheet10.Range(rangenam & ".WDVLastDay").value = 0
        Sheet10.Range(rangenam & ".DepUnderSec38").value = 0
        Sheet10.Range(rangenam & ".ProAgrdep").value = 0
        CGUS50 = strval
        RealizationTotalPeriod = Sheet10.Range(rangenam & ".RealizationTotalPeriod").value
        RealizationPeriodDuringYear = Sheet10.Range(rangenam & ".RealizationPeriodDuringYear").value
        WDVFirstDay = Sheet10.Range(rangenam & ".WDVFirstDay").value
        AdditionsGrThan180Days = Sheet10.Range(rangenam & ".AdditionsGrThan180Days").value
        AdditionsLessThan180Days = Sheet10.Range(rangenam & ".AdditionsLessThan180Days").value
        ExpdrOnTrforSaleAsset = Sheet10.Range(rangenam & ".ExpdrOnTrforSaleAsset").value
        
        If isDPM > 0 Then
        WDVFirstDay2 = Sheet10.Range(rangenam & ".WDVFirstDay2").value
        'NEWCGUS50 = RealizationTotalPeriod + RealizationPeriodDuringYear - WDVFirstDay - AdditionsGrThan180Days - AdditionsLessThan180Days - ExpdrOnTrforSaleAsset - WDVFirstDay2
        'NEWCGUS50 = RealizationPeriodDuringYear - WDVFirstDay - AdditionsGrThan180Days - AdditionsLessThan180Days - ExpdrOnTrforSaleAsset - WDVFirstDay2
        NEWCGUS50 = RealizationTotalPeriod + RealizationPeriodDuringYear - WDVFirstDay - AdditionsGrThan180Days - AdditionsLessThan180Days - ExpdrOnTrforSaleAsset - WDVFirstDay2
        Else
        'NEWCGUS50 = RealizationTotalPeriod + RealizationPeriodDuringYear - WDVFirstDay - AdditionsGrThan180Days - AdditionsLessThan180Days - ExpdrOnTrforSaleAsset
        'NEWCGUS50 = RealizationPeriodDuringYear - WDVFirstDay - AdditionsGrThan180Days - AdditionsLessThan180Days - ExpdrOnTrforSaleAsset
        NEWCGUS50 = RealizationTotalPeriod + RealizationPeriodDuringYear - WDVFirstDay - AdditionsGrThan180Days - AdditionsLessThan180Days - ExpdrOnTrforSaleAsset
        End If
        
        Dim srTemp
        If Mid(rangenam, 1, 3) = "DAO" Then
            srTemp = 17
        Else
            srTemp = 20
        End If
        If (strval > 0) Then
        
            If (NEWCGUS50 > CGUS50) Or CGUS50 < 0 Or (CGUS50 > NEWCGUS50) Then
            If (NEWCGUS50 > CGUS50) Then
                fmsgboxStatus "The calculated value of Capital Gain / Loss u/s 50 in S.No " & srTemp & " is " & NEWCGUS50 & " and is different from what has been entered by you (  " & CGUS50 & " ). Please enter correct value of sale consideration in S.No 5 or 8 of this schedule. Else the entered value will be replaced with the calculated value."
                Sheet10.Range(rangenam & ".CapGainUs50").value = NEWCGUS50
            End If
            End If

        Else
        
            If (NEWCGUS50 < CGUS50) Or CGUS50 < 0 Or (CGUS50 < NEWCGUS50) Then
            If (NEWCGUS50 > CGUS50) Then
                fmsgboxStatus "The calculated value of Capital Gain/Loss u/s 50 in Sl.No " & srTemp & " is " & NEWCGUS50 & " and is different from what has been entered by you ( " & CGUS50 & " ). Please enter correct value of sale consideration in Sl.No 5 or 8 of this schedule. Else the entered value will be replaced with the calculated value"
                Sheet10.Range(rangenam & ".CapGainUs50").value = NEWCGUS50
            End If
            End If
        End If
    End If
endline:
    Sheet10.Protect strpassword
    Application.EnableEvents = True
    
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error Resume Next
If Not Range(ActiveCell.Address).Locked = True Then
       If ActiveCell.Validation.Type = 6 Then
    Application.OnKey "%{DOWN}", ""
    'Sheet1.Range(targetaddress).Value = ""
ElseIf ActiveCell.Validation.Type = 1 Then
 Application.OnKey "%{DOWN}", ""
ElseIf ActiveCell.Validation.Type = 3 Then

    Application.OnKey "%{DOWN}"
End If
Else
Application.OnKey "%{DOWN}", ""
End If

End Sub
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)

    Cancel = True


End Sub


