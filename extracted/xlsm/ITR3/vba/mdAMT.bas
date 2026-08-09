Attribute VB_Name = "mdAMT"
Option Explicit

Public TotalIncItem11_AMT As Variant
Public DeductClaimSec6A_AMT As Variant
Public DeductClaimSec10AA_AMT As Variant
Public DeductClaimSec35AD_AMT As Variant
Public Total_AMT As Variant
Public AdjustedUnderSec115JC_AMT As Variant
Public TaxPayableUnderSec115JC_AMT As Variant
Public msgbox_AMT As Variant
Sub ValidateSchAMT()

 If Not ValidateScheduleAMT Then
 Sheet22.Activate
 fmsgboxsmall msgbox_AMT ', vbOKOnly, "Error(s!)"
 CloseMsg
 End If
 
 
End Sub
Sub ValidateSchAMTclick()
 
 ValidateSchAMT
 If msgbox_AMT = "" Then fmsgboxStatus "Sheet AMT is OK"
 
End Sub


Function ValidateScheduleAMT() As Boolean
    ValidateScheduleAMT = True
    msgbox_AMT = ""
    If Not ValidateTotalIncItem11() Then ValidateScheduleAMT = False
    If Not ValidateDeductClaimSec6A() Then ValidateScheduleAMT = False
    If Not ValidateDeductClaimSec10AA() Then ValidateScheduleAMT = False
    If Not validateDeductClaimSec35AD() Then ValidateScheduleAMT = False
    If Not ValidateTotalAdjustment() Then ValidateScheduleAMT = False
    If Not ValidateAdjustedUnderSec115JC() Then ValidateScheduleAMT = False
    If Not ValidateTaxPayableUnderSec115JC() Then ValidateScheduleAMT = False
    
End Function

Function ValidateTotalIncItem11() As Boolean
    ValidateTotalIncItem11 = True
    TotalIncItem11_AMT = Sheet22.Range("AMT.TotalIncItem11").value
    Dim i As Long
    If Len(Abs(TotalIncItem11_AMT)) > 14 Then
        msgbox_AMT = msgbox_AMT & "* * Total Income as per item 11 in Sheet : Schedule AMT should be at most 14 digits " & Chr(13)
        ValidateTotalIncItem11 = False
        Exit Function
    End If

    
End Function
Function ValidateDeductClaimSec6A() As Boolean
    ValidateDeductClaimSec6A = True
    Dim Value_2a, Value_2b, Value_21, Value_22, Value_23 As Double
    DeductClaimSec6A_AMT = Sheet22.Range("AMT.DeductClaimSec6A").value
    Dim i As Long
    If Len(DeductClaimSec6A_AMT) > 14 Then
        msgbox_AMT = msgbox_AMT & "* Deduction claimed under any section in Sheet : Schedule AMT should be at most 14 digits " & Chr(13)
        ValidateDeductClaimSec6A = False
        Exit Function
    End If
    For i = 1 To Len(DeductClaimSec6A_AMT)
        If Not IsNumeric(DeductClaimSec6A_AMT) Then
            msgbox_AMT = msgbox_AMT & "* Deduction claimed under any section in Sheet : Schedule AMT must contain only digits from 0 to 9" & Chr(13)
            ValidateDeductClaimSec6A = False
            Exit Function
        End If
    Next

    Value_21 = Sheet7.Range("sheet8b.GrossTotalIncome").value
    Value_22 = Sheet7.Range("Sheet8b.IncChargeableTaxSplRates").value
    Value_23 = Sheet7.Range("Sheet8b.DeductionsUnder10Aor10AA").value
    
    If Len(DeductClaimSec6A_AMT) > 0 Then
    
    If DeductClaimSec6A_AMT > WorksheetFunction.Max((Value_21 - Value_22 - Value_23), 0) Then
        msgbox_AMT = msgbox_AMT & "* Deduction claimed under any section in Sheet : Schedule AMT should be limited to Point (10-11-12) under section Part B-TI(PARTB and Verification)" & Chr(13)
        ValidateDeductClaimSec6A = False
        Exit Function
    End If
    End If
    
End Function

Function ValidateDeductClaimSec10AA() As Boolean
    ValidateDeductClaimSec10AA = True
    Dim Value_2a, Value_2b, Value_21, Value_22, Value_23 As Double
    
    DeductClaimSec10AA_AMT = Sheet22.Range("AMT.DeductClaimSec10AA").value
    Dim i As Long
    If Len(DeductClaimSec10AA_AMT) > 14 Then
        msgbox_AMT = msgbox_AMT & "* Deduction Claimed u/s 10AA in Sheet : Schedule AMT should be at most 14 digits " & Chr(13)
        ValidateDeductClaimSec10AA = False
        Exit Function
    End If
    For i = 1 To Len(DeductClaimSec10AA_AMT)
        If Not IsNumeric(DeductClaimSec10AA_AMT) Then
            msgbox_AMT = msgbox_AMT & "* Deduction Claimed u/s 10AA in Sheet : Schedule AMT must contain only digits from 0 to 9" & Chr(13)
            ValidateDeductClaimSec10AA = False
            Exit Function
        End If
    Next
    
    Value_21 = Sheet7.Range("sheet8b.GrossTotalIncome").value
    Value_22 = Sheet7.Range("Sheet8b.IncChargeableTaxSplRates").value
    If Len(DeductClaimSec10AA_AMT) > 0 Then
    If DeductClaimSec10AA_AMT > WorksheetFunction.Max((Value_21 - Value_22), 0) Then
        msgbox_AMT = msgbox_AMT & "* Deduction Claimed u/s 10AA in Sheet : Schedule AMT   should be limited to Point (10-11) under section Part B-TI (PARTB and Verification)" & Chr(13)
        ValidateDeductClaimSec10AA = False
        Exit Function
    End If
    End If
End Function

Function validateDeductClaimSec35AD() As Boolean
    validateDeductClaimSec35AD = True
    DeductClaimSec35AD_AMT = Sheet22.Range("AMT.DeductClaimSec35AD").value
       'Added by ramya 70620
       'Added by Aavula
       
        If Sheet22.Visible = xlSheetVisible Then
           If Sheet22.Range("AMT.DeductClaimSec35AD").value = "" Or Sheet22.Range("AMT.DeductClaimSec35AD").value = "(Select)" Then
               msgbox_AMT = msgbox_AMT & "* Asset on which deduction has been claimed u/s 35AD, normal depreciation cannot be claimed on such asset."
           validateDeductClaimSec35AD = False
           End If
         End If
    
End Function

Function ValidateTotalAdjustment() As Boolean
    ValidateTotalAdjustment = True
    Total_AMT = Sheet22.Range("AMT.Total").value
    Dim i As Long
    If Len(Total_AMT) > 14 Then
        msgbox_AMT = msgbox_AMT & "* Total Adjustment in Sheet : Schedule AMT should be at most 14 digits " & Chr(13)
        ValidateTotalAdjustment = False
        Exit Function
    End If
    For i = 1 To Len(Total_AMT)
        If Not IsNumeric(Total_AMT) Then
            msgbox_AMT = msgbox_AMT & "* Total Adjustment in Sheet : Schedule AMT must contain only digits from 0 to 9" & Chr(13)
            ValidateTotalAdjustment = False
            Exit Function
        End If
    Next
End Function
Function ValidateAdjustedUnderSec115JC() As Boolean
    ValidateAdjustedUnderSec115JC = True
    AdjustedUnderSec115JC_AMT = Sheet22.Range("AMT.AdjustedUnderSec115JC").value
    Dim i As Long
    If Len(AdjustedUnderSec115JC_AMT) > 14 Then
        msgbox_AMT = msgbox_AMT & "* Adjusted Total Income in Sheet : Schedule AMT should be at most 14 digits " & Chr(13)
        ValidateAdjustedUnderSec115JC = False
        Exit Function
    End If
    For i = 1 To Len(AdjustedUnderSec115JC_AMT)
        If Not IsNumeric(AdjustedUnderSec115JC_AMT) Then
            msgbox_AMT = msgbox_AMT & "* Adjusted Total Income in Sheet : Schedule AMT must contain only digits from 0 to 9" & Chr(13)
            ValidateAdjustedUnderSec115JC = False
            Exit Function
        End If
    Next
End Function

Function ValidateTaxPayableUnderSec115JC() As Boolean
    ValidateTaxPayableUnderSec115JC = True
    TaxPayableUnderSec115JC_AMT = Sheet22.Range("AMT.TaxPayableUnderSec115JC").value
    Dim i As Long
    If Len(TaxPayableUnderSec115JC_AMT) > 14 Then
        msgbox_AMT = msgbox_AMT & "* Tax payable under section 115JC in Sheet : Schedule AMT should be at most 14 digits " & Chr(13)
        ValidateTaxPayableUnderSec115JC = False
        Exit Function
    End If
    For i = 1 To Len(TaxPayableUnderSec115JC_AMT)
    If Not IsNumeric(TaxPayableUnderSec115JC_AMT) Then
        msgbox_AMT = msgbox_AMT & "* Tax payable under section 115JC in Sheet : Schedule AMT must contain only digits from 0 to 9" & Chr(13)
        ValidateTaxPayableUnderSec115JC = False
        Exit Function
    End If
    Next
End Function
