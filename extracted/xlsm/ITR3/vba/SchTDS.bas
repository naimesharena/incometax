Attribute VB_Name = "SchTDS"
Option Explicit
Public msgbox_TDS1, msgbox_TDS2, msgbox_TDS3, msgValidateSheetTDS3, msgbox_TCS As String
 Dim i As Long

'----TDS1----
Public TAN_TDS1 As Variant
Public EmployerOrDeductorOrCollecterName_TDS1 As Variant
Public IncChrgSal_TDS1 As Variant
Public TotalTDSSal_TDS1 As Variant

Public end_TDS1, end_TDS12, end_TDS13, end_TDS14 As Variant


'----TDS2----
Public TDSCreditName_TDS2 As Variant
Public TDSPanOfOther_TDS2 As Variant
Public TDSAadharOfOther_TDS2 As Variant

Public TAN_TDS2 As Variant
Public EmployerOrDeductorOrCollecterName_TDS2 As Variant
Public UTN_TDS2 As Variant
Public DeductedYear_TDS2 As Variant
Public BroughtFwdTDSAmt_TDS2 As Variant
Public DeductedInOwnHands_TDS2 As Variant
Public DeductedIncome5A_TDS2 As Variant
Public DeductedTDS5A_TDS2 As Variant
Public DeductedPAN5A_TDS2 As Variant
Public TDSClaimedInOwnHands_TDS2 As Variant
Public GrossAmount_TDS2 As Variant
Public HeadIncome_TDS2 As Variant
Public ClaimedIncome5A_TDS2 As Variant
Public ClaimedTDS5A_TDS2 As Variant
Public ClaimedPAN5A_TDS2 As Variant
Public ClaimedAadhar5A_TDS2 As Variant
Public AmtCarriedForward_TDS2 As Variant

Public SectionTDSDeducted_TDS2 As Variant


Public rngname_TDS2 As Variant
Public end_TDS2, end_TDS2_GrossAmount, end_TDS2_Head, end_TDS2_Pan, end_TDS22, end_TDS23, end_TDS24, end_TDS25, end_TDS26, end_TDS27, end_TDS28, end_TDS29, end_TDS201, end_TDS202, end_TDS203, end_TDS204, end_TDS205, end_TDS206, end_TDS207, end_TDS208, end_TDS209, end_TDS210, end_TDS211, end_TDSTANentry As Variant


'-----TDS3-----
Public TDSCreditName_TDS3 As Variant
Public TDSPanOfOther_TDS3 As Variant
Public Aadhar_TDS3 As Variant
Public PAN_TDS3 As Variant
Public TDS3_Aadhaar As Variant
Public Aadhaar_TDS3 As Variant
Public EmployerOrDeductorOrCollecterName_TDS3 As Variant
Public UTN_TDS3 As Variant
Public DeductedYear_TDS3 As Variant
Public BroughtFwdTDSAmt_TDS3 As Variant
Public DeductedInOwnHands_TDS3 As Variant
Public DeductedIncome5A_TDS3 As Variant
Public DeductedTDS5A_TDS3 As Variant
Public DeductedPAN5A_TDS3 As Variant
Public TDSClaimedInOwnHands_TDS3 As Variant
Public ClaimedIncome5A_TDS3 As Variant
Public ClaimedTDS5A_TDS3 As Variant
Public ClaimedPAN5A_TDS3 As Variant
Public ClaimedAadhar5A_TDS3 As Variant
Public AmtCarriedForward_TDS3 As Variant
Public GrossAmount_TDS3 As Variant
Public HeadIncome_TDS3 As Variant
Public rngname_TDS3 As Variant

Public SectionTDSisDeduct_TDS3 As Variant


Public end_TDS3, end_TDS3_GrossAmount, end_TDS3_Head, end_TDS3_Pan, end_TDS32, end_TDS33, end_TDS34, end_TDS35, end_TDS36, end_TDS37, end_TDS38, end_TDS39, end_TDS301, end_TDS302, end_TDS303, end_TDS304, end_TDS305, end_TDS306, end_TDS307, end_TDS308, end_TDS309, end_TDS310, end_TDS311, end_TDSPANentry As Variant


'----TCS----
Public TAN_TCS As Variant
Public EmployerOrDeductorOrCollecterName_TCS As Variant
Public DeductedYear_TCS As Variant
Public BroughtFwdTDSAmt_TCS As Variant 'added new
Public TotalTCS_TCS As Variant
Public AmtTCSClaimedThisYear_TCS As Variant
Public AmtCarriedFwd_TCS As Variant
Public end_TCS, end_TCS2, end_TCS3, end_TCS4, end_TCS5, end_TCS6 As Variant
Sub Cmd_AddRows_TDS1_Click()
Dim vRows As Long
EfilingCommon.DefinedgridNameRange = ("TDS1.TAN||TDS1.EmployerOrDeductorOrCollecterName||TDS1.IncChrgSal||TDS1.TotalTDSSal")
ActiveCellRange = EfilingCommon.searchLastRow("TDS1.TAN")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub Cmd_AddRows_TDS2_Click()
Dim vRows As Long
'EfilingCommon.DefinedgridNameRange = ("TDS2.TDSCreditName||TDS2.GrossAmount||TDS2.HeadOfIncome||TDS2.PanOtherPerson||TDSAadharOfOther_TDS2||TDS2.TAN||TDS2.DeductedYear||TDS2.BroughtFwdTDSAmt||TDS2.DeductedInOwnHands||TDS2.DeductedIncome5A||TDS2.DeductedTDS5A||TDS2.ClaimedInOwnHands||TDS2.ClaimedIncome5A||TDS2.ClaimedTDS5A||TDS2.ClaimedPAN5A||ClaimedAadhar5A_TDS2||TDS2.AmtCarriedForward")

EfilingCommon.DefinedgridNameRange = ("TDS2.TDSCreditName||TDS2.GrossAmount||TDS2.HeadOfIncome||TDS2.PanOtherPerson||TDSAadharOfOther_TDS2||TDS2.TAN||TDS2.TdsDeducted||TDS2.DeductedYear||TDS2.BroughtFwdTDSAmt||TDS2.DeductedInOwnHands||TDS2.DeductedIncome5A||TDS2.DeductedTDS5A||TDS2.ClaimedInOwnHands||TDS2.ClaimedIncome5A||TDS2.ClaimedTDS5A||TDS2.ClaimedPAN5A||ClaimedAadhar5A_TDS2||TDS2.AmtCarriedForward")

ActiveCellRange = EfilingCommon.searchLastRow("TDS2.TAN")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

'Change-09.01.2023.102.TCS
'Sub Cmd_AddRows_TCS_Click()
'Dim vRows As Long
'EfilingCommon.DefinedgridNameRange = ("TCS.TAN||TCS.EmployerOrDeductorOrCollecterName||TCS.DeductedYear||TCS.BroughtFwdTDSAmt||TCS.TotalTCS||TCS.AmtTCSClaimedThisYear||TCS.AmtCarriedFwd")
'ActiveCellRange = EfilingCommon.searchLastRow("TCS.TAN")
'vRows = EfilingCommon.insertRowUnderSectionWithFormula
'End Sub

'end---

Sub Cmd_AddRows_TDS3_Click()
Dim vRows As Long
'EfilingCommon.DefinedgridNameRange = ("TDS3.TDSCreditName||TDS3.GrossAmount||TDS3.HeadOfIncome||TDS3.PanOtherPerson||Aadhar_TDS3||TDS3.PAN||TDS3.Aadhaar||TDS3.DeductedYear||TDS3.BroughtFwdTDSAmt||TDS3.DeductedInOwnHands||TDS3.DeductedIncome5A||TDS3.DeductedTDS5A||TDS3.ClaimedInOwnHands||TDS3.ClaimedIncome5A||TDS3.ClaimedTDS5A||TDS3.ClaimedPAN5A||ClaimedAadhar5A_TDS3||TDS3.AmtCarriedForward")
EfilingCommon.DefinedgridNameRange = ("TDS3.TDSCreditName||TDS3.GrossAmount||TDS3.HeadOfIncome||TDS3.PanOtherPerson||Aadhar_TDS3||TDS3.PAN||TDS3.Aadhaar||TDS3.TdsIsDeducted||TDS3.DeductedYear||TDS3.BroughtFwdTDSAmt||TDS3.DeductedInOwnHands||TDS3.DeductedIncome5A||TDS3.DeductedTDS5A||TDS3.ClaimedInOwnHands||TDS3.ClaimedIncome5A||TDS3.ClaimedTDS5A||TDS3.ClaimedPAN5A||ClaimedAadhar5A_TDS3||TDS3.AmtCarriedForward")

ActiveCellRange = EfilingCommon.searchLastRow("TDS3.PAN")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub Cmd_Validate_TDS_Click()
ValidateSheetSchTDS
fmsgboxStatus "Sheet TDS is ok"
End Sub



Sub ValidateSheetSchTDS()
    subProcCaption = "Validating TDS 1"
    If Not ValidatesheetTDS1 Then
        Sheet28.Activate
        fmsgboxsmall msgbox_TDS1 ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    subProcCaption = "Validating TDS 2"
    If Not ValidatesheetTDS2 Then
        Sheet28.Activate
        fmsgboxsmall msgbox_TDS2  ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
    subProcCaption = "Validating TCS"
'    Change-09.01.2023.102.TCS2
'    If Not ValidatesheetTCS Then
'        Sheet28.Activate
'        fmsgboxsmall msgbox_TCS ', vbOKOnly, "Error(s)!"
'        CloseMsg
'    End If
    If Not ValidateTCS1 Then
        Sheet28.Activate
        fmsgboxsmall msgValidateSheetTCS ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
'    end---
    subProcCaption = "Validating TDS 3"
    If Not ValidatesheetTDS3 Then
        Sheet28.Activate
        fmsgboxsmall msgbox_TDS3 ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
End Sub
Function ValidatesheetTDS1() As Boolean
     ValidatesheetTDS1 = True
     msgbox_TDS1 = ""
     If Not ValidateTAN_TDS1() Then ValidatesheetTDS1 = False
    If (Len(Sheet28.Range("TDS1.TAN").item(1).value) > 0) Then
     If Not ValidateEmployerOrDeductorOrCollecterName_TDS1() Then ValidatesheetTDS1 = False
     If Not ValidateIncChrgSal_TDS1() Then ValidatesheetTDS1 = False
     If Not ValidateTotalTDSSal_TDS1() Then ValidatesheetTDS1 = False
'Change-25.07.2023.101 -> SIT-48404
    If Not ValidateTAN_TDS1WithSalary() Then ValidatesheetTDS1 = False
'---
  End If

setTblinfo_TDS12
setTblinfo_TDS13
setTblinfo_TDS14

'Added by Aavula Naresh
'    If Sheet24.Range("grosssalandexemptincome_hidden").value < Sheet28.Range("Incm_TDS1").value Then
''        msgbox_TDS1 = msgbox_TDS1 + "*  ""Amount of ""Gross Salary"" and ""exempt incomes"" which are part of salary declared is less than the amount of Gross Salary declared in Schedule TDS 1""" & Chr(13)
''        ValidatesheetTDS1 = False
'     'IPIP-74303_M   Added by Malli on 22/09/2025
'
'        MsgBox "*  ""Amount of ""Gross Salary"" and ""exempt incomes"" which are part of salary declared is less than the amount of Gross Salary declared in Schedule TDS 1""", , "Alert:"
'    End If
'''''''''



If (end_TDS1 <> end_TDS12) Or (end_TDS1 <> end_TDS13) Or (end_TDS1 <> end_TDS14) Then
         msgbox_TDS1 = msgbox_TDS1 + "* Enter all Mandatory Details in Schedule TDS1." & Chr(13)
         ValidatesheetTDS1 = False
End If

End Function

'----TDS1----

Function ValidateTAN_TDS1() As Boolean
    ValidateTAN_TDS1 = True
    setTblinfo_TDS1
    Dim rangecells As Range
   
    Set rangecells = Sheet28.Range("TDS1.TAN").Cells
    ReDim TAN_TDS1(end_TDS1)
    noOfProcessSub = end_TDS1
    
For i = 1 To Sheet28.Range("TDS1.TAN").count
    If Sheet28.Range("TDS1.EmployerOrDeductorOrCollecterName").Cells(i, 1).value <> "" Or Sheet28.Range("TDS1.IncChrgSal").Cells(i, 1).value <> "" Or Sheet28.Range("TDS1.TotalTDSSal").Cells(i, 1).value <> "" Then
        If Sheet28.Range("TDS1.TAN").Cells(i, 1).value = "" Or Sheet28.Range("TDS1.TAN").Cells(i, 1).value = Empty Then
            msgbox_TDS1 = msgbox_TDS1 + "* Please enter the TAN at Sr. No  " & i & "in Schedule TDS1" & Chr(13)
            ValidateTAN_TDS1 = False
            Exit Function
        End If
    End If
Next
    
    For i = 1 To end_TDS1
        TAN_TDS1(i) = UCase(rangecells.item(i).value)
        
 If Not Len(TAN_TDS1(i)) = 0 Then
 
     If Not ValidateTantype_text(Mid(TAN_TDS1(i), 1, 4)) Then
        'SIT-95399 Jyoti 8
         'msgbox_TDS1 = msgbox_TDS1 + "* Please enter the TAN at Sr. No  " & i & " is invalid. First 4 alphabets, next 5 digits, then alphabet in Schedule TDS1" & Chr(13)
         msgbox_TDS1 = msgbox_TDS1 + "* If Invalid TAN  is entered If Invalid TAN  is entered ""Invalid TAN. TAN format should be First 4 alphabets, then 5 digits, then alphabetat Sr. No  " & i & " in schedule TDS1"" & Chr(13)"
          
         ValidateTAN_TDS1 = False
         Exit Function
     End If
     
     If Not IsNumeric(Mid(TAN_TDS1(i), 5, 5)) Then
     'SIT-95399
         'msgbox_TDS1 = msgbox_TDS1 + "* TAN at Sr. No  " & i & " is invalid. First 4 alphabets, next 5 digits, then alphabet.in Schedule TDS1" & Chr(13)
          msgbox_TDS1 = msgbox_TDS1 + "* If Invalid TAN  is entered If Invalid TAN  is entered ""Invalid TAN. TAN format should be First 4 alphabets, then 5 digits, then alphabetat Sr. No  " & i & " in schedule TDS1"" & Chr(13)"
        
         ValidateTAN_TDS1 = False
         Exit Function
     End If
     
     If Not ValidateTantype_text(Right(TAN_TDS1(i), 1)) Then
        
        'msgbox_TDS1 = msgbox_TDS1 + "* TAN at Sr. No  " & i & " is invalid. First 4 alphabets, next 5 digits, then alphabet.in Schedule TDS1" & Chr(13)
        msgbox_TDS1 = msgbox_TDS1 + "* If Invalid TAN  is entered If Invalid TAN  is entered ""Invalid TAN. TAN format should be First 4 alphabets, then 5 digits, then alphabetat Sr. No  " & i & " in schedule TDS1"" & Chr(13)"
        
         ValidateTAN_TDS1 = False
         Exit Function
     End If
     
      If Not ValidateTANCodes(UCase(TAN_TDS1(i))) Then
                msgbox_TDS1 = msgbox_TDS1 + "* Invalid TAN at Sr.No " & i & "  Schedule TDS1.Please enter valid TAN" & Chr(13)
                ValidateTAN_TDS1 = False
                Exit Function
        End If
    
 ElseIf Not chkCompulsory(TAN_TDS1(i)) Then
         msgbox_TDS1 = msgbox_TDS1 + "* Please enter the TAN at Sr. No  " & i & "in Schedule TDS1" & Chr(13)
     ValidateTAN_TDS1 = False
     Exit Function
 End If
 
 UpdateProgressBar
 Next
 
End Function

Function ValidateEmployerOrDeductorOrCollecterName_TDS1() As Boolean
 
    ValidateEmployerOrDeductorOrCollecterName_TDS1 = True
    setTblinfo_TDS1
    Dim rangecells As Range
    Set rangecells = Sheet28.Range("TDS1.EmployerOrDeductorOrCollecterName").Cells
    ReDim EmployerOrDeductorOrCollecterName_TDS1(end_TDS1)
    For i = 1 To end_TDS1
        EmployerOrDeductorOrCollecterName_TDS1(i) = rangecells.item(i).value
        
     If Not chkCompulsory(EmployerOrDeductorOrCollecterName_TDS1(i)) Then
'Change-14.05.2023.101 -> SIT-49877
'         msgbox_TDS1 = msgbox_TDS1 + "* Please enter the EmployerOrDeductorOrCollecterName at Sr. No  " & i & " in Schedule TDS1" & Chr(13)
         msgbox_TDS1 = msgbox_TDS1 + "* Please enter name of employer at Sr. No  " & i & " in Schedule TDS1" & Chr(13)
         ValidateEmployerOrDeductorOrCollecterName_TDS1 = False
         Exit Function
     End If
     
 Next
End Function

Function ValidateIncChrgSal_TDS1() As Boolean
    ValidateIncChrgSal_TDS1 = True
    setTblinfo_TDS1
    Dim rangecells As Range
    Set rangecells = Sheet28.Range("TDS1.IncChrgSal").Cells
    ReDim IncChrgSal_TDS1(end_TDS1)
    For i = 1 To end_TDS1
        IncChrgSal_TDS1(i) = rangecells.item(i).value
        
        If Not chkCompulsory(IncChrgSal_TDS1(i)) Then
             msgbox_TDS1 = msgbox_TDS1 + "* Please enter the Income chargable from salary at Sr. No " & i & " in Schedule TDS1" & Chr(13)
            ValidateIncChrgSal_TDS1 = False
            Exit Function
        End If
 Next
End Function

Function ValidateTotalTDSSal_TDS1() As Boolean
    ValidateTotalTDSSal_TDS1 = True
    setTblinfo_TDS1
    Dim rangecells As Range
    Set rangecells = Sheet28.Range("TDS1.TotalTDSSal").Cells
    ReDim TotalTDSSal_TDS1(end_TDS1)
    For i = 1 To end_TDS1
        TotalTDSSal_TDS1(i) = rangecells.item(i).value
        
        If Not chkCompulsory(TotalTDSSal_TDS1(i)) Then
             msgbox_TDS1 = msgbox_TDS1 + "* Please entr the Total tax deducted at Sr. No " & i & "in Schedule TDS1" & Chr(13)
            ValidateTotalTDSSal_TDS1 = False
            Exit Function
        End If
        
        If TotalTDSSal_TDS1(i) > IncChrgSal_TDS1(i) Then
        'SIT-95399
              ' msgbox_TDS1 = msgbox_TDS1 + "* Please enter the Total tax (5) cannot be greater than Income chargable(4) at Sr. No  " & i & "in Schedule TDS1" & Chr(13)
               msgbox_TDS1 = msgbox_TDS1 + "*""Total tax deducted cannot be more than Income chargeable at Sr. No  " & i & " under Salaries in schedule TDS1""" & Chr(13)
                 
                ValidateTotalTDSSal_TDS1 = False
                Exit Function
        End If
    Next
    
    Dim Total_TDS1 As Variant
    Total_TDS1 = 0
    Total_TDS1 = Sheet28.Range("TDS1.Sum").value
        
    If Len(Total_TDS1) > 14 Then
        msgbox_TDS1 = msgbox_TDS1 + "* Please enter the Total of TDS1 cannot be greater than 14 digits in Schedule TDS1" & Chr(13)
        ValidateTotalTDSSal_TDS1 = False
        Exit Function
    End If
    
End Function

Sub setTblinfo_TDS1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet28.Range("TDS1.TAN").count
    Set rangecells = Sheet28.Range("TDS1.TAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDS1 = ccount
    
End Sub

Sub setTblinfo_TDS12()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet28.Range("TDS1.EmployerOrDeductorOrCollecterName").count
    Set rangecells = Sheet28.Range("TDS1.EmployerOrDeductorOrCollecterName").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDS12 = ccount
End Sub

Sub setTblinfo_TDS13()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet28.Range("TDS1.IncChrgSal").count
    Set rangecells = Sheet28.Range("TDS1.IncChrgSal").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDS13 = ccount
End Sub


Sub setTblinfo_TDS14()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet28.Range("TDS1.TotalTDSSal").count
    Set rangecells = Sheet28.Range("TDS1.TotalTDSSal").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDS14 = ccount
End Sub
'end TDS1

'----TDS2----
Function ValidatesheetTDS2() As Boolean
    ValidatesheetTDS2 = True
    msgbox_TDS2 = "Schedule TDS2 :" & Chr(10)

    If Len(Sheet28.Range("TDS2.Total").value) > 14 Then
       msgbox_TDS2 = msgbox_TDS2 & " Total of TDS2 should not excced 14 digits " & Chr(13)
       ValidatesheetTDS2 = False
       Exit Function
    End If
    
    setTblinfo_TDS2
    setTblinfo_TDS2_Pan
    setTblinfo_TDS2_GrossAmount
    setTblinfo_TDS2_Head
    setTblinfo_TDS22
    setTblinfo_TDS23
    'setTblinfo_TDS24
    setTblinfo_TDS25
    setTblinfo_TDS26
    setTblinfo_TDS27
    setTblinfo_TDS28
    setTblinfo_TDS29
   ' setTblinfo_TDS201
    setTblinfo_TDS202
    setTblinfo_TDS203
    setTblinfo_TDS204
    setTblinfo_TDS205
    setTblinfo_TDS206

    For i = 1 To end_TDS2
        If isdropdownblank(Sheet28.Range("TDS2.TDSCreditName").item(i).value) Then
            msgbox_TDS2 = msgbox_TDS2 + "* Please fill all the mandatory fields of the row " & i & " before filling the next row.in Schedule TDS2" & Chr(13)
            ValidatesheetTDS2 = False
            Exit Function
        End If
    Next

    If Not ValidateTDSCreditName_TDS2() Then ValidatesheetTDS2 = False
     If Not ValidateTAN_TDS2() Then ValidatesheetTDS2 = False
    If (Len(Sheet28.Range("TDS2.TDSCreditName").item(1).value) > 0) Then
       
        'If Not ValidateEmployerOrDeductorOrCollecterName_TDS2() Then ValidatesheetTDS2 = False
        'If Not ValidateUTN_TDS2() Then ValidatesheetTDS2 = False
        
        If Not ValidateSectionTDSDeducted_TDS2() Then ValidatesheetTDS2 = False 'Added by Aavula Naresh for AY 2025-26

        If Not ValidateDeductedYear_TDS2() Then ValidatesheetTDS2 = False
        If Not ValidateBroughtFwdTDSAmt_TDS2() Then ValidatesheetTDS2 = False
        If Not ValidateDeductedInOwnHands_TDS2() Then ValidatesheetTDS2 = False
        If Not ValidateDeductedIncome5A_TDS2() Then ValidatesheetTDS2 = False
        If Not ValidateDeductedTDS5A_TDS2() Then ValidatesheetTDS2 = False
        'If Not ValidateDeductedPAN5A_TDS2() Then ValidatesheetTDS2 = False
        If Not ValidateTDSClaimedInOwnHands_TDS2 Then ValidatesheetTDS2 = False
        If Not ValidateClaimedIncome5A_TDS2 Then ValidatesheetTDS2 = False
        If Not ValidateClaimedTDS5A_TDS2 Then ValidatesheetTDS2 = False
        If Not ValidateClaimedPAN5A_TDS2 Then ValidatesheetTDS2 = False
        If Not ValidateAmtCarriedForward_TDS2 Then ValidatesheetTDS2 = False
        If Not ValidateCalculations_TDS2 Then ValidatesheetTDS2 = False

    End If
    
    For i = 1 To Sheet28.Range("TDS2.TAN").count

        If Len(Sheet28.Range("TDS2.BroughtFwdTDSAmt").item(i).value) > 0 Then
    
            If Len(Sheet28.Range("TDS2.TAN").item(i).value) = 0 Then
                msgbox_TDS2 = msgbox_TDS2 + "* Enter All Mandatory Fields at Sr. No " & i & " in Schedule TDS2" & Chr(13)
                ValidatesheetTDS2 = False
                Exit Function
            End If
            
            If isdropdownblank(Sheet28.Range("TDS2.DeductedYear").item(i).value) Then
'                msgbox_TDS2 = msgbox_TDS2 + "* Please select the financial year in which TDS was deducted at Sr. No " & i & " in Schedule TDS2" & Chr(13)
            'For SIT-69645
                msgbox_TDS2 = msgbox_TDS2 + "* Please select Year of tax deduction from drop-down in schedule TDS2 at Sr. No " & i & " " & Chr(13)
                ValidatesheetTDS2 = False
                Exit Function
            End If
        End If
        
        If Not isdropdownblank(Sheet28.Range("TDS2.DeductedYear").item(i).value) Then
            If Len(Sheet28.Range("TDS2.BroughtFwdTDSAmt").item(i).value) = 0 Then
'                msgbox_TDS2 = msgbox_TDS2 + "* Please enter the the TDS b/f at Sr. No " & i & " in Schedule TDS2" & Chr(13)

                 msgbox_TDS2 = msgbox_TDS2 + "* ""Please enter the Amount of TDS brought forwarded in schedule TDS2"" at Sr.No " & i & " " & Chr(13)
                
                ValidatesheetTDS2 = False
                Exit Function
            End If
        End If
        
        If Len(Sheet28.Range("TDS2.BroughtFwdTDSAmt").item(i).value) <> 0 Then
            If Len(Sheet28.Range("TDS2.DeductedIncome5A").item(i).value) <> 0 Then
            
                msgbox_TDS2 = msgbox_TDS2 + "* Please enter the Unclaimed TDS brought forward and Income details of TDS Deducted in the hands of any other person as per rule 37BA(2) should be provided in different rows in Sch TDS 2 at Sr. No " & i & "in Schedule TDS2" & Chr(13)
                ValidatesheetTDS2 = False
                Exit Function
            End If
        End If
        
        If Len(Sheet28.Range("TDS2.DeductedIncome5A").item(i).value) <> 0 Then
            If Len(Sheet28.Range("TDS2.BroughtFwdTDSAmt").item(i).value) <> 0 Then
            
                msgbox_TDS2 = msgbox_TDS2 + "* Please enter the Unclaimed TDS brought forward and Income details of TDS Deducted in the hands of any other person as per rule 37BA(2) should be provided in different rows in Sch TDS 2 at Sr. No " & i & " in Schedule TDS2" & Chr(13)
                ValidatesheetTDS2 = False
                Exit Function
            End If
        End If
        
    Next
    
    For i = 1 To Sheet28.Range("TDS2.TAN").count
    
        If Len(Sheet28.Range("TDS2.DeductedIncome5A").item(i).value) > 0 Then
    
            If isdropdownblank(Sheet28.Range("TDS2.TDSCreditName").item(i).value) Then
                msgbox_TDS2 = msgbox_TDS2 + "* Enter All Mandatory Fields at Sr. No " & i & " in Schedule TDS2" & Chr(13)
                ValidatesheetTDS2 = False
                Exit Function
            End If
            
            If Len(Sheet28.Range("TDS2.DeductedTDS5A").item(i).value) = 0 Then
            
                msgbox_TDS2 = msgbox_TDS2 + "* TDS is mandatory at Sr. No " & i & " in Schedule TDS2" & Chr(13)
                ValidatesheetTDS2 = False
                Exit Function
            End If

        End If
        
        If Len(Sheet28.Range("TDS2.DeductedTDS5A").item(i).value) > 0 Then
    
            If isdropdownblank(Sheet28.Range("TDS2.TDSCreditName").item(i).value) Then
                msgbox_TDS2 = msgbox_TDS2 + "* Enter All Mandatory Fields at Sr. No " & i & " in Schedule TDS2" & Chr(13)
                ValidatesheetTDS2 = False
                Exit Function
            End If
            
            If Len(Sheet28.Range("TDS2.DeductedIncome5A").item(i).value) = 0 Then
            
                msgbox_TDS2 = msgbox_TDS2 + "* Income is mandatory at Sr. No " & i & " in Schedule TDS2" & Chr(13)
                
                ValidatesheetTDS2 = False
                Exit Function
            End If
            
        End If


    Next
    
    For i = 1 To Sheet28.Range("TDS2.TAN").count
    
        If Len(Sheet28.Range("TDS2.ClaimedIncome5A").item(i).value) > 0 Then
    
            If isdropdownblank(Sheet28.Range("TDS2.TDSCreditName").item(i).value) Then
                msgbox_TDS2 = msgbox_TDS2 + "* Enter All Mandatory Fields at Sr. No " & i & "in Schedule TDS2" & Chr(13)
                ValidatesheetTDS2 = False
                Exit Function
            End If
            'Malli''''''''''''''''
            If Len(Sheet28.Range("TDS2.ClaimedTDS5A").item(i).value) = 0 Then
                'AY_2023_24 'Malli
                'msgbox_TDS2 = msgbox_TDS2 + "* Please enter TDS amount at Claimed in own hands at Sr. No " & i & " in Schedule TDS2" & Chr(13)
                 msgbox_TDS2 = msgbox_TDS2 + "* ""Please enter TDS amount at Claimed in spouse as per section 5A or any other person at Sr. No " & i & " in Schedule TDS2""" & Chr(13)
                '-------------------------------------
                ValidatesheetTDS2 = False
                Exit Function
            End If
            
            If Len(Sheet28.Range("TDS2.ClaimedPAN5A").item(i).value) = 0 Then
                msgbox_TDS2 = msgbox_TDS2 + "* Please enter the  PAN under Col(10) at Sr. No " & i & " in Schedule TDS2" & Chr(13)
                ValidatesheetTDS2 = False
                Exit Function
            End If
            
        End If
        
        If Len(Sheet28.Range("TDS2.ClaimedTDS5A").item(i).value) > 0 Then
    
            If isdropdownblank(Sheet28.Range("TDS2.TDSCreditName").item(i).value) Then
                msgbox_TDS2 = msgbox_TDS2 + "* Enter All Mandatory Fields at Sr. No " & i & " in Schedule TDS2" & Chr(13)
                ValidatesheetTDS2 = False
                Exit Function
            End If
            
            If Len(Sheet28.Range("TDS2.ClaimedIncome5A").item(i).value) = 0 Then
            
                'msgbox_TDS2 = msgbox_TDS2 + "* Please enter the  Income under Col(10) at Sr. No " & i & " in Schedule TDS2" & Chr(13)
                'AY_2024_25 Change 'Malli
                'msgbox_TDS2 = msgbox_TDS2 + "* Please enter Income at Claimed in own hands at Sr. No " & i & " in Schedule TDS2" & Chr(13)
                msgbox_TDS2 = msgbox_TDS2 + "* ""Please enter Income at Claimed in hands of spouse as per section 5A or any other person at Sr. No " & i & " in Schedule TDS2""" & Chr(13)
                '-----------------------------
                ValidatesheetTDS2 = False
                Exit Function
            End If
            
            If Len(Sheet28.Range("TDS2.ClaimedPAN5A").item(i).value) = 0 Then
                msgbox_TDS2 = msgbox_TDS2 + "* Please enter the PAN under Col(10) at Sr. No " & i & " in Schedule TDS2" & Chr(13)
                ValidatesheetTDS2 = False
                Exit Function
            End If
            
        End If
        
        If Len(Sheet28.Range("TDS2.ClaimedPAN5A").item(i).value) > 0 Then
    
            If isdropdownblank(Sheet28.Range("TDS2.TDSCreditName").item(i).value) Then
                msgbox_TDS2 = msgbox_TDS2 + "* Enter All Mandatory Fields at Sr. No " & i & " in Schedule TDS2" & Chr(13)
                ValidatesheetTDS2 = False
                Exit Function
            End If
            
            If Len(Sheet28.Range("TDS2.ClaimedIncome5A").item(i).value) = 0 Then
                'msgbox_TDS2 = msgbox_TDS2 + "* Please enter the  Income under Col(10) at Sr. No " & i & " in Schedule TDS2" & Chr(13)
                
                'Added by Aavula
                msgbox_TDS2 = msgbox_TDS2 + "*Please enter Income at Claimed in hands of spouse as per section 5A or any other person in schedule TDS2" & Chr(13)
                
                ValidatesheetTDS2 = False
                Exit Function
            End If
            
            If Len(Sheet28.Range("TDS2.ClaimedTDS5A").item(i).value) = 0 Then
                msgbox_TDS2 = msgbox_TDS2 + "* Please enter the  TDS under Col(10) at Sr. No " & i & " in Schedule TDS2" & Chr(13)
                ValidatesheetTDS2 = False
                Exit Function
            End If
            
        End If
        
        If Len(Sheet28.Range("ClaimedAadhar5A_TDS2").item(i).value) > 0 Then
            If Not ValidateAadhaar(CStr(Sheet28.Range("ClaimedAadhar5A_TDS2").item(i).value)) Then
            
                 msgbox_TDS2 = msgbox_TDS2 + "* Invalid Aadhaar No at Sr. No  " & i & "" & Chr(13)
                 ValidatesheetTDS2 = False
                 Exit Function
            End If
        End If
        
        
    Next
    
    If (end_TDS2 <> end_TDS22) Or (end_TDS2 <> end_TDS202) Then
             msgbox_TDS2 = msgbox_TDS2 + "* Enter all Mandatory Details in Schedule TDS2" & Chr(13)
             ValidatesheetTDS2 = False
    End If
    
End Function

Function ValidateCalculations_TDS2() As Boolean
ValidateCalculations_TDS2 = True
setTblinfo_TDS2

For i = 1 To end_TDS2

If Mid(TDSCreditName_TDS2(i), 1, 1) = "S" Then
    If TDSClaimedInOwnHands_TDS2(i) > GreaterValue(BroughtFwdTDSAmt_TDS2(i), DeductedInOwnHands_TDS2(i)) Then
       
'        msgbox_TDS2 = msgbox_TDS2 + "* Amount in field Col (10) at Sr. No  " & i & " cannot be more than Col (7) or Col (8) as the case may be in Schedule TDS2" & Chr(13)
        msgbox_TDS2 = msgbox_TDS2 + "* Amount in field Col (9) at Sr. No  " & i & " cannot be more than Col (6) or Col (7) as the case may be in Schedule TDS2" & Chr(13)
        ValidateCalculations_TDS2 = False
        Exit Function
    End If
End If

If Mid(TDSCreditName_TDS2(i), 1, 1) = "O" Then
'    If TDSClaimedInOwnHands_TDS2(i) > BroughtFwdTDSAmt_TDS2(i) + DeductedTDS5A_TDS2(i) Then
'    If TDSClaimedInOwnHands_TDS2(i) > (DeductedTDS5A_TDS2(i)) Then
    If (Sheet28.Range("TDS2.ClaimedInOwnHands").item(i, 1).value > Sheet28.Range("TDS2.BroughtFwdTDSAmt").item(i, 1).value + Sheet28.Range("TDS2.DeductedTDS5A").item(i, 1).value) Then
'        msgbox_TDS2 = msgbox_TDS2 + "* Amount in field Col (10) at Sr. No  " & i & " cannot be more than TDS field of Col (9) in Schedule TDS2" & Chr(13)
        msgbox_TDS2 = msgbox_TDS2 + "* Amount in field Col (9) at Sr. No  " & i & " cannot be more than Col (6) + TDS field of Col (8) in Schedule TDS2" & Chr(13)
        ValidateCalculations_TDS2 = False
        Exit Function
    End If
End If

If Mid(TDSCreditName_TDS2(i), 1, 1) = "S" Then
    If BroughtFwdTDSAmt_TDS2(i) > 0 Then
'        If DeductedInOwnHands_TDS2(I) <> "" Then'
''            msgbox_TDS2 = msgbox_TDS2 + "* ""Please enter the Unclaimed TDS brought forward"" and details of ""TDS of current FY"" should be provided in different rows in Schedule TDS2" & Chr(13)
''            ValidateCalculations_TDS2 = False
''            Exit Function
'        End If
    End If
    
    If DeductedInOwnHands_TDS2(i) <> "" Then
'        If (BroughtFwdTDSAmt_TDS2(I) <> "") Or (BroughtFwdTDSAmt_TDS2(I) > 0) Then
'            msgbox_TDS2 = msgbox_TDS2 + "* ""Please enter the Unclaimed TDS brought forward"" and details of ""TDS of current FY"" should be provided in different rows in Schedule TDS2" & Chr(13)
'            ValidateCalculations_TDS2 = False
'            Exit Function
'        End If
    End If
End If


Next
End Function

Function ValidateTDSCreditName_TDS2() As Boolean

    ValidateTDSCreditName_TDS2 = True
    setTblinfo_TDS2
    Dim rangecells As Range
    Set rangecells = Sheet28.Range("TDS2.TDSCreditName").Cells
    Dim rangecells1 As Range
    Set rangecells1 = Sheet28.Range("TDS2.PanOtherPerson").Cells
    Dim rangecells2 As Range
    Set rangecells2 = Sheet28.Range("TDSAadharOfOther_TDS2").Cells
    ReDim TDSCreditName_TDS2(end_TDS2)
    ReDim TDSPanOfOther_TDS2(end_TDS2)
    ReDim TDSAadharOfOther_TDS2(end_TDS2)
    For i = 1 To end_TDS2
     
        TDSCreditName_TDS2(i) = UCase(rangecells.item(i).value)
        TDSPanOfOther_TDS2(i) = UCase(rangecells1.item(i).value)
        TDSAadharOfOther_TDS2(i) = UCase(rangecells2.item(i).value)
        If isdropdownblank(TDSCreditName_TDS2(i)) And rangecells.item(i).Locked = False Then
            msgbox_TDS2 = msgbox_TDS2 + "* Please enter the TDS Credit Name at Sr. No  " & i & "in Schedule TDS2" & Chr(13)
              ValidateTDSCreditName_TDS2 = False
              Exit Function
        End If
        If Not checkfieldspecialcharacter(TDSCreditName_TDS2(i)) Then
              msgbox_TDS2 = msgbox_TDS2 + "* Special characters are not allowed in TDS Credit Name at Sr. No  " & i & "in Schedule TDS2" & Chr(13)
              ValidateTDSCreditName_TDS2 = False
              Exit Function
        End If
        
        If Mid(TDSCreditName_TDS2(i), 1, 1) = "O" Then
        If isdropdownblank(TDSPanOfOther_TDS2(i)) And rangecells1.item(i).Locked = False Then
            msgbox_TDS2 = msgbox_TDS2 + "* Please enter the PAN Of Other Person at Sr. No  " & i & " in Schedule TDS2" & Chr(13)
              ValidateTDSCreditName_TDS2 = False
              Exit Function
        End If
      
        If Len(TDSPanOfOther_TDS2(i)) > 0 Then
            If Not CheckDoneePAN(CStr(TDSPanOfOther_TDS2(i))) Then
             msgbox_TDS2 = msgbox_TDS2 + "* PAN Of Other Person at Sr. No  " & i & " is invalid. First 5 alphabets, next 4 digits, then alphabet in Schedule TDS2" & Chr(13)
             ValidateTDSCreditName_TDS2 = False
             Exit Function
            End If
       
            If Left(Range("sheet1.AsseseeRepFlg").value, 1) = "N" Then
            
'                If (UCase(TDSPanOfOther_TDS2(i)) =  Then
'                 msgbox_TDS2 = msgbox_TDS2 + "* ddddOther person PAN cannot be same as assesse PAN or verification PAN  at Sr. No  " & i & "  in Schedule TDS2" & Chr(13)
'                 ValidateTDSCreditName_TDS2 = False
'                 Exit Function
'                End If

                If (UCase(TDSPanOfOther_TDS2(i)) = UCase(Sheet1.Range("sheet1.PAN").value) Or UCase(TDSPanOfOther_TDS2(i)) = UCase(Sheet52.Range("sheet9.PAN").value)) Then
                 msgbox_TDS2 = msgbox_TDS2 + "* Other person PAN cannot be same as assesse PAN or verification PAN  at Sr. No  " & i & "  in Schedule TDS2" & Chr(13)
                 ValidateTDSCreditName_TDS2 = False
                 Exit Function
                End If
            ElseIf Left(Range("sheet1.AsseseeRepFlg").value, 1) = "Y" Then
                If UCase(TDSPanOfOther_TDS2(i)) = UCase(Sheet1.Range("sheet1.PAN").value) Then
                 msgbox_TDS2 = msgbox_TDS2 + "* Other person PAN cannot be same as assesse PAN at Sr. No  " & i & "  in Schedule TDS2" & Chr(13)
                 ValidateTDSCreditName_TDS2 = False
                 Exit Function
                End If
            End If
            
        End If
        
        If Len(TDSAadharOfOther_TDS2(i)) > 0 Then
        If Not ValidateAadhaar(CStr(TDSAadharOfOther_TDS2(i))) Then
        
         msgbox_TDS2 = msgbox_TDS2 + "* Aadhaar Of Other Person at Sr. No  " & i & " is invalid in Schedule TDS2" & Chr(13)
         ValidateTDSCreditName_TDS2 = False
         Exit Function
        End If
        End If
        
        
        End If
        
        
        
        
 Next
End Function

Function ValidateTAN_TDS2() As Boolean
    ValidateTAN_TDS2 = True
    setTblinfo_TDS2
    Dim rangecells As Range
    Set rangecells = Sheet28.Range("TDS2.TAN").Cells
    ReDim TAN_TDS2(end_TDS2)
    
    subProcCaption = "Validating TDS2"
    noOfProcessSub = end_TDS2
    
    For i = 1 To end_TDS2
        TAN_TDS2(i) = UCase(rangecells.item(i).value)
    If Not Len(TAN_TDS2(i)) = 0 Then
     
     If Not ValidateTantype_text(Mid(TAN_TDS2(i), 1, 4)) Then
         msgbox_TDS2 = msgbox_TDS2 + "* Please enter the TAN at Sr. No  " & i & " is invalid. First 4 alphabets, next 5 digits, then alphabet.in Schedule TDS2" & Chr(13)
         ValidateTAN_TDS2 = False
         Exit Function
     End If
     
     If Not IsNumeric(Mid(TAN_TDS2(i), 5, 5)) Then
         msgbox_TDS2 = msgbox_TDS2 + "* TAN at Sr. No  " & i & " is invalid. First 4 alphabets, next 5 digits, then alphabet.in Schedule TDS2" & Chr(13)
         ValidateTAN_TDS2 = False
         Exit Function
     End If
     
     If Not ValidateTantype_text(Right(TAN_TDS2(i), 1)) Then
         msgbox_TDS2 = msgbox_TDS2 + "* TAN at Sr. No  " & i & " is invalid. First 4 alphabets, next 5 digits, then alphabet.in Schedule TDS2" & Chr(13)
         ValidateTAN_TDS2 = False
         Exit Function
     End If
     
     If Not ValidateTANCodes(UCase(TAN_TDS2(i))) Then
                msgbox_TDS2 = msgbox_TDS2 + "* Invalid TAN at Sr.No " & i & "  Schedule TDS2.Please enter valid TAN" & Chr(13)
                ValidateTAN_TDS2 = False
                Exit Function
        End If
     
 ElseIf Not chkMandatory(TAN_TDS2(i)) Then
         msgbox_TDS2 = msgbox_TDS2 + "* Please enter the TAN at Sr. No  " & i & " in Schedule TDS2" & Chr(13)
     ValidateTAN_TDS2 = False
     Exit Function
 End If
 UpdateProgressBar
 Next
End Function
Function ValidateSectionTDSDeducted_TDS2() As Boolean
 
    ValidateSectionTDSDeducted_TDS2 = True
    setTblinfo_TDS2
    Dim rangecells As Range
    'Set rangecells = Sheet28.Range("TDS2.TdsDeducted").Cells    Konda updated on 28-05-2025
    Set rangecells = Sheet28.Range("TDS2.TdsDeducted").item(1, 1).Cells
    ReDim SectionTDSDeducted_TDS2(end_TDS2)
    For i = 1 To end_TDS2
        SectionTDSDeducted_TDS2(i) = rangecells.item(i).value
         
         If Not chkCompulsory(SectionTDSDeducted_TDS2(i)) Then
                msgbox_TDS2 = msgbox_TDS2 + "* ""Please select the Section under which TDS is deducted"" at Sr. No " & i & "" & Chr(13)
                ValidateSectionTDSDeducted_TDS2 = False
                Exit Function
        End If
        
        If (SectionTDSDeducted_TDS2(i) = "" Or SectionTDSDeducted_TDS2(i) = "(Select)") Then
                msgbox_TDS2 = msgbox_TDS2 + "* ""Please select the Section under which TDS is deducted"" at Sr. No " & i & "" & Chr(13)
                ValidateSectionTDSDeducted_TDS2 = False
                Exit Function
       End If
 Next
End Function

Function ValidateDeductedYear_TDS2() As Boolean
 
    ValidateDeductedYear_TDS2 = True
    setTblinfo_TDS2
    Dim rangecells As Range
    Set rangecells = Sheet28.Range("TDS2.DeductedYear").Cells
    ReDim DeductedYear_TDS2(end_TDS2)
    For i = 1 To end_TDS2
        DeductedYear_TDS2(i) = rangecells.item(i).value
        If isdropdownblank(DeductedYear_TDS2(i)) Then
            DeductedYear_TDS2(i) = ""
        End If
        If Not checkfieldspecialcharacter(DeductedYear_TDS2(i)) Then
              msgbox_TDS2 = msgbox_TDS2 + "* Special characters are not allowed in Deducted Year at Sr. No  " & i & " in Schedule TDS2" & Chr(13)
              ValidateDeductedYear_TDS2 = False
              Exit Function
        End If
 Next
End Function

Function ValidateBroughtFwdTDSAmt_TDS2() As Boolean
    ValidateBroughtFwdTDSAmt_TDS2 = True
    setTblinfo_TDS2
    
    Dim i As Long
    Dim rangecells As Range
    
    
    ReDim BroughtFwdTDSAmt_TDS2(end_TDS2)
    Set rangecells = Sheet28.Range("TDS2.BroughtFwdTDSAmt").Cells
    
    For i = 1 To end_TDS2
        BroughtFwdTDSAmt_TDS2(i) = rangecells.item(i).value
        
    If Len(BroughtFwdTDSAmt_TDS2(i)) > 14 Then
        msgbox_TDS2 = msgbox_TDS2 + "* Brought Forward Amount at Sr. No  " & i & "  should not be greater than 14 digits.in Schedule TDS2" & Chr(13)
        ValidateBroughtFwdTDSAmt_TDS2 = False
    End If
         
    Next
End Function


Function ValidateDeductedInOwnHands_TDS2() As Boolean
    ValidateDeductedInOwnHands_TDS2 = True
    setTblinfo_TDS2
    
    Dim i As Long
    Dim rangecells As Range
    
    ReDim DeductedInOwnHands_TDS2(end_TDS2)
    Set rangecells = Sheet28.Range("TDS2.DeductedInOwnHands").Cells
    
    For i = 1 To end_TDS2
    If rangecells.item(i).Locked = True Then
        DeductedInOwnHands_TDS2(i) = ""
    Else
        DeductedInOwnHands_TDS2(i) = rangecells.item(i).value
    End If
    If Len(DeductedInOwnHands_TDS2(i)) > 14 Then
        msgbox_TDS2 = msgbox_TDS2 + "* Amount Deducted In Own Hands at Sr. No  " & i & "  should not be greater than 14 digits.in Schedule TDS2" & Chr(13)
        ValidateDeductedInOwnHands_TDS2 = False
    End If
         
    Next
End Function

Function ValidateDeductedIncome5A_TDS2() As Boolean
    ValidateDeductedIncome5A_TDS2 = True
    setTblinfo_TDS28
    
    Dim i As Long
    Dim rangecells As Range
    
    ReDim DeductedIncome5A_TDS2(end_TDS28)
    Set rangecells = Sheet28.Range("TDS2.DeductedIncome5A").Cells
    
    For i = 1 To end_TDS28
    If rangecells.item(i).Locked = True Then
        DeductedIncome5A_TDS2(i) = ""
    Else
        DeductedIncome5A_TDS2(i) = rangecells.item(i).value
    End If

    If Len(DeductedIncome5A_TDS2(i)) > 14 Then
        msgbox_TDS2 = msgbox_TDS2 + "* Income under Col(8) at Sr. No  " & i & "  should not be greater than 14 digits.in Schedule TDS2" & Chr(13)
        ValidateDeductedIncome5A_TDS2 = False
    End If
         
    Next
End Function

Function ValidateDeductedTDS5A_TDS2() As Boolean
    ValidateDeductedTDS5A_TDS2 = True
    setTblinfo_TDS28
    
    Dim i As Long
    Dim rangecells As Range
    
    ReDim DeductedTDS5A_TDS2(end_TDS28)
    Set rangecells = Sheet28.Range("TDS2.DeductedTDS5A").Cells
    
    For i = 1 To end_TDS28
    If rangecells.item(i).Locked = True Then
        DeductedTDS5A_TDS2(i) = ""
    Else
        DeductedTDS5A_TDS2(i) = rangecells.item(i).value
    End If
    If Len(DeductedTDS5A_TDS2(i)) > 14 Then
        msgbox_TDS2 = msgbox_TDS2 + "* TDS under Col(8) at Sr. No  " & i & "  should not be greater than 14 digits.in Schedule TDS2" & Chr(13)
        ValidateDeductedTDS5A_TDS2 = False
    End If
         
    Next
End Function
Function ValidateTDSClaimedInOwnHands_TDS2() As Boolean
    ValidateTDSClaimedInOwnHands_TDS2 = True
    setTblinfo_TDS2
    
    Dim i As Long
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    
    Set rangecells = Sheet28.Range("TDS2.ClaimedInOwnHands").Cells
    Set rangecells1 = Sheet28.Range("TDS2.GrossAmount").Cells
    Set rangecells2 = Sheet28.Range("TDS2.HeadOfIncome").Cells
    
    ReDim TDSClaimedInOwnHands_TDS2(end_TDS2)
    ReDim GrossAmount_TDS2(end_TDS2)
    ReDim HeadIncome_TDS2(end_TDS2)
    
    
    For i = 1 To end_TDS2
        TDSClaimedInOwnHands_TDS2(i) = rangecells.item(i).value
        GrossAmount_TDS2(i) = rangecells1.item(i).value
        HeadIncome_TDS2(i) = rangecells2.item(i).value
        
    If Not chkMandatory(TDSClaimedInOwnHands_TDS2(i)) And rangecells.item(i).Locked = False Then
         msgbox_TDS2 = msgbox_TDS2 + "* Please enter the TDS Claimed in Own Hands at Sr. No  " & i & "in Schedule TDS2" & Chr(13)
         ValidateTDSClaimedInOwnHands_TDS2 = False
         Exit Function
    End If
    
    If TDSClaimedInOwnHands_TDS2(i) > 0 Then
    If Mid(HeadIncome_TDS2(i), 1, 1) <> "N" Then
    If Not chkCompulsory(GrossAmount_TDS2(i)) Then
     msgbox_TDS2 = msgbox_TDS2 + "* Please enter the Gross Amount at Sr. No  " & i & "in Schedule TDS2" & Chr(13)
         ValidateTDSClaimedInOwnHands_TDS2 = False
         Exit Function
    End If
    End If
    
    If isdropdownblank(HeadIncome_TDS2(i)) Then
'     msgbox_TDS2 = msgbox_TDS2 + "* Please enter the Head Income at Sr. No  " & i & "in Schedule TDS2" & Chr(13)
    'SIT-69628  Changed by sai
     msgbox_TDS2 = msgbox_TDS2 + "*""Please select the drop down of head of income for which corresponding income offered in schedule TDS2"" at Sr.No " & i & " " & Chr(13)
     
     
         ValidateTDSClaimedInOwnHands_TDS2 = False
         Exit Function
    End If
    End If
        
    If Len(TDSClaimedInOwnHands_TDS2(i)) > 14 Then
        msgbox_TDS2 = msgbox_TDS2 + "* TDS Claimed in Own Hands at Sr. No  " & i & "  should not be greater than 14 digits in Schedule TDS2" & Chr(13)
        ValidateTDSClaimedInOwnHands_TDS2 = False
    End If
         
    Next
End Function

Function ValidateClaimedIncome5A_TDS2() As Boolean
    ValidateClaimedIncome5A_TDS2 = True
    setTblinfo_TDS203
    
    Dim i As Long
    Dim rangecells As Range
    
    ReDim ClaimedIncome5A_TDS2(end_TDS203)
    Set rangecells = Sheet28.Range("TDS2.ClaimedIncome5A").Cells
    
    For i = 1 To end_TDS203
        ClaimedIncome5A_TDS2(i) = rangecells.item(i).value
        
    If Len(ClaimedIncome5A_TDS2(i)) > 14 Then
        msgbox_TDS2 = msgbox_TDS2 + "* Income under Col(10) at Sr. No  " & i & "  should not be greater than 14 digits.in Schedule TDS2" & Chr(13)
        ValidateClaimedIncome5A_TDS2 = False
    End If
         
    Next
End Function

Function ValidateClaimedTDS5A_TDS2() As Boolean
    ValidateClaimedTDS5A_TDS2 = True
    setTblinfo_TDS203
    
    Dim i As Long
    Dim rangecells As Range
    
    ReDim ClaimedTDS5A_TDS2(end_TDS203)
    Set rangecells = Sheet28.Range("TDS2.ClaimedTDS5A").Cells
    
    For i = 1 To end_TDS203
        ClaimedTDS5A_TDS2(i) = rangecells.item(i).value
        
    If Len(ClaimedTDS5A_TDS2(i)) > 14 Then
        msgbox_TDS2 = msgbox_TDS2 + "* TDS under Col(10) at Sr. No  " & i & "  should not be greater than 14 digits.in Schedule TDS2" & Chr(13)
        ValidateClaimedTDS5A_TDS2 = False
    End If
         
    Next
End Function

Function ValidateClaimedPAN5A_TDS2() As Boolean
ValidateClaimedPAN5A_TDS2 = True
Dim rangecells As Range
Dim rangecells1 As Range
Set rangecells = Sheet28.Range("TDS2.ClaimedPAN5A").Cells
Set rangecells1 = Sheet28.Range("ClaimedAadhar5A_TDS2").Cells
Dim i As Long
ReDim ClaimedPAN5A_TDS2(end_TDS203)
ReDim ClaimedAadhar5A_TDS2(end_TDS203)
For i = 1 To end_TDS203
ClaimedPAN5A_TDS2(i) = rangecells.item(i).value
ClaimedAadhar5A_TDS2(i) = rangecells1.item(i).value
    If Not checkfieldspecialcharacter(ClaimedPAN5A_TDS2(i)) Then
         msgbox_TDS2 = msgbox_TDS2 + "* Please enter the PAN under Col(10) at Sr. No  " & i & "  characters < > & ' " & Chr(34) & " are not allowed in Schedule TDS2" & Chr(13)
         ValidateClaimedPAN5A_TDS2 = False
         Exit Function
    End If
    
    If Not CheckPAN(UCase(ClaimedPAN5A_TDS2(i))) Then
         msgbox_TDS2 = msgbox_TDS2 + "* PAN under Col(10) at Sr. No  " & i & " in Sheet TDS is Invalid" & Chr(34) & "  PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet.in Schedule TDS2" & Chr(13)
         ValidateClaimedPAN5A_TDS2 = False
         Exit Function
    End If
    
    If UCase(ClaimedPAN5A_TDS2(i)) = UCase(Sheet1.Range("sheet1.PAN").value) Then
    
        msgbox_TDS2 = msgbox_TDS2 + "* Other person PAN cannot be same as assesse PAN at Sr. No  " & i & "  in Schedule TDS2" & Chr(13)
        ValidateClaimedPAN5A_TDS2 = False
        Exit Function
    End If
                
    If Len(ClaimedAadhar5A_TDS2(i)) > 0 Then
        If Not ValidateAadhaar(CStr(ClaimedAadhar5A_TDS2(i))) Then
         msgbox_TDS2 = msgbox_TDS2 + "* Aadhaar Of Other Person at Sr. No  " & i & " is invalid in Schedule TDS2" & Chr(13)
         ValidateClaimedPAN5A_TDS2 = False
         Exit Function
        End If
        End If

Next
End Function

Function ValidateAmtCarriedForward_TDS2() As Boolean
    ValidateAmtCarriedForward_TDS2 = True
    setTblinfo_TDS2
    
    Dim i As Long
    Dim rangecells As Range
    
    ReDim AmtCarriedForward_TDS2(end_TDS2)
    Set rangecells = Sheet28.Range("TDS2.AmtCarriedForward").Cells
    
    For i = 1 To end_TDS2
        AmtCarriedForward_TDS2(i) = rangecells.item(i).value
        
    If Not chkMandatory(AmtCarriedForward_TDS2(i)) And rangecells.item(i).Locked = False Then
         msgbox_TDS2 = msgbox_TDS2 + "* Please enter the Amount Carried Forward at Sr. No  " & i & " in Schedule TDS2" & Chr(13)
         ValidateAmtCarriedForward_TDS2 = False
         Exit Function
    End If
        
    If Len(AmtCarriedForward_TDS2(i)) > 14 Then
        msgbox_TDS2 = msgbox_TDS2 + "* Amount Carried Forward at Sr. No  " & i & "  should not be greater than 14 digits.in Schedule TDS2" & Chr(13)
        ValidateAmtCarriedForward_TDS2 = False
    End If
         
    Next
End Function


Sub setTblinfo_TDS2()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet28.Range("TDS2.TDSCreditName").count
 Set rangecells = Sheet28.Range("TDS2.TDSCreditName").Cells
 For mIntCtr = 1 To mIntCells
     If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
    End If
 Next
 end_TDS2 = ccount
 rngname_TDS2 = "TDS2.TDSCreditName;TDS2.GrossAmount;TDS2.HeadOfIncome;TDS2.PanOtherPerson;TDSAadharOfOther_TDS2;TDS2.TAN;TDS2.TdsDeducted;TDS2.DeductedYear;TDS2.BroughtFwdTDSAmt;TDS2.DeductedInOwnHands;TDS2.DeductedIncome5A;TDS2.DeductedTDS5A;TDS2.ClaimedInOwnHands;TDS2.ClaimedIncome5A;TDS2.ClaimedTDS5A;c;ClaimedAadhar5A_TDS2;TDS2.AmtCarriedForward;"
 End Sub
 
 
 Sub setTblinfo_TDS2_Pan()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet28.Range("TDS2.PanOtherPerson").count
 Set rangecells = Sheet28.Range("TDS2.PanOtherPerson").Cells
 For mIntCtr = 1 To mIntCells
    If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS2_Pan = ccount
 End Sub
 
 Sub setTblinfo_TDS2_GrossAmount()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet28.Range("TDS2.GrossAmount").count
 Set rangecells = Sheet28.Range("TDS2.GrossAmount").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS2_GrossAmount = ccount
 End Sub
 
 Sub setTblinfo_TDS2_Head()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet28.Range("TDS2.HeadOfIncome").count
 Set rangecells = Sheet28.Range("TDS2.HeadOfIncome").Cells
 For mIntCtr = 1 To mIntCells
     If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
    End If
 Next
 end_TDS2_Head = ccount
 End Sub
 
 Sub setTblinfo_TDS22()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet28.Range("TDS2.TAN").count
 Set rangecells = Sheet28.Range("TDS2.TAN").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS22 = ccount
 End Sub
 
Sub setTblinfo_TDS23()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet28.Range("TDS2.TdsDeducted").count
 Set rangecells = Sheet28.Range("TDS2.TdsDeducted").Cells
 For mIntCtr = 1 To mIntCells
     If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
    End If
 Next
 end_TDS23 = ccount
 End Sub
 
 Sub setTblinfo_TDS25()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 Dim Year_TDS As Variant
 ccount = 0
 mIntCells = Sheet28.Range("TDS2.DeductedYear").count
 ReDim Year_TDS(mIntCells)
 Set rangecells = Sheet28.Range("TDS2.DeductedYear").Cells
 For mIntCtr = 1 To mIntCells
    Year_TDS(mIntCtr) = rangecells.item(mIntCtr).value
    If isdropdownblank(Year_TDS(mIntCtr)) Then
        Year_TDS(mIntCtr) = ""
    End If
    If Not Year_TDS(mIntCtr) = "" Then
        ccount = ccount + 1
    End If
 Next
    If ccount <= end_TDS2 Then
    end_TDS25 = end_TDS2
    Else
    end_TDS25 = ccount
    End If
 End Sub
 
 Sub setTblinfo_TDS26()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet28.Range("TDS2.BroughtFwdTDSAmt").count
 Set rangecells = Sheet28.Range("TDS2.BroughtFwdTDSAmt").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 If ccount <= end_TDS2 Then
    end_TDS26 = end_TDS2
Else
 end_TDS26 = ccount
End If
 End Sub
 
Sub setTblinfo_TDS27()
Dim rangecells As Range
Dim mIntCells  As Long
Dim mIntCtr  As Long
Dim ccount  As Long
ccount = 0
mIntCells = Sheet28.Range("TDS2.DeductedInOwnHands").count
Set rangecells = Sheet28.Range("TDS2.DeductedInOwnHands").Cells
For mIntCtr = 1 To mIntCells
    If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
    End If
Next
If ccount <= end_TDS2 Then
end_TDS27 = end_TDS2
Else
end_TDS27 = ccount
End If
End Sub

Sub setTblinfo_TDS28()
Dim rangecells As Range
Dim mIntCells  As Long
Dim mIntCtr  As Long
Dim ccount  As Long
ccount = 0
mIntCells = Sheet28.Range("TDS2.DeductedIncome5A").count
Set rangecells = Sheet28.Range("TDS2.DeductedIncome5A").Cells
For mIntCtr = 1 To mIntCells
    If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
    End If
Next

If ccount <= end_TDS2 Then
end_TDS28 = end_TDS2
Else
end_TDS28 = ccount
End If
End Sub

Sub setTblinfo_TDS29()
Dim rangecells As Range
Dim mIntCells  As Long
Dim mIntCtr  As Long
Dim ccount  As Long
ccount = 0
mIntCells = Sheet28.Range("TDS2.DeductedTDS5A").count
Set rangecells = Sheet28.Range("TDS2.DeductedTDS5A").Cells
For mIntCtr = 1 To mIntCells
    If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
    End If
Next
If ccount <= end_TDS2 Then
end_TDS29 = end_TDS2
Else
end_TDS29 = ccount
End If
End Sub

Sub setTblinfo_TDS202()
Dim rangecells As Range
Dim mIntCells  As Long
Dim mIntCtr  As Long
Dim ccount  As Long
ccount = 0
mIntCells = Sheet28.Range("TDS2.ClaimedInOwnHands").count
Set rangecells = Sheet28.Range("TDS2.ClaimedInOwnHands").Cells
For mIntCtr = 1 To mIntCells
    If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
    End If
Next
end_TDS202 = ccount
End Sub

Sub setTblinfo_TDS203()
Dim rangecells As Range
Dim mIntCells  As Long
Dim mIntCtr  As Long
Dim ccount  As Long
ccount = 0
mIntCells = Sheet28.Range("TDS2.ClaimedIncome5A").count
Set rangecells = Sheet28.Range("TDS2.ClaimedIncome5A").Cells
For mIntCtr = 1 To mIntCells
    If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
    End If
Next
If ccount <= end_TDS2 Then
end_TDS203 = end_TDS2
Else
end_TDS203 = ccount
End If
End Sub

Sub setTblinfo_TDS204()
Dim rangecells As Range
Dim mIntCells  As Long
Dim mIntCtr  As Long
Dim ccount  As Long
ccount = 0
mIntCells = Sheet28.Range("TDS2.ClaimedTDS5A").count
Set rangecells = Sheet28.Range("TDS2.ClaimedTDS5A").Cells
For mIntCtr = 1 To mIntCells
    If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
    End If
Next
If ccount <= end_TDS2 Then
end_TDS204 = end_TDS2
Else
end_TDS204 = ccount
End If
End Sub

Sub setTblinfo_TDS205()
Dim rangecells As Range
Dim mIntCells  As Long
Dim mIntCtr  As Long
Dim ccount  As Long
ccount = 0
mIntCells = Sheet28.Range("TDS2.ClaimedPAN5A").count
Set rangecells = Sheet28.Range("TDS2.ClaimedPAN5A").Cells
For mIntCtr = 1 To mIntCells
    If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
    End If
Next
If ccount <= end_TDS2 Then
end_TDS205 = end_TDS2
Else
end_TDS205 = ccount
End If
End Sub

Sub setTblinfo_TDS206()
Dim rangecells As Range
Dim mIntCells  As Long
Dim mIntCtr  As Long
Dim ccount  As Long
ccount = 0
mIntCells = Sheet28.Range("TDS2.AmtCarriedForward").count
Set rangecells = Sheet28.Range("TDS2.AmtCarriedForward").Cells
For mIntCtr = 1 To mIntCells
    If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
    End If
Next
end_TDS206 = ccount
End Sub

 
'TDS3
Function ValidatesheetTDS3() As Boolean
    ValidatesheetTDS3 = True
    msgbox_TDS3 = "Schedule TDS3 :" & Chr(10)

    If Len(Sheet28.Range("TDS3.Total").value) > 14 Then
       msgbox_TDS3 = msgbox_TDS3 & " Total of TDS3 should not excced 14 digits " & Chr(13)
       ValidatesheetTDS3 = False
       Exit Function
    End If
    
    setTblinfo_TDS3
    setTblinfo_TDS3_Pan
    setTblinfo_TDS3_GrossAmount
    setTblinfo_TDS3_Head
    setTblinfo_TDS32
    'setTblinfo_TDS33
    'setTblinfo_TDS34
    setTblinfo_TDS35
    setTblinfo_TDS36
    setTblinfo_TDS37
    setTblinfo_TDS38
    setTblinfo_TDS39
    'setTblinfo_TDS301
    setTblinfo_TDS302
    setTblinfo_TDS303
    setTblinfo_TDS304
    setTblinfo_TDS305
    setTblinfo_TDS306

    For i = 1 To end_TDS3
        If isdropdownblank(Sheet28.Range("TDS3.TDSCreditName").item(i).value) Then
            msgbox_TDS3 = msgbox_TDS3 + "* Please fill all the mandatory fields of the row " & i & " before filling the next row.in Schedule TDS3" & Chr(13)
            ValidatesheetTDS3 = False
            Exit Function
        End If
    Next

    If Not ValidateTDSCreditName_TDS3() Then ValidatesheetTDS3 = False
    If Not ValidatePAN_TDS3() Then ValidatesheetTDS3 = False
    If (Len(Sheet28.Range("TDS3.TDSCreditName").item(1).value) > 0) Then
        
        'If Not ValidateEmployerOrDeductorOrCollecterName_TDS3() Then ValidatesheetTDS3 = False
        'If Not ValidateUTN_TDS3() Then ValidatesheetTDS3 = False
        
        If Not SectionTDSisDeducted_TDS3 Then ValidatesheetTDS3 = False 'Added by Aavula Naresh for AY 2025-26
        
        If Not ValidateDeductedYear_TDS3() Then ValidatesheetTDS3 = False
        If Not ValidateBroughtFwdTDSAmt_TDS3() Then ValidatesheetTDS3 = False
        If Not ValidateDeductedInOwnHands_TDS3() Then ValidatesheetTDS3 = False
        If Not ValidateDeductedIncome5A_TDS3() Then ValidatesheetTDS3 = False
        If Not ValidateDeductedTDS5A_TDS3() Then ValidatesheetTDS3 = False
        'If Not ValidateDeductedPAN5A_TDS3() Then ValidatesheetTDS3 = False
        If Not ValidateTDSClaimedInOwnHands_TDS3 Then ValidatesheetTDS3 = False
        If Not ValidateClaimedIncome5A_TDS3 Then ValidatesheetTDS3 = False
        If Not ValidateClaimedTDS5A_TDS3 Then ValidatesheetTDS3 = False
        If Not ValidateClaimedPAN5A_TDS3 Then ValidatesheetTDS3 = False
        If Not ValidateAmtCarriedForward_TDS3 Then ValidatesheetTDS3 = False
        If Not ValidateCalculations_TDS3 Then ValidatesheetTDS3 = False

    End If
    
    For i = 1 To Sheet28.Range("TDS3.PAN").count

        If Len(Sheet28.Range("TDS3.BroughtFwdTDSAmt").item(i).value) > 0 Then
    
            If Len(Sheet28.Range("TDS3.PAN").item(i).value) = 0 Then
                msgbox_TDS3 = msgbox_TDS3 + "* Enter All Mandatory Fields at Sr. No " & i & " in Schedule TDS3" & Chr(13)
                ValidatesheetTDS3 = False
                Exit Function
            End If
                                  
            If isdropdownblank(Sheet28.Range("TDS3.DeductedYear").item(i).value) Then
'                msgbox_TDS3 = msgbox_TDS3 + "* Please select the financial year in which TDS was deducted at Sr. No " & i & " in Schedule TDS3" & Chr(13)
               'Newly added by Sai on 26/02/2024
                 msgbox_TDS3 = msgbox_TDS3 + "* ""Please select Year of tax deduction from drop-down in schedule TDS3"" at Sr. No " & i & " " & Chr(13)
                ValidatesheetTDS3 = False
                Exit Function
            End If
        End If
        
        If Not isdropdownblank(Sheet28.Range("TDS3.DeductedYear").item(i).value) Then
            If Len(Sheet28.Range("TDS3.BroughtFwdTDSAmt").item(i).value) = 0 Then
               ' msgbox_TDS3 = msgbox_TDS3 + "* Please enter TDS b/f at Sr. No " & i & " in Schedule TDS3" & Chr(13)
                'Ankita
                 msgbox_TDS3 = msgbox_TDS3 + "* Please select the drop down for the field TDS b/f at Sr. No " & i & "  in schedule TDS3" & Chr(13)
                 ValidatesheetTDS3 = False
                Exit Function
            End If
        End If
        
        If Len(Sheet28.Range("TDS3.BroughtFwdTDSAmt").item(i).value) <> 0 Then
            If Len(Sheet28.Range("TDS3.DeductedIncome5A").item(i).value) <> 0 Then
            
                msgbox_TDS3 = msgbox_TDS3 + "* Please enter the Unclaimed TDS brought forward and Income details of TDS Deducted in the hands of any other person as per rule 37BA(2) should be provided in different rows at Sr. No " & i & " in Schedule TDS3" & Chr(13)
                ValidatesheetTDS3 = False
                Exit Function
            End If
        End If
        
        If Len(Sheet28.Range("TDS3.DeductedIncome5A").item(i).value) <> 0 Then
            If Len(Sheet28.Range("TDS3.BroughtFwdTDSAmt").item(i).value) <> 0 Then
           
                msgbox_TDS3 = msgbox_TDS3 + "* Please enter the Unclaimed TDS brought forward and Income details of TDS Deducted in the hands of any other person as per rule 37BA(2) should be provided in different rows at Sr. No " & i & " in Schedule TDS3" & Chr(13)
                ValidatesheetTDS3 = False
                Exit Function
            End If
        End If
        
        
    Next
    
    For i = 1 To Sheet28.Range("TDS3.PAN").count
    
        If Len(Sheet28.Range("TDS3.DeductedIncome5A").item(i).value) > 0 Then
    
            If isdropdownblank(Sheet28.Range("TDS3.TDSCreditName").item(i).value) Then
'                msgbox_TDS3 = msgbox_TDS3 + "* Enter All Mandatory Fields at Sr. No " & i & " in Schedule TDS3" & Chr(13)
             'SIT-70656 Changed by sai
                msgbox_TDS3 = msgbox_TDS3 + "* Please select the drop down for the field ""TDS credit relating to Self/Other Person in schedule TDS3"" at Sr. No " & i & "" & Chr(13)
                
                ValidatesheetTDS3 = False
                Exit Function
            End If
            
            If Len(Sheet28.Range("TDS3.DeductedTDS5A").item(i).value) = 0 Then
            
                msgbox_TDS3 = msgbox_TDS3 + "* TDS is mandatory at Sr. No " & i & " in Schedule TDS3" & Chr(13)
                ValidatesheetTDS3 = False
                Exit Function
            End If

        End If
        
        If Len(Sheet28.Range("TDS3.DeductedTDS5A").item(i).value) > 0 Then
    
            If isdropdownblank(Sheet28.Range("TDS3.TDSCreditName").item(i).value) Then
                msgbox_TDS3 = msgbox_TDS3 + "* Enter All Mandatory Fields at Sr. No " & i & " in Schedule TDS3" & Chr(13)
                ValidatesheetTDS3 = False
                Exit Function
            End If
            
            If Len(Sheet28.Range("TDS3.DeductedIncome5A").item(i).value) = 0 Then
                msgbox_TDS3 = msgbox_TDS3 + "* Income is mandatory at Sr. No " & i & " in Schedule TDS3" & Chr(13)
                ValidatesheetTDS3 = False
                Exit Function
            End If

        End If
    Next
    
    For i = 1 To Sheet28.Range("TDS3.PAN").count
    
        If Len(Sheet28.Range("TDS3.ClaimedIncome5A").item(i).value) > 0 Then
    
            If isdropdownblank(Sheet28.Range("TDS3.TDSCreditName").item(i).value) Then
                msgbox_TDS3 = msgbox_TDS3 + "* TDS credit relating to Self/Other Person in schedule TDS3 at Sr. No " & i & "" & Chr(13)
                ValidatesheetTDS3 = False
                Exit Function
            End If
            
            If Len(Sheet28.Range("TDS3.ClaimedTDS5A").item(i).value) = 0 Then
              '  msgbox_TDS3 = msgbox_TDS3 + "* Enter TDS under Col(10) at Sr. No " & i & " in Schedule TDS3" & Chr(13)
                'Added by ramya 70643
                 msgbox_TDS3 = msgbox_TDS3 + "* Please enter TDS amount at Column 10 in schedule TDS3" & Chr(13)

                ValidatesheetTDS3 = False
                Exit Function
            End If
            
            If Len(Sheet28.Range("TDS3.ClaimedPAN5A").item(i).value) = 0 Then
                msgbox_TDS3 = msgbox_TDS3 + "* Please Enter the PAN under Col(10) at Sr. No " & i & " in Schedule TDS3" & Chr(13)
                ValidatesheetTDS3 = False
                Exit Function
            End If
            
        End If
        
        If Len(Sheet28.Range("TDS3.ClaimedTDS5A").item(i).value) > 0 Then
    
            If isdropdownblank(Sheet28.Range("TDS3.TDSCreditName").item(i).value) Then
                msgbox_TDS3 = msgbox_TDS3 + "* TDS credit relating to Self/Other Person in schedule TDS3 at Sr. No " & i & "" & Chr(13)
                ValidatesheetTDS3 = False
                Exit Function
            End If
            
            If Len(Sheet28.Range("TDS3.ClaimedIncome5A").item(i).value) = 0 Then
                msgbox_TDS3 = msgbox_TDS3 + "* Please enter the Income under Col(10) at Sr. No " & i & " in Schedule TDS3" & Chr(13)
                ValidatesheetTDS3 = False
                Exit Function
            End If
            
            If Len(Sheet28.Range("TDS3.ClaimedPAN5A").item(i).value) = 0 Then
                msgbox_TDS3 = msgbox_TDS3 + "* Please enter the PAN under Col(10) at Sr. No " & i & " in Schedule TDS3" & Chr(13)
                ValidatesheetTDS3 = False
                Exit Function
            End If
            
        End If
        
        If Len(Sheet28.Range("TDS3.ClaimedPAN5A").item(i).value) > 0 Then
    
            If isdropdownblank(Sheet28.Range("TDS3.TDSCreditName").item(i).value) Then
                msgbox_TDS3 = msgbox_TDS3 + "* TDS credit relating to Self/Other Person in schedule TDS3 at Sr. No " & i & "" & Chr(13)
                ValidatesheetTDS3 = False
                Exit Function
            End If
            
            If Len(Sheet28.Range("TDS3.ClaimedIncome5A").item(i).value) = 0 Then
                msgbox_TDS3 = msgbox_TDS3 + "* Please enter the Income under Col(10) at Sr. No " & i & " in Schedule TDS3" & Chr(13)
                ValidatesheetTDS3 = False
                Exit Function
            End If
            
            If Len(Sheet28.Range("TDS3.ClaimedTDS5A").item(i).value) = 0 Then
                msgbox_TDS3 = msgbox_TDS3 + "* Please enter the TDS under Col(10) at Sr. No " & i & " in Schedule TDS3" & Chr(13)
                ValidatesheetTDS3 = False
                Exit Function
            End If
            
        End If
        
        
        
        If Len(Sheet28.Range("ClaimedAadhar5A_TDS3").item(i).value) > 0 Then
        If Not ValidateAadhaar(CStr(Sheet28.Range("ClaimedAadhar5A_TDS3").item(i).value)) Then
         msgbox_TDS3 = msgbox_TDS3 + "* Aadhaar Of Other Person at Sr. No  " & i & " is invalid in Schedule TDS3" & Chr(13)
         ValidatesheetTDS3 = False
         Exit Function
        End If
        End If
        
    Next
    
    If (end_TDS3 <> end_TDS32) Or (end_TDS3 <> end_TDS302) Then
             msgbox_TDS3 = msgbox_TDS3 + "* TDS credit relating to Self/Other Person in schedule TDS3" & Chr(13) 'Added by Aavula
             ValidatesheetTDS3 = False
    End If
    
End Function

Function ValidateCalculations_TDS3() As Boolean
ValidateCalculations_TDS3 = True
setTblinfo_TDS3

For i = 1 To end_TDS3

If Mid(TDSCreditName_TDS3(i), 1, 1) = "S" Then
    If TDSClaimedInOwnHands_TDS3(i) > GreaterValue(BroughtFwdTDSAmt_TDS3(i), DeductedInOwnHands_TDS3(i)) Then
'Change-11.05.2023.101 -> SIT-50688
'        msgbox_TDS3 = msgbox_TDS3 + "* Amount in field Col (10) at Sr. No  " & i & " cannot be more than Col (7) or Col (8) as the case may be in Schedule TDS3" & Chr(13)
        msgbox_TDS3 = msgbox_TDS3 + "* Amount in field Col (9) at Sr. No  " & i & " cannot be more than Col (6) or Col (7) as the case may be in Schedule TDS3" & Chr(13)
        ValidateCalculations_TDS3 = False
        Exit Function
    End If
End If

If Mid(TDSCreditName_TDS3(i), 1, 1) = "O" Then
    If TDSClaimedInOwnHands_TDS3(i) > (BroughtFwdTDSAmt_TDS3(i) + DeductedTDS5A_TDS3(i)) Then
'Change-11.05.2023.101 -> SIT-50688
'        msgbox_TDS3 = msgbox_TDS3 + "* Amount in field Col (10) at Sr. No  " & i & " cannot be more than Col (7) + TDS field of Col (9) in Schedule TDS3" & Chr(13)
        msgbox_TDS3 = msgbox_TDS3 + "* Amount in field Col (9) at Sr. No  " & i & " cannot be more than Col (6) + TDS field of Col (8) in Schedule TDS3" & Chr(13)
        ValidateCalculations_TDS3 = False
        Exit Function
    End If
End If

If Mid(TDSCreditName_TDS3(i), 1, 1) = "S" Then
    If BroughtFwdTDSAmt_TDS3(i) > 0 Then
'        If DeductedInOwnHands_TDS3(I) <> "" Then'
'            msgbox_TDS3 = msgbox_TDS3 + "* ""Unclaimed TDS brought forward"" and details of ""TDS of current FY"" should be provided in different rows in Schedule TDS3" & Chr(13)
'            ValidateCalculations_TDS3 = False
'            Exit Function
'        End If
    End If
    
    If DeductedInOwnHands_TDS3(i) <> "" Then
'        If (BroughtFwdTDSAmt_TDS3(I) <> "") Or (BroughtFwdTDSAmt_TDS3(I) <> 0) Then
'            msgbox_TDS3 = msgbox_TDS3 + "* ""Unclaimed TDS brought forward"" and details of ""TDS of current FY"" should be provided in different rows in Schedule TDS3" & Chr(13)
'            ValidateCalculations_TDS3 = False
'            Exit Function
'        End If
    End If
End If

Next
End Function

Function ValidateTDSCreditName_TDS3() As Boolean
 
    ValidateTDSCreditName_TDS3 = True
    setTblinfo_TDS3
    Dim rangecells As Range
    Set rangecells = Sheet28.Range("TDS3.TDSCreditName").Cells
    Dim rangecells1 As Range
    Set rangecells1 = Sheet28.Range("TDS3.PanOtherPerson").Cells
    Dim rangecells2 As Range
    Set rangecells2 = Sheet28.Range("Aadhar_TDS3").Cells
    Dim rangecells3 As Range
    Set rangecells3 = Sheet28.Range("TDS3.PAN").Cells
    ReDim TDSCreditName_TDS3(end_TDS3)
    ReDim TDSPanOfOther_TDS3(end_TDS3)
    ReDim TDSPan_TDS3(end_TDS3)
    ReDim Aadhar_TDS3(end_TDS3)
    For i = 1 To end_TDS3
        TDSCreditName_TDS3(i) = rangecells.item(i).value
        TDSPanOfOther_TDS3(i) = rangecells1.item(i).value
        Aadhar_TDS3(i) = rangecells2.item(i).value
        TDSPan_TDS3(i) = rangecells3.item(i).value
        
        If isdropdownblank(TDSCreditName_TDS3(i)) And rangecells.item(i).Locked = False Then
            msgbox_TDS3 = msgbox_TDS3 + "* TDS credit relating to Self/Other Person in schedule TDS3 at Sr. No  " & i & "." & Chr(13)
              ValidateTDSCreditName_TDS3 = False
              Exit Function
        End If
        If Not checkfieldspecialcharacter(TDSCreditName_TDS3(i)) Then
              msgbox_TDS3 = msgbox_TDS3 + "* Special characters are not allowed in TDS Credit Name at Sr. No  " & i & " in Schedule TDS3" & Chr(13)
              ValidateTDSCreditName_TDS3 = False
              Exit Function
        End If
        
        If Mid(TDSCreditName_TDS3(i), 1, 1) = "O" Then
        If isdropdownblank(TDSPanOfOther_TDS3(i)) And rangecells1.item(i).Locked = False Then
            msgbox_TDS3 = msgbox_TDS3 + "* Please enter the PAN Of Other Person at Sr. No  " & i & " in Schedule TDS3" & Chr(13)
              ValidateTDSCreditName_TDS3 = False
              Exit Function
        End If

        If Len(TDSPanOfOther_TDS3(i)) > 0 Then
        
                If Not CheckDoneePAN(CStr(TDSPanOfOther_TDS3(i))) Then
                    msgbox_TDS3 = msgbox_TDS3 + "* PAN Of Other Person at Sr. No  " & i & " is invalid. First 5 alphabets, next 4 digits, then alphabet in Schedule TDS3" & Chr(13)
                    ValidateTDSCreditName_TDS3 = False
                    Exit Function
                End If
                
            If Left(Range("sheet1.AsseseeRepFlg").value, 1) = "N" Then
                If (UCase(TDSPanOfOther_TDS3(i)) = UCase(Sheet1.Range("sheet1.PAN").value) Or UCase(TDSPanOfOther_TDS3(i)) = UCase(Sheet52.Range("sheet9.PAN").value)) Then
                    msgbox_TDS3 = msgbox_TDS3 + "* Other person PAN cannot be same as assesse PAN or verification PAN  at Sr. No  " & i & "  in Schedule TDS3" & Chr(13)
                    ValidateTDSCreditName_TDS3 = False
                    Exit Function
                End If
                If UCase(TDSPanOfOther_TDS3(i)) = UCase(TDSPan_TDS3(i)) Then
                    msgbox_TDS3 = msgbox_TDS3 + "* Other person PAN cannot be same as PAN of Tenant or buyer at Sr. No  " & i & "  in Schedule TDS3" & Chr(13)
                    ValidateTDSCreditName_TDS3 = False
                    Exit Function
                End If
'Change-13.05.2023.101 -> SIT-50707
'            ElseIf Left(Range("sheet1.AsseseeRepFlg").value, 1) = "Y" Then
            Else

'                If UCase(TDSPanOfOther_TDS3(i)) = UCase(Sheet1.Range("sheet1.PAN").value) Then
'                    msgbox_TDS3 = msgbox_TDS3 + "* Other person PAN cannot be same as assesse PAN at Sr. No  " & i & "  in Schedule TDS3" & Chr(13)
'                    ValidateTDSCreditName_TDS3 = False
'                    Exit Function
'                End If
             If (UCase(TDSPanOfOther_TDS3(i)) = UCase(Sheet1.Range("sheet1.PAN").value) Or UCase(TDSPanOfOther_TDS3(i)) = UCase(Sheet52.Range("sheet9.PAN").value)) Then
                    msgbox_TDS3 = msgbox_TDS3 + "* Other person PAN cannot be same as assesse PAN or verification PAN  at Sr. No  " & i & "  in Schedule TDS3" & Chr(13)
                    ValidateTDSCreditName_TDS3 = False
                    Exit Function
                End If
'---
            End If
        
        End If
        
        If Len(Aadhar_TDS3(i)) > 0 Then
        If Not ValidateAadhaar(CStr(Aadhar_TDS3(i))) Then
         msgbox_TDS3 = msgbox_TDS3 + "* Aadhaar Of Other Person at Sr. No  " & i & " is invalid in Schedule TDS3" & Chr(13)
         ValidateTDSCreditName_TDS3 = False
         Exit Function
        End If
        End If
    End If
        
        
 Next
End Function

Function ValidatePAN_TDS3() As Boolean
    ValidatePAN_TDS3 = True
    setTblinfo_TDS3
    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Sheet28.Range("TDS3.PAN").Cells
    ReDim PAN_TDS3(end_TDS3)
    Set rangecells1 = Sheet28.Range("TDS3.Aadhaar").Cells
    ReDim TDS3_Aadhaar(end_TDS3)
    subProcCaption = "Validating TDS3"
    noOfProcessSub = end_TDS3
     
    For i = 1 To end_TDS3
            PAN_TDS3(i) = UCase(rangecells.item(i).value)
 If Not Len(PAN_TDS3(i)) = 0 Then
     If Not CheckDoneePAN(CStr(PAN_TDS3(i))) Then
         msgbox_TDS3 = msgbox_TDS3 + "* PAN of Buyer/Tenant at Sr. No  " & i & " is invalid. First 5 alphabets, next 4 digits, then alphabet in Schedule TDS3" & Chr(13)
         ValidatePAN_TDS3 = False
         Exit Function
     End If
     
     
 ElseIf Not chkMandatory(PAN_TDS3(i)) Then
         msgbox_TDS3 = msgbox_TDS3 + "* Please enter the PAN of Buyer/Tenant at Sr. No  " & i & "in Schedule TDS3" & Chr(13)
     ValidatePAN_TDS3 = False
     Exit Function
 End If
 UpdateProgressBar
 
 
 TDS3_Aadhaar(i) = UCase(rangecells1.item(i).value)
 
        If Len(TDS3_Aadhaar(i)) > 0 Then
        If Not ValidateAadhaar(CStr(TDS3_Aadhaar(i))) Then
         msgbox_TDS3 = msgbox_TDS3 + "* Aadhaar Of Other Person at Sr. No  " & i & " is invalid in Schedule TDS3" & Chr(13)
         ValidatePAN_TDS3 = False
         Exit Function
        End If
        End If
 Next
End Function
Function SectionTDSisDeducted_TDS3() As Boolean
    SectionTDSisDeducted_TDS3 = True
    setTblinfo_TDS3
    
    Dim i As Long
    Dim rangecells As Range
    
    ReDim SectionTDSisDeduct_TDS3(end_TDS3)
    'Set rangecells = Sheet28.Range("TDS3.TdsIsDeducted").Cells
    Set rangecells = Sheet28.Range("TDS3.TdsIsDeducted").item(1, 1).Cells
    
    For i = 1 To end_TDS3
        SectionTDSisDeduct_TDS3(i) = rangecells.item(i).value
        
    If Not chkMandatory(SectionTDSisDeduct_TDS3(i)) And rangecells.item(i).Locked = False Then
         msgbox_TDS3 = msgbox_TDS3 + "* Please select the Section under which TDS is deducted at Sr. No  " & i & "" & Chr(13)
         SectionTDSisDeducted_TDS3 = False
         Exit Function
    End If
    
       If (SectionTDSisDeduct_TDS3(i) = "" Or SectionTDSisDeduct_TDS3(i) = "(Select)") Then
         msgbox_TDS3 = msgbox_TDS3 + "* Please select the Section under which TDS is deducted at Sr. No  " & i & "" & Chr(13)
           SectionTDSisDeducted_TDS3 = False
           Exit Function
       End If
         
    Next
End Function
Function ValidateDeductedYear_TDS3() As Boolean
 
    ValidateDeductedYear_TDS3 = True
    setTblinfo_TDS3
    Dim rangecells As Range
    Set rangecells = Sheet28.Range("TDS3.DeductedYear").Cells
    ReDim DeductedYear_TDS3(end_TDS3)
    For i = 1 To end_TDS3
        DeductedYear_TDS3(i) = rangecells.item(i).value
        If isdropdownblank(DeductedYear_TDS3(i)) Then
            DeductedYear_TDS3(i) = ""
        End If
        If Not checkfieldspecialcharacter(DeductedYear_TDS3(i)) Then
              msgbox_TDS3 = msgbox_TDS3 + "* Special characters are not allowed in Deducted Year at Sr. No  " & i & " in Schedule TDS3" & Chr(13)
              ValidateDeductedYear_TDS3 = False
              Exit Function
        End If
 Next
End Function

Function ValidateBroughtFwdTDSAmt_TDS3() As Boolean
    ValidateBroughtFwdTDSAmt_TDS3 = True
    setTblinfo_TDS3
    
    Dim i As Long
    Dim rangecells As Range
    
    
    ReDim BroughtFwdTDSAmt_TDS3(end_TDS3)
    Set rangecells = Sheet28.Range("TDS3.BroughtFwdTDSAmt").Cells
    
    For i = 1 To end_TDS3
        BroughtFwdTDSAmt_TDS3(i) = rangecells.item(i).value
        
    If Len(BroughtFwdTDSAmt_TDS3(i)) > 14 Then
        msgbox_TDS3 = msgbox_TDS3 + "* Brought Forward Amount at Sr. No  " & i & "  should not be greater than 14 digits.in Schedule TDS3" & Chr(13)
        ValidateBroughtFwdTDSAmt_TDS3 = False
    End If
         
    Next
End Function


Function ValidateDeductedInOwnHands_TDS3() As Boolean
    ValidateDeductedInOwnHands_TDS3 = True
    setTblinfo_TDS3
    
    Dim i As Long
    Dim rangecells As Range
    
    ReDim DeductedInOwnHands_TDS3(end_TDS3)
    Set rangecells = Sheet28.Range("TDS3.DeductedInOwnHands").Cells
    
    For i = 1 To end_TDS3
        DeductedInOwnHands_TDS3(i) = rangecells.item(i).value

    If Len(DeductedInOwnHands_TDS3(i)) > 14 Then
        msgbox_TDS3 = msgbox_TDS3 + "* Amount Deducted In Own Hands at Sr. No  " & i & "  should not be greater than 14 digits.in Schedule TDS3" & Chr(13)
        ValidateDeductedInOwnHands_TDS3 = False
    End If
         
    Next
End Function

Function ValidateDeductedIncome5A_TDS3() As Boolean
    ValidateDeductedIncome5A_TDS3 = True
    setTblinfo_TDS38
    
    Dim i As Long
    Dim rangecells As Range
    
    ReDim DeductedIncome5A_TDS3(end_TDS38)
    Set rangecells = Sheet28.Range("TDS3.DeductedIncome5A").Cells
    
    For i = 1 To end_TDS38
    If rangecells.item(i).Locked = True Then
        DeductedIncome5A_TDS3(i) = ""
    Else
        DeductedIncome5A_TDS3(i) = rangecells.item(i).value
    End If

    If Len(DeductedIncome5A_TDS3(i)) > 14 Then
        msgbox_TDS3 = msgbox_TDS3 + "* Deducted Income under section 5A at Sr. No  " & i & "  should not be greater than 14 digits.in Schedule TDS3" & Chr(13)
        ValidateDeductedIncome5A_TDS3 = False
    End If
         
    Next
End Function

Function ValidateDeductedTDS5A_TDS3() As Boolean
    ValidateDeductedTDS5A_TDS3 = True
    setTblinfo_TDS38
    
    Dim i As Long
    Dim rangecells As Range
    
    ReDim DeductedTDS5A_TDS3(end_TDS38)
    Set rangecells = Sheet28.Range("TDS3.DeductedTDS5A").Cells
    
    For i = 1 To end_TDS38
    If rangecells.item(i).Locked = True Then
        DeductedTDS5A_TDS3(i) = ""
    Else
        DeductedTDS5A_TDS3(i) = rangecells.item(i).value
    End If

    If Len(DeductedTDS5A_TDS3(i)) > 14 Then
        msgbox_TDS3 = msgbox_TDS3 + "* Deducted TDS under section 5A at Sr. No  " & i & "  should not be greater than 14 digits.in Schedule TDS3" & Chr(13)
        ValidateDeductedTDS5A_TDS3 = False
    End If
         
    Next
End Function

Function ValidateTDSClaimedInOwnHands_TDS3() As Boolean
    ValidateTDSClaimedInOwnHands_TDS3 = True
    setTblinfo_TDS3
    
    Dim i As Long
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    
    
    Set rangecells = Sheet28.Range("TDS3.ClaimedInOwnHands").Cells
    Set rangecells1 = Sheet28.Range("TDS3.GrossAmount").Cells
    Set rangecells2 = Sheet28.Range("TDS3.HeadOfIncome").Cells
    
    ReDim TDSClaimedInOwnHands_TDS3(end_TDS3)
    ReDim GrossAmount_TDS3(end_TDS3)
    ReDim HeadIncome_TDS3(end_TDS3)
    
    For i = 1 To end_TDS3
        TDSClaimedInOwnHands_TDS3(i) = rangecells.item(i).value
        GrossAmount_TDS3(i) = rangecells1.item(i).value
        HeadIncome_TDS3(i) = rangecells2.item(i).value
        
    If Not chkMandatory(TDSClaimedInOwnHands_TDS3(i)) And rangecells.item(i).Locked = False Then
         msgbox_TDS3 = msgbox_TDS3 + "* Please enter the TDS Claimed in Own Hands at Sr. No  " & i & " in Schedule TDS3" & Chr(13)
         ValidateTDSClaimedInOwnHands_TDS3 = False
         Exit Function
    End If
    
    If TDSClaimedInOwnHands_TDS3(i) > 0 Then
    If Not chkCompulsory(GrossAmount_TDS3(i)) Then
     msgbox_TDS3 = msgbox_TDS3 + "* Please enter the Gross Amount at Sr. No  " & i & "in Schedule TDS3" & Chr(13)
         ValidateTDSClaimedInOwnHands_TDS3 = False
         Exit Function
    End If
    
    If isdropdownblank(HeadIncome_TDS3(i)) Then
     msgbox_TDS3 = msgbox_TDS3 + "* Please enter the Head Income at Sr. No  " & i & " in Schedule TDS3" & Chr(13)
         ValidateTDSClaimedInOwnHands_TDS3 = False
         Exit Function
    End If
    End If
        
    If Len(TDSClaimedInOwnHands_TDS3(i)) > 14 Then
        msgbox_TDS3 = msgbox_TDS3 + "* TDS Claimed in Own Hands at Sr. No  " & i & "  should not be greater than 14 digits.in Schedule TDS3" & Chr(13)
        ValidateTDSClaimedInOwnHands_TDS3 = False
    End If
         
    Next
End Function

Function ValidateClaimedIncome5A_TDS3() As Boolean
    ValidateClaimedIncome5A_TDS3 = True
    setTblinfo_TDS303
    
    Dim i As Long
    Dim rangecells As Range
    
    ReDim ClaimedIncome5A_TDS3(end_TDS303)
    Set rangecells = Sheet28.Range("TDS3.ClaimedIncome5A").Cells
    
    For i = 1 To end_TDS303
        ClaimedIncome5A_TDS3(i) = rangecells.item(i).value
        
    If Len(ClaimedIncome5A_TDS3(i)) > 14 Then
        msgbox_TDS3 = msgbox_TDS3 + "* Claimed Income under section 5A at Sr. No  " & i & "  should not be greater than 14 digits.in Schedule TDS3" & Chr(13)
        ValidateClaimedIncome5A_TDS3 = False
    End If
         
    Next
End Function

Function ValidateClaimedTDS5A_TDS3() As Boolean
    ValidateClaimedTDS5A_TDS3 = True
    setTblinfo_TDS303
    
    Dim i As Long
    Dim rangecells As Range
    
    ReDim ClaimedTDS5A_TDS3(end_TDS303)
    Set rangecells = Sheet28.Range("TDS3.ClaimedTDS5A").Cells
    
    For i = 1 To end_TDS303
        ClaimedTDS5A_TDS3(i) = rangecells.item(i).value
        
    If Len(ClaimedTDS5A_TDS3(i)) > 14 Then
        msgbox_TDS3 = msgbox_TDS3 + "* Claimed TDS under section 5A at Sr. No  " & i & "  should not be greater than 14 digits.in Schedule TDS3" & Chr(13)
        ValidateClaimedTDS5A_TDS3 = False
    End If
         
    Next
End Function

Function ValidateClaimedPAN5A_TDS3() As Boolean
ValidateClaimedPAN5A_TDS3 = True
Dim rangecells As Range
Dim rangecells1 As Range
Set rangecells = Sheet28.Range("TDS3.ClaimedPAN5A").Cells
Set rangecells1 = Sheet28.Range("ClaimedAadhar5A_TDS3").Cells
Dim i As Long
ReDim ClaimedPAN5A_TDS3(end_TDS303)
ReDim ClaimedAadhar5A_TDS3(end_TDS303)
For i = 1 To end_TDS303
ClaimedPAN5A_TDS3(i) = rangecells.item(i).value
ClaimedAadhar5A_TDS3(i) = rangecells1.item(i).value
    If Not checkfieldspecialcharacter(ClaimedPAN5A_TDS3(i)) Then
         msgbox_TDS3 = msgbox_TDS3 + "* PAN at Sr. No  " & i & " in Sheet TDS  characters < > & ' " & Chr(34) & " are not allowedin Schedule TDS3" & Chr(13)
         ValidateClaimedPAN5A_TDS3 = False
         Exit Function
    End If
    
    If Not CheckPAN(UCase(ClaimedPAN5A_TDS3(i))) Then
         msgbox_TDS3 = msgbox_TDS3 + "* PAN at Sr. No  " & i & " in Sheet TDS is Invalid" & Chr(34) & "  PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet.in Schedule TDS3" & Chr(13)
         ValidateClaimedPAN5A_TDS3 = False
         Exit Function
    End If
    
     If Len(ClaimedAadhar5A_TDS3(i)) > 0 Then
        If Not ValidateAadhaar(CStr(ClaimedAadhar5A_TDS3(i))) Then
         msgbox_TDS3 = msgbox_TDS3 + "* Aadhaar Of Other Person at Sr. No  " & i & " is invalid in Schedule TDS3" & Chr(13)
         ValidateClaimedPAN5A_TDS3 = False
         Exit Function
        End If
        End If

Next
End Function

Function ValidateAmtCarriedForward_TDS3() As Boolean
    ValidateAmtCarriedForward_TDS3 = True
    setTblinfo_TDS3
    
    Dim i As Long
    Dim rangecells As Range
    
    ReDim AmtCarriedForward_TDS3(end_TDS3)
    Set rangecells = Sheet28.Range("TDS3.AmtCarriedForward").Cells
    
    For i = 1 To end_TDS3
        AmtCarriedForward_TDS3(i) = rangecells.item(i).value
        
    If Not chkMandatory(AmtCarriedForward_TDS3(i)) And rangecells.item(i).Locked = False Then
         msgbox_TDS3 = msgbox_TDS3 + "* Please enter the Amount Carried Forward at Sr. No  " & i & "in Schedule TDS3" & Chr(13)
         ValidateAmtCarriedForward_TDS3 = False
         Exit Function
    End If
        
    If Len(AmtCarriedForward_TDS3(i)) > 14 Then
        msgbox_TDS3 = msgbox_TDS3 + "* Amount Carried Forward at Sr. No  " & i & "  should not be greater than 14 digits.in Schedule TDS3" & Chr(13)
        ValidateAmtCarriedForward_TDS3 = False
    End If
         
    Next
End Function

Sub setTblinfo_TDS3()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet28.Range("TDS3.TDSCreditName").count
 Set rangecells = Sheet28.Range("TDS3.TDSCreditName").Cells
 For mIntCtr = 1 To mIntCells
     If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
    End If
 Next
 end_TDS3 = ccount
 rngname_TDS3 = "TDS3.TDSCreditName;TDS3.GrossAmount;TDS3.HeadOfIncome;TDS3.PanOtherPerson;Aadhar_TDS3;TDS3.PAN;TDS3.Aadhaar;TDS3.TdsIsDeducted;TDS3.DeductedYear;TDS3.BroughtFwdTDSAmt;TDS3.DeductedInOwnHands;TDS3.DeductedIncome5A;TDS3.DeductedTDS5A;TDS3.ClaimedInOwnHands;TDS3.ClaimedIncome5A;TDS3.ClaimedTDS5A;TDS3.ClaimedPAN5A;ClaimedAadhar5A_TDS3;TDS3.AmtCarriedForward;"
 End Sub
 
 Sub setTblinfo_TDS3_Pan()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet28.Range("TDS3.PanOtherPerson").count
 Set rangecells = Sheet28.Range("TDS3.PanOtherPerson").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS3_Pan = ccount
 End Sub
 
 Sub setTblinfo_TDS3_GrossAmount()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet28.Range("TDS3.GrossAmount").count
 Set rangecells = Sheet28.Range("TDS3.GrossAmount").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS3_GrossAmount = ccount
 End Sub
 
 Sub setTblinfo_TDS3_Head()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet28.Range("TDS3.HeadOfIncome").count
 Set rangecells = Sheet28.Range("TDS3.HeadOfIncome").Cells
 For mIntCtr = 1 To mIntCells
     If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
    End If
 Next
 end_TDS3_Head = ccount
 End Sub
 
 Sub setTblinfo_TDS32()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet28.Range("TDS3.PAN").count
 Set rangecells = Sheet28.Range("TDS3.PAN").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS32 = ccount
 End Sub
  Sub setTblinfo_TDS3_Section()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet28.Range("TDS3.SectionTDS").count
 Set rangecells = Sheet28.Range("TDS3.SectionTDS").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TDS3_Pan = ccount
 End Sub

 Sub setTblinfo_TDS35()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 Dim Year_TDS As Variant
 ccount = 0
 mIntCells = Sheet28.Range("TDS3.DeductedYear").count
 ReDim Year_TDS(mIntCells)
 Set rangecells = Sheet28.Range("TDS3.DeductedYear").Cells
 For mIntCtr = 1 To mIntCells
    Year_TDS(mIntCtr) = rangecells.item(mIntCtr).value
    If isdropdownblank(Year_TDS(mIntCtr)) Then
        Year_TDS(mIntCtr) = ""
    End If
    If Not Year_TDS(mIntCtr) = "" Then
        ccount = ccount + 1
    End If
 Next
    If ccount <= end_TDS3 Then
    end_TDS35 = end_TDS3
    Else
    end_TDS35 = ccount
    End If
 End Sub
 
 Sub setTblinfo_TDS36()
 Dim rangecells As Range
 Dim mIntCells  As Long
 Dim mIntCtr  As Long
 Dim ccount  As Long
 ccount = 0
 mIntCells = Sheet28.Range("TDS3.BroughtFwdTDSAmt").count
 Set rangecells = Sheet28.Range("TDS3.BroughtFwdTDSAmt").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 If ccount <= end_TDS3 Then
    end_TDS36 = end_TDS3
Else
 end_TDS36 = ccount
End If
 End Sub
 
Sub setTblinfo_TDS37()
Dim rangecells As Range
Dim mIntCells  As Long
Dim mIntCtr  As Long
Dim ccount  As Long
ccount = 0
mIntCells = Sheet28.Range("TDS3.DeductedInOwnHands").count
Set rangecells = Sheet28.Range("TDS3.DeductedInOwnHands").Cells
For mIntCtr = 1 To mIntCells
    If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
    End If
Next
If ccount <= end_TDS3 Then
    end_TDS37 = end_TDS3
    Else
    end_TDS37 = ccount
    End If
End Sub

Sub setTblinfo_TDS38()
Dim rangecells As Range
Dim mIntCells  As Long
Dim mIntCtr  As Long
Dim ccount  As Long
ccount = 0
mIntCells = Sheet28.Range("TDS3.DeductedIncome5A").count
Set rangecells = Sheet28.Range("TDS3.DeductedIncome5A").Cells
For mIntCtr = 1 To mIntCells
    If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
    End If
Next
If ccount <= end_TDS3 Then
    end_TDS38 = end_TDS3
    Else
    end_TDS38 = ccount
    End If
End Sub

Sub setTblinfo_TDS39()
Dim rangecells As Range
Dim mIntCells  As Long
Dim mIntCtr  As Long
Dim ccount  As Long
ccount = 0
mIntCells = Sheet28.Range("TDS3.DeductedTDS5A").count
Set rangecells = Sheet28.Range("TDS3.DeductedTDS5A").Cells
For mIntCtr = 1 To mIntCells
    If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
    End If
Next
If ccount <= end_TDS3 Then
    end_TDS39 = end_TDS3
    Else
    end_TDS39 = ccount
    End If
End Sub

Sub setTblinfo_TDS302()
Dim rangecells As Range
Dim mIntCells  As Long
Dim mIntCtr  As Long
Dim ccount  As Long
ccount = 0
mIntCells = Sheet28.Range("TDS3.ClaimedInOwnHands").count
Set rangecells = Sheet28.Range("TDS3.ClaimedInOwnHands").Cells
For mIntCtr = 1 To mIntCells
    If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
    End If
Next
end_TDS302 = ccount
End Sub

Sub setTblinfo_TDS303()
Dim rangecells As Range
Dim mIntCells  As Long
Dim mIntCtr  As Long
Dim ccount  As Long
ccount = 0
mIntCells = Sheet28.Range("TDS3.ClaimedIncome5A").count
Set rangecells = Sheet28.Range("TDS3.ClaimedIncome5A").Cells
For mIntCtr = 1 To mIntCells
    If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
    End If
Next
If ccount <= end_TDS3 Then
    end_TDS303 = end_TDS3
    Else
    end_TDS303 = ccount
    End If
End Sub

Sub setTblinfo_TDS304()
Dim rangecells As Range
Dim mIntCells  As Long
Dim mIntCtr  As Long
Dim ccount  As Long
ccount = 0
mIntCells = Sheet28.Range("TDS3.ClaimedTDS5A").count
Set rangecells = Sheet28.Range("TDS3.ClaimedTDS5A").Cells
For mIntCtr = 1 To mIntCells
    If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
    End If
Next
If ccount <= end_TDS3 Then
    end_TDS304 = end_TDS3
    Else
    end_TDS304 = ccount
    End If
End Sub

Sub setTblinfo_TDS305()
Dim rangecells As Range
Dim mIntCells  As Long
Dim mIntCtr  As Long
Dim ccount  As Long
ccount = 0
mIntCells = Sheet28.Range("TDS3.ClaimedPAN5A").count
Set rangecells = Sheet28.Range("TDS3.ClaimedPAN5A").Cells
For mIntCtr = 1 To mIntCells
    If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
    End If
Next
If ccount <= end_TDS3 Then
    end_TDS305 = end_TDS3
    Else
    end_TDS305 = ccount
    End If
End Sub

Sub setTblinfo_TDS306()
Dim rangecells As Range
Dim mIntCells  As Long
Dim mIntCtr  As Long
Dim ccount  As Long
ccount = 0
mIntCells = Sheet28.Range("TDS3.AmtCarriedForward").count
Set rangecells = Sheet28.Range("TDS3.AmtCarriedForward").Cells
For mIntCtr = 1 To mIntCells
    If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
    End If
Next
end_TDS306 = ccount
End Sub



''TCS
'Function ValidatesheetTCS() As Boolean
'    Dim i As Variant
'
'    ValidatesheetTCS = True
'    setTblinfo_TCS
'    setTblinfo_TCS2
'    setTblinfo_TCS4
'    setTblinfo_TCS5
'   ' setTblinfo_TCS6
'
'    If (end_TCS <> end_TCS2) Or (end_TCS <> end_TCS5) Then
'        msgbox_TCS = msgbox_TCS + "* Enter all Mandatory Details in Schedule TCS." & Chr(13)
'        ValidatesheetTCS = False
'        Exit Function
'    End If
'    If Not ValidateTAN_TCS() Then ValidatesheetTCS = False
'    If (Len(Sheet28.Range("TCS.TAN").item(1).value) > 0) Then
'        If Not ValidateEmployerOrDeductorOrCollecterName_TCS() Then ValidatesheetTCS = False
'        If Not ValidateBroughtFwdTDSAmt() Then ValidatesheetTCS = False
'        If Not ValidateDeductedYear_TCS() Then ValidatesheetTCS = False
'        If Not ValidateTotalTCS_TCS() Then ValidatesheetTCS = False
'
'        If Not ValidateAmtTCSClaimedThisYear_TCS() Then ValidatesheetTCS = False
'    End If
'
'
'    For i = 1 To end_TCS
'    If BroughtFwdTDSAmt_TCS(i) > 0 Then
'       If TotalTCS_TCS(i) > 0 Then
'       msgbox_TCS = msgbox_TCS + "* Unclaimed TCS brought forward and details of TCS of current FY should be provided in different rows in Schedule TCS" & Chr(13)
'        ValidatesheetTCS = False
'        Exit Function
'    End If
'    End If
'    Next
'
'
'    Dim Total_TCS As Variant
'    Total_TCS = 0
'    Total_TCS = Sheet28.Range("TCS.Sum").value
'
'    If Len(Total_TCS) > 14 Then
'        msgbox_TCS = msgbox_TCS + "* Total of TCS cannot be greater than 14 digits in Schedule TCS. " & Chr(13)
'        ValidatesheetTCS = False
'        Exit Function
'    End If
'
'
'End Function

'Function ValidateTAN_TCS() As Boolean
'    Dim rangecells As Range
'    Dim i As Long
'
'    ValidateTAN_TCS = True
'    setTblinfo_TCS
'    Set rangecells = Sheet28.Range("TCS.TAN").Cells
'    ReDim TAN_TCS(end_TCS)
'    For i = 1 To end_TCS
'    noOfProcessSub = end_TCS
'        TAN_TCS(i) = UCase(rangecells.item(i).value)
'        If Not Len(TAN_TCS(i)) = 0 Then
'            If Not ValidateTantype_text(Mid(UCase(TAN_TCS(i)), 1, 4)) Then
'                msgbox_TCS = msgbox_TCS + "* TAN at Sr. No  " & i & " in Schedule TCS  is invalid. First 4 alphabets, next 5 digits, then alphabet.in Schedule TCS" & Chr(13)
'                ValidateTAN_TCS = False
'                Exit Function
'            End If
'
'            If Not IsNumeric(Mid(UCase(TAN_TCS(i)), 5, 5)) Then
'                msgbox_TCS = msgbox_TCS + "* TAN at Sr. No  " & i & "  in Schedule TCS  is invalid. First 4 alphabets, next 5 digits, then alphabet.in Schedule TCS" & Chr(13)
'                ValidateTAN_TCS = False
'                Exit Function
'            End If
'
'            If Not ValidateTantype_text(Right(UCase(TAN_TCS(i)), 1)) Then
'                msgbox_TCS = msgbox_TCS + "* TAN at Sr. No  " & i & "  in Schedule TCS  is invalid. First 4 alphabets, next 5 digits, then alphabet.in Schedule TCS" & Chr(13)
'                ValidateTAN_TCS = False
'                Exit Function
'            End If
'
'            If Not ValidateTANCodes(UCase(TAN_TCS(i))) Then
'                msgbox_TCS = msgbox_TCS + "* Invalid TAN at Sr.No " & i & "  Schedule TCS.Please enter valid TAN" & Chr(13)
'                ValidateTAN_TCS = False
'                Exit Function
'        End If
'
'        ElseIf Not chkCompulsory(TAN_TCS(i)) Then
'            msgbox_TCS = msgbox_TCS + "* Please enter the TAN at Sr. No  " & i & "  in Schedule TCS." & Chr(13)
'            ValidateTAN_TCS = False
'            Exit Function
'        End If
'        UpdateProgressBar
'    Next
'End Function

'Function ValidateEmployerOrDeductorOrCollecterName_TCS() As Boolean
'    Dim rangecells As Range
'    Dim i As Long
'
'    ValidateEmployerOrDeductorOrCollecterName_TCS = True
'    setTblinfo_TCS
'
'    Set rangecells = Sheet28.Range("TCS.EmployerOrDeductorOrCollecterName").Cells
'    ReDim EmployerOrDeductorOrCollecterName_TCS(end_TCS)
'    For i = 1 To end_TCS
'        EmployerOrDeductorOrCollecterName_TCS(i) = rangecells.item(i).value
'
'        If Not chkCompulsory(EmployerOrDeductorOrCollecterName_TCS(i)) Then
'           msgbox_TCS = msgbox_TCS + "* Please enter the Name of collector at Sr. No  " & i & " in Schedule TCS" & Chr(13)
'            ValidateEmployerOrDeductorOrCollecterName_TCS = False
'            Exit Function
'        End If
'
'    Next
'End Function
'Function ValidateDeductedYear_TCS() As Boolean
'    Dim rangecells As Range
'    Dim i As Long
'
'    ValidateDeductedYear_TCS = True
'    setTblinfo_TCS
'    Set rangecells = Sheet28.Range("TCS.DeductedYear").Cells
'    ReDim DeductedYear_TCS(end_TCS)
'    For i = 1 To end_TCS
'        DeductedYear_TCS(i) = rangecells.item(i).value
'
'        If isdropdownblank(DeductedYear_TCS(i)) Then
'            DeductedYear_TCS(i) = ""
'        End If
'
'        If DeductedYear_TCS(i) = "" And BroughtFwdTDSAmt_TCS(i) <> "" Then
'            msgbox_TCS = msgbox_TCS + "* Please select the  financial year in which collected  at Sr. No  " & i & " in Schedule TCS" & Chr(13)
'            ValidateDeductedYear_TCS = False
'            Exit Function
'        End If
'
'        If DeductedYear_TCS(i) <> "" And BroughtFwdTDSAmt_TCS(i) = "" Then
'            msgbox_TCS = msgbox_TCS + "* Please enter the Amount of TCS brought forward  at Sr. No  " & i & " in Schedule TCS." & Chr(13)
'            ValidateDeductedYear_TCS = False
'            Exit Function
'        End If
'
'        If Not checkfieldspecialcharacter(DeductedYear_TCS(i)) Then
'            msgbox_TCS = msgbox_TCS + "* Special characters are not allowed in TCS. Fin. Year at Sr. No  " & i & " in Schedule TCS." & Chr(13)
'            ValidateDeductedYear_TCS = False
'            Exit Function
'        End If
'    Next
'End Function

'Function ValidateBroughtFwdTDSAmt() As Boolean
'    Dim rangecells As Range
'
'    Dim i As Long
'    Dim chkcomp As Boolean
'
'    ValidateBroughtFwdTDSAmt = True
'    setTblinfo_TCS
'    Set rangecells = Sheet28.Range("TCS.BroughtFwdTDSAmt").Cells
'
'    ReDim BroughtFwdTDSAmt_TCS(end_TCS)
'
'    For i = 1 To end_TCS
'        BroughtFwdTDSAmt_TCS(i) = rangecells.item(i).value
'
'        If Not checkfieldspecialcharacter(BroughtFwdTDSAmt_TCS(i)) Then
'            msgbox_TCS = msgbox_TCS + "* Special characters are not allowed in Brought Forward TCS of current financial Year at Sr. No  " & i & " in Schedule TCS." & Chr(13)
'            ValidateBroughtFwdTDSAmt = False
'            Exit Function
'        End If
' Next
'End Function


'Function ValidateTotalTCS_TCS() As Boolean
'    Dim rangecells As Range
'
'    Dim i As Long
'    Dim chkcomp As Boolean
'
'    ValidateTotalTCS_TCS = True
'    setTblinfo_TCS
'    Set rangecells = Sheet28.Range("TCS.TotalTCS").Cells
'
'    ReDim TotalTCS_TCS(end_TCS)
'
'    For i = 1 To end_TCS
'        TotalTCS_TCS(i) = rangecells.item(i).value
'
'        If Not checkfieldspecialcharacter(TotalTCS_TCS(i)) Then
'            msgbox_TCS = msgbox_TCS + "* Special characters are not allowed in TCS of current financial Year at Sr. No  " & i & " in Schedule TCS." & Chr(13)
'            ValidateTotalTCS_TCS = False
'            Exit Function
'        End If
'
'        If DeductedYear_TCS(i) = "" And BroughtFwdTDSAmt_TCS(i) = "" And TotalTCS_TCS(i) = "" Then
'            msgbox_TCS = msgbox_TCS + "* Please enter the Amount in TCS of current financial Year or Brought Forward TCS of current financial Year at Sr. No  " & i & " in Schedule TCS." & Chr(13)
'            ValidateTotalTCS_TCS = False
'            Exit Function
'        End If
'
' Next
'End Function

'Function ValidateAmtTCSClaimedThisYear_TCS() As Boolean
'    Dim rangecells, rangecells1 As Range
'
'    Dim i As Long
'
'    ValidateAmtTCSClaimedThisYear_TCS = True
'    setTblinfo_TCS
'
'    Set rangecells = Sheet28.Range("TCS.AmtTCSClaimedThisYear").Cells
'    Set rangecells1 = Sheet28.Range("TCS.AmtCarriedFwd").Cells
'    ReDim AmtTCSClaimedThisYear_TCS(end_TCS)
'    ReDim AmtCarriedFwd_TCS(end_TCS)
'    For i = 1 To end_TCS
'        AmtTCSClaimedThisYear_TCS(i) = rangecells.item(i).value
'        AmtCarriedFwd_TCS(i) = rangecells1.item(i).value
'        If Not chkCompulsory(AmtTCSClaimedThisYear_TCS(i)) Then
'            msgbox_TCS = msgbox_TCS + "* Amount claimed this year(7) at Sr. No  " & i & " in Schedule TCS is compulsary. Please fill zero if no amount is claimed.in Schedule TCS" & Chr(13)
'            ValidateAmtTCSClaimedThisYear_TCS = False
'            Exit Function
'        End If
'
'        If (AmtTCSClaimedThisYear_TCS(i)) > TotalTCS_TCS(i) + BroughtFwdTDSAmt_TCS(i) Then
'           msgbox_TCS = msgbox_TCS + "* Amount in field Col(7) cannot be more than Col (5) or Col (6) as the case may be at Sr. No  " & i & " in Schedule TCS. " & Chr(13)
'            ValidateAmtTCSClaimedThisYear_TCS = False
'            Exit Function
'        End If
'    Next
'End Function

'Sub setTblinfo_TCS()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'
'    ccount = 0
'    mIntCells = Sheet28.Range("TCS.TAN").count
'    Set rangecells = Sheet28.Range("TCS.TAN").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'    end_TCS = ccount
'    End Sub

'Sub setTblinfo_TCS2()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'
'    ccount = 0
'    mIntCells = Sheet28.Range("TCS.EmployerOrDeductorOrCollecterName").count
'    Set rangecells = Sheet28.Range("TCS.EmployerOrDeductorOrCollecterName").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'    end_TCS2 = ccount
'End Sub
'Sub setTblinfo_TCS3()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    Dim Year_TDS As Variant
'    ccount = 0
'    mIntCells = Sheet28.Range("TCS.DeductedYear").count
'    ReDim Year_TDS(mIntCells)
'    Set rangecells = Sheet28.Range("TCS.DeductedYear").Cells
'    For mIntCtr = 1 To mIntCells
'
'        Year_TDS(mIntCtr) = rangecells.item(mIntCtr).value
'        If isdropdownblank(Year_TDS(mIntCtr)) Then
'            Year_TDS(mIntCtr) = ""
'        End If
'
'        If Not Year_TDS(mIntCtr) = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'    end_TCS3 = ccount
'End Sub
'Sub setTblinfo_TCS4()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'
'    ccount = 0
'    mIntCells = Sheet28.Range("TCS.TotalTCS").count
'    Set rangecells = Sheet28.Range("TCS.TotalTCS").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'    end_TCS4 = ccount
'End Sub
'Sub setTblinfo_TCS5()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'
'    ccount = 0
'    mIntCells = Sheet28.Range("TCS.AmtTCSClaimedThisYear").count
'    Set rangecells = Sheet28.Range("TCS.AmtTCSClaimedThisYear").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'    end_TCS5 = ccount
'End Sub
'
'Sub setTblinfo_TCS6()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'
'    ccount = 0
'    mIntCells = Sheet28.Range("TCS.BroughtFwdTDSAmt").count
'    Set rangecells = Sheet28.Range("TCS.BroughtFwdTDSAmt").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            ccount = ccount + 1
'        End If
'    Next
'    end_TCS6 = ccount
'End Sub

Function ValidateTantype_text(strName As Variant) As Boolean
    Dim len1 As Long
    Dim s1 As String
    Dim i As Long
    
    ValidateTantype_text = True
    len1 = Len(strName)
    For i = 1 To len1
        s1 = Mid(strName, i, 1)
        If (((asc(s1) >= 65) And (asc(s1) <= 90)) Or (asc(s1) = 45)) Then
        Else
            ValidateTantype_text = False
        End If
    Next
End Function

Function GreaterValue(Number1 As Variant, Number2 As Variant) As Double
If (Number1 = "" Or IsEmpty(Number1)) Then
Number1 = 0
End If

If (Number2 = "" Or IsEmpty(Number2)) Then
Number2 = 0
End If

    If (Number1 > Number2) Then
        GreaterValue = Number1
    Else
    
        GreaterValue = Number2
        
    End If
End Function

Function ValidateTANCodes(TAN As String) As Boolean
    ValidateTANCodes = False
    Dim TANCode As Range

    
    Dim TAN1 As Range

    Set TANCode = SheetALL.Range("TANCode")
  
    
    For Each TAN1 In TANCode
        If Mid(TAN1.value, 1, 3) = Mid(TAN, 1, 3) Then
        ValidateTANCodes = True
        Exit For
        End If
    Next

End Function

'Added by Riyaz on 16/06/2025 for 2025-26 New Dev
Sub ImportScheduleTDS1CSV()

Dim filePath, text As String

Dim answer As Integer
 
answer = MsgBox("Do you want to import CSV? If you import csv, All data will be clear.", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
If answer = vbYes Then
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Please select a CSV file."
        .Filters.add "CSV File", "*.csv", 1
        If .Show = True Then
            filePath = .SelectedItems.item(1)
        Else
            Exit Sub
        End If
        On Error Resume Next
    End With
    
    Open filePath For Input As #1
    text = Input$(LOF(1), 1)
    Close #1
    
    importTDS1CSVFunc (text)
End If
End Sub
'Added by Riyaz on 16/06/2025 for 2025-26 New Dev
Sub importTDS1CSVFunc(text As String)
    Const Delim As String = ","
    Dim MyData As String, strData() As String, TmpAr() As String
    Dim TwoDArray() As String
    Dim i As Long, n As Long
    Dim TotalExRow, TotalXMLRow, TotalDiffRow, rowcount As Variant
    strData() = Split(text, vbCrLf)
    For i = LBound(strData) To UBound(strData)
        If Len(Trim(strData(i))) <> 0 And i <> 0 Then
            TmpAr = Split(strData(i), Delim)
            n = n + 1
            ReDim Preserve TwoDArray(3, 1 To n)
        
        
            TwoDArray(0, n) = TmpAr(0)
            TwoDArray(1, n) = TmpAr(1)
            TwoDArray(2, n) = TmpAr(2)
            TwoDArray(3, n) = TmpAr(3)
            
            
        End If
    Next i
   
    Dim SelfOthColNo, TANColNo, PANColNo, DeductedYrColNo
    'Dim CurrYrClaimOwnHandColNo, CurrYrClaimOthHandTCSColNo, CurrYrClaimOthHandPANColNo, AmtCarriedFwdColNo
    
    SelfOthColNo = Sheet28.Range("TDS1.TAN").Column
    TANColNo = Sheet28.Range("TDS1.EmployerOrDeductorOrCollecterName").Column
    PANColNo = Sheet28.Range("TDS1.IncChrgSal").Column
    DeductedYrColNo = Sheet28.Range("TDS1.TotalTDSSal").Column
      
    
    TotalExRow = Sheet28.Range("TDS1.TAN").Rows.count
    TotalXMLRow = n
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    
        If Sheet28.Range("TDS1.TAN").Locked = False Then
            Sheet28.Range("TDS1.TAN").ClearContents
        End If
        If Sheet28.Range("TDS1.EmployerOrDeductorOrCollecterName").Locked = False Then
            Sheet28.Range("TDS1.EmployerOrDeductorOrCollecterName").ClearContents
        End If
        If Sheet28.Range("TDS1.IncChrgSal").Locked = False Then
            Sheet28.Range("TDS1.IncChrgSal").ClearContents
        End If
        If Sheet28.Range("TDS1.TotalTDSSal").Locked = False Then
            Sheet28.Range("TDS1.TotalTDSSal").ClearContents
        End If
                
    End If
    
    If (TotalDiffRow > 0) Then
        AddDiffRows_TDS1 (TotalDiffRow)
    End If
    rowcount = getRowNo(Sheet28.Range("TDS1.TAN").name)
    
    For i = 1 To n
        
        If Sheet28.Cells(rowcount, SelfOthColNo).Locked = False Then
            Sheet28.Cells(rowcount, SelfOthColNo).value = TwoDArray(0, i)
        End If
        
        If Sheet28.Cells(rowcount, TANColNo).Locked = False Then
            Sheet28.Cells(rowcount, TANColNo).value = TwoDArray(1, i)
        End If
        
        If Sheet28.Cells(rowcount, PANColNo).Locked = False Then
            Sheet28.Cells(rowcount, PANColNo).value = TwoDArray(2, i)
        End If
        
        If Sheet28.Cells(rowcount, DeductedYrColNo).Locked = False Then
            Sheet28.Cells(rowcount, DeductedYrColNo).value = TwoDArray(3, i)
        End If
                
        rowcount = rowcount + 1
                
    Next i
    
End Sub

Sub ImportScheduleTDS2CSV()
Dim filePath, text As String

Dim answer As Integer
 
answer = MsgBox("Do you want to import CSV? If you import csv, All data will be clear.", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
If answer = vbYes Then
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Please select a CSV file."
        .Filters.add "CSV File", "*.csv", 1
        If .Show = True Then
            filePath = .SelectedItems.item(1)
        Else
            Exit Sub
        End If
        On Error Resume Next
    End With
    
    Open filePath For Input As #1
    text = Input$(LOF(1), 1)
    Close #1
    
    importTDS2CSVFunc (text)
End If
End Sub
Sub importTDS2CSVFunc(text As String)
    Const Delim As String = ","
    Dim MyData As String, strData() As String, TmpAr() As String
    Dim TwoDArray() As String
    Dim i As Long, n As Long
    Dim TotalExRow, TotalXMLRow, TotalDiffRow As Variant
    strData() = Split(text, vbCrLf)
    For i = LBound(strData) To UBound(strData)
        If Len(Trim(strData(i))) <> 0 And i <> 0 Then
            TmpAr = Split(strData(i), Delim)
            n = n + 1
            ReDim Preserve TwoDArray(16, 1 To n)
                    
            TwoDArray(0, n) = TmpAr(0)      'slno
            TwoDArray(1, n) = TmpAr(1)      'self,other
            TwoDArray(2, n) = TmpAr(2)      'pan of person
            TwoDArray(3, n) = TmpAr(3)      'Aadhaar of other person
            TwoDArray(4, n) = TmpAr(4)      'tan of deductor
            TwoDArray(5, n) = TmpAr(5)      'year
            TwoDArray(6, n) = TmpAr(6)      'tds b/f
            TwoDArray(7, n) = TmpAr(7)      'deducted in own hands
            TwoDArray(8, n) = TmpAr(8)      'income spouse
            TwoDArray(9, n) = TmpAr(9)      'TDS spouse
            TwoDArray(10, n) = TmpAr(10)    'Claimed in own hands
            TwoDArray(11, n) = TmpAr(11)    'income
            TwoDArray(12, n) = TmpAr(12)    'TDS
            TwoDArray(13, n) = TmpAr(13)    'PAN
            TwoDArray(14, n) = TmpAr(14)    'Aadhar
            TwoDArray(15, n) = TmpAr(15)    'Gross amount
            TwoDArray(16, n) = TmpAr(16)    'Head of income
            
        End If
    Next i
    
    Dim TDSCreditName_TDS2, HeadIncome_TDS2, rowcount
    Dim TDSCreditNameColNo, PANofOtherPersonColNo, AadhaarOfOtherPersonColNo, TANOfDeductorColNo, DeductedYrColNo, BroughtFwdTDSAmtColNo
    Dim TaxDeductedOwnHandsColNo, SectionOfDeductorColNo
    Dim TaxDeductedIncomeColNo, TaxDeductedTDSColNo, TaxClaimedOwnHandsColNo, TaxClaimedIncomeColNo, TaxClaimedTDSColNo
    Dim TaxClaimedSpouseOthPrsnPANColNo, SpouseOthPrsnAadhaarColNo
    Dim GrossAmountColNo, HeadOfIncomeColNo, AmtCarriedFwdColNo
    
    TDSCreditNameColNo = Sheet28.Range("TDS2.TDSCreditName").Column
    PANofOtherPersonColNo = Sheet28.Range("TDS2.PanOtherPerson").Column
    AadhaarOfOtherPersonColNo = Sheet28.Range("TDSAadharOfOther_TDS2").Column
    TANOfDeductorColNo = Sheet28.Range("TDS2.TAN").Column
    SectionOfDeductorColNo = Sheet28.Range("TDS2.TdsDeducted").Column 'new
    DeductedYrColNo = Sheet28.Range("TDS2.DeductedYear").Column
    BroughtFwdTDSAmtColNo = Sheet28.Range("TDS2.BroughtFwdTDSAmt").Column
    TaxDeductedOwnHandsColNo = Sheet28.Range("TDS2.DeductedInOwnHands").Column
    TaxDeductedIncomeColNo = Sheet28.Range("TDS2.DeductedIncome5A").Column
    TaxDeductedTDSColNo = Sheet28.Range("TDS2.DeductedTDS5A").Column
    TaxClaimedOwnHandsColNo = Sheet28.Range("TDS2.ClaimedInOwnHands").Column
    TaxClaimedIncomeColNo = Sheet28.Range("TDS2.ClaimedIncome5A").Column
    TaxClaimedTDSColNo = Sheet28.Range("TDS2.ClaimedTDS5A").Column
    TaxClaimedSpouseOthPrsnPANColNo = Sheet28.Range("TDS2.ClaimedPAN5A").Column
    SpouseOthPrsnAadhaarColNo = Sheet28.Range("ClaimedAadhar5A_TDS2").Column
    GrossAmountColNo = Sheet28.Range("TDS2.GrossAmount").Column
    HeadOfIncomeColNo = Sheet28.Range("TDS2.HeadOfIncome").Column
    
    
    TotalExRow = Sheet28.Range("TDS2.TDSCreditName").Rows.count
    TotalXMLRow = n
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Application.EnableEvents = False
        If Sheet28.Range("TDS2.TDSCreditName").Locked = False Then
            Sheet28.Range("TDS2.TDSCreditName").ClearContents
        End If
        If Sheet28.Range("TDS2.PanOtherPerson").Locked = False Then
            Sheet28.Range("TDS2.PanOtherPerson").ClearContents
        End If
        If Sheet28.Range("TDSAadharOfOther_TDS2").Locked = False Then
            Sheet28.Range("TDSAadharOfOther_TDS2").ClearContents
        End If
        If Sheet28.Range("TDS2.TAN").Locked = False Then
            Sheet28.Range("TDS2.TAN").ClearContents
        End If
        If Sheet28.Range("TDS2.TdsDeducted").Locked = False Then
            Sheet28.Range("TDS2.TdsDeducted").ClearContents
        End If
        If Sheet28.Range("TDS2.DeductedYear").Locked = False Then
            Sheet28.Range("TDS2.DeductedYear").ClearContents
        End If
        If Sheet28.Range("TDS2.DeductedIncome5A").Locked = False Then
            Sheet28.Range("TDS2.DeductedIncome5A").ClearContents
        End If
        If Sheet28.Range("TDS2.DeductedTDS5A").Locked = False Then
            Sheet28.Range("TDS2.DeductedTDS5A").ClearContents
        End If
        If Sheet28.Range("TDS2.ClaimedInOwnHands").Locked = False Then
            Sheet28.Range("TDS2.ClaimedInOwnHands").ClearContents
        End If
        If Sheet28.Range("TDS2.ClaimedIncome5A").Locked = False Then
            Sheet28.Range("TDS2.ClaimedIncome5A").ClearContents
        End If
        If Sheet28.Range("TDS2.ClaimedTDS5A").Locked = False Then
            Sheet28.Range("TDS2.ClaimedTDS5A").ClearContents
        End If
        If Sheet28.Range("TDS2.ClaimedPAN5A").Locked = False Then
            Sheet28.Range("TDS2.ClaimedPAN5A").ClearContents
        End If
        If Sheet28.Range("ClaimedAadhar5A_TDS2").Locked = False Then
            Sheet28.Range("ClaimedAadhar5A_TDS2").ClearContents
        End If
        If Sheet28.Range("TDS2.GrossAmount").Locked = False Then
            Sheet28.Range("TDS2.GrossAmount").ClearContents
        End If
        If Sheet28.Range("TDS2.HeadOfIncome").Locked = False Then
            Sheet28.Range("TDS2.HeadOfIncome").ClearContents
        End If
        Application.EnableEvents = True
    End If
    
    If (TotalDiffRow > 0) Then
        AddDiffRows_TDSoth (TotalDiffRow)
    End If
    rowcount = getRowNo(Sheet28.Range("TDS2.TDSCreditName").name)

    For i = 1 To n
        If Sheet28.Cells(rowcount, TDSCreditNameColNo).Locked = False Then
            If UCase(Mid(TwoDArray(0, i), 1, 1)) = "S" Then
                TwoDArray(0, i) = "Self"
            End If
            If UCase(Mid(TwoDArray(0, i), 1, 1)) = "O" Then
                TwoDArray(0, i) = "Other Person"
            End If
            Sheet28.Cells(rowcount, TDSCreditNameColNo).value = TwoDArray(0, i)
        End If
        
        If Sheet28.Cells(rowcount, PANofOtherPersonColNo).Locked = False Then
            Sheet28.Cells(rowcount, PANofOtherPersonColNo).value = TwoDArray(1, i)
        End If
        
        If Sheet28.Cells(rowcount, AadhaarOfOtherPersonColNo).Locked = False Then
            Sheet28.Cells(rowcount, AadhaarOfOtherPersonColNo).value = TwoDArray(2, i)
        End If
        
        If Sheet28.Cells(rowcount, TANOfDeductorColNo).Locked = False Then
            Sheet28.Cells(rowcount, TANOfDeductorColNo).value = TwoDArray(3, i)
        End If
        
'New section
        If Sheet28.Cells(rowcount, SectionOfDeductorColNo).Locked = False Then
            If TwoDArray(4, i) = "192-Salary-Payment to Government employees other than Indian Government employees" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "192-Salary-Payment to Government employees other than Indian Government employees"
            ElseIf TwoDArray(4, i) = "192-Salary-Payment to employees other than Government employees" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "192-Salary-Payment to employees other than Government employees"
            ElseIf TwoDArray(4, i) = "192-Salary-Payment to Indian Government employees" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "192-Salary-Payment to Indian Government employees"
            ElseIf TwoDArray(4, i) = "192A-TDS on PF withdrawal" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "192A-TDS on PF withdrawal"
            ElseIf TwoDArray(4, i) = "193-Interest on Securities" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "193-Interest on Securities"
            ElseIf TwoDArray(4, i) = "194-Dividends" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194-Dividends"
            ElseIf TwoDArray(4, i) = "194A-Interest other than 'Interest on securities'" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194A-Interest other than 'Interest on securities'"
            ElseIf TwoDArray(4, i) = "194B-Winning from lottery or crossword puzzle" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194B-Winning from lottery or crossword puzzle"
            ElseIf TwoDArray(4, i) = "194BA-Winnings from online games" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194BA-Winnings from online games"
            ElseIf TwoDArray(4, i) = "194BB-Winning from horse race" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194BB-Winning from horse race"
            ElseIf TwoDArray(4, i) = "194C-Payments to contractors and sub-contractors" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194C-Payments to contractors and sub-contractors"
            ElseIf TwoDArray(4, i) = "194D-Insurance commission" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194D-Insurance commission"
            ElseIf TwoDArray(4, i) = "194DA-Payment in respect of life insurance policy" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194DA-Payment in respect of life insurance policy"
            ElseIf TwoDArray(4, i) = "194E-Payments to non-resident sportsmen or sports associations" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194E-Payments to non-resident sportsmen or sports associations"
            ElseIf TwoDArray(4, i) = "194EE-Payments in respect of deposits under National Savings" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194EE-Payments in respect of deposits under National Savings"
            ElseIf TwoDArray(4, i) = "194F-Payments on account of repurchase of units by Mutual Fund or Unit Trust of India" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194F-Payments on account of repurchase of units by Mutual Fund or Unit Trust of India"
            ElseIf TwoDArray(4, i) = "194G-Commission, price, etc. on sale of lottery tickets" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194G-Commission, price, etc. on sale of lottery tickets"
            ElseIf TwoDArray(4, i) = "194H-Commission or brokerage" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194H-Commission or brokerage"
            ElseIf TwoDArray(4, i) = "194I(a)-Rent on hiring of plant and machinery" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194I(a)-Rent on hiring of plant and machinery"
            ElseIf TwoDArray(4, i) = "194I(b)-Rent on other than plant and machinery" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194I(b)-Rent on other than plant and machinery"
            ElseIf TwoDArray(4, i) = "194IA-TDS on Sale of immovable property" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194IA-TDS on Sale of immovable property"
            ElseIf TwoDArray(4, i) = "194IB-Payment of rent by certain individuals or Hindu undivided" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194IB-Payment of rent by certain individuals or Hindu undivided"
            ElseIf TwoDArray(4, i) = "194IC-Payment under specified agreement" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194IC-Payment under specified agreement"
            ElseIf TwoDArray(4, i) = "194J(a)-Fees for technical services" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194J(a)-Fees for technical services"
            ElseIf TwoDArray(4, i) = "194J(b)-Fees for professional  services or royalty etc" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194J(b)-Fees for professional  services or royalty etc"
            ElseIf TwoDArray(4, i) = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India"
            ElseIf TwoDArray(4, i) = "194LA-Payment of compensation on acquisition of certain immovable" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LA-Payment of compensation on acquisition of certain immovable"
            ElseIf TwoDArray(4, i) = "194LB-Income by way of Interest from Infrastructure Debt fund" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LB-Income by way of Interest from Infrastructure Debt fund"
            ElseIf TwoDArray(4, i) = "194LC-194LC (2)(i) and (ia) Income under clause (i) and (ia) of sub-section (2) of section 194LC" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LC-194LC (2)(i) and (ia) Income under clause (i) and (ia) of sub-section (2) of section 194LC"
            ElseIf TwoDArray(4, i) = "194LC-194LC (2)(ib) Income under clause (ib) of sub-section (2) of section 194LC" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LC-194LC (2)(ib) Income under clause (ib) of sub-section (2) of section 194LC"
            ElseIf TwoDArray(4, i) = "194LC-194LC (2)(ic) Income under clause (ic) of sub-section (2) of section 194LC" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LC-194LC (2)(ic) Income under clause (ic) of sub-section (2) of section 194LC"
            ElseIf TwoDArray(4, i) = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder"
            ElseIf TwoDArray(4, i) = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder"
            ElseIf TwoDArray(4, i) = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR"
            ElseIf TwoDArray(4, i) = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR"
            ElseIf TwoDArray(4, i) = "194LBA(c)-194LBA(c) Income referred to in section 10(23FCA) from units of a business trust-NR" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LBA(c)-194LBA(c) Income referred to in section 10(23FCA) from units of a business trust-NR"
            ElseIf TwoDArray(4, i) = "194LBB-Income in respect of units of investment fund" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LBB-Income in respect of units of investment fund"
            ElseIf TwoDArray(4, i) = "194R-Benefits or perquisites of business or profession" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194R-Benefits or perquisites of business or profession"
            ElseIf TwoDArray(4, i) = "194S-Payment of consideration for transfer of virtual digital asset by persons other than specified persons" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194S-Payment of consideration for transfer of virtual digital asset by persons other than specified persons"
            ElseIf TwoDArray(4, i) = "Proviso to section 194B-Winnings from lotteries and crossword puzzles where consideration is made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such winnings are released" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "Proviso to section 194B-Winnings from lotteries and crossword puzzles where consideration is made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such winnings are released"
            ElseIf TwoDArray(4, i) = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released"
            ElseIf TwoDArray(4, i) = "Proviso to sub- section(1) of section 194S-Payment for transfer of virtual digital asset where payment is in kind or in exchange of another virtual digital asset and tax required to be deducted is paid before such payment is released" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "Proviso to sub- section(1) of section 194S-Payment for transfer of virtual digital asset where payment is in kind or in exchange of another virtual digital asset and tax required to be deducted is paid before such payment is released"
            ElseIf TwoDArray(4, i) = "194LBC-Income in respect of investment in securitization trust" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LBC-Income in respect of investment in securitization trust"
            ElseIf TwoDArray(4, i) = "194LD-TDS on interest on bonds / government securities" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LD-TDS on interest on bonds / government securities"
            ElseIf TwoDArray(4, i) = "194M-Payment of certain sums by certain individuals or HUF" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194M-Payment of certain sums by certain individuals or HUF"
            ElseIf TwoDArray(4, i) = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso"
            ElseIf TwoDArray(4, i) = "194N -First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194N -First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties"
            ElseIf TwoDArray(4, i) = "194N -Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194N -Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso"
            ElseIf TwoDArray(4, i) = "194N-First Proviso read with Third Proviso Payment of certain amount in cash to non-filers being co-operative societies" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194N-First Proviso read with Third Proviso Payment of certain amount in cash to non-filers being co-operative societies"
            ElseIf TwoDArray(4, i) = "194O-Payment of certain sums by e-commerce operator to e-commerce participant." Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194O-Payment of certain sums by e-commerce operator to e-commerce participant."
            ElseIf TwoDArray(4, i) = "194P-Deduction of tax in case of specified senior citizen" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194P-Deduction of tax in case of specified senior citizen"
            ElseIf TwoDArray(4, i) = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods"
                
    'Newly addded by sadineni on 23/06/2026
            ElseIf TwoDArray(4, i) = "194T-Payments to partners of firms" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194T-Payments to partners of firms"
                
                
            ElseIf TwoDArray(4, i) = "195-Other sums payable to a non-resident" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "195-Other sums payable to a non-resident"
            ElseIf TwoDArray(4, i) = "196A-Income in respect of units of non-residents" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "196A-Income in respect of units of non-residents"
            ElseIf TwoDArray(4, i) = "196B-Payments in respect of units to an offshore fund" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "196B-Payments in respect of units to an offshore fund"
            ElseIf TwoDArray(4, i) = "196C-Income from foreign currency bonds or shares of Indian" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "196C-Income from foreign currency bonds or shares of Indian"
            ElseIf TwoDArray(4, i) = "196D-Income of foreign institutional investors from securities" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "196D-Income of foreign institutional investors from securities"
            ElseIf TwoDArray(4, i) = "196D(1A)-Income of specified fund from securities" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "196D(1A)-Income of specified fund from securities"
            ElseIf TwoDArray(4, i) = "194BA(2)-Sub-section (2) of section 194BA Net Winnings from online games where the net winnings are made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such net winnings are released" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194BA(2)-Sub-section (2) of section 194BA Net Winnings from online games where the net winnings are made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such net winnings are released"
            
            Else
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "(Select)"
            End If
        End If
        
        
        
        If Sheet28.Cells(rowcount, DeductedYrColNo).Locked = False Then
            Sheet28.Cells(rowcount, DeductedYrColNo).value = TwoDArray(5, i)
        End If
        
        If Sheet28.Cells(rowcount, BroughtFwdTDSAmtColNo).Locked = False Then
            Sheet28.Cells(rowcount, BroughtFwdTDSAmtColNo).value = TwoDArray(6, i)
        End If
        
        If Sheet28.Cells(rowcount, TaxDeductedOwnHandsColNo).Locked = False Then
            Sheet28.Cells(rowcount, TaxDeductedOwnHandsColNo).value = TwoDArray(7, i)
        End If
        
        If Sheet28.Cells(rowcount, TaxDeductedIncomeColNo).Locked = False Then
            Sheet28.Cells(rowcount, TaxDeductedIncomeColNo).value = TwoDArray(8, i)
        End If
        
        If Sheet28.Cells(rowcount, TaxDeductedTDSColNo).Locked = False Then
            Sheet28.Cells(rowcount, TaxDeductedTDSColNo).value = TwoDArray(9, i)
        End If
        
        If Sheet28.Cells(rowcount, TaxClaimedOwnHandsColNo).Locked = False Then
            Sheet28.Cells(rowcount, TaxClaimedOwnHandsColNo).value = TwoDArray(10, i)
        End If
        
        If Sheet28.Cells(rowcount, TaxClaimedIncomeColNo).Locked = False Then
            Sheet28.Cells(rowcount, TaxClaimedIncomeColNo).value = TwoDArray(11, i)
        End If
        
        If Sheet28.Cells(rowcount, TaxClaimedTDSColNo).Locked = False Then
            Sheet28.Cells(rowcount, TaxClaimedTDSColNo).value = TwoDArray(12, i)
        End If
        
        If Sheet28.Cells(rowcount, TaxClaimedSpouseOthPrsnPANColNo).Locked = False Then
            Sheet28.Cells(rowcount, TaxClaimedSpouseOthPrsnPANColNo).value = TwoDArray(13, i)
        End If
        
        If Sheet28.Cells(rowcount, SpouseOthPrsnAadhaarColNo).Locked = False Then
            Sheet28.Cells(rowcount, SpouseOthPrsnAadhaarColNo).value = TwoDArray(14, i)
        End If
        
        If Sheet28.Cells(rowcount, GrossAmountColNo).Locked = False Then
            Sheet28.Cells(rowcount, GrossAmountColNo).value = TwoDArray(15, i)
        End If
        
        If Sheet28.Cells(rowcount, HeadOfIncomeColNo).Locked = False Then
            If TwoDArray(16, i) = "Income from House Property" Then
                Sheet28.Cells(rowcount, HeadOfIncomeColNo).value = "Income from House Property"
            ElseIf TwoDArray(16, i) = "Income from Business & Profession" Then
                Sheet28.Cells(rowcount, HeadOfIncomeColNo).value = "Income from Business & Profession"
            ElseIf TwoDArray(16, i) = "Income from Capital Gains" Then
                Sheet28.Cells(rowcount, HeadOfIncomeColNo).value = "Income from Capital Gains"
            ElseIf TwoDArray(16, i) = "Income from Other Sources" Then
                Sheet28.Cells(rowcount, HeadOfIncomeColNo).value = "Income from Other Sources"
            ElseIf TwoDArray(16, i) = "Exempt Income" Then
                Sheet28.Cells(rowcount, HeadOfIncomeColNo).value = "Exempt Income"
            ElseIf TwoDArray(16, i) = "Not applicable ( only in case TDS is deducted u/s 194N)" Then
                Sheet28.Cells(rowcount, HeadOfIncomeColNo).value = "Not applicable ( only in case TDS is deducted u/s 194N)"
            ElseIf TwoDArray(16, i) = "NA" Then
                Sheet28.Cells(rowcount, HeadOfIncomeColNo).value = "Not applicable ( only in case TDS is deducted u/s 194N)"

            Else
                Sheet28.Cells(rowcount, HeadOfIncomeColNo).value = "(Select)"
            End If
        End If
       
        rowcount = rowcount + 1
                
    Next i
    
End Sub
'----end change---

Sub ImportScheduleTDS3CSV()
Dim filePath, text As String

Dim answer As Integer
 
answer = MsgBox("Do you want to import CSV? If you import csv, All data will be clear.", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
If answer = vbYes Then
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Please select a CSV file."
        .Filters.add "CSV File", "*.csv", 1
        If .Show = True Then
            filePath = .SelectedItems.item(1)
        Else
            Exit Sub
        End If
        On Error Resume Next
    End With
    
    Open filePath For Input As #1
    text = Input$(LOF(1), 1)
    Close #1
    
    importTDS3CSVFunc (text)
End If
End Sub
Sub importTDS3CSVFunc(text As String)
    Const Delim As String = ","
    Dim MyData As String, strData() As String, TmpAr() As String
    Dim TwoDArray() As String
    Dim i As Long, n As Long
    Dim TotalExRow, TotalXMLRow, TotalDiffRow As Variant
    strData() = Split(text, vbCrLf)
    For i = LBound(strData) To UBound(strData)
        If Len(Trim(strData(i))) <> 0 And i <> 0 Then
            TmpAr = Split(strData(i), Delim)
            n = n + 1
            ReDim Preserve TwoDArray(18, 1 To n)
               
        
            TwoDArray(0, n) = TmpAr(0)      'slno
            TwoDArray(1, n) = TmpAr(1)      'self,other
            TwoDArray(2, n) = TmpAr(2)      'pan of person
            TwoDArray(3, n) = TmpAr(3)      'Aadhaar of other person
            TwoDArray(4, n) = TmpAr(4)      'pan of deductor
            TwoDArray(5, n) = TmpAr(5)      'aadhar
            TwoDArray(6, n) = TmpAr(6)      'year
            TwoDArray(7, n) = TmpAr(7)      'tds b/f
            TwoDArray(8, n) = TmpAr(8)      'deducted in own hands
            TwoDArray(9, n) = TmpAr(9)      'income spouse
            TwoDArray(10, n) = TmpAr(10)    ' TDS spouse
            TwoDArray(11, n) = TmpAr(11)    'Claimed in own hands
            TwoDArray(12, n) = TmpAr(12)    'income
            TwoDArray(13, n) = TmpAr(13)    'TDS
            TwoDArray(14, n) = TmpAr(14)    'PAN
            TwoDArray(15, n) = TmpAr(15)    'Aadhar
            TwoDArray(16, n) = TmpAr(16)    'Gross amount
            TwoDArray(17, n) = TmpAr(17)    'Head of income
            TwoDArray(18, n) = TmpAr(18)
            
        End If
    Next i
    
    Dim TDSCreditName_TDS3, HeadIncome_TDS3, rowcount, SectionOfDeductorColNo
    Dim TDSCreditNameColNo, PANOfBuyerTenantColNo, AmtCarriedFwdColNo, PANofOtherPersonColNo As Variant
    Dim AadhaarOfOtherPersonColNo, AadhaarOfBuyerTenantColNo As Variant
    Dim DeductedYrColNo, BroughtFwdTDSAmtColNo, GrossAmountColNo, HeadOfIncomeColNo As Variant
    Dim TaxClaimedOwnHandsColNo, TaxDeductedOwnHandsColNo, TaxDeductedIncomeColNo, TaxDeductedTDSColNo, TaxClaimedIncomeColNo As Variant
    Dim TaxClaimedTDSColNo, TaxClaimedSpouseOthPrsnPANColNo, SpouseOthPrsnAadhaarColNo As Variant
    
    TDSCreditNameColNo = Sheet28.Range("TDS3.TDSCreditName").Column
    PANofOtherPersonColNo = Sheet28.Range("TDS3.PanOtherPerson").Column
    AadhaarOfOtherPersonColNo = Sheet28.Range("Aadhar_TDS3").Column
    PANOfBuyerTenantColNo = Sheet28.Range("TDS3.PAN").Column
    AadhaarOfBuyerTenantColNo = Sheet28.Range("TDS3.Aadhaar").Column
    SectionOfDeductorColNo = Sheet28.Range("TDS3.TdsIsDeducted").Column 'New
    DeductedYrColNo = Sheet28.Range("TDS3.DeductedYear").Column
    BroughtFwdTDSAmtColNo = Sheet28.Range("TDS3.BroughtFwdTDSAmt").Column
    TaxDeductedOwnHandsColNo = Sheet28.Range("TDS3.DeductedInOwnHands").Column
    TaxDeductedIncomeColNo = Sheet28.Range("TDS3.DeductedIncome5A").Column
    TaxDeductedTDSColNo = Sheet28.Range("TDS3.DeductedTDS5A").Column
    TaxClaimedOwnHandsColNo = Sheet28.Range("TDS3.ClaimedInOwnHands").Column
    TaxClaimedIncomeColNo = Sheet28.Range("TDS3.ClaimedIncome5A").Column
    TaxClaimedTDSColNo = Sheet28.Range("TDS3.ClaimedTDS5A").Column
    TaxClaimedSpouseOthPrsnPANColNo = Sheet28.Range("TDS3.ClaimedPAN5A").Column
    SpouseOthPrsnAadhaarColNo = Sheet28.Range("ClaimedAadhar5A_TDS3").Column
    GrossAmountColNo = Sheet28.Range("TDS3.GrossAmount").Column
    HeadOfIncomeColNo = Sheet28.Range("TDS3.HeadOfIncome").Column
    AmtCarriedFwdColNo = Sheet28.Range("TDS3.AmtCarriedForward").Column
    
    
    TotalExRow = Sheet28.Range("TDS3.TDSCreditName").Rows.count
    TotalXMLRow = n
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Application.EnableEvents = False
        If Sheet28.Range("TDS3.TDSCreditName").Locked = False Then
            Sheet28.Range("TDS3.TDSCreditName").ClearContents
        End If
        If Sheet28.Range("TDS3.PanOtherPerson").Locked = False Then
            Sheet28.Range("TDS3.PanOtherPerson").ClearContents
        End If
        If Sheet28.Range("Aadhar_TDS3").Locked = False Then
            Sheet28.Range("Aadhar_TDS3").ClearContents
        End If
        If Sheet28.Range("TDS3.PAN").Locked = False Then
            Sheet28.Range("TDS3.PAN").ClearContents
        End If
        If Sheet28.Range("TDS3.Aadhaar").Locked = False Then
            Sheet28.Range("TDS3.Aadhaar").ClearContents
        End If
        If Sheet28.Range("TDS3.TdsIsDeducted").Locked = False Then
            Sheet28.Range("TDS3.TdsIsDeducted").ClearContents
        End If
        If Sheet28.Range("TDS3.DeductedYear").Locked = False Then
            Sheet28.Range("TDS3.DeductedYear").ClearContents
        End If
        If Sheet28.Range("TDS3.BroughtFwdTDSAmt").Locked = False Then
            Sheet28.Range("TDS3.BroughtFwdTDSAmt").ClearContents
        End If
        If Sheet28.Range("TDS3.DeductedInOwnHands").Locked = False Then
            Sheet28.Range("TDS3.DeductedInOwnHands").ClearContents
        End If
        If Sheet28.Range("TDS3.DeductedIncome5A").Locked = False Then
            Sheet28.Range("TDS3.DeductedIncome5A").ClearContents
        End If
        If Sheet28.Range("TDS3.DeductedTDS5A").Locked = False Then
            Sheet28.Range("TDS3.DeductedTDS5A").ClearContents
        End If
        If Sheet28.Range("TDS3.ClaimedInOwnHands").Locked = False Then
            Sheet28.Range("TDS3.ClaimedInOwnHands").ClearContents
        End If
        If Sheet28.Range("TDS3.ClaimedIncome5A").Locked = False Then
            Sheet28.Range("TDS3.ClaimedIncome5A").ClearContents
        End If
        If Sheet28.Range("TDS3.ClaimedTDS5A").Locked = False Then
            Sheet28.Range("TDS3.ClaimedTDS5A").ClearContents
        End If
        If Sheet28.Range("TDS3.ClaimedPAN5A").Locked = False Then
            Sheet28.Range("TDS3.ClaimedPAN5A").ClearContents
        End If
        If Sheet28.Range("ClaimedAadhar5A_TDS3").Locked = False Then
            Sheet28.Range("ClaimedAadhar5A_TDS3").ClearContents
        End If
        If Sheet28.Range("TDS3.GrossAmount").Locked = False Then
            Sheet28.Range("TDS3.GrossAmount").ClearContents
        End If
        If Sheet28.Range("TDS3.HeadOfIncome").Locked = False Then
            Sheet28.Range("TDS3.HeadOfIncome").ClearContents
        End If
    Application.EnableEvents = True
    End If
    
    If (TotalDiffRow > 0) Then
        AddDiffRows_TDS3oth (TotalDiffRow)
    End If
    rowcount = getRowNo(Sheet28.Range("TDS3.TDSCreditName").name)

    For i = 1 To n
        If Sheet28.Cells(rowcount, TDSCreditNameColNo).Locked = False Then
            If UCase(Mid(TwoDArray(0, i), 1, 1)) = "S" Then
                TwoDArray(0, i) = "Self"
            End If
            If UCase(Mid(TwoDArray(0, i), 1, 1)) = "O" Then
                TwoDArray(0, i) = "Other Person"
            End If
            Sheet28.Cells(rowcount, TDSCreditNameColNo).value = TwoDArray(0, i)
        End If
        If Sheet28.Cells(rowcount, PANofOtherPersonColNo).Locked = False Then
            Sheet28.Cells(rowcount, PANofOtherPersonColNo).value = TwoDArray(1, i)
        End If
        If Sheet28.Cells(rowcount, AadhaarOfOtherPersonColNo).Locked = False Then
            Sheet28.Cells(rowcount, AadhaarOfOtherPersonColNo).value = TwoDArray(2, i)
        End If
        If Sheet28.Cells(rowcount, PANOfBuyerTenantColNo).Locked = False Then
            Sheet28.Cells(rowcount, PANOfBuyerTenantColNo).value = TwoDArray(3, i)
        End If
        If Sheet28.Cells(rowcount, AadhaarOfBuyerTenantColNo).Locked = False Then
            Sheet28.Cells(rowcount, AadhaarOfBuyerTenantColNo).value = TwoDArray(4, i)
        End If
        
        'New section
        If Sheet28.Cells(rowcount, SectionOfDeductorColNo).Locked = False Then
            If TwoDArray(5, i) = "192-Salary-Payment to Government employees other than Indian Government employees" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "192-Salary-Payment to Government employees other than Indian Government employees"
            ElseIf TwoDArray(5, i) = "192-Salary-Payment to employees other than Government employees" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "192-Salary-Payment to employees other than Government employees"
            ElseIf TwoDArray(5, i) = "192-Salary-Payment to Indian Government employees" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "192-Salary-Payment to Indian Government employees"
            ElseIf TwoDArray(5, i) = "192A-TDS on PF withdrawal" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "192A-TDS on PF withdrawal"
            ElseIf TwoDArray(5, i) = "193-Interest on Securities" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "193-Interest on Securities"
            ElseIf TwoDArray(5, i) = "194-Dividends" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194-Dividends"
            ElseIf TwoDArray(5, i) = "194A-Interest other than 'Interest on securities'" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194A-Interest other than 'Interest on securities'"
            ElseIf TwoDArray(5, i) = "194B-Winning from lottery or crossword puzzle" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194B-Winning from lottery or crossword puzzle"
            ElseIf TwoDArray(5, i) = "194BA-Winnings from online games" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194BA-Winnings from online games"
            ElseIf TwoDArray(5, i) = "194BB-Winning from horse race" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194BB-Winning from horse race"
            ElseIf TwoDArray(5, i) = "194C-Payments to contractors and sub-contractors" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194C-Payments to contractors and sub-contractors"
            ElseIf TwoDArray(5, i) = "194D-Insurance commission" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194D-Insurance commission"
            ElseIf TwoDArray(5, i) = "194DA-Payment in respect of life insurance policy" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194DA-Payment in respect of life insurance policy"
            ElseIf TwoDArray(5, i) = "194E-Payments to non-resident sportsmen or sports associations" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194E-Payments to non-resident sportsmen or sports associations"
            ElseIf TwoDArray(5, i) = "194EE-Payments in respect of deposits under National Savings" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194EE-Payments in respect of deposits under National Savings"
            ElseIf TwoDArray(5, i) = "194F-Payments on account of repurchase of units by Mutual Fund or Unit Trust of India" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194F-Payments on account of repurchase of units by Mutual Fund or Unit Trust of India"
            ElseIf TwoDArray(5, i) = "194G-Commission, price, etc. on sale of lottery tickets" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194G-Commission, price, etc. on sale of lottery tickets"
            ElseIf TwoDArray(5, i) = "194H-Commission or brokerage" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194H-Commission or brokerage"
            ElseIf TwoDArray(5, i) = "194I(a)-Rent on hiring of plant and machinery" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194I(a)-Rent on hiring of plant and machinery"
            ElseIf TwoDArray(5, i) = "194I(b)-Rent on other than plant and machinery" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194I(b)-Rent on other than plant and machinery"
            ElseIf TwoDArray(5, i) = "194IA-TDS on Sale of immovable property" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194IA-TDS on Sale of immovable property"
            ElseIf TwoDArray(5, i) = "194IB-Payment of rent by certain individuals or Hindu undivided" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194IB-Payment of rent by certain individuals or Hindu undivided"
            ElseIf TwoDArray(5, i) = "194IC-Payment under specified agreement" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194IC-Payment under specified agreement"
            ElseIf TwoDArray(5, i) = "194J(a)-Fees for technical services" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194J(a)-Fees for technical services"
            ElseIf TwoDArray(5, i) = "194J(b)-Fees for professional  services or royalty etc" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194J(b)-Fees for professional  services or royalty etc"
            ElseIf TwoDArray(5, i) = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India"
            ElseIf TwoDArray(5, i) = "194LA-Payment of compensation on acquisition of certain immovable" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LA-Payment of compensation on acquisition of certain immovable"
            ElseIf TwoDArray(5, i) = "194LB-Income by way of Interest from Infrastructure Debt fund" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LB-Income by way of Interest from Infrastructure Debt fund"
            ElseIf TwoDArray(5, i) = "194LC-194LC (2)(i) and (ia) Income under clause (i) and (ia) of sub-section (2) of section 194LC" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LC-194LC (2)(i) and (ia) Income under clause (i) and (ia) of sub-section (2) of section 194LC"
            ElseIf TwoDArray(5, i) = "194LC-194LC (2)(ib) Income under clause (ib) of sub-section (2) of section 194LC" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LC-194LC (2)(ib) Income under clause (ib) of sub-section (2) of section 194LC"
            ElseIf TwoDArray(5, i) = "194LC-194LC (2)(ic) Income under clause (ic) of sub-section (2) of section 194LC" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LC-194LC (2)(ic) Income under clause (ic) of sub-section (2) of section 194LC"
            ElseIf TwoDArray(5, i) = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder"
            ElseIf TwoDArray(5, i) = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder"
            ElseIf TwoDArray(5, i) = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR"
            ElseIf TwoDArray(5, i) = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR"
            ElseIf TwoDArray(5, i) = "194LBA(c)-194LBA(c) Income referred to in section 10(23FCA) from units of a business trust-NR" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LBA(c)-194LBA(c) Income referred to in section 10(23FCA) from units of a business trust-NR"
            ElseIf TwoDArray(5, i) = "194LBB-Income in respect of units of investment fund" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LBB-Income in respect of units of investment fund"
            ElseIf TwoDArray(5, i) = "194R-Benefits or perquisites of business or profession" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194R-Benefits or perquisites of business or profession"
            ElseIf TwoDArray(5, i) = "194S-Payment of consideration for transfer of virtual digital asset by persons other than specified persons" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194S-Payment of consideration for transfer of virtual digital asset by persons other than specified persons"
            ElseIf TwoDArray(5, i) = "Proviso to section 194B-Winnings from lotteries and crossword puzzles where consideration is made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such winnings are released" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "Proviso to section 194B-Winnings from lotteries and crossword puzzles where consideration is made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such winnings are released"
            ElseIf TwoDArray(5, i) = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released"
            ElseIf TwoDArray(5, i) = "Proviso to sub- section(1) of section 194S-Payment for transfer of virtual digital asset where payment is in kind or in exchange of another virtual digital asset and tax required to be deducted is paid before such payment is released" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "Proviso to sub- section(1) of section 194S-Payment for transfer of virtual digital asset where payment is in kind or in exchange of another virtual digital asset and tax required to be deducted is paid before such payment is released"
            ElseIf TwoDArray(5, i) = "194LBC-Income in respect of investment in securitization trust" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LBC-Income in respect of investment in securitization trust"
            ElseIf TwoDArray(5, i) = "194LD-TDS on interest on bonds / government securities" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194LD-TDS on interest on bonds / government securities"
            ElseIf TwoDArray(5, i) = "194M-Payment of certain sums by certain individuals or HUF" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194M-Payment of certain sums by certain individuals or HUF"
            ElseIf TwoDArray(5, i) = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso"
            ElseIf TwoDArray(5, i) = "194N -First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194N -First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties"
            ElseIf TwoDArray(5, i) = "194N -Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194N -Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso"
            ElseIf TwoDArray(5, i) = "194N-First Proviso read with Third Proviso Payment of certain amount in cash to non-filers being co-operative societies" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194N-First Proviso read with Third Proviso Payment of certain amount in cash to non-filers being co-operative societies"
            ElseIf TwoDArray(5, i) = "194O-Payment of certain sums by e-commerce operator to e-commerce participant." Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194O-Payment of certain sums by e-commerce operator to e-commerce participant."
            ElseIf TwoDArray(5, i) = "194P-Deduction of tax in case of specified senior citizen" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194P-Deduction of tax in case of specified senior citizen"
            ElseIf TwoDArray(5, i) = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods"
                
    'Newly addded by sadineni on 23/06/2026
            ElseIf TwoDArray(5, i) = "194T-Payments to partners of firms" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194T-Payments to partners of firms"
        
            ElseIf TwoDArray(5, i) = "195-Other sums payable to a non-resident" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "195-Other sums payable to a non-resident"
            ElseIf TwoDArray(5, i) = "196A-Income in respect of units of non-residents" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "196A-Income in respect of units of non-residents"
            ElseIf TwoDArray(5, i) = "196B-Payments in respect of units to an offshore fund" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "196B-Payments in respect of units to an offshore fund"
            ElseIf TwoDArray(5, i) = "196C-Income from foreign currency bonds or shares of Indian" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "196C-Income from foreign currency bonds or shares of Indian"
            ElseIf TwoDArray(5, i) = "196D-Income of foreign institutional investors from securities" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "196D-Income of foreign institutional investors from securities"
            ElseIf TwoDArray(5, i) = "196D(1A)-Income of specified fund from securities" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "196D(1A)-Income of specified fund from securities"
            ElseIf TwoDArray(5, i) = "194BA(2)-Sub-section (2) of section 194BA Net Winnings from online games where the net winnings are made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such net winnings are released" Then
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "194BA(2)-Sub-section (2) of section 194BA Net Winnings from online games where the net winnings are made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such net winnings are released"
            
            Else
                Sheet28.Cells(rowcount, SectionOfDeductorColNo).value = "(Select)"
            End If
        End If
        
        If Sheet28.Cells(rowcount, DeductedYrColNo).Locked = False Then
            Sheet28.Cells(rowcount, DeductedYrColNo).value = TwoDArray(6, i)
        End If
        If Sheet28.Cells(rowcount, BroughtFwdTDSAmtColNo).Locked = False Then
            Sheet28.Cells(rowcount, BroughtFwdTDSAmtColNo).value = TwoDArray(7, i)
        End If
        If Sheet28.Cells(rowcount, TaxDeductedOwnHandsColNo).Locked = False Then
            Sheet28.Cells(rowcount, TaxDeductedOwnHandsColNo).value = TwoDArray(8, i)
        End If
        If Sheet28.Cells(rowcount, TaxDeductedIncomeColNo).Locked = False Then
            Sheet28.Cells(rowcount, TaxDeductedIncomeColNo).value = TwoDArray(9, i)
        End If
        If Sheet28.Cells(rowcount, TaxDeductedTDSColNo).Locked = False Then
            Sheet28.Cells(rowcount, TaxDeductedTDSColNo).value = TwoDArray(10, i)
        End If
        If Sheet28.Cells(rowcount, TaxClaimedOwnHandsColNo).Locked = False Then
            Sheet28.Cells(rowcount, TaxClaimedOwnHandsColNo).value = TwoDArray(11, i)
        End If
        If Sheet28.Cells(rowcount, TaxClaimedIncomeColNo).Locked = False Then
            Sheet28.Cells(rowcount, TaxClaimedIncomeColNo).value = TwoDArray(12, i)
        End If
        If Sheet28.Cells(rowcount, TaxClaimedTDSColNo).Locked = False Then
            Sheet28.Cells(rowcount, TaxClaimedTDSColNo).value = TwoDArray(13, i)
        End If
        If Sheet28.Cells(rowcount, TaxClaimedSpouseOthPrsnPANColNo).Locked = False Then
            Sheet28.Cells(rowcount, TaxClaimedSpouseOthPrsnPANColNo).value = TwoDArray(14, i)
        End If
        If Sheet28.Cells(rowcount, SpouseOthPrsnAadhaarColNo).Locked = False Then
            Sheet28.Cells(rowcount, SpouseOthPrsnAadhaarColNo).value = TwoDArray(15, i)
        End If
        If Sheet28.Cells(rowcount, GrossAmountColNo).Locked = False Then
            Sheet28.Cells(rowcount, GrossAmountColNo).value = TwoDArray(16, i)
        End If
        If Sheet28.Cells(rowcount, HeadOfIncomeColNo).Locked = False Then
            If TwoDArray(17, i) = "Income from House Property" Then
                Sheet28.Cells(rowcount, HeadOfIncomeColNo).value = "Income from House Property"
            ElseIf TwoDArray(17, i) = "Income from Business & Profession" Then
                Sheet28.Cells(rowcount, HeadOfIncomeColNo).value = "Income from Business & Profession"
            ElseIf TwoDArray(17, i) = "Income from Capital Gains" Then
                Sheet28.Cells(rowcount, HeadOfIncomeColNo).value = "Income from Capital Gains"
            ElseIf TwoDArray(17, i) = "Income from Other Sources" Then
                Sheet28.Cells(rowcount, HeadOfIncomeColNo).value = "Income from Other Sources"
            ElseIf TwoDArray(17, i) = "Exempt Income" Then
                Sheet28.Cells(rowcount, HeadOfIncomeColNo).value = "Exempt Income"
            ElseIf TwoDArray(17, i) = "Not applicable ( only in case TDS is deducted u/s 194N)" Then
                Sheet28.Cells(rowcount, HeadOfIncomeColNo).value = "Not applicable ( only in case TDS is deducted u/s 194N)"
            ElseIf TwoDArray(17, i) = "NA" Then
                Sheet28.Cells(rowcount, HeadOfIncomeColNo).value = "Not applicable ( only in case TDS is deducted u/s 194N)"
            Else
                Sheet28.Cells(rowcount, HeadOfIncomeColNo).value = "(Select)"
            End If
        End If
        If Sheet28.Cells(rowcount, AmtCarriedFwdColNo).Locked = False Then
            Sheet28.Cells(rowcount, AmtCarriedFwdColNo).value = TwoDArray(18, i)
        End If
        rowcount = rowcount + 1
        
    Next i
    
End Sub

