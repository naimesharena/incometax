Attribute VB_Name = "md80EEA"
Option Explicit

Public MsgBox_80EEA As Variant
Public end_80EEA As Variant
Public end_80EEALoanfrm As Variant
Public end_80EEAIFSC As Variant
Public end_80EEAbankName As Variant
Public end_80EEAPAN As Variant
Public end_80EEAAccntNum As Variant
Public end_80EEALoanDate As Variant
Public end_80EEALoanAmt As Variant
Public end_80EEALoanOutstanding As Variant
Public end_80EEAIntrst As Variant
Public Deduction_80EE_and_80EEA_chk As Boolean

Sub ValidateSheet80EEA_Click()
Dim vbMessgaeCaption As String

vbMessgaeCaption = "ITR 3: AY: 2024-25"

If Validate80EEA_All Then
Sheet60.Activate
fmsgboxStatus "Schedule 80EEA is OK"
CloseMsg
End If
End Sub

Function Validate80EEA_All()
 Validate80EEA_All = True
Dim vbMessgaeCaption As String
vbMessgaeCaption = "Error"
   Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("80E_80EE_80EEA_80EEB")
  
If Not Validate_80EEA Then
    sourceSheet.Activate
    fmsgbox (MsgBox_80EEA)
    CloseMsg
End If


Dim rowcount_80EEA, i, j, rowcount_Int24B   As Long


'for80EEA
setTblinfo_80EEALoanfrm
setTblinfo_80EEABankName
setTblinfo_80EEAAccntNum
setTblinfo_80EEALoanDate
setTblinfo_80EEALoanAmt
setTblinfo_80EEALoanOutstanding
setTblinfo_80EEAIntrst

Dim Data80EEA, Data24b, Combination_80EEA_24B As Boolean

Data80EEA = end_80EEALoanfrm > 0 And end_80EEAbankName > 0 And end_80EEAAccntNum > 0 And end_80EEALoanDate > 0 And end_80EEALoanAmt > 0 And end_80EEALoanOutstanding > 0 And end_80EEAIntrst > 0
end_80EEA = WorksheetFunction.Max(0, end_80EEALoanfrm, end_80EEAbankName, end_80EEAAccntNum, end_80EEALoanDate, end_80EEALoanAmt, end_80EEALoanOutstanding, end_80EEAIntrst)

If Data80EEA = True Then

    setTblinfo_24bLoanfrm
    setTblinfo_24bBankName
    setTblinfo_24bAccntNum
    setTblinfo_24bLoanDate
    setTblinfo_24bLoanAmt
    setTblinfo_24bLoanOutstanding
    setTblinfo_24bIntrst
    
     Data24b = end_24bLoanfrm > 0 And end_24bbankName > 0 And end_24bAccntNum > 0 And end_24bLoanDate > 0 And end_24bLoanAmt > 0 And end_24bLoanOutstanding > 0 And end_24bIntrst > 0
     end_24b = WorksheetFunction.Max(0, end_24bLoanfrm, end_24bbankName, end_24bAccntNum, end_24bLoanDate, end_24bLoanAmt, end_24bLoanOutstanding, end_24bIntrst)
    
    rowcount_80EEA = getRowNo(Sheet60.Range("LoanfrmBankOrInstitute.80EEA").name)
    
    If end_80EEA = 1 Then
       end_80EEA = rowcount_80EEA
    Else
       end_80EEA = ((end_80EEA - 1) + rowcount_80EEA)
    End If
    
    
    
    Dim Dedn_80EEA, Dedn_24B, HpCount
    
    Dedn_80EEA = Sheet60.Range("Combination_80EEA").Column
    
    
    If Data24b = True Then
        Combination_80EEA_24B = False
         For HpCount = 1 To Sheet8.Range("PropertySectionCOunt").value
               Dedn_24B = Sheet8.Range("Combination_24B" & HpCount).Column
               rowcount_Int24B = getRowNo(Sheet8.Range("LoanfrmBankOrInstitute.24b" & HpCount).name)
               end_24b = WorksheetFunction.Max(0, end_24bLoanfrm, end_24bbankName, end_24bAccntNum, end_24bLoanDate, end_24bLoanAmt, end_24bLoanOutstanding, end_24bIntrst)
    
                If end_24b = 1 Then
                    end_24b = rowcount_Int24B
                Else
                    end_24b = ((end_24b - 1) + rowcount_Int24B)
                End If
                For i = rowcount_80EEA To end_80EEA
                    For j = rowcount_Int24B To end_24b
                        If UCase(Trim(Sheet60.Cells(i, Dedn_80EEA).value)) = UCase(Trim(Sheet8.Cells(j, Dedn_24B).value)) Then
                           Combination_80EEA_24B = True
                           GoTo checking
                        End If
                    Next
                Next
           Next
    Else
        Combination_80EEA_24B = False
    End If

checking:
If Combination_80EEA_24B = False Then
        MsgBox_80EEA = MsgBox_80EEA & ("""Deduction u/s 80EEA can be claimed only if the limit u/s 24(b) is exhausted. So please note that the loan details of respective property under which 80EEA is being claimed should be the same as in Table 24(b) of Schedule HP""") & Chr(13)
        fmsgboxStatus MsgBox_80EEA
        Validate80EEA_All = False
        CloseMsg
End If


Dim k, Total As Long

For k = 1 To Sheet8.Range("PropertySectionCOunt").value
Total = Total + Sheet8.Range("TotAmt.24b" & k).value

Next

If (Sheet60.Range("TotAmt.80EEA").value > 0 And Not Total > 0) Then
        sourceSheet.Activate
        MsgBox_80EE = MsgBox_80EE & " ""Deduction u/s 80EEA can be claimed only if the limit u/s 24(b) is exhausted.""" & Chr(13)
       fmsgboxStatus MsgBox_80EE
       CloseMsg
    End If
    
If Combination_80EEA_24B = False Then
        MsgBox_80EEA = MsgBox_80EEA & ("""Deduction u/s 80EEA can be claimed only if the limit u/s 24(b) is exhausted. So please note that the loan details of respective property under which 80EEA is being claimed should be the same as in Table 24(b) of Schedule HP""") & Chr(13)
        fmsgboxStatus MsgBox_80EEA
        Validate80EEA_All = False
        CloseMsg
End If

End If

 If MsgBox_80EEA <> "" Then
 sourceSheet.Activate
  Validate80EEA_All = False
  Else
  sourceSheet.Activate
  Validate80EEA_All = True
  End If
              
End Function




Function Validate_80EEA()
Validate_80EEA = True
MsgBox_80EEA = ""

setTblinfo_80EEALoanfrm
setTblinfo_80EEABankName
setTblinfo_80EEAAccntNum
setTblinfo_80EEALoanDate
setTblinfo_80EEALoanAmt
setTblinfo_80EEALoanOutstanding
setTblinfo_80EEAIntrst


end_80EEA = WorksheetFunction.Max(0, end_80EEALoanfrm, end_80EEAbankName, , end_80EEAAccntNum, end_80EEALoanDate, end_80EEALoanAmt, end_80EEALoanOutstanding, end_80EEAIntrst)

If end_80EEA > 0 Then
    If (Sheet60.Range("Stampduty.80EEA").value) = "" Then
        MsgBox_80EEA = MsgBox_80EEA & " ""Stamp value of residential house property is mandatory in schedule 80EEA""" & Chr(13)
        Validate_80EEA = False
    End If
End If

     If (Sheet60.Range("Stampduty.80EEA").value) > 4500000 Then
        MsgBox_80EEA = MsgBox_80EEA & " ""Stamp value of residential house property shall not be more than Rs. 45 Lakhs""" & Chr(13)
        Validate_80EEA = False
    
    
    End If
    
    
If Not ValidateLoanfrm_80EEA Then Validate_80EEA = False
If Not ValidateBankName_80EEA Then Validate_80EEA = False
If Not ValidateAccntNum_80EEA Then Validate_80EEA = False
If Not ValidateLoanDate_80EEA Then Validate_80EEA = False
If Not ValidateLoanAmt_80EEA Then Validate_80EEA = False
If Not ValidateLoanOutstanding_80EEA Then Validate_80EEA = False
If Not ValidateIntrst_80EEA Then Validate_80EEA = False


    If (Sheet60.Range("TotAmt.80EEA").value) > 2000000 Then
        MsgBox_80EEA = MsgBox_80EEA & " To claim ""Interest payable on borrowed capital more than Rs.20 lakhs"", you may please consider filing ITR 3. Refer Rule 12 for further details" & Chr(13)
        Validate_80EEA = False
    End If
    
    If Len(Sheet60.Range("TotAmt.80EEA").value) > 14 Then
        MsgBox_80EEA = MsgBox_80EEA & "* Total of Interest u/s 80EEA in schedule 80EEA cannot exceed 14 digits" & Chr(13)
        Validate_80EEA = False
    End If
    
    

End Function

Sub setTblinfo_80EEALoanfrm()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("LoanfrmBankOrInstitute.80EEA").count
    Set rangecells = Range("LoanfrmBankOrInstitute.80EEA").Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
    end_80EEALoanfrm = ccount
End Sub

Sub setTblinfo_80EEABankName()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("bankName.80EEa").count
    Set rangecells = Range("bankName.80EEA").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEAbankName = ccount
End Sub

Sub setTblinfo_80EEAAccntNum()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanAccNum.80EEA").count
    Set rangecells = Range("loanAccNum.80EEA").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEAAccntNum = ccount
End Sub

Sub setTblinfo_80EEALoanDate()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanDate.80EEA").count
    Set rangecells = Range("loanDate.80EEA").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEALoanDate = ccount
End Sub

Sub setTblinfo_80EEALoanAmt()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanAmt.80EEA").count
    Set rangecells = Range("loanAmt.80EEA").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEALoanAmt = ccount
End Sub

Sub setTblinfo_80EEALoanOutstanding()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanOutstanding.80EEA").count
    Set rangecells = Range("loanOutstanding.80EEA").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEALoanOutstanding = ccount
End Sub
Sub setTblinfo_80EEAIntrst()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Intrst.80EEA").count
    Set rangecells = Range("Intrst.80EEA").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEAIntrst = ccount
End Sub


Function ValidateLoanfrm_80EEA() As Boolean
    ValidateLoanfrm_80EEA = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("LoanfrmBankOrInstitute.80EEA").Cells
    ReDim Loanfrm_80EEA(end_80EEA)
    For i = 1 To end_80EEA
        Loanfrm_80EEA(i) = rangecells.item(i).value
        If isdropdownblank(Loanfrm_80EEA(i)) Then
              MsgBox_80EEA = MsgBox_80EEA + "* Please select dropdown from ""Loan taken from"" at Sr. No " & i & "" & Chr(13)
            ValidateLoanfrm_80EEA = False
            Exit Function
        End If
Next
End Function

Function ValidateBankName_80EEA() As Boolean
    ValidateBankName_80EEA = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("bankName.80EEA").Cells
    ReDim BankName_80EEA(end_80EEA)
    For i = 1 To end_80EEA
        BankName_80EEA(i) = rangecells.item(i).value
        If Not chkCompulsory(BankName_80EEA(i)) Then
             MsgBox_80EEA = MsgBox_80EEA + "* Please provide ""Name of the Bank/ Institution from which the loan is taken"" at Sr. No " & i & "" & Chr(13)
            ValidateBankName_80EEA = False
            Exit Function
        End If
         If Len(BankName_80EEA(i)) > 125 Then
          MsgBox_80EEA = MsgBox_80EEA + "* Name of the Bank/ Institution from which the loan is taken in schedule 80EEA at Sr. No " & i & "  should be less than or equal to 125 characters." & Chr(13)
            ValidateBankName_80EEA = False
            Exit Function
        End If
        
        If Not checkfieldSuperSpecialcharacter(BankName_80EEA(i)) Then
             MsgBox_80EEA = MsgBox_80EEA + "* Name of the Bank/ Institution from which the loan is taken in schedule 80EEA at Sl.no. " & i & " should not Contain <, >, characters." & Chr(13)
            ValidateBankName_80EEA = False
            Exit Function
        End If
         
Next
End Function

Function ValidateAccntNum_80EEA() As Boolean
    ValidateAccntNum_80EEA = True

    Dim rangecells As Range
    Dim rangecells1 As Range
    
    Dim i As Long
    Set rangecells = Sheet60.Range("loanAccNum.80EEA").Cells
    Set rangecells1 = Sheet60.Range("LoanfrmBankOrInstitute.80EEA").Cells
    ReDim AccntNum_80EEA(end_80EEA)
    ReDim BankorInst_80EEA(end_80EEA)
    
    For i = 1 To end_80EEA
        AccntNum_80EEA(i) = rangecells.item(i).value
        BankorInst_80EEA(i) = rangecells1.item(i).value
        If Not chkCompulsory(AccntNum_80EEA(i)) Then
             MsgBox_80EEA = MsgBox_80EEA + "* Please provide ""Loan Account number of the Bank / Institution from which loan is taken"" at Sr. No " & i & "" & Chr(13)
            ValidateAccntNum_80EEA = False
            Exit Function
        End If
         If Len(AccntNum_80EEA(i)) > 20 Then
          MsgBox_80EEA = MsgBox_80EEA + "* Loan Account number  at Sr. No " & i & " in Sheet 80EEA less than 20 characters." & Chr(13)
            ValidateAccntNum_80EEA = False
            Exit Function
        End If
        
        If BankorInst_80EEA(i) = "Bank" Then
        If Not ValidateBankAccountNumber_80EEA(AccntNum_80EEA(i), i) Then
        ValidateAccntNum_80EEA = False
        Exit Function
        End If
        End If
    
        If BankorInst_80EEA(i) = "Institution" Then
        If Not checkfieldspecialcharacter1(AccntNum_80EEA(i)) Then
            MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number of the Bank / Institution at Sr.No " & i & " is invalid  in schedule 80EEA, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
           ValidateAccntNum_80EEA = False
            Exit Function
        End If
        End If
        
         
Next
End Function

Function ValidateLoanDate_80EEA() As Boolean
    ValidateLoanDate_80EEA = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("loanDate.80EEA").Cells
    ReDim LoanDate_80EEA(end_80EEA)
    For i = 1 To end_80EEA
        LoanDate_80EEA(i) = rangecells.item(i).value
        If Not chkCompulsory(LoanDate_80EEA(i)) Then
             MsgBox_80EEA = MsgBox_80EEA + "* Please provide ""Date of sanction of loan"" in schedule 80EEA at Sr. No " & i & "" & Chr(13)
            ValidateLoanDate_80EEA = False
            Exit Function
        End If
         If Len(LoanDate_80EEA(i)) > 10 Then
          MsgBox_80EEA = MsgBox_80EEA + "* Date of sanction of loan  at Sr. No " & i & " in Sheet 80EEA less than 10 characters." & Chr(13)
            ValidateLoanDate_80EEA = False
            Exit Function
        End If
        'CheckDateddmmyyyy
         If Not CheckDateddmmyyyy(LoanDate_80EEA(i)) Then
         MsgBox_80EEA = MsgBox_80EEA + " ""Please enter date in valid format"" at Sr. No " & i & "." & Chr(13)
        ValidateLoanDate_80EEA = False
                 

          Exit Function
        End If
        
        If Not ChkMaxDate_80EEA(Trim(LoanDate_80EEA(i)), "31-03-2022") Then
            MsgBox_80EEA = MsgBox_80EEA + "* ""Date of sanction of Loan"" shall be between 01/04/2019 to 31/03/2022 in schedule 80EEA at Sr. No " & i & "." & Chr(13)
            ValidateLoanDate_80EEA = False
                 

          Exit Function
        End If
        
        If Not ChkMinInclusiveDate(Trim(Dformat(LoanDate_80EEA(i), "yyyy-mm-dd")), "2019-04-01") Then
           MsgBox_80EEA = MsgBox_80EEA + "* ""Date of sanction of Loan"" shall be between 01/04/2019 to 31/03/2022 in schedule 80EEA at Sr. No " & i & "." & Chr(13)
            ValidateLoanDate_80EEA = False
                 

          Exit Function
        End If
          
         
Next
End Function

Function ValidateLoanAmt_80EEA() As Boolean
    ValidateLoanAmt_80EEA = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("loanAmt.80EEA").Cells
    ReDim LoanAmt_80EEA(end_80EEA)
    For i = 1 To end_80EEA
        LoanAmt_80EEA(i) = rangecells.item(i).value
        If Not chkCompulsory(LoanAmt_80EEA(i)) Then
             MsgBox_80EEA = MsgBox_80EEA + "* ""Please provide Total amount of Loan in schedule 80EEA"" at Sr. No " & i & "" & Chr(13)
            ValidateLoanAmt_80EEA = False
            Exit Function
        End If

        
        If Not IsNumeric(LoanAmt_80EEA(i)) Then
            MsgBox_80EEA = MsgBox_80EEA & " Loan amount at Sr. No  " & i & "  in schedule 80EEA should be Numeric value" & Chr(13)
            ValidateLoanAmt_80EEA = False
            Exit Function
        End If
        
        If LoanAmt_80EEA(i) > 99999999999999# Then
            MsgBox_80EEA = MsgBox_80EEA & "* Loan amount at Sr. No  " & i & "  in schedule 80EEA cannot exceed 14 digits" & Chr(13)
            ValidateLoanAmt_80EEA = False
            Exit Function
        End If
        
        If LoanAmt_80EEA(i) < 0 Or LoanAmt_80EEA(i) = 0 Then
            MsgBox_80EEA = MsgBox_80EEA & "* Total amount of Loan should be more than 0 at Sr. No  " & i & "" & Chr(13)
            ValidateLoanAmt_80EEA = False
            Exit Function
        End If
         
Next
End Function

Function ValidateLoanOutstanding_80EEA() As Boolean
    ValidateLoanOutstanding_80EEA = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("loanOutstanding.80EEA").Cells
    ReDim LoanOutStanding_80EEA(end_80EEA)
    For i = 1 To end_80EEA
        LoanOutStanding_80EEA(i) = rangecells.item(i).value
        If Not chkCompulsory(LoanOutStanding_80EEA(i)) Then
             MsgBox_80EEA = MsgBox_80EEA + "* ""Loan  outstanding as on last date of financial year is mandatory in schedule 80EEA. Please enter 0 in case the entire loan is repaid during the year"" at Sr. No " & i & "" & Chr(13)
            ValidateLoanOutstanding_80EEA = False
            Exit Function
        End If
        
        If Not IsNumeric(LoanOutStanding_80EEA(i)) Then
            MsgBox_80EEA = MsgBox_80EEA & " Loan outstanding at Sr. No  " & i & "  in schedule 80EEA should be Numeric value" & Chr(13)
            ValidateLoanOutstanding_80EEA = False
            Exit Function
        End If
        
        If LoanOutStanding_80EEA(i) > 99999999999999# Then
            MsgBox_80EEA = MsgBox_80EEA & "* Loan outstanding at Sr. No  " & i & " in schedule 80EEA cannot exceed 14 digits" & Chr(13)
            ValidateLoanOutstanding_80EEA = False
            Exit Function
        End If
        
        If LoanOutStanding_80EEA(i) < 0 Then
            MsgBox_80EEA = MsgBox_80EEA & "* Loan outstanding as on  last date of financial year can't be less than 0 at Sr. No  " & i & ". You may please enter as 0 if it become negative as result of excess payment." & Chr(13)
            ValidateLoanOutstanding_80EEA = False
            Exit Function
        End If
         
Next
End Function


Function ValidateIntrst_80EEA() As Boolean
    ValidateIntrst_80EEA = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("Intrst.80EEA").Cells
    ReDim Intrst_80EEA(end_80EEA)
    For i = 1 To end_80EEA
        Intrst_80EEA(i) = rangecells.item(i).value
        If Not chkCompulsory(Intrst_80EEA(i)) Then
             MsgBox_80EEA = MsgBox_80EEA + "* ""Please provide Interest u/s 80EEA"" at Sr. No " & i & "" & Chr(13)
            ValidateIntrst_80EEA = False
            Exit Function
        End If


        If Not IsNumeric(Intrst_80EEA(i)) Then
            MsgBox_80EEA = MsgBox_80EEA & "* Interest at Sr. No  " & i & "  in schedule 80EEA should be Numeric value" & Chr(13)
            ValidateIntrst_80EEA = False
            Exit Function
        End If
        
        If Intrst_80EEA(i) > 99999999999999# Then
            MsgBox_80EEA = MsgBox_80EEA & "* Interest at Sr. No  " & i & "  in schedule 80EEA cannot exceed 14 digits" & Chr(13)
            ValidateIntrst_80EEA = False
            Exit Function
        End If
        
        'Interest u/s 80EEA should be more than 0
        
        If Intrst_80EEA(i) < 0 Or Intrst_80EEA(i) = 0 Then
            MsgBox_80EEA = MsgBox_80EEA & "* ""Interest u/s 80EEA"" should be more than 0 at Sr. No  " & i & "" & Chr(13)
            ValidateIntrst_80EEA = False
            Exit Function
        End If
         
Next
End Function

Function ValidateBankAccountNumber_80EEA(BankAccountNumber As Variant, CC As Long) As Boolean
    ValidateBankAccountNumber_80EEA = True
    Dim numfound As Boolean
    Dim countnum As Long
    Dim myB(), ValidateIFSC As Variant
    Dim i As Long
    Dim zeroCount As Long
    Dim BeforeZero, AfterZero As String
    errmsgVerification = ""
    numfound = False
    countnum = 0
    BeforeZero = ""
    AfterZero = ""
    zeroCount = 1
    
           
    If Len(BankAccountNumber) > 0 Then
        If Not checkfieldspecialcharacter1(BankAccountNumber) Then
            MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & CC & " is invalid  in schedule 80EEA, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber_80EEA = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & CC & " is invalid  in schedule 80EEA " & Chr(13)
            ValidateBankAccountNumber_80EEA = False
            Exit Function
        End If
    
    End If
  '----------------------------------------------------------------
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        MsgBox_80EEA = MsgBox_80EEA & "*  Please enter the Loan Account number in Bank Details at Sr.No " & CC & " in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If
    
    
    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & CC & " is mandatory in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If
'--------------------------------------------------------------------

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0) Or (InStr(BankAccountNumber, "-/") > 0) Or (InStr(BankAccountNumber, "/-") > 0)) Then
        MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & CC & " is invalid  in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & CC & " is invalid  in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If
    
    If (Not checkfieldspecialcharacter(Mid(BankAccountNumber, 1, 1))) Then
        MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & CC & " is invalid  in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If


    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & CC & " is invalid  in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & CC & " is invalid  in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If

    ReDim myB(Len(BankAccountNumber) - 1)
    For i = 1 To Len(BankAccountNumber)
        myB(i - 1) = Mid(BankAccountNumber, i, 1)
    Next

    For i = LBound(myB) To UBound(myB)
        If IsNumeric(myB(i)) Then
            countnum = countnum + 1
        End If

        If i > LBound(myB) And i < UBound(myB) Then
            If myB(i) = 0 Then
                If myB(i - 1) = 0 Then
                    zeroCount = zeroCount + 1
                    AfterZero = IIf(Not IsNumeric(myB(i + 1)), myB(i + 1), "")
                Else
                    BeforeZero = IIf(Not IsNumeric(myB(i - 1)), myB(i - 1), "")
                End If
            End If
        End If

    Next

    If BeforeZero <> "" And AfterZero <> "" Then
        If zeroCount > 1 Then
            MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & CC & " is invalid  in schedule 80EEA" & Chr(13)
            ValidateBankAccountNumber_80EEA = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "* Loan Account number at Sr.No " & CC & "  in schedule 80EEA is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If
End Function


Public Function ChkMaxDate_80EEA(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

     ChkMaxDate_80EEA = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

        If Year > 2022 Then
            ChkMaxDate_80EEA = False
            Exit Function
        Else
                If Year = 2022 Then
                If month > 4 Then
                    ChkMaxDate_80EEA = False
                    Exit Function
                Else
                    If month = 4 Then
                        If dat > 1 Then
                           ChkMaxDate_80EEA = False
                            Exit Function
                        Else
                            If dat = 1 Then
                               ChkMaxDate_80EEA = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function

