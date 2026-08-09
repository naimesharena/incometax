Attribute VB_Name = "md80EEB"
Public MsgBox_80EEB As Variant
Dim end_80EEB As Variant
Dim end_80EEBLoanfrm As Variant
Dim end_80EEBIFSC As Variant
Dim end_80EEBbankName As Variant
Dim end_80EEBPAN As Variant
Dim end_80EEBAccntNum As Variant
Dim end_80EEBLoanDate As Variant
Dim end_80EEBLoanAmt As Variant
Dim end_80EEBLoanOutstanding As Variant
Dim end_80EEBVehicleValue As Variant
Dim end_80EEBVehicleReg As Variant
Dim end_80EEBIntrst As Variant

Sub ValidateSheet80EEB_Click()
Dim vbMessgaeCaption As String

vbMessgaeCaption = "ITR 2: AY: 2025-26"
Validate80EEB_All


fmsgboxStatus "Schedule 80EEB is OK"
End Sub
Sub Validate80EEB_All()
Dim vbMessgaeCaption As String
vbMessgaeCaption = "Error"
   Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("80E_80EE_80EEA_80EEB")
If Not Validate_80EEB Then
    sourceSheet.Activate
    fmsgboxStatus MsgBox_80EEB
    CloseMsg
End If
End Sub

Function Validate_80EEB()
Validate_80EEB = True
MsgBox_80EEB = ""
end_80EEB = 0

setTblinfo_80EEBLoanfrm
'setTblinfo_80EEBIFSC
'setTblinfo_80EEBPAN
setTblinfo_80EEBBankName
setTblinfo_80EEBAccntNum
setTblinfo_80EEBLoanDate
setTblinfo_80EEBLoanAmt
setTblinfo_80EEBLoanOutstanding
'setTblinfo_80EEBVehicleValue
setTblinfo_80EEBVehicleReg
setTblinfo_80EEBIntrst


'end_80EEB = WorksheetFunction.Max(0, end_80EEB, end_80EEBLoanfrm, end_80EEBbankName, end_80EEBAccntNum, end_80EEBLoanDate, end_80EEBLoanAmt, end_80EEBLoanOutstanding, end_80EEBVehicleValue, end_80EEBVehicleReg, end_80EEBIntrst)
end_80EEB = WorksheetFunction.Max(0, end_80EEBLoanfrm, end_80EEBbankName, end_80EEBAccntNum, end_80EEBLoanDate, end_80EEBLoanAmt, end_80EEBLoanOutstanding, end_80EEBVehicleReg, end_80EEBIntrst)

If Not ValidateLoanfrm_80EEB Then Validate_80EEB = False
If Not ValidateBankName_80EEB Then Validate_80EEB = False
If Not ValidateAccntNum_80EEB Then Validate_80EEB = False
If Not ValidateLoanDate_80EEB Then Validate_80EEB = False
If Not ValidateLoanAmt_80EEB Then Validate_80EEB = False

'    If (Sheet20.Range("MaxLoan_80EEB").value) > 3500000 Then
'        MsgBox_80EEB = MsgBox_80EEB & "* ""Total loan taken"" shall not be more than Rs. 35 Lakhs in schedule 80EE" & Chr(13)
'        Validate_80EEB = False
'    End If

If Not ValidateLoanOutstanding_80EEB Then Validate_80EEB = False
'If Not ValidateVehicleValue_80EEB Then Validate_80EEB = False
If Not ValidateVehicleReg_80EEB Then Validate_80EEB = False
If Not ValidateIntrst_80EEB Then Validate_80EEB = False
'If Not ValidateLoanfrm_80EEB Then Validate_80EEB = False


    
    If Len(Sheet60.Range("TotAmt.80EEB").value) > 14 Then
        MsgBox_80EEB = MsgBox_80EEB & "* Total of Interest u/s 80EEB in schedule 80EEB cannot exceed 14 digits" & Chr(13)
        Validate_80EEB = False
    End If

End Function

Sub setTblinfo_80EEBLoanfrm()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("LoanfrmBankOrInstitute.80EEB").count
    Set rangecells = Range("LoanfrmBankOrInstitute.80EEB").Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
    end_80EEBLoanfrm = ccount
End Sub




Sub setTblinfo_80EEBBankName()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("bankName.80EEB").count
    Set rangecells = Range("bankName.80EEB").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEBbankName = ccount
End Sub




Sub setTblinfo_80EEBAccntNum()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanAccNum.80EEB").count
    Set rangecells = Range("loanAccNum.80EEB").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEBAccntNum = ccount
End Sub
Sub setTblinfo_80EEBLoanDate()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanDate.80EEB").count
    Set rangecells = Range("loanDate.80EEB").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEBLoanDate = ccount
End Sub
Sub setTblinfo_80EEBLoanAmt()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanAmt.80EEB").count
    Set rangecells = Range("loanAmt.80EEB").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEBLoanAmt = ccount
End Sub

Sub setTblinfo_80EEBLoanOutstanding()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanOutstanding.80EEB").count
    Set rangecells = Range("loanOutstanding.80EEB").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEBLoanOutstanding = ccount
End Sub

Sub setTblinfo_80EEBVehicleReg()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("VehicleRegNum.80EEB").count
    Set rangecells = Range("VehicleRegNum.80EEB").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEBVehicleReg = ccount
End Sub

Sub setTblinfo_80EEBIntrst()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Intrst.80EEB").count
    Set rangecells = Range("Intrst.80EEB").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEBIntrst = ccount
End Sub

Function ValidateLoanfrm_80EEB() As Boolean
    ValidateLoanfrm_80EEB = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("LoanfrmBankOrInstitute.80EEB").Cells
    ReDim Loanfrm_80EEB(end_80EEB)
    For i = 1 To end_80EEB
        Loanfrm_80EEB(i) = rangecells.item(i).value
       If isdropdownblank(Loanfrm_80EEB(i)) Then
           '  MsgBox_80EEB = MsgBox_80EEB + "* Please select dropdown ""Loan taken from"" in schedule 80EEB at Sr. No " & i & "" & Chr(13)
              MsgBox_80EEB = MsgBox_80EEB + "* Please select dropdown ""Loan taken from"" in schedule 80EEB at Sr. No " & i & "" & Chr(13)
            ValidateLoanfrm_80EEB = False
            Exit Function
        End If

         
Next
End Function

Function ValidateBankName_80EEB() As Boolean
    ValidateBankName_80EEB = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("bankName.80EEB").Cells
    ReDim BankName_80EEB(end_80EEB)
    For i = 1 To end_80EEB
        BankName_80EEB(i) = rangecells.item(i).value
        If Not chkCompulsory(BankName_80EEB(i)) Then
             MsgBox_80EEB = MsgBox_80EEB + "* ""Please provide Name of the Bank/Institution from which the loan is taken in schedule 80EEB"" at Sr. No " & i & "" & Chr(13)
            ValidateBankName_80EEB = False
            Exit Function
        End If
         If Len(BankName_80EEB(i)) > 125 Then
          MsgBox_80EEB = MsgBox_80EEB + "* Name of the Bank/Institution at Sr. No " & i & " in Sheet 80EEB less than 125 characters." & Chr(13)
            ValidateBankName_80EEB = False
            Exit Function
        End If
         If Not checkfieldSuperSpecialcharacter(BankName_80EEB(i)) Then
             MsgBox_80EEB = MsgBox_80EEB + "* Name of the Bank/Institution from which the loan is taken in schedule 80EEB at Sr.no. " & i & " should not Contain <, >, characters." & Chr(13)
            ValidateBankName_80EEB = False
            Exit Function
        End If
Next
End Function

Function ValidateAccntNum_80EEB() As Boolean
    ValidateAccntNum_80EEB = True

    Dim rangecells, rangecells1 As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("loanAccNum.80EEB").Cells
    Set rangecells1 = Sheet60.Range("LoanfrmBankOrInstitute.80EEB").Cells
    
    ReDim AccntNum_80EEB(end_80EEB)
     ReDim Bankorothebank_80EEB(end_80EEB)
     
    For i = 1 To end_80EEB
        AccntNum_80EEB(i) = rangecells.item(i).value
        Bankorothebank_80EEB(i) = rangecells1.item(i).value
         
        If Not chkCompulsory(AccntNum_80EEB(i)) Then
          MsgBox_80EEB = MsgBox_80EEB + "* ""Please provide Loan Account number of the Bank/Institution from which loan is taken in schedule 80EEB"" at Sr. No " & i & "" & Chr(13)
            ValidateAccntNum_80EEB = False
            Exit Function
        End If
         If Len(AccntNum_80EEB(i)) > 20 Then
            MsgBox_80EEB = MsgBox_80EEB + "* Loan Account number in Sheet 80EEB should be less than or equal to 20 characters at Sr. No " & i & "." & Chr(13)
            ValidateAccntNum_80EEB = False
            Exit Function
        End If
        
        If Bankorothebank_80EEB(i) = "Bank" Then
        If Not ValidateBankAccountNumber_80EEB(AccntNum_80EEB(i), i) Then
        'Msgbox_BA = Msgbox_BA + "Account Number at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
        ValidateAccntNum_80EEB = False
        Exit Function
    End If
    End If
    
    
    If Bankorothebank_80EEB(i) = "Institution" Then
     If Not checkfieldspecialcharacter_BthacntReferencenumber(AccntNum_80EEB(i)) Then
            MsgBox_80EEB = MsgBox_80EEB + "* Loan Account number of the Bank/Institution at Sr.No " & i & " is invalid in schedule 80EEB, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateAccntNum_80EEB = False
            Exit Function
     End If
    End If
         
Next
End Function

Function ValidateVehicleReg_80EEB() As Boolean
    ValidateVehicleReg_80EEB = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("VehicleRegNum.80EEB").Cells
    ReDim VehicleReg_80EEB(end_80EEB)
    For i = 1 To end_80EEB
        VehicleReg_80EEB(i) = rangecells.item(i).value
        If Not chkCompulsory(VehicleReg_80EEB(i)) Then
             MsgBox_80EEB = MsgBox_80EEB + "* ""Please enter vehicle Registration number in schedule 80EEB"" at Sr. No " & i & "" & Chr(13)
            ValidateVehicleReg_80EEB = False
            Exit Function
        End If
         If Len(VehicleReg_80EEB(i)) > 11 Then
          MsgBox_80EEB = MsgBox_80EEB + "*  ""Vehicle Registration number cannot exceed 11 characters in schedule 80EEB"" at Sr. No " & i & "" & Chr(13)
            ValidateVehicleReg_80EEB = False
            Exit Function
        End If
    
    'CheckSpecialCharacter_New_1
        
        'If Not CheckSpecialCharacter(VehicleReg_80EEB(i)) Then
        
'SIT-95387 BY SAI ON 26/06/2025
        If Not CheckSpecialCharacter_New_1(VehicleReg_80EEB(i)) Then
         MsgBox_80EEB = MsgBox_80EEB + "* Vehicle Registration number should be alphanumeic at Sr. No " & i & " in Sheet 80EEB " & Chr(13)
            ValidateVehicleReg_80EEB = False
            Exit Function
        End If
        
         
Next
End Function

Function ValidateLoanDate_80EEB() As Boolean
    ValidateLoanDate_80EEB = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("loanDate.80EEB").Cells
    ReDim LoanDate_80EEB(end_80EEB)
    For i = 1 To end_80EEB
        LoanDate_80EEB(i) = rangecells.item(i).value
        If Not chkCompulsory(LoanDate_80EEB(i)) Then
        
             MsgBox_80EEB = MsgBox_80EEB + "* ""Please provide Date of sanction of loan in schedule 80EEB"" at Sr. No " & i & "" & Chr(13)
            ValidateLoanDate_80EEB = False
            Exit Function
        End If
         If Len(LoanDate_80EEB(i)) > 10 Then
        
          MsgBox_80EEB = MsgBox_80EEB + "* Date of sanction of loan at Sr. No " & i & " in schedule 80EEB less than 10 characters." & Chr(13)
            ValidateLoanDate_80EEB = False
            Exit Function
        End If
        'CheckDateddmmyyyy
         If Not CheckDateddmmyyyy(LoanDate_80EEB(i)) Then
         MsgBox_80EEB = MsgBox_80EEB + "* ""Please enter date in valid format in schedule 80EEB"" at Sr. No " & i & "." & Chr(13)
        ValidateLoanDate_80EEB = False
                 

          Exit Function
        End If
        
        If Not ChkMaxDate_80EEB(Trim(LoanDate_80EEB(i)), "31-03-2023") Then
          MsgBox_80EEB = MsgBox_80EEB + "* ""Date of sanction of Loan shall be between 01/04/2019 to 31/03/2023"" in schedule 80EEB at Sr. No " & i & "." & Chr(13)
         ValidateLoanDate_80EEB = False
          Exit Function
        End If
                 
      If Not ChkMinInclusiveDate(Trim(Dformat(LoanDate_80EEB(i), "yyyy-mm-dd")), "2019-04-01") Then
             MsgBox_80EEB = MsgBox_80EEB + "* ""Date of sanction of Loan shall be between 01/04/2019 to 31/03/2023"" in schedule 80EEB at Sr. No " & i & "." & Chr(13)
            ValidateLoanDate_80EEB = False
                 

          Exit Function
        End If
   
Next
End Function

Function ValidateLoanAmt_80EEB() As Boolean
    ValidateLoanAmt_80EEB = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("loanAmt.80EEB").Cells
    ReDim LoanAmt_80EEB(end_80EEB)
    For i = 1 To end_80EEB
        LoanAmt_80EEB(i) = rangecells.item(i).value
        If Not chkCompulsory(LoanAmt_80EEB(i)) Then
             MsgBox_80EEB = MsgBox_80EEB + "* ""Please provide Total amount of Loan in schedule 80EEB"" at Sr. No " & i & "" & Chr(13)
            ValidateLoanAmt_80EEB = False
            Exit Function
        End If
        
        If Not IsNumeric(LoanAmt_80EEB(i)) Then
            MsgBox_80EEB = MsgBox_80EEB & "* Loan amount at Sr. No  " & i & " in schedule 80EEB should be Numeric value" & Chr(13)
            ValidateLoanAmt_80EEB = False
            Exit Function
        End If
        
        If LoanAmt_80EEB(i) > 99999999999999# Then
            MsgBox_80EEB = MsgBox_80EEB & "* Loan amount at Sr. No  " & i & " in schedule 80EEB cannot exceed 14 digits" & Chr(13)
            ValidateLoanAmt_80EEB = False
            Exit Function
        End If
        
        If LoanAmt_80EEB(i) < 0 Or LoanAmt_80EEB(i) = 0 Then
          '  MsgBox_80EEB = MsgBox_80EEB & "* Total amount of Loan should be more than 0 in schedule 80EEB at Sr. No  " & i & "" & Chr(13)
           MsgBox_80EEB = MsgBox_80EEB & "* Total loan taken should be more than 0 in schedule 80EEB at Sr. No  " & i & "" & Chr(13)
            ValidateLoanAmt_80EEB = False
            Exit Function
        End If
         
Next
End Function


Function ValidateLoanOutstanding_80EEB() As Boolean
    ValidateLoanOutstanding_80EEB = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("loanOutstanding.80EEB").Cells
    ReDim LoanOutStanding_80EEB(end_80EEB)
    For i = 1 To end_80EEB
        LoanOutStanding_80EEB(i) = rangecells.item(i).value
        If Not chkCompulsory(LoanOutStanding_80EEB(i)) Then
           '  MsgBox_80EEB = MsgBox_80EEB + "* ""Loan outstanding as on last date of financial year is mandatory in schedule 80EEB"" at Sr. No " & i & "" & Chr(13)
             MsgBox_80EEB = MsgBox_80EEB + "* ""Loan outstanding as on  last date of financial year is mandatory in schedule 80EEB. Please enter 0 in case the entire loan is repaid during the year""at Sr. No " & i & "" & Chr(13)
            ValidateLoanOutstanding_80EEB = False
            Exit Function
        End If

        
        If Not IsNumeric(LoanOutStanding_80EEB(i)) Then
            MsgBox_80EEB = MsgBox_80EEB & "* Loan outstanding at Sr. No  " & i & " in schedule 80EEB should be Numeric value" & Chr(13)
            ValidateLoanOutstanding_80EEB = False
            Exit Function
        End If
        
        If LoanOutStanding_80EEB(i) > 99999999999999# Then
            MsgBox_80EEB = MsgBox_80EEB & "* Loan outstanding at Sr. No  " & i & " in schedule 80EEB cannot exceed 14 digits" & Chr(13)
            ValidateLoanOutstanding_80EEB = False
            Exit Function
        End If
        
'        If LoanOutStanding_80EEB(i) < 0 Then
'            MsgBox_80EEB = MsgBox_80EEB & "* Loan outstanding as on last date of financial year can't be less than 0 in schedule 80EEB at Sr. No  " & i & ". You may please enter as 0 if it become negative as result of excess payment." & Chr(13)
'            ValidateLoanOutstanding_80EEB = False
'            Exit Function
'        End If
         
Next
End Function

Function ValidateIntrst_80EEB() As Boolean
    ValidateIntrst_80EEB = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("Intrst.80EEB").Cells
    ReDim Intrst_80EEB(end_80EEB)
    For i = 1 To end_80EEB
        Intrst_80EEB(i) = rangecells.item(i).value
        If Not chkCompulsory(Intrst_80EEB(i)) Then
             MsgBox_80EEB = MsgBox_80EEB + "* ""Please provide Interest u/s 80EEB"" at Sr. No " & i & "" & Chr(13)
            ValidateIntrst_80EEB = False
            Exit Function
        End If


        If Not IsNumeric(Intrst_80EEB(i)) Then
            MsgBox_80EEB = MsgBox_80EEB & "* Interest at Sr. No  " & i & " in schedule 80EEB should be Numeric value" & Chr(13)
            ValidateIntrst_80EEB = False
            Exit Function
        End If
        
        If Intrst_80EEB(i) > 99999999999999# Then
            MsgBox_80EEB = MsgBox_80EEB & "* Interest at Sr. No  " & i & " in schedule 80EEB cannot exceed 14 digits" & Chr(13)
            ValidateIntrst_80EEB = False
            Exit Function
        End If
        
        'Interest u/s 80EEB should be more than 0 in schedule 80EEB
        
        If Intrst_80EEB(i) < 0 Or Intrst_80EEB(i) = 0 Then
            MsgBox_80EEB = MsgBox_80EEB & "* Interest u/s 80EEB should be more than 0 in schedule 80EEB at Sr. No  " & i & "" & Chr(13)
            ValidateIntrst_80EEB = False
            Exit Function
        End If
         
Next
End Function


Function ValidateBankAccountNumber_80EEB(BankAccountNumber As Variant, cc As Long) As Boolean
    ValidateBankAccountNumber_80EEB = True
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
    
   ' BankAccountNumber = Sheet5.Range("IncD.BankAccountNumber")

           
    If Len(BankAccountNumber) > 0 Then
        
        If Not checkfieldspecialcharacter80E_EE_EEA_EEB(BankAccountNumber) Then
            MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EEB, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber_80EEB = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EEB " & Chr(13)
            ValidateBankAccountNumber_80EEB = False
            Exit Function
        End If
    End If
  '----------------------------------------------------------------
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        MsgBox_80EEB = MsgBox_80EEB & "*  Please enter the Loan Account number in Bank Details at Sr.No " & cc & " in schedule 80EEB" & Chr(13)
        ValidateBankAccountNumber_80EEB = False
        Exit Function
    End If
    
    
    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc & " is mandatory in schedule 80EEB" & Chr(13)
        ValidateBankAccountNumber_80EEB = False
        Exit Function
    End If
'--------------------------------------------------------------------

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0) Or (InStr(BankAccountNumber, "-/") > 0) Or (InStr(BankAccountNumber, "/-") > 0)) Then
        MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 80EEB" & Chr(13)
        ValidateBankAccountNumber_80EEB = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EEB" & Chr(13)
        ValidateBankAccountNumber_80EEB = False
        Exit Function
    End If
    
    If (Not checkfieldspecialcharacter(Mid(BankAccountNumber, 1, 1))) Then
        MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EEB" & Chr(13)
        ValidateBankAccountNumber_80EEB = False
        Exit Function
    End If


    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EEB" & Chr(13)
        ValidateBankAccountNumber_80EEB = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EEB" & Chr(13)
        ValidateBankAccountNumber_80EEB = False
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
            MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc + 1 & " is invalid in schedule 80EEB" & Chr(13)
            ValidateBankAccountNumber_80EEB = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "* Loan Account number at Sr.No " & cc + 1 & "  in schedule 80EEB is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber_80EEB = False
        Exit Function
    End If
End Function


Public Function ChkMaxDate_80EEB(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

     ChkMaxDate_80EEB = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))


        If Year > 2023 Then
            ChkMaxDate_80EEB = False
            Exit Function
        Else

                If Year = 2023 Then
                If month > 4 Then
                    ChkMaxDate_80EEB = False
                    Exit Function
                Else
                    If month = 4 Then
                        If dat > 1 Then
                           ChkMaxDate_80EEB = False
                            Exit Function
                        Else
                            If dat = 1 Then
                               ChkMaxDate_80EEB = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function


Sub CmdSchedule80EEBAdd_Click()
Dim vRows As Long
    Sheets("80E_80EE_80EEA_80EEB").Activate
    EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEB||bankName.80EEB||loanAccNum.80EEB||loanDate.80EEB||loanAmt.80EEA||loanOutstanding.80EEB||VehicleRegNum.80EEB||Intrst.80EEB"
    ActiveCellRange = EfilingCommon.searchLastRow("LoanfrmBankOrInstitute.80EEB")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub




