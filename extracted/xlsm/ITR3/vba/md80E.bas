Attribute VB_Name = "md80E"
Public MsgBox_80E As Variant
Dim end_80E As Variant
Dim end_80ELoanfrm As Variant
Dim end_80EIFSC As Variant
Dim end_80EbankName As Variant
Dim end_80EPAN As Variant
Dim end_80EAccntNum As Variant
Dim end_80ELoanDate As Variant
Dim end_80ELoanAmt As Variant
Dim end_80ELoanOutstanding As Variant
Dim end_80EIntrst As Variant

Sub ValidateSheet80E_Click()
Dim vbMessgaeCaption As String

vbMessgaeCaption = "ITR 3: AY: 2025-26"
Validate80E_All

fmsgboxStatus "Schedule 80E is OK"

End Sub
Sub Validate80E_All()
Dim vbMessgaeCaption As String
vbMessgaeCaption = "Error"
   Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("80E_80EE_80EEA_80EEB")
If Not Validate_80E Then
    sourceSheet.Activate
    fmsgbox (MsgBox_80E)
    CloseMsg
End If
End Sub
Function Validate_80E()
Validate_80E = True

MsgBox_80E = ""

setTblinfo_80ELoanfrm
setTblinfo_80EBankName
setTblinfo_80EAccntNum
setTblinfo_80ELoanDate
setTblinfo_80ELoanAmt
setTblinfo_80ELoanOutstanding
setTblinfo_80EIntrst


end_80E = WorksheetFunction.Max(0, end_80ELoanfrm, end_80EbankName, , end_80EAccntNum, end_80ELoanDate, end_80ELoanAmt, end_80ELoanOutstanding, end_80EIntrst)

If Not ValidateLoanfrm_80E Then Validate_80E = False
If Not ValidateBankName_80E Then Validate_80E = False
If Not ValidateAccntNum_80E Then Validate_80E = False
If Not ValidateLoanDate_80E Then Validate_80E = False
If Not ValidateLoanAmt_80E Then Validate_80E = False
If Not ValidateLoanOutstanding_80E Then Validate_80E = False
If Not ValidateIntrst_80E Then Validate_80E = False
If Len(Sheet60.Range("TotAmt.80E").value) > 14 Then
        MsgBox_80E = MsgBox_80E & "* Total of Interest u/s 80E in schedule 80E cannot exceed 14 digits" & Chr(13)
        Validate_80E = False
    End If
End Function
Sub setTblinfo_80ELoanfrm()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("LoanfrmBankOrInstitute.80E").count
    Set rangecells = Range("LoanfrmBankOrInstitute.80E").Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
    end_80ELoanfrm = ccount
End Sub

Sub setTblinfo_80EBankName()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("bankName.80E").count
    Set rangecells = Range("bankName.80E").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EbankName = ccount
End Sub

Sub setTblinfo_80EAccntNum()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanAccNum.80E").count
    Set rangecells = Range("loanAccNum.80E").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EAccntNum = ccount
End Sub

Sub setTblinfo_80ELoanDate()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanDate.80E").count
    Set rangecells = Range("loanDate.80E").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80ELoanDate = ccount
End Sub
Sub setTblinfo_80ELoanAmt()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanAmt.80E").count
    Set rangecells = Range("loanAmt.80E").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80ELoanAmt = ccount
End Sub

Sub setTblinfo_80ELoanOutstanding()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanOutstanding.80E").count
    Set rangecells = Range("loanOutstanding.80E").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80ELoanOutstanding = ccount
End Sub
Sub setTblinfo_80EIntrst()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Intrst.80E").count
    Set rangecells = Range("Intrst.80E").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EIntrst = ccount
End Sub

Function ValidateLoanfrm_80E() As Boolean
    ValidateLoanfrm_80E = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("LoanfrmBankOrInstitute.80E").Cells
    ReDim Loanfrm_80E(end_80E)
    For i = 1 To end_80E
        Loanfrm_80E(i) = rangecells.item(i).value
        If isdropdownblank(Loanfrm_80E(i)) Then
             MsgBox_80E = MsgBox_80E + "* Please select dropdown from ""Loan taken from"" in 80E schedule at Sr. No." & i & "" & Chr(13)
            ValidateLoanfrm_80E = False
            Exit Function
        End If
Next
End Function
Function ValidateBankName_80E() As Boolean
    ValidateBankName_80E = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("bankName.80E").Cells
    ReDim BankName_80E(end_80E)
    For i = 1 To end_80E
        BankName_80E(i) = rangecells.item(i).value
        If Not chkCompulsory(BankName_80E(i)) Then
             MsgBox_80E = MsgBox_80E + "* ""Please provide Name of the Bank/ Institution from which the loan is taken"" at Sr. No." & i & "" & Chr(13)
            ValidateBankName_80E = False
            Exit Function
        End If
         If Len(BankName_80E(i)) > 125 Then
          MsgBox_80E = MsgBox_80E + "* Name of the Bank/ Institution from which the loan is taken in Schedule 80E at Sr. No." & i & " should be less than or equal to 125 characters." & Chr(13)
            ValidateBankName_80E = False
            Exit Function
        End If
        
        If Not checkfieldSuperSpecialcharacter(BankName_80E(i)) Then
             MsgBox_80E = MsgBox_80E + "* Name of the Bank/ Institution from which the loan is taken in schedule 80E at Sl.no. " & i & " should not Contain <, >, characters." & Chr(13)
            ValidateBankName_80E = False
            Exit Function
        End If
         
Next
End Function

Function ValidateAccntNum_80E() As Boolean
    ValidateAccntNum_80E = True

    Dim rangecells As Range
    Dim rangeceels1 As Range
    
    Dim i As Long
    Set rangecells = Sheet60.Range("loanAccNum.80E").Cells
    Set rangecells1 = Sheet60.Range("LoanfrmBankOrInstitute.80E").Cells
    ReDim AccntNum_80E(end_80E)
    ReDim BankorInst_80E(end_80E)
    For i = 1 To end_80E
        AccntNum_80E(i) = rangecells.item(i).value
        BankorInst_80E(i) = rangecells1.item(i).value
        If Not chkCompulsory(AccntNum_80E(i)) Then
              MsgBox_80E = MsgBox_80E + "* ""Please provide Loan Account number of the Bank / Institution from which loan is taken or any reference number if loan is taken from an Institution which does not provide loan account number"" at Sr. No." & i & "" & Chr(13)
            
            ValidateAccntNum_80E = False
            Exit Function
        End If
         If Len(AccntNum_80E(i)) > 20 Then
          MsgBox_80E = MsgBox_80E + "* Loan Account number  at Sr. No " & i & " in Schedule 80E less than 20 characters." & Chr(13)
            ValidateAccntNum_80E = False
            Exit Function
        End If
        
        If BankorInst_80E(i) = "Bank" Then
        If Not ValidateBankAccountNumber_80E(AccntNum_80E(i), i) Then
        'Msgbox_BA = Msgbox_BA + "Account Number at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
        ValidateAccntNum_80E = False
        Exit Function
    End If
    End If
    
    
    If BankorInst_80E(i) = "Institution" Then
        If Not checkfieldspecialcharacter1(AccntNum_80E(i)) Then
            MsgBox_80E = MsgBox_80E & "* Loan Account number of the Bank/Institution at Sr.No " & i & " is invalid  in schedule 80E, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
           ValidateAccntNum_80E = False
            Exit Function
        End If
        End If
         
Next
End Function

Function ValidateLoanDate_80E() As Boolean
    ValidateLoanDate_80E = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("loanDate.80E").Cells
    ReDim LoanDate_80E(end_80E)
    For i = 1 To end_80E
        LoanDate_80E(i) = rangecells.item(i).value
        If Not chkCompulsory(LoanDate_80E(i)) Then
              MsgBox_80E = MsgBox_80E + "* ""Please provide Date of sanction of Loan"" at Sr. No." & i & "" & Chr(13)
            ValidateLoanDate_80E = False
            Exit Function
        End If
         If Len(LoanDate_80E(i)) > 10 Then
          MsgBox_80E = MsgBox_80E + "* Date of sanction of Loan at Sr. No " & i & " in schedule 80E less than 10 characters." & Chr(13)
            ValidateLoanDate_80E = False
            Exit Function
        End If
        'CheckDateddmmyyyy
         If Not CheckDateddmmyyyy(LoanDate_80E(i)) Then
         MsgBox_80E = MsgBox_80E + "* ""Please enter date in valid format"" at Sr. No " & i & "." & Chr(13)
        ValidateLoanDate_80E = False
                 

          Exit Function
        End If
        
        If Not ChkMaxDate_80E(Trim(LoanDate_80E(i)), "31-03-2026") Then
            MsgBox_80E = MsgBox_80E + "* Date can not be after 31/03/2026 at Sr. No " & i & "." & Chr(13)
            ValidateLoanDate_80E = False
                 

          Exit Function
        End If
Next
End Function
Function ValidateLoanAmt_80E() As Boolean
    ValidateLoanAmt_80E = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("loanAmt.80E").Cells
    ReDim LoanAmt_80E(end_80E)
    For i = 1 To end_80E
        LoanAmt_80E(i) = rangecells.item(i).value
        If Not chkCompulsory(LoanAmt_80E(i)) Then
              MsgBox_80E = MsgBox_80E + "* ""Please provide Total amount of loan"" at Sr. No." & i & "" & Chr(13)
            ValidateLoanAmt_80E = False
            Exit Function
        End If
        
        If Not IsNumeric(LoanAmt_80E(i)) Then
            MsgBox_80E = MsgBox_80E & "* Loan amount at Sr. No." & i & " in schedule 80E should be Numeric value" & Chr(13)
            ValidateLoanAmt_80E = False
            Exit Function
        End If
        
        If LoanAmt_80E(i) > 99999999999999# Then
            MsgBox_80E = MsgBox_80E & "* Loan amount at Sr. No." & i & " in schedule 80E cannot exceed 14 digits" & Chr(13)
            ValidateLoanAmt_80E = False
            Exit Function
        End If
        
        If LoanAmt_80E(i) < 0 Or LoanAmt_80E(i) = 0 Then
            MsgBox_80E = MsgBox_80E & "* ""Total amount of loan should be more than 0"" at Sr. No." & i & "" & Chr(13)
            ValidateLoanAmt_80E = False
            Exit Function
        End If
         
Next
End Function

Function ValidateLoanOutstanding_80E() As Boolean
    ValidateLoanOutstanding_80E = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("loanOutstanding.80E").Cells
    ReDim LoanOutStanding_80E(end_80E)
    For i = 1 To end_80E
        LoanOutStanding_80E(i) = rangecells.item(i).value
        If Not chkCompulsory(LoanOutStanding_80E(i)) Then
                MsgBox_80E = MsgBox_80E + "* ""Loan outstanding as on last date of finacial year is mandatory in schedule 80E. Please enter 0 in case the entire loan is repaid during the year"" at Sr. No." & i & "" & Chr(13)
            ValidateLoanOutstanding_80E = False
            Exit Function
        End If
        If Not IsNumeric(LoanOutStanding_80E(i)) Then
            MsgBox_80E = MsgBox_80E & "* Loan outstanding at Sr. No." & i & "  in schedule 80E should be Numeric value" & Chr(13)
            ValidateLoanOutstanding_80E = False
            Exit Function
        End If
        
        If LoanOutStanding_80E(i) > 99999999999999# Then
            MsgBox_80E = MsgBox_80E & "* Loan outstanding at Sr. No." & i & " in schedule 80E cannot exceed 14 digits" & Chr(13)
            ValidateLoanOutstanding_80E = False
            Exit Function
        End If
        
        If LoanOutStanding_80E(i) < 0 Then
            MsgBox_80E = MsgBox_80E & "* Loan outstanding as on last date of financial year can't be less than 0 in schedule 80E at Sr. No." & i & ". You may please enter as 0 if it become negative as result of excess payment." & Chr(13)
            ValidateLoanOutstanding_80E = False
            Exit Function
        End If
         
Next
End Function

Function ValidateIntrst_80E() As Boolean
    ValidateIntrst_80E = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("Intrst.80E").Cells
    ReDim Intrst_80E(end_80E)
    For i = 1 To end_80E
        Intrst_80E(i) = rangecells.item(i).value
        If Not chkCompulsory(Intrst_80E(i)) Then
             MsgBox_80E = MsgBox_80E + "* ""Please provide Interest u/s 80E"" at Sr. No." & i & "" & Chr(13)
            ValidateIntrst_80E = False
            Exit Function
        End If
        If Not IsNumeric(Intrst_80E(i)) Then
            MsgBox_80E = MsgBox_80E & "* Interest at Sr. No." & i & "  in schedule 80E should be Numeric value" & Chr(13)
            ValidateIntrst_80E = False
            Exit Function
        End If
        
        If Intrst_80E(i) > 99999999999999# Then
            MsgBox_80E = MsgBox_80E & "* Interest at Sr. No." & i & "  in schedule 80E cannot exceed 14 digits" & Chr(13)
            ValidateIntrst_80E = False
            Exit Function
        End If
        
        'Interest u/s 80E should be more than 0 in schedule 80E
        
        If Intrst_80E(i) < 0 Or Intrst_80E(i) = 0 Then
            MsgBox_80E = MsgBox_80E & "* Interest u/s 80E should be more than 0 at Sr. No." & i & "" & Chr(13)
            ValidateIntrst_80E = False
            Exit Function
        End If
         
Next
End Function

Function ValidateBankAccountNumber_80E(BankAccountNumber As Variant, CC As Long) As Boolean
    ValidateBankAccountNumber_80E = True
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
        If Not checkfieldspecialcharacter1(BankAccountNumber) Then
            MsgBox_80E = MsgBox_80E & "* Loan Account number at Sr.No " & CC & " is invalid  in schedule 80E, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber_80E = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            MsgBox_80E = MsgBox_80E & "* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from an Institution which does not provide loan account number""" & Chr(13)
            ValidateBankAccountNumber_80E = False
            Exit Function
        End If
    
    End If
  '----------------------------------------------------------------
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        MsgBox_80E = MsgBox_80E & "*  Please enter the Loan Account number in Bank Details at Sr.No " & CC & " in schedule 80E" & Chr(13)
        ValidateBankAccountNumber_80E = False
        Exit Function
    End If
    
    
    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        MsgBox_80E = MsgBox_80E & "* Loan Account number at Sr.No " & CC & " is mandatory in schedule 80E" & Chr(13)
        ValidateBankAccountNumber_80E = False
        Exit Function
    End If
'--------------------------------------------------------------------

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0) Or (InStr(BankAccountNumber, "-/") > 0) Or (InStr(BankAccountNumber, "/-") > 0)) Then
        MsgBox_80E = MsgBox_80E & "* Loan Account number at Sr.No " & CC & " is invalid  in schedule 80E" & Chr(13)
        ValidateBankAccountNumber_80E = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        MsgBox_80E = MsgBox_80E & "* ""Please provide correct Loan Account number of the Bank/Institution or any reference number if loan is taken from an Institution which does not provide loan account number""" & Chr(13)
        ValidateBankAccountNumber_80E = False
        Exit Function
    End If
    
    If (Not checkfieldspecialcharacter(Mid(BankAccountNumber, 1, 1))) Then
        MsgBox_80E = MsgBox_80E & "* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from an Institution which does not provide loan account number""" & Chr(13)
        ValidateBankAccountNumber_80E = False
        Exit Function
    End If

''
    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        MsgBox_80E = MsgBox_80E & "* ""Please provide correct Loan Account number of the Bank/Institution  or any reference number if loan is taken from an Institution which does not provide loan account number""" & Chr(13)
        ValidateBankAccountNumber_80E = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        MsgBox_80E = MsgBox_80E & "* ""Please provide correct Loan Account number of the Bank/Institution or any reference number if loan is taken from an Institution which does not provide loan account number""" & Chr(13)
        ValidateBankAccountNumber_80E = False
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
            MsgBox_80E = MsgBox_80E & "* ""Please provide correct Loan Account number of the Bank/Institution or any reference number if loan is taken from an Institution which does not provide loan account number""" & Chr(13)
            ValidateBankAccountNumber_80E = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "* Loan Account number at Sr.No " & CC & "  in schedule 80E is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber_80E = False
        Exit Function
    End If
End Function


Public Function ChkMaxDate_80E(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

    ChkMaxDate_80E = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

        If Year > 2026 Then
            ChkMaxDate_80E = False
            Exit Function
        Else
            If Year = 2026 Then
                If month > 4 Then
                    ChkMaxDate_80E = False
                    Exit Function
                Else
                    If month = 4 Then

                        If dat > 1 Then
                           ChkMaxDate_80E = False
                            Exit Function
                        Else
                            If dat = 1 Then
                               ChkMaxDate_80E = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function
