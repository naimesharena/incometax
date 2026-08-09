Attribute VB_Name = "md80EE"
Public MsgBox_80EE As Variant
Dim end_80EE As Variant
Dim end_80EELoanfrm As Variant
Dim end_80EEIFSC As Variant
Dim end_80EEbankName As Variant
Dim end_80EEPAN As Variant
Dim end_80EEAccntNum As Variant
Dim end_80EELoanDate As Variant
Dim end_80EELoanAmt As Variant
Dim end_80EELoanOutstanding As Variant
Dim end_80EEIntrst As Variant
'Dim end_80EELoanfrm As Variant


Sub ValidateSheet80EE_Click()
Dim vbMessgaeCaption As String

vbMessgaeCaption = "ITR 2: AY: 2025-26"
'Validate80EE_All  'malli comented

If Validate80EE_All Then
Sheet60.Activate


fmsgboxStatus "Schedule 80EE is OK"
CloseMsg
End If
End Sub

'Sub Validate80EE_All()
Function Validate80EE_All()
 Validate80EE_All = True

Dim vbMessgaeCaption As String
vbMessgaeCaption = "Error"
   Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("80E_80EE_80EEA_80EEB")
If Not Validate_80EE Then
    sourceSheet.Activate
   ' MsgBox (MsgBox_80GA), vbOKOnly, vbMessgaeCaption
    fmsgboxStatus MsgBox_80EE
    CloseMsg
End If




''Ankita_18/04
'Dim Data24b, rowcount_80EE, i, j, rowcount_Int24B As Long
'
'rowcount_80EE = getRowNo(Sheet60.Range("LoanfrmBankOrInstitute.80EE").name)
'
'If end_80EE = 1 Then
'   end_80EE = rowcount_80EE
'Else
'   end_80EE = (end_80EE + rowcount_80EE)
'End If
'
'
'setTblinfo_24bLoanfrm
''setTblinfo_24bIFSC
'
'setTblinfo_24bPAN
'setTblinfo_24bBankName
'setTblinfo_24bAccntNum
'setTblinfo_24bLoanDate
'setTblinfo_24bLoanAmt
'setTblinfo_24bLoanOutstanding
'setTblinfo_24bIntrst
'
'Data24b = end_24bLoanfrm > 0 And end_24bbankName > 0 And end_24bAccntNum > 0 And end_24bLoanDate > 0 And end_24bLoanAmt > 0 And end_24bLoanOutstanding > 0 And end_24bIntrst > 0
'
'end_24b = WorksheetFunction.Max(0, end_24b, end_24bLoanfrm, end_24bbankName, , end_24bAccntNum, end_24bLoanDate, end_24bLoanAmt, end_24bLoanOutstanding, end_24bIntrst)
'
'rowcount_Int24B = getRowNo(Sheet16.Range("LoanfrmBankOrInstitute.24b").name)
'
'If end_24b = 1 Then
'   end_24b = rowcount_Int24B
'Else
'   end_24b = (end_24b + rowcount_Int24B)
'End If
'
'
'Dim IFSC_80EE, IFSC_24B
'
'IFSC_80EE = Sheet60.Range("Combination_80EE").Column
'IFSC_24B = Sheet16.Range("Combination_24B").Column
'
'
'If Data24b = True Then
'    For i = rowcount_80EE To end_80EE
'        For j = rowcount_Int24B To end_24b
'            If Sheet60.Cells(i, IFSC_80EE).Value = Sheet16.Cells(j, IFSC_24B).Value Then
'               Exit For
'            End If
'        Next
'
'    Next
'Else
'fmsgboxStatus ("Deduction u/s 80EE can be claimed only if the limit u/s 24(b) is exhausted. So please note that the loan details of respective property under which 80EE is being claimed should be the same as in schedule 24(b)")
'End If
'
'_____________________

'Malli______________________
'Ankita_----------------
Dim rowcount_80EE, i, j, rowcount_Int24B   As Long


'for80EE
setTblinfo_80EELoanfrm
'setTblinfo_80EEIFSC   'Ankita_05/05/2025_Commented as per DESheet_v0.7
'setTblinfo_80EEPAN    'Ankita_05/05/2025_Commented as per DESheet_v0.7
setTblinfo_80EEBankName
setTblinfo_80EEAccntNum
setTblinfo_80EELoanDate
setTblinfo_80EELoanAmt
setTblinfo_80EELoanOutstanding
setTblinfo_80EEIntrst

Dim Data80EE, Data24b, Combination_80EE_24B As Boolean

'Data80EE = end_80EELoanfrm > 0 And end_80EEIFSC > 0 Or end_80EEPAN > 0 And end_80EEbankName > 0 And end_80EEAccntNum > 0 And end_80EELoanDate > 0 And end_80EELoanAmt > 0 And end_80EELoanOutstanding > 0 And end_80EEIntrst > 0
'end_80EE = WorksheetFunction.Max(0, end_80EELoanfrm, end_80EEbankName, end_80EEIFSC, end_80EEPAN, end_80EEAccntNum, end_80EELoanDate, end_80EELoanAmt, end_80EELoanOutstanding, end_80EEIntrst)
'Ankita_09/05/2025_Commented as per DESheet_v0.7
Data80EE = end_80EELoanfrm > 0 And end_80EEbankName > 0 And end_80EEAccntNum > 0 And end_80EELoanDate > 0 And end_80EELoanAmt > 0 And end_80EELoanOutstanding > 0 And end_80EEIntrst > 0
end_80EE = WorksheetFunction.Max(0, end_80EELoanfrm, end_80EEbankName, end_80EEAccntNum, end_80EELoanDate, end_80EELoanAmt, end_80EELoanOutstanding, end_80EEIntrst)
If Data80EE = True Then

    setTblinfo_24bLoanfrm
    'setTblinfo_24bIFSC   'Ankita_05/05/2025_Commented as per DESheet_v0.7
    'setTblinfo_24bPAN    'Ankita_05/05/2025_Commented as per DESheet_v0.7
    setTblinfo_24bBankName
    setTblinfo_24bAccntNum
    setTblinfo_24bLoanDate
    setTblinfo_24bLoanAmt
    setTblinfo_24bLoanOutstanding
    setTblinfo_24bIntrst
    
    'Ankita_05/05/2025_Commented as per DESheet_v0.7
'    Data24b = end_24bLoanfrm > 0 And end_24bIFSC > 0 Or end_24bPAN > 0 And end_24bbankName > 0 And end_24bAccntNum > 0 And end_24bLoanDate > 0 And end_24bLoanAmt > 0 And end_24bLoanOutstanding > 0 And end_24bIntrst > 0
     Data24b = end_24bLoanfrm > 0 And end_24bbankName > 0 And end_24bAccntNum > 0 And end_24bLoanDate > 0 And end_24bLoanAmt > 0 And end_24bLoanOutstanding > 0 And end_24bIntrst > 0
'    end_24b = WorksheetFunction.Max(0, end_24bLoanfrm, end_24bbankName, end_24bIFSC, end_24bPAN, end_24bAccntNum, end_24bLoanDate, end_24bLoanAmt, end_24bLoanOutstanding, end_24bIntrst)
     end_24b = WorksheetFunction.Max(0, end_24bLoanfrm, end_24bbankName, end_24bAccntNum, end_24bLoanDate, end_24bLoanAmt, end_24bLoanOutstanding, end_24bIntrst)


'--Bindu
   ' rowcount_Int24B = getRowNo(Sheet16.Range("LoanfrmBankOrInstitute.24b").name)
    
'    If end_24b = 1 Then
'       end_24b = rowcount_Int24B
'    Else
'       end_24b = ((end_24b - 1) + rowcount_Int24B) 'Newly Update on 02/05/2025
'    End If
'---Bindu
    
    rowcount_80EE = getRowNo(Sheet60.Range("LoanfrmBankOrInstitute.80EE").name)
    
    If end_80EE = 1 Then
       end_80EE = rowcount_80EE
    Else
       end_80EE = ((end_80EE - 1) + rowcount_80EE)   'Newly Update on 02/05/2025
    End If
    
    
    
    Dim Dedn_80EE, Dedn_24B, HpCount
    
    Dedn_80EE = Sheet60.Range("Combination_80EE").Column

     
    
    
    
    If Data24b = True Then
       For HpCount = 1 To Sheet8.Range("PropertySectionCOunt").value '--Bindu
           Dedn_24B = Sheet8.Range("Combination_24B" & HpCount).Column
           
           rowcount_Int24B = getRowNo(Sheet8.Range("LoanfrmBankOrInstitute.24b" & HpCount).name)
           end_24b = WorksheetFunction.Max(0, end_24bLoanfrm, end_24bbankName, end_24bAccntNum, end_24bLoanDate, end_24bLoanAmt, end_24bLoanOutstanding, end_24bIntrst)

            If end_24b = 1 Then
                end_24b = rowcount_Int24B
            Else
                end_24b = ((end_24b - 1) + rowcount_Int24B) 'Newly Update on 02/05/2025
            End If

            Combination_80EE_24B = False
            For i = rowcount_80EE To end_80EE
                For j = rowcount_Int24B To end_24b
    '                 Debug.Print Sheet60.Cells(i, Dedn_80EE).Value
    '                  Debug.Print Sheet16.Cells(j, Dedn_24B).Value
                    If UCase(Trim(Sheet60.Cells(i, Dedn_80EE).value)) = UCase(Trim(Sheet8.Cells(j, Dedn_24B).value)) Then
                       Combination_80EE_24B = True
                       GoTo checking
                    End If
                Next
                
            Next
          Next
    Else
        Combination_80EE_24B = False
    End If
    
    
    '--Bindu
checking:
If Combination_80EE_24B = False Then
        sourceSheet.Activate
         'MsgBox_80EE = MsgBox_80EE & ("""Deduction u/s 80EE can be claimed only if the limit u/s 24(b) is exhausted. So please note that the loan details of respective property under which 80EE is being claimed should be the same as in schedule 24(b)""") & Chr(13)
         MsgBox_80EE = MsgBox_80EE & ("""Deduction u/s 80EE can be claimed only if the limit u/s 24(b) is exhausted. So please note that the loan details of respective property under which 80EE is being claimed should be the same as in table 24(b)""") & Chr(13) 'Changed done (ver0.6)
         Validate80EE_All = False
         fmsgboxStatus MsgBox_80EE
         CloseMsg
        Exit Function
End If
    
'End If

'Added by 'Shrutika(23/05/2025)
Dim k, Total As Long

    For k = 1 To Sheet8.Range("PropertySectionCOunt").value
       Total = Total + Sheet8.Range("TotAmt.24b" & k).value
    
    Next
    
    If (Sheet60.Range("TotAmt.80EE").value > 0 And Not Total > 0) Then
        sourceSheet.Activate
        MsgBox_80EE = MsgBox_80EE & " ""Deduction u/s 80EE can be claimed only if the limit u/s 24(b) is exhausted.""" & Chr(13)
       fmsgboxStatus MsgBox_80EE
       CloseMsg
    End If
    
    

End If

'
 
'

 If MsgBox_80EE <> "" Then
 sourceSheet.Activate
  Validate80EE_All = False
  Else
  sourceSheet.Activate
  Validate80EE_All = True
  End If
'

 

'___________________________


End Function

Function Validate_80EE()
Validate_80EE = True
MsgBox_80EE = ""
setTblinfo_80EELoanfrm
'setTblinfo_80EEIFSC
end_80EE = 0

'setTblinfo_80EEPAN
setTblinfo_80EEBankName
setTblinfo_80EEAccntNum
setTblinfo_80EELoanDate
setTblinfo_80EELoanAmt
setTblinfo_80EELoanOutstanding
setTblinfo_80EEIntrst


'end_80EE = WorksheetFunction.Max(0, end_80EE, end_80EELoanfrm, end_80EEbankName, , end_80EEAccntNum, end_80EELoanDate, end_80EELoanAmt, end_80EELoanOutstanding, end_80EEIntrst)

'end_80EE = WorksheetFunction.Max(0, end_80EELoanfrm, end_80EEbankName, end_80EEPAN, end_80EEAccntNum, end_80EELoanDate, end_80EELoanAmt, end_80EELoanOutstanding, end_80EEIntrst)
end_80EE = WorksheetFunction.Max(0, end_80EELoanfrm, end_80EEbankName, end_80EEAccntNum, end_80EELoanDate, end_80EELoanAmt, end_80EELoanOutstanding, end_80EEIntrst)
If end_80EE > 0 Then


End If
If Not ValidateLoanfrm_80EE Then Validate_80EE = False
'If Not ValidateIFSC_80EE Then Validate_80EE = False
'If Not ValidatePAN_80EE Then Validate_80EE = False
If Not ValidateBankName_80EE Then Validate_80EE = False
If Not ValidateAccntNum_80EE Then Validate_80EE = False
If Not ValidateLoanDate_80EE Then Validate_80EE = False
If Not ValidateLoanAmt_80EE Then Validate_80EE = False

    If (Sheet60.Range("MaxLoan_80EE").value) > 3500000 Then
      '  MsgBox_80EE = MsgBox_80EE & "* ""Total loan taken"" shall not be more than Rs. 35 Lakhs in schedule 80EE" & Chr(13)
      MsgBox_80EE = MsgBox_80EE & "* ""Total amount of Loan shall not be more than Rs. 35 Lakhs in schedule 80EE" & Chr(13)
        Validate_80EE = False
    End If

If Not ValidateLoanOutstanding_80EE Then Validate_80EE = False
If Not ValidateIntrst_80EE Then Validate_80EE = False
'If Not ValidateLoanfrm_80EE Then Validate_80EE = False


    
        If Len(Sheet60.Range("TotAmt.80EE").value) > 14 Then
        MsgBox_80EE = MsgBox_80EE & " Total of Interest u/s 80EE in schedule 80EE cannot exceed 14 digits" & Chr(13)
        Validate_80EE = False
    End If


End Function

Sub setTblinfo_80EELoanfrm()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Sheet60.Activate
    mIntCells = Range("LoanfrmBankOrInstitute.80EE").count
    Set rangecells = Range("LoanfrmBankOrInstitute.80EE").Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
    end_80EELoanfrm = ccount
    
End Sub



Sub setTblinfo_80EEBankName()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("bankName.80EE").count
    Set rangecells = Range("bankName.80EE").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEbankName = ccount
    
End Sub



Sub setTblinfo_80EEAccntNum()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanAccNum.80EE").count
    Set rangecells = Range("loanAccNum.80EE").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEAccntNum = ccount
    
End Sub
Sub setTblinfo_80EELoanDate()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanDate.80EE").count
    Set rangecells = Range("loanDate.80EE").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EELoanDate = ccount
    
End Sub
Sub setTblinfo_80EELoanAmt()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanAmt.80EE").count
    Set rangecells = Range("loanAmt.80EE").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EELoanAmt = ccount
     
End Sub

Sub setTblinfo_80EELoanOutstanding()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanOutstanding.80EE").count
    Set rangecells = Range("loanOutstanding.80EE").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EELoanOutstanding = ccount
     
End Sub
Sub setTblinfo_80EEIntrst()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Intrst.80EE").count
    Set rangecells = Range("Intrst.80EE").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEIntrst = ccount
    
End Sub



Function ValidateLoanfrm_80EE() As Boolean
    ValidateLoanfrm_80EE = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("LoanfrmBankOrInstitute.80EE").Cells
    ReDim Loanfrm_80EE(end_80EE)
    For i = 1 To end_80EE
        Loanfrm_80EE(i) = rangecells.item(i).value
        
      '  If Not chkCompulsory(Loanfrm_80EE(i)) Then '
        If isdropdownblank(Loanfrm_80EE(i)) Then 'added by Shrutika(SIT-93632)
             MsgBox_80EE = MsgBox_80EE + "* Please select dropdown ""Loan taken from"" in schedule 80EE at Sr. No " & i & "" & Chr(13)
            ValidateLoanfrm_80EE = False
            Exit Function
        End If

         
Next
End Function


Function ValidateBankName_80EE() As Boolean
    ValidateBankName_80EE = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("bankName.80EE").Cells
    ReDim BankName_80EE(end_80EE)
    For i = 1 To end_80EE
        BankName_80EE(i) = rangecells.item(i).value
        If Not chkCompulsory(BankName_80EE(i)) Then
             'MsgBox_80EE = MsgBox_80EE + "* Please provide Name of the Bank/ Institution/Person from which the loan is taken in schedule 80EE at Sr. No " & i & "" & Chr(13)
              MsgBox_80EE = MsgBox_80EE + "* Please provide ""Name of the Bank/Institution from which the loan is taken"" in schedule 80EE at Sr. No " & i & "" & Chr(13)
            ValidateBankName_80EE = False
            Exit Function
        End If
        
         If Len(BankName_80EE(i)) > 125 Then
         ' MsgBox_80EE = MsgBox_80EE + "* Name of the Bank/ Institution/Person at Sr. No " & i & " in Sheet 80EE less than 75 characters." & Chr(13)
            MsgBox_80EE = MsgBox_80EE + "* Name of the Bank/Institution in Sheet 80EE should be less than or equal to 125 characters at Sr. No " & i & " ." & Chr(13)
            ValidateBankName_80EE = False
            Exit Function
        End If
         If Not checkfieldSuperSpecialcharacter(BankName_80EE(i)) Then
             MsgBox_80EE = MsgBox_80EE + "* Name of the Bank/Institution from which the loan is taken in schedule 80EE at Sl.no. " & i & " should not Contain <, >, characters." & Chr(13)
            ValidateBankName_80EE = False
            Exit Function
        End If
         
Next
End Function


Function ValidateAccntNum_80EE() As Boolean
    ValidateAccntNum_80EE = True

    Dim rangecells, rangecells1 As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("loanAccNum.80EE").Cells
    Set rangecells1 = Sheet60.Range("LoanfrmBankOrInstitute.80EE").Cells
    
    ReDim AccntNum_80EE(end_80EE)
    ReDim Bankorothebank_80EE(end_80EE)
    
    For i = 1 To end_80EE
        AccntNum_80EE(i) = rangecells.item(i).value
        Bankorothebank_80EE(i) = rangecells1.item(i).value
        
        If Not chkCompulsory(AccntNum_80EE(i)) Then
        
'            MsgBox_80EE = MsgBox_80EE + "* Please provide ""Loan Account number of the Bank / Reference number of the Institution"" from which loan is taken in schedule 80EE at Sr. No " & i & "" & Chr(13)
             MsgBox_80EE = MsgBox_80EE + "* Please provide ""Loan Account number of the Bank/Institution"" from which loan is taken in schedule 80EE at Sr. No " & i & "" & Chr(13)
            ValidateAccntNum_80EE = False
            Exit Function
        End If
         If Len(AccntNum_80EE(i)) > 20 Then
'          MsgBox_80EE = MsgBox_80EE + "* Loan Account number  at Sr. No " & i & " in Sheet 80EE less than 20 characters." & Chr(13)
            MsgBox_80EE = MsgBox_80EE + "* Loan Account number in Sheet 80EE should be less than or equal to 20 characters at Sr. No " & i & "." & Chr(13)
            ValidateAccntNum_80EE = False
            Exit Function
        End If
       
       If Bankorothebank_80EE(i) = "Bank" Then
        If Not ValidateBankAccountNumber_80EE(AccntNum_80EE(i), i) Then
        'Msgbox_BA = Msgbox_BA + "Account Number at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
        ValidateAccntNum_80EE = False
        Exit Function
    End If
    End If
     
     If Bankorothebank_80EE(i) = "Institution" Then
     If Not checkfieldspecialcharacter_BthacntReferencenumber(AccntNum_80EE(i)) Then
        MsgBox_80EE = MsgBox_80EE + "* Loan Account number of the Bank / Institution at Sr.No " & i & " is invalid  in schedule 80EE, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
        ValidateAccntNum_80EE = False
        Exit Function
      End If
    End If
Next
End Function

Function ValidateLoanDate_80EE() As Boolean
    ValidateLoanDate_80EE = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("loanDate.80EE").Cells
    ReDim LoanDate_80EE(end_80EE)
    For i = 1 To end_80EE
        LoanDate_80EE(i) = rangecells.item(i).value
        If Not chkCompulsory(LoanDate_80EE(i)) Then
             'MsgBox_80EE = MsgBox_80EE + "* Please provide Date of taking loan in schedule 80EE at Sr. No " & i & "" & Chr(13)
               MsgBox_80EE = MsgBox_80EE + "* Please provide ""Date of sanction of loan"" in schedule 80EE at Sr. No " & i & "" & Chr(13)
            ValidateLoanDate_80EE = False
            Exit Function
        End If
        
         If Len(LoanDate_80EE(i)) > 10 Then
          'MsgBox_80EE = MsgBox_80EE + "* Date of taking loan  at Sr. No " & i & " in Sheet 80EE less than 10 characters." & Chr(13)
          MsgBox_80EE = MsgBox_80EE + "* Date of sanction of loan  at Sr. No " & i & " in Sheet 80EE less than 10 characters." & Chr(13)
            ValidateLoanDate_80EE = False
            Exit Function
        End If
        
        'CheckDateddmmyyyy
         If Not CheckDateddmmyyyy(LoanDate_80EE(i)) Then
         MsgBox_80EE = MsgBox_80EE + "* ""Please enter date in valid format"" at Sr. No " & i & "." & Chr(13)
        ValidateLoanDate_80EE = False
                 

          Exit Function
        End If
        
        If Not ChkMaxDate_80EE(Trim(LoanDate_80EE(i)), "31-03-2017") Then
            MsgBox_80EE = MsgBox_80EE + "* ""Date of sanction of Loan"" shall be between 01/04/2016 to 31/03/2017 in schedule 80EE at Sr. No " & i & "." & Chr(13)
            ValidateLoanDate_80EE = False
                 

          Exit Function
        End If
        
       If Not ChkMinInclusiveDate(Trim(Dformat(LoanDate_80EE(i), "yyyy-mm-dd")), "2016-04-01") Then
           MsgBox_80EE = MsgBox_80EE + "* ""Date of sanction of Loan"" shall be between 01/04/2016 to 31/03/2017 in schedule 80EE at Sr. No " & i & "." & Chr(13)
            ValidateLoanDate_80EE = False
                 
                 

          Exit Function
        End If
                 
                 
          
         
Next
End Function

Function ValidateLoanAmt_80EE() As Boolean
    ValidateLoanAmt_80EE = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("loanAmt.80EE").Cells
    ReDim LoanAmt_80EE(end_80EE)
    For i = 1 To end_80EE
        LoanAmt_80EE(i) = rangecells.item(i).value
        If Not chkCompulsory(LoanAmt_80EE(i)) Then
        
             MsgBox_80EE = MsgBox_80EE + "* ""Please provide ""Total amount of Loan"" in schedule 80EE"" at Sr. No " & i & "" & Chr(13)
            ValidateLoanAmt_80EE = False
            Exit Function
        End If

        
        If Not IsNumeric(LoanAmt_80EE(i)) Then
            MsgBox_80EE = MsgBox_80EE & "* Loan amount at Sr. No  " & i & "  in Sheet 80EE should be Numeric value" & Chr(13)
            ValidateLoanAmt_80EE = False
            Exit Function
        End If
        
        If LoanAmt_80EE(i) > 99999999999999# Then
            MsgBox_80EE = MsgBox_80EE & "* Loan amount at Sr. No  " & i & "  in Sheet 80EE cannot exceed 14 digits" & Chr(13)
            ValidateLoanAmt_80EE = False
            Exit Function
        End If
        
        If LoanAmt_80EE(i) < 0 Or LoanAmt_80EE(i) = 0 Then
        
            MsgBox_80EE = MsgBox_80EE & "* ""Total amount of Loan"" should be more than 0 in schedule 80EE" & Chr(13)
            ValidateLoanAmt_80EE = False
            Exit Function
        End If
         
Next
End Function

Function ValidateLoanOutstanding_80EE() As Boolean
    ValidateLoanOutstanding_80EE = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("loanOutstanding.80EE").Cells
    ReDim LoanOutStanding_80EE(end_80EE)
    For i = 1 To end_80EE
        LoanOutStanding_80EE(i) = rangecells.item(i).value
        If Not chkCompulsory(LoanOutStanding_80EE(i)) Then
       
         '    MsgBox_80EE = MsgBox_80EE + "* ""Loan outstanding as on last date of financial year is mandatory in schedule 80EE"" at Sr. No " & i & "" & Chr(13)
             MsgBox_80EE = MsgBox_80EE + "*  ""Loan  outstanding as on last date of financial year is mandatory. Please enter 0 in case the entire loan is repaid during the year"" in schedule 80EE at Sr. No " & i & "" & Chr(13)
          
            ValidateLoanOutstanding_80EE = False
            Exit Function
        End If

        
        If Not IsNumeric(LoanOutStanding_80EE(i)) Then
            MsgBox_80EE = MsgBox_80EE & "* Loan outstanding at Sr. No  " & i & "  in Sheet 80EE should be Numeric value" & Chr(13)
            ValidateLoanOutstanding_80EE = False
            Exit Function
        End If
        
        If LoanOutStanding_80EE(i) > 99999999999999# Then
            MsgBox_80EE = MsgBox_80EE & "* Loan outstanding at Sr. No  " & i & "  in Sheet 80EE cannot exceed 14 digits" & Chr(13)
            ValidateLoanOutstanding_80EE = False
            Exit Function
        End If
        
        If LoanOutStanding_80EE(i) < 0 Then
      
'            MsgBox_80EE = MsgBox_80EE & "* ""Loan outstanding as on 31.03.2025"" can't be less than 0 in schedule 80EE at Sr. No  " & i & ". You may please enter as 0 if it become negative as result of excess payment"" & Chr(13)"
            MsgBox_80EE = MsgBox_80EE & "* ""Loan outstanding as on last date of financial year"" can't be less than 0 in schedule 80EE at Sr. No  " & i & ". You may please enter as 0 if it become negative as result of excess payment"" & Chr(13)"
            ValidateLoanOutstanding_80EE = False
            Exit Function
        End If
         
Next
End Function

Function ValidateIntrst_80EE() As Boolean
    ValidateIntrst_80EE = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet60.Range("Intrst.80EE").Cells
    ReDim Intrst_80EE(end_80EE)
    For i = 1 To end_80EE
        Intrst_80EE(i) = rangecells.item(i).value
        If Not chkCompulsory(Intrst_80EE(i)) Then
             MsgBox_80EE = MsgBox_80EE + "* ""Please provide Interest u/s 80EE"" at Sr. No " & i & "" & Chr(13)
            ValidateIntrst_80EE = False
            Exit Function
        End If


        If Not IsNumeric(Intrst_80EE(i)) Then
            MsgBox_80EE = MsgBox_80EE & "* Interest at Sr. No  " & i & "  in Sheet 80EE should be Numeric value" & Chr(13)
            ValidateIntrst_80EE = False
            Exit Function
        End If
        
        If Intrst_80EE(i) > 99999999999999# Then
            MsgBox_80EE = MsgBox_80EE & "* Interest at Sr. No  " & i & "  in Sheet 80EE cannot exceed 14 digits" & Chr(13)
            ValidateIntrst_80EE = False
            Exit Function
        End If
        
        'Interest u/s 80EE should be more than 0 in schedule 80EE
        
        If Intrst_80EE(i) < 0 Or Intrst_80EE(i) = 0 Then
            MsgBox_80EE = MsgBox_80EE & "* ""Interest u/s 80EE"" should be more than 0 in schedule 80EE at Sr. No  " & i & "" & Chr(13)
            ValidateIntrst_80EE = False
            Exit Function
        End If
         
Next
End Function


Function ValidateBankAccountNumber_80EE(BankAccountNumber As Variant, cc As Long) As Boolean
    ValidateBankAccountNumber_80EE = True
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
            MsgBox_80EE = MsgBox_80EE & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EE, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber_80EE = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            MsgBox_80EE = MsgBox_80EE & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EE " & Chr(13)
            ValidateBankAccountNumber_80EE = False
            Exit Function
        End If

    
    End If
  '----------------------------------------------------------------
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        MsgBox_80EE = MsgBox_80EE & "*  Please enter the Loan Account number in Bank Details at Sr.No " & cc & " in schedule 80EE" & Chr(13)
        ValidateBankAccountNumber_80EE = False
        Exit Function
    End If
    
    
    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        MsgBox_80EE = MsgBox_80EE & "* Loan Account number at Sr.No " & cc & " is mandatory in schedule 80EE" & Chr(13)
        ValidateBankAccountNumber_80EE = False
        Exit Function
    End If
'--------------------------------------------------------------------

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0) Or (InStr(BankAccountNumber, "-/") > 0) Or (InStr(BankAccountNumber, "/-") > 0)) Then
        MsgBox_80EE = MsgBox_80EE & " Loan Account number at Sr.No " & cc & " is invalid in schedule 80EE" & Chr(13)
        ValidateBankAccountNumber_80EE = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        MsgBox_80EE = MsgBox_80EE & " Loan Account number at Sr.No " & cc & " is invalid in schedule 80EE" & Chr(13)
        ValidateBankAccountNumber_80EE = False
        Exit Function
    End If
    
    If (Not checkfieldspecialcharacter(Mid(BankAccountNumber, 1, 1))) Then
        MsgBox_80EE = MsgBox_80EE & " Loan Account number at Sr.No " & cc & " is invalid in schedule 80EE" & Chr(13)
        ValidateBankAccountNumber_80EE = False
        Exit Function
    End If


    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        MsgBox_80EE = MsgBox_80EE & " Loan Account number at Sr.No " & cc & " is invalid in schedule 80EE" & Chr(13)
        ValidateBankAccountNumber_80EE = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        MsgBox_80EE = MsgBox_80EE & " Loan Account number at Sr.No " & cc & " is invalid in schedule 80EE" & Chr(13)
        ValidateBankAccountNumber_80EE = False
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
            MsgBox_80EE = MsgBox_80EE & "* Loan Account number at Sr.No " & cc + 1 & " is invalid in schedule 80EE" & Chr(13)
            ValidateBankAccountNumber_80EE = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "* Loan Account number at Sr.No " & cc + 1 & "  in schedule 80EE is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber_80EE = False
        Exit Function
    End If
End Function


Public Function ChkMaxDate_80EE(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

     ChkMaxDate_80EE = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

        If Year > 2017 Then
            ChkMaxDate_80EE = False
            Exit Function
        Else
                If Year = 2017 Then
                If month > 4 Then
                    ChkMaxDate_80EE = False
                    Exit Function
                Else
                    If month = 4 Then
                        If dat > 1 Then
                           ChkMaxDate_80EE = False
                            Exit Function
                        Else
                            If dat = 1 Then
                               ChkMaxDate_80EE = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function


Sub CmdSchedule80EEAdd_Click()
Dim vRows As Long
    Sheets("80E_80EE_80EEA_80EEB").Activate
    EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80EE||bankName.80EE||loanAccNum.80EE||loanDate.80EE||loanAmt.80EE||loanOutstanding.80EE||Intrst.80EE"
    ActiveCellRange = EfilingCommon.searchLastRow("LoanfrmBankOrInstitute.80EE")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

