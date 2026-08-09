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
Public HpCount As Variant
Public end_24bLoanfrm2 As Variant
Public end_24bbankName2 As Variant
Public end_24bAccntNum2 As Variant
Public end_24bLoanDate2 As Variant
Public end_24bLoanAmt2 As Variant
Public end_24bIntrst2 As Variant
Public end_24bLoanOutstanding2 As Variant
Public end_24b2 As Variant


Sub ValidateSheet80EEA_Click()
Dim vbMessgaeCaption As String

vbMessgaeCaption = "ITR 4: AY: 2024-25"

If Validate80EEA_All Then
Sheet20.Activate
   ' MsgBox (MsgBox_80GA), vbOKOnly, vbMessgaeCaption
    'fmsgboxStatus MsgBox_80EEA
   ' CloseMsg
fmsgboxStatus "Schedule 80EEA is OK"

CloseMsg
End If
'fmsgboxStatus "Sheet 80EEA is OK"
End Sub
'Sub Validate80EEA_All()
Function Validate80EEA_All()
 Validate80EEA_All = True
Dim vbMessgaeCaption As String
vbMessgaeCaption = "Error"
   Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("80E_80EE_80EEA_80EEB")
    
'    If Not lock_80EEA_flag = False Then
'      MsgBox_80EEA = MsgBox_80EEA & " ""Deduction u/s 80EE and 80EEA can't be claimed together""" & Chr(13)
'      fmsgboxStatus MsgBox_80EEA
'       CloseMsg
'    End If
    
If Not Validate_80EEA Then
    sourceSheet.Activate
   ' MsgBox (MsgBox_80GA), vbOKOnly, vbMessgaeCaption
    fmsgboxStatus MsgBox_80EEA
    CloseMsg
End If

    

'Ayush_21/04----------------------
Dim rowcount_80EEA, i, j, rowcount_Int24B, rowcount_80EE   As Long


'for80EEA
setTblinfo_80EEALoanfrm
'setTblinfo_80EEAIFSC

'setTblinfo_80EEAPAN
setTblinfo_80EEABankName
setTblinfo_80EEAAccntNum
setTblinfo_80EEALoanDate
setTblinfo_80EEALoanAmt
setTblinfo_80EEALoanOutstanding
setTblinfo_80EEAIntrst

Dim Data80EEA, Data24b, Combination_80EEA_24B As Boolean

'Data80EEA = end_80EEALoanfrm > 0 And end_80EEAIFSC > 0 Or end_80EEAPAN > 0 And end_80EEAbankName > 0 And end_80EEAAccntNum > 0 And end_80EEALoanDate > 0 And end_80EEALoanAmt > 0 And end_80EEALoanOutstanding > 0 And end_80EEAIntrst > 0
Data80EEA = end_80EEALoanfrm > 0 And end_80EEAbankName > 0 And end_80EEAAccntNum > 0 And end_80EEALoanDate > 0 And end_80EEALoanAmt > 0 And end_80EEALoanOutstanding > 0 And end_80EEAIntrst > 0

'end_80EEA = WorksheetFunction.Max(0, end_80EEALoanfrm, end_80EEAbankName, end_80EEAIFSC, end_80EEAPAN, end_80EEAAccntNum, end_80EEALoanDate, end_80EEALoanAmt, end_80EEALoanOutstanding, end_80EEAIntrst)
end_80EEA = WorksheetFunction.Max(0, end_80EEALoanfrm, end_80EEAbankName, end_80EEAAccntNum, end_80EEALoanDate, end_80EEALoanAmt, end_80EEALoanOutstanding, end_80EEAIntrst)

If Data80EEA = True Then

    setTblinfo_24bLoanfrm_com

   setTblinfo_24bBankName_com
   setTblinfo_24bAccntNum_com
   setTblinfo_24bLoanDate_com
    setTblinfo_24bLoanAmt_com
   setTblinfo_24bLoanOutstanding_com
   setTblinfo_24bIntrst_com
    
'    Data24b = end_24bLoanfrm > 0 And end_24bIFSC > 0 Or end_24bPAN > 0 And end_24bbankName > 0 And end_24bAccntNum > 0 And end_24bLoanDate > 0 And end_24bLoanAmt > 0 And end_24bLoanOutstanding > 0 And end_24bIntrst > 0
    Data24b = end_24bLoanfrm2 > 0 And end_24bbankName2 > 0 And end_24bAccntNum2 > 0 And end_24bLoanDate2 > 0 And end_24bLoanAmt2 > 0 And end_24bLoanOutstanding2 > 0 And end_24bIntrst2 > 0
    
'    end_24b = WorksheetFunction.Max(0, end_24bLoanfrm, end_24bbankName, end_24bIFSC, end_24bPAN, end_24bAccntNum, end_24bLoanDate, end_24bLoanAmt, end_24bLoanOutstanding, end_24bIntrst)
    end_24b2 = WorksheetFunction.Max(0, end_24bLoanfrm2, end_24bbankName2, end_24bAccntNum2, end_24bLoanDate2, end_24bLoanAmt2, end_24bLoanOutstanding2, end_24bIntrst2)
    
'    rowcount_Int24B = getRowNo(Sheet21.Range("LoanfrmBankOrInstitute.24b").name)
'
'    If end_24b = 1 Then
'       end_24b = rowcount_Int24B
'    Else
'       end_24b = ((end_24b - 1) + rowcount_Int24B)
'    End If
'
'    rowcount_80EEA = getRowNo(Sheet20.Range("LoanfrmBankOrInstitute.80EEA").name)
'
'    If end_80EEA = 1 Then
'       end_80EEA = rowcount_80EEA
'    Else
'       end_80EEA = ((end_80EEA - 1) + rowcount_80EEA)
'    End If
    
    
    rowcount_80EEA = getRowNo(Sheet20.Range("LoanfrmBankOrInstitute.80EEA").name)
    If end_80EEA = 1 Then
       end_80EEA = rowcount_80EEA
    Else
       end_80EEA = ((end_80EEA - 1) + rowcount_80EEA)
    End If


    Dim Dedn_80EEA, Dedn_24B
    
    Dedn_80EEA = Sheet20.Range("Combination_80EEA").Column
'    Dedn_24B = Sheet21.Range("Combination_24B").Column
    
    
    
'    If Data24b = True Then
'        Combination_80EEA_24B = False
'        For i = rowcount_80EEA To end_80EEA
'            For j = rowcount_Int24B To end_24b
'                If UCase(Trim(Sheet20.Cells(i, Dedn_80EEA).value)) = UCase(Trim(Sheet21.Cells(j, Dedn_24B).value)) Then
'                   Combination_80EEA_24B = True
'                   Exit For
'
'                End If
'            Next
'
'        Next
'    Else
'        Combination_80EEA_24B = False
'        'MsgBox_80EEA = MsgBox_80EEA & ("Deduction u/s 80EEA can be claimed only if the limit u/s 24(b) is exhausted. So please note that the loan details of respective property under which 80EE is being claimed should be the same as in schedule 24(b)") & Chr(13)
'        'fmsgboxStatus MsgBox_80EEA
'    End If
''End If

    'Ankita_24/01/2026===========

    If Data24b = True Then
        Combination_80EEA_24B = False
         For HpCount = 1 To Sheet22.Range("PropertySectionCOunt").value
               Dedn_24B = Sheet22.Range("Combination_24B" & HpCount).Column
               
               rowcount_Int24B = getRowNo(Sheet22.Range("LoanfrmBankOrInstitute.24b" & HpCount).name)
               end_24b2 = WorksheetFunction.Max(0, end_24bLoanfrm2, end_24bbankName2, end_24bAccntNum2, end_24bLoanDate2, end_24bLoanAmt2, end_24bLoanOutstanding2, end_24bIntrst2)
    
                If end_24b2 = 1 Then
                    end_24b2 = rowcount_Int24B
                Else
                    end_24b2 = ((end_24b2 - 1) + rowcount_Int24B)
                End If
                For i = rowcount_80EEA To end_80EEA
                    For j = rowcount_Int24B To end_24b2
                        If UCase(Trim(Sheet20.Cells(i, Dedn_80EEA).value)) = UCase(Trim(Sheet22.Cells(j, Dedn_24B).value)) Then
                           Combination_80EEA_24B = True
                           GoTo checking
                        End If
                    Next
                Next
           Next
    Else
        Combination_80EEA_24B = False
    End If
'End If

checking:


If Combination_80EEA_24B = False Then
'SIT-92061 by sai on 25/04/2024
        MsgBox_80EEA = MsgBox_80EEA & ("""Deduction u/s 80EEA can be claimed only if the limit u/s 24(b) is exhausted. So please note that the loan details of respective property under which 80EEA is being claimed should be the same as in schedule 24(b)""") & Chr(13)
        Validate80EEA_All = False 'Ayush-09/05
        fmsgboxStatus MsgBox_80EEA
        CloseMsg
'        Validate80EEA_All = False
        Exit Function
End If

End If

Dim k, Total As Long

    For k = 1 To Sheet22.Range("PropertySectionCount").value
       Total = Total + Sheet22.Range("TotAmt.24b" & k).value
    
    Next

'Ayush_21/04
            
'  If (Sheet20.Range("TotAmt.80EEA").value > 0 And Not Sheet1.Range("IncD.IncomeFromHP_v").value > 0) Then
  If (Sheet20.Range("TotAmt.80EEA").value > 0 And Not Total > 0) Then
        MsgBox_80EEA = MsgBox_80EEA & " ""Deduction u/s 80EEA can be claimed only if the limit u/s 24(b) is exhausted.""" & Chr(13)
        fmsgboxStatus MsgBox_80EEA
       CloseMsg
    End If
    
'Ayush_21/04
 If MsgBox_80EEA <> "" Then
  Validate80EEA_All = False
  Else
  Validate80EEA_All = True
  End If
  
  '-------------------------------


End Function

Function Validate_80EEA()
Validate_80EEA = True
MsgBox_80EEA = ""


setTblinfo_80EEALoanfrm
'setTblinfo_80EEAIFSC

'setTblinfo_80EEAPAN
setTblinfo_80EEABankName
setTblinfo_80EEAAccntNum
setTblinfo_80EEALoanDate
setTblinfo_80EEALoanAmt
setTblinfo_80EEALoanOutstanding
setTblinfo_80EEAIntrst


end_80EEA = WorksheetFunction.Max(0, end_80EEALoanfrm, end_80EEAbankName, , end_80EEAAccntNum, end_80EEALoanDate, end_80EEALoanAmt, end_80EEALoanOutstanding, end_80EEAIntrst)

'Ayush_15/05
If Sheet20.Range("Stampduty.80EEA").value > 0 And end_80EEA = 0 Then
     MsgBox_80EEA = MsgBox_80EEA & " ""Please provide details in respect of interest on loan taken for house property""" & Chr(13)
        Validate_80EEA = False
End If



If end_80EEA > 0 Then
    If (Sheet20.Range("Stampduty.80EEA").value) = "" Then
        MsgBox_80EEA = MsgBox_80EEA & " ""Stamp value of residential house property"" is mandatory in schedule 80EEA" & Chr(13)
        Validate_80EEA = False
    End If

     If (Sheet20.Range("Stampduty.80EEA").value) > 4500000 Then
        MsgBox_80EEA = MsgBox_80EEA & " ""Stamp value of residential house property"" shall not be more than Rs. 45 Lakhs in schedule 80EEA" & Chr(13)
        Validate_80EEA = False
    End If
    
    End If
    
    
If Not ValidateLoanfrm_80EEA Then Validate_80EEA = False
'If Not ValidateIFSC_80EEA Then Validate_80EEA = False
'If Not ValidatePAN_80EEA Then Validate_80EEA = False
If Not ValidateBankName_80EEA Then Validate_80EEA = False
If Not ValidateAccntNum_80EEA Then Validate_80EEA = False
If Not ValidateLoanDate_80EEA Then Validate_80EEA = False
If Not ValidateLoanAmt_80EEA Then Validate_80EEA = False

'    If (Sheet20.Range("MaxLoan_80EEA").value) > 3500000 Then
'        MsgBox_80EEA = MsgBox_80EEA & "* ""Total loan taken"" shall not be more than Rs. 35 Lakhs in schedule 80EEA" & Chr(13)
'        Validate_80EEA = False
'    End If

If Not ValidateLoanOutstanding_80EEA Then Validate_80EEA = False
If Not ValidateIntrst_80EEA Then Validate_80EEA = False
'If Not ValidateLoanfrm_80EEA Then Validate_80EEA = False

'ValidateIFSC_80EEA

    If (Sheet20.Range("TotAmt.80EEA").value) > 2000000 Then
        MsgBox_80EEA = MsgBox_80EEA & " To claim ""Interest payable on borrowed capital more than Rs.20 lakhs"", you may please consider filing ITR 3. Refer Rule 12 for further details" & Chr(13)
        Validate_80EEA = False
    End If
    
    If Len(Sheet20.Range("TotAmt.80EEA").value) > 14 Then
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

'Commented by Ayush_05/05/2025 for 0.7 changes
'Sub setTblinfo_80EEAIFSC()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("IFSC.80EEA").count
'    Set rangecells = Range("IFSC.80EEA").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80EEAIFSC = ccount
'End Sub


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

'Commented by Ayush_05/05/2025 for 0.7 changes

'Sub setTblinfo_80EEAPAN()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("PAN.80EEA").count
'    Set rangecells = Range("PAN.80EEA").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80EEAPAN = ccount
'End Sub

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
    Set rangecells = Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Cells
    ReDim Loanfrm_80EEA(end_80EEA)
    For i = 1 To end_80EEA
        Loanfrm_80EEA(i) = rangecells.item(i).value
        If isdropdownblank(Loanfrm_80EEA(i)) Then
'             MsgBox_80EEA = MsgBox_80EEA + "* Please select dropdown from ""Loan taken from"" in 80EEA schedule at Sr. No " & i & "" & Chr(13)

'SIT-91985 by sai on 25/04/2025
              MsgBox_80EEA = MsgBox_80EEA + "* Please select dropdown from ""Loan taken from"" in schedule 80EEA at Sr. No " & i & "" & Chr(13)
            ValidateLoanfrm_80EEA = False
            Exit Function
        End If
'         If Loanfrm_80EEA(i) = "(Select)" Then
'          MsgBox_80EEA = MsgBox_80EEA + "* Please select dropdown from ""Loan taken from""  in 80EEA schedule at Sr. No " & i & "" & Chr(13)
'            ValidateLoanfrm_80EEA = False
'            Exit Function
'        End If
         
Next
End Function
'Commented by Ayush_05/05/2025 for 0.7 changes
'Function ValidateIFSC_80EEA() As Boolean
'ValidateIFSC_80EEA = True
'
'   ' setTblinfo_80EEAIFSC
'   setTblinfo_80EEALoanfrm
'    Dim rangecells As Range
'    Dim rangecells1 As Range
'    Set rangecells = Range("LoanfrmBankOrInstitute.80EEA").Cells
'    Set rangecells1 = Range("IFSC.80EEA").Cells
'    Dim cellrange As String
'    Dim cellRange1 As String
'    Dim i As Long
'    ReDim Others_Loan(end_80EEALoanfrm)
'    ReDim Others_IFSC(end_80EEALoanfrm)
'
'
'
'
'
'
'
'    For i = 1 To end_80EEALoanfrm
'    cellrange = GetMergedAddressCell(rangecells, i)
'    cellRange1 = GetMergedAddressCell(rangecells1, i)
'
'        Others_Loan(i) = Sheet20.Range(cellrange).value
'        Others_IFSC(i) = Sheet20.Range(cellRange1).value
'
'
'
'         If (Others_Loan(i) = "Bank") Then
'         If Others_IFSC(i) = "" Then
'         'Change.03.03.2023.102.IDS.113
'             'msgError = msgError & "* Please enter description in Nature of Income (Exempt income (For reporting Purposes)) in Sheet Income Details" & Chr(13)
'             MsgBox_80EEA = MsgBox_80EEA & "* Please provide  IFSC of the Bank from which loan is taken in schedule 80EEA at Sr. No " & i & "" & Chr(13)
'         'End Change.IDS113
'             ValidateIFSC_80EEA = False
'             Exit Function
'         End If
'         End If
'
'         If Len(Others_IFSC(i)) > 11 Then
'          MsgBox_80EEA = MsgBox_80EEA + "* IFSC of the Bank at Sr. No " & i & " in Sheet 80EEA should be 11 characters." & Chr(13)
'            ValidateIFSC_80EEA = False
'            Exit Function
'        End If
'
'    Next
'
'End Function

'Commented by Ayush_05/05/2025 for 0.7 changes

'Function ValidatePAN_80EEA() As Boolean
'ValidatePAN_80EEA = True
'
'   ' setTblinfo_80EEAIFSC
'   setTblinfo_80EEALoanfrm
'    Dim rangecells As Range
'    Dim rangecells1 As Range
'    Set rangecells = Range("LoanfrmBankOrInstitute.80EEA").Cells
'    Set rangecells1 = Range("PAN.80EEA").Cells
'    Dim cellrange As String
'    Dim cellRange1 As String
'    Dim i As Long
'    ReDim Others_Loan(end_80EEALoanfrm)
'    ReDim Others_PAN(end_80EEALoanfrm)
'
'
'
'
'
'
'
'    For i = 1 To end_80EEALoanfrm
'    cellrange = GetMergedAddressCell(rangecells, i)
'    cellRange1 = GetMergedAddressCell(rangecells1, i)
'
'        Others_Loan(i) = Sheet20.Range(cellrange).value
'        Others_PAN(i) = Sheet20.Range(cellRange1).value
'
'
'
'         If (Others_Loan(i) = "Institution") Then
'         If Others_PAN(i) = "" Then
'         'Change.03.03.2023.102.IDS.113
'             'msgError = msgError & "* Please enter description in Nature of Income (Exempt income (For reporting Purposes)) in Sheet Income Details" & Chr(13)
'             MsgBox_80EEA = MsgBox_80EEA & "* Please provide  PAN of the institution from which loan is taken in schedule 80EEA. at Sr. No " & i & "" & Chr(13)
'         'End Change.IDS113
'             ValidatePAN_80EEA = False
'             Exit Function
'         End If
'         End If
'
'         If Len(Others_PAN(i)) > 10 Then
'          MsgBox_80EEA = MsgBox_80EEA + "* PAN of the institution from which loan is taken at Sr. No " & i & " in Sheet 80EEA should be 10 characters." & Chr(13)
'            ValidatePAN_80EEA = False
'            Exit Function
'        End If
'
'    Next
'
'End Function


Function ValidateBankName_80EEA() As Boolean
    ValidateBankName_80EEA = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet20.Range("bankName.80EEA").Cells
    ReDim BankName_80EEA(end_80EEA)
    For i = 1 To end_80EEA
        BankName_80EEA(i) = rangecells.item(i).value
        If Not chkCompulsory(BankName_80EEA(i)) Then
             MsgBox_80EEA = MsgBox_80EEA + "* Please provide ""Name of the Bank/ Institution from which the loan is taken"" in schedule 80EEA at Sr. No " & i & "" & Chr(13)
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

'Function ValidateIFSC_80EEA() As Boolean
'    ValidateIFSC_80EEA = True
'
'    Dim rangecells As Range
'    Dim rangecells1 As Range
'    Dim i As Long
'    Set rangecells = Sheet20.Range("IFSC.80EEA").Cells
'    Set rangecells1 = Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Cells
'    ReDim IFSC_80EEA(end_80EEA)
'    ReDim Loanfrom_80EEA(end_80EEA)
'    For i = 1 To end_80EEA
'        IFSC_80EEA(i) = rangecells.item(i).value
'       ' Loanfrom_80EEA(i) = rangecells.item(i).value
'     '   If Loanfrom_80EEA(i) = "Bank" Then
'        If Not chkCompulsory(IFSC_80EEA(i)) Then
'             MsgBox_80EEA = MsgBox_80EEA + "* Please select dropdown from ""Loan taken from"" at Sr. No " & i & " in 80EEA schedule." & Chr(13)
'            ValidateIFSC_80EEA = False
'            Exit Function
'        End If
'      '  End If
'
'         If Len(IFSC_80EEA(i)) > 11 Then
'          MsgBox_80EEA = MsgBox_80EEA + "* IFSC at Sr. No " & i & " in Sheet 80EEA should be 11 characters." & Chr(13)
'            ValidateIFSC_80EEA = False
'            Exit Function
'        End If
'
'Next
''
''If (Sheet20.Range("LoanfrmBankOrInstitute.80EE").value = "Bank") Then
''        If Sheet20.Range("IFSC.80EE").value = "" Then
'''            msgError = msgError & "*If A23a is ""Yes"" Date of filing of Form 10IEA is mandatory. " & Chr(13)
'''Ayush_10/02
'''            msgError = msgError & "*""Date of filing of Form 10IEA is mandatory.""" & Chr(13) 'Modified by sai on 27/01/2025
''            MsgBox_80EEA = MsgBox_80EEA & "*""IFSC""" & Chr(13)
''            ValidateIFSC_80EEA = False
''        End If
''        End If
'
'End Function
Function ValidateAccntNum_80EEA() As Boolean
    ValidateAccntNum_80EEA = True

    Dim rangecells As Range
    Dim rangecells1 As Range
    
    Dim i As Long
    Set rangecells = Sheet20.Range("loanAccNum.80EEA").Cells
    Set rangecells1 = Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Cells
    ReDim AccntNum_80EEA(end_80EEA)
    ReDim BankorInst_80EEA(end_80EEA)
    
    For i = 1 To end_80EEA
        AccntNum_80EEA(i) = rangecells.item(i).value
        BankorInst_80EEA(i) = rangecells1.item(i).value
        If Not chkCompulsory(AccntNum_80EEA(i)) Then
             MsgBox_80EEA = MsgBox_80EEA + "* Please provide ""Loan Account number of the Bank / Institution"" from which loan is taken in schedule 80EEA at Sr. No " & i & "" & Chr(13)
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
        'Msgbox_BA = Msgbox_BA + "Account Number at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
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
    Set rangecells = Sheet20.Range("loanDate.80EEA").Cells
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
    Set rangecells = Sheet20.Range("loanAmt.80EEA").Cells
    ReDim LoanAmt_80EEA(end_80EEA)
    For i = 1 To end_80EEA
        LoanAmt_80EEA(i) = rangecells.item(i).value
        If Not chkCompulsory(LoanAmt_80EEA(i)) Then
             MsgBox_80EEA = MsgBox_80EEA + "* ""Please provide Total amount of Loan in schedule 80EEA"" at Sr. No " & i & "" & Chr(13)
            ValidateLoanAmt_80EEA = False
            Exit Function
        End If
'         If Len(LoanAmt_80EEA(i)) > 14 Then
'          MsgBox_80EEA = MsgBox_80EEA + "* Loan Amount  at Sr. No " & i & " in Sheet 80EEA less than 15 characters." & Chr(13)
'            ValidateLoanAmt_80EEA = False
'            Exit Function
'        End If
        
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
            MsgBox_80EEA = MsgBox_80EEA & "* Total amount of Loan should be more than 0 in schedule 80EEA at Sr. No  " & i & "" & Chr(13)
            ValidateLoanAmt_80EEA = False
            Exit Function
        End If
         
Next
End Function

Function ValidateLoanOutstanding_80EEA() As Boolean
    ValidateLoanOutstanding_80EEA = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet20.Range("loanOutstanding.80EEA").Cells
    ReDim LoanOutStanding_80EEA(end_80EEA)
    For i = 1 To end_80EEA
        LoanOutStanding_80EEA(i) = rangecells.item(i).value
        If Not chkCompulsory(LoanOutStanding_80EEA(i)) Then
             MsgBox_80EEA = MsgBox_80EEA + "* ""Loan outstanding as on  last date of financial year is mandatory in schedule 80EEA"" at Sr. No " & i & "" & Chr(13)
            ValidateLoanOutstanding_80EEA = False
            Exit Function
        End If
'         If Len(LoanOutstanding_80EEA(i)) > 14 Then
'          MsgBox_80EEA = MsgBox_80EEA + "* Loan Outstanding  at Sr. No " & i & " in Sheet 80EEA less than 15 characters." & Chr(13)
'            ValidateLoanOutstanding_80EEA = False
'            Exit Function
'        End If
        
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
            MsgBox_80EEA = MsgBox_80EEA & "* Loan outstanding as on  last date of financial year can't be less than 0 in schedule 80EEA at Sr. No  " & i & ". You may please enter as 0 if it become negative as result of excess payment." & Chr(13)
            ValidateLoanOutstanding_80EEA = False
            Exit Function
        End If
         
Next
End Function

Function ValidateIntrst_80EEA() As Boolean
    ValidateIntrst_80EEA = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet20.Range("Intrst.80EEA").Cells
    ReDim Intrst_80EEA(end_80EEA)
    For i = 1 To end_80EEA
        Intrst_80EEA(i) = rangecells.item(i).value
        If Not chkCompulsory(Intrst_80EEA(i)) Then
             MsgBox_80EEA = MsgBox_80EEA + "* ""Please provide Interest u/s 80EEA"" at Sr. No " & i & "" & Chr(13)
            ValidateIntrst_80EEA = False
            Exit Function
        End If
'         If Len(Intrst_80EEA(i)) > 15 Then
'          MsgBox_80EEA = MsgBox_80EEA + "* Interest  at Sr. No " & i & " in Sheet 80EEA less than 15 characters." & Chr(13)
'            ValidateIntrst_80EEA = False
'            Exit Function
'        End If

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
        
        'Interest u/s 80EEA should be more than 0 in schedule 80EEA
        
        If Intrst_80EEA(i) < 0 Or Intrst_80EEA(i) = 0 Then
            MsgBox_80EEA = MsgBox_80EEA & "* ""Interest u/s 80EEA"" should be more than 0 in schedule 80EEA at Sr. No  " & i & "" & Chr(13)
            ValidateIntrst_80EEA = False
            Exit Function
        End If
         
Next
End Function


Function ValidateBankAccountNumber_80EEA(BankAccountNumber As Variant, cc As Long) As Boolean
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
    
   ' BankAccountNumber = Sheet5.Range("IncD.BankAccountNumber")

           
    If Len(BankAccountNumber) > 0 Then
        If Not checkfieldspecialcharacter1(BankAccountNumber) Then
            MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 80EEA, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber_80EEA = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 80EEA " & Chr(13)
            ValidateBankAccountNumber_80EEA = False
            Exit Function
        End If
    
    End If
  '----------------------------------------------------------------
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        MsgBox_80EEA = MsgBox_80EEA & "*  Please enter the Loan Account number in Bank Details at Sr.No " & cc & " in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If
    
    
    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc & " is mandatory in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If
'--------------------------------------------------------------------

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0) Or (InStr(BankAccountNumber, "-/") > 0) Or (InStr(BankAccountNumber, "/-") > 0)) Then
        MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If
    
    If (Not checkfieldspecialcharacter(Mid(BankAccountNumber, 1, 1))) Then
        MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If


    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 80EEA" & Chr(13)
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
            MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 80EEA" & Chr(13)
            ValidateBankAccountNumber_80EEA = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "* Loan Account number at Sr.No " & cc & "  in schedule 80EEA is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
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

'        If Year > 2023 Then
        'PAG_C1 AY 2024-25 Change
        'Ayush_25-26
'        If Year > 2024 Then
        If Year > 2022 Then
            ChkMaxDate_80EEA = False
            Exit Function
        Else
'            If Year = 2024 Then
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


'end_80EEALoanfrm, end_80EEAbankName, , end_80EEAAccntNum, end_80EEALoanDate, end_80EEALoanAmt, end_80EEALoanOutstanding, end_80EEAIntrst)


Sub LockUnlock80EEA_OLD()
On Error Resume Next
Application.EnableEvents = False
Sheet20.Unprotect Password:=getmsgstate

end_80EE = 0

Deduction_80EE_and_80EEA_chk = ""

md80EE.setTblinfo_80EEAccntNum
md80EE.setTblinfo_80EEBankName
'md80EE.setTblinfo_80EEIFSC
md80EE.setTblinfo_80EELoanfrm
'md80EE.setTblinfo_80EEPAN
md80EE.setTblinfo_80EEIntrst
md80EE.setTblinfo_80EELoanAmt
md80EE.setTblinfo_80EELoanDate
md80EE.setTblinfo_80EELoanOutstanding

end_80EE = WorksheetFunction.Max(0, md80EE.end_80EE, md80EE.end_80EELoanfrm, md80EE.end_80EEbankName, , md80EE.end_80EEAccntNum, md80EE.end_80EELoanDate, md80EE.end_80EELoanAmt, md80EE.end_80EELoanOutstanding, md80EE.end_80EEIntrst)



If end_80EE > 0 Then
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("LoanfrmBankOrInstitute.80EEA").ClearContents
            Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Locked = True
            Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Interior.Color = "&HD8D8D8"
           
'Commented by Ayush_05/05/2025 for 0.7 changes
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range("IFSC.80EEA").ClearContents
'            Sheet20.Range("IFSC.80EEA").Locked = True
'            Sheet20.Range("IFSC.80EEA").Interior.Color = "&HD8D8D8"
'            Sheet20.Range("IFSC.80EEA").value = ""

'Commented by Ayush_05/05/2025 for 0.7 changes
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range("PAN.80EEA").ClearContents
'            Sheet20.Range("PAN.80EEA").Locked = True
'            Sheet20.Range("PAN.80EEA").Interior.Color = "&HD8D8D8"
'            Sheet20.Range("PAN.80EEA").value = ""
            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("bankName.80EEA").ClearContents
            Sheet20.Range("bankName.80EEA").Locked = True
            Sheet20.Range("bankName.80EEA").Interior.Color = "&HD8D8D8"
            Sheet20.Range("bankName.80EEA").value = ""
            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanAccNum.80EEA").ClearContents
            Sheet20.Range("loanAccNum.80EEA").Locked = True
            Sheet20.Range("loanAccNum.80EEA").Interior.Color = "&HD8D8D8"
            Sheet20.Range("loanAccNum.80EEA").value = ""
            
           Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanDate.80EEA").ClearContents
            Sheet20.Range("loanDate.80EEA").Locked = True
            Sheet20.Range("loanDate.80EEA").Interior.Color = "&HD8D8D8"
            Sheet20.Range("loanDate.80EEA").value = ""
            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanAmt.80EEA").ClearContents
            Sheet20.Range("loanAmt.80EEA").Locked = True
            Sheet20.Range("loanAmt.80EEA").Interior.Color = "&HD8D8D8"
            Sheet20.Range("loanAmt.80EEA").value = ""
            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanOutstanding.80EEA").ClearContents
            Sheet20.Range("loanOutstanding.80EEA").Locked = True
            Sheet20.Range("loanOutstanding.80EEA").Interior.Color = "&HD8D8D8"
            Sheet20.Range("loanOutstanding.80EEA").value = ""
            
'            Sheet20.Range("loanOutstanding.80EEA").ClearContents
'            Sheet20.Range("loanOutstanding.80EEA").Locked = True
'            Sheet20.Range("loanOutstanding.80EEA").Interior.Color = "&HD8D8D8"
'            Sheet20.Range("loanOutstanding.80EEA").value = ""
    
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("Intrst.80EEA").ClearContents
            Sheet20.Range("Intrst.80EEA").Locked = True
            Sheet20.Range("Intrst.80EEA").Interior.Color = "&HD8D8D8"
            Sheet20.Range("Intrst.80EEA").value = ""
            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("Stampduty.80EEA").ClearContents
            Sheet20.Range("Stampduty.80EEA").Locked = True
            Sheet20.Range("Stampduty.80EEA").Interior.Color = "&HD8D8D8"
            Sheet20.Range("Stampduty.80EEA").value = ""

            Deduction_80EE_and_80EEA_chk = True

    
    Else
    
    
            
            Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Locked = False
            Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Interior.Color = "&HCCFFCC"
    
 'Commented by Ayush_05/05/2025 for 0.7 changes
'            Sheet20.Range("IFSC.80EEA").Locked = False
'            Sheet20.Range("IFSC.80EEA").Interior.Color = "&HCCFFCC"
'
'
'            Sheet20.Range("PAN.80EEA").Locked = False
'            Sheet20.Range("PAN.80EEA").Interior.Color = "&HCCFFCC"
            
            Sheet20.Range("bankName.80EEA").Locked = False
            Sheet20.Range("bankName.80EEA").Interior.Color = "&HCCFFCC"
            
            Sheet20.Range("loanAccNum.80EEA").Locked = False
            Sheet20.Range("loanAccNum.80EEA").Interior.Color = "&HCCFFCC"
            
            
            Sheet20.Range("loanDate.80EEA").Locked = False
            Sheet20.Range("loanDate.80EEA").Interior.Color = "&HCCFFCC"
            
            
            Sheet20.Range("loanAmt.80EEA").Locked = False
            Sheet20.Range("loanAmt.80EEA").Interior.Color = "&HCCFFCC"
            
            Sheet20.Range("loanOutstanding.80EEA").Locked = False
            Sheet20.Range("loanOutstanding.80EEA").Interior.Color = "&HCCFFCC"
            
            Sheet20.Range("Intrst.80EEA").Locked = False
            Sheet20.Range("Intrst.80EEA").Interior.Color = "&HCCFFCC"
            
    
           
            Sheet20.Range("Stampduty.80EEA").Locked = False
            Sheet20.Range("Stampduty.80EEA").Interior.Color = "&HCCFFCC"
            

            Deduction_80EE_and_80EEA_chk = False

    
    
    End If




Sheet20.Protect Password:=getmsgstate
Application.EnableEvents = True
End Sub







Sub setTblinfo_24bLoanfrm_com()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
    mIntCells = Range("LoanfrmBankOrInstitute.24b" & i).count
    Set rangecells = Range("LoanfrmBankOrInstitute.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bLoanfrm2 = ccount
End Sub

Sub setTblinfo_24bBankName_com()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
    mIntCells = Range("bankName.24b" & i).count
    Set rangecells = Range("bankName.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bbankName2 = ccount
End Sub

Sub setTblinfo_24bAccntNum_com()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
    mIntCells = Range("loanAccNum.24b" & i).count
    Set rangecells = Range("loanAccNum.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bAccntNum2 = ccount
End Sub

Sub setTblinfo_24bLoanDate_com()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
    mIntCells = Range("loanDate.24b" & i).count
    Set rangecells = Range("loanDate.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bLoanDate2 = ccount
End Sub

Sub setTblinfo_24bLoanAmt_com()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
    mIntCells = Range("loanAmt.24b" & i).count
    Set rangecells = Range("loanAmt.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bLoanAmt2 = ccount
End Sub

Sub setTblinfo_24bLoanOutstanding_com()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
    mIntCells = Range("loanOutstanding.24b" & i).count
    Set rangecells = Range("loanOutstanding.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bLoanOutstanding2 = ccount
End Sub

Sub setTblinfo_24bIntrst_com()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet22.Range("PropertySectionCOunt").value
    mIntCells = Range("Intrst.24b" & i).count
    Set rangecells = Range("Intrst.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bIntrst2 = ccount
End Sub
