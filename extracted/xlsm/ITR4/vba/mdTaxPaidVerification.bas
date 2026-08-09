Attribute VB_Name = "mdTaxPaidVerification"
Option Explicit

Public AdvanceTaxAmnt As Variant
Public TDSAmnt As Variant
Public TCSAmnt As Variant
Public SelfAssmntTax As Variant
Public TotalTaxPaid As Variant
Public TotalTaxPayable As Variant
Public RefundAmnt As Variant
Public BankAccountNum As Variant
Public IfsCode As Variant
Public TypeOfAccnt As Variant
Public AssesseName As Variant
'Public capacity As Variant
Public AssesseFatherName As Variant
Public VerificationPlace As Variant
Public VerificationDate As Variant
Public VerificationPAN As Variant
Public TRPID As Variant
Public TRPName As Variant
Public TRPreimbursementAmnt As Variant
Public TaxExmpIntInc_AIR As Variant
'Public CtCheck As Variant
Public Capacity As Variant
'Public BankAccountNumber As Variant


Public errmsgVerification As Variant
Public Others_NOI, Others_NOI1, Others_NOI22, Others_Amt, end_OthersNOI, end_OthersAmt As Variant

Public Others_NOI_BP, Others_NOI1_BP, Others_Amt_BP As Variant

'Malli-----AY_2025-26 LTCG
Public IncD_Sale_LTCG  As Variant
Public IncD_Cost_LTCG As Variant
Public IncD_CG_LTCG As Variant
Public msgError_LTCG As String
'---------------

Dim MgBx_14 As String

Sub cmdValidateVerify_Click()
On Error Resume Next
ValidateSheetTaxPaidVerificaton
fmsgboxStatus "Sheet Tax Paid and Verification is Ok" ', vbOKOnly, vbMessageCaption
End Sub

Sub cmdPrev_Click_Verify()
'Ayush
If Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" Then
       Sheet202.Activate
    Else:
Sheet17.Activate
End If
End Sub
Sub cmdNext_Click_Verify()
If Sheet8.Range("BacValue").value = 2 Then
Sheet16.Activate

'ElseIf Sheet1.Range("IncD.TypeOfHP").value <> "Self Occupied" And Sheet8.Range("bacvalue").value <> 1 Then
'Sheet21.Activate
Else
MsgBox "Can Not go to next sheet as you have selected 115BAC = Y "
End If
End Sub
Sub AddRows_Others()
    Dim vRows As Long
    Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Sheets("Taxes Paid and Verification")
    sourceSheet.Activate
    mdCommon.DefinedgridNameRange = "Sheet5.Nature||Sheet5.Description||Sheet5.Amount||SEC38Check"
    ActiveCellRange = mdCommon.searchLastRow("Sheet5.Nature")
    vRows = mdCommon.insertRowUnderSectionWithFormula
End Sub

Sub Sheet1AddRows_Others()
    Dim vRows As Long
    Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Sheets("Income Details")
    sourceSheet.Activate
    mdCommon.DefinedgridNameRange = "Sheet1.Nature||Sheet1.Description||Sheet1.Amount" '||SEC38Check"
    ActiveCellRange = mdCommon.searchLastRow("Sheet1.Nature")
    vRows = mdCommon.insertRowUnderSectionWithFormula
End Sub
Sub Sheet1AddRows_Others1()
    Dim vRows As Long
    Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Sheets("Income Details")
    sourceSheet.Activate
    mdCommon.DefinedgridNameRange = "Sheet1.Nature1||Sheet1.Description1||Sheet1.Amount1||Sheet1.Amount11||Sheet1.Amount22||Sheet1.Amount23||Sheet1.Amount24||Sheet1.Amount25" '||SEC38Check"
    ActiveCellRange = mdCommon.searchLastRow("Sheet1.Nature1")
    vRows = mdCommon.insertRowUnderSectionWithFormula
End Sub
Function ValidateOthersEI() As Boolean
ValidateOthersEI = True

    If Len((Range("Sheet5.Nature").item(1).value) > 0) Then
        If Not ValidateNatureOfIncome Then ValidateOthersEI = False
        'Ayush_05/07/2026
        If Not ValidateSub_category Then ValidateOthersEI = False
        If Not ValidateAmount Then ValidateOthersEI = False
    End If
     
    setTblinfo_OthersNOI
    setTblinfo_OthersAmt
   
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim cellrange As String
    Dim cellRange1 As String
    Dim i As Long
    
    Set rangecells = Range("Sheet5.Nature").Cells
 
    ReDim Others_NOI(end_OthersAmt)

    For i = 1 To end_OthersAmt
   
    cellrange = GetMergedAddressCell(rangecells, i)
   
    Others_NOI(i) = Sheet5.Range(cellrange).value

        
            If Others_NOI(i) = "(Select)" Or Others_NOI(i) = "" Then
'             msgError = msgError & "Please select an option from the drop down in Exempt Income at Sr. No  " & i & "  in Sheet Taxes paid and verification" & Chr(13)
                 msgError = msgError & "* Please select category from the drop down in Exempt Income in Sheet Taxes paid and verification" & Chr(13)
             ValidateOthersEI = False
             Exit Function
         End If

   Next


End Function

Sub setTblinfo_OthersNOI()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sheet5.Nature").count
    Set rangecells = Range("Sheet5.Nature").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
 end_OthersNOI = ccount
 End Sub
 
 Sub setTblinfo_OthersAmt()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sheet5.Amount").count
    Set rangecells = Range("Sheet5.Amount").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_OthersAmt = ccount
End Sub

Function ValidateNatureOfIncome() As Boolean
ValidateNatureOfIncome = False

    setTblinfo_OthersNOI
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    
    Set rangecells = Range("Sheet5.Nature").Cells
    Set rangecells1 = Range("Sheet5.Description").Cells
    Set rangecells2 = Range("Sheet5.Amount").Cells
    Set rangecells3 = Sheet5.Range("Sheet5.DescEI").Cells

    Dim i As Long
    ReDim Others_NOI(end_OthersNOI)
    ReDim Others_NOI1(end_OthersNOI)
    ReDim Others_Amt(end_OthersNOI)
    ReDim Others_NOI22(end_OthersNOI)
    
    For i = 1 To end_OthersNOI
        Others_NOI(i) = rangecells.item(i).value
        Others_NOI1(i) = rangecells1.item(i).value
        Others_Amt(i) = rangecells2.item(i).value
        Others_NOI22(i) = rangecells3.item(i).value
        
'         If (Others_NOI(i) = "(Select)" Or Others_NOI(i) = "") Then
'             msgError = msgError & "* Please select category from the drop down in Exempt Income in Sheet Taxes paid and verification" & Chr(13)
'             ValidateNatureOfIncome = False
'             Exit Function
'         End If
            
         
         
         If Others_NOI1(i) = "" Or Others_NOI1(i) = "(Select)" Then
             msgError = msgError & "* Please select sub-category from the drop down in Exempt Income in Sheet Taxes paid and verification" & Chr(13)
             ValidateNatureOfIncome = False
             Exit Function
         End If
                  
'Added by Riyaz on 05/03/2026
   ' If Mid(Others_NOI(i), 1, 4) = "Agri" Then
   If Mid(Others_NOI1(i), 1, 5) = "10(1)" Then
        If Others_Amt(i) > 5000 Then
         '   msgError = msgError & "* If agricultural income is more than Rs 5000/- then use ITR 3 or 5" & Chr(13)
          msgError = msgError & "*If sec 10(1) agricultural income is more than Rs 5000/- then use ITR 3 or 5" & Chr(13)
            ValidateNatureOfIncome = False
            Exit Function
        End If
    End If
        
'         If Len(Others_NOI1(i)) > 125 Then
'             msgError = msgError & "* Description at Sr. No  " & i & " cannot exceed 125 characters in Sheet Taxes paid and verification" & Chr(13)
'             ValidateNatureOfIncome = False
'             Exit Function
'         End If
'    Next
    'Ankita_18/06/2026======
    
          If Others_NOI1(i) = "Income exempt as per CBDT Circular" Then
             If Others_NOI22(i) = "" Then
                msgError = msgError & """Please enter the relevant Circular for sub-category “Income exempt as per CBDT Circular” in Exempt Income at Sl.No. """ & i & Chr(13)
                ValidateNatureOfIncome = False
             End If
        End If
          If Others_NOI1(i) = "Income exempt as per CBDT Notification" Then
             If Others_NOI22(i) = "" Then
                msgError = msgError & """Please enter the relevant Notification No. for other exempt income for sub-category “Income exempt as per CBDT Notification” in Exempt Income at Sl.No. """ & i & Chr(13)
                ValidateNatureOfIncome = False
             End If
        End If
        If Others_NOI1(i) = "Receipts not in the nature of income" Then
               If Others_NOI22(i) = "" Then
                  msgError = msgError & """Please enter the relevant Act/Section reference for sub-category “Receipts not in the nature of income” in Exempt Income at Sl.No. """ & i & Chr(13)
                  ValidateNatureOfIncome = False
               End If
        End If
    Next
End Function


Function ValidateAmount() As Boolean
ValidateAmount = False

    setTblinfo_OthersNOI
    Dim rangecells As Range
    Set rangecells = Range("Sheet5.Amount").Cells
    Dim i As Long
    ReDim Others_Amt(end_OthersNOI)
    For i = 1 To end_OthersNOI
        Others_Amt(i) = rangecells.item(i).value
        
        If Not chkCompulsory(Others_Amt(i)) Then
            msgError = msgError & "* Please enter Amount of exempt income in Sheet Taxes paid and verification" & Chr(13)
            ValidateAmount = False
            Exit Function
        End If
                
        If Not IsNumeric(Others_Amt(i)) Then
            msgError = msgError & "* Others Amount at Sr. No  " & i & "  in Sheet Taxes paid and verificationshould be Numeric value" & Chr(13)
            ValidateAmount = False
            Exit Function
        End If
        
        If Others_Amt(i) > 99999999999999# Then
            msgError = msgError & "* Others Amount at Sr. No  " & i & "  in Sheet Taxes paid and verification cannot exceed 14 digits" & Chr(13)
            ValidateAmount = False
            Exit Function
        End If
    Next
End Function
Sub ValidateSheetTaxPaidVerificaton()
msgError = ""
AssignValueTaxandVerification

If Not CheckifFieldexceed14digits Then
Sheet5.Activate
fmsgboxsmall MgBx_14 ', vbOKOnly, "Error(s)!"
End If


'Malli-----AY_2025_26
If Not validateLTCG() Then
Sheet5.Activate
fmsgboxsmall_LTCG msgError_LTCG
CloseMsg
End If
'----------------------

ValidateBA

If Not validateSheetTaxandVerification() Then
Sheet5.Activate
fmsgboxsmall msgError ' , vbOKOnly, "Error(s)!"
CloseMsg
End If
End Sub

Function AssignValueTaxandVerification()
On Error Resume Next

AdvanceTaxAmnt = Trim(Range("IncD.AdvanceTax").value)
TDSAmnt = Trim(Range("IncD.TDS").value)
TCSAmnt = Trim(Range("IncD.TCS").value)
SelfAssmntTax = Trim(Range("IncD.SelfAssessmentTax").value)
TotalTaxPaid = Trim(Range("IncD.TotalTaxesPaid").value)
TotalTaxPayable = Trim(Range("IncD.BalTaxPayable").value)
RefundAmnt = Trim(Range("IncD.RefundDue").value)
BankAccountNum = Trim(Range("IncD.BankAccountNumber").value)
IfsCode = Trim(Range("IncD.MICRCode").value)
TypeOfAccnt = Trim(Range("IncD.BankAccountType").value)
AssesseName = Trim(Range("Ver.AssesseeVerName").value)
AssesseFatherName = Trim(Range("Ver.FatherName").value)
'capacity = Trim(Range("Ver.Capacity").value)
VerificationPlace = Trim(Range("Ver.Place").value)
VerificationDate = Trim(Range("Ver.Date").value)
VerificationPAN = Trim(Range("Ver.PAN").value)
TRPID = Trim(Range("Ver.IdentificationNoOfTRP").value)
TRPName = Trim(Range("Ver.NameOfTRP").value)
TRPreimbursementAmnt = Trim(Range("Ver.ReImbFrmGov").value)
TaxExmpIntInc_AIR = Trim(Range("AIR.TaxExmpIntInc").value)
'CtCheck = Trim(Range("Sheet5.CTcheck").value)
Capacity = Trim(Range("Ver.Capacity").value)

'Malli--AY_2025_26 LTCG
IncD_Sale_LTCG = Trim(Range("IncD.Sale_LTCG").value)
IncD_Cost_LTCG = Trim(Range("IncD.Cost_LTCG").value)
IncD_CG_LTCG = Trim(Range("IncD.CG_LTCG").value)
'----------------

End Function


Function CheckifFieldexceed14digits() As Boolean
On Error Resume Next
CheckifFieldexceed14digits = True

MgBx_14 = ""

If Sheet5.Range("IncD.AdvanceTax").value > 99999999999999# Then
    CheckifFieldexceed14digits = False
    MgBx_14 = MgBx_14 + "* Advance Tax cannot exceed 14 digits" & Chr(13)
End If


If Sheet5.Range("IncD.TDS").value > 99999999999999# Then
    CheckifFieldexceed14digits = False
    MgBx_14 = MgBx_14 + "* TDS cannot exceed 14 digits" & Chr(13)
End If

If Sheet5.Range("IncD.TCS").value > 99999999999999# Then
    MgBx_14 = MgBx_14 + "* TCS cannot exceed 14 digits" & Chr(13)
End If

If Sheet5.Range("IncD.SelfAssessmentTax").value > 99999999999999# Then
    CheckifFieldexceed14digits = False
    MgBx_14 = MgBx_14 + "* Self Assessment Tax cannot exceed 14 digits" & Chr(13)
End If


If Sheet5.Range("IncD.TotalTaxesPaid").value > 99999999999999# Then
    CheckifFieldexceed14digits = False
    MgBx_14 = MgBx_14 + "* Total Tax Paid cannot exceed 14 digits" & Chr(13)
End If

If Sheet5.Range("IncD.BalTaxPayable").value > 99999999999999# Then
    CheckifFieldexceed14digits = False
    MgBx_14 = MgBx_14 + "* Tax Payable cannot exceed 14 digits" & Chr(13)
End If

If Sheet5.Range("IncD.RefundDue").value > 99999999999999# Then
    CheckifFieldexceed14digits = False
    MgBx_14 = MgBx_14 + "* Refund cannot exceed 14 digits" & Chr(13)
End If


If Len(Sheet5.Range("Ver.AssesseeVerName").value) > 127 Then
    CheckifFieldexceed14digits = False
    MgBx_14 = MgBx_14 + "* Verification Name cannot exceed 127 charcters" & Chr(13)
End If

If Len(Sheet5.Range("Ver.FatherName").value) > 125 Then
    CheckifFieldexceed14digits = False
    MgBx_14 = MgBx_14 + "* Father Name cannot exceed 125 charcters" & Chr(13)
End If
'If Len(Sheet5.Range("Ver.Capacity").value) > 125 Then
'    CheckifFieldexceed14digits = False
'    MgBx_14 = MgBx_14 + "Capacity cannot exceed 125 charcters" & Chr(13)
'End If

If Len(Sheet5.Range("Ver.Place").value) > 50 Then
    CheckifFieldexceed14digits = False
    MgBx_14 = MgBx_14 + "* Place Name in Verification cannot exceed 50 charcters" & Chr(13)
End If


'Newly added by Bindu for LTCG on 3rd Feb 2025 as per DE sheet version 3
If Sheet5.Range("IncD.Sale_LTCG").value > 99999999999999# Then
    CheckifFieldexceed14digits = False
    MgBx_14 = MgBx_14 + "* Total sale consideration cannot exceed 14 digits" & Chr(13)
End If

If Sheet5.Range("IncD.Cost_LTCG").value > 99999999999999# Then
    CheckifFieldexceed14digits = False
    MgBx_14 = MgBx_14 + "* Total cost of acquisition cannot exceed 14 digits" & Chr(13)
End If


If Sheet5.Range("IncD.CG_LTCG").value > 99999999999999# Then
    CheckifFieldexceed14digits = False
    MgBx_14 = MgBx_14 + "* Long term capital gains as per sec 112A cannot exceed 14 digits" & Chr(13)
End If

End Function


Function validateSheetTaxandVerification() As Boolean
On Error Resume Next
validateSheetTaxandVerification = True
subProcCaption = "Validating Tax and Verification"
noOfProcessSub = 4


ValidateOthersEI

'If Sheet5.Range("Agri_income").value > 5000 Then
'msgError = msgError & "* If agricultural income is more than Rs 5000/- then use ITR 3 or 5" & Chr(13)
'End If



'If capacity = "Representative" And Representativeassesseeflg = "" Then
'msgError = msgError & " Details of representative assessee is mandatory as the return is being filed by representative" & Chr(13)
'End If

If AssesseName = "" Then msgError = msgError & "* Assessee Name in sheet Tax Paid and Verification is Mandatory" & Chr(13)
If AssesseFatherName = "" Then msgError = msgError & "* Assessee Father Name in sheet Tax Paid and Verification is Mandatory" & Chr(13)
If (Capacity = "" Or isdropdownblank(Capacity)) Then msgError = msgError & "* Selection of capacity in verification Part is mandatory. Please select appropriate option from drop down " & Chr(13)

''Ayush_17/06
'If Sheet1.Range("sheet1.RepAssessee").value = "Yes" Then
'If Capacity <> "Representative" Then
'msgError = msgError & "* Please select ""Representative"" if return is filed by Represenattive." & Chr(13)
'End If
'End If
'Ayush_SIT-111524
If Dformat(Trim(Range("sheet1.DOB").value), "yyyy-mm-dd") >= Dformat(Trim("01/04/2008"), "yyyy-mm-dd") And Sheet5.Range("Ver.Capacity").value = "Self" Then
    msgError = msgError & "* Since Date of Birth is on or after 01/04/2008 the return cannot be verified through ""Self"", please verify the return through ""Representative""." & Chr(13)
End If

If Dformat(Trim(Range("sheet1.DOB").value), "yyyy-mm-dd") >= Dformat(Trim("01/04/2008"), "yyyy-mm-dd") And Sheet5.Range("Ver.PAN").value = Sheet1.Range("sheet1.PAN").value Then
    msgError = msgError & "* If date of birth is on or after 1 April 2008 then in verification PAN cannot be same as mentioned in general information." & Chr(13)
End If


If (UCase(Capacity) <> "SELF" And UCase(Capacity) <> "REPRESENTATIVE" And UCase(Capacity) <> "KARTA" And UCase(Capacity) <> "PARTNER" And UCase(Capacity) <> "(SELECT)") Then msgError = msgError & "* Please select valid option from dropdown for the capacity in Verification." & Chr(13)

If VerificationPlace = "" Then msgError = msgError & "* Place in sheet Tax Paid and Verification is Mandatory" & Chr(13)
'If (Mid(CtCheck, 1, 1) = "" Or Mid(CtCheck, 1, 1) = "(") Then msgError = msgError & "Please select an option from the drop down" & Chr(13)

If VerificationDate = "" Then
    msgError = msgError & "* Date in sheet Tax Paid and Verification is Mandatory" & Chr(13)
Else
    If Not FormatNCheckDate(VerificationDate) Then
        msgError = msgError & "* Please enter a valid date in DD/MM/YYYY format." & Chr(13)
    Else
    If Not ChkMinInclusiveDate(Dformat(VerificationDate, "yyyy-mm-dd"), "2024-04-01") Then msgError = msgError & "* Date cannot be before 01/04/2024 in sheet Tax Paid and Verification." & Chr(13)
    End If
End If
UpdateProgressBar

'If (Range("EI.Sec10_34").value) > 1000000 Then msgError = msgError + "Sec.10(34) shall not be more than Rs.1000000" & Chr(13)

If VerificationPAN = "" Then
    msgError = msgError & "* PAN in sheet Tax Paid and Verification is Mandatory" & Chr(13)
Else
    If Not CheckVerPAN(UCase(VerificationPAN)) Then msgError = msgError & "* Invalid PAN" & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet in Verifcation"
    'msgError = msgError & "* Invalid PAN" & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet in Verifcation"
End If

UpdateProgressBar


'            If calculateAge(Trim(Range("sheet1.DOB").value)) < 18 And Mid(Range("sheet1.status").value, 1, 1) = "I" Then
'                If Sheet5.Range("Ver.PAN").value = Sheet1.Range("sheet1.PAN").value Then
'                    msgError = msgError & "*If status is individual & date of birth is on or after 2 April 2002 then in verification PAN cannot be same as mentioned in general information."
'                     'Sheet5.Range("Ver.PAN").value = ""
'                End If
'            End If
'
'            If calculateAge(Trim(Range("sheet1.DOB").value)) < 18 And Mid(Range("sheet1.status").value, 1, 1) = "I" Then
'                If Sheet5.Range("Ver.Capacity").value <> "Representative" Then
'                    msgError = msgError & "*If status is individual & date of birth is on or after 2 April 2002 then in verification drop down cannot be selected as self/Karta/ Partner"
'                    'Sheet5.Range("Ver.Capacity").value = ""
'                End If
'            End If

If TRPID <> "" And TRPName = "" Then msgError = msgError & "* Please enter TRP name" & Chr(13)
If TRPName <> "" And TRPID = "" Then msgError = msgError & "* Please enter TRP ID" & Chr(13)
If TRPID <> "" Then
    If Not ChkTRPID(TRPID) Then msgError = msgError & "* Please Enter a valid TRP ID"
End If
UpdateProgressBar
If ((TRPreimbursementAmnt <> "0" And TRPreimbursementAmnt <> "") And (TRPName = "")) Then msgError = msgError & "* Please enter TRP Name in Verifcation" & Chr(13)
If ((TRPreimbursementAmnt <> "" And TRPreimbursementAmnt <> "0") And (TRPID = "")) Then msgError = msgError & "* Please enter TRP ID in Verifcation" & Chr(13)

If (Range("Ver.Capacity").value) = "Representative" And (Sheet1.Range("sheet1.RepAssessee").value = "" Or UCase(Sheet1.Range("sheet1.RepAssessee").value) = UCase("(Select)")) Then


         msgError = msgError & "* Details of representative assessee is mandatory as the return is being filed by representative in Sheet Income Details" & Chr(13)
End If


If Sheet1.Range("sheet1.RepAssessee").value = "No" Then
'If Capacity <> "Representative" Then

If Mid(Sheet1.Range("sheet1.Status"), 1, 1) = "I" And (UCase(Capacity) <> "SELF" And UCase(Capacity) <> UCase("(Select)")) Then
    msgError = msgError & "*  Please select capacity as ""Self"" if return is filed with status ""Individual""." & Chr(13)
                
ElseIf Mid(Sheet1.Range("sheet1.Status"), 1, 1) = "H" And (UCase(Capacity) <> "KARTA" And UCase(Capacity) <> UCase("(Select)")) Then
        msgError = msgError & "*  Please select ""Karta"" if return is filed with status ""HUF""." & Chr(13)
        
ElseIf Mid(Sheet1.Range("sheet1.Status"), 1, 1) = "F" And (UCase(Capacity) <> "PARTNER" And UCase(Capacity) <> UCase("(Select)")) Then

 msgError = msgError & "*  Please select ""Partner"" if return is filed with status ""Firm""." & Chr(13)
 End If
 
Else
If Capacity <> "Representative" And UCase(Capacity) <> UCase("(Select)") Then
    msgError = msgError & "* Please select ""Representative"" if return is filed by Representative." & Chr(13)
 End If





End If

If msgError <> "" Then validateSheetTaxandVerification = False

UpdateProgressBar

End Function

Function ChkTRPID(item As Variant) As Boolean
On Error Resume Next
Dim i As Long
ChkTRPID = True

If Len(item) = 10 Then
    
    If UCase(Mid(item, 1, 1)) <> "T" Then
        ChkTRPID = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(item, 2, 9)) Then
        ChkTRPID = False
        Exit Function
    End If
    
ElseIf Len(item) = 6 Then

    If Not IsNumeric(Mid(item, 1, 6)) Then
        ChkTRPID = False
        Exit Function
    End If
    
Else
    ChkTRPID = False
    Exit Function
End If
End Function


Function ValidateBankAccountNumber(BankAccountNumber As Variant) As Boolean
    ValidateBankAccountNumber = True
    Dim numfound As Boolean
    Dim countnum As Long
    Dim myB() As Variant
    Dim i As Long
    Dim zeroCount As Long
    Dim BeforeZero, AfterZero As String
    errmsgVerification = ""
    numfound = False
    countnum = 0
    BeforeZero = ""
    AfterZero = ""
    zeroCount = 1
    
    'BankAccountNumber = Sheet5.Range("IncD.BankAccountNumber")
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        errmsgVerification = errmsgVerification & "* Bank Account Number is Mandatory in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber = False
        Exit Function
    End If

    If Len(BankAccountNumber) > 0 Then
        If Not checkfieldspecialcharacter1(BankAccountNumber) Then
            errmsgVerification = errmsgVerification & "* Bank Account Number is Invalid in Sheet Taxes Paid and Verification, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            ValidateBankAccountNumber = False
            errmsgVerification = errmsgVerification & "* Bank Account Number is invalid in sheet Taxes Paid and Verification" & Chr(13)
            Exit Function
        End If
    
    End If

    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        errmsgVerification = errmsgVerification & "* Bank Account Number in Sheet Taxes Paid and Verification is Compulsory for all cases" & Chr(13)
        ValidateBankAccountNumber = False
        Exit Function
    End If

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0) Or (InStr(BankAccountNumber, "-/") > 0) Or (InStr(BankAccountNumber, "/-") > 0)) Then
        errmsgVerification = errmsgVerification & "* Bank Account Number is invalid in sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        errmsgVerification = errmsgVerification & "* Bank Account Number is Invalid in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber = False
        Exit Function
    End If
    
    If (Not checkfieldspecialcharacter(Mid(BankAccountNumber, 1, 1))) Then
        errmsgVerification = errmsgVerification & "* Bank Account Number is invalid in sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber = False
        Exit Function
    End If


    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        errmsgVerification = errmsgVerification & "* Bank Account Number is invalid in sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        errmsgVerification = errmsgVerification & "* Bank Account Number is invalid in sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber = False
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
            errmsgVerification = errmsgVerification & "* Bank Account Number is invalid in sheet Taxes Paid and Verification" & Chr(13)
            ValidateBankAccountNumber = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "* Bank Account Number in Sheet Taxes Paid and Verification is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber = False
        Exit Function
    End If
End Function


Function checkfieldspecialcharacter1(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter1 = True
    Dim arr As Variant
    arr = Array(".", "@", "*", "!", "&", "#", "~", ";", "?", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<", "_", "|")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter1 = False
            Exit Function
        End If
        Next
    Next
End Function
Function checkfieldspecialcharacter_Transaction(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter_Transaction = True
    Dim arr As Variant
    arr = Array(".", "@", "*", "!", "&", "#", "~", ";", "?", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<", "_", "|")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter_Transaction = False
            Exit Function
        End If
        Next
    Next
End Function

''AY_2025-26 Newly added by Malli on 31st Jan 2025 as per V3 DE sheet

Function validateLTCG() As Boolean
On Error Resume Next
validateLTCG = True

msgError_LTCG = ""


If (IncD_Sale_LTCG > 0 And IncD_Sale_LTCG <> "") Or (IncD_Cost_LTCG > 0 And IncD_Cost_LTCG <> "") Then
        'Change by Ayush on 18Feb for SIT-88367
'        If IncD_Sale_LTCG > 0 And (IncD_Cost_LTCG = "" Or IncD_Cost_LTCG = 0) Then
         If IncD_Sale_LTCG > 0 And (IncD_Cost_LTCG = "") Then
            validateLTCG = False
            msgError_LTCG = msgError_LTCG & "* ""Please fill Total cost of acquisition""" & Chr(13)
            Exit Function
        End If
        'Change by Ayush on 18Feb for SIT-88367
'        If IncD_Cost_LTCG > 0 And (IncD_Sale_LTCG = "" Or IncD_Sale_LTCG = 0) Then
        If IncD_Cost_LTCG > 0 And (IncD_Sale_LTCG = "") Then
            validateLTCG = False
            msgError_LTCG = msgError_LTCG & "* ""Please fill Total sale consideration """ & Chr(13)
            Exit Function
        End If
        
        'Updated by Ayush as per DE V3 of ITR1 post confirmation with BA
        If (IncD_Sale_LTCG - IncD_Cost_LTCG) < 0 Then
            fmsgboxsmall """To avail the benefit of carry forward and set of loss, please use ITR -3/5"""
        End If

End If

'IncD_Sale_LTCG
If IncD_Sale_LTCG <> "" And IncD_Sale_LTCG > 0 Then
    If Not IsNumeric(IncD_Sale_LTCG) Or IncD_Sale_LTCG > 99999999999999# Or IncD_Sale_LTCG < 0 Then
        validateLTCG = False
        msgError_LTCG = msgError_LTCG & "* ""Amount should be numeric, Non negative, no decimal, upto 99,999,999,999,999 at Total sale consideration""" & Chr(13)
        Sheet5.Range("IncD.Sale_LTCG").value = ""
        'Exit Function
    End If
End If
            
'IncD_Cost_LTCG
If IncD_Cost_LTCG <> "" And IncD_Cost_LTCG > 0 Then
    If Not IsNumeric(IncD_Cost_LTCG) Or IncD_Cost_LTCG > 99999999999999# Or IncD_Cost_LTCG < 0 Then
        validateLTCG = False
        msgError_LTCG = msgError_LTCG & "* ""Amount should be numeric, Non negative, no decimal, upto 99,999,999,999,999 at Total cost of acquisition""" & Chr(13)
        Sheet5.Range("IncD.Cost_LTCG").value = ""
        ' Exit Function
    End If
End If


            
'IncD_CG_LTCG
Dim IncD_Sale_Cost_LTCG As Variant
IncD_Sale_Cost_LTCG = IncD_Sale_LTCG - IncD_Cost_LTCG


If IncD_Sale_Cost_LTCG > 125000 Then
    validateLTCG = False
    msgError_LTCG = msgError_LTCG & "* ""In ITR 4, the maximum gains as per Section 112A can be INR 1,25,000/-. Please file ITR 3/5 if gains u/s 112A is more than INR 1,25,000/-""" & Chr(13)
    Sheet5.Range("IncD.Cost_LTCG").value = ""
    Exit Function
End If
UpdateProgressBar

End Function
Function ValidateSub_category() As Boolean
ValidateSub_category = False

    setTblinfo_OthersNOI
    Dim rangecells As Range
    Set rangecells = Range("Sheet5.Description").Cells
    Dim i As Long
    ReDim Others_sub_cat(end_OthersNOI)
    For i = 1 To end_OthersNOI
        Others_sub_cat(i) = rangecells.item(i).value
        
        
        If Sheet8.Range("bacvalue") = 1 Then
       If Mid(Others_sub_cat(i), 1, 6) = "10(32)" Then
            msgError = msgError & "* Sub category ""10(32)-Minor child's income—small exemption"" is not applicable for new tax regime" & Chr(13)
            Others_sub_cat(i) = "(Select)"
            ValidateSub_category = False
            
            Exit Function
        End If
        End If
    Next
End Function
