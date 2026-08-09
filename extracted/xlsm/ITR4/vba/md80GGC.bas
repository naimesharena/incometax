Attribute VB_Name = "md80GGC"
Option Explicit

Public MsgBox_80GGC As String

Public DateofDonation_80GGC As Variant
'Public NatureofTransaction_80GGC As Variant
Public Donation_cash_80GGC As Variant
Public Donation_other_80GGC As Variant
Public Donation_total_80GGC As Variant
Public Donation_Eligible_80GGC As Variant
Public ChequeNumber_80GGC As Variant
Public BankAccntnum_80GGC, BankIFSC_80GGC, BankName_80GGC As Variant
Public Total_DonationInCash_80GGC, Total_DonationInOThMode_80GGC, Total_TotalDonation_80GGC As Variant

Public rngname_80GGC As Variant
Public Total_Donation_Eligible_80GGC As Variant

Public end80GGC, end80GGC1, end80GGC2, end80GGC3, end80GGC4, end80GGC5, end80GGC6, end80GGC7, end80GGC8 As Variant 'As Long
'----------

Sub cmdNext80GGC_Click()
If Sheet20.Visible = True Then
    Sheet20.Activate
ElseIf Sheet19.Visible = True Then
    Sheet19.Activate

End If

End Sub

 Sub pppp()
 Application.EnableEvents = True
 End Sub
Sub Prev80GGC_Click()
Sheet14.Activate
End Sub

Sub Cmd80GGC()
    Sheets("80GGC").Activate
'    mdCommon.DefinedgridNameRange = "SCH_80GGC_DonationDate||SCH_80GGC_DonationAmtCash ||SCH_80GGC_DonationAmtOtherMode||SCH_80GGC_TransactionRefNumUPI||SCH_80GGC_IFSCCode||SCH_80GGC_BankName||SCH_80GGC_BankAccountNo||SCH_80GGC_ChnImpsNeftRtgsNum||SCH_80GGC_DonationAmt||SCH_80GGC_EligibleDonationAmt"
     mdCommon.DefinedgridNameRange = "SCH_80GGC_DonationDate||SCH_80GGC_DonationAmtCash ||SCH_80GGC_DonationAmtOtherMode||SCH_80GGC_IFSCCode"
     ||SCH_80GGC_ChnImpsNeftRtgsNum||SCH_80GGC_DonationAmt||SCH_80GGC_EligibleDonationAmt"
    'ActiveCellRange = mdCommon.searchLastRow("SCH_80GGC_DonationDate")
    ActiveCellRange = mdCommon.searchLastRow("SCH_80GGC_DonationDate")
    'insertRowUnderSectionWithFormula80GGC
    'mdCommon.insertRowUnderSectionWithFormula
    mdCommon.insertRowUnderSectionWithFormula80GGC HOIflag:=1
End Sub


Function CheckIFSC2(Tval As Variant, Tadd As Variant)
On Error Resume Next
Application.EnableEvents = False
 
 
If Tval <> "" Then
 
      
       If Not CheckIFSC(UCase(Tval)) Then
           'fmsgboxsmall "* Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
Msg80GGC = Msg80GGC + "* IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets in Bank Details at Sr.No " & i & " in Sheet 80GGC" & Chr(13)
           Sheet11.Range(Tadd) = ""
           Sheet11.Range(Tadd).Offset(0, 1) = ""
           GoTo endfd
     End If
 
      
     If Not ValidateIFSCList(UCase(Tval)) Then
            'fmsgboxsmall "* Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
Msg80GGC = Msg80GGC + "* IFSC Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets in Bank Details at Sr.No " & i & " in Sheet 80GGC" & Chr(13)
            Sheet11.Range(Tadd) = ""
            Sheet11.Range(Tadd).Offset(0, 1) = ""
            GoTo endfd
      End If
    Sheet11.Range(Tadd).Select
    Sheet11.Range(Tadd).Offset(0, 1) = UCase(Sheet11.Range(Tadd).Offset(0, 10))
End If
If Tval = "" Then
    Sheet11.Range(Tadd).Offset(0, 1) = ""
End If
 
endfd:
 
Application.EnableEvents = True
 
End Function


Sub ValidateSheet80GGC_Click()
 Dim vbMessgaeCaption As String
 vbMessgaeCaption = "ITR 1: AY: 2024-25"
Validate80GGC
'MsgBox "Sheet 80GGC is OK", vbOKOnly, vbMessgaeCaption
fmsgboxoK "Sheet 80GGC is OK"
End Sub
'--------------------------------------------------------------------------
Sub Validate80GGC()
    If Not Validate80GGC_1 Then
        Sheet11.Activate
        fmsgbox (MsgBox_80GGC)
        CloseMsg
    End If
End Sub
'-------------------------------------------------------------
Sub setTableInfo80GGC() '80GGC_C1
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet11.Range("SCH_80GGC_DonationDate").Cells
    mIntCells = Sheet11.Range("SCH_80GGC_DonationDate").count
    For mIntCtr = 1 To mIntCells
         If Not rangecells.item(mIntCtr).value = "" Then
          ccount = ccount + 1
        End If
    Next
    end80GGC = ccount
    rngname_80GGC = "SCH_80GGC_DonationDate;SCH_80GGC_DonationAmtCash;SCH_80GGC_DonationAmtOtherMode;SCH_80GGC_DonationAmt;SCH_80GGC_EligibleDonationAmt;SCH_80GGC_ChnImpsNeftRtgsNum;SCH_80GGC_IFSCCode;SCH_80GGC_Political_Name;SCH_80GGC_Political_PAN"
'------------------------------------------------------------------------------------------

End Sub
Sub setTableInfo80GGC1() '80GGC_C2
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet11.Range("SCH_80GGC_DonationAmtCash").Cells
    mIntCells = Sheet11.Range("SCH_80GGC_DonationAmtCash").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGC1 = ccount
End Sub
'-----------------------------------------------------------------
Sub setTableInfo80GGC2() '80GGC_C3
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet11.Range("SCH_80GGC_DonationAmtOtherMode").Cells
    mIntCells = Sheet11.Range("SCH_80GGC_DonationAmtOtherMode").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGC2 = ccount
End Sub

Sub setTableInfo80GGC4() '80GGC_C5
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet11.Range("SCH_80GGC_ChnImpsNeftRtgsNum").Cells
    mIntCells = Sheet11.Range("SCH_80GGC_ChnImpsNeftRtgsNum").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end80GGC4 = ccount
End Sub
'----------------------------------------------------------------------------------------------------------
Sub setTableInfo80GGC5() '80GGC_C6
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet11.Range("SCH_80GGC_IFSCCode").Cells
    mIntCells = Sheet11.Range("SCH_80GGC_IFSCCode").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGC5 = ccount
End Sub
'Ayush_15/03/2026
Sub setTableInfo80GGC6()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet11.Range("SCH_80GGC_Political_Name").Cells
    mIntCells = Sheet11.Range("SCH_80GGC_Political_Name").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGC6 = ccount
End Sub
Sub setTableInfo80GGC7()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet11.Range("SCH_80GGC_Political_PAN").Cells
    mIntCells = Sheet11.Range("SCH_80GGC_Political_PAN").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGC7 = ccount
End Sub


Sub AddRows80GGC() '80CCG
    Dim vRows  As Long
    Sheets("80GGC").Activate
'    EfilingCommon.DefinedgridNameRange = "||DateofDonation_80GGC||Donationincash_80GGC||Donationinothermode_80GGC||TotalDonation_80GGC||EligibleAmountofDonation_80GGC||NatureofTransaction_80GGC||Chequeno_80GGC||IFSC_80GGC||NameofDonor_80GGC||AccountofDonor_80GGC||"
    EfilingCommon.DefinedgridNameRange = "||DateofDonation_80GGC||Donationincash_80GGC||Donationinothermode_80GGC||TotalDonation_80GGC||EligibleAmountofDonation_80GGC||Chequeno_80GGC||IFSC_80GGC||"
    ActiveCellRange = EfilingCommon.searchLastRow("DateofDonation_80GGC")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
'-------------------------------------------------------------------------------
Function CheckIFSC_80GGC(Tval As Variant, Tadd As Variant) '80GGC_C5
On Error Resume Next
Application.EnableEvents = False
 'sPassword = getmsgstate
   Sheet11.Unprotect Password:=getmsgstate


If Tval <> "" Then
    
     If Not CheckIFSC(UCase(Tval)) Then
        fmsgbox "Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets in Bank Details in Schedule 80GGC Sheet."
           Sheet11.Range(Tadd) = ""
           'Sheet11.Range(Tadd).Offset(0, 1) = ""
           GoTo endfd
     End If
    
    If Not ValidateIFSCList(UCase(Tval)) Then
           fmsgbox "Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets in Bank Details in Schedule 80GGC Sheet."
           Sheet11.Range(Tadd) = ""
           'Sheet11.Range(Tadd).Offset(0, 1) = ""
           GoTo endfd
     End If
    'Sheet11.Range(Tadd).Select
    Dim uiffsc As Variant
    uiffsc = UCase(Sheet11.Range(Tadd).value)
    Sheet11.Range(Tadd).value = uiffsc
    'Sheet11.Range(Tadd).Offset(0, 1) = UCase(Sheet11.Range(Tadd).Offset(0, 10).value)
End If

'If Tval = "" Then
'    Sheet11.Range(Tadd).Offset(0, 1) = ""
'End If

endfd:
Sheet11.Protect Password:=getmsgstate
Application.EnableEvents = True

End Function
'-------------------------------------------------------------------------------------
Function Validate80GGC_1() As Boolean
    Validate80GGC_1 = True
    
    MsgBox_80GGC = "Schedule 80GGC : " & Chr(10)
        setTableInfo80GGC
        setTableInfo80GGC1
        setTableInfo80GGC2
'        setTableInfo80GGC3
        setTableInfo80GGC4
        setTableInfo80GGC5
        setTableInfo80GGC6
        setTableInfo80GGC7
        
        
    end80GGC = WorksheetFunction.Max(0, end80GGC, end80GGC1, end80GGC2, end80GGC3, end80GGC4, end80GGC5, end80GGC6, end80GGC7)

        If end80GGC > 0 Then
            If Not ValidateDateofDonation_80GGC Then Validate80GGC_1 = False ' DATE
            'Chandru
            If Not ValidateMandatorySh80GGC Then Validate80GGC_1 = False
            
            If Not ValidateDonationAmt_80GGC Then Validate80GGC_1 = False 'DONATION IN CASH/OTHER MODE
'            If Not ValidateNatureofTransaction_80GGC Then Validate80GGC_1 = False 'NATURE OF TRANSACTION
            If Not ValidateChequeNumber_80GGC Then Validate80GGC_1 = False 'CHEQUE
            If Not ValidateBank_IFSC_Account Then Validate80GGC_1 = False 'IFSC,NAME OF BANK, ACCOUNT NUMBER
            If Not ValidateDonation_total_80GGC Then Validate80GGC_1 = False ' TOTAL DONATION
            If Not ValidateDonation_Eligible_80GGC Then Validate80GGC_1 = False 'ELIGIBLE DONATION

            If Not ValidateTotal_Donation_InCash_80GGC Then Validate80GGC_1 = False 'TOTAL DONATION IN CASH
            If Not ValidateTotal_Donation_OtherMode_80GGC Then Validate80GGC_1 = False 'TOTAL DONATION IN OTHER MODE
            If Not ValidateTotalof_Total_Donation_80GGC Then Validate80GGC_1 = False 'TOTAL - TOATL DONATION
            'Ayush_15/03/2026
            If Not ValidatePoliticalParty_Name_80GGC Then Validate80GGC_1 = False
            If Not ValidatePoliticalParty_PAN_80GGC Then Validate80GGC_1 = False
        End If
        
End Function
'--------------------------------------------------------------------------------------------------
Function ValidateDateofDonation_80GGC() As Boolean
    ValidateDateofDonation_80GGC = True
    'setTableInfo80GGC
    Dim rangecells As Range
    Set rangecells = Sheet11.Range("SCH_80GGC_DonationDate").Cells
    Dim i As Long
    ReDim DateofDonation_80GGC(end80GGC)

    subProcCaption = "Validating 80GGC"
    noOfProcessSub = end80GGC

    For i = 1 To end80GGC
        DateofDonation_80GGC(i) = rangecells.item(i).value

'        If Not chkCompulsory(DateofDonation_80GGC(i)) Then
'            MsgBox_80GGC = MsgBox_80GGC + "* please select "" date of Contribution "" at Sr. No." & i & "  in Schedule 80GGC" & Chr(13)
'            ValidateDateofDonation_80GGC = False
'            Exit Function
'        End If
'        UpdateProgressBar
    Next
End Function

'Chandru
Function ValidateMandatorySh80GGC() As Boolean
ValidateMandatorySh80GGC = True
Dim i As Long
Dim flag As Boolean

Dim rangecells As Range



flag = True
For i = 1 To Sheet11.Range("SCH_80GGC_DonationDate").Rows.count

    

    If Sheet11.Range("SCH_80GGC_DonationDate").item(i).value <> "" Or Sheet11.Range("SCH_80GGC_DonationAmtCash").item(i).value <> "" Or _
        Sheet11.Range("SCH_80GGC_DonationAmtOtherMode").item(i).value <> "" Or Sheet11.Range("SCH_80GGC_ChnImpsNeftRtgsNum").item(i).value <> "" _
        Or Sheet11.Range("SCH_80GGC_ChnImpsNeftRtgsNum").item(i).value <> "" Then
        
        
        
        If Sheet11.Range("SCH_80GGC_DonationDate").item(i).value = "" Then
            flag = False
'Ayush_Legacy17
'                MsgBox_80GGC = MsgBox_80GGC + "* please select "" date of Contribution "" " & Chr(13)
                MsgBox_80GGC = MsgBox_80GGC + "* please enter "" date of Contribution "" " & Chr(13)

        End If
        

        If flag = False Then
            ValidateMandatorySh80GGC = False
            Exit Function
        End If
        
    End If
Next i
End Function

'-------------------------------------------------------------------------------------------------
Function ValidateTotal_Donation_InCash_80GGC() As Boolean
 ValidateTotal_Donation_InCash_80GGC = True
 Total_DonationInCash_80GGC = Sheet11.Range("SCH_80GGC_TotalDonationAmtCash80GGC").value
 
 
 If Len(Total_DonationInCash_80GGC) > 14 Then
    MsgBox_80GGC = MsgBox_80GGC + "* Total amount of Contribution in Cash cannot be greater than 14 digits in Schedule 80GGC." & Chr(13)
    ValidateTotal_Donation_InCash_80GGC = False
    Exit Function
End If


End Function
'--------------------------------------------------------------
Function ValidateDonationAmt_80GGC() As Boolean
ValidateDonationAmt_80GGC = True
'    setTableInfo80GGC
    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Sheet11.Range("SCH_80GGC_DonationAmtCash").Cells
    Set rangecells1 = Sheet11.Range("SCH_80GGC_DonationAmtOtherMode").Cells
    Dim i As Long
    ReDim Donation_cash_80GGC(end80GGC)
    ReDim Donation_other_80GGC(end80GGC)
    
    For i = 1 To end80GGC
        Donation_cash_80GGC(i) = rangecells.item(i).value
        Donation_other_80GGC(i) = rangecells1.item(i).value
        

'    If rangecells.item(i).Locked = False And rangecells1.item(i).Locked = False Then
'       If rangecells1.item(i).Locked = False Then
       
           If rangecells.item(i).Locked = False Then
            If Not chkCompulsory(Donation_cash_80GGC(i)) Then
                MsgBox_80GGC = MsgBox_80GGC + "* Enter the amount of Contribution either in field ""Contribution in cash"" or ""Contribution in other mode at Sr. No " & i & "  in Schedule 80GGC" & Chr(13)
                ValidateDonationAmt_80GGC = False
                Exit Function
            End If
       End If
       
       If rangecells1.item(i).Locked = False Then
            If Not chkCompulsory(Donation_other_80GGC(i)) Then
                MsgBox_80GGC = MsgBox_80GGC + "* Enter the amount of Contribution either in field ""Contribution in cash"" or ""Contribution in other mode at Sr. No " & i & "  in Schedule 80GGC" & Chr(13)
                ValidateDonationAmt_80GGC = False
                Exit Function
            End If
       End If
'    End If
    Next
End Function

Function ValidateChequeNumber_80GGC_old() As Boolean
    ValidateChequeNumber_80GGC = True
'    setTableInfo80GGC
    Dim rangecells As Range
    Set rangecells = Sheet11.Range("SCH_80GGC_ChnImpsNeftRtgsNum").Cells
    Dim i As Long
    ReDim ChequeNumber_80GGC(end80GGC)
    
    For i = 1 To end80GGC
        ChequeNumber_80GGC(i) = rangecells.item(i).value
        
        If rangecells.item(i).Locked = False Then
        
            If Not chkCompulsory(ChequeNumber_80GGC(i)) Then
                MsgBox_80GGC = MsgBox_80GGC + "* please enter Cheque number/IMPS/NEFT/RTGS of Contribution transaction at Sr. No " & i & " in Schedule 80GGC" & Chr(13)
                ValidateChequeNumber_80GGC = False
                Exit Function
            End If
        End If
        UpdateProgressBar
    Next
End Function
'--------------------------------------------------------------------------------
Function ValidateBank_IFSC_Account() As Boolean
ValidateBank_IFSC_Account = True

Dim rangecells3 As Range


Set rangecells3 = Range("SCH_80GGC_IFSCCode").Cells


Dim i As Long
ReDim BankIFSC_80GGC(end80GGC)



For i = 1 To end80GGC
    
'    BankName_80GGC(i) = rangecells2.item(i).value
'    BankAccntnum_80GGC(i) = rangecells1.item(i).value
    BankIFSC_80GGC(i) = rangecells3.item(i).value
    
    
 If rangecells3.item(i).Locked = False Then
 
'    If BankName_80GGC(i) = "" Then
'        MsgBox_80GGC = MsgBox_80GGC + "* please enter your bank name from which Donation is made at Sr.No " & i & " in Schedule 80GGC" & Chr(13)
'        ValidateBank_IFSC_Account = False
'        'Exit Function
'    End If
'
'    If Len(BankName_80GGC(i)) > 125 Then
'        MsgBox_80GGC = MsgBox_80GGC + "* Bank Name at Sr.No " & i & " in Schedule 80GGC cannot exceed 125 characters" & Chr(13)
'        ValidateBank_IFSC_Account = False
'        'Exit Function
'    End If
    
    If isdropdownblank(BankIFSC_80GGC(i)) Then
        MsgBox_80GGC = MsgBox_80GGC + "* please enter ""your bank IFSC from which Contribution is made"" at Sr.No " & i & " in Schedule 80GGC." & Chr(13)
        ValidateBank_IFSC_Account = False
        'Exit Function

    End If
    

    If Len(BankIFSC_80GGC(i)) > 11 Then
        MsgBox_80GGC = MsgBox_80GGC + "* IFS Code at Sr.No " & i & " in Schedule 80GGC cannot exceed 11 characters" & Chr(13)
        ValidateBank_IFSC_Account = False
        'Exit Function
    End If
    
    If Not ValidateIFSCList(UCase(BankIFSC_80GGC(i))) Then
        MsgBox_80GGC = MsgBox_80GGC + "* Invalid IFS Code at Sr.No " & i & " in Schedule 80GGC.Refer to your bank for valid IFS Codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
        ValidateBank_IFSC_Account = False
        'Exit Function
    End If
    
    UpdateProgressBar
    
'    If Not chkCompulsory(BankAccntnum_80GGC(i)) Then
'        MsgBox_80GGC = MsgBox_80GGC + "* please enter your bank account number from which Donation is made at Sr.No " & i & " in Schedule 80GGC" & Chr(13)
'        ValidateBank_IFSC_Account = False
'        'Exit Function
'    End If
'
'
'
'    If Not ValidateBankAccountNumber_80GGC(BankAccntnum_80GGC(i), i) Then
'        ValidateBank_IFSC_Account = False
'       ' Exit Function
'    End If

  End If

Next
End Function
'-----------------------------------------------------------------------------------------------------------
Function ValidateBankAccountNumber_80GGC(BankAccountNumber As Variant, cc As Long) As Boolean
    ValidateBankAccountNumber_80GGC = True
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
           
    If Len(BankAccountNumber) > 0 Then
        If Not checkfieldspecialcharacter_Bank(BankAccountNumber) Then
            MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & cc & " is invalid in 80GGC, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber_80GGC = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & cc & " is invalid in 80GGC " & Chr(13)
            ValidateBankAccountNumber_80GGC = False
            Exit Function
        End If
    
    End If
    
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & cc & " is mandatory in 80GGC" & Chr(13)
        ValidateBankAccountNumber_80GGC = False
        Exit Function
    End If

    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & cc & " is mandatory in 80GGC" & Chr(13)
        ValidateBankAccountNumber_80GGC = False
        Exit Function
    End If

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0)) Then
        MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & cc & " is invalid in 80GGC  " & Chr(13)
        ValidateBankAccountNumber_80GGC = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & cc & " is invalid in 80GGC" & Chr(13)
        ValidateBankAccountNumber_80GGC = False
        Exit Function
    End If

    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & cc & " is invalid in 80GGC " & Chr(13)
        ValidateBankAccountNumber_80GGC = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & cc & " is invalid in 80GGC" & Chr(13)
        ValidateBankAccountNumber_80GGC = False
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
            MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & cc & " is invalid in 80GGC." & Chr(13)
            ValidateBankAccountNumber_80GGC = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "* Bank Account Number at Sr.No " & cc & " in 80GGC is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber_80GGC = False
        Exit Function
    End If
End Function
'---------------------------------------------------------------
Function ValidateDonation_total_80GGC() As Boolean
ValidateDonation_total_80GGC = True
'    setTableInfo80GGC
    Dim rangecells As Range
    Set rangecells = Sheet11.Range("SCH_80GGC_DonationAmt").Cells
    Dim i As Long
    ReDim Donation_total_80GGC(end80GGC)
    
    For i = 1 To end80GGC
        Donation_total_80GGC(i) = rangecells.item(i).value
        If Len(Donation_total_80GGC(i)) > 14 Then
            MsgBox_80GGC = MsgBox_80GGC + "* Total Contribution Amount at Sr. No " & i & " cannot exceed 14 digits in Schedule 80GGC." & Chr(13)
            ValidateDonation_total_80GGC = False
            Exit Function
        End If
    Next
End Function
'---------------------------------------------------------------------------------------
Function ValidateDonation_Eligible_80GGC() As Boolean
ValidateDonation_Eligible_80GGC = True
'    setTableInfo80GGC
    Dim rangecells As Range
    Set rangecells = Sheet11.Range("SCH_80GGC_EligibleDonationAmt").Cells
    Dim i As Long
    ReDim Donation_Eligible_80GGC(end80GGC)
    
    For i = 1 To end80GGC
        Donation_Eligible_80GGC(i) = rangecells.item(i).value
        If Len(Donation_Eligible_80GGC(i)) > 14 Then
            MsgBox_80GGC = MsgBox_80GGC + "* Eligible Contribution Amount at Sr. No " & i & " cannot exceed 14 digits in Schedule 80GGC." & Chr(13)
            ValidateDonation_Eligible_80GGC = False
            Exit Function
        End If
    Next
End Function
'-----------------------------------------------------------------------
Function ValidateTotal_Donation_OtherMode_80GGC() As Boolean
 ValidateTotal_Donation_OtherMode_80GGC = True
 Total_DonationInOThMode_80GGC = Sheet11.Range("SCH_80GGC_TotalDonationAmtOtherMode80GGC").value
 
 If Len(Total_DonationInOThMode_80GGC) > 14 Then
    MsgBox_80GGC = MsgBox_80GGC + "* Total amount of Contribution in Other Mode cannot be greater than 14 digits in Schedule 80GGC." & Chr(13)
    ValidateTotal_Donation_OtherMode_80GGC = False
    Exit Function
End If

End Function
'------------------------------------------------------------------------------------------
Function ValidateTotalof_Total_Donation_80GGC() As Boolean
 ValidateTotalof_Total_Donation_80GGC = True
 Total_TotalDonation_80GGC = Sheet11.Range("SCH_80GGC_TotalDonationsUs80GGC").value
 
 If Len(Total_TotalDonation_80GGC) > 14 Then
    MsgBox_80GGC = MsgBox_80GGC + "* Total amount Total Contribution cannot be greater than 14 digits in Schedule 80GGC." & Chr(13)
    ValidateTotalof_Total_Donation_80GGC = False
    Exit Function
End If

End Function
'------------------------------------------------------------------------------------------
Function ValidateTotal_Donation_Eligible_80GGC() As Boolean
 ValidateTotal_Donation_Eligible_80GGC = True
 Total_Donation_Eligible_80GGC = Sheet11.Range("SCH_80GGC_TotalEligibleDonationAmt80GGC").value
 
 If Len(Total_Donation_Eligible_80GGC) > 14 Then
    MsgBox_80GGC = MsgBox_80GGC + "* Total Eligible Contribution amount cannot be greater than 14 digits in Schedule 80GGC." & Chr(13)
    ValidateTotal_Donation_Eligible_80GGC = False
    Exit Function
End If

End Function


Public Sub fmsgboxoK(iintmsg As Variant)
With MessageBox
    .Height = 130
    .LMessagebox.Height = 60
    .CommandButton1.Top = 78
    .LMessagebox = "*" + iintmsg
    .LMessagebox.TextAlign = fmTextAlignLeft
    .Show
End With

End Sub

Function checkfieldspecialcharacter_Bank(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter_Bank = True
    Dim arr As Variant
    arr = Array("@", "*", "!", ".", "&", "#", "~", ";", "?", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter_Bank = False
            Exit Function
        End If
        Next
    Next
End Function

Function ValidateChequeNumber_80GGC() As Boolean
    ValidateChequeNumber_80GGC = True
'    setTableInfo80GGC
    Dim rangecells As Range
    Dim rangecells1 As Range
    
    Set rangecells = Sheet11.Range("SCH_80GGC_ChnImpsNeftRtgsNum").Cells
    Dim i As Long
    ReDim ChequeNumber_80GGC(end80GGC)
    
'    Set rangecells1 = Sheet11.Range("SCH_80GGC_TransactionRefNumUPI").Cells
'    ReDim NatureofTransaction_80GGC(end80GGC)
    
    For i = 1 To end80GGC
        ChequeNumber_80GGC(i) = rangecells.item(i).value
'        NatureofTransaction_80GGC(i) = rangecells1.item(i).value
        
        If rangecells.item(i).Locked = False Then  'Or rangecells1.item(i).Locked <> "(Select)"
          
'          If (NatureofTransaction_80GGC(i) <> "(Select)") Or (NatureofTransaction_80GGC(i) <> "") Then
          
            If (Donation_other_80GGC(i) > "0") Then
          
            If Not chkCompulsory(ChequeNumber_80GGC(i)) Then
                MsgBox_80GGC = MsgBox_80GGC + "* please enter ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of Contribution transaction at Sr. No " & i & " in Schedule 80GGC " & Chr(13)
'                 MsgBox_80GGC = MsgBox_80GGC + "* please enter transaction reference number of donation transaction at Sr. No " & i & " in Schedule 80GGC" & Chr(13)
                ValidateChequeNumber_80GGC = False
                Exit Function
            End If
            
            If Not checkfieldspecialcharacter_Bank(ChequeNumber_80GGC(i)) Then
                MsgBox_80GGC = MsgBox_80GGC + "*  Transaction Reference number/Cheque number/IMPS/NEFT/RTGS cannot contain special characters in Schedule 80GGC. Only "" / "" and "" - "" special characters are allowed. " & Chr(13)
'                 MsgBox_80GGC = MsgBox_80GGC + "* please enter transaction reference number of donation transaction at Sr. No " & i & " in Schedule 80GGC" & Chr(13)
                ValidateChequeNumber_80GGC = False
                Exit Function
            End If
            
          End If
            
'          If Len(ChequeNumber_80GGC(i)) > 0 Then
''             If (NatureofTransaction_80GGC(i) = "(Select)") Or (NatureofTransaction_80GGC(i) = "") Then
'
'
'             If (Donation_other_80GGC(i) > "0") Or (Donation_other_80GGC(i) = "") Then
'
'                 MsgBox_80GGC = MsgBox_80GGC + "* please select "" Nature of transaction "" of Contribution in other mode at Sr. No " & i & " in Schedule 80GGC" & Chr(13)
'
'                ValidateChequeNumber_80GGC = False
'                Exit Function
'
'             End If
'          End If
            
            
            
        End If
        UpdateProgressBar
    Next
End Function
Function ValidatePoliticalParty_PAN_80GGC() As Boolean
    ValidatePoliticalParty_PAN_80GGC = True

    Dim rangecells As Range
    Dim rangecells1 As Range
    
    Set rangecells = Sheet11.Range("SCH_80GGC_Political_PAN").Cells
    Dim i As Long
    ReDim Political_PAN_80GGC(end80GGC)
    

    
    For i = 1 To end80GGC
        Political_PAN_80GGC(i) = rangecells.item(i).value

        
        If rangecells.item(i).Locked = False Then  'Or rangecells1.item(i).Locked <> "(Select)"
          
     
            If (Donation_other_80GGC(i) > 0 Or Donation_cash_80GGC(i) > 0) Then
          
            If Not chkCompulsory(Political_PAN_80GGC(i)) Then
                MsgBox_80GGC = MsgBox_80GGC + "* ""Please enter PAN of Political Party in schedule 80GGC"" at Sr. No " & i & "" & Chr(13)
               ValidatePoliticalParty_PAN_80GGC = False
                Exit Function
            End If
            
           If ((UCase(Political_PAN_80GGC(i)) = UCase(Sheet1.Range("sheet1.PAN").value)) Or (UCase(Political_PAN_80GGC(i)) = UCase(Sheet5.Range("Ver.PAN").value))) Then
                 MsgBox_80GGC = MsgBox_80GGC + "* ""Political party PAN cannot be same as assesse PAN or verification PAN in schedule 80GGC "" at Sr. No " & i & "" & Chr(13)
               ValidatePoliticalParty_PAN_80GGC = False
                Exit Function
            End If
            
          End If
            

            
            
            
        End If
        UpdateProgressBar
    Next
End Function



Function ValidatePoliticalParty_Name_80GGC() As Boolean
ValidatePoliticalParty_Name_80GGC = True

    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Sheet11.Range("SCH_80GGC_Political_Name").Cells
    
    Dim i As Long
    ReDim Political_Name_80GGC(end80GGC)
    
    
    For i = 1 To end80GGC
        Political_Name_80GGC(i) = rangecells.item(i).value

        
        If rangecells.item(i).Locked = False Then  'Or rangecells1.item(i).Locked <> "(Select)"
          
     
            If (Donation_other_80GGC(i) > 0 Or Donation_cash_80GGC(i) > 0) Then
          
            If Not chkCompulsory(Political_Name_80GGC(i)) Then
                MsgBox_80GGC = MsgBox_80GGC + "* ""Please enter name of Political Party in schedule 80GGC"" at Sr. No " & i & "" & Chr(13)
               ValidatePoliticalParty_Name_80GGC = False
                Exit Function
            End If
            
            If Not CheckSpecialCharacter4New(Political_Name_80GGC(i)) Then
                MsgBox_80GGC = MsgBox_80GGC + "* Name of political party characters < >  are not allowed in schedule 80GGC" & Chr(13)
                ValidatePoliticalParty_Name_80GGC = False
                Exit Function
            End If
            
          End If
            
    End If
    Next
End Function

