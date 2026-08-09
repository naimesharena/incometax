Attribute VB_Name = "Sch80GGC"
Option Explicit

Public MsgBox_80GGC As String

Public DateofDonation_80GGC As Variant
Public NatureofTransaction_80GGC As Variant
Public Donation_cash_80GGC As Variant
Public Donation_other_80GGC As Variant
Public Donation_total_80GGC As Variant
Public Donation_Eligible_80GGC As Variant
Public ChequeNumber_80GGC As Variant
Public BankAccntnum_80GGC, BankIFSC_80GGC, BankName_80GGC As Variant
Public Total_DonationInCash_80GGC, Total_DonationInOThMode_80GGC, Total_TotalDonation_80GGC As Variant

Public rngName_80GGC As Variant
Public Total_Donation_Eligible_80GGC As Variant

Public end80GGC, end80GGC1, end80GGC2, end80GGC3, end80GGC4, end80GGC5, end80GGC6, end80GGC7, end80GGC8, end80GGC9 As Long
Sub ValidateSheet80GGC_Click()
 Dim vbMessgaeCaption As String
  vbMessgaeCaption = "ITR 3: AY: 2026-27"
  Validate80GGC
'MsgBox "Sheet 80GGC is OK", vbOKOnly, vbMessgaeCaption
fmsgboxStatus "Sheet 80GGC is OK"
End Sub
Sub Validate80GGC()
    If Not Validate80GGC_1 Then
        Sheet58.Activate
        fmsgbox (MsgBox_80GGC)
        CloseMsg
    End If
End Sub

Sub setTableInfo80GGC() '80GGC_C1
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("DateofDonation_80GGC").Cells
    mIntCells = Sheet58.Range("DateofDonation_80GGC").count
    For mIntCtr = 1 To mIntCells
         If Not rangecells.item(mIntCtr).value = "" Then
          ccount = ccount + 1
        End If
    Next
    end80GGC = ccount
'    rngName_80GGC = "DateofDonation_80GGC;Donationincash_80GGC;Donationinothermode_80GGC;TotalDonation_80GGC;EligibleAmountofDonation_80GGC;NatureofTransaction_80GGC;Chequeno_80GGC;IFSC_80GGC;"
    rngName_80GGC = "DateofDonation_80GGC;Donationincash_80GGC;Donationinothermode_80GGC;TotalDonation_80GGC;EligibleAmountofDonation_80GGC;NatureofTransaction_80GGC;NameOfParty_80GGC;PANofParty_80GGC;Chequeno_80GGC;IFSC_80GGC;"

End Sub
Sub setTableInfo80GGC1() '80GGC_C2
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("Donationincash_80GGC").Cells
    mIntCells = Sheet58.Range("Donationincash_80GGC").count
    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
        If Not rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).Locked = True Then
        
        ccount = ccount + 1
        End If
    Next
    end80GGC1 = ccount
End Sub

Sub setTableInfo80GGC2() '80GGC_C3
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("Donationinothermode_80GGC").Cells
    mIntCells = Sheet58.Range("Donationinothermode_80GGC").count
    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
        If Not rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).Locked = True Then
        ccount = ccount + 1
        End If
    Next
    end80GGC2 = ccount
End Sub

Sub setTableInfo80GGC3() '80GGC_C4
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("NatureofTransaction_80GGC").Cells
    mIntCells = Sheet58.Range("NatureofTransaction_80GGC").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end80GGC3 = ccount
End Sub

Sub setTableInfo80GGC4() '80GGC_C5
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("Chequeno_80GGC").Cells
    mIntCells = Sheet58.Range("Chequeno_80GGC").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end80GGC4 = ccount
End Sub

Sub setTableInfo80GGC5() '80GGC_C6
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("IFSC_80GGC").Cells
    mIntCells = Sheet58.Range("IFSC_80GGC").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGC5 = ccount
End Sub

Sub setTableInfo80GGC8()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("NameOfParty_80GGC").Cells
    mIntCells = Sheet58.Range("NameOfParty_80GGC").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGC8 = ccount
End Sub

Sub setTableInfo80GGC9()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet58.Range("PANofParty_80GGC").Cells
    mIntCells = Sheet58.Range("PANofParty_80GGC").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGC9 = ccount
End Sub

'Sub setTableInfo80GGC6() '80GGC_C7
'    Dim rangecells As Range
'    Dim mIntCells  As Long
'    Dim mIntCtr  As Long
'    Dim ccount  As Long
'    ccount = 0
'    Set rangecells = Sheet59.Range("NameofDonor_80GGC").Cells
'    mIntCells = Sheet59.Range("NameofDonor_80GGC").count
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).Value = "" Then
'        ccount = ccount + 1
'        End If
'    Next
'    end80GGC6 = ccount
'End Sub

'Sub setTableInfo80GGC7() '80GGC_C8
'    Dim rangecells As Range
'    Dim mIntCells  As Long
'    Dim mIntCtr  As Long
'    Dim ccount  As Long
'    ccount = 0
'    Set rangecells = Sheet59.Range("AccountofDonor_80GGC").Cells
'    mIntCells = Sheet59.Range("AccountofDonor_80GGC").count
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).Value = "" Then
'        ccount = ccount + 1
'        End If
'    Next
'    end80GGC7 = ccount
'End Sub



Sub AddRows80GGC() '80CCG
    Dim vRows  As Long
    Sheets("80GGC").Activate
'    EfilingCommon.DefinedgridNameRange = "DateofDonation_80GGC||Donationincash_80GGC||Donationinothermode_80GGC||TotalDonation_80GGC||EligibleAmountofDonation_80GGC||Chequeno_80GGC||IFSC_80GGC||"
     EfilingCommon.DefinedgridNameRange = "DateofDonation_80GGC||Donationincash_80GGC||Donationinothermode_80GGC||TotalDonation_80GGC||EligibleAmountofDonation_80GGC||NameOfParty_80GGC||PANofParty_80GGC||Chequeno_80GGC||IFSC_80GGC||"
    ActiveCellRange = EfilingCommon.searchLastRow("DateofDonation_80GGC")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula_80GGC
End Sub
Function CheckIFSC_80GGC(Tval As Variant, Tadd As Variant) '80GGC_C5
On Error Resume Next
Application.EnableEvents = False
 sPassword = EfilingCommon.getmsgstate
   Sheet58.Unprotect Password:=sPassword


If Tval <> "" Then
    
     If Not CheckIFSC(UCase(Tval)) Then
        fmsgbox "Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets in Bank Details in Schedule 80GGC Sheet."
           Sheet58.Range(Tadd) = ""
           'Sheet59.Range(Tadd).Offset(0, 1) = ""
           GoTo endfd
     End If
    
    If Not ValidateIFSCList(UCase(Tval)) Then
           fmsgbox "Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets in Bank Details in Schedule 80GGC Sheet."
           Sheet58.Range(Tadd) = ""
           'Sheet59.Range(Tadd).Offset(0, 1) = ""
           GoTo endfd
     End If
    'Sheet59.Range(Tadd).Select
   
    'Sheet59.Range(Tadd).Offset(0, 1) = UCase(Sheet59.Range(Tadd).Offset(0, 8).Value)
End If

'If Tval = "" Then
'    Sheet59.Range(Tadd).Offset(0, 1) = ""
'End If

endfd:
Sheet58.Protect Password:=sPassword
Application.EnableEvents = True

End Function

Function Validate80GGC_1() As Boolean
    Validate80GGC_1 = True
    
    MsgBox_80GGC = "Schedule 80GGC : " & Chr(10)
        setTableInfo80GGC
        setTableInfo80GGC1
        setTableInfo80GGC2
       ' setTableInfo80GGC3
        
        setTableInfo80GGC8
        setTableInfo80GGC9
       
        setTableInfo80GGC4
        setTableInfo80GGC5
        'setTableInfo80GGC6
        'setTableInfo80GGC7
        
        
    end80GGC = WorksheetFunction.Max(0, end80GGC, end80GGC1, end80GGC2, end80GGC4, end80GGC5, end80GGC8, end80GGC9)

        If end80GGC > 0 Then
            If Not ValidateDateofDonation_80GGC Then Validate80GGC_1 = False ' DATE
            If Not ValidateDonationAmt_80GGC Then Validate80GGC_1 = False 'DONATION IN CASH/OTHER MODE
            'If Not ValidateNatureofTransaction_80GGC Then Validate80GGC_1 = False 'NATURE OF TRANSACTION
            If Not ValidateChequeNumber_80GGC Then Validate80GGC_1 = False 'TRANSACTION REF & NATURE OF TRANSACTION




            'Added by Aavula Naresh AY 2026-27
            If Not ValidateNamePP_80GGC Then Validate80GGC_1 = False 'Name of Political Party
            If Not ValidatePanPP_80GGC Then Validate80GGC_1 = False 'Pan of Political Party
            ''''''''''''
            
            'If Not ValidateBank_IFSC_Account Then Validate80GGC_1 = False 'IFSC
            If Not ValidateDonation_total_80GGC Then Validate80GGC_1 = False ' TOTAL DONATION
            If Not ValidateDonation_Eligible_80GGC Then Validate80GGC_1 = False 'ELIGIBLE DONATION

            If Not ValidateTotal_Donation_InCash_80GGC Then Validate80GGC_1 = False 'TOTAL DONATION IN CASH
            If Not ValidateTotal_Donation_OtherMode_80GGC Then Validate80GGC_1 = False 'TOTAL DONATION IN OTHER MODE
            If Not ValidateTotalof_Total_Donation_80GGC Then Validate80GGC_1 = False 'TOTAL - TOATL DONATION
            

            
        End If
        
End Function

Function ValidateDateofDonation_80GGC() As Boolean
    ValidateDateofDonation_80GGC = True
'    setTableInfo80GGC
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("DateofDonation_80GGC").Cells
    Dim i As Long
    ReDim DateofDonation_80GGC(end80GGC)
    
    subProcCaption = "Validating 80GGC"
    noOfProcessSub = end80GGC
    
    For i = 1 To end80GGC
        DateofDonation_80GGC(i) = rangecells.item(i).value
        
        If Not chkCompulsory(DateofDonation_80GGC(i)) Then
        
        'Added by Aavula
        
'            MsgBox_80GGC = MsgBox_80GGC + "* Please enter date of contribution at Sr. No." & i & "" & Chr(13)

            MsgBox_80GGC = MsgBox_80GGC + "* please select ""date of contribution"" at Sr. No." & i & "" & Chr(13)
            ValidateDateofDonation_80GGC = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function
Function ValidateTotal_Donation_InCash_80GGC() As Boolean
 ValidateTotal_Donation_InCash_80GGC = True
 Total_DonationInCash_80GGC = Sheet58.Range("Total_DonationInCash_80GGC").value
 
 
 If Len(Total_DonationInCash_80GGC) > 14 Then
    MsgBox_80GGC = MsgBox_80GGC + "* Total amount of Donation in Cash cannot be greater than 14 digits in Schedule 80GGC." & Chr(13)

    ValidateTotal_Donation_InCash_80GGC = False
    Exit Function
End If


End Function
Function ValidateDonationAmt_80GGC() As Boolean
ValidateDonationAmt_80GGC = True
'    setTableInfo80GGC
    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Sheet58.Range("Donationincash_80GGC").Cells
    Set rangecells1 = Sheet58.Range("Donationinothermode_80GGC").Cells
    Dim i As Long
    ReDim Donation_cash_80GGC(end80GGC)
    ReDim Donation_other_80GGC(end80GGC)
    
    For i = 1 To end80GGC
        Donation_cash_80GGC(i) = rangecells.item(i).value
        Donation_other_80GGC(i) = rangecells1.item(i).value
        
       If rangecells.item(i).Locked = False Then
            If Not chkCompulsory(Donation_cash_80GGC(i)) Then
                'MsgBox_80GGC = MsgBox_80GGC + "* Enter the amount of donation either in field Donation in cash or Donation in other mode at Sr. No " & i & " is Mandatory in Schedule 80GGC." &
                
                MsgBox_80GGC = MsgBox_80GGC + "*Enter the amount of contribution either in field ""contribution in cash"" or ""contribution in other mode at Sr. No " & i & "  in schedule 80GGC"" " & Chr(13)

                
                ValidateDonationAmt_80GGC = False
                Exit Function
            End If
       End If
       
       If rangecells1.item(i).Locked = False Then
            If Not chkCompulsory(Donation_other_80GGC(i)) Then
                'MsgBox_80GGC = MsgBox_80GGC + "* Enter the amount of donation either in field Donation in cash or Donation in other mode at Sr. No " & i & " is Mandatory in Schedule 80GGC." &
                
                MsgBox_80GGC = MsgBox_80GGC + "*Enter the amount of contribution either in field ""contribution in cash"" or ""contribution  in other mode at Sr. No " & i & " in schedule 80GGC"" " & Chr(13)

                
                ValidateDonationAmt_80GGC = False
                Exit Function
            End If
       End If
       
       

    Next
End Function
Function ValidateNatureofTransaction_80GGC() As Boolean
    ValidateNatureofTransaction_80GGC = True
'    setTableInfo80GGC
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("NatureofTransaction_80GGC").Cells
    Dim i As Long
    ReDim NatureofTransaction_80GGC(end80GGC)
    
    For i = 1 To end80GGC
        NatureofTransaction_80GGC(i) = rangecells.item(i).value
        
        If rangecells.item(i).Locked = False Then
        
            If (NatureofTransaction_80GGC(i) = "(Select)") Or (NatureofTransaction_80GGC(i) = "") Then
                'MsgBox_80GGC = MsgBox_80GGC + "* please enter transaction reference number of donation transaction at Sr. No " & i & " in Schedule 80GGC" &
                
                MsgBox_80GGC = MsgBox_80GGC + "* please enter ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of contribution transaction at Sr. No " & i & " in Schedule 80GGC" & Chr(13) & Chr(13)


                ValidateNatureofTransaction_80GGC = False
                Exit Function
            End If
        End If
        UpdateProgressBar
    Next
End Function
Function ValidateChequeNumber_80GGC() As Boolean
    ValidateChequeNumber_80GGC = True
'    setTableInfo80GGC
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    
    Set rangecells2 = Sheet58.Range("Donationinothermode_80GGC").Cells
    ReDim Donation_other_80GGC(end80GGC)
    
    Set rangecells = Sheet58.Range("Chequeno_80GGC").Cells
    Dim i As Long
    ReDim ChequeNumber_80GGC(end80GGC)
    
    'Set rangecells1 = Sheet59.Range("NatureofTransaction_80GGC").Cells
    'ReDim NatureofTransaction_80GGC(end80GGC)
    
    
    Dim rangecells3 As Range
    Set rangecells3 = Range("IFSC_80GGC").Cells
    ReDim BankIFSC_80GGC(end80GGC)
    
    
    For i = 1 To end80GGC
        ChequeNumber_80GGC(i) = rangecells.item(i).value
        'NatureofTransaction_80GGC(i) = rangecells1.item(i).value
        BankIFSC_80GGC(i) = rangecells3.item(i).value
        Donation_other_80GGC(i) = rangecells2.item(i).value
        
        
        
    If rangecells2.item(i).Locked = False Then
           
         
'         If Donation_other_80GGC(i) > 0 Then
'           If (NatureofTransaction_80GGC(i) = "(Select)") Or (NatureofTransaction_80GGC(i) = "") Then
'                MsgBox_80GGC = MsgBox_80GGC + "* please select Nature of  transaction of donation transaction at Sr. No " & i & " in Schedule 80GGC" &
'                ValidateChequeNumber_80GGC = False
'                Exit Function
'             End If
'
'
'          End If
        
          
          'If (NatureofTransaction_80GGC(i) <> "(Select)") And (NatureofTransaction_80GGC(i) <> "") Then
           If Len(Donation_other_80GGC(i)) > 0 And Donation_other_80GGC(i) <> 0 Then
            
            If Not chkCompulsory(ChequeNumber_80GGC(i)) Then
                'MsgBox_80GGC = MsgBox_80GGC + "* please enter Transaction Reference number/Cheque number/IMPS/NEFT/RTGS of donation transaction at Sr. No " & i & " in Schedule 80GGC" &
                'MsgBox_80GGC = MsgBox_80GGC + "* please enter ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of contribution transaction at Sr. No " & i & " in Schedule 80GGC"
                
                MsgBox_80GGC = MsgBox_80GGC + "*please enter ""transaction reference number"" of Contribution transaction at Sr. No " & i & " in Schedule 80GGC" & Chr(13)

                ValidateChequeNumber_80GGC = False
                Exit Function
            End If
         ' End If
         
'          If Len(ChequeNumber_80GGC(i)) > 0 Or Len(BankIFSC_80GGC(i)) > 1 Then
'             If (NatureofTransaction_80GGC(i) = "(Select)") Or (NatureofTransaction_80GGC(i) = "") Then
'                MsgBox_80GGC = MsgBox_80GGC + "* please select Nature of  transaction of donation transaction at Sr. No " & i & " in Schedule 80GGC" &
'                ValidateChequeNumber_80GGC = False
'                Exit Function
'             End If
'          End If


          'If (NatureofTransaction_80GGC(i) <> "(Select)") And (NatureofTransaction_80GGC(i) <> "") Then
            If isdropdownblank(BankIFSC_80GGC(i)) Then
                'MsgBox_80GGC = MsgBox_80GGC + "* please enter your bank IFSC from which Donation is made at Sr.No " & i & " in Schedule 80GGC." &
                MsgBox_80GGC = MsgBox_80GGC + "* please enter ""your bank IFSC from which contribution is made"" at Sr.No " & i & " in Schedule 80GGC."
                
                
                ValidateChequeNumber_80GGC = False
                Exit Function
            End If
          End If
            
            If Len(ChequeNumber_80GGC(i)) > 50 Then
                MsgBox_80GGC = MsgBox_80GGC + "* Transaction Reference number/Cheque number/IMPS/NEFT/RTGS of Contribution transaction at Sr.No " & i & " in Schedule 80GGC cannot exceed 50 characters"
                ValidateChequeNumber_80GGC = False
                Exit Function
            End If
            
            
            If Len(BankIFSC_80GGC(i)) > 11 Then
                MsgBox_80GGC = MsgBox_80GGC + "* IFS Code at Sr.No " & i & " in Schedule 80GGC cannot exceed 11 characters"
                ValidateChequeNumber_80GGC = False
                Exit Function
            End If
            
            If Not EfilingCommon.checkListIFSC(UCase(BankIFSC_80GGC(i))) Then
                MsgBox_80GGC = MsgBox_80GGC + "* Invalid IFS Code at Sr.No " & i & " in Schedule 80GGC.Refer to your bank for valid IFS Codes." & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)"
                ValidateChequeNumber_80GGC = False
                Exit Function
            End If
        
     End If
        UpdateProgressBar
    Next
End Function

Function ValidateNamePP_80GGC() As Boolean
    ValidateNamePP_80GGC = True
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("NameOfParty_80GGC").Cells
    Dim i As Long
    ReDim NamePP_80GGC(end80GGC)
    
    For i = 1 To end80GGC
        NamePP_80GGC(i) = rangecells.item(i).value
        
        If Not chkMandatory(NamePP_80GGC(i)) Then
            MsgBox_80GGC = MsgBox_80GGC + "* ""Please enter name of Political Party in schedule 80GGC""" & Chr(13)
            ValidateNamePP_80GGC = False
            Exit Function
        End If
        
        If Len(NamePP_80GGC(i)) > 125 Then
            MsgBox_80GGC = MsgBox_80GGC + "* Name of the Share/Unit at Sr. No." & i & " should not exceed 125 characters.in Schedule 112A" & Chr(13)
            ValidateNamePP_80GGC = False
            Exit Function
        End If
    Next
End Function

Function ValidatePanPP_80GGC() As Boolean
    Dim rangecells As Range
    Dim tempPan As String
    Dim i As Long
    
    ValidatePanPP_80GGC = True
    
    Set rangecells = Sheet58.Range("PANofParty_80GGC").Cells
    ReDim PanPP_80GGC(end80GGC)
    For i = 1 To end80GGC
        PanPP_80GGC(i) = rangecells.item(i).value
        
        tempPan = PanPP_80GGC(i)
        If Not chkCompulsory(PanPP_80GGC(i)) Then
            MsgBox_80GGC = MsgBox_80GGC + "* ""Please enter PAN of Political Party in schedule 80GGC""" & Chr(13)
            ValidatePanPP_80GGC = False
            Exit Function
        End If
        
        If Not CheckPAN(tempPan) Then
           
        'changed by Chetan C M for AY 2026-27 as per DE V0.2
        'start--
           'MsgBox_80GGC = MsgBox_80GGC + "* ""Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet  in schedule 80GGC." & Chr(13)
'           MsgBox_80GGC = MsgBox_80GGC + "* ""Invalid PAN in schedule 80GGC." & Chr(13)
  'changed by Naresh for AY 2026-27 as per DE V0.3
  MsgBox_80GGC = MsgBox_80GGC + "* ""PAN of the political party is invalid in schedule 80GGC." & Chr(13)
  
        '--end
           
            ValidatePanPP_80GGC = False
            Exit Function
        End If
        
'SIT-117302 by sadineni on 11/04/2026
'****************************************************************************************************
        If ((UCase(PanPP_80GGC(i)) = UCase(Sheet1.Range("sheet1.PAN").value)) Or (UCase(PanPP_80GGC(i)) = UCase(Sheet52.Range("sheet9.PAN").value))) Then
                 MsgBox_80GGC = MsgBox_80GGC + "* ""Political party PAN cannot be same as assesse PAN or verification PAN in schedule 80GGC "" at Sr. No " & i & "" & Chr(13)
              ValidatePanPP_80GGC = False
                Exit Function
        End If
'*****************************************************************************************************
    Next
End Function

Function ValidateBank_IFSC_Account() As Boolean
ValidateBank_IFSC_Account = True
'setTableInfo80GGC
'Dim rangecells As Range
'Dim rangecells1 As Range
'Dim rangecells2 As Range
Dim rangecells3 As Range


'Set rangecells1 = Range("AccountofDonor_80GGC").Cells
'Set rangecells2 = Range("NameofDonor_80GGC").Cells
Set rangecells3 = Range("IFSC_80GGC").Cells


Dim i As Long
ReDim BankIFSC_80GGC(end80GGC)
'ReDim BankName_80GGC(end80GGC)
'ReDim BankAccntnum_80GGC(end80GGC)



For i = 1 To end80GGC
    
    'BankName_80GGC(i) = rangecells2.item(i).Value
    'BankAccntnum_80GGC(i) = rangecells1.item(i).Value
    BankIFSC_80GGC(i) = rangecells3.item(i).value
    
    
 If rangecells3.item(i).Locked = False Then
 
'    If BankName_80GGC(i) = "" Then
'        MsgBox_80GGC = MsgBox_80GGC + "* please enter your bank name from which Donation is made at Sr.No " & i & " in Schedule 80GGC" &
'        ValidateBank_IFSC_Account = False
'        'Exit Function
'    End If
'
'    If Len(BankName_80GGC(i)) > 125 Then
'        MsgBox_80GGC = MsgBox_80GGC + "* Bank Name at Sr.No " & i & " in Schedule 80GGC cannot exceed 125 characters" &
'        ValidateBank_IFSC_Account = False
'        'Exit Function
'    End If
'
    If isdropdownblank(BankIFSC_80GGC(i)) Then
        MsgBox_80GGC = MsgBox_80GGC + "* please enter ""your bank IFSC from which contribution is made"" at Sr.No " & i & " in Schedule 80GGC."
        ValidateBank_IFSC_Account = False
        'Exit Function

    End If
    

    If Len(BankIFSC_80GGC(i)) > 11 Then
        MsgBox_80GGC = MsgBox_80GGC + "* IFS Code at Sr.No " & i & " in Schedule 80GGC cannot exceed 11 characters"
        ValidateBank_IFSC_Account = False
        'Exit Function
    End If
    
    If Not ValidateIFSCList(UCase(BankIFSC_80GGC(i))) Then
        MsgBox_80GGC = MsgBox_80GGC + "* Invalid IFS Code at Sr.No " & i & " in Schedule 80GGC.Refer to your bank for valid IFS Codes." & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)"
        ValidateBank_IFSC_Account = False
        'Exit Function
    End If
    
    UpdateProgressBar
    
'    If Not chkCompulsory(BankAccntnum_80GGC(i)) Then
'        MsgBox_80GGC = MsgBox_80GGC + "* please enter your bank account number from which Donation is made at Sr.No " & i & " in Schedule 80GGC" &
'        ValidateBank_IFSC_Account = False
'        'Exit Function
'    End If
    
    
    
'    If Not ValidateBankAccountNumber_80GGC(BankAccntnum_80GGC(i), i) Then
'        ValidateBank_IFSC_Account = False
'       ' Exit Function
'    End If

  End If

Next
End Function
Function ValidateBankAccountNumber_80GGC(BankAccountNumber As Variant, CC As Long) As Boolean
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
            MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & CC & " is invalid in 80GGC, Only "" / "" and "" - "" special characters are allowed."
            ValidateBankAccountNumber_80GGC = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & CC & " is invalid in 80GGC "
            ValidateBankAccountNumber_80GGC = False
            Exit Function
        End If
    
    End If
    
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & CC & " is mandatory in 80GGC"
        ValidateBankAccountNumber_80GGC = False
        Exit Function
    End If

    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & CC & " is mandatory in 80GGC"
        ValidateBankAccountNumber_80GGC = False
        Exit Function
    End If

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0)) Then
        MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & CC & " is invalid in 80GGC  "
        ValidateBankAccountNumber_80GGC = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & CC & " is invalid in 80GGC"
        ValidateBankAccountNumber_80GGC = False
        Exit Function
    End If

    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & CC & " is invalid in 80GGC "
        ValidateBankAccountNumber_80GGC = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & CC & " is invalid in 80GGC"
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
            MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & CC & " is invalid in 80GGC."
            ValidateBankAccountNumber_80GGC = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "* Bank Account Number at Sr.No " & CC & " in 80GGC is invalid. Account Number should have a minimum of 1 numeric digit (1-9)"
        ValidateBankAccountNumber_80GGC = False
        Exit Function
    End If
End Function
Function ValidateDonation_total_80GGC() As Boolean
ValidateDonation_total_80GGC = True
'    setTableInfo80GGC
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("TotalDonation_80GGC").Cells
    Dim i As Long
    ReDim Donation_total_80GGC(end80GGC)
    
    For i = 1 To end80GGC
        Donation_total_80GGC(i) = rangecells.item(i).value
        If Len(Donation_total_80GGC(i)) > 14 Then
            MsgBox_80GGC = MsgBox_80GGC + "* Total contribution Amount at Sr. No " & i & " cannot exceed 14 digits in Schedule 80GGC."
            ValidateDonation_total_80GGC = False
            Exit Function
        End If
    Next
End Function

Function ValidateDonation_Eligible_80GGC() As Boolean
ValidateDonation_Eligible_80GGC = True
'    setTableInfo80GGC
    Dim rangecells As Range
    Set rangecells = Sheet58.Range("EligibleAmountofDonation_80GGC").Cells
    Dim i As Long
    ReDim Donation_Eligible_80GGC(end80GGC)
    
    For i = 1 To end80GGC
        Donation_Eligible_80GGC(i) = rangecells.item(i).value
        If Len(Donation_Eligible_80GGC(i)) > 14 Then
            MsgBox_80GGC = MsgBox_80GGC + "* Eligible contribution Amount at Sr. No " & i & " cannot exceed 14 digits in Schedule 80GGC."
            ValidateDonation_Eligible_80GGC = False
            Exit Function
        End If
    Next
End Function
Function ValidateTotal_Donation_OtherMode_80GGC() As Boolean
 ValidateTotal_Donation_OtherMode_80GGC = True
 Total_DonationInOThMode_80GGC = Sheet58.Range("Total_DonationInOtherMode_80GGC").value
 
 If Len(Total_DonationInOThMode_80GGC) > 14 Then
    MsgBox_80GGC = MsgBox_80GGC + "* Total amount of contribution in Other Mode cannot be greater than 14 digits in Schedule 80GGC."
    ValidateTotal_Donation_OtherMode_80GGC = False
    Exit Function
End If

End Function
Function ValidateTotalof_Total_Donation_80GGC() As Boolean
 ValidateTotalof_Total_Donation_80GGC = True
 Total_TotalDonation_80GGC = Sheet58.Range("Total_Donation_80GGC").value
 
 If Len(Total_TotalDonation_80GGC) > 14 Then
    MsgBox_80GGC = MsgBox_80GGC + "* Total amount Total contribution cannot be greater than 14 digits in Schedule 80GGC."
    ValidateTotalof_Total_Donation_80GGC = False
    Exit Function
End If

End Function
Function ValidateTotal_Donation_Eligible_80GGC() As Boolean
 ValidateTotal_Donation_Eligible_80GGC = True
 Total_Donation_Eligible_80GGC = Sheet58.Range("Total_Donation_Eligible_80GGC").value
 
 If Len(Total_Donation_Eligible_80GGC) > 14 Then
    MsgBox_80GGC = MsgBox_80GGC + "* Total Eligible contribution amount cannot be greater than 14 digits in Schedule 80GGC."
    ValidateTotal_Donation_Eligible_80GGC = False
    Exit Function
End If

End Function

 




