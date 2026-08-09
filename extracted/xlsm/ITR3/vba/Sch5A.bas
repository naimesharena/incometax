Attribute VB_Name = "Sch5A"
Option Explicit

Dim Msg5ASheet, Msg5ASheetHP, Msg5ASheetBP, Msg5ASheetCG, Msg5ASheetOS As String

Public NameOfSpouse_S5A, PANOfSpouse_S5A, AadhaarOfSpouse_S5A As String

Public IncmRecvdHP As Variant
Public IncmHeadHP As Variant
Public TDSAmtDedHP As Variant
Public TDSApprndHP As Variant

Public IncmRecvdBP As Variant
Public IncmHeadBP As Variant
Public TDSAmtDedBP As Variant
Public TDSApprndBP As Variant

Public IncmRecvdCG As Variant
Public IncmHeadCG As Variant
Public TDSAmtDedCG As Variant
Public TDSApprndCG As Variant

Public IncmRecvdOS As Variant
Public IncmHeadOS As Variant
Public TDSAmtDedOS As Variant
Public TDSApprndOS As Variant

Public IncmRecvdTot As Variant 'added new
Public IncmHeadTot As Variant   'added new
Public TDSAmtDedTot As Variant  'aded new
Public TDSApprndTot As Variant 'added new


Sub Cmd_Validate_5A_Click()
 ValidateSchedule5A
 fmsgboxStatus "Sheet 5A is ok" ', vbOKOnly, "ITR 3"
End Sub



Sub ValidateSchedule5A()
'On Error Resume Next
If Not ValidatesheetSch5A Then
        Sheet29.Activate
    If Msg5ASheet <> "" Then
    fmsgboxsmall Msg5ASheet ', vbOKOnly, "Error(s)!"
    CloseMsg
    End If
    
    If Msg5ASheetHP <> "" Then
    fmsgboxsmall Msg5ASheetHP ', vbOKOnly, "Error(s)!"
    CloseMsg
    End If
    
    If Msg5ASheetBP <> "" Then
    fmsgboxsmall Msg5ASheetBP ', vbOKOnly, "Error(s)!"
    CloseMsg
    End If
    
    If Msg5ASheetCG <> "" Then
    fmsgboxsmall Msg5ASheetCG ', vbOKOnly, "Error(s)!"
    CloseMsg
    End If
    
    If Msg5ASheetOS <> "" Then
    fmsgboxsmall Msg5ASheetOS ', vbOKOnly, "Error(s)!"
    CloseMsg
    End If
    
    
End If
    
End Sub

Function ValidatesheetSch5A() As Boolean

Msg5ASheet = ""



ValidatesheetSch5A = True

If Mid(Sheet1.Range("sheet1.PortugeseCC5A").value, 1, 1) = "Y" Then
        If Sheet29.Range("BooksSpouse44ABFlg").value = "" Or Sheet29.Range("BooksSpouse44ABFlg").value = "(Select)" Or IsEmpty(Sheet29.Range("BooksSpouse44ABFlg").value) Then
             Msg5ASheet = Msg5ASheet + "* Please select the option in whether books of accounts of spouse is audited u/s 44AB?  or Whether your  spouse  is a  partner of a firm whose accounts are required to be audited u/s 44AB under this Act" & Chr(13)
             ValidatesheetSch5A = False
        End If
        If Sheet29.Range("BooksSpouse92EFlg").value = "" Or Sheet29.Range("BooksSpouse92EFlg").value = "(Select)" Or IsEmpty(Sheet29.Range("BooksSpouse92EFlg").value) Then
'             Msg5ASheet = Msg5ASheet + "* Please select the option in whether books of accounts of spouse is audited u/s 92E? or Whether your spouse is a  partner of a firm whose accounts are required to be audited u/s 92E under this Act " & Chr(13)
           
           'Added by Aavula Naresh for A.Y 2025_26 Dev
              Msg5ASheet = Msg5ASheet + "* Please select the option ""Whether your spouse is liable for audit u/s 92E? or Whether your spouse is a partner of a firm  which is liable for audit u/s 92E under this Act?"" " & Chr(13)
            ''''
             ValidatesheetSch5A = False
        End If
        If Not ValidateNameOfSpouse_S5A() Then ValidatesheetSch5A = False
        If Not ValidatePANOfSpouse_S5A() Then ValidatesheetSch5A = False
        If Not ValidateHP_S5A() Then ValidatesheetSch5A = False
        If Not ValidateBP_S5A() Then ValidatesheetSch5A = False
        If Not ValidateCG_S5A() Then ValidatesheetSch5A = False
        If Not ValidateOS_S5A() Then ValidatesheetSch5A = False
End If
End Function

Function ValidateNameOfSpouse_S5A() As Boolean

ValidateNameOfSpouse_S5A = True
 NameOfSpouse_S5A = Sheet29.Range("S5A_NameOfSpouse").value
 
 If NameOfSpouse_S5A = "" Or IsEmpty(NameOfSpouse_S5A) Then
    Msg5ASheet = Msg5ASheet + "* Please enter  Name  of the spouse in schedule 5A" & Chr(13)
    ValidateNameOfSpouse_S5A = False
End If

End Function

Function ValidatePANOfSpouse_S5A()

ValidatePANOfSpouse_S5A = True
'Sheet29.Range("S5A_PANOfSpouse").value = UCase(Sheet29.Range("S5A_PANOfSpouse").value)
 PANOfSpouse_S5A = Sheet29.Range("S5A_PANOfSpouse").value
 AadhaarOfSpouse_S5A = Sheet29.Range("S5A_AadhaarOfSpouse").value
 
            If Not chkCompulsory(PANOfSpouse_S5A) Then
                Msg5ASheet = Msg5ASheet + Chr(13) & "Please enter PAN  of the spouse in schedule 5A" & Chr(13)
                ValidatePANOfSpouse_S5A = False
                Exit Function
            End If
            
'            If Not chkCompulsory(AadhaarOfSpouse_S5A) Then
'            If PANOfSpouse_S5A = "" Then
'                Msg5ASheet = Msg5ASheet + Chr(13) & "Please enter Aadhaar of the spouse in schedule 5A" & Chr(13)
'                ValidatePANOfSpouse_S5A = False
'                Exit Function
'            End If
'            End If
            
            If Not checkfieldspecialcharacter(PANOfSpouse_S5A) Then
                 Msg5ASheet = Msg5ASheet + Chr(13) & "PAN of Spouse in Sheet Sch 5A Cannot Contain Special Characters in schedule 5A & Chr(34)"
                ValidatePANOfSpouse_S5A = False
                Exit Function
            End If
            If Not CheckSpousePAN(PANOfSpouse_S5A) Then
                Msg5ASheet = Msg5ASheet + Chr(13) & "Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet  in schedule 5A" & Chr(13)
                ValidatePANOfSpouse_S5A = False
                Exit Function
            End If
            
    
            
        If AadhaarOfSpouse_S5A <> "" Then
        If Not IsNumeric(AadhaarOfSpouse_S5A) Then
             Msg5ASheet = Msg5ASheet + Chr(13) & "Please enter valid Aadhaar Number in schedule 5A" & Chr(13)
                ValidatePANOfSpouse_S5A = False
                Exit Function
            End If
        
    
        If AadhaarOfSpouse_S5A = "000000000000" Then
           Msg5ASheet = Msg5ASheet + Chr(13) & "Please enter valid Aadhaar Number in schedule 5A" & Chr(13)
                ValidatePANOfSpouse_S5A = False
                Exit Function
        End If
        
        If AadhaarOfSpouse_S5A = "111111111111" Then
            Msg5ASheet = Msg5ASheet + Chr(13) & "Please enter valid Aadhaar Number in schedule 5A" & Chr(13)
                ValidatePANOfSpouse_S5A = False
                Exit Function
        End If
        
    
        If Len(AadhaarOfSpouse_S5A) <> 12 Then
            Msg5ASheet = Msg5ASheet + Chr(13) & "Please enter valid 12 digit Aadhaar Number in schedule 5A" & Chr(13)
                ValidatePANOfSpouse_S5A = False
                Exit Function
        End If
    End If

 
End Function

Function ValidateHP_S5A() As Boolean

ValidateHP_S5A = True
Dim Total As Variant
IncmRecvdHP = Sheet29.Range("S5A_IncRecvdUndHeadHP").value
IncmHeadHP = Sheet29.Range("S5A_HPHeadIncome").value
TDSAmtDedHP = Sheet29.Range("S5A_AmtTDSDeductedHP").value
TDSApprndHP = Sheet29.Range("S5A_TDSApprndOfSpouseHP").value
Total = Sheet29.Range("S5A_IncRecvdUndHeadTotal").value

    If Not chkCompulsory(IncmRecvdHP) Then
'            Msg5ASheetHP = Msg5ASheetHP + Chr(13) & "Receipts recieved from House Property in Sheet SCH5A  is Mandatory.Please fill zero if no Receipts recieved" & Chr(13)
        'SIT-69819 Changed by sai
           Msg5ASheetHP = Msg5ASheetHP + Chr(13) & """Please enter Amount of TDS deducted on income in Schedule 5A. If null, enter Zero""" & Chr(13)
            
            
            
            ValidateHP_S5A = False
    End If
    
    If Not chkCompulsory(IncmHeadHP) Then
    
    'SIT-69819 Changed by sai
            Msg5ASheetHP = Msg5ASheetHP + Chr(13) & """Please enter Amount of TDS deducted on income in Schedule 5A. If null, enter Zero""" & Chr(13)
            ValidateHP_S5A = False
    End If
    
    If Not chkCompulsory(TDSAmtDedHP) Then
'            Msg5ASheetHP = Msg5ASheetHP + Chr(13) & "Amount of TDS deducted on Receipts from House Property in Sheet SCH5A  is Mandatory.Please fill zero if no TDS deducted" & Chr(13)

       'SIT-69818 Changed by sai
             Msg5ASheetHP = Msg5ASheetHP + Chr(13) & """Please enter Amount of TDS deducted on income in Schedule-5A. If null, enter Zero""" & Chr(13)

            ValidateHP_S5A = False
    End If
    
    If Not chkCompulsory(TDSApprndHP) Then
'            Msg5ASheetHP = Msg5ASheetHP + Chr(13) & "TDS apportioned in the hands of spouse from House Property in Sheet SCH5A  is Mandatory.Please fill zero if no TDS apportioned in hands of spouse" & Chr(13)
            
         'SIT-69816 Changed by sai
             Msg5ASheetHP = Msg5ASheetHP + Chr(13) & """Please enter TDS apportioned in the hands of spouse in Schedule 5A. If null, enter Zero""" & Chr(13)
            
            
            ValidateHP_S5A = False
    End If
    
    If IncmHeadHP > IncmRecvdHP Then
    
            Msg5ASheetHP = Msg5ASheetHP + Chr(13) & "Amount apportioned in the hands of the spouse cannot be greater than Receipts received under the head. " & Chr(13)
            
       
           
            
            
            ValidateHP_S5A = False
    End If
    
    
    If TDSApprndHP > TDSAmtDedHP Then
            Msg5ASheetHP = Msg5ASheetHP + Chr(13) & "TDS apportioned in the hands of spouse cannot be greater than Amount of TDS deducted on income at (ii)." & Chr(13)
            ValidateHP_S5A = False
    End If
    
    
    
    If Len(Total) > 14 Then
            Msg5ASheetHP = Msg5ASheetHP + Chr(13) & "total of Receipts recieved under head (i) cannot be greater than 14 digits" & Chr(13)
            ValidateHP_S5A = False
    End If
    
    If Msg5ASheetHP <> "" Then
    Msg5ASheetHP = "House Property in SCH5A" & Chr(13) + Msg5ASheetHP
    End If
    
End Function

Function ValidateBP_S5A() As Boolean

ValidateBP_S5A = True
Dim Total As Variant
IncmRecvdBP = Sheet29.Range("S5A_IncRecvdUndHeadBP").value
IncmHeadBP = Sheet29.Range("S5A_BusHeadIncome").value
TDSAmtDedBP = Sheet29.Range("S5A_AmtTDSDeductedBP").value
TDSApprndBP = Sheet29.Range("S5A_TDSApprndOfSpouseBP").value
Total = Sheet29.Range("S5A_TotalHeadIncome").value
    If Not chkCompulsory(IncmRecvdBP) Then
            Msg5ASheetBP = Msg5ASheetBP + Chr(13) & "Receipts recieved from Business or Profession in Sheet SCH5A  is Mandatory.Please fill zero if no Receipts recieved" & Chr(13)
            ValidateBP_S5A = False
    End If
    
    If Not chkCompulsory(IncmHeadBP) Then
            Msg5ASheetBP = Msg5ASheetBP + Chr(13) & "Amount apportioned in hands of spouse from Business or Profession in Sheet SCH5A  is Mandatory.Please fill zero if no amount is apportioned in hands of spouse" & Chr(13)
            ValidateBP_S5A = False
    End If
    
    If Not chkCompulsory(TDSAmtDedBP) Then
            Msg5ASheetBP = Msg5ASheetBP + Chr(13) & "Amount of TDS deducted on Receipts from Business or Profession in Sheet SCH5A  is Mandatory.Please fill zero if no TDS deducted" & Chr(13)
            ValidateBP_S5A = False
    End If
    
    If Not chkCompulsory(TDSApprndBP) Then
            Msg5ASheetBP = Msg5ASheetBP + Chr(13) & "TDS apportioned in the hands of spouse from Business or Profession in Sheet SCH5A  is Mandatory.Please fill zero if no TDS apportioned in hands of spouse" & Chr(13)
            ValidateBP_S5A = False
    End If
    
    If TDSApprndBP > TDSAmtDedBP Then
            Msg5ASheetBP = Msg5ASheetBP + Chr(13) & "TDS apportioned in the hands of spouse cannot be greater than Amount of TDS deducted on income at (ii). " & Chr(13)
            ValidateBP_S5A = False
    End If
    
    
    If IncmHeadBP > IncmRecvdBP Then
            Msg5ASheetBP = Msg5ASheetBP + Chr(13) & "Amount apportioned in the hands of the spouse cannot be greater than Receipts received under the head. " & Chr(13)
            ValidateBP_S5A = False
    End If

    If Len(Total) > 14 Then
            Msg5ASheetBP = Msg5ASheetBP + Chr(13) & "total of Receipts recieved under head (ii) cannot be greater than 14 digits" & Chr(13)
            ValidateBP_S5A = False
    End If
    
    If Msg5ASheetBP <> "" Then
    Msg5ASheetBP = "Business or Profession in SCH5A" & Chr(13) + Msg5ASheetBP
    End If
    
End Function

Function ValidateCG_S5A() As Boolean
Dim Total As Variant
ValidateCG_S5A = True
Total = Sheet29.Range("S5A_AmtTDSDeductedTotal").value
IncmRecvdCG = Sheet29.Range("S5A_IncRecvdUndHeadCG").value
IncmHeadCG = Sheet29.Range("S5A_CapGainHeadIncome").value
TDSAmtDedCG = Sheet29.Range("S5A_AmtTDSDeductedCG").value
TDSApprndCG = Sheet29.Range("S5A_TDSApprndOfSpouseCG").value

    If Not chkCompulsory(IncmRecvdCG) Then
            Msg5ASheetCG = Msg5ASheetCG + Chr(13) & "Receipts recieved from Capital Gains in Sheet SCH5A  is Mandatory.Please fill zero if no Receipts recieved" & Chr(13)
            ValidateCG_S5A = False
    End If
    
    If Not chkCompulsory(IncmHeadCG) Then
            Msg5ASheetCG = Msg5ASheetCG + Chr(13) & "Amount apportioned in hands of spouse from Capital Gains in Sheet SCH5A  is Mandatory.Please fill zero if no amount is apportioned in hands of spouse" & Chr(13)
            ValidateCG_S5A = False
    End If
    
    If Not chkCompulsory(TDSAmtDedCG) Then
            Msg5ASheetCG = Msg5ASheetCG + Chr(13) & "Amount of TDS deducted on Receipts from Capital Gains in Sheet SCH5A  is Mandatory.Please fill zero if no TDS deducted" & Chr(13)
            ValidateCG_S5A = False
    End If
    
'Change-14.05.2023.101.5A -> SIT-49873
'    If Not chkCompulsory(TDSAmtDedCG) Then
    If Not chkCompulsory(TDSApprndCG) Then
'---
            Msg5ASheetCG = Msg5ASheetCG + Chr(13) & "TDS apportioned in the hands of spouse from Capital Gains in Sheet SCH5A  is Mandatory.Please fill zero if no TDS apportioned in hands of spouse" & Chr(13)
            ValidateCG_S5A = False
    End If
    
    If IncmHeadCG > IncmRecvdCG Then
            Msg5ASheetCG = Msg5ASheetCG + Chr(13) & "Amount apportioned in the hands of the spouse cannot be greater than Receipts received under the head. " & Chr(13)
            ValidateCG_S5A = False
    End If
    
    If TDSApprndCG > TDSAmtDedCG Then
            Msg5ASheetCG = Msg5ASheetCG + Chr(13) & "TDS apportioned in the hands of the spouse cannot be greater than Amount of TDS deducted on income at (ii). " & Chr(13)
            ValidateCG_S5A = False
    End If

    If Len(Total) > 14 Then
            Msg5ASheetCG = Msg5ASheetCG + Chr(13) & "total of Receipts recieved under head (iii) cannot be greater than 14 digits" & Chr(13)
            ValidateCG_S5A = False
    End If
    
     If Msg5ASheetCG <> "" Then
        Msg5ASheetCG = "Capital Gains in SCH5A" & Chr(13) + Msg5ASheetCG
    End If
End Function

Function ValidateOS_S5A() As Boolean
Dim Total As Variant
ValidateOS_S5A = True
Total = Sheet29.Range("S5A_TDSApprndOfSpouseTotal").value
IncmRecvdOS = Sheet29.Range("S5A_IncRecvdUndHeadOS").value
IncmHeadOS = Sheet29.Range("S5A_OtherSourcesHeadIncome").value
TDSAmtDedOS = Sheet29.Range("S5A_AmtTDSDeductedOS").value
TDSApprndOS = Sheet29.Range("S5A_TDSApprndOfSpouseOS").value

    If Not chkCompulsory(IncmRecvdOS) Then
            Msg5ASheetOS = Msg5ASheetOS + Chr(13) & "Receipts recieved from Other Sources in Sheet SCH5A  is Mandatory.Please fill zero if no Receipts recieved" & Chr(13)
            ValidateOS_S5A = False
    End If
    
    If Not chkCompulsory(IncmHeadOS) Then
            Msg5ASheetOS = Msg5ASheetOS + Chr(13) & "Amount apportioned in hands of spouse from Other Sources in Sheet SCH5A  is Mandatory.Please fill zero if no amount is apportioned in hands of spouse" & Chr(13)
            ValidateOS_S5A = False
    End If
    
    If Not chkCompulsory(TDSAmtDedOS) Then
            Msg5ASheetOS = Msg5ASheetOS + Chr(13) & "Amount of TDS deducted on Receipts from Other Sources in Sheet SCH5A  is Mandatory.Please fill zero if no TDS deducted" & Chr(13)
            ValidateOS_S5A = False
    End If
    
'Change-14.05.2023.101.5A -> SIT-49873
'    If Not chkCompulsory(TDSAmtDedOS) Then
    If Not chkCompulsory(TDSApprndOS) Then
'---
            Msg5ASheetOS = Msg5ASheetOS + Chr(13) & "TDS apportioned in the hands of spouse from Other Sources in Sheet SCH5A  is Mandatory.Please fill zero if no TDS apportioned in hands of spouse" & Chr(13)
            ValidateOS_S5A = False
    End If
    
    If IncmHeadOS > IncmRecvdOS Then
            Msg5ASheetOS = Msg5ASheetOS + Chr(13) & "Amount apportioned in the hands of the spouse cannot be greater than Receipts received under the head. " & Chr(13)
            ValidateOS_S5A = False
    End If
    
    If TDSApprndOS > TDSAmtDedOS Then
            Msg5ASheetOS = Msg5ASheetOS + Chr(13) & "TDS apportioned in the hands of the spouse cannot be greater than Amount of TDS deducted on income at (ii). " & Chr(13)
            ValidateOS_S5A = False
    End If

    If Len(Total) > 14 Then
            Msg5ASheetOS = Msg5ASheetOS + Chr(13) & "total of Receipts recieved under head (iv) cannot be greater than 14 digits" & Chr(13)
            ValidateOS_S5A = False
    End If
    
    If Msg5ASheetOS <> "" Then
        Msg5ASheetOS = "Other Sources in SCH5A" & Chr(13) + Msg5ASheetOS
    End If
    
End Function

Function CheckSpousePAN(PAN As Variant) As Boolean
On Error Resume Next
    CheckSpousePAN = True
    If Len(PAN) > 0 Then
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
            CheckSpousePAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
            CheckSpousePAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
            CheckSpousePAN = False
            Exit Function
        End If
         If Mid(PAN, 4, 1) <> "P" Then
                CheckSpousePAN = False
                Exit Function
        End If
            
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
            CheckSpousePAN = False
            Exit Function
        End If
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
            CheckSpousePAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
            CheckSpousePAN = False
            Exit Function
        End If
    End If
End Function


