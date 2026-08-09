Attribute VB_Name = "Sheet20"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Public lockcount_80EEA, lockcount_80EE As Variant
Public lock_80EEA_flag, lock_80EE_flag As Boolean
Public IfsCode2 As Variant
Public sPassword As String

Private Sub Worksheet_Activate()
 Sheet20.Protect Password:=getmsgstate

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
Dim TargtVal, targetadd As Variant
Dim IffCode As Range
Dim IfsCode As Range


Application.EnableEvents = False
Application.ScreenUpdating = False

CalculateTaxCheckFlag = False
Sheet20.Unprotect Password:=getmsgstate
'Target.EntireColumn.AutoFit

Dim lock80EE_chk, lock80EEA_chk As Boolean

'If Not Application.Intersect(Sheet20.Range("Stampduty.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet20.Range("LoanfrmBankOrInstitute.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet20.Range("IFSC.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet20.Range("PAN.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet20.Range("bankName.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet20.Range("loanAccNum.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet20.Range("loanDate.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet20.Range("loanAmt.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet20.Range("loanOutstanding.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet20.Range("Intrst.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Then


'Commented by Ayush_05/05/2025 for 0.7 changes
If Not Application.Intersect(Sheet20.Range("Stampduty.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet20.Range("LoanfrmBankOrInstitute.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet20.Range("bankName.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet20.Range("loanAccNum.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet20.Range("loanDate.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet20.Range("loanAmt.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet20.Range("loanOutstanding.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet20.Range("Intrst.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Then


'If Target.Value <> "" Then
     lock80EE_80EEAcount
'End If

If lockcount_80EEA > 0 Then
  lockcount_80EE = 0
  lock80EE_80EEA_chk
Else
  lock80EE_80EEA_chk
End If


End If


'Commented by Ayush_05/05/2025 for 0.7 changes
'If Not Application.Intersect(Sheet20.Range("ResidentialHP.80E"), Sheet20.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet20.Range("LoanfrmBankOrInstitute.80EE"), Sheet20.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet20.Range("IFSC.80EE"), Sheet20.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet20.Range("PAN.80EE"), Sheet20.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet20.Range("bankName.80EE"), Sheet20.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet20.Range("loanAccNum.80EE"), Sheet20.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet20.Range("loanDate.80EE"), Sheet20.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet20.Range("loanAmt.80EE"), Sheet20.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet20.Range("loanOutstanding.80EE"), Sheet20.Range(Target.Address)) Is Nothing Or _
'Not Application.Intersect(Sheet20.Range("Intrst.80EE"), Sheet20.Range(Target.Address)) Is Nothing Then


If Not Application.Intersect(Sheet20.Range("LoanfrmBankOrInstitute.80EE"), Sheet20.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet20.Range("bankName.80EE"), Sheet20.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet20.Range("loanAccNum.80EE"), Sheet20.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet20.Range("loanDate.80EE"), Sheet20.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet20.Range("loanAmt.80EE"), Sheet20.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet20.Range("loanOutstanding.80EE"), Sheet20.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet20.Range("Intrst.80EE"), Sheet20.Range(Target.Address)) Is Nothing Then


'If Target.Value <> "" Then
     lock80EE_80EEAcount
'End If
'80EE
If lockcount_80EE > 0 Then
  lockcount_80EEA = 0
  lock80EE_80EEA_chk
  
  Else
  lock80EE_80EEA_chk
End If


End If

'----------------------------
'Ayush_18/04

'If Not Application.Intersect(Sheet20.Range("LoanfrmBankOrInstitute.80EE"), Sheet20.Range(Target.Address)) Is Nothing Then
'LockUnlock80EEA_OLD
'End If
'
'If Not Application.Intersect(Sheet20.Range("LoanfrmBankOrInstitute.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Then
'LockUnlock80EE_OLD
'End If


If Not Application.Intersect(Sheet20.Range("loanDate.80E"), Sheet20.Range(Target.Address)) Is Nothing Then
       
        If Not (ValidateDate_80E(Target)) Then
        
         End If
           GoTo endline
End If

If Not Application.Intersect(Sheet20.Range("loanDate.80EE"), Sheet20.Range(Target.Address)) Is Nothing Then
       
        If Not (ValidateDate_80EE(Target)) Then
        
         End If
           GoTo endline
End If

If Not Application.Intersect(Sheet20.Range("loanDate.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Then
       
        If Not (ValidateDate_80EEA(Target)) Then
        
         End If
           GoTo endline
End If


If Not Application.Intersect(Sheet20.Range("loanDate.80EEB"), Sheet20.Range(Target.Address)) Is Nothing Then
       
        If Not (ValidateDate_80EEB(Target)) Then
        
         End If
           GoTo endline
End If
'--------------------------------------




'Commented by Ayush_05/05/2025 for 0.7 changes
'Ayush---------------16/04/2025
 
'If Not Application.Intersect(Sheet20.Range("LoanfrmBankOrInstitute.80E"), Sheet20.Range(Target.Address)) Is Nothing Then
'
'
'If Sheet20.Range(Target.Address).value <> "" Or Sheet11.Range(Target.Address).value <> "(Select)" Then
'
'            If Mid(Sheet20.Range(Target.Address).value, 1, 1) = "I" Then
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "E")).ClearContents
'
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Locked = True
'
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Interior.Color = "&HD8D8D8"
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "G")).ClearContents
'
'            'Sheet20.Range(Replace(Target.Address, "D", "G")).Locked = True
'
'            Sheet20.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HCCFFCC"
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).ClearContents
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Locked = False
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HCCFFCC"
'            ElseIf Mid(Sheet20.Range(Target.Address).value, 1, 1) = "B" Then
'
'
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).ClearContents
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Locked = False
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Interior.Color = "&HCCFFCC"
'
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).ClearContents
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Locked = True
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HD8D8D8"
'
'
'            Else
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "E")).ClearContents
'
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Locked = False
'
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Interior.Color = "&HCCFFCC"
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "G")).ClearContents
'
'            Sheet20.Range(Replace(Target.Address, "D", "G")).Locked = False
'
'            Sheet20.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HCCFFCC"
'
'
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).ClearContents
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Locked = False
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HCCFFCC"
'
'
'            End If
'End If
'End If
'
'
''-----------------------


'Commented by Ayush_05/05/2025 for 0.7 changes
'If Not Application.Intersect(Sheet20.Range("PAN.80E"), Sheet20.Range(Target.Address)) Is Nothing Then
'
''If (Target.name.name) = "sheet1.PANRepAssessee" Then
'
'    Dim TargetPAN1  As String
'
'    TargetPAN1 = ""
'
'    TargetPAN1 = Sheet20.Range(Target.Address).value
'
'    'Target.value = UCase(Target.value)
'
'    If TargetPAN1 <> "" Then
'
'        'If Not mdIncomeDetails.CheckPAN1(UCase(Trim(Range(TargetPAN1).value))) Then
'
'        If Not mdIncomeDetails.CheckPAN(UCase(Trim(TargetPAN1))) Then
'
'            'Ayush_PAN
'
'            'fmsgboxsmall "* Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet." & Chr(10) & Chr(13) & "4th letter must be        'P' in Sheet Income Details" ', vbOKCancel, mdCommon.vbMessageCaption
'
'            fmsgboxsmall "* Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet." ', vbOKCancel, mdCommon.vbMessageCaption
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'           Sheet20.Range(Target.Address).ClearContents
'
'      '      TargetPAN1 = ""
'
''            Sheet20.Range(TargetPAN1) = ""
'
''                    Sheet20.Range(TargetPAN1).Offset(0, 1) = ""
'
''
'
'            GoTo endline
'
'        End If
'
'       ' TargetPAN1 = UCase(TargetPAN1)
'
'       Sheet20.Range(Target.Address).value = UCase(TargetPAN1)
'
'
''        If Range(TargetPAN1).value = Sheet1.Range("sheet1.PAN").value Then
'
''
'
''            fmsgboxsmall "* PAN of the representative assessee And assessee (under personal information) cannot be same in Sheet Income Details"        ', vbOKCancel, mdCommon.vbMessageCaption
'
''
'
''            Sheet20.Range(TargetPAN1).Select
'
''
'
''            Sheet20.Range(TargetPAN1).value = ""
'
''
'
''            GoTo endline
'
''
'
''        End If
'
'    End If
'
'End If

 




'Commented by Ayush_05/05/2025 for 0.7 changes

'If Target.Validation.Type = 6 Then
'    Target.value = UCase(Target.value)
'
'
'
'
'
'
''--------17/12/2024----
''17/12/2024
' Set IfsCode = Range("IFSC.80E")
'
'
'    Dim sourceAddress, Address, address1 As Variant
'    Dim rangecell As Range
'    Dim BankNameFromIFSC As Variant
'
'    If Not Application.Intersect(IfsCode, Range(Target.Address)) Is Nothing Then
'         For Each rangecell In Sheet20.Range("IFSC.80E")
'            rangecell.value = UCase(rangecell.value)
'
'     If rangecell.value <> "" Then
'
'           If Not CheckIFSC(UCase(TargtVal)) Then
'                    'fmsgboxsmall "* Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
'                     fmsgbox ("* ""Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
'                    Sheet20.Range(targetadd) = ""
'                    Sheet20.Range(targetadd).Offset(0, 1) = ""
'                    GoTo endline
'
'            ElseIf Not ValidateIFSCList(UCase(TargtVal)) Then
'                   ' fmsgboxsmall "* Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
'                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
'                    Sheet20.Range(targetadd) = ""
'                    Sheet20.Range(targetadd).Offset(0, 1) = ""
'                    GoTo endline
'
'                  'Call CheckIFSC1(TargtVal, targetadd)
'                  'GoTo endline
'            Else:
'                       sourceAddress = rangecell.Address
'                       Address = Replace(sourceAddress, "E", "G")
'
'                     If Mid(rangecell, 1, 4) = "GIFT" Then
'                             BankNameFromIFSC = GetBankName(rangecell)
'                     Else
'                             BankNameFromIFSC = GetBankName(Mid(rangecell, 1, 4))
'                     End If
'
'
'                If Mid(rangecell, 1, 4) = "GIFT" And BankNameFromIFSC <> "" Then
'                   Sheet20.Range(Address).value = BankNameFromIFSC
'                ElseIf Mid(rangecell, 1, 4) <> "GIFT" Then
'                   Sheet20.Range(Address).value = BankNameFromIFSC
'                End If
'
'                If Mid(rangecell, 1, 4) <> "GIFT" Then
'                    If (Sheet20.Range(Address).value = "") Then
'                     'MsgBox "IFS Code: Please enter valid IFS Code. Refer to your bank for valid codes. (1st 4 Alpha, followed by 0 and 6 digits)"
'                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
'                     Address = Replace(sourceAddress, "K", "J")
'                     Sheet20.Range(Address).value = ""
'                    GoTo endline
'                    End If
'                End If
'
'            End If
'         End If
'    Next
'    GoTo endline
'End If
'End If



'80EE______________starts
'Sheet20.Unprotect Password:=getmsgstate


'Commented by Ayush_05/05/2025 for 0.7 changes

''Ayush---------------16/04/2025
'
'If Not Application.Intersect(Sheet20.Range("LoanfrmBankOrInstitute.80EE"), Sheet20.Range(Target.Address)) Is Nothing Then
'
'
'If Sheet20.Range(Target.Address).value <> "" Or Sheet11.Range(Target.Address).value <> "(Select)" Then
'
'            If Mid(Sheet20.Range(Target.Address).value, 1, 1) = "I" Then
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "E")).ClearContents
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Locked = True
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Interior.Color = "&HD8D8D8"
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "G")).ClearContents
'
'            'Sheet20.Range(Replace(Target.Address, "D", "G")).Locked = True
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HCCFFCC"
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).ClearContents
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Locked = False
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HCCFFCC"
'
'            ElseIf Mid(Sheet20.Range(Target.Address).value, 1, 1) = "B" Then
'
'
'
'
'
'           Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).ClearContents
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Locked = False
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Interior.Color = "&HCCFFCC"
'
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).ClearContents
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Locked = True
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HD8D8D8"
'
'
'            Else
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "E")).ClearContents
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Locked = False
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Interior.Color = "&HCCFFCC"
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "G")).ClearContents
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "G")).Locked = False
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HCCFFCC"
'
'
'
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "F")).ClearContents
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Locked = False
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HCCFFCC"
'
'
'            End If
'End If
'End If
 
'-----------------------

'Commented by Ayush_05/05/2025 for 0.7 changes

'If Not Application.Intersect(Sheet20.Range("PAN.80EE"), Sheet20.Range(Target.Address)) Is Nothing Then
'
''If (Target.name.name) = "sheet1.PANRepAssessee" Then
'
'    Dim TargetPAN2  As String
'
'    TargetPAN2 = ""
'
'    TargetPAN2 = Sheet20.Range(Target.Address).value
'
'    'Target.value = UCase(Target.value)
'
'    If TargetPAN2 <> "" Then
'
'        'If Not mdIncomeDetails.CheckPAN1(UCase(Trim(Range(TargetPAN1).value))) Then
'
'        If Not mdIncomeDetails.CheckPAN(UCase(Trim(TargetPAN2))) Then
'
'            'Ayush_PAN
'
'            'fmsgboxsmall "* Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet." & Chr(10) & Chr(13) & "4th letter must be        'P' in Sheet Income Details" ', vbOKCancel, mdCommon.vbMessageCaption
'
'            fmsgboxsmall "* Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet." ', vbOKCancel, mdCommon.vbMessageCaption
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'           Sheet20.Range(Target.Address).ClearContents
'
'      '      TargetPAN1 = ""
'
''            Sheet20.Range(TargetPAN1) = ""
'
''                    Sheet20.Range(TargetPAN1).Offset(0, 1) = ""
'
''
'
'            GoTo endline
'
'        End If
'
'       ' TargetPAN1 = UCase(TargetPAN1)
'
'       Sheet20.Range(Target.Address).value = UCase(TargetPAN2)
'
'
''
'
'    End If
'
'End If
'
'If Target.Validation.Type = 6 Then
'    Target.value = UCase(Target.value)
'
'
'
'
'
'
''--------17/12/2024----
''17/12/2024
''Public IfsCode2 As Variant
'
' Set IfsCode2 = Range("IFSC.80EE")
'
'
'    Dim sourceAddress80EE, Address80EE, Address80EE1 As Variant
'    Dim rangecell80EE As Range
'    Dim BankNameFromIFSC80EE As Variant
'
'    If Not Application.Intersect(IfsCode2, Range(Target.Address)) Is Nothing Then
'         For Each rangecell80EE In Sheet20.Range("IFSC.80EE")
'            rangecell80EE.value = UCase(rangecell80EE.value)
'
'     If rangecell80EE.value <> "" Then
'
'           If Not CheckIFSC(UCase(TargtVal)) Then
'                    'fmsgboxsmall "* Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
'                     fmsgbox ("* ""Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
'                    Sheet20.Range(targetadd) = ""
'                    Sheet20.Range(targetadd).Offset(0, 1) = ""
'                    GoTo endline
'
'            ElseIf Not ValidateIFSCList(UCase(TargtVal)) Then
'                   ' fmsgboxsmall "* Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
'                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
'                    Sheet20.Range(targetadd) = ""
'                    Sheet20.Range(targetadd).Offset(0, 1) = ""
'                    GoTo endline
'
'                  'Call CheckIFSC1(TargtVal, targetadd)
'                  'GoTo endline
'            Else:
'                       sourceAddress80EE = rangecell80EE.Address
'                       Address80EE = Replace(sourceAddress80EE, "E", "G")
'
'                     If Mid(rangecell80EE, 1, 4) = "GIFT" Then
'                             BankNameFromIFSC80EE = GetBankName(rangecell80EE)
'                     Else
'                             BankNameFromIFSC80EE = GetBankName(Mid(rangecell80EE, 1, 4))
'                     End If
'
'
'                If Mid(rangecell80EE, 1, 4) = "GIFT" And BankNameFromIFSC80EE <> "" Then
'                   Sheet20.Range(Address80EE).value = BankNameFromIFSC80EE
'                ElseIf Mid(rangecell80EE, 1, 4) <> "GIFT" Then
'                   Sheet20.Range(Address80EE).value = BankNameFromIFSC80EE
'                End If
'
'                If Mid(rangecell80EE, 1, 4) <> "GIFT" Then
'                    If (Sheet20.Range(Address80EE).value = "") Then
'                     'MsgBox "IFS Code: Please enter valid IFS Code. Refer to your bank for valid codes. (1st 4 Alpha, followed by 0 and 6 digits)"
'                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
'                     Address80EE = Replace(sourceAddress80EE, "K", "J")
'                     Sheet20.Range(Address80EE).value = ""
'                    GoTo endline
'                    End If
'                End If
'
'            End If
'         End If
'    Next
'    GoTo endline
'End If
'End If
'
'
''80EE______________Ends
'
''80EEA_________starts
'
'Sheet20.Unprotect Password:=getmsgstate


'Commented by Ayush_05/05/2025 for 0.7 changes

'Ayush---------------16/04/2025
 
'If Not Application.Intersect(Sheet20.Range("LoanfrmBankOrInstitute.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Then
'
'
'If Sheet20.Range(Target.Address).value <> "" Or Sheet11.Range(Target.Address).value <> "(Select)" Then
'
'            If Mid(Sheet20.Range(Target.Address).value, 1, 1) = "I" Then
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "E")).ClearContents
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Locked = True
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Interior.Color = "&HD8D8D8"
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "G")).ClearContents
'
'            'Sheet20.Range(Replace(Target.Address, "D", "G")).Locked = True
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HCCFFCC"
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).ClearContents
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Locked = False
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HCCFFCC"
'
'            ElseIf Mid(Sheet20.Range(Target.Address).value, 1, 1) = "B" Then
'
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).ClearContents
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Locked = False
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Interior.Color = "&HCCFFCC"
'
'
'
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).ClearContents
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Locked = True
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HD8D8D8"
'
'
'            Else
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "E")).ClearContents
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Locked = False
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Interior.Color = "&HCCFFCC"
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "G")).ClearContents
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "G")).Locked = False
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HCCFFCC"
'
'
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).ClearContents
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Locked = False
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HCCFFCC"
'
'
'            End If
'End If
'End If
'
''-----------------------
'
'
'
'If Not Application.Intersect(Sheet20.Range("PAN.80EEA"), Sheet20.Range(Target.Address)) Is Nothing Then
'
''If (Target.name.name) = "sheet1.PANRepAssessee" Then
'
'    Dim TargetPAN80EEA  As String
'
'    TargetPAN80EEA = ""
'
'    TargetPAN80EEA = Sheet20.Range(Target.Address).value
'
'    'Target.value = UCase(Target.value)
'
'    If TargetPAN80EEA <> "" Then
'
'        'If Not mdIncomeDetails.CheckPAN1(UCase(Trim(Range(TargetPAN1).value))) Then
'
'        If Not mdIncomeDetails.CheckPAN(UCase(Trim(TargetPAN80EEA))) Then
'
'            'Ayush_PAN
'
'            'fmsgboxsmall "* Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet." & Chr(10) & Chr(13) & "4th letter must be        'P' in Sheet Income Details" ', vbOKCancel, mdCommon.vbMessageCaption
'
'            fmsgboxsmall "* Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet." ', vbOKCancel, mdCommon.vbMessageCaption
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'           Sheet20.Range(Target.Address).ClearContents
'
'
''
'
'            GoTo endline
'
'        End If
'
'       ' TargetPAN1 = UCase(TargetPAN1)
'
'       Sheet20.Range(Target.Address).value = UCase(TargetPAN80EEA)
'
'
'    End If
'
'End If
'
'
'
'If Target.Validation.Type = 6 Then
'    Target.value = UCase(Target.value)
'
'
'
'
'
'
''--------17/12/2024----
''17/12/2024
' Set IfsCode2 = Range("IFSC.80EEA")
'
'
'    Dim sourceAddress80EEA, Address80EEA, Address80EEA1 As Variant
'    Dim rangecell80EEA As Range
'    Dim BankNameFromIFSC80EEA As Variant
'
'    If Not Application.Intersect(IfsCode2, Range(Target.Address)) Is Nothing Then
'         For Each rangecell80EEA In Sheet20.Range("IFSC.80EEA")
'            rangecell80EEA.value = UCase(rangecell80EEA.value)
'
'     If rangecell80EEA.value <> "" Then
'
'           If Not CheckIFSC(UCase(TargtVal)) Then
'                    'fmsgboxsmall "* Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
'                     fmsgbox ("* ""Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
'                    Sheet20.Range(targetadd) = ""
'                    Sheet20.Range(targetadd).Offset(0, 1) = ""
'                    GoTo endline
'
'            ElseIf Not ValidateIFSCList(UCase(TargtVal)) Then
'                   ' fmsgboxsmall "* Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
'                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
'                    Sheet20.Range(targetadd) = ""
'                    Sheet20.Range(targetadd).Offset(0, 1) = ""
'                    GoTo endline
'
'                  'Call CheckIFSC1(TargtVal, targetadd)
'                  'GoTo endline
'            Else:
'                       sourceAddress80EEA = rangecell80EEA.Address
'                       Address80EEA = Replace(sourceAddress80EEA, "E", "G")
'
'                     If Mid(rangecell80EEA, 1, 4) = "GIFT" Then
'                             BankNameFromIFSC80EEA = GetBankName(rangecell80EEA)
'                     Else
'                             BankNameFromIFSC80EEA = GetBankName(Mid(rangecell80EEA, 1, 4))
'                     End If
'
'
'                If Mid(rangecell80EEA, 1, 4) = "GIFT" And BankNameFromIFSC80EEA <> "" Then
'                   Sheet20.Range(Address80EEA).value = BankNameFromIFSC80EEA
'                ElseIf Mid(rangecell80EEA, 1, 4) <> "GIFT" Then
'                   Sheet20.Range(Address80EEA).value = BankNameFromIFSC80EEA
'                End If
'
'                If Mid(rangecell80EEA, 1, 4) <> "GIFT" Then
'                    If (Sheet20.Range(Address80EEA).value = "") Then
'                     'MsgBox "IFS Code: Please enter valid IFS Code. Refer to your bank for valid codes. (1st 4 Alpha, followed by 0 and 6 digits)"
'                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
'                     Address80EEA = Replace(sourceAddress80EEA, "K", "J")
'                     Sheet20.Range(Address80EEA).value = ""
'                    GoTo endline
'                    End If
'                End If
'
'            End If
'         End If
'    Next
'    GoTo endline
'End If
'End If
'
''80EEA_________Ends

'Commented by Ayush_05/05/2025 for 0.7 changes
'80EEB________Starts
'Sheet20.Unprotect Password:=getmsgstate
'
'
'
'
''Ayush---------------16/04/2025
'
'If Not Application.Intersect(Sheet20.Range("LoanfrmBankOrInstitute.80EEB"), Sheet20.Range(Target.Address)) Is Nothing Then
'
'
'If Sheet20.Range(Target.Address).value <> "" Or Sheet11.Range(Target.Address).value <> "(Select)" Then
'
'            If Mid(Sheet20.Range(Target.Address).value, 1, 1) = "I" Then
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "E")).ClearContents
'
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Locked = True
'
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Interior.Color = "&HD8D8D8"
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "G")).ClearContents
'
'            'Sheet20.Range(Replace(Target.Address, "D", "G")).Locked = True
'
'            Sheet20.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HCCFFCC"
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).ClearContents
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Locked = False
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HCCFFCC"
'            ElseIf Mid(Sheet20.Range(Target.Address).value, 1, 1) = "B" Then
'
'
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).ClearContents
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Locked = False
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Interior.Color = "&HCCFFCC"
'
'
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).ClearContents
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Locked = True
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HD8D8D8"
'
'
'            Else
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "E")).ClearContents
'
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Locked = False
'
'            Sheet20.Range(Replace(Target.Address, "D", "E")).Interior.Color = "&HCCFFCC"
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "G")).ClearContents
'
'            Sheet20.Range(Replace(Target.Address, "D", "G")).Locked = False
'
'            Sheet20.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HCCFFCC"
'
'
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).ClearContents
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Locked = False
'
'            Sheet20.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HCCFFCC"
'
'
'            End If
'End If
'End If
'
''-----------------------
'
'
'If Not Application.Intersect(Sheet20.Range("PAN.80EEB"), Sheet20.Range(Target.Address)) Is Nothing Then
'
''If (Target.name.name) = "sheet1.PANRepAssessee" Then
'
'    Dim TargetPAN80EEB  As String
'
'    TargetPAN80EEB = ""
'
'    TargetPAN80EEB = Sheet20.Range(Target.Address).value
'
'    'Target.value = UCase(Target.value)
'
'    If TargetPAN80EEB <> "" Then
'
'        'If Not mdIncomeDetails.CheckPAN1(UCase(Trim(Range(TargetPAN80EEB).value))) Then
'
'        If Not mdIncomeDetails.CheckPAN(UCase(Trim(TargetPAN80EEB))) Then
'
'            'Ayush_PAN
'
'            'fmsgboxsmall "* Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet." & Chr(10) & Chr(13) & "4th letter must be        'P' in Sheet Income Details" ', vbOKCancel, mdCommon.vbMessageCaption
'
'            fmsgboxsmall "* Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet." ', vbOKCancel, mdCommon.vbMessageCaption
'
'            Sheet20.Unprotect Password:=getmsgstate
'
'           Sheet20.Range(Target.Address).ClearContents
'
'
''
'
'            GoTo endline
'
'        End If
'
'       ' TargetPAN1 = UCase(TargetPAN1)
'
'       Sheet20.Range(Target.Address).value = UCase(TargetPAN80EEB)
'
'
'    End If
'
'End If
'
'If Target.Validation.Type = 6 Then
'    Target.value = UCase(Target.value)
'
'
'
'
'
'
''--------17/12/2024----
''17/12/2024
' Set IfsCode2 = Range("IFSC.80EEB")
'
'
'    Dim sourceAddress80EEB, Address80EEB, Address80EEB1 As Variant
'    Dim rangecell80EEB As Range
'    Dim BankNameFromIFSC80EEB As Variant
'
'    If Not Application.Intersect(IfsCode2, Range(Target.Address)) Is Nothing Then
'         For Each rangecell80EEB In Sheet20.Range("IFSC.80EEB")
'            rangecell80EEB.value = UCase(rangecell80EEB.value)
'
'     If rangecell80EEB.value <> "" Then
'
'           If Not CheckIFSC(UCase(TargtVal)) Then
'                    'fmsgboxsmall "* Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
'                     fmsgbox ("* ""Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
'                    Sheet20.Range(targetadd) = ""
'                    Sheet20.Range(targetadd).Offset(0, 1) = ""
'                    GoTo endline
'
'            ElseIf Not ValidateIFSCList(UCase(TargtVal)) Then
'                   ' fmsgboxsmall "* Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
'                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
'                    Sheet20.Range(targetadd) = ""
'                    Sheet20.Range(targetadd).Offset(0, 1) = ""
'                    GoTo endline
'
'                  'Call CheckIFSC1(TargtVal, targetadd)
'                  'GoTo endline
'            Else:
'                       sourceAddress80EEB = rangecell80EEB.Address
'                       Address80EEB = Replace(sourceAddress80EEB, "E", "G")
'
'                     If Mid(rangecell80EEB, 1, 4) = "GIFT" Then
'                             BankNameFromIFSC80EEB = GetBankName(rangecell80EEB)
'                     Else
'                             BankNameFromIFSC80EEB = GetBankName(Mid(rangecell80EEB, 1, 4))
'                     End If
'
'
'                If Mid(rangecell80EEB, 1, 4) = "GIFT" And BankNameFromIFSC80EEB <> "" Then
'                   Sheet20.Range(Address80EEB).value = BankNameFromIFSC80EEB
'                ElseIf Mid(rangecell80EEB, 1, 4) <> "GIFT" Then
'                   Sheet20.Range(Address80EEB).value = BankNameFromIFSC80EEB
'                End If
'
'                If Mid(rangecell80EEB, 1, 4) <> "GIFT" Then
'                    If (Sheet20.Range(Address80EEB).value = "") Then
'                     'MsgBox "IFS Code: Please enter valid IFS Code. Refer to your bank for valid codes. (1st 4 Alpha, followed by 0 and 6 digits)"
'                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
'                     Address80EEB = Replace(sourceAddress80EEB, "K", "J")
'                     Sheet20.Range(Address80EEB).value = ""
'                    GoTo endline
'                    End If
'                End If
'
'            End If
'         End If
'    Next
'    GoTo endline
'End If
'End If



'80EEB______Ends


endline:

Application.EnableEvents = True
Sheet20.Protect Password:=getmsgstate
Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub

Function GetBankName(IFSC As Variant) As String
On Error Resume Next
GetBankName = ""
Dim rangecells As Range
Dim Address As Variant
For Each rangecells In Sheet8.Range("IFSC_StartLetter")

If UCase(rangecells.value) = UCase(IFSC) Then
Address = rangecells.Address
Address = Replace(Address, "E", "F")
GetBankName = Sheet8.Range(Address).value
Exit Function
End If

Next

'If GetBankName = "" Then
'GetBankName = False
'End If

End Function


Function ValidateDate_80E(ByVal Target As Range) As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim Date_80E
    Dim mIntCells As Long
    Dim mIntCtr As Long
    
    
    mIntCells = Sheet20.Range("loanDate.80E").count
    ValidateDate_80E = True

    Dim end_80E
    end_80E = 0
    Set rangecells = Sheet20.Range("loanDate.80E").Cells

    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            end_80E = end_80E + 1
        End If
    Next

    ReDim Date_80E(end_80E)

    For i = 1 To end_80E
        Date_80E(i) = rangecells.item(i).value
'        If Not chkCompulsory(Date_80E(i)) Then
'            'fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is mandatory") '
'            fmsgboxsmall ("* ""Please provide Date of taking loan in schedule 80E"" at Sr. No  " & i & "")
'            ValidateDate_80E = False
'            Exit Function
'        End If

        If Not CheckDateddmmyyyy(Date_80E(i)) Then
            ValidateDate_80E = False
'            fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is invalid. Pl enter in dd/mm/yyyy format")
'Ayush_10/02
'              fmsgboxsmall ("Please enter a valid date in dd/mm/yyyy format at Sr. No  " & i & "  in Sheet IT ")
                fmsgboxsmall ("* ""Please enter date in valid format"" at Sr. No  " & i & "")
            Target.value = ""
            Exit Function
        Else
            Date_80E(i) = Dformat(Date_80E(i), "")
'            'Ayush_13
''            If Not ChkMinInclusiveDate(DateDep_IT(i), "2023-04-01") Then
'             If Not ChkMinInclusiveDate(Date_24b(i), "2024-04-01") Then
'            'Error Message Changed By Ayush for the SIT-56341
'               ' fmsgboxsmall ("Date of deposit cannot be prior to 01-04-2023 at Sr. No " & i & " in Sheet IT  ")
''               fmsgboxsmall ("""Date of deposit cannot be prior to 01-04-2023""")
'                fmsgboxsmall ("""Date of deposit cannot be prior to 01-04-2024""")
'                Target.value = ""
'                ValidateDate_24b = False
'                Exit Function
'            End If
            
            
            If Not ChkMaxDate_24b(Trim(Date_80E(i)), "31-03-2025") Then
                fmsgboxsmall ("* Date can not be after 31/03/2025 at Sr. No " & i & ".")
                Target.value = ""
                ValidateDate_80E = False
                 

          Exit Function
        End If
            
'             If Date_24b(i) > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'
'                fmsgboxsmall ("Date of Deposit at Sr. No  " & i & " in Sheet IT can not be beyond system date")
'                Target.value = ""
'                Target.Select
'                ValidateDate_24b = False
'                Exit Function
'            End If

             
          '-------------------------------------------------

'            If Sheet1.Range("Sheet1.ReturnFileSec") = "139(8A)" Then
''                If ChkMinInclusiveDate(DateDep_IT(i), "2024-01-01") Then
''                    fmsgboxsmall ("Please enter the details of 140B in Sl No 14A of Part B-ATI and not in this schedule, so as to get the correct computation of your tax liability")
' 'Ayush_UR
'            If ChkMinInclusiveDate(DateDep_IT(i), "2025-01-01") Then
'                   fmsgboxsmall ("Date of deposit can't be after 31-12-2024 in schedule IT. Please enter payments made after 31st December 2024 in schedule IT 1.")
'
'
'                    Target.value = ""
'                End If
'            End If
        End If
    Next

End Function


Function ValidateDate_80EE(ByVal Target As Range) As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim Date_80EE
    Dim mIntCells As Long
    Dim mIntCtr As Long
    
    
    mIntCells = Sheet20.Range("loanDate.80EE").count
    ValidateDate_80EE = True

    Dim end_80EE
    end_80EE = 0
    Set rangecells = Sheet20.Range("loanDate.80EE").Cells

    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            end_80EE = end_80EE + 1
        End If
    Next

    ReDim Date_80EE(end_80EE)

    For i = 1 To end_80EE
        Date_80EE(i) = rangecells.item(i).value
'        If Not chkCompulsory(Date_80EE(i)) Then
'            'fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is mandatory") '
'            fmsgboxsmall (" Please provide ""Date of sanction of loan"" in schedule 80EE at Sr. No  " & i & "")
'            ValidateDate_80EE = False
'            Exit Function
'        End If

        If Not CheckDateddmmyyyy(Date_80EE(i)) Then
            ValidateDate_80EE = False
'            fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is invalid. Pl enter in dd/mm/yyyy format")
'Ayush_10/02
'              fmsgboxsmall ("Please enter a valid date in dd/mm/yyyy format at Sr. No  " & i & "  in Sheet IT ")
                fmsgboxsmall ("* ""Please enter date in valid format"" at Sr. No  " & i & "")
            Target.value = ""
            Exit Function
        Else
            Date_80EE(i) = Dformat(Date_80EE(i), "")
'            'Ayush_13
''            If Not ChkMinInclusiveDate(DateDep_IT(i), "2023-04-01") Then
'             If Not ChkMinInclusiveDate(Date_24b(i), "2024-04-01") Then
'            'Error Message Changed By Ayush for the SIT-56341
'               ' fmsgboxsmall ("Date of deposit cannot be prior to 01-04-2023 at Sr. No " & i & " in Sheet IT  ")
''               fmsgboxsmall ("""Date of deposit cannot be prior to 01-04-2023""")
'                fmsgboxsmall ("""Date of deposit cannot be prior to 01-04-2024""")
'                Target.value = ""
'                ValidateDate_24b = False
'                Exit Function
'            End If
            
            
            If Not ChkMaxDate_24b(Trim(Date_80EE(i)), "31-03-2025") Then
                fmsgboxsmall ("* Date can not be after 31/03/2025 at Sr. No " & i & ".")
                Target.value = ""
                ValidateDate_80EE = False
                 

          Exit Function
        End If
            
'             If Date_24b(i) > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'
'                fmsgboxsmall ("Date of Deposit at Sr. No  " & i & " in Sheet IT can not be beyond system date")
'                Target.value = ""
'                Target.Select
'                ValidateDate_24b = False
'                Exit Function
'            End If

             
          '-------------------------------------------------

'            If Sheet1.Range("Sheet1.ReturnFileSec") = "139(8A)" Then
''                If ChkMinInclusiveDate(DateDep_IT(i), "2024-01-01") Then
''                    fmsgboxsmall ("Please enter the details of 140B in Sl No 14A of Part B-ATI and not in this schedule, so as to get the correct computation of your tax liability")
' 'Ayush_UR
'            If ChkMinInclusiveDate(DateDep_IT(i), "2025-01-01") Then
'                   fmsgboxsmall ("Date of deposit can't be after 31-12-2024 in schedule IT. Please enter payments made after 31st December 2024 in schedule IT 1.")
'
'
'                    Target.value = ""
'                End If
'            End If
        End If
    Next

End Function



Function ValidateDate_80EEA(ByVal Target As Range) As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim Date_80EEA
    Dim mIntCells As Long
    Dim mIntCtr As Long
    
    
    mIntCells = Sheet20.Range("loanDate.80EEA").count
    ValidateDate_80EEA = True

    Dim end_80EEA
    end_80EEA = 0
    Set rangecells = Sheet20.Range("loanDate.80EEA").Cells

    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            end_80EEA = end_80EEA + 1
        End If
    Next

    ReDim Date_80EEA(end_80EEA)

    For i = 1 To end_80EEA
        Date_80EEA(i) = rangecells.item(i).value
'        If Not chkCompulsory(Date_80EEA(i)) Then
'            'fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is mandatory") '
'            fmsgboxsmall ("* Please provide ""Date of sanction of loan"" in schedule 80EEA at Sr. No  " & i & "")
'            ValidateDate_80EEA = False
'            Exit Function
'        End If

        If Not CheckDateddmmyyyy(Date_80EEA(i)) Then
            ValidateDate_80EEA = False
'            fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is invalid. Pl enter in dd/mm/yyyy format")
'Ayush_10/02
'              fmsgboxsmall ("Please enter a valid date in dd/mm/yyyy format at Sr. No  " & i & "  in Sheet IT ")
                fmsgboxsmall ("* ""Please enter date in valid format"" at Sr. No  " & i & "")
            Target.value = ""
            Exit Function
        Else
            Date_80EEA(i) = Dformat(Date_80EEA(i), "")
'            'Ayush_13
''            If Not ChkMinInclusiveDate(DateDep_IT(i), "2023-04-01") Then
'             If Not ChkMinInclusiveDate(Date_24b(i), "2024-04-01") Then
'            'Error Message Changed By Ayush for the SIT-56341
'               ' fmsgboxsmall ("Date of deposit cannot be prior to 01-04-2023 at Sr. No " & i & " in Sheet IT  ")
''               fmsgboxsmall ("""Date of deposit cannot be prior to 01-04-2023""")
'                fmsgboxsmall ("""Date of deposit cannot be prior to 01-04-2024""")
'                Target.value = ""
'                ValidateDate_24b = False
'                Exit Function
'            End If
            
            
            If Not ChkMaxDate_24b(Trim(Date_80EEA(i)), "31-03-2025") Then
                fmsgboxsmall ("* Date can not be after 31/03/2025 at Sr. No " & i & ".")
                Target.value = ""
                ValidateDate_80EEA = False
                 

          Exit Function
        End If
            
'             If Date_24b(i) > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'
'                fmsgboxsmall ("Date of Deposit at Sr. No  " & i & " in Sheet IT can not be beyond system date")
'                Target.value = ""
'                Target.Select
'                ValidateDate_24b = False
'                Exit Function
'            End If

             
          '-------------------------------------------------

'            If Sheet1.Range("Sheet1.ReturnFileSec") = "139(8A)" Then
''                If ChkMinInclusiveDate(DateDep_IT(i), "2024-01-01") Then
''                    fmsgboxsmall ("Please enter the details of 140B in Sl No 14A of Part B-ATI and not in this schedule, so as to get the correct computation of your tax liability")
' 'Ayush_UR
'            If ChkMinInclusiveDate(DateDep_IT(i), "2025-01-01") Then
'                   fmsgboxsmall ("Date of deposit can't be after 31-12-2024 in schedule IT. Please enter payments made after 31st December 2024 in schedule IT 1.")
'
'
'                    Target.value = ""
'                End If
'            End If
        End If
    Next

End Function






Function ValidateDate_80EEB(ByVal Target As Range) As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim Date_80EEB
    Dim mIntCells As Long
    Dim mIntCtr As Long
    
    
    mIntCells = Sheet20.Range("loanDate.80EEB").count
    ValidateDate_80EEB = True

    Dim end_80EEB
    end_80EEB = 0
    Set rangecells = Sheet20.Range("loanDate.80EEB").Cells

    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            end_80EEB = end_80EEB + 1
        End If
    Next

    ReDim Date_80EEB(end_80EEB)

    For i = 1 To end_80EEB
        Date_80EEB(i) = rangecells.item(i).value
'        If Not chkCompulsory(Date_80EEB(i)) Then
'            'fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is mandatory") '
'            fmsgboxsmall ("* ""Please provide Date of taking loan in schedule 80EEB"" at Sr. No  " & i & "")
'            ValidateDate_80EEB = False
'            Exit Function
'        End If

        If Not CheckDateddmmyyyy(Date_80EEB(i)) Then
            ValidateDate_80EEB = False
'            fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is invalid. Pl enter in dd/mm/yyyy format")
'Ayush_10/02
'              fmsgboxsmall ("Please enter a valid date in dd/mm/yyyy format at Sr. No  " & i & "  in Sheet IT ")
                fmsgboxsmall ("* ""Please enter date in valid format"" at Sr. No  " & i & "")
            Target.value = ""
            Exit Function
        Else
            Date_80EEB(i) = Dformat(Date_80EEB(i), "")
'            'Ayush_13
''            If Not ChkMinInclusiveDate(DateDep_IT(i), "2023-04-01") Then
'             If Not ChkMinInclusiveDate(Date_24b(i), "2024-04-01") Then
'            'Error Message Changed By Ayush for the SIT-56341
'               ' fmsgboxsmall ("Date of deposit cannot be prior to 01-04-2023 at Sr. No " & i & " in Sheet IT  ")
''               fmsgboxsmall ("""Date of deposit cannot be prior to 01-04-2023""")
'                fmsgboxsmall ("""Date of deposit cannot be prior to 01-04-2024""")
'                Target.value = ""
'                ValidateDate_24b = False
'                Exit Function
'            End If
            
            
            If Not ChkMaxDate_24b(Trim(Date_80EEB(i)), "31-03-2025") Then
                fmsgboxsmall ("* Date can not be after 31/03/2025 at Sr. No " & i & ".")
                Target.value = ""
                ValidateDate_80EEB = False
                 

          Exit Function
        End If
            
'             If Date_24b(i) > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'
'                fmsgboxsmall ("Date of Deposit at Sr. No  " & i & " in Sheet IT can not be beyond system date")
'                Target.value = ""
'                Target.Select
'                ValidateDate_24b = False
'                Exit Function
'            End If

             
          '-------------------------------------------------

'            If Sheet1.Range("Sheet1.ReturnFileSec") = "139(8A)" Then
''                If ChkMinInclusiveDate(DateDep_IT(i), "2024-01-01") Then
''                    fmsgboxsmall ("Please enter the details of 140B in Sl No 14A of Part B-ATI and not in this schedule, so as to get the correct computation of your tax liability")
' 'Ayush_UR
'            If ChkMinInclusiveDate(DateDep_IT(i), "2025-01-01") Then
'                   fmsgboxsmall ("Date of deposit can't be after 31-12-2024 in schedule IT. Please enter payments made after 31st December 2024 in schedule IT 1.")
'
'
'                    Target.value = ""
'                End If
'            End If
        End If
    Next

End Function


Function lock80EE_80EEA_chk()


'lock_80EEA_flag, lock_80EE_flag

lock_80EEA_flag = False
lock_80EE_flag = False
            
        If lockcount_80EEA > 0 Then
        '80EE
        'Commented by Ayush_05/05/2025 for 0.7 changes
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range("ResidentialHP.80E").MergeArea.ClearContents
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range("ResidentialHP.80E").MergeArea.Locked = True
'            Sheet20.Range("ResidentialHP.80E").MergeArea.Interior.Color = "&HD8D8D8"

            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("LoanfrmBankOrInstitute.80EE").ClearContents
            Sheet20.Range("LoanfrmBankOrInstitute.80EE").Locked = True
            Sheet20.Range("LoanfrmBankOrInstitute.80EE").Interior.Color = "&HD8D8D8"

'Commented by Ayush_05/05/2025 for 0.7 changes
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range("IFSC.80EE").ClearContents
'            Sheet20.Range("IFSC.80EE").Locked = True
'            Sheet20.Range("IFSC.80EE").Interior.Color = "&HD8D8D8"
'
'
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range("PAN.80EE").ClearContents
'            Sheet20.Range("PAN.80EE").Locked = True
'            Sheet20.Range("PAN.80EE").Interior.Color = "&HD8D8D8"

            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("bankName.80EE").ClearContents
            Sheet20.Range("bankName.80EE").Locked = True
            Sheet20.Range("bankName.80EE").Interior.Color = "&HD8D8D8"

            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanAccNum.80EE").ClearContents
            Sheet20.Range("loanAccNum.80EE").Locked = True
            Sheet20.Range("loanAccNum.80EE").Interior.Color = "&HD8D8D8"

            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanDate.80EE").ClearContents
            Sheet20.Range("loanDate.80EE").Locked = True
            Sheet20.Range("loanDate.80EE").Interior.Color = "&HD8D8D8"

            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanAmt.80EE").ClearContents
            Sheet20.Range("loanAmt.80EE").Locked = True
            Sheet20.Range("loanAmt.80EE").Interior.Color = "&HD8D8D8"

            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanOutstanding.80EE").ClearContents
            Sheet20.Range("loanOutstanding.80EE").Locked = True
            Sheet20.Range("loanOutstanding.80EE").Interior.Color = "&HD8D8D8"

            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("Intrst.80EE").ClearContents
            Sheet20.Range("Intrst.80EE").Locked = True
            Sheet20.Range("Intrst.80EE").Interior.Color = "&HD8D8D8"
            
            lock_80EE_flag = True

          ElseIf lockcount_80EE > 0 Then
            '80EEA
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("Stampduty.80EEA").MergeArea.ClearContents
            Sheet20.Range("Stampduty.80EEA").MergeArea.Locked = True
            Sheet20.Range("Stampduty.80EEA").MergeArea.Interior.Color = "&HD8D8D8"
            
            
            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("LoanfrmBankOrInstitute.80EEA").ClearContents
            Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Locked = True
            Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Interior.Color = "&HD8D8D8"
            
 'Commented by Ayush_05/05/2025 for 0.7 changes
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range("IFSC.80EEA").ClearContents
'            Sheet20.Range("IFSC.80EEA").Locked = True
'            Sheet20.Range("IFSC.80EEA").Interior.Color = "&HD8D8D8"
            
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range("PAN.80EEA").ClearContents
'            Sheet20.Range("PAN.80EEA").Locked = True
'            Sheet20.Range("PAN.80EEA").Interior.Color = "&HD8D8D8"
            
            
           Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("bankName.80EEA").ClearContents
            Sheet20.Range("bankName.80EEA").Locked = True
            Sheet20.Range("bankName.80EEA").Interior.Color = "&HD8D8D8"
            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanAccNum.80EEA").ClearContents
            Sheet20.Range("loanAccNum.80EEA").Locked = True
            Sheet20.Range("loanAccNum.80EEA").Interior.Color = "&HD8D8D8"
            
           Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanDate.80EEA").ClearContents
            Sheet20.Range("loanDate.80EEA").Locked = True
            Sheet20.Range("loanDate.80EEA").Interior.Color = "&HD8D8D8"
            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanAmt.80EEA").ClearContents
            Sheet20.Range("loanAmt.80EEA").Locked = True
            Sheet20.Range("loanAmt.80EEA").Interior.Color = "&HD8D8D8"
            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanOutstanding.80EEA").ClearContents
            Sheet20.Range("loanOutstanding.80EEA").Locked = True
            Sheet20.Range("loanOutstanding.80EEA").Interior.Color = "&HD8D8D8"
            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("Intrst.80EEA").ClearContents
            Sheet20.Range("Intrst.80EEA").Locked = True
            Sheet20.Range("Intrst.80EEA").Interior.Color = "&HD8D8D8"
            
             lock_80EEA_flag = True
            
       Else
       
       'Commented by Ayush_05/05/2025 for 0.7 changes
       
'        Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range("ResidentialHP.80E").MergeArea.ClearContents
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range("ResidentialHP.80E").MergeArea.Locked = False
'            Sheet20.Range("ResidentialHP.80E").MergeArea.Interior.Color = "&HCCFFCC"

            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("LoanfrmBankOrInstitute.80EE").ClearContents
            Sheet20.Range("LoanfrmBankOrInstitute.80EE").Locked = False
            Sheet20.Range("LoanfrmBankOrInstitute.80EE").Interior.Color = "&HCCFFCC"

            'Commented by Ayush_05/05/2025 for 0.7 changes
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range("IFSC.80EE").ClearContents
'            Sheet20.Range("IFSC.80EE").Locked = False
'            Sheet20.Range("IFSC.80EE").Interior.Color = "&HCCFFCC"

            
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range("PAN.80EE").ClearContents
'            Sheet20.Range("PAN.80EE").Locked = False
'            Sheet20.Range("PAN.80EE").Interior.Color = "&HCCFFCC"

            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("bankName.80EE").ClearContents
            Sheet20.Range("bankName.80EE").Locked = False
            Sheet20.Range("bankName.80EE").Interior.Color = "&HCCFFCC"

            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanAccNum.80EE").ClearContents
            Sheet20.Range("loanAccNum.80EE").Locked = False
            Sheet20.Range("loanAccNum.80EE").Interior.Color = "&HCCFFCC"

            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanDate.80EE").ClearContents
            Sheet20.Range("loanDate.80EE").Locked = False
            Sheet20.Range("loanDate.80EE").Interior.Color = "&HCCFFCC"

            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanAmt.80EE").ClearContents
            Sheet20.Range("loanAmt.80EE").Locked = False
            Sheet20.Range("loanAmt.80EE").Interior.Color = "&HCCFFCC"

            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanOutstanding.80EE").ClearContents
            Sheet20.Range("loanOutstanding.80EE").Locked = False
            Sheet20.Range("loanOutstanding.80EE").Interior.Color = "&HCCFFCC"

            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("Intrst.80EE").ClearContents
            Sheet20.Range("Intrst.80EE").Locked = False
            Sheet20.Range("Intrst.80EE").Interior.Color = "&HCCFFCC"

          
           Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("Stampduty.80EEA").MergeArea.ClearContents
            Sheet20.Range("Stampduty.80EEA").MergeArea.Locked = False
            Sheet20.Range("Stampduty.80EEA").MergeArea.Interior.Color = "&HCCFFCC"
            
            
            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("LoanfrmBankOrInstitute.80EEA").ClearContents
            Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Locked = False
            Sheet20.Range("LoanfrmBankOrInstitute.80EEA").Interior.Color = "&HCCFFCC"
            
            
            'Commented by Ayush_05/05/2025 for 0.7 changes
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range("IFSC.80EEA").ClearContents
'            Sheet20.Range("IFSC.80EEA").Locked = False
'            Sheet20.Range("IFSC.80EEA").Interior.Color = "&HCCFFCC"
'
'            Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range("PAN.80EEA").ClearContents
'            Sheet20.Range("PAN.80EEA").Locked = False
'            Sheet20.Range("PAN.80EEA").Interior.Color = "&HCCFFCC"
            
            
           Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("bankName.80EEA").ClearContents
            Sheet20.Range("bankName.80EEA").Locked = False
            Sheet20.Range("bankName.80EEA").Interior.Color = "&HCCFFCC"
            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanAccNum.80EEA").ClearContents
            Sheet20.Range("loanAccNum.80EEA").Locked = False
            Sheet20.Range("loanAccNum.80EEA").Interior.Color = "&HCCFFCC"
            
           Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanDate.80EEA").ClearContents
            Sheet20.Range("loanDate.80EEA").Locked = False
            Sheet20.Range("loanDate.80EEA").Interior.Color = "&HCCFFCC"
            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanAmt.80EEA").ClearContents
            Sheet20.Range("loanAmt.80EEA").Locked = False
            Sheet20.Range("loanAmt.80EEA").Interior.Color = "&HCCFFCC"
            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("loanOutstanding.80EEA").ClearContents
            Sheet20.Range("loanOutstanding.80EEA").Locked = False
            Sheet20.Range("loanOutstanding.80EEA").Interior.Color = "&HCCFFCC"
            
            Sheet20.Unprotect Password:=getmsgstate
            Sheet20.Range("Intrst.80EEA").ClearContents
            Sheet20.Range("Intrst.80EEA").Locked = False
            Sheet20.Range("Intrst.80EEA").Interior.Color = "&HCCFFCC"
            
            lock_80EEA_flag = False
            lock_80EE_flag = False
            
           End If

End Function



Function lock80EE_80EEAcount()

Dim ResidentialHP_80E, Stampduty_80EEA As Variant




lockcount_80EE = 0
lockcount_80EEA = 0


'------------------------------------------------------------------------
'80EE

    Dim mIntCells As Long
    Dim rangecells, rangecells1, rangecells2, rangecells3, rangecells4, rangecells5, rangecells6, rangecells7, rangecells8, rangecells9 As Range
     Dim rangecells11, rangecells22, rangecells33, rangecells44, rangecells55, rangecells66, rangecells77, rangecells88, rangecells99 As Range
    'Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount, ccount1, ccount2, ccount3, ccount4, ccount5, ccount6, ccount7, ccount8 As Long
     Dim ccount9, ccount11, ccount22, ccount33, ccount44, ccount55, ccount66, ccount77, ccount88 As Long
    
    
    mIntCells = Range("LoanfrmBankOrInstitute.80EE").count
    Set rangecells9 = Range("LoanfrmBankOrInstitute.80EE").Cells
    'Commented by Ayush_05/05/2025 for 0.7 changes
'    Set rangecells11 = Range("IFSC.80EE").Cells
'    Set rangecells22 = Range("PAN.80EE").Cells
    Set rangecells33 = Range("bankName.80EE").Cells
    Set rangecells44 = Range("loanAccNum.80EE").Cells
    Set rangecells55 = Range("loanDate.80EE").Cells
    Set rangecells66 = Range("loanAmt.80EE").Cells
    Set rangecells77 = Range("loanOutstanding.80EE").Cells
    Set rangecells88 = Range("Intrst.80EE").Cells
    
    For mIntCtr = 1 To mIntCells
            If Not (rangecells9.item(mIntCtr).value = "" Or rangecells9.item(mIntCtr).value = "(Select)") Then
                ccount9 = ccount9 + 1
            End If
            'Commented by Ayush_05/05/2025 for 0.7 changes
'            If Not (rangecells11.item(mIntCtr).value = "") Then
'                ccount11 = ccount11 + 1
'            End If
'            If Not (rangecells22.item(mIntCtr).value = "") Then
'                ccount22 = ccount22 + 1
'            End If
            If Not (rangecells33.item(mIntCtr).value = "") Then
                ccount33 = ccount33 + 1
            End If
            If Not (rangecells44.item(mIntCtr).value = "") Then
                ccount44 = ccount44 + 1
            End If
            If Not (rangecells55.item(mIntCtr).value = "") Then
                ccount55 = ccount55 + 1
            End If
            If Not (rangecells66.item(mIntCtr).value = "") Then
                ccount66 = ccount66 + 1
            End If
            If Not (rangecells77.item(mIntCtr).value = "") Then
                ccount77 = ccount77 + 1
            End If
            If Not (rangecells88.item(mIntCtr).value = "") Then
                ccount88 = ccount88 + 1
            End If
            
            
            
    Next
    
    'Commented by Ayush_05/05/2025 for 0.7 changes
    
'    If Sheet20.Range("ResidentialHP.80E").value > 0 Then
'ResidentialHP_80E = 1
'Else
'ResidentialHP_80E = 0
'End If

'Commented by Ayush_05/05/2025 for 0.7 changes
'    lockcount_80EE = Application.WorksheetFunction.Max(0, ResidentialHP_80E, ccount9, ccount11, ccount22, ccount33, ccount44, ccount55, ccount66, ccount77, ccount88)
    lockcount_80EE = Application.WorksheetFunction.Max(0, ccount9, ccount11, ccount22, ccount33, ccount44, ccount55, ccount66, ccount77, ccount88)






mIntCells = Range("LoanfrmBankOrInstitute.80EEA").count
    Set rangecells = Range("LoanfrmBankOrInstitute.80EEA").Cells
    'Commented by Ayush_05/05/2025 for 0.7 changes
   ' Set rangecells1 = Range("IFSC.80EEA").Cells
   ' Set rangecells2 = Range("PAN.80EEA").Cells
    Set rangecells3 = Range("bankName.80EEA").Cells
    Set rangecells4 = Range("loanAccNum.80EEA").Cells
    Set rangecells5 = Range("loanDate.80EEA").Cells
    Set rangecells6 = Range("loanAmt.80EEA").Cells
    Set rangecells7 = Range("loanOutstanding.80EEA").Cells
    Set rangecells8 = Range("Intrst.80EEA").Cells
    
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
            'Commented by Ayush_05/05/2025 for 0.7 changes
'            If Not (rangecells1.item(mIntCtr).value = "") Then
'                ccount1 = ccount1 + 1
'            End If
'            If Not (rangecells2.item(mIntCtr).value = "") Then
'                ccount2 = ccount2 + 1
'            End If
            If Not (rangecells3.item(mIntCtr).value = "") Then
                ccount3 = ccount3 + 1
            End If
            If Not (rangecells4.item(mIntCtr).value = "") Then
                ccount4 = ccount4 + 1
            End If
            If Not (rangecells5.item(mIntCtr).value = "") Then
                ccount5 = ccount5 + 1
            End If
            If Not (rangecells6.item(mIntCtr).value = "") Then
                ccount6 = ccount6 + 1
            End If
            If Not (rangecells7.item(mIntCtr).value = "") Then
                ccount7 = ccount7 + 1
            End If
            If Not (rangecells8.item(mIntCtr).value = "") Then
                ccount8 = ccount8 + 1
            End If
            
            
            
    Next
    
If Sheet20.Range("Stampduty.80EEA").value > 0 Then
Stampduty_80EEA = 1
Else
Stampduty_80EEA = 0
End If

    lockcount_80EEA = WorksheetFunction.Max(0, Stampduty_80EEA, ccount, ccount1, ccount2, ccount3, ccount4, ccount5, ccount6, ccount7, ccount8)


'-------------------------------------------------------------------------
End Function

