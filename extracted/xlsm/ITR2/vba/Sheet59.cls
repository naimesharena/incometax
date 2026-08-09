Attribute VB_Name = "Sheet59"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Activate()
On Error Resume Next
Application.EnableEvents = False
ActiveWindow.DisplayHeadings = False

Sheet59.Unprotect Password:=EfilingCommon.getmsgstate
fmsgbox "Please fill date wise donation details. If Donation is made in cash and in other mode on same day, please fill details in separate rows"
Sheet59.Protect Password:=EfilingCommon.getmsgstate
Application.EnableEvents = True
'ActiveWindow.DisplayHeadings = True

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

Application.EnableEvents = False


Dim IfsCode As Range
Dim sPassword

Dim IffCode As Range
Dim TargtVal, targetadd As Variant

TargtVal = Target.value
targetadd = Target.Address
targetadd = Replace(targetadd, "$", "")


sPassword = EfilingCommon.getmsgstate
Sheet59.Unprotect Password:=sPassword

'Added by Shrutika- 30/03/2026 V0.2 AY-26-27

If Not Application.Intersect(Sheet59.Range("politicalPan_80GGC"), Sheet59.Range(Target.Address)) Is Nothing Then

       Target.value = UCase(Target.value)

End If
'----------------------
If Not Application.Intersect(Sheet59.Range("DateofDonation_80GGC"), Sheet59.Range(Target.Address)) Is Nothing Then

    If Target.value <> "00/00/0000" And Target.value <> "" Then
        If Not FormatNCheckDate(Trim(Target.value)) Then
            fmsgbox "* Date is invalid in Schedule 80GGC. Please enter a valid Date in dd/mm/yyyy format." ', vbOKOnly, mdCommon.vbMessageCaption
            Target.ClearContents
        End If
     End If
     
 '--------------------------------

 
     
'changed by Chetan C M on 21/01/2026 for AY 2026-27
    'start--
''Shrutika-------------------AY_2025-26 change
'    If Not ChkMinInclusiveDate(Dformat(Target.value, ""), "2024-04-01") Or _
'       Not ChkMaxInclusiveDate(Dformat(Target.value, ""), "2025-03-31") Then
'       MsgBox "* ""Date of contribution should be within the 1st April 2024 to 31st March 2025.""", , "Error(s)" ',vbOKOnly, mdCommon.vbMessageCaption
''-----------------------------------------
    If Not ChkMinInclusiveDate(Dformat(Target.value, ""), "2025-04-01") Or _
       Not ChkMaxInclusiveDate(Dformat(Target.value, ""), "2026-03-31") Then
       MsgBox "* ""Date of contribution should be within the 1st April 2025 to 31st March 2026.""", , "Error(s)"
       Target.ClearContents
    End If
    '--end
End If


'--------------------
            
Set IfsCode = Range("IFSC_80GGC")
If Not Application.Intersect(IfsCode, Range(Target.Address)) Is Nothing Then
        'Call CheckIFSC_80GGC(TargtVal, targetadd)
        
        Target.value = UCase(Target.value)
        If Target.value <> "" Then
            If Not checkListIFSC(UCase(Target.value)) Then

                allowifscnnn = False
                tempvar = Sheet1.Range("sheet1.ResidentialStatus1").value
                If Mid(tempvar, 1, 3) = "NRI" Then
                   allowifscnnn = True
                End If

                    fmsgboxStatus "Please enter valid IFS Code."
                    Sheet59.Range(Target.Address).value = ""
                    Sheet59.Range(Target.Address).Select
                    sourceAddress = rangecell.Address
                    
             ElseIf Not ValidateIFSCCode(Target.value) Then
             
                fmsgboxStatus "Invalid IFS Code: Please enter valid IFS Code. Refer to your bank for valid codes. (1st 4 Alpha, followed by 0 and 6 digits) "
                Range(Target.Address) = ""
                Range(Target.Address).Select
                
            End If
        End If
    
        'GoTo Endline
End If
       
Application.EnableEvents = False
sPassword = EfilingCommon.getmsgstate
Sheet59.Unprotect Password:=sPassword

If Not Application.Intersect(Sheet59.Range("Donationincash_80GGC"), Sheet59.Range(Target.Address)) Is Nothing Then

    If Target.value > 0 And Not IsEmpty(Target.value) Then
        
        
        
        'DONATION IN OTHER MODE
        Sheet59.Range(Replace(Target.Address, "F", "G")).ClearContents
        Sheet59.Range(Replace(Target.Address, "F", "G")).Locked = True
        Sheet59.Range(Replace(Target.Address, "F", "G")).Interior.Color = "&HD8D8D8"
        
        
'        'NATURE OF TRANSACTION
'        Sheet59.Range(Replace(Target.Address, "F", "J")).ClearContents
'        Sheet59.Range(Replace(Target.Address, "F", "J")).Locked = True
'        Sheet59.Range(Replace(Target.Address, "F", "J")).Interior.Color = "&HD8D8D8"
        
        
'        'REF/CHEQUE NUMBER
'        Sheet59.Range(Replace(Target.Address, "F", "J")).ClearContents
'        Sheet59.Range(Replace(Target.Address, "F", "J")).Locked = True
'        Sheet59.Range(Replace(Target.Address, "F", "J")).Interior.Color = "&HD8D8D8"
'
'        'IFSC
'        Sheet59.Range(Replace(Target.Address, "F", "K")).ClearContents
'        Sheet59.Range(Replace(Target.Address, "F", "K")).Locked = True
'        Sheet59.Range(Replace(Target.Address, "F", "K")).Interior.Color = "&HD8D8D8"
        '------------'Added by Shrutika- 27/03/2026 V0.2 AY-26-27

            'REF/CHEQUE NUMBER
        Sheet59.Range(Replace(Target.Address, "F", "L")).ClearContents
        Sheet59.Range(Replace(Target.Address, "F", "L")).Locked = True
        Sheet59.Range(Replace(Target.Address, "F", "L")).Interior.Color = "&HD8D8D8"
        
        'IFSC
        Sheet59.Range(Replace(Target.Address, "F", "M")).ClearContents
        Sheet59.Range(Replace(Target.Address, "F", "M")).Locked = True
        Sheet59.Range(Replace(Target.Address, "F", "M")).Interior.Color = "&HD8D8D8"
        '----------------------------'Added by Shrutika- 27/03/2026 V0.2 AY-26-27
        
'        'BANK NAME OF DONoR
'        Sheet59.Range(Replace(Target.Address, "F", "M")).ClearContents
'        Sheet59.Range(Replace(Target.Address, "F", "M")).Locked = True
'        Sheet59.Range(Replace(Target.Address, "F", "M")).Interior.Color = "&HD8D8D8"
'
'
'        'BANK ACCOUNT OF Donor
'        Sheet59.Range(Replace(Target.Address, "F", "N")).ClearContents
'        Sheet59.Range(Replace(Target.Address, "F", "N")).Locked = True
'        Sheet59.Range(Replace(Target.Address, "F", "N")).Interior.Color = "&HD8D8D8"
        
    Else
    'DONATION IN OTHER MODE
        Sheet59.Range(Replace(Target.Address, "F", "G")).ClearContents
        Sheet59.Range(Replace(Target.Address, "F", "G")).Locked = False
        Sheet59.Range(Replace(Target.Address, "F", "G")).Interior.Color = "&HCCFFCC"
        
        
'        'NATURE OF TRANSACTION
'        Sheet59.Range(Replace(Target.Address, "F", "J")).Value = "(Select)"
'        Sheet59.Range(Replace(Target.Address, "F", "J")).Locked = False
'        Sheet59.Range(Replace(Target.Address, "F", "J")).Interior.Color = "&HCCFFCC"
        
        
'        'REF/CHEQUE NUMBER
'        Sheet59.Range(Replace(Target.Address, "F", "J")).ClearContents
'        Sheet59.Range(Replace(Target.Address, "F", "J")).Locked = False
'        Sheet59.Range(Replace(Target.Address, "F", "J")).Interior.Color = "&HCCFFCC"
'
'        'IFSC
'        Sheet59.Range(Replace(Target.Address, "F", "K")).ClearContents
'        Sheet59.Range(Replace(Target.Address, "F", "K")).Locked = False
'        Sheet59.Range(Replace(Target.Address, "F", "K")).Interior.Color = "&HCCFFCC"
        
        '--------'Added by Shrutika- 27/03/2026 V0.2 AY-26-27

         'REF/CHEQUE NUMBER
        Sheet59.Range(Replace(Target.Address, "F", "L")).ClearContents
        Sheet59.Range(Replace(Target.Address, "F", "L")).Locked = False
        Sheet59.Range(Replace(Target.Address, "F", "L")).Interior.Color = "&HCCFFCC"
        
        'IFSC
        Sheet59.Range(Replace(Target.Address, "F", "M")).ClearContents
        Sheet59.Range(Replace(Target.Address, "F", "M")).Locked = False
        Sheet59.Range(Replace(Target.Address, "F", "M")).Interior.Color = "&HCCFFCC"
        '-----------------------------------------------------------------------------------
        
'        'BANK NAME OF DONoR
'        Sheet59.Range(Replace(Target.Address, "F", "M")).ClearContents
'        Sheet59.Range(Replace(Target.Address, "F", "M")).Locked = False
'        Sheet59.Range(Replace(Target.Address, "F", "M")).Interior.Color = "&HCCFFCC"
'
'
'        'BANK ACCOUNT OF Donor
'        Sheet59.Range(Replace(Target.Address, "F", "N")).ClearContents
'        Sheet59.Range(Replace(Target.Address, "F", "N")).Locked = False
'        Sheet59.Range(Replace(Target.Address, "F", "N")).Interior.Color = "&HCCFFCC"
        
        
    End If
End If


'DONATION IN OTHER MODE
If Not Application.Intersect(Sheet59.Range("Donationinothermode_80GGC"), Sheet59.Range(Target.Address)) Is Nothing Then
     If (Target.value > 0 And Not IsEmpty(Target.value)) Or (Len(Range(Replace(Target.Address, "G", "J")).value) > 0) Or Len(Range(Replace(Target.Address, "G", "K")).value) > 0 Then
    Sheet59.Unprotect Password:=sPassword
    'DONATION IN CASH
        Sheet59.Range(Replace(Target.Address, "G", "F")).ClearContents
        Sheet59.Range(Replace(Target.Address, "G", "F")).Locked = True
        Sheet59.Range(Replace(Target.Address, "G", "F")).Interior.Color = "&HD8D8D8"
    Else
        Sheet59.Range(Replace(Target.Address, "G", "F")).Locked = False
        Sheet59.Range(Replace(Target.Address, "G", "F")).Interior.Color = "&HCCFFCC"
    End If
              
End If



''Nature TRANSACTION
'If Not Application.Intersect(Sheet59.Range("NatureofTransaction_80GGC"), Sheet59.Range(Target.Address)) Is Nothing Then
'
'    If ((Not IsEmpty(Target.Value) And UCase(Target.Value) <> UCase("(Select)"))) Or (Range(Replace(Target.Address, "J", "G")).Value > 0 And Not IsEmpty(Range(Replace(Target.Address, "J", "G")).Value)) Or Len(Range(Replace(Target.Address, "J", "K")).Value) > 0 Or Len(Range(Replace(Target.Address, "J", "L")).Value) > 0 Then
'
'
'        'DONATION IN CASH
'        Sheet59.Range(Replace(Target.Address, "J", "F")).ClearContents
'        Sheet59.Range(Replace(Target.Address, "J", "F")).Locked = True
'        Sheet59.Range(Replace(Target.Address, "J", "F")).Interior.Color = "&HD8D8D8"
'    Else
'
'       'DONATION IN CASH
'        Sheet59.Range(Replace(Target.Address, "J", "F")).ClearContents
'        Sheet59.Range(Replace(Target.Address, "J", "F")).Locked = False
'        Sheet59.Range(Replace(Target.Address, "J", "F")).Interior.Color = "&HCCFFCC"
'
'    End If
'
'End If


''Transaction number
If Not Application.Intersect(Sheet59.Range("Chequeno_80GGC"), Sheet59.Range(Target.Address)) Is Nothing Then
      Sheet59.Unprotect Password:=sPassword
    If Len(Target.value) > 1 Or (Range(Replace(Target.Address, "L", "G")).value > 0 And Not IsEmpty(Range(Replace(Target.Address, "J", "G")).value)) Or Len(Range(Replace(Target.Address, "J", "K")).value) > 0 Then


        'DONATION IN CASH
        Sheet59.Range(Replace(Target.Address, "L", "F")).ClearContents
        Sheet59.Range(Replace(Target.Address, "L", "F")).Locked = True
        Sheet59.Range(Replace(Target.Address, "L", "F")).Interior.Color = "&HD8D8D8"
    Else

       'DONATION IN CASH
        Sheet59.Range(Replace(Target.Address, "L", "F")).Locked = False
        Sheet59.Range(Replace(Target.Address, "L", "F")).Interior.Color = "&HCCFFCC"

    End If

End If



'IFSC
If Not Application.Intersect(Sheet59.Range("IFSC_80GGC"), Sheet59.Range(Target.Address)) Is Nothing Then
     Sheet59.Unprotect Password:=sPassword
     If Len(Target.value) > 1 Or (Range(Replace(Target.Address, "M", "G")).value > 0 And Not IsEmpty(Range(Replace(Target.Address, "K", "G")).value)) Or Len(Range(Replace(Target.Address, "K", "J")).value) > 0 Then

        'DONATION IN CASH
        Sheet59.Range(Replace(Target.Address, "M", "F")).ClearContents
        Sheet59.Range(Replace(Target.Address, "M", "F")).Locked = True
        Sheet59.Range(Replace(Target.Address, "M", "F")).Interior.Color = "&HD8D8D8"
    Else

       'DONATION IN CASH
        Sheet59.Range(Replace(Target.Address, "M", "F")).Locked = False
        Sheet59.Range(Replace(Target.Address, "M", "F")).Interior.Color = "&HCCFFCC"

    End If

End If


          
''NAME OF BANK
' If Not Application.Intersect(Sheet59.Range("NameofDonor_80GGC"), Sheet59.Range(Target.Address)) Is Nothing Then
'
'    If Len(Target.Value) > 1 Or (Range(Replace(Target.Address, "M", "G")).Value > 0 And Not IsEmpty(Range(Replace(Target.Address, "M", "G")).Value)) Or Len(Range(Replace(Target.Address, "M", "J")).Value) > 0 Or Len(Range(Replace(Target.Address, "M", "K")).Value) > 0 Or Len(Range(Replace(Target.Address, "M", "L")).Value) > 0 Or Len(Range(Replace(Target.Address, "M", "N")).Value) > 0 Then
'
'
'        'DONATION IN CASH
'        Sheet59.Range(Replace(Target.Address, "M", "F")).ClearContents
'        Sheet59.Range(Replace(Target.Address, "M", "F")).Locked = True
'        Sheet59.Range(Replace(Target.Address, "M", "F")).Interior.Color = "&HD8D8D8"
'    Else
'
'       'DONATION IN CASH
'        Sheet59.Range(Replace(Target.Address, "M", "F")).ClearContents
'        Sheet59.Range(Replace(Target.Address, "M", "F")).Locked = False
'        Sheet59.Range(Replace(Target.Address, "M", "F")).Interior.Color = "&HCCFFCC"
'
'    End If
'
'End If
            



endline:

Application.EnableEvents = True


sPassword = EfilingCommon.getmsgstate
Sheet59.Protect Password:=sPassword
End Sub



