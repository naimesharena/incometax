Attribute VB_Name = "Sheet11"
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

Sheet11.Unprotect Password:=getmsgstate
fmsgboxStatus "Please fill date wise Contribution details. If Contribution is made in cash and in other mode on same day, please fill details in separate rows"
Sheet11.Protect Password:=getmsgstate
Application.EnableEvents = True
ActiveWindow.DisplayHeadings = True

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
Dim validateType As String
Dim TargtVal, targetadd As Variant
Dim IffCode As Range
Dim IfsCode As Range


TargtVal = Target.value
targetadd = Target.Address
targetadd = Replace(targetadd, "$", "")

 


 
 
Sheet11.Unprotect Password:=getmsgstate

If Not Application.Intersect(Sheet11.Range("SCH_80GGC_DonationDate"), Sheet11.Range(Target.Address)) Is Nothing Then

    If Target.value <> "00/00/0000" And Target.value <> "" Then
        If Not FormatNCheckDate(Trim(Target.value)) Then
        'Ayush_10/02
'            fmsgbox "* Date of Contribution is invalid in Sheet 80GGC. Please enter a valid date in dd/mm/yyyy format."        ', vbOKOnly, mdCommon.vbMessageCaption
            fmsgbox "* ""Date of Contribution - Please enter valid date in valid format."""        ', vbOKOnly, mdCommon.vbMessageCaption
            Target.ClearContents
        End If
    End If
    
    '===================
    'Malli
    Dim DateDep_80ggc As Variant
    DateDep_80ggc = Dformat(Trim(Target.value), "")
    Dim PY_date_80GGC As Date
    PY_date_80GGC = CDate(Sheet8.Range("Date_80GGC_PY").value)
'Ayush_25-26
'    If Not ChkMinInclusiveDate(DateDep_80ggc, "2023-04-01") Then
'    fmsgbox "*""Date of Contribution cannot be before 01/04/2023."""
'     If Not ChkMinInclusiveDate(DateDep_80ggc, "2024-04-01") Then
    If Not ChkMinInclusiveDate(DateDep_80ggc, Dformat(Trim(Sheet8.Range("Date_80GGC_PY").value), "yyyy-mm-dd")) Then
'    fmsgbox "*""Date of Contribution cannot be before 01/04/2025."""
     fmsgbox "*""Date of Contribution cannot be before " & Dformat4(PY_date_80GGC, "dd/mm/yyyy") & "."""
                'fmsgboxsmall ("Malli")
                Target.ClearContents
                ''Exit Function
            End If
    
'    If Not ChkMaxDate_80GGc(Trim(Target.value), "31-03-2024") Then
'    fmsgbox "*""Date of Contribution cannot be more than 31/03/2024."""
'Ayush_26-27_80GGC_C5
'    If Not ChkMaxDate_80GGc(Trim(Target.value), "31-03-2025") Then
Dim cutoff_80GGC As Date
cutoff_80GGC = CDate(Sheet8.Range("AY_80GGC").value)

    If Not ChkMaxDate_80GGc(Trim(Target.value), Sheet8.Range("DOB_1").value) Then
    fmsgbox "*""Date of Contribution cannot be more than " & Dformat5(cutoff_80GGC, "dd/mm/yyyy") & "."""
    'fmsgboxsmall (" Date of Contribution cannot be more than 01/04/2025")
    'fmsgboxsmall (" Date of Contribution cannot be more than 01/04/2024")
                Target.ClearContents
                End If
    '=================
    
    
End If

'-----------------------------
    
validateType = ""
Dim sourceSheet As Worksheet
Set sourceSheet = ThisWorkbook.Sheets("80GGC")

Application.ScreenUpdating = False
Application.EnableEvents = False

Sheet11.Unprotect Password:=getmsgstate
If Not Intersect(Range("SCH_80GGC_DonationAmtCash"), Range(Target.Address)) Is Nothing Then
Sheet11.Unprotect Password:=getmsgstate
        'If Sheet11.Range(Target.Address).value <> "" Or Sheet11.Range(Target.Address).value > 0 Then

           If Sheet11.Range(Target.Address).value > 0 Then
           Sheet11.Range(Target.Address).Interior.Color = "&HCCFFCC"
           
            
           
            Sheet11.Range(Replace(Target.Address, "F", "G")).ClearContents
            Sheet11.Range(Replace(Target.Address, "F", "G")).Locked = True
            Sheet11.Range(Replace(Target.Address, "F", "G")).Interior.Color = "&HD8D8D8"
'            Sheet11.Range(Replace(Target.Address, "F", "G")).value = 0
            
            
'            Sheet11.Range(Replace(Target.Address, "F", "J")).ClearContents
'            Sheet11.Range(Replace(Target.Address, "F", "J")).Locked = True
'            Sheet11.Range(Replace(Target.Address, "F", "J")).Interior.Color = "&HD8D8D8"
'
'
'            Sheet11.Range(Replace(Target.Address, "F", "K")).ClearContents
'            Sheet11.Range(Replace(Target.Address, "F", "K")).Locked = True
'            Sheet11.Range(Replace(Target.Address, "F", "K")).Interior.Color = "&HD8D8D8"
            
             
            Sheet11.Range(Replace(Target.Address, "F", "L")).ClearContents
            Sheet11.Range(Replace(Target.Address, "F", "L")).Locked = True
            Sheet11.Range(Replace(Target.Address, "F", "L")).Interior.Color = "&HD8D8D8"
            
            
            Sheet11.Range(Replace(Target.Address, "F", "M")).ClearContents
            Sheet11.Range(Replace(Target.Address, "F", "M")).Locked = True
            Sheet11.Range(Replace(Target.Address, "F", "M")).Interior.Color = "&HD8D8D8"
'
'
'            Sheet11.Range(Replace(Target.Address, "F", "N")).ClearContents
'            Sheet11.Range(Replace(Target.Address, "F", "N")).Locked = True
'            Sheet11.Range(Replace(Target.Address, "F", "N")).Interior.Color = "&HD8D8D8"
'
             
            
            
        Else
                Sheet11.Unprotect Password:=getmsgstate
                Sheet11.Range(Replace(Target.Address, "F", "G")).ClearContents
                Sheet11.Range(Replace(Target.Address, "F", "G")).Locked = False
                Sheet11.Range(Replace(Target.Address, "F", "G")).Interior.Color = "&HCCFFCC"
                
'                Sheet11.Range(Replace(Target.Address, "F", "J")).Locked = False
'                Sheet11.Range(Replace(Target.Address, "F", "J")).Interior.Color = "&HCCFFCC"
'
'                Sheet11.Range(Replace(Target.Address, "F", "K")).Locked = False
'                Sheet11.Range(Replace(Target.Address, "F", "K")).Interior.Color = "&HCCFFCC"
                
                Sheet11.Range(Replace(Target.Address, "F", "L")).Locked = False
                Sheet11.Range(Replace(Target.Address, "F", "L")).Interior.Color = "&HCCFFCC"
'
                Sheet11.Range(Replace(Target.Address, "F", "M")).Locked = False
                Sheet11.Range(Replace(Target.Address, "F", "M")).Interior.Color = "&HCCFFCC"
'
'                Sheet11.Range(Replace(Target.Address, "F", "N")).Locked = False
'                Sheet11.Range(Replace(Target.Address, "F", "N")).Interior.Color = "&HCCFFCC"
                
                
                
                 
Sheet11.Protect Password:=getmsgstate
    End If
End If

'DONATION IN OTHER MODE
        
If Not Application.Intersect(Sheet11.Range("SCH_80GGC_DonationAmtOtherMode"), Sheet11.Range(Target.Address)) Is Nothing Then
'If (Target.value > 0 And Not IsEmpty(Target.value)) Or ((Not IsEmpty(Range(Replace(Target.Address, "G", "J")).value) And UCase(Range(Replace(Target.Address, "G", "J")).value <> UCase("(Select)")))) Or Len(Range(Replace(Target.Address, "G", "K")).value) > 0 Then


If (Target.value > 0 And Not IsEmpty(Target.value)) Or (Not IsEmpty(Range(Replace(Target.Address, "G", "j")).value) And Len(Range(Replace(Target.Address, "G", "J")).value) > 0) Or (Not IsEmpty(Range(Replace(Target.Address, "G", "k")).value) And Len(Range(Replace(Target.Address, "G", "k")).value) > 0) Then



'    If (Target.value > 0 And Not IsEmpty(Target.value)) Or Len(Range(Replace(Target.Address, "G", "J")).value) > 0 Or Len(Range(Replace(Target.Address, "G", "K")).value) > 0 Or Len(Range(Replace(Target.Address, "G", "L")).value) > 0 Or Len(Range(Replace(Target.Address, "G", "M")).value) > 0 Or Len(Range(Replace(Target.Address, "G", "N")).value) > 0 Then
    Sheet11.Unprotect Password:=getmsgstate
    'DONATION IN CASH
        Sheet11.Range(Replace(Target.Address, "G", "F")).ClearContents
        Sheet11.Range(Replace(Target.Address, "G", "F")).Locked = True
        Sheet11.Range(Replace(Target.Address, "G", "F")).Interior.Color = "&HD8D8D8"
'        Sheet11.Range(Replace(Target.Address, "G", "F")).value = 0
        
    Else
    
'DONATION IN CASH
        'Sheet11.Range(Replace(Target.Address, "G", "F")).ClearContents
        Sheet11.Range(Replace(Target.Address, "G", "F")).Locked = False
        Sheet11.Range(Replace(Target.Address, "G", "F")).Interior.Color = "&HCCFFCC"
    
    End If
              
End If
                        
'TRANSACTION
'If Not Application.Intersect(Sheet11.Range("SCH_80GGC_TransactionRefNumUPI"), Sheet11.Range(Target.Address)) Is Nothing Then
'
''    If Len(Target.value) > 1 Or (Range(Replace(Target.Address, "J", "G")).value > 0 And Not IsEmpty(Range(Replace(Target.Address, "J", "G")).value)) Or Len(Range(Replace(Target.Address, "J", "K")).value) > 0 Or Len(Range(Replace(Target.Address, "J", "L")).value) > 0 Or Len(Range(Replace(Target.Address, "J", "M")).value) > 0 Or Len(Range(Replace(Target.Address, "J", "N")).value) > 0 Then
'    If ((Not IsEmpty(Target.value) And UCase(Target.value) <> UCase("(Select)"))) Or (Range(Replace(Target.Address, "J", "G")).value > 0 And Not IsEmpty(Range(Replace(Target.Address, "J", "G")).value)) Or Len(Range(Replace(Target.Address, "J", "K")).value) > 0 Or Len(Range(Replace(Target.Address, "J", "L")).value) > 0 Then
'    Sheet11.Unprotect Password:=getmsgstate
'        'DONATION IN CASH
'        Sheet11.Range(Replace(Target.Address, "J", "F")).ClearContents
'        Sheet11.Range(Replace(Target.Address, "J", "F")).Locked = True
'        Sheet11.Range(Replace(Target.Address, "J", "F")).Interior.Color = "&HD8D8D8"
''        Sheet11.Range(Replace(Target.Address, "J", "F")).value = 0
'    Else
'
'       'DONATION IN CASH
'        'Sheet11.Range(Replace(Target.Address, "J", "F")).ClearContents
'        Sheet11.Range(Replace(Target.Address, "J", "F")).Locked = False
'        Sheet11.Range(Replace(Target.Address, "J", "F")).Interior.Color = "&HCCFFCC"
'
'    End If
'
'End If


'CHEQUE
'If Not Application.Intersect(Sheet11.Range("SCH_80GGC_ChnImpsNeftRtgsNum"), Sheet11.Range(Target.Address)) Is Nothing Then
'
''    If Len(Target.value) > 1 Or (Range(Replace(Target.Address, "K", "G")).value > 0 And Not IsEmpty(Range(Replace(Target.Address, "K", "G")).value)) Or Len(Range(Replace(Target.Address, "K", "J")).value) > 0 Or Len(Range(Replace(Target.Address, "K", "L")).value) > 0 Or Len(Range(Replace(Target.Address, "K", "M")).value) > 0 Or Len(Range(Replace(Target.Address, "K", "N")).value) > 0 Then
''     If Len(Target.value) > 1 Or (Range(Replace(Target.Address, "J", "G")).value > 0 And Not IsEmpty(Range(Replace(Target.Address, "J", "G")).value)) And ((Not IsEmpty(Range(Replace(Target.Address, "K", "J")).value)) Or UCase(Range(Replace(Target.Address, "K", "J")).value <> UCase("(Select)"))) Or Len(Range(Replace(Target.Address, "K", "L")).value) > 0 Then
'
'
'
'If (Len(Target.value) > 1) Or (Not IsEmpty(Range(Replace(Target.Address, "J", "G")).value) And Len(Range(Replace(Target.Address, "J", "G")).value) > 0) Or (Not IsEmpty(Range(Replace(Target.Address, "J", "K")).value) And Len(Range(Replace(Target.Address, "J", "K")).value) > 0) Then
'
'
'    Sheet11.Unprotect Password:=getmsgstate
'        'DONATION IN CASH
'        Sheet11.Range(Replace(Target.Address, "J", "F")).ClearContents
'        Sheet11.Range(Replace(Target.Address, "J", "F")).Locked = True
'        Sheet11.Range(Replace(Target.Address, "J", "F")).Interior.Color = "&HD8D8D8"
''        Sheet11.Range(Replace(Target.Address, "K", "F")).value = 0
'    Else
'
'       'DONATION IN CASH
'        'Sheet11.Range(Replace(Target.Address, "K", "F")).ClearContents
'        Sheet11.Range(Replace(Target.Address, "J", "F")).Locked = False
'        Sheet11.Range(Replace(Target.Address, "J", "F")).Interior.Color = "&HCCFFCC"
'
'    End If
'
'End If

'IFSC
'If Not Application.Intersect(Sheet11.Range("SCH_80GGC_IFSCCode"), Sheet11.Range(Target.Address)) Is Nothing Then
'
''  If Len(Target.value) > 1 Or (Range(Replace(Target.Address, "L", "G")).value > 0 And Not IsEmpty(Range(Replace(Target.Address, "L", "G")).value)) Or Len(Range(Replace(Target.Address, "L", "J")).value) > 0 Or Len(Range(Replace(Target.Address, "L", "K")).value) > 0 Or Len(Range(Replace(Target.Address, "L", "M")).value) > 0 Or Len(Range(Replace(Target.Address, "L", "N")).value) > 0 Then
''    If Len(Target.value) > 1 Or (Range(Replace(Target.Address, "K", "G")).value > 0 And Not IsEmpty(Range(Replace(Target.Address, "L", "G")).value)) Or ((Not IsEmpty(Range(Replace(Target.Address, "L", "J")).value)) And UCase(Range(Replace(Target.Address, "L", "J")).value <> UCase("(Select)"))) Or Len(Range(Replace(Target.Address, "L", "K")).value) > 0 Then
'
'
'    If (Target.value > 0 And Not IsEmpty(Target.value)) Or (Not IsEmpty(Range(Replace(Target.Address, "K", "G")).value) And Len(Range(Replace(Target.Address, "K", "G")).value) > 0) Or (Not IsEmpty(Range(Replace(Target.Address, "K", "J")).value) And Len(Range(Replace(Target.Address, "K", "J")).value) > 0) Then
'
'    Sheet11.Unprotect Password:=getmsgstate
'        'DONATION IN CASH
'        Sheet11.Range(Replace(Target.Address, "K", "F")).ClearContents
'        Sheet11.Range(Replace(Target.Address, "K", "F")).Locked = True
'        Sheet11.Range(Replace(Target.Address, "K", "F")).Interior.Color = "&HD8D8D8"
''        Sheet11.Range(Replace(Target.Address, "L", "F")).value = 0
'
'    Else
'
'       'DONATION IN CASH
'        'Sheet11.Range(Replace(Target.Address, "L", "F")).ClearContents
'        Sheet11.Range(Replace(Target.Address, "K", "F")).Locked = False
'        Sheet11.Range(Replace(Target.Address, "K", "F")).Interior.Color = "&HCCFFCC"
'
'    End If
'
'End If
          

        
    Sheet11.Unprotect Password:=getmsgstate
    
    'CheckPAN_80GGC
    
    If Not Intersect(Range("SCH_80GGC_Political_PAN"), Range(Target.Address)) Is Nothing Then
        Sheet11.Unprotect Password:=getmsgstate
    If Sheet11.Range(Target.Address).value <> "" Then
        
     Call CheckPolitical_PAN_80GGC(Sheet11.Range(Target.Address).value, Target.Address)
    End If

End If
If Not Intersect(Range("SCH_80GGC_IFSCCode"), Range(Target.Address)) Is Nothing Then
        Sheet11.Unprotect Password:=getmsgstate
    If Sheet11.Range(Target.Address).value <> "" Then
        
     Call CheckIFSC_80GGC(Sheet11.Range(Target.Address).value, Target.Address)
    End If

End If


Sheet11.Protect Password:=getmsgstate
 
endline:

Application.EnableEvents = True

End Sub
 Sub testttt()
 Application.EnableEvents = True
 End Sub

Public Function ChkMaxDate_80GGc(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

     ChkMaxDate_80GGc = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

'        If Year > 2023 Then
        'PAG_C1 AY 2024-25 Change
        'Ayush_25-26
'        If Year > 2024 Then
'Ayush_26-27
'        If Year > 2025 Then
        If Year > CInt(Sheet8.Range("DOB_Year").value) Then
            ChkMaxDate_80GGc = False
            Exit Function
        Else
'            If Year = 2024 Then
'                If Year = 2025 Then
                If Year = CInt(Sheet8.Range("DOB_Year").value) Then
                If month > 4 Then
                    ChkMaxDate_80GGc = False
                    Exit Function
                Else
                    If month = 4 Then
                        If dat > 1 Then
                           ChkMaxDate_80GGc = False
                            Exit Function
                        Else
                            If dat = 1 Then
                               ChkMaxDate_80GGc = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function

Function CheckPAN_80GGC(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.

    CheckPAN_80GGC = True
    If Len(PAN) > 0 Then
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
            CheckPAN_80GGC = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
            CheckPAN_80GGC = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
            CheckPAN_80GGC = False
            Exit Function
        End If
                
        If Not ChkAlphabet(Mid(PAN, 4, 1)) Then
            CheckPAN_80GGC = False
            Exit Function
        End If
            
        
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
            CheckPAN_80GGC = False
            Exit Function
        End If
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
            CheckPAN_80GGC = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
            CheckPAN_80GGC = False
            Exit Function
        End If
        
        If Not CheckSpecialCharacterdot(PAN) Then
        CheckPAN_80GGC = False
            Exit Function
        End If
               
    End If
End Function


Function CheckPolitical_PAN_80GGC(Tval As Variant, Tadd As Variant) '80GGC_C5
On Error Resume Next
Application.EnableEvents = False
 'sPassword = getmsgstate
   Sheet11.Unprotect Password:=getmsgstate


If Tval <> "" Then
    
     If Not CheckPAN_80GGC(UCase(Tval)) Then
        fmsgbox "Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet  in schedule 80GGC."
           Sheet11.Range(Tadd) = ""
           'Sheet11.Range(Tadd).Offset(0, 1) = ""
           GoTo endfd
     End If
     
     Dim PAN_Ucase As Variant
    PAN_Ucase = UCase(Sheet11.Range(Tadd).value)
    Sheet11.Range(Tadd).value = PAN_Ucase
    
 End If
 

endfd:
Sheet11.Protect Password:=getmsgstate
Application.EnableEvents = True

End Function
