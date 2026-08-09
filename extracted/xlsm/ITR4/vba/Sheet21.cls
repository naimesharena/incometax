Attribute VB_Name = "Sheet21"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Private Sub Worksheet_Activate()
 Sheet21.Protect Password:=getmsgstate

End Sub

'Commented by Ayush_05/05/2025 for 0.7 changes

'Function GetBankName(IFSC As Variant) As String
'On Error Resume Next
'GetBankName = ""
'Dim rangecells As Range
'Dim Address As Variant
'For Each rangecells In Sheet8.Range("IFSC_StartLetter")
'
'If UCase(rangecells.value) = UCase(IFSC) Then
'Address = rangecells.Address
'Address = Replace(Address, "E", "F")
'GetBankName = Sheet8.Range(Address).value
'Exit Function
'End If
'
'Next
'
''If GetBankName = "" Then
''GetBankName = False
''End If
'
'End Function
'
'Sub LockUnlockIFSC()
'On Error Resume Next
'Application.EnableEvents = False
'Sheet21.Unprotect Password:=getmsgstate
'    If Mid(Sheet21.Range("LoanfrmBankOrInstitute.24b").value, 1, 1) = "O" Then
'    Sheet21.Unprotect Password:=getmsgstate
'        Sheet21.Range("IFSC.24b").ClearContents
'       ' Sheet21.Range("IFSC.24b").Locked = True
'        Sheet21.Range("IFSC.24b").Interior.Color = "&HD8D8D8"
'        Sheet21.Range("IFSC.24b").value = ""
'        Sheet21.Range("IFSC.24b").Locked = True
'
'    Else
'    Sheet21.Unprotect Password:=getmsgstate
'        Sheet21.Range("IFSC.24b").ClearContents
'        Sheet21.Range("IFSC.24b").Locked = False
'        Sheet21.Range("IFSC.24b").Interior.Color = "&HCCFFCC"
'
'    End If
'Sheet21.Protect Password:=getmsgstate
'Application.EnableEvents = True
'End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
Dim TargtVal, targetadd As Variant
Dim IffCode As Range
Dim IfsCode As Range


Application.EnableEvents = False
Application.ScreenUpdating = False

CalculateTaxCheckFlag = False
Sheet21.Unprotect Password:=getmsgstate

'Ayush_18/04

'Newly added by Bindu
TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")


If Not Application.Intersect(Sheet21.Range("loanDate.24b"), Sheet21.Range(Target.Address)) Is Nothing Then
       
        If Not (ValidateDate_24b(Target)) Then
        
         End If
           GoTo endline
End If
        
         



'Commented by Ayush_05/05/2025 for 0.7 changes
'Ayush---------------16/04/2025
 
'If Not Application.Intersect(Sheet21.Range("LoanfrmBankOrInstitute.24b"), Sheet21.Range(Target.Address)) Is Nothing Then
'
'
'If Sheet21.Range(Target.Address).value <> "" Or Sheet11.Range(Target.Address).value <> "(Select)" Then
'
'            If Mid(Sheet21.Range(Target.Address).value, 1, 1) = "O" Then
'
'            Sheet21.Unprotect Password:=getmsgstate
'
'            Sheet21.Range(Replace(Target.Address, "D", "E")).ClearContents
'
'            Sheet21.Range(Replace(Target.Address, "D", "E")).Locked = True
'
'            Sheet21.Range(Replace(Target.Address, "D", "E")).Interior.Color = "&HD8D8D8"
'
'            Sheet21.Unprotect Password:=getmsgstate
'
'            Sheet21.Range(Replace(Target.Address, "D", "G")).ClearContents
'
'            'Sheet21.Range(Replace(Target.Address, "D", "G")).Locked = True
'
'            Sheet21.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HCCFFCC"
'            Sheet21.Unprotect Password:=getmsgstate
'
'            Sheet21.Range(Replace(Target.Address, "D", "F")).ClearContents
'
'            Sheet21.Range(Replace(Target.Address, "D", "F")).Locked = False
'
'            Sheet21.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HCCFFCC"
'            ElseIf Mid(Sheet21.Range(Target.Address).value, 1, 1) = "B" Then
'
'            Sheet21.Unprotect Password:=getmsgstate
'            Sheet21.Range(Replace(Target.Address, "D", "E")).ClearContents
'            Sheet21.Unprotect Password:=getmsgstate
'            Sheet21.Range(Replace(Target.Address, "D", "E")).Locked = False
'            Sheet21.Unprotect Password:=getmsgstate
'            Sheet21.Range(Replace(Target.Address, "D", "E")).Interior.Color = "&HCCFFCC"
'
'
'            Sheet21.Unprotect Password:=getmsgstate
'
'            Sheet21.Range(Replace(Target.Address, "D", "F")).ClearContents
'
'            Sheet21.Range(Replace(Target.Address, "D", "F")).Locked = True
'
'            Sheet21.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HD8D8D8"
'
'
'            Else
'
'            Sheet21.Unprotect Password:=getmsgstate
'
'            Sheet21.Range(Replace(Target.Address, "D", "E")).ClearContents
'
'            Sheet21.Range(Replace(Target.Address, "D", "E")).Locked = False
'
'            Sheet21.Range(Replace(Target.Address, "D", "E")).Interior.Color = "&HCCFFCC"
'
'            Sheet21.Unprotect Password:=getmsgstate
'
'            Sheet21.Range(Replace(Target.Address, "D", "G")).ClearContents
'
'            Sheet21.Range(Replace(Target.Address, "D", "G")).Locked = False
'
'            Sheet21.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HCCFFCC"
'
'
'
'            Sheet21.Unprotect Password:=getmsgstate
'
'            Sheet21.Range(Replace(Target.Address, "D", "F")).ClearContents
'
'            Sheet21.Range(Replace(Target.Address, "D", "F")).Locked = False
'
'            Sheet21.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HCCFFCC"
'
'
'            End If
'End If
'End If
 
'-----------------------


'Commented by Ayush_05/05/2025 for 0.7 changes
'If Not Application.Intersect(Sheet21.Range("PAN.24b"), Sheet21.Range(Target.Address)) Is Nothing Then
'
''If (Target.name.name) = "sheet1.PANRepAssessee" Then
'
'    Dim TargetPAN1  As String
'
'    TargetPAN1 = ""
'
'    TargetPAN1 = Sheet21.Range(Target.Address).value
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
'            Sheet21.Unprotect Password:=getmsgstate
'
'           Sheet21.Range(Target.Address).ClearContents
'
'      '      TargetPAN1 = ""
'
''            Sheet21.Range(TargetPAN1) = ""
'
''                    Sheet21.Range(TargetPAN1).Offset(0, 1) = ""
'
''
'
'            GoTo endline
'
'        End If
'
'       ' TargetPAN1 = UCase(TargetPAN1)
'
'       Sheet21.Range(Target.Address).value = UCase(TargetPAN1)
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
''            Sheet21.Range(TargetPAN1).Select
'
''
'
''            Sheet21.Range(TargetPAN1).value = ""
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
' Set IfsCode = Range("IFSC.24b")
'
'
'    Dim sourceAddress, Address, address1 As Variant
'    Dim rangecell As Range
'    Dim BankNameFromIFSC As Variant
'
'    If Not Application.Intersect(IfsCode, Range(Target.Address)) Is Nothing Then
'         For Each rangecell In Sheet21.Range("IFSC.24b")
'            rangecell.value = UCase(rangecell.value)
'
'     If rangecell.value <> "" Then
'
'           If Not CheckIFSC(UCase(TargtVal)) Then
'                    'fmsgboxsmall "* Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
'                     fmsgbox ("* ""Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
'                    Sheet21.Range(targetadd) = ""
'                    Sheet21.Range(targetadd).Offset(0, 1) = ""
'                    GoTo endline
'
'            ElseIf Not ValidateIFSCList(UCase(TargtVal)) Then
'                   ' fmsgboxsmall "* Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
'                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
'                    Sheet21.Range(targetadd) = ""
'                    Sheet21.Range(targetadd).Offset(0, 1) = ""
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
'                   Sheet21.Range(Address).value = BankNameFromIFSC
'                ElseIf Mid(rangecell, 1, 4) <> "GIFT" Then
'                   Sheet21.Range(Address).value = BankNameFromIFSC
'                End If
'
'                If Mid(rangecell, 1, 4) <> "GIFT" Then
'                    If (Sheet21.Range(Address).value = "") Then
'                     'MsgBox "IFS Code: Please enter valid IFS Code. Refer to your bank for valid codes. (1st 4 Alpha, followed by 0 and 6 digits)"
'                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
'                     Address = Replace(sourceAddress, "K", "J")
'                     Sheet21.Range(Address).value = ""
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

endline:

Application.EnableEvents = True
Sheet21.Protect Password:=sPassword
Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub


Function ValidateDate_24b(ByVal Target As Range) As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim DateDep_IT
    Dim mIntCells As Long
    Dim mIntCtr As Long
    
    
    mIntCells = Sheet21.Range("loanDate.24b").count
    ValidateDate_24b = True

    Dim end_24
    end_24 = 0
    Set rangecells = Sheet21.Range("loanDate.24b").Cells

    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            end_24 = end_24 + 1
        End If
    Next

    ReDim Date_24b(end_24)

    For i = 1 To end_24
        Date_24b(i) = rangecells.item(i).value
'        If Not chkCompulsory(Date_24b(i)) Then
'            'fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is mandatory") '
'            fmsgboxsmall ("Please provide Date of taking loan in schedule 24(b) at Sr. No  " & i & "")
'            ValidateDate_24b = False
'            Exit Function
'        End If

        If Not CheckDateddmmyyyy(Date_24b(i)) Then
            ValidateDate_24b = False
'            fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is invalid. Pl enter in dd/mm/yyyy format")
'Ayush_10/02
'              fmsgboxsmall ("Please enter a valid date in dd/mm/yyyy format at Sr. No  " & i & "  in Sheet IT ")
                fmsgboxsmall ("""Please enter date in valid format"" at Sr. No  " & i & "")
            Target.value = ""
            Exit Function
        Else
            Date_24b(i) = Dformat(Date_24b(i), "")
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
            
            
            If Not ChkMaxDate_24b(Trim(Date_24b(i)), "31-03-2025") Then
                fmsgboxsmall (" Date can not be after 31/03/2025 at Sr. No " & i & ".")
                Target.value = ""
                ValidateDate_24b = False
                 

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

