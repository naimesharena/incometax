Attribute VB_Name = "Sheet6"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Activate()
On Error GoTo endline

Application.EnableEvents = False

 Sheet6.Unprotect Password:=getmsgstate

    ActiveSheet.Shapes("Button 1").Left = Range("Sch80G.Button1").Left
    ActiveSheet.Shapes("Button 1").Top = Range("Sch80G.Button1").Top

    ActiveSheet.Shapes("Button 2").Left = Range("Sch80G.Button2").Left
    ActiveSheet.Shapes("Button 2").Top = Range("Sch80G.Button2").Top

    ActiveSheet.Shapes("Button 3").Left = Range("Sch80G.Button3").Left
    ActiveSheet.Shapes("Button 3").Top = Range("Sch80G.Button3").Top

    ActiveSheet.Shapes("Button 4").Left = Range("Sch80G.Button4").Left
    ActiveSheet.Shapes("Button 4").Top = Range("Sch80G.Button4").Top
endline:
 Sheet6.Protect Password:=getmsgstate

Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
Dim rangenm As Variant
    Application.EnableEvents = False

CalculateTaxCheckFlag = False



    If Not checkfieldSuperSpecialcharacter(Target.value) Then
        fmsgbox ("* Please do not enter Special characters < , >" & Chr(13))
        Target.value = ""
    End If
    
'Added by Riyaz on 06/03/2026 for AY 2026-27 TableA
If Not Application.Intersect(Sheet6.Range("Per10080G.IFSC"), Sheet6.Range(Target.Address)) Is Nothing Then

Dim IfsCode As Range
Set IfsCode = Sheet6.Range("Per10080G.IFSC")

    Dim sourceAddress, Address As Variant
    Dim rangecell As Range
    Dim BankNameFromIFSC As Variant
    
    If Not Application.Intersect(IfsCode, Range(Target.Address)) Is Nothing Then
         For Each rangecell In Sheet6.Range("Per10080G.IFSC")
            rangecell.value = UCase(rangecell.value)
     
     If rangecell.value <> "" Then
           
           If Not CheckIFSC(UCase(Target.value)) Then
                    fmsgbox ("* ""Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Sheet6.Unprotect Password:=getmsgstate
                    Target.value = ""
                    GoTo endline
           
            ElseIf Not ValidateIFSCList(UCase(Target.value)) Then
                    fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Target.value = ""
                    GoTo endline
     
            Else:
                       Address = rangecell.Address
                     
                     End If
                
                If Mid(rangecell, 1, 4) <> "GIFT" Then
                    If (Sheet6.Range(Address).value = "") Then
                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                     Address = Replace(sourceAddress, "K", "J")
                     Sheet6.Range(Address).value = ""
                    GoTo endline
                    End If
                End If
                
            End If
    Next
    GoTo endline
End If
End If
'------

'Added by Riyaz on 06/03/2026 for AY 2026-27 TableB
If Not Application.Intersect(Sheet6.Range("PerNO5080G.IFSC"), Sheet6.Range(Target.Address)) Is Nothing Then

Dim IfsCode1 As Range
Set IfsCode1 = Sheet6.Range("PerNO5080G.IFSC")

    Dim sourceAddress1, Address1 As Variant
    Dim rangecell1 As Range
'    Dim BankNameFromIFSC As Variant
    
    If Not Application.Intersect(IfsCode1, Range(Target.Address)) Is Nothing Then
         For Each rangecell1 In Sheet6.Range("PerNO5080G.IFSC")
            rangecell1.value = UCase(rangecell1.value)
     
     If rangecell1.value <> "" Then
           
           If Not CheckIFSC(UCase(Target.value)) Then
                    fmsgbox ("* ""Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Sheet6.Unprotect Password:=getmsgstate
                    Target.value = ""
                    GoTo endline
           
            ElseIf Not ValidateIFSCList(UCase(Target.value)) Then
                    fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Target.value = ""
                    GoTo endline
     
            Else:
                       Address1 = rangecell1.Address
                     
                     End If
                
                If Mid(rangecell1, 1, 4) <> "GIFT" Then
                    If (Sheet6.Range(Address1).value = "") Then
                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                     Address = Replace(sourceAddress1, "K", "J")
                     Sheet6.Range(Address1).value = ""
                    GoTo endline
                    End If
                End If
                
            End If
    Next
    GoTo endline
End If
End If
'------

'Added by Riyaz on 06/03/2026 for AY 2026-27 TableC
If Not Application.Intersect(Sheet6.Range("PerYES10080G.IFSC"), Sheet6.Range(Target.Address)) Is Nothing Then

Dim IfsCode2 As Range
Set IfsCode2 = Sheet6.Range("PerYES10080G.IFSC")

    Dim sourceAddress2, Address2  As Variant
    Dim rangecell2 As Range
'    Dim BankNameFromIFSC As Variant
    
    If Not Application.Intersect(IfsCode2, Range(Target.Address)) Is Nothing Then
         For Each rangecell2 In Sheet6.Range("PerYES10080G.IFSC")
            rangecell2.value = UCase(rangecell2.value)
     
     If rangecell2.value <> "" Then
           
           If Not CheckIFSC(UCase(Target.value)) Then
                    fmsgbox ("* ""Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Sheet6.Unprotect Password:=getmsgstate
                    Target.value = ""
                    GoTo endline
           
            ElseIf Not ValidateIFSCList(UCase(Target.value)) Then
                    fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Target.value = ""
                    GoTo endline
     
            Else:
                       Address2 = rangecell2.Address
                     
                     End If
                
                If Mid(rangecell, 1, 4) <> "GIFT" Then
                    If (Sheet6.Range(Address2).value = "") Then
                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                     Address = Replace(sourceAddress2, "K", "J")
                     Sheet6.Range(Address2).value = ""
                    GoTo endline
                    End If
                End If
                
            End If
    Next
    GoTo endline
End If
End If
'------

'Added by Riyaz on 06/03/2026 for AY 2026-27 TableD
If Not Application.Intersect(Sheet6.Range("Per5080G.IFSC"), Sheet6.Range(Target.Address)) Is Nothing Then

Dim IfsCode3 As Range
Set IfsCode3 = Sheet6.Range("Per5080G.IFSC")

    Dim sourceAddress3, Address3 As Variant
    Dim rangecell3 As Range
'    Dim BankNameFromIFSC As Variant
    
    If Not Application.Intersect(IfsCode3, Range(Target.Address)) Is Nothing Then
         For Each rangecell3 In Sheet6.Range("Per5080G.IFSC")
            rangecell3.value = UCase(rangecell3.value)
     
     If rangecell3.value <> "" Then
           
           If Not CheckIFSC(UCase(Target.value)) Then
                    fmsgbox ("* ""Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Sheet6.Unprotect Password:=getmsgstate
                    Target.value = ""
                    GoTo endline
           
            ElseIf Not ValidateIFSCList(UCase(Target.value)) Then
                    fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Target.value = ""
                    GoTo endline
     
            Else:
                       Address3 = rangecell3.Address
                     
                     End If
                
                If Mid(rangecell3, 1, 4) <> "GIFT" Then
                    If (Sheet6.Range(Address3).value = "") Then
                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                     Address = Replace(sourceAddress3, "K", "J")
                     Sheet6.Range(Address3).value = ""
                    GoTo endline
                    End If
                End If
                
            End If
    Next
    GoTo endline
End If
End If
'------

If Not Application.Intersect(Sheet6.Range("Per10080G.DoneePan"), Sheet6.Range(Target.Address)) Is Nothing Then
       Target.value = UCase(Target.value)
 End If

 If Not Application.Intersect(Sheet6.Range("PerNO5080G.DoneePAN"), Sheet6.Range(Target.Address)) Is Nothing Then
       Target.value = UCase(Target.value)
 End If

 If Not Application.Intersect(Sheet6.Range("PerYES10080G.DoneePAN"), Sheet6.Range(Target.Address)) Is Nothing Then
       Target.value = UCase(Target.value)
 End If
'Change-25.01.2023.103.80G.02B
' If Not Application.Intersect(Sheet6.Range("PerYES10080G.DoneeARN"), Sheet6.Range(Target.Address)) Is Nothing Then
'       Target.value = UCase(Target.value)
' End If
 'End change
 
 If Not Application.Intersect(Sheet6.Range("Per5080G.DoneePAN"), Sheet6.Range(Target.Address)) Is Nothing Then
       Target.value = UCase(Target.value)
 End If


Dim var As Integer
Dim isPAN As Boolean
Dim isCash As Boolean

If Target.Column = 10 Or Target.Column = 12 Then
isPAN = _
    Not Application.Intersect(Target, Sheet6.Range("Per10080G.DoneePAN")) Is Nothing _
 Or Not Application.Intersect(Target, Sheet6.Range("PerNO5080G.DoneePAN")) Is Nothing _
 Or Not Application.Intersect(Target, Sheet6.Range("Per5080G.DoneePAN")) Is Nothing _
 Or Not Application.Intersect(Target, Sheet6.Range("PerYES10080G.DoneePAN")) Is Nothing

isCash = _
    Not Application.Intersect(Target, Sheet6.Range("Per10080G.Donationcash")) Is Nothing _
 Or Not Application.Intersect(Target, Sheet6.Range("PerNO5080G.Donationcash")) Is Nothing _
 Or Not Application.Intersect(Target, Sheet6.Range("Per5080G.DonationCash")) Is Nothing _
 Or Not Application.Intersect(Target, Sheet6.Range("PerYES10080G.Donationcash")) Is Nothing
        
        If Trim(Range("J" & Target.row).value) = "AAAAR1077P" Then
'            If 1 < Application.CountIf(Range("Per10080G.DoneePAN"), Range("J" & Target.row).value) Or 1 < Application.CountIf(Range("PerNO5080G.DoneePAN"), Range("J" & Target.row).value) Or 1 < Application.CountIf(Range("PerYES10080G.DoneePAN"), Range("J" & Target.row).value) Then 'Or 1 < Application.CountIf(Range("Per5080G.DoneePAN"), Range("J" & Target.row).Value)
'                 MsgBox "* Duplicate Pan AAAAR1077P is not allowed for same table"
'                 Range(Target.Address).value = ""
'                 Range("K" & Target.row).value = ""
'                 Range("L" & Target.row).value = ""
'            End If


   ElseIf _
       isPAN Or isCash Then

        Dim rngCash As Range
        Dim rngPAN As Range
        Dim rngVal As Range
        Dim cntOnes As Long
        Dim cell As Range

        'Combine all cash blocks
        Set rngCash = Union( _
            Sheet6.Range("Per10080G.Donationcash"), _
            Sheet6.Range("PerNO5080G.Donationcash"), _
            Sheet6.Range("Per5080G.DonationCash"), _
            Sheet6.Range("PerYES10080G.Donationcash") _
        )
        
        Set rngPAN = Union( _
            Sheet6.Range("Per10080G.DoneePAN"), _
            Sheet6.Range("PerNO5080G.DoneePAN"), _
            Sheet6.Range("Per5080G.DoneePAN"), _
            Sheet6.Range("PerYES10080G.DoneePAN") _
        )
        
        If isPAN Then
            Set ranVal = rngPAN
        ElseIf isCash Then
            Set ranVal = rngCash
        End If


        Application.EnableEvents = False
        Application.Calculate

        'Column V
        cntOnes = Application.WorksheetFunction.CountIf(Sheet6.Range("V:V"), 1)

        If cntOnes > 1 Then

            'Clear
            For Each cell In Application.Intersect(Target, ranVal)
                If isPAN Then
                    If cell.value <> "" Then
                        cell.MergeArea.ClearContents
                    End If

                    ElseIf isCash Then
                    If cell.value <> "" And cell.value <> 0 Then
                        cell.ClearContents
                    End If
                End If
            Next cell

            MsgBox "Same PAN cannot be entered for ""Donation in cash"""

        End If



        Application.EnableEvents = True

End If
End If

Dim comb90Val As Long

If Target.Column = 10 Then   'PAN column

    If Trim(Range("J" & Target.row).value) <> "AAAAR1077P" Then
        '---- Get comb_90 value (USE ONE) ----
        'comb90Val = Sheet6.Range("comb_91").Cells(Target.row).value
        comb90Val = Sheet6.Cells(Target.row, "X").value
        'OR if comb_90 is column W:
        'comb90Val = Sheet6.Cells(Target.Row, "W").Value
       '------------------------------------
    
            If Trim(Target.value) <> "" Then
        
                    If comb90Val > 0 Then
            
                        Application.EnableEvents = False
            
                       ' MsgBox "Same PAN cannot be entered in other block", vbExclamation
                        MsgBox "Same PAN cannot be entered in other block"
                        Target.value = ""
            '            Range("K" & Target.row).value = ""
            '            Range("L" & Target.row).value = ""
            
                        Application.EnableEvents = True
            
                    End If
        
            End If
    
    End If
End If





'Change-25.01.2023.103.80G.02A
'If Target.Column = 11 Then
'If Range("K" & Target.row).value > 2000 Then
'fmsgbox "* If Donation in cash is more than Rs. 2000, 80G Deduction is not allowed"
'End If
'End If

If Target.Column = 12 Then
If Range("L" & Target.row).value > 2000 Then
fmsgbox "* If Donation in cash is more than Rs. 2000, 80G Deduction is not allowed"
End If
End If
'End Change







 If Target.Column = 8 Then
    
'    Dim TargtVal, PIN_targetadd, state_targetadd As Variant
'      state_targetadd = Target.Address
'      state_targetadd = Replace(state_targetadd, "$", "")
'      PIN_targetadd = Replace(state_targetadd, "H", "I")
'
'
'    Call state_Validation(PIN_targetadd, state_targetadd)
'17/09/2024
   
   Dim PIN_targetadd, state_targetadd As String
      'state_targetadd = Target.address
      state_targetadd = Replace(Target.Address, "$", "")
      PIN_targetadd = Replace(state_targetadd, "H", "I")
      
      
    If state_Validation(PIN_targetadd, state_targetadd) = False Then Sheet6.Range(Replace(state_targetadd, "H", "I")).value = ""
       
   End If
   
  If Target.Column = 9 Then
            
            If (Target.value <> "") Then
            
'            Dim pinrng_chk, staterng_chk  As Variant
'
'      pinrng_chk = Target.Address
'      pinrng_chk = Replace(pinrng_chk, "$", "")
'      staterng_chk = Replace(pinrng_chk, "I", "H")
'
'      Call PinState_codes_validation_80G(pinrng_chk, staterng_chk)
  
     
     Dim pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk As String
     Dim ws As Worksheet
      'pinrng_chk = Target.address
      pinrng_chk = Replace(Target.Address, "$", "")
      staterng_chk = Replace(pinrng_chk, "I", "H")
'      cntryrng_chk = Replace(pinrng_chk, "J", "I")
'      ZIPrng_chk = Replace(pinrng_chk, "J", "K")
      
      Set ws = Worksheets("80G")
      
      Call PinState_codes_validation(ws, pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk)
            
             End If
      End If

'If Target.Column = 8 Then
'
'Dim PIN_targetadd, state_targetadd As String
'      'state_targetadd = Target.address
'      state_targetadd = Replace(Target.Address, "$", "")
'      PIN_targetadd = Replace(state_targetadd, "H", "I")
'
'
'    If state_Validation(PIN_targetadd, state_targetadd) = False Then Sheet6.Range(Replace(state_targetadd, "H", "I")).value = ""
'
'   End If
'
'
'
'
'   If Target.Column = 9 Then
'
'    Dim pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk As String
'     Dim ws As Worksheet
'      'pinrng_chk = Target.address
'      pinrng_chk = Replace(Target.Address, "$", "")
'      staterng_chk = Replace(pinrng_chk, "I", "H")
''      cntryrng_chk = Replace(pinrng_chk, "J", "I")
''      ZIPrng_chk = Replace(pinrng_chk, "J", "K")
'
'      Set ws = Worksheets("80G")
'
'      Call PinState_codes_validation(ws, pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk)
'
'             End If
    
'        Dim PinCode As Range
'        For Each PinCode In Sheet8.Range("PincodeList")
'            Dim code As String
'            code = Left(CStr(Target.value), 3)
'            If CStr(PinCode.value) = CStr(code) Then
'                Sheet6.Range("H" & Target.row).value = Sheet8.Range("CM" & PinCode.row).value
'                Exit For
'            End If
'         Next PinCode
'    End If
    
    
    'Ayush1
'Dim pin As Range
'Dim state1, State2, State3, DupFlag, riyaz As Variant
'Dim h, i, j As Long
'Dim StateName
'Dim StateCode
'If Not Application.Intersect(Sheet6.Range("Per10080G.StateCode"), Sheet6.Range(Target.Address)) Is Nothing Then
'           If Sheet6.Range(Replace(Target.Address, "H", "I")).value <> "" Then
'                state1 = Application.VLookup(Sheet6.Range(Replace(Target.Address, "H", "I")).value, Sheet8.Range("All_PinCode_V"), 2, False)
'                State2 = Application.VLookup(Sheet6.Range(Replace(Target.Address, "H", "I")).value, Sheet8.Range("Duplicate_V"), 2, False)
'                State3 = Application.VLookup(Sheet6.Range(Replace(Target.Address, "H", "I")).value, Sheet8.Range("Duplicate_V"), 3, False)
'                If state1 = UCase(Mid(Target.value, 4)) Then
'
'                End If
'                If Not (State2 = UCase(Mid(Target.value, 4)) Or State3 = UCase(Mid(Target.value, 4))) Then
'                Sheet6.Range(Replace(Target.Address, "H", "I")).value = ""
'                End If
'            End If
'        End If
'If Not Application.Intersect(Sheet6.Range("Per10080G.PinCode"), Sheet6.Range(Target.Address)) Is Nothing Then
'         DupFlag = False
'                 For Each pin In Sheet8.Range("Duplicate_Pincode_list")
'                    If Target.value = pin Then
'                        DupFlag = True
'                        Exit For
'                    End If
'                 Next pin
'
'                 If DupFlag = False Then
'
'                    StateName = Application.VLookup(Target.value, Sheet8.Range("All_PinCode_V"), 2, False)
'                     If Not StateName = UCase(Mid(Sheet6.Range(Replace(Target.Address, "I", "H")).value, 4)) Then
'                        StateCode = Application.VLookup(StateName, Sheet8.Range("EW:EX"), 2, False)
'                        Sheet6.Range(Replace(Target.Address, "I", "H")).value = StateCode
'                        'Sheet2.Range(Replace(Target.value, "J", "I")) = "91-INDIA"
'                        'Sheet1.Range("sheet1.MobileCountryCode").value = "91"
'                    End If
'                 End If
'
'                 If DupFlag = True Then
'                 Application.EnableEvents = False
'                    'If Sheet6.Range(Replace(Target.address, "I", "H")) <> "" Or Sheet6.Range(Replace(Target.address, "I", "H")) = "" Then
'
'                     state1 = Application.VLookup(Target.value, Sheet8.Range("Duplicate_V"), 2, False)
'                     State2 = Application.VLookup(Target.value, Sheet8.Range("Duplicate_V"), 3, False)
'
'                      If Not UCase(Mid(Sheet6.Range(Replace(Target.Address, "I", "H")), 4)) = state1 Or UCase(Mid(Sheet6.Range(Replace(Target.Address, "I", "H")), 4)) = State2 Then
'                      Sheet6.Range(Replace(Target.Address, "I", "H")) = "(Select)"
'                       End If
'                    'End If
'        End If
'        End If
'        '-------------------------------------------
'
'
'
' 'Ayush2
'' Dim pin As Range
''Dim State1, State2, State3, DupFlag, riyaz As Variant
''Dim h, i, j As Long
''Dim StateName
''Dim StateCode
'If Not Application.Intersect(Sheet6.Range("PerNO5080G.StateCode"), Sheet6.Range(Target.Address)) Is Nothing Then
'           If Sheet6.Range(Replace(Target.Address, "H", "I")).value <> "" Then
'                state1 = Application.VLookup(Sheet6.Range(Replace(Target.Address, "H", "I")).value, Sheet8.Range("All_PinCode_V"), 2, False)
'                State2 = Application.VLookup(Sheet6.Range(Replace(Target.Address, "H", "I")).value, Sheet8.Range("Duplicate_V"), 2, False)
'                State3 = Application.VLookup(Sheet6.Range(Replace(Target.Address, "H", "I")).value, Sheet8.Range("Duplicate_V"), 3, False)
'                If state1 = UCase(Mid(Target.value, 4)) Then
'
'                End If
'                If Not (State2 = UCase(Mid(Target.value, 4)) Or State3 = UCase(Mid(Target.value, 4))) Then
'                Sheet6.Range(Replace(Target.Address, "H", "I")).value = ""
'                End If
'            End If
'        End If
'If Not Application.Intersect(Sheet6.Range("PerNO5080G.PinCode"), Sheet6.Range(Target.Address)) Is Nothing Then
'         DupFlag = False
'                 For Each pin In Sheet8.Range("Duplicate_Pincode_list")
'                    If Target.value = pin Then
'                        DupFlag = True
'                        Exit For
'                    End If
'                 Next pin
'
'                 If DupFlag = False Then
'
'                    StateName = Application.VLookup(Target.value, Sheet8.Range("All_PinCode_V"), 2, False)
'                     If Not StateName = UCase(Mid(Sheet6.Range(Replace(Target.Address, "I", "H")).value, 4)) Then
'                        StateCode = Application.VLookup(StateName, Sheet8.Range("EW:EX"), 2, False)
'                        Sheet6.Range(Replace(Target.Address, "I", "H")).value = StateCode
'                        'Sheet2.Range(Replace(Target.value, "J", "I")) = "91-INDIA"
'                        'Sheet1.Range("sheet1.MobileCountryCode").value = "91"
'                    End If
'                 End If
'
'                 If DupFlag = True Then
'                 Application.EnableEvents = False
'                    'If Sheet6.Range(Replace(Target.address, "I", "H")) <> "" Or Sheet6.Range(Replace(Target.address, "I", "H")) = "" Then
'
'                     state1 = Application.VLookup(Target.value, Sheet8.Range("Duplicate_V"), 2, False)
'                     State2 = Application.VLookup(Target.value, Sheet8.Range("Duplicate_V"), 3, False)
'
'                      If Not UCase(Mid(Sheet6.Range(Replace(Target.Address, "I", "H")), 4)) = state1 Or UCase(Mid(Sheet6.Range(Replace(Target.Address, "I", "H")), 4)) = State2 Then
'                      Sheet6.Range(Replace(Target.Address, "I", "H")) = "(Select)"
'                       End If
'                    'End If
'        End If
'        End If
'
'
'        'Ayush3
'
'If Not Application.Intersect(Sheet6.Range("PerYES10080G.StateCode"), Sheet6.Range(Target.Address)) Is Nothing Then
'           If Sheet6.Range(Replace(Target.Address, "H", "I")).value <> "" Then
'                state1 = Application.VLookup(Sheet6.Range(Replace(Target.Address, "H", "I")).value, Sheet8.Range("All_PinCode_V"), 2, False)
'                State2 = Application.VLookup(Sheet6.Range(Replace(Target.Address, "H", "I")).value, Sheet8.Range("Duplicate_V"), 2, False)
'                State3 = Application.VLookup(Sheet6.Range(Replace(Target.Address, "H", "I")).value, Sheet8.Range("Duplicate_V"), 3, False)
'                If state1 = UCase(Mid(Target.value, 4)) Then
'
'                End If
'                If Not (State2 = UCase(Mid(Target.value, 4)) Or State3 = UCase(Mid(Target.value, 4))) Then
'                Sheet6.Range(Replace(Target.Address, "H", "I")).value = ""
'                End If
'            End If
'        End If
'If Not Application.Intersect(Sheet6.Range("PerYES10080G.PinCode"), Sheet6.Range(Target.Address)) Is Nothing Then
'         DupFlag = False
'                 For Each pin In Sheet8.Range("Duplicate_Pincode_list")
'                    If Target.value = pin Then
'                        DupFlag = True
'                        Exit For
'                    End If
'                 Next pin
'
'                 If DupFlag = False Then
'
'                    StateName = Application.VLookup(Target.value, Sheet8.Range("All_PinCode_V"), 2, False)
'                     If Not StateName = UCase(Mid(Sheet6.Range(Replace(Target.Address, "I", "H")).value, 4)) Then
'                        StateCode = Application.VLookup(StateName, Sheet8.Range("EW:EX"), 2, False)
'                        Sheet6.Range(Replace(Target.Address, "I", "H")).value = StateCode
'                        'Sheet2.Range(Replace(Target.value, "J", "I")) = "91-INDIA"
'                        'Sheet1.Range("sheet1.MobileCountryCode").value = "91"
'                    End If
'                 End If
'
'                 If DupFlag = True Then
'                 Application.EnableEvents = False
'                    'If Sheet6.Range(Replace(Target.address, "I", "H")) <> "" Or Sheet6.Range(Replace(Target.address, "I", "H")) = "" Then
'
'                     state1 = Application.VLookup(Target.value, Sheet8.Range("Duplicate_V"), 2, False)
'                     State2 = Application.VLookup(Target.value, Sheet8.Range("Duplicate_V"), 3, False)
'
'                      If Not UCase(Mid(Sheet6.Range(Replace(Target.Address, "I", "H")), 4)) = state1 Or UCase(Mid(Sheet6.Range(Replace(Target.Address, "I", "H")), 4)) = State2 Then
'                      Sheet6.Range(Replace(Target.Address, "I", "H")) = "(Select)"
'                       End If
'                    'End If
'        End If
'        End If
'
'
'        'Ayush4
'
'If Not Application.Intersect(Sheet6.Range("Per5080G.StateCode"), Sheet6.Range(Target.Address)) Is Nothing Then
'           If Sheet6.Range(Replace(Target.Address, "H", "I")).value <> "" Then
'                state1 = Application.VLookup(Sheet6.Range(Replace(Target.Address, "H", "I")).value, Sheet8.Range("All_PinCode_V"), 2, False)
'                State2 = Application.VLookup(Sheet6.Range(Replace(Target.Address, "H", "I")).value, Sheet8.Range("Duplicate_V"), 2, False)
'                State3 = Application.VLookup(Sheet6.Range(Replace(Target.Address, "H", "I")).value, Sheet8.Range("Duplicate_V"), 3, False)
'                If state1 = UCase(Mid(Target.value, 4)) Then
'
'                End If
'                If Not (State2 = UCase(Mid(Target.value, 4)) Or State3 = UCase(Mid(Target.value, 4))) Then
'                Sheet6.Range(Replace(Target.Address, "H", "I")).value = ""
'                End If
'            End If
'        End If
'If Not Application.Intersect(Sheet6.Range("Per5080G.PinCode"), Sheet6.Range(Target.Address)) Is Nothing Then
'         DupFlag = False
'                 For Each pin In Sheet8.Range("Duplicate_Pincode_list")
'                    If Target.value = pin Then
'                        DupFlag = True
'                        Exit For
'                    End If
'                 Next pin
'
'                 If DupFlag = False Then
'
'                    StateName = Application.VLookup(Target.value, Sheet8.Range("All_PinCode_V"), 2, False)
'                     If Not StateName = UCase(Mid(Sheet6.Range(Replace(Target.Address, "I", "H")).value, 4)) Then
'                        StateCode = Application.VLookup(StateName, Sheet8.Range("EW:EX"), 2, False)
'                        Sheet6.Range(Replace(Target.Address, "I", "H")).value = StateCode
'                        'Sheet2.Range(Replace(Target.value, "J", "I")) = "91-INDIA"
'                        'Sheet1.Range("sheet1.MobileCountryCode").value = "91"
'                    End If
'                 End If
'
'                 If DupFlag = True Then
'                 Application.EnableEvents = False
'                    'If Sheet6.Range(Replace(Target.address, "I", "H")) <> "" Or Sheet6.Range(Replace(Target.address, "I", "H")) = "" Then
'
'                     state1 = Application.VLookup(Target.value, Sheet8.Range("Duplicate_V"), 2, False)
'                     State2 = Application.VLookup(Target.value, Sheet8.Range("Duplicate_V"), 3, False)
'
'                      If Not UCase(Mid(Sheet6.Range(Replace(Target.Address, "I", "H")), 4)) = state1 Or UCase(Mid(Sheet6.Range(Replace(Target.Address, "I", "H")), 4)) = State2 Then
'                      Sheet6.Range(Replace(Target.Address, "I", "H")) = "(Select)"
'                       End If
'                    'End If
'        End If
'        End If
'



If Not Intersect(Range("Per10080G.Donationcash"), Range(Target.Address)) Is Nothing Then
Sheet6.Unprotect Password:=getmsgstate
        'If Sheet6.Range(Target.Address).value <> "" Or Sheet6.Range(Target.Address).value > 0 Then

           If Sheet6.Range(Target.Address).value > 0 Then
           Sheet6.Range(Target.Address).Interior.Color = "&HCCFFCC"
           
            
           
            Sheet6.Range(Replace(Target.Address, "L", "M")).ClearContents
            Sheet6.Range(Replace(Target.Address, "L", "M")).Locked = True
            Sheet6.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HD8D8D8"
'            Sheet6.Range(Replace(Target.Address, "L", "M")).value = 0
            
            
            Sheet6.Range(Replace(Target.Address, "L", "N")).ClearContents
            Sheet6.Range(Replace(Target.Address, "L", "N")).Locked = True
            Sheet6.Range(Replace(Target.Address, "L", "N")).Interior.Color = "&HD8D8D8"
            
             
            Sheet6.Range(Replace(Target.Address, "L", "O")).ClearContents
            Sheet6.Range(Replace(Target.Address, "L", "O")).Locked = True
            Sheet6.Range(Replace(Target.Address, "L", "O")).Interior.Color = "&HD8D8D8"
        
        Else
                Sheet6.Unprotect Password:=getmsgstate
                Sheet6.Range(Replace(Target.Address, "L", "M")).ClearContents
                Sheet6.Range(Replace(Target.Address, "L", "M")).Locked = False
                Sheet6.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HCCFFCC"
                
                Sheet6.Range(Replace(Target.Address, "L", "N")).Locked = False
                Sheet6.Range(Replace(Target.Address, "L", "N")).Interior.Color = "&HCCFFCC"
                
                Sheet6.Range(Replace(Target.Address, "L", "O")).Locked = False
                Sheet6.Range(Replace(Target.Address, "L", "O")).Interior.Color = "&HCCFFCC"
                
                 
Sheet6.Protect Password:=getmsgstate
    End If
End If

If Not Intersect(Range("PerNO5080G.DonationCash"), Range(Target.Address)) Is Nothing Then
Sheet6.Unprotect Password:=getmsgstate
        'If Sheet6.Range(Target.Address).value <> "" Or Sheet6.Range(Target.Address).value > 0 Then

           If Sheet6.Range(Target.Address).value > 0 Then
           Sheet6.Range(Target.Address).Interior.Color = "&HCCFFCC"
           
            
           
            Sheet6.Range(Replace(Target.Address, "L", "M")).ClearContents
            Sheet6.Range(Replace(Target.Address, "L", "M")).Locked = True
            Sheet6.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HD8D8D8"
'            Sheet6.Range(Replace(Target.Address, "L", "M")).value = 0
            
            
            Sheet6.Range(Replace(Target.Address, "L", "N")).ClearContents
            Sheet6.Range(Replace(Target.Address, "L", "N")).Locked = True
            Sheet6.Range(Replace(Target.Address, "L", "N")).Interior.Color = "&HD8D8D8"
            
             
            Sheet6.Range(Replace(Target.Address, "L", "O")).ClearContents
            Sheet6.Range(Replace(Target.Address, "L", "O")).Locked = True
            Sheet6.Range(Replace(Target.Address, "L", "O")).Interior.Color = "&HD8D8D8"
        
        Else
                Sheet6.Unprotect Password:=getmsgstate
                Sheet6.Range(Replace(Target.Address, "L", "M")).ClearContents
                Sheet6.Range(Replace(Target.Address, "L", "M")).Locked = False
                Sheet6.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HCCFFCC"
                
                Sheet6.Range(Replace(Target.Address, "L", "N")).Locked = False
                Sheet6.Range(Replace(Target.Address, "L", "N")).Interior.Color = "&HCCFFCC"
                
                Sheet6.Range(Replace(Target.Address, "L", "O")).Locked = False
                Sheet6.Range(Replace(Target.Address, "L", "O")).Interior.Color = "&HCCFFCC"
                
                 
Sheet6.Protect Password:=getmsgstate
    End If
End If


If Not Intersect(Range("PerYES10080G.DonationCash"), Range(Target.Address)) Is Nothing Then
Sheet6.Unprotect Password:=getmsgstate
        'If Sheet6.Range(Target.Address).value <> "" Or Sheet6.Range(Target.Address).value > 0 Then

           If Sheet6.Range(Target.Address).value > 0 Then
           Sheet6.Range(Target.Address).Interior.Color = "&HCCFFCC"
           
            
           
            Sheet6.Range(Replace(Target.Address, "L", "M")).ClearContents
            Sheet6.Range(Replace(Target.Address, "L", "M")).Locked = True
            Sheet6.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HD8D8D8"
'            Sheet6.Range(Replace(Target.Address, "L", "M")).value = 0
            
            
            Sheet6.Range(Replace(Target.Address, "L", "N")).ClearContents
            Sheet6.Range(Replace(Target.Address, "L", "N")).Locked = True
            Sheet6.Range(Replace(Target.Address, "L", "N")).Interior.Color = "&HD8D8D8"
            
             
            Sheet6.Range(Replace(Target.Address, "L", "O")).ClearContents
            Sheet6.Range(Replace(Target.Address, "L", "O")).Locked = True
            Sheet6.Range(Replace(Target.Address, "L", "O")).Interior.Color = "&HD8D8D8"
        
        Else
                Sheet6.Unprotect Password:=getmsgstate
                Sheet6.Range(Replace(Target.Address, "L", "M")).ClearContents
                Sheet6.Range(Replace(Target.Address, "L", "M")).Locked = False
                Sheet6.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HCCFFCC"
                
                Sheet6.Range(Replace(Target.Address, "L", "N")).Locked = False
                Sheet6.Range(Replace(Target.Address, "L", "N")).Interior.Color = "&HCCFFCC"
                
                Sheet6.Range(Replace(Target.Address, "L", "O")).Locked = False
                Sheet6.Range(Replace(Target.Address, "L", "O")).Interior.Color = "&HCCFFCC"
                
                 
Sheet6.Protect Password:=getmsgstate
    End If
End If


If Not Intersect(Range("Per5080G.DonationCash"), Range(Target.Address)) Is Nothing Then
Sheet6.Unprotect Password:=getmsgstate
        'If Sheet6.Range(Target.Address).value <> "" Or Sheet6.Range(Target.Address).value > 0 Then

           If Sheet6.Range(Target.Address).value > 0 Then
           Sheet6.Range(Target.Address).Interior.Color = "&HCCFFCC"
           
            
           
            Sheet6.Range(Replace(Target.Address, "L", "M")).ClearContents
            Sheet6.Range(Replace(Target.Address, "L", "M")).Locked = True
            Sheet6.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HD8D8D8"
'            Sheet6.Range(Replace(Target.Address, "L", "M")).value = 0
            
            
            Sheet6.Range(Replace(Target.Address, "L", "N")).ClearContents
            Sheet6.Range(Replace(Target.Address, "L", "N")).Locked = True
            Sheet6.Range(Replace(Target.Address, "L", "N")).Interior.Color = "&HD8D8D8"
            
             
            Sheet6.Range(Replace(Target.Address, "L", "O")).ClearContents
            Sheet6.Range(Replace(Target.Address, "L", "O")).Locked = True
            Sheet6.Range(Replace(Target.Address, "L", "O")).Interior.Color = "&HD8D8D8"
        
        Else
                Sheet6.Unprotect Password:=getmsgstate
                Sheet6.Range(Replace(Target.Address, "L", "M")).ClearContents
                Sheet6.Range(Replace(Target.Address, "L", "M")).Locked = False
                Sheet6.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HCCFFCC"
                
                Sheet6.Range(Replace(Target.Address, "L", "N")).Locked = False
                Sheet6.Range(Replace(Target.Address, "L", "N")).Interior.Color = "&HCCFFCC"
                
                Sheet6.Range(Replace(Target.Address, "L", "O")).Locked = False
                Sheet6.Range(Replace(Target.Address, "L", "O")).Interior.Color = "&HCCFFCC"
                
                 
Sheet6.Protect Password:=getmsgstate
    End If
End If


If Not Application.Intersect(Sheet6.Range("Per10080G.DonationAmt"), Sheet6.Range(Target.Address)) Is Nothing Then

If Target.value > 0 Then



  Sheet6.Unprotect Password:=getmsgstate
    'DONATION IN CASH
        Sheet6.Range(Replace(Target.Address, "M", "L")).ClearContents
        Sheet6.Range(Replace(Target.Address, "M", "L")).Locked = True
        Sheet6.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"

        
    Else
    
'DONATION IN CASH
        'Sheet6.Range(Replace(Target.Address, "M", "L")).ClearContents
        Sheet6.Range(Replace(Target.Address, "M", "L")).Locked = False
        Sheet6.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"
    
    End If
              
End If

If Not Application.Intersect(Sheet6.Range("PerNO5080G.DonationAmt"), Sheet6.Range(Target.Address)) Is Nothing Then

If Target.value > 0 Then



  Sheet6.Unprotect Password:=getmsgstate
    'DONATION IN CASH
        Sheet6.Range(Replace(Target.Address, "M", "L")).ClearContents
        Sheet6.Range(Replace(Target.Address, "M", "L")).Locked = True
        Sheet6.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"

        
    Else
    
'DONATION IN CASH
        'Sheet6.Range(Replace(Target.Address, "M", "L")).ClearContents
        Sheet6.Range(Replace(Target.Address, "M", "L")).Locked = False
        Sheet6.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"
    
    End If
              
End If


If Not Application.Intersect(Sheet6.Range("PerYES10080G.DonationAmt"), Sheet6.Range(Target.Address)) Is Nothing Then

If Target.value > 0 Then



  Sheet6.Unprotect Password:=getmsgstate
    'DONATION IN CASH
        Sheet6.Range(Replace(Target.Address, "M", "L")).ClearContents
        Sheet6.Range(Replace(Target.Address, "M", "L")).Locked = True
        Sheet6.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"

        
    Else
    
'DONATION IN CASH
        'Sheet6.Range(Replace(Target.Address, "M", "L")).ClearContents
        Sheet6.Range(Replace(Target.Address, "M", "L")).Locked = False
        Sheet6.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"
    
    End If
              
End If


If Not Application.Intersect(Sheet6.Range("Per5080G.DonationAmt"), Sheet6.Range(Target.Address)) Is Nothing Then

If Target.value > 0 Then



  Sheet6.Unprotect Password:=getmsgstate
    'DONATION IN CASH
        Sheet6.Range(Replace(Target.Address, "M", "L")).ClearContents
        Sheet6.Range(Replace(Target.Address, "M", "L")).Locked = True
        Sheet6.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"

        
    Else
    
'DONATION IN CASH
        'Sheet6.Range(Replace(Target.Address, "M", "L")).ClearContents
        Sheet6.Range(Replace(Target.Address, "M", "L")).Locked = False
        Sheet6.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"
    
    End If
              
End If



endline:
    Application.EnableEvents = True
End Sub


Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
Cancel = True
End Sub
