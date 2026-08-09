Attribute VB_Name = "Sheet20"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Dim ActiveCellRange As String


Private Sub Worksheet_Activate()
On Error Resume Next
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet20.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Sheet20.Shapes("Button 17").Left = Range("Button_80G1").Left
Sheet20.Shapes("Button 17").Top = Range("Button_80G1").Top + 2

Sheet20.Shapes("Button 18").Left = Range("Button_80G2").Left
Sheet20.Shapes("Button 18").Top = Range("Button_80G2").Top + 2

Sheet20.Shapes("Button 19").Left = Range("Button_80G3").Left
Sheet20.Shapes("Button 19").Top = Range("Button_80G3").Top + 2

Sheet20.Shapes("Button 20").Left = Range("Button_80G4").Left
Sheet20.Shapes("Button 20").Top = Range("Button_80G4").Top + 2

Sheet20.Protect Password:=sPassword


ActiveSheet.Unprotect getmsgstate
    Dim rCell
    For Each rCell In Range("Per10080G.DoneePAN").Cells
        With rCell
            .ClearComments
            .AddComment
            .Comment.text text:="Enter the PAN of the person to whom the donation is made"
        End With
    Next
    For Each rCell In Range("PerNO5080G.DoneePAN").Cells
        With rCell
            .ClearComments
            .AddComment
            .Comment.text text:="Enter the PAN of the person to whom the donation is made"
        End With
    Next
    For Each rCell In Range("PerYES10080G.DoneePAN").Cells
        With rCell
            .ClearComments
            .AddComment
            .Comment.text text:="Enter the PAN of the person to whom the donation is made"
        End With
    Next
    For Each rCell In Range("Per5080G.DoneePAN").Cells
        With rCell
            .ClearComments
            .AddComment
            .Comment.text text:="Enter the PAN of the person to whom the donation is made"
        End With
    Next
    ActiveSheet.Protect getmsgstate

End Sub

Private Sub Worksheet_Calculate()

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

CalculateTaxCheckFlag = False

Application.EnableEvents = False
'Sheet19.DI_Hidden
Target.value = UCase(Target.value)

'Change-29.04.2023.101.80G -> SIT-49822
'Change14.03.2023.102.IDS10
'If Not Application.Intersect(Range("Per5080G.ARN_Reference"), Range(Target.Address)) Is Nothing Then
'    If Not isAlphanumeric(Target.value) Then
'        MsgBox ("ARN can be alphanumerics only")
'        Target.value = ""
'    End If
'End If
'-----


'Added by Aavula Naresh for AY 2026-27

If Not Application.Intersect(Sheet20.Range("IFSC_80GA"), Sheet20.Range(Target.Address)) Is Nothing Then

Dim IfsCode As Range
Set IfsCode = Sheet20.Range("IFSC_80GA")

    Dim sourceAddress, Address As Variant
    Dim rangecell As Range
    Dim BankNameFromIFSC As Variant
    
    If Not Application.Intersect(IfsCode, Range(Target.Address)) Is Nothing Then
         For Each rangecell In Sheet20.Range("IFSC_80GA")
            rangecell.value = UCase(rangecell.value)
     
     If rangecell.value <> "" Then
           
           If Not CheckIFSC(UCase(Target.value)) Then
                    fmsgbox ("* ""Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Sheet20.Unprotect Password:=getmsgstate
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
                    If (Sheet20.Range(Address).value = "") Then
                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                     Address = Replace(sourceAddress, "K", "J")
                     Sheet20.Range(Address).value = ""
                    GoTo endline
                    End If
                End If
                
            End If
    Next
    GoTo endline
End If
End If

If Not Application.Intersect(Sheet20.Range("IFSC_80GB"), Sheet20.Range(Target.Address)) Is Nothing Then

Set IfsCode = Sheet20.Range("IFSC_80GB")

    
    If Not Application.Intersect(IfsCode, Range(Target.Address)) Is Nothing Then
         For Each rangecell In Sheet20.Range("IFSC_80GB")
            rangecell.value = UCase(rangecell.value)
     
     If rangecell.value <> "" Then
           
           If Not CheckIFSC(UCase(Target.value)) Then
                    fmsgbox ("* ""Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Sheet20.Unprotect Password:=getmsgstate
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
                    If (Sheet20.Range(Address).value = "") Then
                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                     Address = Replace(sourceAddress, "K", "J")
                     Sheet20.Range(Address).value = ""
                    GoTo endline
                    End If
                End If
                
            End If
    Next
    GoTo endline
End If
End If

If Not Application.Intersect(Sheet20.Range("IFSC_80GC"), Sheet20.Range(Target.Address)) Is Nothing Then

Set IfsCode = Sheet20.Range("IFSC_80GC")


    If Not Application.Intersect(IfsCode, Range(Target.Address)) Is Nothing Then
         For Each rangecell In Sheet20.Range("IFSC_80GC")
            rangecell.value = UCase(rangecell.value)
     
     If rangecell.value <> "" Then
           
           If Not CheckIFSC(UCase(Target.value)) Then
                    fmsgbox ("* ""Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Sheet20.Unprotect Password:=getmsgstate
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
                    If (Sheet20.Range(Address).value = "") Then
                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                     Address = Replace(sourceAddress, "K", "J")
                     Sheet20.Range(Address).value = ""
                    GoTo endline
                    End If
                End If
                
            End If
    Next
    GoTo endline
End If
End If

If Not Application.Intersect(Sheet20.Range("IFSC_80GD"), Sheet20.Range(Target.Address)) Is Nothing Then

Set IfsCode = Sheet20.Range("IFSC_80GD")
    
    If Not Application.Intersect(IfsCode, Range(Target.Address)) Is Nothing Then
         For Each rangecell In Sheet20.Range("IFSC_80GD")
            rangecell.value = UCase(rangecell.value)
     
     If rangecell.value <> "" Then
           
           If Not CheckIFSC(UCase(Target.value)) Then
                    fmsgbox ("* ""Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Sheet20.Unprotect Password:=getmsgstate
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
                    If (Sheet20.Range(Address).value = "") Then
                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                     Address = Replace(sourceAddress, "K", "J")
                     Sheet20.Range(Address).value = ""
                    GoTo endline
                    End If
                End If
                
            End If
    Next
    GoTo endline
End If
End If



 Dim var As Integer
    If Not checkfieldSuperSpecialcharacter(Target.value) Then
        fmsgbox ("* Please do not enter Special characters < , >" & Chr(13))
        Target.value = ""
    End If
'Change-27.01.2023.101.80G.5A



'    If target.Column = 10 Then
'        If Trim(Range("J" & target.row).value) = "AAAAR1077P" Then
'            If 1 < Application.CountIf(Range("Per10080G.DoneePAN"), Range("J" & target.row).value) Or 1 < Application.CountIf(Range("PerNO5080G.DoneePAN"), Range("J" & target.row).value) Or 1 < Application.CountIf(Range("PerYES10080G.DoneePAN"), Range("J" & target.row).value) Then 'Or 1 < Application.CountIf(Range("Per5080G.DoneePAN"), Range("J" & Target.row).Value)
'                 MsgBox "* Duplicate Pan AAAAR1077P is not allowed for same table"
'                 Range(target.Address).value = ""
'                 Range("K" & target.row).value = ""
'                 Range("L" & target.row).value = ""
'            End If
'        Else
'        'Change-ANK, 28.01.2023 ramya 70567
' If Application.Intersect(Sheet20.Range("Per5080G.DoneePAN"), Sheet20.Range(target.Address)) Is Nothing Then
'                var = Application.CountIf(Range("J1:J" & Rows.count), Range("J" & target.row).value)
'                    If 1 < var Then
'                            'MsgBox "* Duplicate Pan value is not allowed"
'
'                        'ramya 70567
''                             MsgBox "* Please enter aggregate amount of donation made in cash to same Donee once"
'
'                             MsgBox """Same PAN cannot be entered for different donation type"""
'
'
'                            Range(target.Address).value = ""
'                            Range("K" & target.row).value = ""
'                            'Range("L" & Target.row).value = ""
'                    End If
'            Else
'                var = Application.CountIf(Range("Per10080G.DoneePAN"), Range("J" & target.row).value) + Application.CountIf(Range("PerNO5080G.DoneePAN"), Range("J" & target.row).value) + Application.CountIf(Range("PerYES10080G.DoneePAN"), Range("J" & target.row).value)
'                If 1 <= var And Range("J" & target.row).value <> "" Then
'                    'MsgBox "* Duplicate Pan value is not allowed"
'                    'ramya 70567
''                             MsgBox "* Please enter aggregate amount of donation made in cash to same Donee once"
'
'                             MsgBox """Same PAN cannot be entered for different donation type"""
'
'
'
'                    Range(target.Address).value = ""
''                    Range("K" & Target.row).Value = ""
''                    Range("L" & Target.row).Value = ""
'                End If
'
'            End If
'        End If
'    End If

'Added by Naresh as per 0.2v

''==============================================================================
If Not Application.Intersect(Sheet20.Range("Per10080G.DoneePAN"), Sheet20.Range(Target.Address)) Is Nothing Then
       Target.value = UCase(Target.value)
 End If

 If Not Application.Intersect(Sheet20.Range("PerNO5080G.DoneePAN"), Sheet20.Range(Target.Address)) Is Nothing Then
       Target.value = UCase(Target.value)
 End If

 If Not Application.Intersect(Sheet20.Range("PerYES10080G.DoneePAN"), Sheet20.Range(Target.Address)) Is Nothing Then
       Target.value = UCase(Target.value)
 End If

 If Not Application.Intersect(Sheet20.Range("Per5080G.DoneePAN"), Sheet20.Range(Target.Address)) Is Nothing Then
       Target.value = UCase(Target.value)
 End If


'Dim var As Integer
Dim isPAN As Boolean
Dim isCash As Boolean

If Target.Column = 10 Or Target.Column = 12 Then
isPAN = _
    Not Application.Intersect(Target, Sheet20.Range("Per10080G.DoneePAN")) Is Nothing _
 Or Not Application.Intersect(Target, Sheet20.Range("PerNO5080G.DoneePAN")) Is Nothing _
 Or Not Application.Intersect(Target, Sheet20.Range("Per5080G.DoneePAN")) Is Nothing _
 Or Not Application.Intersect(Target, Sheet20.Range("PerYES10080G.DoneePAN")) Is Nothing

isCash = _
    Not Application.Intersect(Target, Sheet20.Range("Per10080G.DonationInCash")) Is Nothing _
 Or Not Application.Intersect(Target, Sheet20.Range("PerNO5080G.DonationInCash")) Is Nothing _
 Or Not Application.Intersect(Target, Sheet20.Range("PerYES10080G.DonationInCash")) Is Nothing _
 Or Not Application.Intersect(Target, Sheet20.Range("Per5080G.DonationInCash")) Is Nothing

        If Trim(Range("J" & Target.row).value) = "AAAAR1077P" Then



   ElseIf _
       isPAN Or isCash Then

        Dim rngCash As Range
        Dim rngPAN As Range
        Dim rngVal As Range
        Dim ranVal As Range
        Dim cntOnes As Long
        Dim cell As Range

        'Combine all cash blocks
        Set rngCash = Union( _
            Sheet20.Range("Per10080G.DonationInCash"), _
            Sheet20.Range("PerNO5080G.DonationInCash"), _
            Sheet20.Range("PerYES10080G.DonationInCash"), _
            Sheet20.Range("Per5080G.DonationInCash") _
        )

        Set rngPAN = Union( _
            Sheet20.Range("Per10080G.DoneePAN"), _
            Sheet20.Range("PerNO5080G.DoneePAN"), _
            Sheet20.Range("Per5080G.DoneePAN"), _
            Sheet20.Range("PerYES10080G.DoneePAN") _
        )

        If isPAN Then
            Set ranVal = rngPAN
        ElseIf isCash Then
            Set ranVal = rngCash
        End If


        Application.EnableEvents = False
        Application.Calculate

        'Column V
        cntOnes = Application.WorksheetFunction.CountIf(Sheet20.Range("V:V"), 1)

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

            MsgBox """Same PAN cannot be entered for ""Donation in cash"" in Schedule 80G"""

        End If



        Application.EnableEvents = True

End If
End If

Dim comb90Val As Long

If Target.Column = 10 Then   'PAN column

    If Trim(Range("J" & Target.row).value) <> "AAAAR1077P" Then
        comb90Val = Sheet20.Cells(Target.row, "X").value


            If Trim(Target.value) <> "" Then

                    If comb90Val > 0 Then

                        Application.EnableEvents = False

                        MsgBox """Same PAN cannot be entered in other block."""
                        Target.value = ""
                        Application.EnableEvents = True

                    End If

            End If

    End If
End If
''==============================================================================
    '
    
'Change-20.01.2023.101.80G.2A
    If Target.Column = 12 Then
        If Range("L" & Target.row).value <> "" And Range("L" & Target.row).value > 2000 Then
'        fmsgbox "* If Donation in cash is more than Rs. 2000, 80G Deduction is not allowed"
        fmsgbox "* ""Aggregate donation in cash against one PAN cannot exceed Rs. 2000"""
        Target.value = ""
        End If
    End If
    
'If Target.Column = 11 Then
'        If Range("K" & Target.row).value > 2000 Then
'        fmsgbox "* If Donation in cash is more than Rs. 2000, 80G Deduction is not allowed"
'        End If
'    End If
    
'---

'commented by sai 12/09/2024=======================================


'    If Target.Column = 9 Then
'           Dim PinCode As Range
'           Dim found As Variant
'            If (Target.value <> "") Then
'           found = False
'
'           For Each PinCode In SheetALL.Range("PincodeList")
'               Dim code As String
'               code = Left(CStr(Target.value), 3)
'               If CStr(PinCode.value) = CStr(code) Then
'                   Range("H" & Target.row).value = SheetALL.Range("CX" & PinCode.row).value
'                   found = True
'                   Exit For
'               End If
'            Next PinCode
'            If found = False Then
'            MsgBox "Enter correct Pincode"
'            Target.value = ""
'            End If
'            End If
'            End If
    
    '=====================================================================
    
 ' by Sai 12/09/2024
  If Target.Column = 8 Then

'by sai on 17/09/2024=======================================================================
   
   Dim PIN_targetadd, state_targetadd As String
      'state_targetadd = Target.address
      state_targetadd = Replace(Target.Address, "$", "")
      PIN_targetadd = Replace(state_targetadd, "H", "I")
      
      
    If state_Validation(PIN_targetadd, state_targetadd) = False Then Sheet20.Range(Replace(state_targetadd, "H", "I")).value = ""
    
   End If
   
  If Target.Column = 9 Then
            
            If (Target.value <> "") Then

  
     
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
      
'=========================================================================================

'Added by Aavula Naresh for AY 2026-27

If Not Intersect(Range("Per10080G.DonationInCash"), Range(Target.Address)) Is Nothing Then
Sheet20.Unprotect Password:=getmsgstate

           If Sheet20.Range(Target.Address).value > 0 Then
           Sheet20.Range(Target.Address).Interior.Color = "&HCCFFCC"
           
            
            Sheet20.Range(Replace(Target.Address, "L", "M")).ClearContents
            Sheet20.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HD8D8D8"
            Sheet20.Range(Replace(Target.Address, "L", "M")).Locked = True
            
            Sheet20.Range(Replace(Target.Address, "L", "P")).ClearContents
            Sheet20.Range(Replace(Target.Address, "L", "P")).Locked = True
            Sheet20.Range(Replace(Target.Address, "L", "P")).Interior.Color = "&HD8D8D8"
        
             
            Sheet20.Range(Replace(Target.Address, "L", "Q")).ClearContents
            Sheet20.Range(Replace(Target.Address, "L", "Q")).Locked = True
            Sheet20.Range(Replace(Target.Address, "L", "Q")).Interior.Color = "&HD8D8D8"
            
            Sheet20.Protect Password:=getmsgstate
        
        Else
                Sheet20.Unprotect Password:=getmsgstate
                
                Sheet20.Range(Replace(Target.Address, "L", "M")).Locked = False
                Sheet20.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HCCFFCC"
                
                Sheet20.Range(Replace(Target.Address, "L", "P")).Locked = False
                Sheet20.Range(Replace(Target.Address, "L", "P")).Interior.Color = "&HCCFFCC"
                
                Sheet20.Range(Replace(Target.Address, "L", "Q")).Locked = False
                Sheet20.Range(Replace(Target.Address, "L", "Q")).Interior.Color = "&HCCFFCC"
                
                 
Sheet20.Protect Password:=getmsgstate
    End If
End If


If Not Intersect(Range("PerNO5080G.DonationInCash"), Range(Target.Address)) Is Nothing Then
Sheet20.Unprotect Password:=getmsgstate

           If Sheet20.Range(Target.Address).value > 0 Then
           Sheet20.Range(Target.Address).Interior.Color = "&HCCFFCC"
           
            
           
            Sheet20.Range(Replace(Target.Address, "L", "M")).ClearContents
            
            Sheet20.Range(Replace(Target.Address, "L", "M")).Locked = True
            Sheet20.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HD8D8D8"
            
            
            Sheet20.Range(Replace(Target.Address, "L", "P")).ClearContents
            Sheet20.Range(Replace(Target.Address, "L", "P")).Locked = True
            Sheet20.Range(Replace(Target.Address, "L", "P")).Interior.Color = "&HD8D8D8"
            
             
            Sheet20.Range(Replace(Target.Address, "L", "Q")).ClearContents
            Sheet20.Range(Replace(Target.Address, "L", "Q")).Locked = True
            Sheet20.Range(Replace(Target.Address, "L", "Q")).Interior.Color = "&HD8D8D8"
        
        Else
                Sheet20.Unprotect Password:=getmsgstate
'                Sheet20.Range(Replace(Target.Address, "L", "M")).ClearContents
                Sheet20.Range(Replace(Target.Address, "L", "M")).Locked = False
                Sheet20.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HCCFFCC"
                
                Sheet20.Range(Replace(Target.Address, "L", "P")).Locked = False
                Sheet20.Range(Replace(Target.Address, "L", "P")).Interior.Color = "&HCCFFCC"
                
                Sheet20.Range(Replace(Target.Address, "L", "Q")).Locked = False
                Sheet20.Range(Replace(Target.Address, "L", "Q")).Interior.Color = "&HCCFFCC"
                
                 
Sheet20.Protect Password:=getmsgstate
    End If
End If

If Not Intersect(Range("PerYES10080G.DonationInCash"), Range(Target.Address)) Is Nothing Then
Sheet20.Unprotect Password:=getmsgstate

           If Sheet20.Range(Target.Address).value > 0 Then
           Sheet20.Range(Target.Address).Interior.Color = "&HCCFFCC"
           
            
           
            Sheet20.Range(Replace(Target.Address, "L", "M")).ClearContents
            Sheet20.Range(Replace(Target.Address, "L", "M")).Locked = True
            Sheet20.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HD8D8D8"
            
            
            Sheet20.Range(Replace(Target.Address, "L", "P")).ClearContents
            Sheet20.Range(Replace(Target.Address, "L", "P")).Locked = True
            Sheet20.Range(Replace(Target.Address, "L", "P")).Interior.Color = "&HD8D8D8"
            
             
            Sheet20.Range(Replace(Target.Address, "L", "Q")).ClearContents
            Sheet20.Range(Replace(Target.Address, "L", "Q")).Locked = True
            Sheet20.Range(Replace(Target.Address, "L", "Q")).Interior.Color = "&HD8D8D8"
        
        Else
                Sheet20.Unprotect Password:=getmsgstate
'                Sheet20.Range(Replace(Target.Address, "L", "M")).ClearContents
                Sheet20.Range(Replace(Target.Address, "L", "M")).Locked = False
                Sheet20.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HCCFFCC"
                
                Sheet20.Range(Replace(Target.Address, "L", "P")).Locked = False
                Sheet20.Range(Replace(Target.Address, "L", "P")).Interior.Color = "&HCCFFCC"
                
                Sheet20.Range(Replace(Target.Address, "L", "Q")).Locked = False
                Sheet20.Range(Replace(Target.Address, "L", "Q")).Interior.Color = "&HCCFFCC"
                
                 
Sheet20.Protect Password:=getmsgstate
    End If
End If

If Not Intersect(Range("Per5080G.DonationInCash"), Range(Target.Address)) Is Nothing Then
Sheet20.Unprotect Password:=getmsgstate

           If Sheet20.Range(Target.Address).value > 0 Then
           Sheet20.Range(Target.Address).Interior.Color = "&HCCFFCC"
           
            
           
            Sheet20.Range(Replace(Target.Address, "L", "M")).ClearContents
            Sheet20.Range(Replace(Target.Address, "L", "M")).Locked = True
            Sheet20.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HD8D8D8"
            
            
            Sheet20.Range(Replace(Target.Address, "L", "P")).ClearContents
            Sheet20.Range(Replace(Target.Address, "L", "P")).Locked = True
            Sheet20.Range(Replace(Target.Address, "L", "P")).Interior.Color = "&HD8D8D8"
            
             
            Sheet20.Range(Replace(Target.Address, "L", "Q")).ClearContents
            Sheet20.Range(Replace(Target.Address, "L", "Q")).Locked = True
            Sheet20.Range(Replace(Target.Address, "L", "Q")).Interior.Color = "&HD8D8D8"
        
        Else
                Sheet20.Unprotect Password:=getmsgstate
'                Sheet20.Range(Replace(Target.Address, "L", "M")).ClearContents
                Sheet20.Range(Replace(Target.Address, "L", "M")).Locked = False
                Sheet20.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HCCFFCC"
                
                Sheet20.Range(Replace(Target.Address, "L", "P")).Locked = False
                Sheet20.Range(Replace(Target.Address, "L", "P")).Interior.Color = "&HCCFFCC"
                
                Sheet20.Range(Replace(Target.Address, "L", "Q")).Locked = False
                Sheet20.Range(Replace(Target.Address, "L", "Q")).Interior.Color = "&HCCFFCC"
                
                 
Sheet20.Protect Password:=getmsgstate
    End If
End If

If Not Application.Intersect(Sheet20.Range("Per10080G.DonationAmt"), Sheet20.Range(Target.Address)) Is Nothing Then

If Target.value > 0 Then



  Sheet20.Unprotect Password:=getmsgstate
    'DONATION IN CASH
        Sheet20.Range(Replace(Target.Address, "M", "L")).ClearContents
        Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = True
        Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"

        
    Else
    
'DONATION IN CASH
        'Sheet20.Range(Replace(Target.Address, "M", "L")).ClearContents
        Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = False
        Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"
    
    End If
              
End If

If Not Application.Intersect(Sheet20.Range("PerNO5080G.DonationAmt"), Sheet20.Range(Target.Address)) Is Nothing Then

If Target.value > 0 Then



  Sheet20.Unprotect Password:=getmsgstate
    'DONATION IN CASH
        Sheet20.Range(Replace(Target.Address, "M", "L")).ClearContents
        Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = True
        Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"

        
    Else
    
'DONATION IN CASH
        'Sheet20.Range(Replace(Target.Address, "M", "L")).ClearContents
        Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = False
        Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"
    
    End If
              
End If

If Not Application.Intersect(Sheet20.Range("PerYES10080G.DonationAmt"), Sheet20.Range(Target.Address)) Is Nothing Then

If Target.value > 0 Then



  Sheet20.Unprotect Password:=getmsgstate
    'DONATION IN CASH
        Sheet20.Range(Replace(Target.Address, "M", "L")).ClearContents
        Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = True
        Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"

        
    Else
    
'DONATION IN CASH
        'Sheet20.Range(Replace(Target.Address, "M", "L")).ClearContents
        Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = False
        Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"
    
    End If
              
End If

If Not Application.Intersect(Sheet20.Range("Per5080G.DonationAmt"), Sheet20.Range(Target.Address)) Is Nothing Then

If Target.value > 0 Then



  Sheet20.Unprotect Password:=getmsgstate
    'DONATION IN CASH
        Sheet20.Range(Replace(Target.Address, "M", "L")).ClearContents
        Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = True
        Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"

        
    Else
    
'DONATION IN CASH
        'Sheet20.Range(Replace(Target.Address, "M", "L")).ClearContents
        Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = False
        Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"
    
    End If
              
End If

'If Not Intersect(Range("PerNO5080G.DonationAmt"), Range(Target.Address)) Is Nothing Then
'Sheet20.Unprotect Password:=getmsgstate
'        If Sheet20.Range(Target.Address).value = 0 Then
'
'            Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = False
'            Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"
'
'            Sheet20.Protect Password:=getmsgstate
'        Else
'        Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "M", "L")).ClearContents
'            Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = True
'            Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"
'            Sheet20.Protect Password:=getmsgstate
'
'    End If
'End If
'
'
'If Not Intersect(Range("PerYES10080G.DonationAmt"), Range(Target.Address)) Is Nothing Then
'Sheet20.Unprotect Password:=getmsgstate
'        If Sheet20.Range(Target.Address).value = 0 Then
'
'            Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = False
'            Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"
'
'            Sheet20.Protect Password:=getmsgstate
'        Else
'        Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "M", "L")).ClearContents
'            Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = True
'            Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"
'            Sheet20.Protect Password:=getmsgstate
'
'    End If
'End If
'
'
'If Not Intersect(Range("Per5080G.DonationAmt"), Range(Target.Address)) Is Nothing Then
'Sheet20.Unprotect Password:=getmsgstate
'        If Sheet20.Range(Target.Address).value = 0 Then
'
'            Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = False
'            Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"
'
'            Sheet20.Protect Password:=getmsgstate
'        Else
'        Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "M", "L")).ClearContents
'            Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = True
'            Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"
'            Sheet20.Protect Password:=getmsgstate
'
'    End If
'End If

'If Not Application.Intersect(Sheet20.Range("Per10080G.DonationAmt"), Sheet20.Range(Target.Address)) Is Nothing Then
'If Target.value > 0 Then
'  Sheet20.Unprotect Password:=getmsgstate
'    'DONATION IN CASH
'        Sheet20.Range(Replace(Target.Address, "M", "L")).ClearContents
'        Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = True
'        Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"
'Sheet20.Protect Password:=getmsgstate
'    Else
'      Sheet20.Unprotect Password:=getmsgstate
''DONATION IN CASH
'        Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = False
'        Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"
'
'    Sheet20.Protect Password:=getmsgstate
'    End If
'End If
'
''''''''''''''-==============================================''''''''''''''''''''''''
''If Not Application.Intersect(Sheet20.Range("Per10080G.DonationAmt"), Sheet20.Range(Target.Address)) Is Nothing Then
''If Target.value = 0 Then
'''Sheet20.Range("Per10080G.DonationAmt").value = ""
''  Sheet20.Unprotect Password:=getmsgstate
''
''
''            Sheet20.Range(Replace(Target.Address, "M", "P")).ClearContents
''            Sheet20.Range(Replace(Target.Address, "M", "P")).Locked = True
''            Sheet20.Range(Replace(Target.Address, "M", "P")).Interior.Color = "&HD8D8D8"
''
''            Sheet20.Range(Replace(Target.Address, "M", "Q")).ClearContents
''            Sheet20.Range(Replace(Target.Address, "M", "Q")).Locked = True
''            Sheet20.Range(Replace(Target.Address, "M", "Q")).Interior.Color = "&HD8D8D8"
''
''            Sheet20.Protect Password:=getmsgstate
''
''
''    End If
''End If
''If Not Application.Intersect(Sheet20.Range("Per10080G.DonationAmt"), Sheet20.Range(Target.Address)) Is Nothing Then
'' If Target.value = "" Then
''  Sheet20.Unprotect Password:=getmsgstate
''            Sheet20.Range(Replace(Target.Address, "M", "P")).Locked = False
''            Sheet20.Range(Replace(Target.Address, "M", "P")).Interior.Color = "&HCCFFCC"
''
''            Sheet20.Range(Replace(Target.Address, "M", "Q")).Locked = False
''            Sheet20.Range(Replace(Target.Address, "M", "Q")).Interior.Color = "&HCCFFCC"
''            Sheet20.Protect Password:=getmsgstate
'' End If
''End If
'
''''''=====================================================================
'
'
'If Not Application.Intersect(Sheet20.Range("PerNO5080G.DonationAmt"), Sheet20.Range(Target.Address)) Is Nothing Then
'
'If Target.value > 0 Then
'  Sheet20.Unprotect Password:=getmsgstate
'    'DONATION IN CASH
'        Sheet20.Range(Replace(Target.Address, "M", "L")).ClearContents
'        Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = True
'        Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"
'        Sheet20.Protect Password:=getmsgstate
'    Else
''DONATION IN CASH
'Sheet20.Unprotect Password:=getmsgstate
'        Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = False
'        Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"
'
'        Sheet20.Protect Password:=getmsgstate
'    End If
'End If
'
'
''''''''''''''-==============================================''''''''''''''''''''''''
''If Not Application.Intersect(Sheet20.Range("PerNO5080G.DonationAmt"), Sheet20.Range(Target.Address)) Is Nothing Then
''If Target.value = 0 Then
''  Sheet20.Unprotect Password:=getmsgstate
''            Sheet20.Range(Replace(Target.Address, "M", "P")).ClearContents
''            Sheet20.Range(Replace(Target.Address, "M", "P")).Locked = True
''            Sheet20.Range(Replace(Target.Address, "M", "P")).Interior.Color = "&HD8D8D8"
''
''            Sheet20.Range(Replace(Target.Address, "M", "Q")).ClearContents
''            Sheet20.Range(Replace(Target.Address, "M", "Q")).Locked = True
''            Sheet20.Range(Replace(Target.Address, "M", "Q")).Interior.Color = "&HD8D8D8"
''            Sheet20.Protect Password:=getmsgstate
''    End If
''End If
''If Not Application.Intersect(Sheet20.Range("PerNO5080G.DonationAmt"), Sheet20.Range(Target.Address)) Is Nothing Then
'' If Target.value = "" Then
''  Sheet20.Unprotect Password:=getmsgstate
''            Sheet20.Range(Replace(Target.Address, "M", "P")).Locked = False
''            Sheet20.Range(Replace(Target.Address, "M", "P")).Interior.Color = "&HCCFFCC"
''
''            Sheet20.Range(Replace(Target.Address, "M", "Q")).Locked = False
''            Sheet20.Range(Replace(Target.Address, "M", "Q")).Interior.Color = "&HCCFFCC"
''            Sheet20.Protect Password:=getmsgstate
'' End If
''End If
'
''''''=====================================================================
'
'If Not Application.Intersect(Sheet20.Range("PerYES10080G.DonationAmt"), Sheet20.Range(Target.Address)) Is Nothing Then
'
'If Target.value > 0 Then
'
'
'
'  Sheet20.Unprotect Password:=getmsgstate
'    'DONATION IN CASH
'        Sheet20.Range(Replace(Target.Address, "M", "L")).ClearContents
'        Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = True
'        Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"
'
'        Sheet20.Protect Password:=getmsgstate
'    Else
'
''DONATION IN CASH
'Sheet20.Unprotect Password:=getmsgstate
'        Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = False
'        Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"
'    Sheet20.Protect Password:=getmsgstate
'    End If
'
'End If
'
''''''''''''''-==============================================''''''''''''''''''''''''
''If Not Application.Intersect(Sheet20.Range("PerYES10080G.DonationAmt"), Sheet20.Range(Target.Address)) Is Nothing Then
''If Target.value = 0 Then
''  Sheet20.Unprotect Password:=getmsgstate
''            Sheet20.Range(Replace(Target.Address, "M", "P")).ClearContents
''            Sheet20.Range(Replace(Target.Address, "M", "P")).Locked = True
''            Sheet20.Range(Replace(Target.Address, "M", "P")).Interior.Color = "&HD8D8D8"
''
''            Sheet20.Range(Replace(Target.Address, "M", "Q")).ClearContents
''            Sheet20.Range(Replace(Target.Address, "M", "Q")).Locked = True
''            Sheet20.Range(Replace(Target.Address, "M", "Q")).Interior.Color = "&HD8D8D8"
''            Sheet20.Protect Password:=getmsgstate
''    End If
''End If
''If Not Application.Intersect(Sheet20.Range("PerYES10080G.DonationAmt"), Sheet20.Range(Target.Address)) Is Nothing Then
'' If Target.value = "" Then
''  Sheet20.Unprotect Password:=getmsgstate
''            Sheet20.Range(Replace(Target.Address, "M", "P")).Locked = False
''            Sheet20.Range(Replace(Target.Address, "M", "P")).Interior.Color = "&HCCFFCC"
''
''            Sheet20.Range(Replace(Target.Address, "M", "Q")).Locked = False
''            Sheet20.Range(Replace(Target.Address, "M", "Q")).Interior.Color = "&HCCFFCC"
''            Sheet20.Protect Password:=getmsgstate
'' End If
''End If
'
''''''=====================================================================
'
'If Not Application.Intersect(Sheet20.Range("Per5080G.DonationAmt"), Sheet20.Range(Target.Address)) Is Nothing Then
'
'If Target.value > 0 Then
'
'
'
'  Sheet20.Unprotect Password:=getmsgstate
'    'DONATION IN CASH
'        Sheet20.Range(Replace(Target.Address, "M", "L")).ClearContents
'        Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = True
'        Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"
'Sheet20.Protect Password:=getmsgstate
'
'    Else
'
''DONATION IN CASH
' Sheet20.Unprotect Password:=getmsgstate
'        Sheet20.Range(Replace(Target.Address, "M", "L")).Locked = False
'        Sheet20.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"
'    Sheet20.Protect Password:=getmsgstate
'    End If
'
'End If
    
    '''''''''''''-==============================================''''''''''''''''''''''''
'If Not Application.Intersect(Sheet20.Range("Per5080G.DonationAmt"), Sheet20.Range(Target.Address)) Is Nothing Then
'If Target.value = 0 Then
'  Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "M", "P")).ClearContents
'            Sheet20.Range(Replace(Target.Address, "M", "P")).Locked = True
'            Sheet20.Range(Replace(Target.Address, "M", "P")).Interior.Color = "&HD8D8D8"
'
'            Sheet20.Range(Replace(Target.Address, "M", "Q")).ClearContents
'            Sheet20.Range(Replace(Target.Address, "M", "Q")).Locked = True
'            Sheet20.Range(Replace(Target.Address, "M", "Q")).Interior.Color = "&HD8D8D8"
'            Sheet20.Protect Password:=getmsgstate
'    End If
'End If
'If Not Application.Intersect(Sheet20.Range("Per5080G.DonationAmt"), Sheet20.Range(Target.Address)) Is Nothing Then
' If Target.value = "" Then
'  Sheet20.Unprotect Password:=getmsgstate
'            Sheet20.Range(Replace(Target.Address, "M", "P")).Locked = False
'            Sheet20.Range(Replace(Target.Address, "M", "P")).Interior.Color = "&HCCFFCC"
'
'            Sheet20.Range(Replace(Target.Address, "M", "Q")).Locked = False
'            Sheet20.Range(Replace(Target.Address, "M", "Q")).Interior.Color = "&HCCFFCC"
'            Sheet20.Protect Password:=getmsgstate
' End If
'End If

'''''=====================================================================
    
        If Intersect(Target, Range("Per10080G.PinCode")) Then
            If Not (ValidatePinCode_Per10080G()) Then
                MsgBox "Invalid PinCode"
            End If
            GoTo endline
        End If

        If Intersect(Target, Range("PerNO5080G.PinCode")) Then
            If Not (ValidatePinCode_PerNO5080G()) Then
                MsgBox "Invalid PinCode"
            End If
            GoTo endline
        End If

        If (getRangeName(Target) = "Per5080G.DoneePAN") Then
            If Not (ValidateDoneePAN_Per5080G()) Then
                MsgBox "Invalid DoneePAN"
            End If
            GoTo endline
        End If

        If (getRangeName(Target) = "Per5080G.PinCode") Then
            If Not (ValidatePinCode_Per5080G()) Then
                MsgBox "Invalid PinCode"
            End If
            GoTo endline
        End If
endline:
Application.EnableEvents = True
End Sub

Private Sub Worksheet_Changeold1(ByVal Target As Range)
On Error GoTo endline

CalculateTaxCheckFlag = False

Application.EnableEvents = False
'Sheet19.DI_Hidden
Target.value = UCase(Target.value)


 Dim var As Integer
'    If Target.Column = 10 Then
'        If Not Range("K" & Target.Row).value = "" Then
'        var = Application.CountIf(Range("J1:J" & Rows.count), Range("J" & Target.Row).value)
'            If 1 < var And Range("K" & Target.Row).value <> "AAAAR1077P" Then
'            MsgBox "d2"
'                 MsgBox "* Duplicate Pan value is not allowed"
'                 Range(Target.Address).value = ""
'                 Range("K" & Target.Row).value = ""
'            End If
'        End If
'    End If

    If Target.Column = 10 Then
'        If Not Range("K" & Target.Row).value = "" Then
        var = Application.CountIf(Range("J1:J" & Rows.count), Range("J" & Target.row).value)
            If (1 < var) And (Target.value <> "AAAAR1077P") Then
            
                ' MsgBox "* Duplicate Pan value is not allowed"
                'ramya 70567
            MsgBox """Same PAN cannot be entered for different donation type"""
                 Range("J" & Target.row).value = ""
                 Range(Target.Address).value = ""
            End If
'        End If
    
    End If

    If Target.Column = 9 Then
           Dim PinCode As Range
           Dim found As Variant
            If (Target.value <> "") Then
           found = False
           
           For Each PinCode In SheetALL.Range("PincodeList")
               Dim code As String
               code = Left(CStr(Target.value), 3)
               If CStr(PinCode.value) = CStr(code) Then
                   Range("H" & Target.row).value = SheetALL.Range("CX" & PinCode.row).value
                   found = True
                   Exit For
               End If
            Next PinCode
            If found = False Then
            MsgBox "Enter correct Pincode"
            Target.value = ""
            End If
            End If
            End If
    
        If Intersect(Target, Range("Per10080G.PinCode")) Then
            If Not (ValidatePinCode_Per10080G()) Then
                MsgBox "Invalid PinCode"
            End If
            GoTo endline
        End If

        If Intersect(Target, Range("PerNO5080G.PinCode")) Then
            If Not (ValidatePinCode_PerNO5080G()) Then
                MsgBox "Invalid PinCode"
            End If
            GoTo endline
        End If

        If (getRangeName(Target) = "Per5080G.DoneePAN") Then
            If Not (ValidateDoneePAN_Per5080G()) Then
                MsgBox "Invalid DoneePAN"
            End If
            GoTo endline
        End If

        If (getRangeName(Target) = "Per5080G.PinCode") Then
            If Not (ValidatePinCode_Per5080G()) Then
                MsgBox "Invalid PinCode"
            End If
            GoTo endline
        End If
endline:
Application.EnableEvents = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error Resume Next
If Not Range(ActiveCell.Address).Locked = True Then
       If ActiveCell.Validation.Type = 6 Then
    Application.OnKey "%{DOWN}", ""
    'Sheet1.Range(targetaddress).Value = ""
ElseIf ActiveCell.Validation.Type = 1 Then
 Application.OnKey "%{DOWN}", ""
ElseIf ActiveCell.Validation.Type = 3 Then

    Application.OnKey "%{DOWN}"
End If
Else
Application.OnKey "%{DOWN}", ""
End If

End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    Cancel = True
End Sub

Function getRangeName(ByVal Target As Range) As String
    Dim start As Long
    start = Len(name)

    If InStr(1, Target.name.name, "'") > 0 Then
        start = start + 3 + 1
    Else
        start = start + 1 + 1
    End If

    If InStr(1, Target.name.name, ActiveSheet.name) = 0 Then
        start = 1
    End If
    getRangeName = Mid(Target.name.name, start)
End Function
Sub test()
Application.EnableEvents = True

End Sub
