Attribute VB_Name = "Sheet58"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Private Sub Worksheet_Change(ByVal Target As Range)

On Error GoTo endline
Application.EnableEvents = False


sPassword = EfilingCommon.getmsgstate
Sheet58.Unprotect Password:=sPassword

'added by Chetan C M on 26/06/2025 for AY 2025-26 based on V0.6 DE sheet changes

'start--
 
If Not Intersect(Range("Disability_80U"), Range(Target.Address)) Is Nothing Then

Sheet58.Unprotect Password:=getmsgstate
 
    If Sheet58.Range(Target.Address).value <> "(Select)" Then
 
         If Sheet58.Range(Target.Address).value <> "(i)autism, cerebral palsy, or multiple disabilities" Then

            If Sheet58.Range(Target.Address).value = "(ii) others" And Sheet58.Range(Replace(Target.Address, "E", "H")).value <> "" Then

                MsgBox """Type of disability shall be selected as ""(i)autism, cerebral palsy, or multiple disabilities"" when acknowledgement no. of form 10IA is provided"""

                Sheet58.Range(Target.Address).value = "(Select)"

            Else:

                Sheet58.Range(Replace(Target.Address, "E", "H")).ClearContents

                Sheet58.Range(Replace(Target.Address, "E", "H")).Locked = True

                Sheet58.Range(Replace(Target.Address, "E", "H")).Interior.Color = "&HD8D8D8"

            End If

        Else
 
            Sheet58.Unprotect Password:=getmsgstate

            Sheet58.Range(Replace(Target.Address, "E", "H")).Locked = False

            Sheet58.Range(Replace(Target.Address, "E", "H")).Interior.Color = "&HCCFFCC"

            Sheet58.Protect Password:=getmsgstate
 
        End If
 
    Else

            Sheet58.Unprotect Password:=getmsgstate

            Sheet58.Range(Replace(Target.Address, "E", "H")).Locked = False

            Sheet58.Range(Replace(Target.Address, "E", "H")).Interior.Color = "&HCCFFCC"
 
    End If
 
End If
 
 
If Not Intersect(Range("Disability_80DD"), Range(Target.Address)) Is Nothing Then

Sheet58.Unprotect Password:=getmsgstate
 
    If Sheet58.Range(Target.Address).value <> "(Select)" Then
 
        If Sheet58.Range(Target.Address).value <> "(i)autism, cerebral palsy, or multiple disabilities" Then
 
         If Sheet58.Range(Target.Address).value = "(ii) others" And Sheet58.Range(Replace(Target.Address, "E", "K")).value <> "" Then

            MsgBox """Type of disability shall be selected as ""(i)autism, cerebral palsy, or multiple disabilities"" when acknowledgement no. of form 10IA is provided"""

            Sheet58.Range(Target.Address).value = "(Select)"

         Else:

             Sheet58.Range(Replace(Target.Address, "E", "K")).ClearContents

             Sheet58.Range(Replace(Target.Address, "E", "K")).Locked = True

             Sheet58.Range(Replace(Target.Address, "E", "K")).Interior.Color = "&HD8D8D8"

             End If
 
 
        Else

             Sheet58.Unprotect Password:=getmsgstate

             Sheet58.Range(Replace(Target.Address, "E", "K")).Locked = False

             Sheet58.Range(Replace(Target.Address, "E", "K")).Interior.Color = "&HCCFFCC"
 
 
    End If

    Else

             Sheet58.Unprotect Password:=getmsgstate

             Sheet58.Range(Replace(Target.Address, "E", "K")).Locked = False

             Sheet58.Range(Replace(Target.Address, "E", "K")).Interior.Color = "&HCCFFCC"
 
    End If

End If
 
 
'--end

Dim dob_80U As Variant
dob_80U = Range("DatefilingFm10IA_80U").value



If Not Intersect(Range("DatefilingFm10IA_80U"), Range(Target.Address)) Is Nothing Then
             If Trim(Range("DatefilingFm10IA_80U").value) <> "" Then
               If Not FormatNCheckDate(dob_80U) Then
                        fmsgboxStatus ("Date of filing of Form 10IA in Schedule 80U  must be a valid dd/mm/yyyy format")
                        Range("DatefilingFm10IA_80U").Select
                        Range("DatefilingFm10IA_80U").value = ""
                    End If
                If Dformat(Trim(Range("DatefilingFm10IA_80U").value), "yyyy-mm-dd") > Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") Then
                          fmsgboxStatus "* Date of filing of Form 10IA can not be more than system date in Schedule 80U"
                          Range("DatefilingFm10IA_80U").Select
                        Range("DatefilingFm10IA_80U").value = ""
                 End If
              End If
End If

''added by Chetan C M for AY 2025-26
''start
'If Not Intersect(Range("Disability_80U"), Range(Target.Address)) Is Nothing Then
'
''    If Sheet58.Range("Disability_80U").value = "(i)autism, cerebral palsy, or multiple disabilities" Then
''        Sheet58.Unprotect Password:=sPassword
''        Sheet58.Range("AcknowledgeNum11A2_80U").Locked = True
''        Sheet58.Range("AcknowledgeNum11A2_80U").Interior.Color = (&HD8D8D8)
''        Sheet58.Protect Password:=sPassword
''    Else
''        Sheet58.Unprotect Password:=sPassword
''        Sheet58.Range("AcknowledgeNum11A2_80U").Locked = False
''        Sheet58.Range("AcknowledgeNum11A2_80U").Interior.Color = (&HCCFFCC)
''        Sheet58.Protect Password:=sPassword
''    End If
'
'    If Sheet58.Range("Disability_80U").value = "(ii) others" Then
'        Sheet58.Unprotect Password:=sPassword
'        Sheet58.Range("DatefilingFm10IA_80U").Locked = True
'        Sheet58.Range("DatefilingFm10IA_80U").Interior.Color = (&HD8D8D8)
'        Sheet58.Range("AckNoFm10IAfiled_80U").Locked = True
'        Sheet58.Range("AckNoFm10IAfiled_80U").Interior.Color = (&HD8D8D8)
'        Sheet58.Protect Password:=sPassword
'
'    Else
'
'        Sheet58.Unprotect Password:=sPassword
'        Sheet58.Range("DatefilingFm10IA_80U").Locked = False
'        Sheet58.Range("DatefilingFm10IA_80U").Interior.Color = (&HCCFFCC)
'        Sheet58.Range("AckNoFm10IAfiled_80U").Locked = False
'        Sheet58.Range("AckNoFm10IAfiled_80U").Interior.Color = (&HCCFFCC)
'        Sheet58.Protect Password:=sPassword
'
'    End If
'
'End If
'--end
'Added by Aavula Naresh SIT-93651
If Not Intersect(Range("Disability_80U"), Range(Target.Address)) Is Nothing Then
Sheet58.Unprotect Password:=getmsgstate
        'If Sheet58.Range(Target.Address).value = "" Or Sheet58.Range(Target.Address).value = "(Select)" Or Sheet58.Range(Target.Address).value = "(i)autism, cerebral palsy, or multiple disabilities" Then   'sit_93967
         
         If Sheet58.Range(Target.Address).value = "" Or Sheet58.Range(Target.Address).value = "(Select)" Or Sheet58.Range(Target.Address).value = "(i)autism, cerebral palsy, or multiple disabilities" Then
            
            Sheet58.Range(Replace(Target.Address, "E", "G")).Locked = False
            Sheet58.Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HCCFFCC"
            
            Sheet58.Range(Replace(Target.Address, "E", "H")).Locked = False
            Sheet58.Range(Replace(Target.Address, "E", "H")).Interior.Color = "&HCCFFCC"
            Sheet58.Protect Password:=getmsgstate
        Else
            Sheet58.Unprotect Password:=getmsgstate
            Sheet58.Range(Replace(Target.Address, "E", "G")).ClearContents
            Sheet58.Range(Replace(Target.Address, "E", "G")).Locked = True
            Sheet58.Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HD8D8D8"
                        
            Sheet58.Range(Replace(Target.Address, "E", "H")).ClearContents
            Sheet58.Range(Replace(Target.Address, "E", "H")).Locked = True
            Sheet58.Range(Replace(Target.Address, "E", "H")).Interior.Color = "&HD8D8D8"
            Sheet58.Protect Password:=getmsgstate

    End If
End If

'commented by Chetan C M on 26/06/2025 for AY 2025-26 based on V0.6 DE sheet changes
'start--

''Ack Number
'If Not Intersect(Range("AckNoFm10IAfiled_80U"), Range(Target.Address)) Is Nothing Then
'    If Not checkallfieldspecialcharacter_80U_80DD(Range("AckNoFm10IAfiled_80U").value) Then
'     fmsgboxStatus "* Ack. No. of Form 10IA filed cannot contain special characters in Schedule 80U."
'     Range("AckNoFm10IAfiled_80U").Select
'     Range("AckNoFm10IAfiled_80U").value = ""
'
'    End If
'End If

'--end


'UDID NUMBER

If Not Intersect(Range("UDIDNum_80U"), Range(Target.Address)) Is Nothing Then
    If Not checkallfieldspecialcharacter_80U_80DD(Range("UDIDNum_80U").value) Then
     fmsgboxStatus "* UDID number cannot contain special characters in Schedule 80U."
     Range("UDIDNum_80U").Select
     Range("UDIDNum_80U").value = ""
    
    End If
End If


'-------------80DD

''added by Chetan C M for AY 2025-26
''start
'If Not Intersect(Range("Disability_80DD"), Range(Target.Address)) Is Nothing Then
'
''    If Sheet58.Range("Disability_80DD").value = "(i)autism, cerebral palsy, or multiple disabilities" Then
''        Sheet58.Unprotect Password:=sPassword
''        Sheet58.Range("AcknowledgeNum11A2_80DD").Locked = True
''        Sheet58.Range("AcknowledgeNum11A2_80DD").Interior.Color = (&HD8D8D8)
''        Sheet58.Protect Password:=sPassword
''    Else
''        Sheet58.Unprotect Password:=sPassword
''        Sheet58.Range("AcknowledgeNum11A2_80DD").Locked = False
''        Sheet58.Range("AcknowledgeNum11A2_80DD").Interior.Color = (&HCCFFCC)
''        Sheet58.Protect Password:=sPassword
''    End If
'
'    If Sheet58.Range("Disability_80DD").value = "(ii) others" Then
'        Sheet58.Unprotect Password:=sPassword
'        Sheet58.Range("DatefilingFm10IA_80DD").Locked = True
'        Sheet58.Range("DatefilingFm10IA_80DD").Interior.Color = (&HD8D8D8)
'        Sheet58.Range("AckNoFm10IAfiled_80DD").Locked = True
'        Sheet58.Range("AckNoFm10IAfiled_80DD").Interior.Color = (&HD8D8D8)
'        Sheet58.Protect Password:=sPassword
'
'    Else
'
'        Sheet58.Unprotect Password:=sPassword
'        Sheet58.Range("DatefilingFm10IA_80DD").Locked = False
'        Sheet58.Range("DatefilingFm10IA_80DD").Interior.Color = (&HCCFFCC)
'        Sheet58.Range("AckNoFm10IAfiled_80DD").Locked = False
'        Sheet58.Range("AckNoFm10IAfiled_80DD").Interior.Color = (&HCCFFCC)
'        Sheet58.Protect Password:=sPassword
'
'    End If
'
'End If
''--end
'Added by Aavula Naresh SIT-93651
If Not Intersect(Range("Disability_80DD"), Range(Target.Address)) Is Nothing Then
Sheet58.Unprotect Password:=getmsgstate
        If Sheet58.Range(Target.Address).value = "" Or Sheet58.Range(Target.Address).value = "(Select)" Or Sheet58.Range(Target.Address).value = "(i)autism, cerebral palsy, or multiple disabilities" Then
            
            Sheet58.Range(Replace(Target.Address, "E", "J")).Locked = False
            Sheet58.Range(Replace(Target.Address, "E", "J")).Interior.Color = "&HCCFFCC"
            
            Sheet58.Range(Replace(Target.Address, "E", "K")).Locked = False
            Sheet58.Range(Replace(Target.Address, "E", "K")).Interior.Color = "&HCCFFCC"
            Sheet58.Protect Password:=getmsgstate
        Else
            Sheet58.Unprotect Password:=getmsgstate
            Sheet58.Range(Replace(Target.Address, "E", "J")).ClearContents
            Sheet58.Range(Replace(Target.Address, "E", "J")).Locked = True
            Sheet58.Range(Replace(Target.Address, "E", "J")).Interior.Color = "&HD8D8D8"
                        
            Sheet58.Range(Replace(Target.Address, "E", "K")).ClearContents
            Sheet58.Range(Replace(Target.Address, "E", "K")).Locked = True
            Sheet58.Range(Replace(Target.Address, "E", "K")).Interior.Color = "&HD8D8D8"
            Sheet58.Protect Password:=getmsgstate

    End If
End If

Dim dob_80dd As Variant
dob_80dd = Range("DatefilingFm10IA_80DD").value
If Not Intersect(Range("DatefilingFm10IA_80DD"), Range(Target.Address)) Is Nothing Then
'              If Trim(Range("DatefilingFm10IA_80DD").value) <> "" Or Trim(Range("DatefilingFm10IA_80DD").value) <> "00/00/0000" Then
              
              'Added by Aavula on 04/03/2025 SIT-89174
              
              If Trim(Range("DatefilingFm10IA_80DD").value) <> "" Then
                    If Not FormatNCheckDate(dob_80dd) Then
                        fmsgboxStatus ("Date of filing of Form 10IA in Schedule 80DD  must be a valid dd/mm/yyyy format")
                        Range("DatefilingFm10IA_80DD").Select
                        Range("DatefilingFm10IA_80DD").value = ""
                    End If
        
                  If Dformat(Trim(Range("DatefilingFm10IA_80DD").value), "yyyy-mm-dd") > Dformat(Trim(Sheet52.Range("sheet9.Date").value), "yyyy-mm-dd") Then
                        fmsgboxStatus "* Date of filing of Form 10IA can not be more than system date in Schedule 80DD"
                        Range("DatefilingFm10IA_80DD").Select
                        Range("DatefilingFm10IA_80DD").value = ""
                 End If
              End If
End If


If Not Intersect(Range("Aadhaardependent_80DD"), Range(Target.Address)) Is Nothing Then
              If Trim(Range("Aadhaardependent_80DD").value) <> "" Then
              If UCase(Trim(Range("Aadhaardependent_80DD").value)) = UCase(Trim(Sheet1.Range("sheet1.adhaarno").value)) Then
                 fmsgboxStatus "*Aadhaar of the dependent cannot be same as assessee Aadhaar  in Part-A General Information."
                       Range("Aadhaardependent_80DD").Select
                        Range("Aadhaardependent_80DD").value = ""
              End If
              End If
End If

If Not Intersect(Range("PANdependent_80DD"), Range(Target.Address)) Is Nothing Then
              If Trim(Range("PANdependent_80DD").value) <> "" Then
              If UCase(Trim(Range("PANdependent_80DD").value)) = UCase(Trim(Sheet1.Range("sheet1.PAN").value)) Then
                 fmsgboxStatus "*PAN of the dependent cannot be same as assessee PAN in Part-A General Information."
                       Range("PANdependent_80DD").Select
                        Range("PANdependent_80DD").value = ""
              End If
              End If
              If Not Part_A_General.CheckPAB_80DD(UCase(Trim(Target.value))) Then
                 fmsgboxStatus "*Invalid PAN in Schedule 80DD"
                 Range("PANdependent_80DD").Select
                  Range("PANdependent_80DD").value = ""
            Else
                Range("PANdependent_80DD").value = UCase(Trim(Range("PANdependent_80DD").value))
                 
              End If
              
              
End If

'commented by Chetan C M on 26/06/2025 for AY 2025-26 based on V0.6 DE sheet changes
'start--

''Ack Number
'If Not Intersect(Range("AckNoFm10IAfiled_80DD"), Range(Target.Address)) Is Nothing Then
'    If Not checkallfieldspecialcharacter_80U_80DD(Range("AckNoFm10IAfiled_80DD").value) Then
'     fmsgboxStatus "* Ack. No. of Form 10IA filed cannot contain special characters in Schedule 80DD."
'     Range("AckNoFm10IAfiled_80DD").Select
'     Range("AckNoFm10IAfiled_80DD").value = ""
'
'    End If
'End If

'--end


'UDID NUMBER

If Not Intersect(Range("UDIDNum_80DD"), Range(Target.Address)) Is Nothing Then
    If Not checkallfieldspecialcharacter_80U_80DD(Range("UDIDNum_80DD").value) Then
     fmsgboxStatus "* UDID number cannot contain special characters in Schedule 80DD."
     Range("UDIDNum_80DD").Select
     Range("UDIDNum_80DD").value = ""
    
    End If
End If

'Aadhar
If Not Intersect(Range("Aadhaardependent_80DD"), Range(Target.Address)) Is Nothing Then
  If ChkAdhaarNo_80DD = False Then
    fmsgboxStatus "*Invalid Aadhaar in Schedule 80DD"
    Range("Aadhaardependent_80DD").Select
    Range("Aadhaardependent_80DD").value = ""
  End If
End If
'Added by Aavula Naresh for AY 2025-26
If Not Intersect(Range("AckNoFm10IAfiled_80U"), Range(Target.Address)) Is Nothing Then
           If Trim(Range("AckNoFm10IAfiled_80U").value) <> "" Then
              If UCase(Trim(Range("AckNoFm10IAfiled_80U").value)) = UCase(Trim(Sheet58.Range("AckNoFm10IAfiled_80DD").value)) Then
'                 fmsgboxStatus "* Acknowledgement number entered in schedule 80U and 80DD cant be same."
                fmsgboxStatus "* Acknowledgement number of Form 10IA filed for self and Dependent can't be same. Please provide proper acknowledgement number"  'added by Chetan C M on 26/06/2025 for AY 2025-26 based on V0.6 DE sheet changes
                    Range("AckNoFm10IAfiled_80U").Select
                    Range("AckNoFm10IAfiled_80U").value = ""
              End If
           End If
 End If

 '80DD Ack No
  If Not Intersect(Range("AckNoFm10IAfiled_80DD"), Range(Target.Address)) Is Nothing Then
           If Trim(Range("AckNoFm10IAfiled_80DD").value) <> "" Then
              If UCase(Trim(Range("AckNoFm10IAfiled_80DD").value)) = UCase(Trim(Sheet58.Range("AckNoFm10IAfiled_80U").value)) Then
'                 fmsgboxStatus "* Acknowledgement number entered in schedule 80U and 80DD cant be same."
                fmsgboxStatus "* Acknowledgement number of Form 10IA filed for self and Dependent can't be same. Please provide proper acknowledgement number"  'added by Chetan C M on 26/06/2025 for AY 2025-26 based on V0.6 DE sheet changes
                    Range("AckNoFm10IAfiled_80DD").Select
                    Range("AckNoFm10IAfiled_80DD").value = ""
              End If
           End If
 End If



 


sPassword = EfilingCommon.getmsgstate
Sheet58.Protect Password:=sPassword
endline:
Application.EnableEvents = True

End Sub
 
