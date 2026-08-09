Attribute VB_Name = "Sheet59"
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
Sheet59.Unprotect Password:=sPassword

'Added by Aavula Naresh for AY 2025-26

If Not Intersect(Range("Disability_80DD"), Range(Target.Address)) Is Nothing Then
   Sheet59.Unprotect Password:=getmsgstate
        If Sheet59.Range(Target.Address).value = "" Or Sheet59.Range(Target.Address).value = "(Select)" Or Sheet59.Range(Target.Address).value = "(i) autism, cerebral palsy, or multiple disabilities" Then
            
            
            Sheet59.Range(Replace(Target.Address, "D", "I")).Locked = False
            Sheet59.Range(Replace(Target.Address, "D", "I")).Interior.Color = "&HCCFFCC"
            
            Sheet59.Range(Replace(Target.Address, "D", "J")).Locked = False
            Sheet59.Range(Replace(Target.Address, "D", "J")).Interior.Color = "&HCCFFCC"
            
            Sheet59.Protect Password:=getmsgstate
        
        Else
        
            Sheet59.Unprotect Password:=getmsgstate

            Sheet59.Range(Replace(Target.Address, "D", "I")).ClearContents
            Sheet59.Range(Replace(Target.Address, "D", "I")).Locked = True
            Sheet59.Range(Replace(Target.Address, "D", "I")).Interior.Color = "&HD8D8D8"
            
            Sheet59.Range(Replace(Target.Address, "D", "J")).ClearContents
            Sheet59.Range(Replace(Target.Address, "D", "J")).Locked = True
            Sheet59.Range(Replace(Target.Address, "D", "J")).Interior.Color = "&HD8D8D8"
            
            Sheet59.Protect Password:=getmsgstate
    End If
End If


If Not Intersect(Range("Disability_80U"), Range(Target.Address)) Is Nothing Then
Sheet59.Unprotect Password:=getmsgstate
        If Sheet59.Range(Target.Address).value = "" Or Sheet59.Range(Target.Address).value = "(Select)" Or Sheet59.Range(Target.Address).value = "(i) autism, cerebral palsy, or multiple disabilities" Then
            
            Sheet59.Range(Replace(Target.Address, "D", "F")).Locked = False
            Sheet59.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HCCFFCC"
            
            Sheet59.Range(Replace(Target.Address, "D", "G")).Locked = False
            Sheet59.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HCCFFCC"
            
            Sheet59.Protect Password:=getmsgstate
        Else
            Sheet59.Unprotect Password:=getmsgstate
            
            Sheet59.Range(Replace(Target.Address, "D", "F")).ClearContents
            Sheet59.Range(Replace(Target.Address, "D", "F")).Locked = True
            Sheet59.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HD8D8D8"
                        
            Sheet59.Range(Replace(Target.Address, "D", "G")).ClearContents
            Sheet59.Range(Replace(Target.Address, "D", "G")).Locked = True
            Sheet59.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HD8D8D8"
            
            Sheet59.Protect Password:=getmsgstate

    End If
End If
'''''''''''
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


'Ack Number
If Not Intersect(Range("AckNoFm10IAfiled_80U"), Range(Target.Address)) Is Nothing Then
    If Not checkallfieldspecialcharacter_80U_80DD(Range("AckNoFm10IAfiled_80U").value) Then
     fmsgboxStatus "* Ack. No. of Form 10IA filed cannot contain special characters in Schedule 80U."
     Range("AckNoFm10IAfiled_80U").Select
     Range("AckNoFm10IAfiled_80U").value = ""
    
    End If
End If

'UDID NUMBER

If Not Intersect(Range("UDIDNum_80U"), Range(Target.Address)) Is Nothing Then
    If Not checkallfieldspecialcharacter_80U_80DD(Range("UDIDNum_80U").value) Then
     fmsgboxStatus "* UDID number cannot contain special characters in Schedule 80U."
     Range("UDIDNum_80U").Select
     Range("UDIDNum_80U").value = ""
    
    End If
End If


'-------------80DD

Dim dob_80dd As Variant
dob_80dd = Range("DatefilingFm10IA_80DD").value
If Not Intersect(Range("DatefilingFm10IA_80DD"), Range(Target.Address)) Is Nothing Then
              If Trim(Range("DatefilingFm10IA_80DD").value) <> "" Or Trim(Range("DatefilingFm10IA_80DD").value) <> "00/00/0000" Then
                  
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
              
             ' If Trim(Range("PANdependent_80DD").value) <> "" Then

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


'Ack Number
If Not Intersect(Range("AckNoFm10IAfiled_80DD"), Range(Target.Address)) Is Nothing Then
    If Not checkallfieldspecialcharacter_80U_80DD(Range("AckNoFm10IAfiled_80DD").value) Then
     fmsgboxStatus "* Ack. No. of Form 10IA filed cannot contain special characters in Schedule 80DD."
     Range("AckNoFm10IAfiled_80DD").Select
     Range("AckNoFm10IAfiled_80DD").value = ""
    
    End If
End If

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

'80U Ack No
 If Not Intersect(Range("AckNoFm10IAfiled_80U"), Range(Target.Address)) Is Nothing Then
           If Trim(Range("AckNoFm10IAfiled_80U").value) <> "" Then
              If UCase(Trim(Range("AckNoFm10IAfiled_80U").value)) = UCase(Trim(Sheet59.Range("AckNoFm10IAfiled_80DD").value)) Then
                 fmsgboxStatus "* Acknowledgement number entered in schedule 80U and 80DD cant be same."
                    Range("AckNoFm10IAfiled_80U").Select
                    Range("AckNoFm10IAfiled_80U").value = ""
              End If
           End If
 End If
 
 '80DD Ack No
  If Not Intersect(Range("AckNoFm10IAfiled_80DD"), Range(Target.Address)) Is Nothing Then
           If Trim(Range("AckNoFm10IAfiled_80DD").value) <> "" Then
              If UCase(Trim(Range("AckNoFm10IAfiled_80DD").value)) = UCase(Trim(Sheet59.Range("AckNoFm10IAfiled_80U").value)) Then
                 fmsgboxStatus "* Acknowledgement number entered in schedule 80U and 80DD cant be same."
                    Range("AckNoFm10IAfiled_80DD").Select
                    Range("AckNoFm10IAfiled_80DD").value = ""
              End If
           End If
 End If
 
 
'Added by Aavula Naresh 0.6V
 
    If Not Intersect(Range("Disability_80U"), Range(Target.Address)) Is Nothing Then

        If Trim(Range("DatefilingFm10IA_80U").value) = "" And Range("Disability_80U").value = "(ii) others" Then
            fmsgboxStatus """Type of disability shall be selected as “(i) autism, cerebral palsy, or multiple disabilities” when date of form 10IA is provided""."        ', vbOKOnly, mdCommon.vbMessageCaption
        End If
    End If
    
    If Not Intersect(Range("Disability_80U"), Range(Target.Address)) Is Nothing Then

        If Trim(Range("AckNoFm10IAfiled_80U").value) = "" And Range("Disability_80U").value = "(ii) others" Then
            fmsgboxStatus """Type of disability shall be selected as “(i) autism, cerebral palsy, or multiple disabilities” when acknowledgement no. of form 10IA is provided""."        ', vbOKOnly, mdCommon.vbMessageCaption
        End If
    End If
    
    ''''''''''''''
    
    If Not Intersect(Range("Disability_80DD"), Range(Target.Address)) Is Nothing Then

        If Trim(Range("DatefilingFm10IA_80DD").value) = "" And Range("Disability_80DD").value = "(ii) others" Then
            fmsgboxStatus """Type of disability shall be selected as “(i) autism, cerebral palsy, or multiple disabilities” when date of form 10IA is provided""."        ', vbOKOnly, mdCommon.vbMessageCaption
        End If
    End If
    
    If Not Intersect(Range("Disability_80DD"), Range(Target.Address)) Is Nothing Then

        If Trim(Range("AckNoFm10IAfiled_80DD").value) = "" And Range("Disability_80DD").value = "(ii) others" Then
            fmsgboxStatus """Type of disability shall be selected as “(i) autism, cerebral palsy, or multiple disabilities” when acknowledgement no. of form 10IA is provided""."        ', vbOKOnly, mdCommon.vbMessageCaption
        End If
    End If
    
''''''''''''''''''''''''''''
 

sPassword = EfilingCommon.getmsgstate
Sheet59.Protect Password:=sPassword
endline:
Application.EnableEvents = True

End Sub
 



