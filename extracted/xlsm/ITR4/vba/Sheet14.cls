Attribute VB_Name = "Sheet14"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Activate()
 Sheet14.Protect Password:=getmsgstate

End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
 
On Error GoTo endline
Application.EnableEvents = False

Sheet14.Unprotect Password:=getmsgstate

'Ayush_20

If Not Intersect(Range("Disability_80DD"), Range(Target.Address)) Is Nothing Then
Sheet14.Unprotect Password:=getmsgstate
        'If Sheet14.Range(Target.Address).Value = "(viii) autism" Then
          If Sheet14.Range(Target.Address).value <> "(Select)" Then
'        If Sheet14.Range(Target.Address).Value <> "(viii) autism" And _
'        Sheet14.Range(Target.Address).Value <> "(ix) cerebral palsy" And _
'        Sheet14.Range(Target.Address).Value <> "(x) multiple disability" Then
'        If Sheet14.Range(Target.Address).Value <> "(i) autism, cerebral palsy, or multiple disabilities and" Then

        If Sheet14.Range(Target.Address).value <> "(i) autism, cerebral palsy, or multiple disabilities" Then  'Ankita_14/05/2025
             
         If Sheet14.Range(Target.Address).value = "(ii) others" And Sheet14.Range(Replace(Target.Address, "D", "I")).value <> "" Then
            MsgBox """Type of disability shall be selected as ""(i) autism, cerebral palsy, or multiple disabilities"" when acknowledgement no. of form 10IA is provided"""
            Sheet14.Range(Target.Address).value = "(Select)"
         Else:
             Sheet14.Range(Replace(Target.Address, "D", "I")).ClearContents
             Sheet14.Range(Replace(Target.Address, "D", "I")).Locked = True
             Sheet14.Range(Replace(Target.Address, "D", "I")).Interior.Color = "&HD8D8D8"
             End If
'            Sheet14.Range(Replace(Target.Address, "D", "J")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "J")).Locked = True
'            Sheet14.Range(Replace(Target.Address, "D", "J")).Interior.Color = "&HD8D8D8"
'            Sheet14.Range(Replace(Target.Address, "D", "K")).Locked = False
'            Sheet14.Range(Replace(Target.Address, "D", "K")).Interior.Color = "&HCCFFCC"

        Else
             Sheet14.Unprotect Password:=getmsgstate
'            Sheet14.Range(Replace(Target.Address, "D", "I")).ClearContents      'Ankita_28/05/2025
             Sheet14.Range(Replace(Target.Address, "D", "I")).Locked = False
             Sheet14.Range(Replace(Target.Address, "D", "I")).Interior.Color = "&HCCFFCC"
'            Sheet14.Range(Replace(Target.Address, "D", "J")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "J")).Locked = False
'            Sheet14.Range(Replace(Target.Address, "D", "J")).Interior.Color = "&HCCFFCC"
'            Sheet14.Range(Replace(Target.Address, "D", "K")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "K")).Locked = True
'            Sheet14.Range(Replace(Target.Address, "D", "K")).Interior.Color = "&HD8D8D8"

 
    End If
    Else
             Sheet14.Unprotect Password:=getmsgstate
'            Sheet14.Range(Replace(Target.Address, "D", "I")).ClearContents     'Ankita_28/05/2025
             Sheet14.Range(Replace(Target.Address, "D", "I")).Locked = False
             Sheet14.Range(Replace(Target.Address, "D", "I")).Interior.Color = "&HCCFFCC"
'            Sheet14.Range(Replace(Target.Address, "D", "J")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "J")).Locked = False
'            Sheet14.Range(Replace(Target.Address, "D", "J")).Interior.Color = "&HCCFFCC"
'
'            Sheet14.Range(Replace(Target.Address, "D", "K")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "K")).Locked = False
'            Sheet14.Range(Replace(Target.Address, "D", "K")).Interior.Color = "&HCCFFCC"
    End If
End If




'80U------------------------
'Malli---------23/04/2025
If Not Intersect(Range("Disability_80U"), Range(Target.Address)) Is Nothing Then
Sheet14.Unprotect Password:=getmsgstate
        'If Sheet14.Range(Target.Address).Value = "(viii) autism" Then
          If Sheet14.Range(Target.Address).value <> "(Select)" Then
        
'        If Sheet14.Range(Target.Address).Value <> "(viii) autism" And _
'        Sheet14.Range(Target.Address).Value <> "(ix) cerebral palsy" And _
'        Sheet14.Range(Target.Address).Value <> "(x) multiple disability" Then
             
'        If Sheet14.Range(Target.Address).Value <> "(i) autism, cerebral palsy, or multiple disabilities and" Then
         If Sheet14.Range(Target.Address).value <> "(i) autism, cerebral palsy, or multiple disabilities" Then  'Ankita_14/05/2025
            
            If Sheet14.Range(Target.Address).value = "(ii) others" And Sheet14.Range(Replace(Target.Address, "D", "F")).value <> "" Then
            MsgBox """Type of disability shall be selected as ""(i) autism, cerebral palsy, or multiple disabilities"" when acknowledgement no. of form 10IA is provided"""
            Sheet14.Range(Target.Address).value = "(Select)"
            Else:
            Sheet14.Range(Replace(Target.Address, "D", "F")).ClearContents
            Sheet14.Range(Replace(Target.Address, "D", "F")).Locked = True
            Sheet14.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HD8D8D8"
            End If
'            Sheet14.Range(Replace(Target.Address, "D", "G")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "G")).Locked = True
'            Sheet14.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HD8D8D8"
            
'            Sheet14.Range(Replace(Target.Address, "D", "H")).Locked = False
'            Sheet14.Range(Replace(Target.Address, "D", "H")).Interior.Color = "&HCCFFCC"
            
            
        Else
        
        Sheet14.Unprotect Password:=getmsgstate
        
'            Sheet14.Range(Replace(Target.Address, "D", "F")).ClearContents 'Ankita_28/05/2025
            Sheet14.Range(Replace(Target.Address, "D", "F")).Locked = False
            Sheet14.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HCCFFCC"
            
'            Sheet14.Range(Replace(Target.Address, "D", "G")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "G")).Locked = False
'            Sheet14.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HCCFFCC"
            
'            Sheet14.Range(Replace(Target.Address, "D", "H")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "H")).Locked = True
'            Sheet14.Range(Replace(Target.Address, "D", "H")).Interior.Color = "&HD8D8D8"
            Sheet14.Protect Password:=getmsgstate
             

    End If
    
    Else
            Sheet14.Unprotect Password:=getmsgstate
'            Sheet14.Range(Replace(Target.Address, "D", "F")).ClearContents  'Ankita_28/05/2025
            Sheet14.Range(Replace(Target.Address, "D", "F")).Locked = False
            Sheet14.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HCCFFCC"
''
'            Sheet14.Range(Replace(Target.Address, "D", "G")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "G")).Locked = False
'            Sheet14.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HCCFFCC"
'
'            Sheet14.Range(Replace(Target.Address, "D", "H")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "H")).Locked = False
'            Sheet14.Range(Replace(Target.Address, "D", "H")).Interior.Color = "&HCCFFCC"
    End If
    
End If






'Commented by Ayush_06/05/2025 for 0.7 changes
'If Not Application.Intersect(Sheet14.Range("Form10IA_80DD"), Sheet14.Range(Target.Address)) Is Nothing Then
'
'    If Target.value <> "00/00/0000" And Target.value <> "" Then
'       If Not CheckDateddmmyyyy(Trim(Target.value)) Then
'       'Ayush_10/02
''            fmsgbox "* Please enter a valid Date of filing of Form 10IA in dd/mm/yyyy format in Schedule 80DD." ', vbOKOnly, mdCommon.vbMessageCaption
'            fmsgbox "* ""Date of filing of Form 10IA - Please enter valid date in valid format.""" ', vbOKOnly, mdCommon.vbMessageCaption
'            Target.ClearContents
'        End If
'    End If
'End If

'Commented by Ayush_06/05/2025 for 0.7 changes
'If Not Intersect(Range("Form10IA_80DD"), Range(Target.Address)) Is Nothing Then
'    If Trim(Range("Form10IA_80DD").value) <> "" Then
'            If Dformat(Trim(Range("Form10IA_80DD").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'                 fmsgbox "* Date of filing of Form 10IA can not be after the system date in Schedule 80DD"
'                       Range("Form10IA_80DD").Select
'                        Range("Form10IA_80DD").value = ""
'            End If
'    End If
'End If


'Ayush_Legacy1
If Not Intersect(Range("Dependent_80DD"), Range(Target.Address)) Is Nothing Then
    If Sheet1.Range("sheet1.Status").value = "I - INDIVIDUAL" Or Sheet1.Range("sheet1.Status").value = "F - FIRM(Other than LLP)" Then
            If Range("Dependent_80DD").value = "8-Member of the HUF (in case of HUF)" Then
                 fmsgbox "*Only HUF can select  Type of dependent as ""Member of the HUF"". Please select applicable dropdown in ""Type of Dependent""."
                       Range("Dependent_80DD").Select
                        Range("Dependent_80DD").value = ""
            End If
    End If
End If
 
 
If Not Intersect(Range("AadhaarDependent_80DD"), Range(Target.Address)) Is Nothing Then
    If Trim(Range("AadhaarDependent_80DD").value) <> "" Then
              If UCase(Trim(Range("AadhaarDependent_80DD").value)) = UCase(Trim(Sheet1.Range("Sheet1.Aadhaar").value)) Then
                 fmsgbox "*""Aadhaar of the dependent cannot be same as assessee Aadhaar  in Part-A General Information""."
                       Range("AadhaarDependent_80DD").Select
                        Range("AadhaarDependent_80DD").value = ""
                   End If
    End If
     If Len(Sheet14.Range("AadhaarDependent_80DD").value) >= 12 Then
            If chkNumeric(Sheet14.Range("AadhaarDependent_80DD").value) Then
                If Sheet14.Range("AadhaarDependent_80DD").value <> "000000000000" Then
                    Sheet14.Range("AadhaarDependent_80DD").value = Sheet14.Range("AadhaarDependent_80DD").value
                Else
                    fmsgboxStatus "* ""Invalid Aadhaar Number in Schedule 80DD """
                    Sheet14.Range("AadhaarDependent_80DD").value = ""
                    Sheet14.Range("AadhaarDependent_80DD").Select
                End If
                If Sheet14.Range("AadhaarDependent_80DD").value < 0 Then
                    fmsgboxStatus "* ""Invalid Aadhaar Number in Schedule 80DD """
                    Sheet14.Range("AadhaarDependent_80DD").value = ""
                    Sheet14.Range("AadhaarDependent_80DD").Select
                End If
                If Sheet14.Range("AadhaarDependent_80DD").value <> "111111111111" Then
                    Sheet14.Range("AadhaarDependent_80DD").value = Sheet14.Range("AadhaarDependent_80DD").value
                Else
                    fmsgboxStatus "* ""Invalid Aadhaar Number in Schedule 80DD"""
                    Sheet14.Range("AadhaarDependent_80DD").value = ""
                    Sheet14.Range("AadhaarDependent_80DD").Select
                End If
            Else
                fmsgboxStatus "* ""Invalid Aadhaar Number in Schedule 80DD"""
                Sheet14.Range("AadhaarDependent_80DD").value = ""
                Sheet14.Range("AadhaarDependent_80DD").Select
            End If
        Else
            Sheet14.Range("AadhaarDependent_80DD").value = ""
        End If
End If


'Newly added by Bindu
 
If Not Intersect(Range("PanDependent_80DD"), Range(Target.Address)) Is Nothing Then
    If Trim(Range("PanDependent_80DD").value) <> "" Then
    
              Target.value = UCase(Target.value)
              If UCase(Trim(Range("PanDependent_80DD").value)) = UCase(Trim(Sheet1.Range("sheet1.PAN").value)) Then
                 fmsgbox "*""PAN of the dependent cannot be same as assessee PAN in Part-A General Information""."
                       Range("PanDependent_80DD").Select
                        Range("PanDependent_80DD").value = ""
                End If
                
                If Not checkfieldspecialcharacter(Range("PanDependent_80DD").value) Then
                       fmsgbox "*Invalid PAN in Schedule 80DD."
                       Range("PanDependent_80DD").Select
                       Range("PanDependent_80DD").value = ""
                End If
                If Not Check80DDPAN(UCase(Range("PanDependent_80DD").value)) Then
                      fmsgbox "*Invalid PAN in Schedule 80DD."
                       Range("PanDependent_80DD").Select
                       Range("PanDependent_80DD").value = ""
                End If
    End If
End If


'AcknowledgeNum_80DD
If Not Intersect(Range("AcknowledgeNum_80DD"), Range(Target.Address)) Is Nothing Then
    If Trim(Range("AcknowledgeNum_80DD").value) <> "" Then
       If Not checkfieldspecialcharacter(Range("AcknowledgeNum_80DD")) Then
         fmsgbox "Ack. No. of Form 10IA filed cannot contain special characters in Schedule 80DD."
         Range("AcknowledgeNum_80DD").Select
         Range("AcknowledgeNum_80DD").value = ""
       End If
    End If
End If

'UDID 80DD
If Not Intersect(Range("UDIDNumber_80DD"), Range(Target.Address)) Is Nothing Then
    If Trim(Range("UDIDNumber_80DD").value) <> "" Then
    'Ayush_Legacy3
       If Not checkfieldspecialcharacter80DD_80U(Range("UDIDNumber_80DD")) Then
         fmsgbox "UDID Number filed cannot contain special characters in Schedule 80DD."
         Range("UDIDNumber_80DD").Select
         Range("UDIDNumber_80DD").value = ""
       End If
    End If
End If


'AcknowledgeNum_80U
If Not Intersect(Range("AckNumber10IA_80U"), Range(Target.Address)) Is Nothing Then
    If Trim(Range("AckNumber10IA_80U").value) <> "" Then
       If Not checkfieldspecialcharacter(Range("AckNumber10IA_80U")) Then
         fmsgbox "Ack. No. of Form 10IA filed cannot contain special characters in Schedule 80U."
         Range("AckNumber10IA_80U").Select
         Range("AckNumber10IA_80U").value = ""
       End If
    End If
End If

'UDID 80U
If Not Intersect(Range("UDIDNum_80U"), Range(Target.Address)) Is Nothing Then
    If Trim(Range("UDIDNum_80U").value) <> "" Then
    'Ayush_Legacy3
       If Not checkfieldspecialcharacter80DD_80U(Range("UDIDNum_80U")) Then
         fmsgbox "UDID Number filed cannot contain special characters in Schedule 80U."
         Range("UDIDNum_80U").Select
         Range("UDIDNum_80U").value = ""
       End If
    End If
End If

'Commented by Ayush_06/05/2025 for 0.7 changes
'If Not Application.Intersect(Sheet14.Range("DateFiling10IA_80U"), Sheet14.Range(Target.Address)) Is Nothing Then
'
'    If Target.value <> "00/00/0000" And Target.value <> "" Then
'       If Not CheckDateddmmyyyy(Trim(Target.value)) Then
'       'Ayush_10/02
''            fmsgbox "* Please enter a valid Date of filing of Form 10IA in dd/mm/yyyy format in Schedule 80U." ', vbOKOnly, mdCommon.vbMessageCaption
'                fmsgbox "* ""Date of filing of Form 10IA - Please enter valid date in valid format.""" ', vbOKOnly, mdCommon.vbMessageCaption
'            Target.ClearContents
'        End If
'    End If
'End If
 'Commented by Ayush_06/05/2025 for 0.7 changes
' If Not Intersect(Range("DateFiling10IA_80U"), Range(Target.Address)) Is Nothing Then
'    If Trim(Range("DateFiling10IA_80U").value) <> "" Then
'            If Dformat(Trim(Range("DateFiling10IA_80U").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'                 fmsgbox "* Date of filing of Form 10IA can not be after the system date in Schedule 80U"
'                       Range("DateFiling10IA_80U").Select
'                        Range("DateFiling10IA_80U").value = ""
'            End If
'    End If
'End If
 
  
  
  If Not Application.Intersect(Sheet14.Range("PanDependent_80DD"), Sheet14.Range(Target.Address)) Is Nothing Then
        
        'If Not mdIncomeDetails.CheckPAN1(UCase(Trim(Range(TargetPAN1).value))) Then
        'If Not mdIncomeDetails.CheckPAN(UCase(Trim(Range(PanDependent_80DD).value))) Then
        If Not mdIncomeDetails.CheckPAN(UCase(Trim(Target.value))) Then
            
            'fmsgboxsmall "* Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet." & Chr(10) & Chr(13) & "4th letter must be        'P' in Sheet Income Details" ', vbOKCancel, mdCommon.vbMessageCaption
            fmsgbox "* Invalid PAN in Schedule 80DD" ', vbOKCancel, mdCommon.vbMessageCaption
            
            Sheet14.Range("PanDependent_80DD").Select
            
            Sheet14.Range("PanDependent_80DD").value = ""
            
            GoTo endline
            
        End If
 End If
 
 
 
 

Sheet14.Protect Password:=getmsgstate
endline:
Application.EnableEvents = True
 
End Sub

Sub pannnn()
Application.EnableEvents = True
End Sub
