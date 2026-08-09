Attribute VB_Name = "Verification"
Public MsgVerificationSheet As Variant

Sub Verification_Click()
    Dim vbMessgaeCaption As String
    'vbMessgaeCaption = "ITR 2: AY: 2024-25" 'dpk1601
     vbMessgaeCaption = "ITR 2: AY: 2025-26" 'Shrutika(27-02-25)
    
    ValidateVerification1
     fmsgboxStatus "Sheet Verification is ok" ', vbOKOnly, vbMessgaeCaption
End Sub

Function ValidateVerification1() As Boolean
ValidateAssesseeVerName_9
ValidateFatherName_9
ValidatePAN_Ver
'ValidatePlace_9
ValidateDate_9


If MsgVerificationSheet <> "" Then
fmsgboxsmall MsgVerificationSheet
Sheet52.Activate
CloseMsg
End If
End Function

Function ValidateAssesseeVerName_9() As Boolean
    ValidateAssesseeVerName_9 = True
    AssesseeVerName_9 = Sheet52.Range("sheet9.AssesseeVerName")
    
    If Len(AssesseeVerName_9) > 125 Then
       MsgVerificationSheet = MsgVerificationSheet + "* Name cannot be more than 125 characters in Schedule Verification" & Chr(13)
        ValidateAssesseeVerName_9 = False
        Exit Function
    End If
    
    If AssesseeVerName_9 = "" Or IsEmpty(AssesseeVerName_9) Then
       MsgVerificationSheet = MsgVerificationSheet + "* Please enter the name in verification" & Chr(13)
        ValidateAssesseeVerName_9 = False
        Exit Function
    End If
End Function

Function ValidateFatherName_9() As Boolean
    ValidateFatherName_9 = True
    FatherName_9 = Sheet52.Range("sheet9.FatherName")
    
    If Len(FatherName_9) > 125 Then
       MsgVerificationSheet = MsgVerificationSheet + "* Name cannot be more than 125 characters in Schedule Verification" & Chr(13)
        ValidateFatherName_9 = False
        Exit Function
    End If
    
    If FatherName_9 = "" Or IsEmpty(FatherName_9) Then
       MsgVerificationSheet = MsgVerificationSheet + "* Please enter the name in verification" & Chr(13)
        ValidateFatherName_9 = False
        Exit Function
    End If
End Function

Function ValidatePAN_Ver() As Boolean
    Dim tempPAN_ver As String
    
    ValidatePAN_Ver = True
    verPAN = Sheet52.Range("sheet9.PAN").value
    verCapacity = Sheet52.Range("sheet9.Capacity")
    
    tempPAN_ver = verPAN
    
    If Not chkCompulsory(verPAN) Then
        
       MsgVerificationSheet = MsgVerificationSheet + "* Please enter the PAN in verification" & Chr(13)
        
        ValidatePAN_Ver = False
        Exit Function
    End If
    
    If Len(verPAN) > 10 Then
       MsgVerificationSheet = MsgVerificationSheet + "* PAN should be 10 digits in verification" & Chr(13)
        ValidatePAN_Ver = False
        Exit Function
    End If
    'Commented by Shrutika_EN
'    If Not (Mid(tempPAN_ver, 4, 1) = "P") Then
'       MsgVerificationSheet = MsgVerificationSheet + "* PAN is invalid (PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet. 4th character must be 'P') in verification" & Chr(13)
'        ValidatePAN_Ver = False
'        Exit Function
'    End If
    
    If Not ValidatePAN(tempPAN_ver) Then
       MsgVerificationSheet = MsgVerificationSheet + "* PAN in Sheet : Verification is invalid (10 digits valid PAN)" & Chr(13)
        ValidatePAN_Ver = False
        Exit Function
    End If
        
    If Not checkfieldspecialcharacter(verPAN) Then
       MsgVerificationSheet = MsgVerificationSheet + "* PAN characters < > & ' " & Chr(34) & " are not allowed in verification" & Chr(13)
        ValidatePAN_Ver = False
        Exit Function
    End If
    
    If isdropdownblank(verCapacity) Then
       MsgVerificationSheet = MsgVerificationSheet + "* Selection of capacity in verification Part is mandatory. Please select appropriate option from drop down" & Chr(13)
        ValidatePAN_Ver = False
        Exit Function
    End If
    
    If Not isdropdownblank(verCapacity) Then
    If verCapacity <> "Self" And verCapacity <> "Representative" And verCapacity <> "Karta" And verCapacity <> "Authorised Signatory" Then
       MsgVerificationSheet = MsgVerificationSheet + "* Selection of capacity in verification Part is mandatory. Please select appropriate option from drop down" & Chr(13)
        ValidatePAN_Ver = False
        Exit Function
    End If
    End If
'
'     If calculateAge(Trim(Sheet1.Range("sheet1.DOB").value)) < 18 And Mid(Sheet1.Range("sheet1.status").value, 1, 1) = "I" Then
'          If Sheet52.Range("Sheet9.Capacity").value <> "Representative" Then
'            MsgVerificationSheet = MsgVerificationSheet + "*If status is individual & date of birth is on or after 2 April 2002 then in verification drop down cannot be selected as self/Karta/ Partner" & Chr(13)
'            Sheet52.Range("Sheet9.Capacity").value = "(Select)"
'            ValidatePAN_Ver = False
'          End If
'        End If
        
'      If calculateAge(Trim(Sheet1.Range("sheet1.DOB").value)) < 18 And Mid(Sheet1.Range("sheet1.status").value, 1, 1) = "I" Then
'          If Sheet52.Range("Sheet9.Capacity").value = "Representative" Then
'            If Sheet52.Range("sheet9.PAN").value = Sheet1.Range("sheet1.PAN").value Then
'            MsgVerificationSheet = MsgVerificationSheet + "*If status is individual & date of birth is on or after 2 April 2002 then in verification PAN cannot be same as mentioned in general information." & Chr(13)
'            Sheet52.Range("sheet9.PAN").value = ""
'            ValidatePAN_Ver = False
'            End If
'          End If
'        End If


    'Added by Shrutika(20/06/5)
       
If Sheet1.Range("sheet1.AsseseeRepFlg").value = "Y -Yes" Then
If verCapacity <> "Representative" Then
MsgVerificationSheet = MsgVerificationSheet + "* Please select ""Representative"" if return is filed by Representative." & Chr(13)
ValidatePAN_Ver = False
End If
End If

If Sheet1.Range("sheet1.AsseseeRepFlg").value = "N -No" Then
If Sheet1.Range("sheet1.Status").value = "I-INDIVIDUAL" Then
If verCapacity <> "Self" And verCapacity <> "Authorised Signatory" Then
MsgVerificationSheet = MsgVerificationSheet + "* Please select ""Self"" or ""Authorised Signatory"" as applicable as capacity" & Chr(13)
ValidatePAN_Ver = False
End If
End If
End If


If Sheet1.Range("sheet1.AsseseeRepFlg").value = "N -No" Then
If Sheet1.Range("sheet1.Status").value = "H-HUF" Then
If verCapacity <> "Karta" And verCapacity <> "Authorised Signatory" Then
MsgVerificationSheet = MsgVerificationSheet + "* Please select ""Karta"" or ""Authorised signatory"" as applicable as capacity" & Chr(13)
ValidatePAN_Ver = False
End If
End If
End If

End Function


Function ValidateDate_9() As Boolean
    Dim i As Long
    
    ValidateDate_9 = True
    Date_9 = Sheet52.Range("sheet9.Date").value
    If Date_9 = "" Or IsEmpty(Date_9) Then
       MsgVerificationSheet = MsgVerificationSheet + "* Date in Sheet : Verification  is Mandatory" & Chr(13)
        ValidateDate_9 = False
        Exit Function
    End If
    
    If Not CheckDateddmmyyyy(Date_9) Then
        ValidateDate_9 = False
       MsgVerificationSheet = MsgVerificationSheet + "* Date in Sheet : Verification    must be a valid dd/mm/yyyy format" & Chr(13)
        Exit Function
    Else
        Date_9 = Dformat(Sheet52.Range("sheet9.Date"), "yyyy-mm-dd")
    End If
    
    If Not CheckDateMinDDMMYYYY(Date_9, 1, 4, 2024, "Verification date cannot be less than 01/04/2024") Then 'dpk1601
        ValidateDate_9 = False
       MsgVerificationSheet = MsgVerificationSheet + "* Date in Verification , Sheet Verification must not be less than 01/04/2024" & Chr(13) 'dpk1601
        Exit Function
    Else
        Date_9 = Dformat(Sheet52.Range("sheet9.Date"), "yyyy-mm-dd")
    End If
End Function



