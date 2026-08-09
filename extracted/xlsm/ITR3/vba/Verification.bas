Attribute VB_Name = "Verification"
Public MsgVerificationSheet As Variant

Sub Verification_Click()
    Dim vbMessgaeCaption As String
    vbMessgaeCaption = "ITR 3: AY: 2025-26"
    ValidateVerification1
     fmsgboxStatus "Sheet Verification is ok" ', vbOKOnly, vbMessgaeCaption
End Sub

Function ValidateVerification1() As Boolean
ValidateAssesseeVerName_9
ValidateFatherName_9
ValidatePAN_Ver
'Change 102 13.12.2023 SIT-64711 A
ValidatePlace_9
'End Change
ValidateDate_9


If MsgVerificationSheet <> "" Then
fmsgboxsmall MsgVerificationSheet
Sheet52.Activate
CloseMsg
End If
End Function
'Change 102 13.12.2023 SIT-64711 B
Function ValidatePlace_9() As Boolean
    ValidatePlace_9 = True
    Place_9 = Sheet52.Range("sheet9.Place")

    If Place_9 = "" Or IsEmpty(Place_9) Then
       MsgVerificationSheet = MsgVerificationSheet + "* Please enter Place in Sheet : verification" & Chr(13)
        ValidatePlace_9 = False
        Exit Function
    End If
    
End Function
'End Change
Function ValidateAssesseeVerName_9() As Boolean
    ValidateAssesseeVerName_9 = True
    AssesseeVerName_9 = Sheet52.Range("sheet9.AssesseeVerName")
    
    If Len(AssesseeVerName_9) > 125 Then
       MsgVerificationSheet = MsgVerificationSheet + "* Assessee Verification Name should be at most 125 characters in verification" & Chr(13)
        ValidateAssesseeVerName_9 = False
        Exit Function
    End If
    
    If AssesseeVerName_9 = "" Or IsEmpty(AssesseeVerName_9) Then
      ' MsgVerificationSheet = MsgVerificationSheet + "* Please enter the Assessee Verification Name in Sheet : verification" & Chr(13)
     
      'Added by Aavula SIT-69112
      
        MsgVerificationSheet = MsgVerificationSheet + "* Please enter the name in verification." & Chr(13)

        ValidateAssesseeVerName_9 = False
        Exit Function
    End If
End Function

Function ValidateFatherName_9() As Boolean
    ValidateFatherName_9 = True
    FatherName_9 = Sheet52.Range("sheet9.FatherName")
    
    If Len(FatherName_9) > 125 Then
       MsgVerificationSheet = MsgVerificationSheet + "* FatherName should be at most 125 characters in verification" & Chr(13)
        ValidateFatherName_9 = False
        Exit Function
    End If
    
    If FatherName_9 = "" Or IsEmpty(FatherName_9) Then
'       MsgVerificationSheet = MsgVerificationSheet + "* Please enter the Son/daughter in verification" & Chr(13)
       
       'Added by Aavula SIT-69114
       
       MsgVerificationSheet = MsgVerificationSheet + "* Please enter the name in verification." & Chr(13)

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
    
    If Not (Mid(tempPAN_ver, 4, 1) = "P") Then
       MsgVerificationSheet = MsgVerificationSheet + "* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet in verification." & Chr(13)
        ValidatePAN_Ver = False
        Exit Function
    End If
    
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


'Added by Riyaz on 25/06/2025
'If Sheet1.Range("sheet1.AsseseeRepFlg").value = "Y -Yes" Then
 If Sheet1.Range("sheet1.AsseseeRepFlg").value = "Y - Yes" Then ' by sai on 10/07/2025

        If verCapacity <> "Representative" Then
             'If verCapacity <> "Self" And verCapacity <> "Karta" And verCapacity <> "Authorised Signatory" Then
            MsgVerificationSheet = MsgVerificationSheet + "* Please select ""Representative"" if return is filed by Representative." & Chr(13)
            ValidatePAN_Ver = False
        End If
End If

'If Sheet1.Range("sheet1.AsseseeRepFlg").value = "N -No" Then
If UCase(Mid(Sheet1.Range("sheet1.AsseseeRepFlg").value, 1, 1)) = "N" Then ' by sai on 10/07/2025
        If Sheet1.Range("sheet1.Status").value = "I-INDIVIDUAL" Then
            If verCapacity <> "Self" And verCapacity <> "Authorised Signatory" Then
                MsgVerificationSheet = MsgVerificationSheet + "* Please select ""Self"" or ""Authorised Signatory"" as applicable as capacity" & Chr(13)
                ValidatePAN_Ver = False
            End If
        End If
End If

'Changed by Riyaz on 23/06/2025
'If Sheet1.Range("sheet1.AsseseeRepFlg").value = "N -No" Then

 If UCase(Mid(Sheet1.Range("sheet1.AsseseeRepFlg").value, 1, 1)) = "N" Then ' by sai on 10/07/2025
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
    
'    If Not CheckDateMinDDMMYYYY(Date_9, 1, 4, 2023, "Verification date cannot be less than 01/04/2023") Then
     If Not CheckDateMinDDMMYYYY(Date_9, 1, 4, 2025, "Verification date cannot be less than 01/04/2025") Then 'Year changed by Sai from 2024 to 2025
        ValidateDate_9 = False
'       MsgVerificationSheet = MsgVerificationSheet + "* Date in Verification , Sheet Verification must not be less than 01/04/2023" & Chr(13)
        MsgVerificationSheet = MsgVerificationSheet + "* Date in Verification , Sheet Verification must not be less than 01/04/2025" & Chr(13) 'Year changed by Sai from 2024 to 2025
        Exit Function
    Else
        Date_9 = Dformat(Sheet52.Range("sheet9.Date"), "yyyy-mm-dd")
    End If
End Function



