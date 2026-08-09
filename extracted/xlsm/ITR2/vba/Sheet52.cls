Attribute VB_Name = "Sheet52"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True


Private Sub Worksheet_Deactivate()

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


Sub representativecheck()

sPassword = EfilingCommon.getmsgstate
  Sheet1.Unprotect Password:=sPassword
If Mid(Sheet1.Range("Sheet1.AsseseeRepFlg").value, 1, 1) = "N" Then


       Sheet1.Range("sheet1.RepName").MergeArea.Locked = True
       Sheet1.Range("sheet1.RepName").MergeArea.Interior.Color = (&HD8D8D8)
       Sheet1.Range("sheet1.RepName").MergeArea.ClearContents
       Sheet1.Range("sheet1.RepCapacity").MergeArea.Locked = True
       Sheet1.Range("sheet1.RepCapacity").MergeArea.Interior.Color = (&HD8D8D8)
       Sheet1.Range("sheet1.RepCapacity").MergeArea.ClearContents


       Sheet1.Range("sheet1.RepAddress").MergeArea.Locked = True
       Sheet1.Range("sheet1.RepAddress").MergeArea.Interior.Color = (&HD8D8D8)
       Sheet1.Range("sheet1.RepAddress").MergeArea.ClearContents
       Sheet1.Range("sheet1.RepPAN").MergeArea.Locked = True
       Sheet1.Range("sheet1.RepPAN").MergeArea.Interior.Color = (&HD8D8D8)
       Sheet1.Range("sheet1.RepPAN").MergeArea.ClearContents
       
       Sheet1.Range("sheet1.RepAadhar").MergeArea.Locked = True
       Sheet1.Range("sheet1.RepAadhar").MergeArea.Interior.Color = (&HD8D8D8)
       Sheet1.Range("sheet1.RepAadhar").MergeArea.ClearContents
   Else:
       Sheet1.Range("sheet1.RepName").MergeArea.Locked = False
       Sheet1.Range("sheet1.RepName").MergeArea.Interior.Color = (&HCCFFCC)
       If Sheet1.Range("sheet1.RepCapacity").value = "" Then
            Sheet1.Range("sheet1.RepCapacity").value = "(Select)"
            Sheet1.Range("sheet1.RepCapacity").MergeArea.Locked = False
            Sheet1.Range("sheet1.RepCapacity").MergeArea.Interior.Color = (&HCCFFCC)
       End If
       Sheet1.Range("sheet1.RepAddress").MergeArea.Locked = False
       Sheet1.Range("sheet1.RepAddress").MergeArea.Interior.Color = (&HCCFFCC)
       Sheet1.Range("sheet1.RepPAN").MergeArea.Locked = False
       Sheet1.Range("sheet1.RepPAN").MergeArea.Interior.Color = (&HCCFFCC)
       
       Sheet1.Range("sheet1.RepAadhar").MergeArea.Locked = False
       Sheet1.Range("sheet1.RepAadhar").MergeArea.Interior.Color = (&HCCFFCC)
   End If
   sPassword = EfilingCommon.getmsgstate
  Sheet1.Protect Password:=sPassword
   End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
Application.EnableEvents = False

If (Target.name.name) = "Sheet9.Capacity" Then
If Sheet52.Range("Sheet9.Capacity").value = "Representative" Then
fmsgboxStatus "Details of representative assessee is mandatory as the return is being filed by representative in schedule Part A General"
Sheet1.Range("Sheet1.AsseseeRepFlg").value = "Y -Yes"
representativecheck
End If

' If calculateAge(Trim(Sheet1.Range("sheet1.DOB").value)) < 18 And Mid(Sheet1.Range("sheet1.status").value, 1, 1) = "I" Then
'          If Sheet52.Range("Sheet9.Capacity").value <> "Representative" Then
'            fmsgboxStatus "*If status is individual & date of birth is on or after 2 April 2002 then in verification drop down cannot be selected as self/Karta/ Partner"
'            Sheet52.Range("Sheet9.Capacity").value = "(Select)"
'          End If
'        End If
End If

If Target.name.name = "sheet9.PAN" Then
    Target.value = UCase(Target.value)
    If Sheet52.Range("sheet9.PAN").value <> "" Then
    If Not CheckPAN(Sheet52.Range("sheet9.PAN").value) Then
    'Ankita(23-Jan)
        'fmsgboxStatus "Please Enter a Valid Pan"
        'fmsgboxStatus "Invalid PAN in Schedule Verification"
        fmsgboxStatus """Invalid PAN in Schedule Verification.""" & Chr(13)
        '--------------------
        Sheet52.Range("sheet9.PAN").value = ""
        Sheet52.Range("sheet9.PAN").Select
    Else
'Commented by Shrutika_EN
'    If Mid(Sheet52.Range(Target.name.name).value, 4, 1) <> "P" Then
'    'Ankita(23-Jan)
'        ' fmsgboxStatus "Please Enter a Valid Individual Pan"
'         fmsgboxStatus """Invalid PAN in Schedule Verification.""" & Chr(13)
'         '---------------------
'        Sheet52.Range("sheet9.PAN").value = ""
'        Sheet52.Range("sheet9.PAN").Select
'    End If
    End If
    End If
End If

 If Target.name.name = "sheet9.Date" Then

    CalculateTaxCheckFlag = False

                'If Trim(Range("sheet9.Date").value) <> "00/00/0000" Then
                    If Not CheckDateBeforeinSheet9(Sheet7.Range("sheet9.Date").value, "Verification Date") Then
                        Range("sheet9.Date").Select
                        Range("sheet9.Date").value = ""
                    Else
                    Range("sheet9.Date").value = CreditDate
                    End If
                'End If
               GoTo endline
            End If

            If Target.name.name = "Sheet9.IdentificationNoOfTRP" Then
    Target.value = UCase(Target.value)
End If
endline:
Application.EnableEvents = True
  ActiveSheet.Protect Password:=EfilingCommon.getmsgstate
End Sub



Function CheckDateBeforeinSheet9(dob As Variant, errmsg As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.
'DoF should not be after 31/03/2018, for A.Y. 2018-19

   CheckDateBeforeinSheet9 = True
   If Trim(dob) = "" Or Not IsEmpty(dob) Then
       If Not FormatNCheckDate(dob, "2100") Then
           CheckDateBeforeinSheet9 = False
          fmsgboxStatus (errmsg & " PartB-TI TTI  must be a valid dd/mm/yyyy format")
           Exit Function
       End If

       If Not ChkMinDOBDate9(dob, "31/03/2024") Then 'dpk1601
       
           fmsgboxStatus (errmsg & " should be on or after 01/04/2024") 'dpk1601
           CheckDateBeforeinSheet9 = False
           Exit Function
       Else
           CreditDate = dob
       End If
   End If
End Function

Function ChkMinDOBDate9(dob As Variant, minDefinedDOB As Variant) As Boolean
On Error Resume Next
    ChkMinDOBDate9 = True
    If Len(dob) > 0 Then
       If val(Mid(dob, 7, 4)) < 2024 Then 'dpk1601
           ChkMinDOBDate9 = False
           Exit Function
       ElseIf (val(Mid(dob, 7, 4)) = 2024 And val(Mid(dob, 4, 2)) <= 3) Then
           ChkMinDOBDate9 = False
           Exit Function
       End If
   End If
End Function


