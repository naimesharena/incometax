Attribute VB_Name = "AY23_24Changes"
Function CheckDOBAY_2324(dob As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.
'DoF should not be after 31/03/2018, for A.Y. 2015-16
    Dim nstatus As String
    nstatus = Sheet1.Range("sheet1.Status").value
    CheckDOBAY_2324 = True
    nstatus = Mid(nstatus, 1, 1)
    If Trim(dob) <> "" Or Not IsEmpty(dob) Then
        If (dob = "01/01/0001" And nstatus = "H") Then
            CheckDOBAY_2324 = True
            AssesseeDob = dob
            Exit Function
        Else
            If Not FormatNCheckDate(dob) Then
             
                CheckDOBAY_2324 = False
               
                fmsgboxStatus ("Please enter valid date in dd/mm/yyyy format  in schedule Part A General under Personal Information")
                
                Exit Function
            End If
        'Chandru (Part A Gen DOB Year change)
            'If Not ChkMaxDOBDate(dob, "31/03/2024") Then
        'Added by Shrutika(20-12-24)AY-25
           ' If Not ChkMaxDOBDate(dob, "31/03/2025") Then
    'Added by Shrutika- 16/01/2026 AY-26-27
                 Dim cutoff As Date
                 cutoff = CDate(SheetALL.Range("DOB_1").value)
                 If Not ChkMaxDOBDate1(dob, SheetALL.Range("DOB_1").value) Then
                 
                 'fmsgboxStatus ("Date should not be after 31/03/2024 for AY 2024-25 in schedule Part A General under Personal Information")
              '   fmsgboxStatus ("Date should not be after 31/03/2025 for AY 2025-26 in schedule Part A General under Personal Information")
'                 fmsgboxStatus ("Date should not be after " & Dformat1(cutoff, "yyyy-mm-dd") & " for AY 2026-27 in schedule Part A General under Personal Information")
                  fmsgboxStatus ("Date should not be after 31/03/2026 for AY 2026-27 in schedule Part A General under Personal Information")

                CheckDOBAY_2324 = False
                Exit Function
            Else
                AssesseeDob = dob
            End If
        End If
    End If
 End Function


Function ChkMaxDOBDate(dob As Variant, maxDefinedDOB As Variant) As Boolean
On Error Resume Next
     ChkMaxDOBDate = True
     If Len(dob) > 0 Then
     'CG_E3_AY_2023-24 OLD 'Malli
     'If val(Mid(dob, 7, 4)) > 2023 Then
     'CG_C3_AY_2024-25 Change
       'If val(Mid(dob, 7, 4)) > 2024 Then
       'Added by Shrutika(20-12-24)AY-25
'Year 2025 To 2026 changed by Riyaz on 22/01/2026
       If val(Mid(dob, 7, 4)) > 2026 Then
                    ChkMaxDOBDate = False
                    Exit Function
        'ElseIf val(Mid(dob, 7, 4)) = 2024 Then
        ElseIf val(Mid(dob, 7, 4)) = 2026 Then
            If val(Mid(dob, 4, 2)) > 3 Then
                If val(Mid(dob, 1, 2)) >= 1 Then
                    ChkMaxDOBDate = False
                    Exit Function
                End If
            End If
        End If
     End If
End Function

 'Added by Shrutika- 16/01/2026 AY-26-27

Function ChkMaxDOBDate1(dob As Variant, maxDefinedDOB As Variant) As Boolean
On Error Resume Next
     ChkMaxDOBDate1 = True
     If Len(dob) > 0 Then
     'CG_E3_AY_2023-24 OLD 'Malli
     'If val(Mid(dob, 7, 4)) > 2023 Then
     'CG_C3_AY_2024-25 Change
       'If val(Mid(dob, 7, 4)) > 2024 Then
       'Added by Shrutika(20-12-24)AY-25
     '  If val(Mid(dob, 7, 4)) > 2025 Then
       If val(Mid(dob, 7, 4)) > CInt(SheetALL.Range("DOB_Year").value) Then  'Added by Shrutika- 02/01/2026 AY-26-27
                    ChkMaxDOBDate1 = False
                    Exit Function
        'ElseIf val(Mid(dob, 7, 4)) = 2024 Then
       ' ElseIf val(Mid(dob, 7, 4)) = 2025 Then
        ElseIf val(Mid(dob, 7, 4)) = CInt(SheetALL.Range("DOB_Year").value) Then   'Added by Shrutika- 02/01/2026 AY-26-27
            If val(Mid(dob, 4, 2)) > 3 Then
                If val(Mid(dob, 1, 2)) >= 1 Then
                    ChkMaxDOBDate1 = False
                    Exit Function
                End If
            End If
        End If
     End If
End Function


Sub PartAGen_Country()   '17/10/2024
      If UCase(Sheet1.Range("sheet1.CountryCode")) = "91-INDIA" Then
      'Malli------------------------16/10/2024
           'Sheet1.Range("sheet1.StateCode1") = "(Select)"
            Sheet1.Range("sheet1.PinCode") = ""
            'Malli------------------------16/10/2024
            'If UCase(Sheet1.Range("sheet1.StateCode1")) = "99-STATE OUTSIDE INDIA" Then
            
            If Sheet1.Range("sheet1.StateCode1") = "99-Foreign" Or UCase(Sheet1.Range("sheet1.StateCode1")) = "99-FOREIGN" Then
            '---------------------------------------------
                
               'fmsgboxStatus "*Country Should not be India for Selected State" & Chr(13)
'by sadineni on 16/06/2026 as per DE V0.4
                fmsgboxStatus "*""Country cannot be India since state is Foreign""" & Chr(13)
                
                Sheet1.Range("sheet1.CountryCode").value = "(Select)"
                Sheet1.Range("sheet1.CountryCode").Select
            End If
            Sheet1.Unprotect Password:=getmsgstate
               Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
               Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)

               Sheet1.Range("HASZIP").MergeArea.value = ""
               Sheet1.Range("HASZIP").MergeArea.Locked = True
               Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)

               Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
               Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
               Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Protect Password:=getmsgstate
      ElseIf Sheet1.Range("sheet1.CountryCode").value = "(Select)" Or Sheet1.Range("sheet1.CountryCode").value = "" Then
            Sheet1.Range("sheet1.StateCode1") = "(Select)"

            Sheet1.Unprotect Password:=getmsgstate
               Sheet1.Range("Countrycode1").MergeArea.value = ""
               Sheet1.Range("Countrycode2").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
               Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)

               Sheet1.Range("HASZIP").MergeArea.value = ""
               Sheet1.Range("HASZIP").MergeArea.Locked = True
               Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)

               Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
               Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
               Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
            Sheet1.Protect Password:=getmsgstate
             'Malli------------------------16/10/2024
             'ElseIf UCase(Sheet1.Range("sheet1.CountryCode").value) <> "91-INDIA" Then
             'If UCase(Sheet1.Range("sheet1.StateCode1")) <> "99-STATE OUTSIDE INDIA" And (Sheet1.Range("sheet1.StateCode1") <> "" And Sheet1.Range("sheet1.StateCode1") <> "(Select)") Then
             
              ElseIf UCase(Sheet1.Range("sheet1.CountryCode").value) <> "91-INDIA" Then
              If UCase(Sheet1.Range("sheet1.StateCode1")) <> "99-FOREIGN" Or Sheet1.Range("sheet1.StateCode1") <> "99-Foreign" And (Sheet1.Range("sheet1.StateCode1") <> "" And Sheet1.Range("sheet1.StateCode1") <> "(Select)") Then
                '----------------------------------------
'                fmsgboxStatus "*Country Should be India for Selected State" & Chr(13)
'by sadinnei on 16/06/2026 as per DE V0.4
                 fmsgboxStatus "*""Country cannot be other than India as you have selected an Indian state in Part A General.""" & Chr(13)
                Sheet1.Range("sheet1.CountryCode").value = "91-INDIA"
                Return
            End If
            'Malli------------------------16/10/2024
            'Sheet1.Range("sheet1.StateCode1").value = "99-STATE OUTSIDE INDIA"
             Sheet1.Range("sheet1.StateCode1").value = "99-Foreign"
            '---------------------------------
            
            Sheet1.Unprotect Password:=getmsgstate
               Sheet1.Range("Countrycode1").MergeArea.value = ""
               Sheet1.Range("Countrycode2").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode").MergeArea.Locked = True
               Sheet1.Range("sheet1.PinCode").MergeArea.value = ""
               Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HD8D8D8)

               Sheet1.Range("HASZIP").MergeArea.value = ""
               Sheet1.Range("HASZIP").MergeArea.Locked = False
               Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HCCFFCC)

               Sheet1.Range("sheet1.ZipCode").MergeArea.value = ""
               Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = False
               Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HCCFFCC)
            Sheet1.Protect Password:=getmsgstate
      End If
End Sub


Sub PartAGen_Adhaarno()
    If ChkAdhaarNo = False Then
        MsgBox ("Please enter the valid Aadhaar number in Part A General")
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.adhaarno").MergeArea.ClearContents
        Sheet1.Protect Password:=getmsgstate
    End If
End Sub

'Commented by Shrutika_EN
'Sub PartAGen_AdhaarEnroll()
'    If ChkAdhaarEnrollID = False Then
'        MsgBox ("Please enter the valid Aadhaar Enrollment ID in Part A General")
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("sheet1.adhaarenrol").MergeArea.ClearContents
'        Sheet1.Protect Password:=getmsgstate
'    End If
'End Sub
'-----------------
Function CheckDateBefore(dob As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.
'DoF should not be after 31/03/2019, for A.Y. 2018-19

    CheckDateBefore = True
    If Len(dob) > 0 Then
        If Not FormatNCheckDate(dob) Then
            CheckDateBefore = False
           fmsgboxStatus ("Date of Filing in Sheet : PART A-GENERAL  must be a valid dd/mm/yyyy format")
            Exit Function
        End If
        
       ' If EfilingCommon.checkFirstDateBefore(dob, "31/03/2024") Then 'Chandru PAG Change
        'Added by Shrutika(20-12-24)AY-25
      '  If EfilingCommon.checkFirstDateBefore(dob, "31/03/2025") Then
      
      'Added by Shrutika- 02/01/2026 AY-26-27)
         Dim cutoff As Date
        cutoff = CDate(SheetALL.Range("Date_Filing1").value)
          If EfilingCommon.checkFirstDateBefore(dob, SheetALL.Range("Date_Filing1").value) Then
            'fmsgboxStatus ("Date of filing in Sheet : PART A-GENERAL  should be on or after 01/04/2024")
          '  fmsgboxStatus ("Date of filing in Sheet : PART A-GENERAL  should be on or after 01/04/2025")
     '   fmsgboxStatus ("Date of filing in Sheet : PART A-GENERAL should be on or after " & Dformat2(cutoff, "yyyy-mm-dd"))
         'Added by Shrutika- 23/04/2026 V0.3
         fmsgboxStatus ("""Date of filing of original return cannot be prior to 01/04/2026 of AY 26-27.""")
            CheckDateBefore = False
            Exit Function
        Else
                EfilingDate = dob
       End If
    End If
 End Function
 
 'Added by Shrutika- 02/01/2026 AY-26-27)
 
 Function Dformat2(dt As Variant, timepass As String) As String

'yyyy-mm-dd'
 
Dim formateddate As String

Dim day As String

Dim month As String

Dim Year As String

If Len(SheetALL.Range("Date_Filing2").value) > 0 Then
 
Year = Mid(SheetALL.Range("Date_Filing2").value, 7, 4)

month = Mid(SheetALL.Range("Date_Filing2").value, 4, 2)

day = Mid(SheetALL.Range("Date_Filing2").value, 1, 2)

formateddate = day & "/" & month & "/" & Year

Dformat2 = formateddate
 
Else

Dformat2 = ""

End If

End Function

Sub IT1_DateofDeposit(ByVal Target As Range)
        Dim temp
        temp = Dformat(Sheet202.Range(Target.Address).value, "")
        'If Not ChkMinInclusiveDate(temp, "2024-04-01") Then
        'Shrutika_Ur
        'If Not ChkMinInclusiveDate(temp, "2025-01-01") Then
        
        If Not ChkMinInclusiveDate(temp, "2026-01-01") Then
        
        'UR_Shrutika 03/11/2025
            'fmsgboxsmall ("The date of credit to Central government cannot be prior to 01/04/2024 in schedule IT.")
            'Shrutika_Ur
           'fmsgboxsmall ("""Please enter valid date in dd/mm/yyyy format in schedule IT. Also date of credit to Central government cannot be prior to 01/01/2025 in schedule IT.""")
           fmsgboxsmall ("""Please enter valid date in dd/mm/yyyy format in schedule IT. Also date of credit to Central government cannot be prior to 01/01/2026 in schedule IT.""")
   
            Sheet202.Range(Target.Address).value = ""
        End If
        
End Sub

Function ValidateDate_IT1() As Boolean
ValidateDate_IT1 = True
Dim rangecells As Range
Set rangecells = Sheet202.Range("U_DateDep1").Cells
Dim i As Long
ReDim DateCredit_IT1(ColCount2_IT1)
For i = 1 To ColCount2_IT1
DateCredit_IT1(i) = rangecells.item(i).value
If Len(DateCredit_IT1(i)) = 0 Then
End If
If Not FormatNCheckDate(DateCredit_IT1(i)) Then
    ValidateDate_IT1 = False
   msgbox_1398AATI = msgbox_1398AATI + "* Date of credit into Govt Account must be a valid dd/mm/yyyy format at Sr.NO " & i & " in schedule IT" & Chr(13)
   
    Exit Function
End If

If Not chkCompulsory_1398A(DateCredit_IT1(i)) Then
    'msgbox_1398AATI = msgbox_1398AATI + "* Please enter the Date of credit into Govt Account at Sr.NO " & i & " in schedule IT" & Chr(13)
   'Added by Shruika
    msgbox_1398AATI = msgbox_1398AATI + "* Date of credit into Govt Account must be a valid dd/mm/yyyy format at Sr.NO " & i & " in schedule IT-1" & Chr(13)
    ValidateDate_IT1 = False
    Exit Function
End If

'Newly changed from 2023-2024
'If Not ChkMinInclusiveDate_1398A(Dformat(DateCredit_IT1(i), ""), "2024-04-01") Then ''
'Shrutika_Ur
'If Not ChkMinInclusiveDate_1398A(Dformat(DateCredit_IT1(i), ""), "2025-01-01") Then
If Not ChkMinInclusiveDate_1398A(Dformat(DateCredit_IT1(i), ""), "2026-01-01") Then  'UR_Shrutika 03/11/2025

    'msgbox_1398AATI = msgbox_1398AATI + "*Date of credit to Central government at Sr. No  " & i & " cannot be prior to 01/04/2024 in schedule IT" & Chr(13)
   'Added by Shrutika_Ur
   ' msgbox_1398AATI = msgbox_1398AATI + "* ""Please enter valid date in dd/mm/yyyy format in schedule IT. Also date of credit to Central government cannot be prior to 01/01/2025 in schedule IT.""" & Chr(13)
    msgbox_1398AATI = msgbox_1398AATI + "* ""Please enter valid date in dd/mm/yyyy format in schedule IT. Also date of credit to Central government cannot be prior to 01/01/2026 in schedule IT.""" & Chr(13)

    ValidateDate_IT1 = False
    Exit Function
End If

Next
End Function

Sub IT2_DateofDeposit(ByVal Target As Range)
    Dim temp1
    temp1 = Dformat(Sheet202.Range(Target.Address).value, "")
   ' If Not ChkMinInclusiveDate(temp1, "2024-04-01") Then 'Newly changed from 2023 to 2024
       'Shrutika_Ur
    '   If Not ChkMinInclusiveDate(temp1, "2023-04-01") Then
       If Not ChkMinInclusiveDate(temp1, "2024-04-01") Then  'UR_Shrutika 03/11/2025
       
        'fmsgboxsmall ("The date of credit to Central government cannot be prior to 01/04/2024 in schedule IT.")
        'Shrutika_Ur
      '  fmsgboxsmall ("""Please enter valid date in dd/mm/yyyy format in schedule IT. Also date of credit to Central government cannot be prior to 01/04/2023 in schedule IT.""")
        fmsgboxsmall ("""Please enter valid date in dd/mm/yyyy format in schedule IT. Also date of credit to Central government cannot be prior to 01/04/2024 in schedule IT.""")

        Sheet202.Range(Target.Address).value = ""

    End If
 

End Sub
Function ValidateDate_IT2() As Boolean
ValidateDate_IT2 = True
Dim rangecells As Range
Set rangecells = Sheet202.Range("U_DateDep2").Cells
Dim i As Long
ReDim DateCredit_IT2(ColCount2_IT2)
For i = 1 To ColCount2_IT2
DateCredit_IT2(i) = rangecells.item(i).value
If Len(DateCredit_IT2(i)) = 0 Then
End If
If Not FormatNCheckDate_1398A(DateCredit_IT2(i)) Then
    ValidateDate_IT2 = False
   msgbox_1398AATI = msgbox_1398AATI + "* Date of credit into Govt Account must be a valid dd/mm/yyyy format at Sr. No. " & i & " in Schedule IT" & Chr(13)
    Exit Function
End If

If Not chkCompulsory_1398A(DateCredit_IT2(i)) Then
   ' msgbox_1398AATI = msgbox_1398AATI + "* Please enter the Date of credit into Govt Account at Sr. No. " & i & " in Schedule IT" & Chr(13)
    'Added by Shrutika
    msgbox_1398AATI = msgbox_1398AATI + "* Please enter the Date of credit into Govt Account at Sr. No. " & i & " in Schedule IT-2" & Chr(13)
    ValidateDate_IT2 = False
    Exit Function
End If
'Newly changed from 2023 to 2024
'If Not ChkMinInclusiveDate_1398A(Dformat(DateCredit_IT2(i), ""), "2024-04-01") Then ''
'Shrutika_Ur
'If Not ChkMinInclusiveDate_1398A(Dformat(DateCredit_IT2(i), ""), "2023-04-01") Then ''
If Not ChkMinInclusiveDate_1398A(Dformat(DateCredit_IT2(i), ""), "2024-04-01") Then '' 'UR_Shrutika 03/11/2025

   ' msgbox_1398AATI = msgbox_1398AATI + "*Date of credit to Central government at Sr. No.  " & i & " cannot be prior to 01/04/2024 in Schedule IT" & Chr(13)
    'Shrutika_Ur
   ' msgbox_1398AATI = msgbox_1398AATI + "* ""Please enter valid date in dd/mm/yyyy format in schedule IT. Also date of credit to Central government cannot be prior to 01/04/2023 in schedule IT.""" & Chr(13)
     msgbox_1398AATI = msgbox_1398AATI + "* ""Please enter valid date in dd/mm/yyyy format in schedule IT. Also date of credit to Central government cannot be prior to 01/04/2024 in schedule IT.""" & Chr(13)
 
    ValidateDate_IT2 = False
    Exit Function
End If
Next
End Function


Function CheckNoticeDateBefore(dob As Variant) As Boolean
On Error Resume Next

    CheckNoticeDateBefore = True
    If Len(dob) > 0 Then
        If Not FormatNCheckDate(dob) Then
            CheckNoticeDateBefore = False
           fmsgboxStatus ("Notice Date in Sheet : PART A-GENERAL  must be a valid dd/mm/yyyy format")
            Exit Function
        End If
        
       ' If EfilingCommon.checkFirstDateBefore(dob, "31/03/2024") Then ' Chandru PAG Changes
        'Added by Shrutika(AY-2025)
'        If EfilingCommon.checkFirstDateBefore(dob, "31/03/2025") Then

'by sadineni on 30/01/2026 for AY 2026-27
         If EfilingCommon.checkFirstDateBefore(dob, "31/03/2026") Then
          '  fmsgboxStatus ("Notice Date in Sheet : PART A-GENERAL  should be on or after 01/04/2024")
'             fmsgboxStatus ("Notice Date in Sheet : PART A-GENERAL  should be on or after 01/04/2025")
              fmsgboxStatus ("Notice Date in Sheet : PART A-GENERAL  should be on or after 01/04/2026")
            CheckNoticeDateBefore = False
            Exit Function
        Else
                noticeDate = dob
       End If
    End If
 End Function
 
 Sub chekUniqueNo()
   If Not CheckSpecialCharacterSlash(Sheet1.Range("Sheet1.UniqueNo").value) Then
     fmsgboxStatus "Spacial Characters not allowed "
     Range(Target.Address).value = ""
    End If
 End Sub
 
 
 Sub PartA_139LockUnlockTable(ByVal Target As Range)
    Application.EnableEvents = False
    Sheet201.Unprotect Password:=getmsgstate
    'If Target.value = "2024-25" Or Target.value = "2025-26" Then
    'Shrutika_Ur
  ' If Target.value = "2025-26" Or Target.value = "2026-27" Then
    If Target.value = "2026-27" Or Target.value = "2027-28" Then 'UR_03/11/2025)
   '--------------
        Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
        Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
        Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.Locked = False
        Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
    Else
        Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
        Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
        Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
    
        Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.ClearContents
        Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.Locked = True
        Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
    End If
    Application.EnableEvents = True
    Sheet201.Protect Password:=getmsgstate
End Sub


Sub ReturnDropdownA23_24(RowIndex As Long)
On Error GoTo endline
Application.EnableEvents = False
Sheet201.Unprotect Password:=getmsgstate
'Commented by Shrutika
'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).value = "2024-25" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
'        'Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'        If (Dformat("01/01/2025", "") <= Dformat(Sheet52.Range("sheet9.Date").value, "")) And (Dformat("31/03/2025", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AY23Return"
'        ElseIf (Dformat("   ", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'        End If
'    End If
'End If
'
'
'
'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).value = "2025-26" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
'        'Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'        If (Dformat("01/01/2026", "") <= Dformat(Sheet52.Range("sheet9.Date").value, "")) And (Dformat("31/03/2026", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AY23Return"
'        ElseIf (Dformat("31/12/2025", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'        End If
'    End If
'End If
'
''Shrutika_Ur
''If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).value = "2025-26" Then
''    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).value = "No" Then
''        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
''        'Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
''        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
''        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
''    Else
''
''        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
''        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
''        If (Dformat("31/12/2025", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
''            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
''            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
''            'Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AfterDateReturn"
''        ElseIf (Dformat("01/01/2026", "") < Dformat(Sheet52.Range("sheet9.Date").value, "")) Or (Dformat("31/03/2026", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
''            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
''            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AfterDateReturn"
''           ' Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
''        End If
''    End If
''End If
'
'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).value = "2026-27" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
'        'Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'        If (Dformat("31/12/2026", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'        ElseIf (Dformat("01/01/2027", "") < Dformat(Sheet52.Range("sheet9.Date").value, "")) Or (Dformat("31/03/2027", "") >= Dformat(Sheet52.Range("sheet9.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AfterDateReturn"
'        End If
'    End If
'End If
endline:
Sheet201.Protect Password:=getmsgstate
    Application.EnableEvents = True
End Sub
Sub x()
Application.EnableEvents = True
End Sub
Function ChkMaxDOBDateDef(dob As Variant, maxDefinedDOB As Variant) As Boolean
On Error Resume Next
     ChkMaxDOBDateDef = True
     If Len(dob) > 0 Then
        If val(Mid(dob, 7, 4)) > 2025 Then 'Chetan - changed year on 20/01/2026 for AY 2026-27
                    ChkMaxDOBDateDef = False
                    Exit Function
        ElseIf val(Mid(dob, 7, 4)) = 2025 Then 'Chetan - changed year on 20/01/2026 for AY 2026-27
            If val(Mid(dob, 4, 2)) >= 4 Then
                If val(Mid(dob, 1, 2)) >= 1 Then
                    ChkMaxDOBDateDef = False
                    Exit Function
                End If
            End If
        End If
     End If
End Function


Function calculateAge(dob As Variant) As Long

On Error Resume Next
'AY_2024-25 OLD
'calculateAge = (2024) - val(Mid(dob, 7, 4)) 'Year Changed - Chandru
'AY_2025_26 Change
'calculateAge = (2025) - val(Mid(dob, 7, 4)) 'Year Changed - Malli
'AY_2026-27 Change
calculateAge = (2026) - val(Mid(dob, 7, 4)) 'Year Changed - Chetan on 22/01/2026
'--------------------------------
    If 4 < val(Mid(dob, 4, 2)) Then
        calculateAge = calculateAge - 1
    ElseIf val(Mid(dob, 4, 2)) = 4 And 1 < val(Mid(dob, 1, 2)) Then
                calculateAge = calculateAge - 1
    End If

End Function

Function LTCG_SaleTrfDate(ByVal Target As Range)
    If (Target.value <> "") Then
        'If Not ChkMinInclusiveDate(Dformat(Target.value, ""), "2021-04-01") Then
        'CG_E4_AY_2023-24 OLD  'Malli
        'If Not ChkMinInclusiveDate(Dformat(Target.value, ""), "2022-04-01") Then
        'CG_C4_AY_2024-25 Change
        If Not ChkMinInclusiveDate(Dformat(Target.value, ""), "2023-04-01") Then
            'fmsgbox ("*Date of sale should not be before 01/04/2021 for AY 2022-23 B1 In schedule CG")
            'CG_E4_AY_2023-24 OLD
            'fmsgbox ("*Date of sale should not be before 01/04/2022 for AY 2022-23 B1 In schedule CG")
            'CG_C4_AY_2024-25 Change
            fmsgbox ("*Date of sale should not be before 01/04/2023 for AY 2024-25 B1 In schedule CG")
            Target.value = ""
        End If
    End If
End Function


Sub az()
Application.EnableEvents = True

End Sub


'Change-20.12.2022.101.TCS.7A
Function ValidateTCS1() As Boolean
ValidateTCS1 = True

If Not ValidateMandatoryTCS1() Then ValidateTCS1 = False
If Not ValidateTCS1ClaimedOwnHands Then ValidateTCS1 = False
If Not ValidateTCS1TAN Then ValidateTCS1 = False
If Not ValidateTCS1PANOP Then ValidateTCS1 = False
If Not ValidateTCS1ClaimedOPPAN Then ValidateTCS1 = False

End Function
Function ValidateMandatoryTCS1() As Boolean
ValidateMandatoryTCS1 = True
Dim i As Long


'First row must be filled first before filling second row - the first column must be filled

'Dim firstFilled
'Dim nextFilled
'For i = 1 To Sheet28.Range("TCS1.CreditName").count
'To be done later
'
'
'Next i
For i = 1 To Sheet28.Range("TCS1.CreditName").count

If (Sheet28.Range("TCS1.CreditName").Cells(i, 1).value <> "" And UCase(Sheet28.Range("TCS1.CreditName").Cells(i, 1).value) <> "(SELECT)") Or _
    Sheet28.Range("TCS1.TAN").Cells(i, 1).value <> "" Or Sheet28.Range("TCS1.PANOP").Cells(i, 1).value <> "" Or _
    (Sheet28.Range("TCS1.CollectYear").Cells(i, 1).value <> "" And Sheet28.Range("TCS1.CollectYear").Cells(i, 1).value <> "(Select)") Or _
    Sheet28.Range("TCS1.UnclaimedBFAmount").Cells(i, 1).value <> "" Or Sheet28.Range("TCS1.CollectedOwnHands").Cells(i, 1).value <> "" Or Sheet28.Range("TCS1.Collected37i").Cells(i, 1).value <> "" Or _
    Sheet28.Range("TCS1.ClaimedOwnHands").Cells(i, 1).value <> "" Or Sheet28.Range("TCS1.ClaimedOPTCS").Cells(i, 1).value <> "" Or Sheet28.Range("TCS1.ClaimedOPPAN").Cells(i, 1).value <> "" Then
    'Some thing is filled in the row
    
    If (Sheet28.Range("TCS1.CreditName").Cells(i, 1).value = "" Or UCase(Sheet28.Range("TCS1.CreditName").Cells(i, 1).value) = "(SELECT)") Then
        ValidateMandatoryTCS1 = False
'Change.21.03.2023.102.IDS44
'        msgbox_TCS = msgbox_TCS + ("Please select the drop down for the field ""TCS credit in the name of"" at Sr. No. " & i) & Chr(13)
        msgbox_TCS ("Please select the drop down for the field TCS credit  relating to Self/Other Person in Schedule TCS at Sr. No. " & i) & Chr(13)
'End Change.IDS44
    End If
    
    If Sheet28.Range("TCS1.TAN").Cells(i, 1).value = "" Then
        ValidateMandatoryTCS1 = False
        'msgbox_TCS ("Please enter TAN of Collector at Sr. No. " & i)
        
        'Shrutika
        msgbox_TCS ("Please enter the TAN of Collector in Schedule TCS.") & Chr(13)
        
    End If
    
    'Ankita
   ' If Sheet28.Range("TCS1.CollectYear").Cells(i, 1).value = "" Or Sheet28.Range("TCS1.CollectYear").Cells(i, 1).value = "(Select)" Then
      'Added by Riyaz
      If Sheet28.Range("TCS1.CollectYear").Locked = False Then
      If Sheet28.Range("TCS1.CollectYear").Cells(i, 1).value = "" Then
       ValidateMandatoryTCS1 = False
       msgbox_TCS "Please select Year of tax deduction from drop-down" & Chr(13)
     End If
    End If
    
    If Sheet28.Range("TCS1.PANOP").Cells(i, 1).value = "" And Sheet28.Range("TCS1.PANOP").Cells(i, 1).Locked = False Then
        ValidateMandatoryTCS1 = False
        'msgbox_TCS ("Please enter the PAN of Other Person at Sr. No. " & i)
        'Shrutika
        msgbox_TCS ("Please enter the PAN/ Aadhaar of Other Person in schedule TCS") & Chr(13)
        
    End If
    
    If Sheet28.Range("TCS1.CollectedOwnHands").Cells(i, 1).value = "" And Sheet28.Range("TCS1.CollectedOwnHands").Cells(i, 1).Locked = False Then
        ValidateMandatoryTCS1 = False
        msgbox_TCS ("""Please enter TCS collected in own hands at Sr. No.  & i """)
    End If
    
    If Sheet28.Range("TCS1.Collected37i").Cells(i, 1).value = "" And Sheet28.Range("TCS1.Collected37i").Cells(i, 1).Locked = False Then
        ValidateMandatoryTCS1 = False
        'Change-13.04.2023.103.TCS.01
        'msgbox_TCS ("Please enter TCS collected in the hands of any other person at Sr. No. " & i)
        msgbox_TCS ("Please enter TCS collected in the hands of spouse or any other person at Sr. No. " & i)
    End If
    
    
    
    'These two field are mandatory for each other
    If (Sheet28.Range("TCS1.CollectYear").Cells(i, 1).value <> "" And Sheet28.Range("TCS1.CollectYear").Cells(i, 1).value <> "(Select)") Or _
        Sheet28.Range("TCS1.UnclaimedBFAmount").Cells(i, 1).value <> "" Then
        
        'If Sheet28.Range("TCS1.CollectYear").Cells(i, 1).value = "" And Sheet28.Range("TCS1.CollectYear").Cells(i, 1).Locked = False Then
            
          'SUDHA
        If Sheet28.Range("TCS1.CollectYear").Cells(i, 1).value = "" Or Sheet28.Range("TCS1.CollectYear").Cells(i, 1).value = "(Select)" And Sheet28.Range("TCS1.CollectYear").Cells(i, 1).Locked = False Then
            ValidateMandatoryTCS1 = False
            'msgbox_TCS ("Please select the Financial year in which TCS is collected at Sr. No. " & i)
            
            'Shrutika
            msgbox_TCS ("""Please select the Financial year in which TCS is collected""")
        End If
        
        If Sheet28.Range("TCS1.UnclaimedBFAmount").Cells(i, 1).value = "" And Sheet28.Range("TCS1.UnclaimedBFAmount").Cells(i, 1).Locked = False Then
            ValidateMandatoryTCS1 = False
            msgbox_TCS ("Please enter the Amount of TCS brought forwarded at Sr. No. " & i)
        End If

    End If
    
End If
Next i

End Function

Function ValidateTCS1ClaimedOwnHands() As Boolean
ValidateTCS1ClaimedOwnHands = True

Dim i As Long

For i = 1 To Sheet28.Range("TCS1.CreditName").count
    If Sheet28.Range("TCS1.UnclaimedBFAmount").Cells(i, 1).Locked = False Or Sheet28.Range("TCS1.CollectedOwnHands").Cells(i, 1).Locked = False Then
    
        If CDbl(Sheet28.Range("TCS1.UnclaimedBFAmount").Cells(i, 1).value) + CDbl(Sheet28.Range("TCS1.CollectedOwnHands").Cells(i, 1).value) < CDbl(Sheet28.Range("TCS1.ClaimedOwnHands").Cells(i, 1).value) Then
            ValidateTCS1ClaimedOwnHands = False
            msgbox_TCS ("Amount in field Col 7(i) cannot be more than Col 5 or Col 6(i) as the case may be at Sr. No. " & i)
        End If
    End If
    If (Sheet28.Range("TCS1.TAN").Cells(i, 1).value <> "" And Sheet28.Range("TCS1.ClaimedOwnHands").Cells(i, 1).value = "") Then
            ValidateTCS1ClaimedOwnHands = False
            'Change-13.04.2023.103.TCS.02
            'msgbox_TCS ("Please enter Tax collected in own hands at Sr. No. " & i)
             msgbox_TCS ("Please enter Tax claimed in own hands at Sr. No. " & i)
    End If
Next i
    
    

End Function

Function ValidateTCS1TAN() As Boolean
ValidateTCS1TAN = True

For i = 1 To Sheet28.Range("TCS1.CreditName").count

If Not ValidateTantype_text(Mid(Sheet28.Range("TCS1.TAN").Cells(i, 1).value, 1, 4)) Then
         msgbox_TCS ("Invalid TAN in schedule TCS at Sr. No. " & i & " TAN format should be First 4 alphabets, then 5 digits, then alphabet")
         ValidateTCS1TAN = False
         Exit Function
     End If
     
     If Not IsNumeric(Mid(Sheet28.Range("TCS1.TAN").Cells(i, 1).value, 5, 5)) And Sheet28.Range("TCS1.TAN").Cells(i, 1).value <> "" Then
         msgbox_TCS ("Invalid TAN in schedule TCS at Sr. No. " & i & " TAN format should be First 4 alphabets, then 5 digits, then alphabet")
         ValidateTCS1TAN = False
         Exit Function
     End If
     
     If Not ValidateTantype_text(Right(Sheet28.Range("TCS1.TAN").Cells(i, 1).value, 1)) Then
         msgbox_TCS ("Invalid TAN in schedule TCS at Sr. No. " & i & " TAN format should be First 4 alphabets, then 5 digits, then alphabet")
         ValidateTCS1TAN = False
         Exit Function
     End If
     
     If Not ValidateTANFirstThreeLetters(Sheet28.Range("TCS1.TAN").Cells(i, 1).value) And Sheet28.Range("TCS1.TAN").Cells(i, 1).value <> "" Then
        msgbox_TCS ("Invalid TAN in schedule TCS at Sr. No. " & i & " TAN format should be First 4 alphabets, then 5 digits, then alphabet")
        ValidateTCS1TAN = False
        Exit Function
     End If
     
Next i

End Function

Function ValidateTCS1PANOP() As Boolean
ValidateTCS1PANOP = True

For i = 1 To Sheet28.Range("TCS1.CreditName").count

If Not checkfieldspecialcharacter(Sheet28.Range("TCS1.PANOP").Cells(i, 1).value) Then
         msgbox_TCS ("Invalid PAN in TCS Sr. No." & i & ", PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet")
         ValidateTCS1PANOP = False
         Exit Function
    End If
    
    If Not ValidatePAN(UCase(Sheet28.Range("TCS1.PANOP").Cells(i, 1).value)) Then
         msgbox_TCS ("Invalid PAN in TCS Sr. No." & i & ", PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet")
         ValidateTCS1PANOP = False
         Exit Function
    End If
    
     If Sheet28.Range("TCS1.PANOP").Cells(i, 1).value = Sheet1.Range("sheet1.PAN").value And Sheet1.Range("sheet1.PAN").value <> "" Then
  'Ankita
       'msgbox_TCS ("PAN of Other Person at Sr. No  " & i & " can not be same as Assessee PAN")
        msgbox_TCS ("Other person PAN cannot be same as assesse PAN or verification PAN") & Chr(13)
         ValidateTCS1PANOP = False
         Exit Function
    End If
    
    
Next i

End Function

Function ValidateTCS1ClaimedOPPAN() As Boolean
ValidateTCS1ClaimedOPPAN = True

For i = 1 To Sheet28.Range("TCS1.CreditName").count

If Not checkfieldspecialcharacter(Sheet28.Range("TCS1.ClaimedOPPAN").Cells(i, 1).value) Then
         msgbox_TCS ("Invalid PAN in TCS Sr. No." & i & ", PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet")
         ValidateTCS1ClaimedOPPAN = False
         Exit Function
    End If
    
    If Not ValidatePAN(UCase(Sheet28.Range("TCS1.ClaimedOPPAN").Cells(i, 1).value)) Then
         msgbox_TCS ("Invalid PAN in TCS Sr. No." & i & ", PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet")
         ValidateTCS1ClaimedOPPAN = False
         Exit Function
    End If
    
     If Sheet28.Range("TCS1.ClaimedOPPAN").Cells(i, 1).value = Sheet1.Range("sheet1.PAN").value And Sheet1.Range("sheet1.PAN").value <> "" Then
        msgbox_TCS ("PAN of Other Person at Sr. No  " & i & " can not be same as Assessee PAN")
         ValidateTCS1ClaimedOPPAN = False
         Exit Function
    End If
    
    If (Sheet28.Range("TCS1.ClaimedOPTCS").Cells(i, 1).value <> "" And (Sheet28.Range("TCS1.ClaimedOPPAN").Cells(i, 1).value = "") And Sheet28.Range("TCS1.ClaimedOPPAN").Cells(i, 1).Locked = False) Then
            ValidateTCS1ClaimedOPPAN = False
            'Ankita
          ' msgbox_TCS ("Please enter PAN at Sr.NO." & i)
            msgbox_TCS ("Please enter the PAN/Aadhaar of Other Person in schedule TCS") & Chr(13)
    End If
    
    If (Sheet28.Range("TCS1.ClaimedOPPAN").Cells(i, 1).value <> "" And (Sheet28.Range("TCS1.ClaimedOPTCS").Cells(i, 1).value = "") And Sheet28.Range("TCS1.ClaimedOPTCS").Cells(i, 1).Locked = False) Then
            ValidateTCS1ClaimedOPPAN = False
            'Ankita
         '  msgbox_TCS ("Please enter Tax collected in Claimed in hands of spouse/other person deducted in own hands at Sr.NO." & i)
            msgbox_TCS ("Please enter Tax claimed in own hands") & Chr(13)
    End If
Next i
    
End Function

Sub ChangeTCSTable(ByVal Target As Range)
On Error GoTo endline

Sheet28.Unprotect Password:=getmsgstate
Application.EnableEvents = False

If Not Intersect(Sheet28.Range("TCS1.CreditName"), Sheet28.Range(Target.Address)) Is Nothing Then
    
    If UCase(Mid(Sheet28.Range(Target.Address).value, 1, 1)) = "S" Then
        Sheet28.Range(Replace(Target.Address, "F", "H")).ClearContents
        Sheet28.Range(Replace(Target.Address, "F", "H")).Locked = True
        Sheet28.Range(Replace(Target.Address, "F", "H")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(Target.Address, "F", "K")).ClearContents
        Sheet28.Range(Replace(Target.Address, "F", "K")).Locked = False
        Sheet28.Range(Replace(Target.Address, "F", "K")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "F", "L")).ClearContents
        Sheet28.Range(Replace(Target.Address, "F", "L")).Locked = True
        Sheet28.Range(Replace(Target.Address, "F", "L")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(Target.Address, "F", "N")).ClearContents
        Sheet28.Range(Replace(Target.Address, "F", "N")).Locked = False
        Sheet28.Range(Replace(Target.Address, "F", "N")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "F", "O")).ClearContents
        Sheet28.Range(Replace(Target.Address, "F", "O")).Locked = False
        Sheet28.Range(Replace(Target.Address, "F", "O")).Interior.Color = "&HCCFFCC"

    ElseIf UCase(Mid(Sheet28.Range(Target.Address).value, 1, 1)) = "O" Then
        Sheet28.Range(Replace(Target.Address, "F", "H")).ClearContents
        Sheet28.Range(Replace(Target.Address, "F", "H")).Locked = False
        Sheet28.Range(Replace(Target.Address, "F", "H")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "F", "K")).ClearContents
        Sheet28.Range(Replace(Target.Address, "F", "K")).Locked = True
        Sheet28.Range(Replace(Target.Address, "F", "K")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(Target.Address, "F", "L")).ClearContents
        Sheet28.Range(Replace(Target.Address, "F", "L")).Locked = False
        Sheet28.Range(Replace(Target.Address, "F", "L")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "F", "N")).ClearContents
        Sheet28.Range(Replace(Target.Address, "F", "N")).Locked = True
        Sheet28.Range(Replace(Target.Address, "F", "N")).Interior.Color = "&HD8D8D8"
        
        Sheet28.Range(Replace(Target.Address, "F", "O")).ClearContents
        Sheet28.Range(Replace(Target.Address, "F", "O")).Locked = True
        Sheet28.Range(Replace(Target.Address, "F", "O")).Interior.Color = "&HD8D8D8"
        
    Else
        Sheet28.Range(Replace(Target.Address, "F", "H")).ClearContents
        Sheet28.Range(Replace(Target.Address, "F", "H")).Locked = False
        Sheet28.Range(Replace(Target.Address, "F", "H")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "F", "I")).ClearContents
        Sheet28.Range(Replace(Target.Address, "F", "I")).Locked = False
        Sheet28.Range(Replace(Target.Address, "F", "I")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "F", "J")).ClearContents
        Sheet28.Range(Replace(Target.Address, "F", "J")).Locked = False
        Sheet28.Range(Replace(Target.Address, "F", "J")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "F", "K")).ClearContents
        Sheet28.Range(Replace(Target.Address, "F", "K")).Locked = False
        Sheet28.Range(Replace(Target.Address, "F", "K")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "F", "L")).ClearContents
        Sheet28.Range(Replace(Target.Address, "F", "L")).Locked = False
        Sheet28.Range(Replace(Target.Address, "F", "L")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "F", "N")).ClearContents
        Sheet28.Range(Replace(Target.Address, "F", "N")).Locked = False
        Sheet28.Range(Replace(Target.Address, "F", "N")).Interior.Color = "&HCCFFCC"
        
        Sheet28.Range(Replace(Target.Address, "F", "O")).ClearContents
        Sheet28.Range(Replace(Target.Address, "F", "O")).Locked = False
        Sheet28.Range(Replace(Target.Address, "F", "O")).Interior.Color = "&HCCFFCC"
        
    End If

End If

If Not Intersect(Sheet28.Range("TCS1.UnclaimedBFAmount"), Sheet28.Range(Target.Address)) Is Nothing Then

    If UCase(Mid(Sheet28.Range(Replace(Target.Address, "J", "F")).value, 1, 1)) = "S" Then
        If Range(Target.Address).value = 0 Then
            Sheet28.Range(Replace(Target.Address, "J", "K")).ClearContents
            Sheet28.Range(Replace(Target.Address, "J", "K")).Locked = False
            Sheet28.Range(Replace(Target.Address, "J", "K")).Interior.Color = "&HCCFFCC"
        Else
            Sheet28.Range(Replace(Target.Address, "J", "K")).ClearContents
            Sheet28.Range(Replace(Target.Address, "J", "K")).Locked = True
            Sheet28.Range(Replace(Target.Address, "J", "K")).Interior.Color = "&HD8D8D8"
            
            If Sheet28.Range(Replace(Target.Address, "J", "M")).value <> "" And (Target.value < Sheet28.Range(Replace(Target.Address, "J", "M"))) Then
                MsgBox ("Amount in field Col 7(i) cannot be more than Col 5 or Col 6(i) in Schedule TCS")
                Target.value = ""
                
                Sheet28.Range(Replace(Target.Address, "J", "K")).ClearContents
                Sheet28.Range(Replace(Target.Address, "J", "K")).Locked = False
                Sheet28.Range(Replace(Target.Address, "J", "K")).Interior.Color = "&HCCFFCC"
            End If
         
        End If
    
    ElseIf UCase(Mid(Sheet28.Range(Replace(Target.Address, "J", "F")).value, 1, 1)) = "O" Then
        If Range(Target.Address).value = 0 Then
            Sheet28.Range(Replace(Target.Address, "J", "L")).ClearContents
            Sheet28.Range(Replace(Target.Address, "J", "L")).Locked = False
            Sheet28.Range(Replace(Target.Address, "J", "L")).Interior.Color = "&HCCFFCC"
        Else
          
            Sheet28.Range(Replace(Target.Address, "J", "L")).ClearContents
            Sheet28.Range(Replace(Target.Address, "J", "L")).Locked = True
            Sheet28.Range(Replace(Target.Address, "J", "L")).Interior.Color = "&HD8D8D8"

            If Sheet28.Range(Replace(Target.Address, "J", "M")).value <> "" And (Target.value < Sheet28.Range(Replace(Target.Address, "J", "M"))) Then
                MsgBox ("Amount in field Col 7(i) cannot be more than Col 5 + TCS field of Col 6(i)")
                Target.value = ""
                
                Sheet28.Range(Replace(Target.Address, "J", "L")).ClearContents
                Sheet28.Range(Replace(Target.Address, "J", "L")).Locked = False
                Sheet28.Range(Replace(Target.Address, "J", "L")).Interior.Color = "&HCCFFCC"
            End If
        End If
        
    End If
End If


If Not Intersect(Sheet28.Range("TCS1.CollectedOwnHands"), Sheet28.Range(Target.Address)) Is Nothing Then
        If Range(Target.Address).value = 0 Then
            
            Sheet28.Range(Replace(Target.Address, "K", "J")).ClearContents
            Sheet28.Range(Replace(Target.Address, "K", "J")).Locked = False
            Sheet28.Range(Replace(Target.Address, "K", "J")).Interior.Color = "&HCCFFCC"
            
            Sheet28.Range(Replace(Target.Address, "K", "I")).ClearContents
            Sheet28.Range(Replace(Target.Address, "K", "I")).Locked = False
            Sheet28.Range(Replace(Target.Address, "K", "I")).Interior.Color = "&HCCFFCC"
            
        Else
            
            Sheet28.Range(Replace(Target.Address, "K", "J")).ClearContents
            Sheet28.Range(Replace(Target.Address, "K", "J")).Locked = True
            Sheet28.Range(Replace(Target.Address, "K", "J")).Interior.Color = "&HD8D8D8"
            
            Sheet28.Range(Replace(Target.Address, "K", "I")).ClearContents
            Sheet28.Range(Replace(Target.Address, "K", "I")).Locked = True
            Sheet28.Range(Replace(Target.Address, "K", "I")).Interior.Color = "&HD8D8D8"
            
            If Sheet28.Range(Replace(Target.Address, "K", "M")).value <> "" And (Target.value < Sheet28.Range(Replace(Target.Address, "K", "M"))) Then
                MsgBox ("Amount in field Col 7(i) cannot be more than Col 5 + TCS field of Col 6(i)")
                Target.value = ""
                
                Sheet28.Range(Replace(Target.Address, "K", "J")).ClearContents
                Sheet28.Range(Replace(Target.Address, "K", "J")).Locked = False
                Sheet28.Range(Replace(Target.Address, "K", "J")).Interior.Color = "&HCCFFCC"
                
                Sheet28.Range(Replace(Target.Address, "K", "I")).ClearContents
                Sheet28.Range(Replace(Target.Address, "K", "I")).Locked = False
                Sheet28.Range(Replace(Target.Address, "K", "I")).Interior.Color = "&HCCFFCC"
            End If
            
    End If
End If

If Not Intersect(Sheet28.Range("TCS1.Collected37i"), Sheet28.Range(Target.Address)) Is Nothing Then
        If Range(Target.Address).value = 0 Then
            
            Sheet28.Range(Replace(Target.Address, "L", "J")).ClearContents
            Sheet28.Range(Replace(Target.Address, "L", "J")).Locked = False
            Sheet28.Range(Replace(Target.Address, "L", "J")).Interior.Color = "&HCCFFCC"
            
            Sheet28.Range(Replace(Target.Address, "L", "I")).ClearContents
            Sheet28.Range(Replace(Target.Address, "L", "I")).Locked = False
            Sheet28.Range(Replace(Target.Address, "L", "I")).Interior.Color = "&HCCFFCC"
            
        Else
            
            Sheet28.Range(Replace(Target.Address, "L", "J")).ClearContents
            Sheet28.Range(Replace(Target.Address, "L", "J")).Locked = True
            Sheet28.Range(Replace(Target.Address, "L", "J")).Interior.Color = "&HD8D8D8"
            
            Sheet28.Range(Replace(Target.Address, "L", "I")).ClearContents
            Sheet28.Range(Replace(Target.Address, "L", "I")).Locked = True
            Sheet28.Range(Replace(Target.Address, "L", "I")).Interior.Color = "&HD8D8D8"
            
    End If
End If

If Not Intersect(Sheet28.Range("TCS1.ClaimedOwnHands"), Sheet28.Range(Target.Address)) Is Nothing Then

    If UCase(Mid(Sheet28.Range(Replace(Target.Address, "M", "F")).value, 1, 1)) = "S" Then
         If Sheet28.Range(Replace(Target.Address, "M", "J")).Locked = False And Sheet28.Range(Replace(Target.Address, "M", "J")).value <> "" And (Target.value > Sheet28.Range(Replace(Target.Address, "M", "J"))) Or _
            Sheet28.Range(Replace(Target.Address, "M", "K")).Locked = False And Sheet28.Range(Replace(Target.Address, "M", "K")).value <> "" And (Target.value > Sheet28.Range(Replace(Target.Address, "M", "K"))) Then
            MsgBox ("Amount in field Col 7(i) cannot be more than Col 5 or Col 6(i) in Schedule TCS")
            Target.value = ""
         End If
        
    ElseIf UCase(Mid(Sheet28.Range(Replace(Target.Address, "M", "F")).value, 1, 1)) = "O" Then
         If Sheet28.Range(Replace(Target.Address, "M", "J")).Locked = False And Sheet28.Range(Replace(Target.Address, "M", "J")).value <> "" And (Target.value > Sheet28.Range(Replace(Target.Address, "M", "J"))) Then
            'Sheet28.Range(Replace(Target.Address, "M", "L")).Locked = False And (Target.Value > Sheet28.Range(Replace(Target.Address, "M", "L"))) Then
            MsgBox ("Amount in field Col 7(i) cannot be more than Col 5 + TCS field of Col 6(i)")
            Target.value = ""
         End If
    End If
End If

endline:
    Application.EnableEvents = True
    Sheet28.Protect Password:=getmsgstate
End Sub

'Change.21.03.2023.102.IDS.69

'Sub AddRows_TCS1()
'    Dim numberofrows As Long
'    Dim rngname_TCS1 As Variant
'    rngname_TCS1 = "TCS1.CreditName;TCS1.TAN;TCS1.PANOP;TCS1.CollectYear;TCS1.UnclaimedBFAmount;TCS1.CollectedOwnHands;TCS1.Collected37i;TCS1.ClaimedOwnHands;TCS1.ClaimedOPTCS;TCS1.ClaimedOPPAN;TCS1.CreditForward;"
'    SelectLastRow ("TCS1.CreditName")
'    numberofrows = InsertRowsAndFillFormulasBPA()
'    Call ExendRangeNameToTable(numberofrows, rngname_TCS1)
'End Sub
Sub AddRows_TCS1()
Dim vRows As Long
    Sheets("TDS").Activate
    EfilingCommon.DefinedgridNameRange = "TCS1.CreditName||TCS1.TAN||TCS1.PANOP||TCS1.CollectYear||TCS1.UnclaimedBFAmount||TCS1.CollectedOwnHands||TCS1.Collected37i||TCS1.ClaimedOwnHands||TCS1.ClaimedOPTCS||TCS1.ClaimedOPPAN||TCS1.CreditForward"
    ActiveCellRange = EfilingCommon.searchLastRow("TCS1.TAN")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
'End Change.IDS.69

 Sub AddRows_TCS1_import(DiffRows As Long)
    Sheet28.Activate
    Dim rngname_TCS1 As Variant
    rngname_TCS1 = "TCS1.CreditName;TCS1.TAN;TCS1.PANOP;TCS1.CollectYear;TCS1.UnclaimedBFAmount;TCS1.CollectedOwnHands;TCS1.Collected37i;TCS1.ClaimedOwnHands;TCS1.ClaimedOPTCS;TCS1.ClaimedOPPAN;TCS1.CreditForward;"
    SelectLastRow ("TCS1.CreditName")
    InsertDiffRowsAndFillFormulas (DiffRows)
    Call ExendRangeNameToTable(DiffRows, rngname_TCS1)
End Sub


''Change-22.01.2023.101.S.3B
'Sub SalaryExemptAllowanceDropdown()
'
'        Dim isGovFlag, j As Variant
'        isGovFlag = False
'
'        For j = 1 To Sheet39.Range("SalarySectionCount").value
'            If Sheet39.Range("sheet.EmployerCategory" & j).value = "Central Government" Or _
'            Sheet39.Range("sheet.EmployerCategory" & j).value = "State Government" Or _
'            Sheet39.Range("sheet.EmployerCategory" & j).value = "CG-Pensioners" Or _
'            Sheet39.Range("sheet.EmployerCategory" & j).value = "SG-Pensioners" Then
'                If isGovFlag = False Then
'                   isGovFlag = True
'                End If
'            End If
'        Next j
'        Application.EnableEvents = False
'        Sheet39.Unprotect Password:=getmsgstate
'        If SheetALL.Range("bacValue").value = 1 Then
'            If isGovFlag = True Then
'                Sheet39.Range("Salary.Nature1").Validation.Delete
'                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACYesCGSGPCGPSG"
'                Sheet39.Range("Salary.Nature1").value = "(Select)"
'                Sheet39.Range("Salary.Amount1").ClearContents
'            Else
'                Sheet39.Range("Salary.Nature1").Validation.Delete
'                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACYes"
'                Sheet39.Range("Salary.Nature1").value = "(Select)"
'                Sheet39.Range("Salary.Amount1").ClearContents
'            End If
'        ElseIf SheetALL.Range("bacValue").value = 2 Then
'            If isGovFlag = True Then
'                Sheet39.Range("Salary.Nature1").Validation.Delete
'                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACNoCGSGPCGPSG"
'                Sheet39.Range("Salary.Nature1").value = "(Select)"
'                Sheet39.Range("Salary.Amount1").ClearContents
'            Else
'                Sheet39.Range("Salary.Nature1").Validation.Delete
'                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACNo"
'                Sheet39.Range("Salary.Nature1").value = "(Select)"
'                Sheet39.Range("Salary.Amount1").ClearContents
'            End If
'        End If
'        Sheet39.Protect Password:=getmsgstate
'        Application.EnableEvents = True
'End Sub

Sub SalaryExemptAllowanceDropdown()

        Dim isPGov, isCGSG, isAllPGov, j As Variant

        isPGov = False
        isCGSG = False

        isAllPGov = False

        For j = 1 To Sheet39.Range("SalarySectionCount").value
            
            If Sheet39.Range("sheet.EmployerCategory" & j).value = "CG-Pensioners" Or _
            Sheet39.Range("sheet.EmployerCategory" & j).value = "SG-Pensioners" Then
                If isPGov = False Then
                   isPGov = True
                End If
            End If
            
            If Sheet39.Range("sheet.EmployerCategory" & j).value = "Central Government" Or _
            Sheet39.Range("sheet.EmployerCategory" & j).value = "State Government" Then
                If isCGSG = False Then
                   isCGSG = True
                End If
            End If
            
            If Sheet39.Range("sheet.EmployerCategory" & j).value = "PSU-Pensioners" Or _
            Sheet39.Range("sheet.EmployerCategory" & j).value = "Others-Pensioners" Then
                If isAllPGov = False Then
                   isAllPGov = True
                End If
            End If
            
        Next j
        
        Sheet39.Unprotect Password:=getmsgstate
        
        If SheetALL.Range("bacValue").value = 1 Then
            
            If isPGov = True And isCGSG = True Then
                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACYesCGSGPCGPSG"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
                
            End If
            
            If isCGSG = True And isPGov = False Then
                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACYesCGSG"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
            End If
            
            If isCGSG = False And isPGov = True Then
                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACYesPCGPSG"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
            End If
            
            If isCGSG = False And isPGov = False And isAllPGov = False Then

                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACYes"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
            End If
            
            If isCGSG = False And isPGov = False And isAllPGov = True Then
                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACYesAllPGov"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
            End If

        ElseIf SheetALL.Range("bacValue").value = 2 Then
            If isCGSG = True And isPGov = True Then
                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACNoCGSGPCGPSG"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
            End If
            
            If isCGSG = True And isPGov = False Then
                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACNoCGSG"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
            End If
            
            If isCGSG = False And isPGov = True Then
                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACNoPCGPSG"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
            End If
                

            If isCGSG = False And isPGov = False And isAllPGov = False Then

                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACNo"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
            End If
            
            If isCGSG = False And isPGov = False And isAllPGov = True Then
                Sheet39.Range("Salary.Nature1").Validation.Delete
                Sheet39.Range("Salary.Nature1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AllowanceBACNoAllPGov"
                Sheet39.Range("Salary.Nature1").value = "(Select)"
                Sheet39.Range("Salary.Amount1").ClearContents
            End If

        End If
        Sheet39.Protect Password:=getmsgstate
    
End Sub




Function ValidateShortTermLongTermTotal_8b() As Boolean
    Dim i As Long
    
    ValidateShortTermLongTermTotal_8b = True
    TotalCapGains_8b = Sheet7.Range("Sheet8b.TotalCapGains").value
    
    If Len(TotalCapGains_8b) > 14 Then
        msgbox8b ("Sum of Short-Term/Long Term Capital Gain should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateShortTermLongTermTotal_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(TotalCapGains_8b)
        If Not IsNumeric(TotalCapGains_8b) Then
            msgbox8b ("Sum of Short-Term/Long Term Capital Gain must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateShortTermLongTermTotal_8b = False
            Exit Function
        End If
    Next
End Function

Function ValidateCapGains30Per115BBH_8b() As Boolean
    Dim i As Long
    
    ValidateCapGains30Per115BBH_8b = True
    CapGains30percent_8b = Sheet7.Range("Sheet8b.CapGains30percent").value
    
    If Len(CapGains30percent_8b) > 14 Then
        msgbox8b ("Capital gains chargable @30% u/s 115BBH should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateCapGains30Per115BBH_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(CapGains30percent_8b)
        If Not IsNumeric(CapGains30percent_8b) Then
            msgbox8b ("Capital gains chargable @30% u/s 115BBH must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateCapGains30Per115BBH_8b = False
            Exit Function
        End If
    Next
End Function
Function ValidateTotalCapGains_8b() As Boolean
    Dim i As Long
    
    ValidateTotalCapGains_8b = True
    CapGainsTotal_8b = Sheet7.Range("Sheet8b.CapGainsTotal").value
    
    If Len(CapGainsTotal_8b) > 14 Then
        msgbox8b ("TotalCapGains should not greater than 14 digits in Sheet PART - TI - TTI ")
        ValidateTotalCapGains_8b = False
        Exit Function
    End If
    
    For i = 1 To Len(CapGainsTotal_8b)
        If Not IsNumeric(CapGainsTotal_8b) Then
            msgbox8b ("TotalCapGains must contain only digits from 0 to 9 in Sheet PART - TI - TTI ")
            ValidateTotalCapGains_8b = False
            Exit Function
        End If
    Next
End Function

Sub AnyOtherDeductions80CCH(ByVal Target As Range)
    
    Dim NatureFlag
    NatureFlag = True
    
    For i = 1 To Sheet39.Range("SalarySectionCount").value
        If Sheet39.Range("sheet.EmployerCategory" & i).value <> "Central Government" And Sheet39.Range("sheet.EmployerCategory" & i).value <> "(Select)" And Sheet39.Range("sheet.EmployerCategory" & i).value <> "" Then
            NatureFlag = False
        End If
    Next
    
    Application.EnableEvents = False
    Sheet19.Unprotect Password:=getmsgstate
    
    'added by Chetan C M on 29/01/2026 for AY 2026-27
    'start--
''    If NatureFlag = False And (Dformat((Sheet1.Range("sheet1.DOB").value), "") < "2001-03-31" Or Dformat((Sheet1.Range("sheet1.DOB").value), "") > "2007-04-01") Then
'    If NatureFlag = False And (Dformat((Sheet1.Range("sheet1.DOB").value), "") < "2001-03-31" Or Dformat((Sheet1.Range("sheet1.DOB").value), "") > "2008-04-01") Then
    If NatureFlag = False And (Dformat((Sheet1.Range("sheet1.DOB").value), "") < "1999-03-31" Or Dformat((Sheet1.Range("sheet1.DOB").value), "") > "2009-04-01") Then
    '--end
    
        'Sheet19.Range("Section80CCH").MergeArea.ClearContents
        'Sheet19.Range("Section80CCH").MergeArea.Locked = True
        'Sheet19.Range("Section80CCH").MergeArea.Interior.Color = "&HD8D8D8"
        
        Sheet19.Range("Anyother80CCH").MergeArea.value = 0
        Sheet19.Range("Anyother80CCH").MergeArea.Locked = True
        Sheet19.Range("Anyother80CCH").MergeArea.Interior.Color = "&HD8D8D8"
    Else
        'Sheet19.Range("Section80CCH").MergeArea.Locked = False
        'Sheet19.Range("Section80CCH").MergeArea.Interior.Color = "&HCCFFCC"
'        Sheet19.Range("Section80CCH").MergeArea.value = "80CCH (applicable for only Central govt emplyees)"
        'Sheet19.Range("Section80CCH").MergeArea.value = "(Select)"
        
        Sheet19.Range("Anyother80CCH").MergeArea.value = 0
        Sheet19.Range("Anyother80CCH").MergeArea.Locked = False
        Sheet19.Range("Anyother80CCH").MergeArea.Interior.Color = "&HCCFFCC"
    End If

endline:
    Sheet19.Protect Password:=getmsgstate
    Application.EnableEvents = True
End Sub

'Sub TCS_Change(ByVal Target As Range)
'    If UCase(Mid(Sheet28.Range(Target.Address).value, 1, 1)) = "S" Then
'        Sheet28.Range(Replace(Target.Address, "E", "G")).ClearContents
'        Sheet28.Range(Replace(Target.Address, "E", "G")).Locked = True
'        Sheet28.Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HD8D8D8"
'
'        Sheet28.Range(Replace(Target.Address, "E", "K")).ClearContents
'        Sheet28.Range(Replace(Target.Address, "E", "K")).Locked = True
'        Sheet28.Range(Replace(Target.Address, "E", "K")).Interior.Color = "&HD8D8D8"
'
'        Sheet28.Range(Replace(Target.Address, "E", "L")).ClearContents
'        Sheet28.Range(Replace(Target.Address, "E", "L")).Locked = False
'        Sheet28.Range(Replace(Target.Address, "E", "L")).Interior.Color = "&HCCFFCC"
'
'        Sheet28.Range(Replace(Target.Address, "E", "M")).ClearContents
'        Sheet28.Range(Replace(Target.Address, "E", "M")).Locked = False
'        Sheet28.Range(Replace(Target.Address, "E", "M")).Interior.Color = "&HCCFFCC"
'
'        Sheet28.Range(Replace(Target.Address, "E", "N")).ClearContents
'        Sheet28.Range(Replace(Target.Address, "E", "N")).Locked = True
'        Sheet28.Range(Replace(Target.Address, "E", "N")).Interior.Color = "&HD8D8D8"
'
'    ElseIf UCase(Mid(Sheet28.Range(Target.Address).value, 1, 1)) = "O" Then
'        Sheet28.Range(Replace(Target.Address, "E", "G")).ClearContents
'        Sheet28.Range(Replace(Target.Address, "E", "G")).Locked = False
'        Sheet28.Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HCCFFCC"
'
'        Sheet28.Range(Replace(Target.Address, "E", "K")).ClearContents
'        Sheet28.Range(Replace(Target.Address, "E", "K")).Locked = False
'        Sheet28.Range(Replace(Target.Address, "E", "K")).Interior.Color = "&HCCFFCC"
'
'        Sheet28.Range(Replace(Target.Address, "E", "L")).ClearContents
'        Sheet28.Range(Replace(Target.Address, "E", "L")).Locked = True
'        Sheet28.Range(Replace(Target.Address, "E", "L")).Interior.Color = "&HD8D8D8"
'
'        Sheet28.Range(Replace(Target.Address, "E", "M")).ClearContents
'        Sheet28.Range(Replace(Target.Address, "E", "M")).Locked = True
'        Sheet28.Range(Replace(Target.Address, "E", "M")).Interior.Color = "&HD8D8D8"
'
'        Sheet28.Range(Replace(Target.Address, "E", "N")).ClearContents
'        Sheet28.Range(Replace(Target.Address, "E", "N")).Locked = False
'        Sheet28.Range(Replace(Target.Address, "E", "N")).Interior.Color = "&HCCFFCC"
'
'    Else
'        Sheet28.Range(Replace(Target.Address, "E", "N")).ClearContents
'        Sheet28.Range(Replace(Target.Address, "E", "N")).Locked = True
'        Sheet28.Range(Replace(Target.Address, "E", "N")).Interior.Color = "&HD8D8D8"
'
'        Sheet28.Range(Replace(Target.Address, "E", "K")).ClearContents
'        Sheet28.Range(Replace(Target.Address, "E", "K")).Locked = False
'        Sheet28.Range(Replace(Target.Address, "E", "K")).Interior.Color = "&HCCFFCC"
'
'        Sheet28.Range(Replace(Target.Address, "E", "L")).ClearContents
'        Sheet28.Range(Replace(Target.Address, "E", "L")).Locked = False
'        Sheet28.Range(Replace(Target.Address, "E", "L")).Interior.Color = "&HCCFFCC"
'
'        Sheet28.Range(Replace(Target.Address, "E", "M")).ClearContents
'        Sheet28.Range(Replace(Target.Address, "E", "M")).Locked = False
'        Sheet28.Range(Replace(Target.Address, "E", "M")).Interior.Color = "&HCCFFCC"
'
'        Sheet28.Range(Replace(Target.Address, "E", "G")).ClearContents
'        Sheet28.Range(Replace(Target.Address, "E", "G")).Locked = True
'        Sheet28.Range(Replace(Target.Address, "E", "G")).Interior.Color = "&HD8D8D8"
'
'    End If
'End Sub
'Sub TCS_UnclaimedBFAmount(ByVal Target As Range)
'Sheet28.Unprotect Password:=getmsgstate
'Application.EnableEvents = False
'        If Range(Target.Address).value = 0 Then
'
'            Sheet28.Range(Replace(Target.Address, "I", "J")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "I", "J")).Locked = False
'            Sheet28.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HCCFFCC"
'
'            Sheet28.Range(Replace(Target.Address, "I", "K")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "I", "K")).Locked = False
'            Sheet28.Range(Replace(Target.Address, "I", "K")).Interior.Color = "&HCCFFCC"
'
'        Else
'
'            Sheet28.Range(Replace(Target.Address, "I", "K")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "I", "K")).Locked = True
'            Sheet28.Range(Replace(Target.Address, "I", "K")).Interior.Color = "&HD8D8D8"
'
'            Sheet28.Range(Replace(Target.Address, "I", "J")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "I", "J")).Locked = True
'            Sheet28.Range(Replace(Target.Address, "I", "J")).Interior.Color = "&HD8D8D8"
'    End If
'Application.EnableEvents = True
'Sheet28.Protect Password:=getmsgstate
'End Sub
'
'
'Sub TCS_CollectedOwnHands(ByVal Target As Range)
'Sheet28.Unprotect Password:=getmsgstate
'Application.EnableEvents = False
'        If Range(Target.Address).value = 0 Then
'
'            Sheet28.Range(Replace(Target.Address, "J", "I")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "J", "I")).Locked = False
'            Sheet28.Range(Replace(Target.Address, "J", "I")).Interior.Color = "&HCCFFCC"
'
'            Sheet28.Range(Replace(Target.Address, "J", "L")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "J", "L")).Locked = False
'            Sheet28.Range(Replace(Target.Address, "J", "L")).Interior.Color = "&HCCFFCC"
'
'        Else
'
'            Sheet28.Range(Replace(Target.Address, "J", "I")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "J", "I")).Locked = True
'            Sheet28.Range(Replace(Target.Address, "J", "I")).Interior.Color = "&HD8D8D8"
'
'            Sheet28.Range(Replace(Target.Address, "J", "L")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "J", "L")).Locked = True
'            Sheet28.Range(Replace(Target.Address, "J", "L")).Interior.Color = "&HD8D8D8"
'    End If
'Application.EnableEvents = True
'Sheet28.Protect Password:=getmsgstate
'End Sub
'
'
'Sub TCS_ClaimedOwnHands(ByVal Target As Range)
'Sheet28.Unprotect Password:=getmsgstate
'Application.EnableEvents = False
'        If Range(Target.Address).value = 0 Then
'
'            Sheet28.Range(Replace(Target.Address, "L", "H")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "L", "H")).Locked = False
'            Sheet28.Range(Replace(Target.Address, "L", "H")).Interior.Color = "&HCCFFCC"
'
'            Sheet28.Range(Replace(Target.Address, "L", "I")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "L", "I")).Locked = False
'            Sheet28.Range(Replace(Target.Address, "L", "I")).Interior.Color = "&HCCFFCC"
'
'            Sheet28.Range(Replace(Target.Address, "L", "J")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "L", "J")).Locked = False
'            Sheet28.Range(Replace(Target.Address, "L", "J")).Interior.Color = "&HCCFFCC"
'
'            Sheet28.Range(Replace(Target.Address, "L", "K")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "L", "K")).Locked = False
'            Sheet28.Range(Replace(Target.Address, "L", "K")).Interior.Color = "&HCCFFCC"
'
'        Else
'
'            Sheet28.Range(Replace(Target.Address, "L", "H")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "L", "H")).Locked = True
'            Sheet28.Range(Replace(Target.Address, "L", "H")).Interior.Color = "&HD8D8D8"
'
'            Sheet28.Range(Replace(Target.Address, "L", "I")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "L", "I")).Locked = True
'            Sheet28.Range(Replace(Target.Address, "L", "I")).Interior.Color = "&HD8D8D8"
'
'            Sheet28.Range(Replace(Target.Address, "L", "J")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "L", "J")).Locked = True
'            Sheet28.Range(Replace(Target.Address, "L", "J")).Interior.Color = "&HD8D8D8"
'
'            Sheet28.Range(Replace(Target.Address, "L", "K")).ClearContents
'            Sheet28.Range(Replace(Target.Address, "L", "K")).Locked = True
'            Sheet28.Range(Replace(Target.Address, "L", "K")).Interior.Color = "&HD8D8D8"
'
'    End If
'Application.EnableEvents = True
'Sheet28.Protect Password:=getmsgstate
'End Sub

'Change-25.04.2023.101.VIA -> SIT-49420
Sub CalcVIA80CCH()
On Error GoTo endline
Dim blockcount, salaryAsPer17_1, nat80CCH, VIA80CCH As Long
blockcount = 0
salaryAsPer17_1 = 0
nat80CCH = 0
For blockcount = 1 To Sheet39.Range("SalarySectionCount").value
salaryAsPer17_1 = salaryAsPer17_1 + Sheet39.Range("SAL.AllowancesNotExempt" & blockcount).value

nat80CCH = nat80CCH + Application.WorksheetFunction.SumIfs(Sheet39.Range("salary.SalaryAmount" & blockcount), Sheet39.Range("salary.Section" & blockcount), "Contribution made by the central government towards Agnipath scheme as referred  under section 80CCH")

Next blockcount

VIA80CCH = 0.6 * (salaryAsPer17_1 - nat80CCH)
Sheet19.Unprotect Password:=getmsgstate
Application.EnableEvents = False
Sheet19.Range("Anyother80CCH_PreCalc").value = VIA80CCH
Application.EnableEvents = True
Sheet19.Protect Password:=getmsgstate
endline:
End Sub
'------


'Chandru AY2024-25 New changes

Sub CalcVIA80CCH_NewRegme()
On Error GoTo endline
Dim blockcountN, salaryAsPer17_1N, nat80CCHN, VIA80CCHN As Long
blockcountN = 0
salaryAsPer17_1N = 0
nat80CCHN = 0
For blockcountN = 1 To Sheet39.Range("SalarySectionCount").value
salaryAsPer17_1N = salaryAsPer17_1N + Sheet39.Range("SAL.AllowancesNotExempt" & blockcountN).value

nat80CCHN = nat80CCHN + Application.WorksheetFunction.SumIfs(Sheet39.Range("salary.SalaryAmount" & blockcountN), Sheet39.Range("salary.Section" & blockcountN), "Contribution made by the central government towards Agnipath scheme as referred  under section 80CCH")

Next blockcountN

VIA80CCHN = salaryAsPer17_1N - nat80CCHN
'Newly changed by Bindu as per Ayushi
VIA80CCHN = nat80CCHN


Sheet19.Unprotect Password:=getmsgstate
Application.EnableEvents = False
Sheet19.Range("Anyother80CCH_PreCalc_Newreg").value = VIA80CCHN
Application.EnableEvents = True
Sheet19.Protect Password:=getmsgstate
endline:
End Sub



'Change-18.07.2023.103 -> CSV implementation
Function DformatExcel(inputDate As Variant) As String
    DformatExcel = ""
    If inputDate = "" Then
        DformatExcel = ""
    Else
        DformatExcel = Mid(inputDate, 9, 2) & "/" & Mid(inputDate, 6, 2) & "/" & Mid(inputDate, 1, 4)
    End If
End Function


'For SIT-89202 by sai on 04/03/2025

Sub LTCGSecUndDedClaim(RowIndex As Long)
'Target will be Range("STCG.A7_PreviousYer_1")
sPassword = EfilingCommon.getmsgstate
Sheet13.Unprotect Password:=sPassword

Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_DeductionSection_1").Column).value = ""


'Year Changed by Riyaz on 27/01/2026
If Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_PreviousYer_1").Column).value = "2022-23" Then
    Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_DeductionSection_1").Column).Validation.Delete
    Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_DeductionSection_1").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="(Select),54,54F"
End If

'Year Changed by Riyaz on 27/01/2026
If Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_PreviousYer_1").Column).value = "2023-24" Then
    Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_DeductionSection_1").Column).Validation.Delete
    Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_DeductionSection_1").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="(Select),54,54B,54F"
End If

'Year Changed by Riyaz on 27/01/2026
If Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_PreviousYer_1").Column).value = "2024-25" Then
    
    Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_DeductionSection_1").Column).Validation.Delete
    Sheet13.Cells(RowIndex, Sheet13.Range("LTCG.B9_DeductionSection_1").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="(Select),54,54B,54F"
End If

Sheet13.Protect Password:=sPassword
End Sub
Function checkspecialcharacterForPinCode(field As Variant) As Boolean
    Dim i, j As Long
    checkspecialcharacterForPinCode = True
    Dim arr As Variant
    arr = Array(".", "@", "*", "!", "-", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkspecialcharacterForPinCode = False
            Exit Function
        End If
        Next
    Next
End Function

Sub AutoPopulateSecondaryAddress()
    Application.EnableEvents = False
 
    If Sheet1.Range("Secondary_Address").value = "Yes" Then
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.ResidenceNo2").value = Sheet1.Range("sheet1.ResidenceNo").value
        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.LocalityOrArea2").value = Sheet1.Range("sheet1.LocalityOrArea").value
        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.ResidenceName2").value = Sheet1.Range("sheet1.ResidenceName").value
        Sheet1.Range("sheet1.ResidenceName2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.ResidenceName2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.RoadOrStreet2").value = Sheet1.Range("sheet1.RoadOrStreet").value
        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.CityOrTownOrDistrict2").value = Sheet1.Range("sheet1.CityOrTownOrDistrict").value
        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range("sheet1.PinCode").Locked = True Then
            Sheet1.Range("sheet1.PinCode2").MergeArea.ClearContents
            Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = True
            Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = "&HD8D8D8"
        Else
            Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = RGB(217, 217, 217)
            Sheet1.Range("sheet1.PinCode2").value = Sheet1.Range("sheet1.PinCode").value
            Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = True
        End If
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.StateCode2").value = Sheet1.Range("sheet1.StateCode1").value
        Sheet1.Range("sheet1.StateCode2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.StateCode2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.CountryCode2").value = Sheet1.Range("sheet1.CountryCode").value
        Sheet1.Range("sheet1.CountryCode2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.CountryCode2").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range("HASZIP").Locked = True Then
            Sheet1.Range("HASZIP2").MergeArea.ClearContents
            Sheet1.Range("HASZIP2").MergeArea.Locked = True
            Sheet1.Range("HASZIP2").MergeArea.Interior.Color = "&HD8D8D8"
        Else
            Sheet1.Range("HASZIP2").MergeArea.Interior.Color = RGB(217, 217, 217)
            Sheet1.Range("HASZIP2").value = Sheet1.Range("HASZIP").value
            Sheet1.Range("HASZIP2").MergeArea.Locked = True
        End If
'         If Sheet1.Range("HASZIP").value = "Yes" Then
'            Sheet1.Range("sheet1.ZipCode2").value = "XXXXXX"
'            Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
'            Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = "&HD8D8D8"
'        End If
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range("sheet1.ZipCode").Locked = True Then
            If Sheet1.Range("HASZIP").value = "Yes" Then
                Sheet1.Range("sheet1.ZipCode2").value = "XXXXXX"
                Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
                Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = "&HD8D8D8"
            Else
            Sheet1.Range("sheet1.ZipCode2").MergeArea.ClearContents
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = "&HD8D8D8"
            End If
        Else
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = RGB(217, 217, 217)
            Sheet1.Range("sheet1.ZipCode2").value = Sheet1.Range("sheet1.ZipCode").value
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
        End If
Else

        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.ClearContents

        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Range("sheet1.ResidenceNo2").MergeArea.Locked = False

        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.ClearContents

        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Range("sheet1.LocalityOrArea2").MergeArea.Locked = False

        Sheet1.Range("sheet1.ResidenceName2").MergeArea.ClearContents

        Sheet1.Range("sheet1.ResidenceName2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Range("sheet1.ResidenceName2").MergeArea.Locked = False

        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.ClearContents

        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Range("sheet1.RoadOrStreet2").MergeArea.Locked = False

        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.ClearContents

        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Range("sheet1.CityOrTownOrDistrict2").MergeArea.Locked = False

        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("sheet1.PinCode2").MergeArea.ClearContents

        Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = False

        Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("sheet1.StateCode2").value = "(Select)"

        Sheet1.Range("sheet1.StateCode2").MergeArea.Locked = False
 
        Sheet1.Range("sheet1.StateCode2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("sheet1.CountryCode2").value = "(Select)"

        Sheet1.Range("sheet1.CountryCode2").MergeArea.Locked = False

        Sheet1.Range("sheet1.CountryCode2").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("HASZIP2").MergeArea.ClearContents
        
'Konda updated on 05-05-2026--SIT-118819
'        Sheet1.Range("HASZIP2").MergeArea.Locked = False
         Sheet1.Range("HASZIP2").MergeArea.Locked = True

        Sheet1.Range("HASZIP2").MergeArea.Interior.Color = "&HD8D8D8"

        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("sheet1.ZipCode2").MergeArea.ClearContents

'Konda updated on 11-05-2026--SIT-119434
'        Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = False
        Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True

        Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = "&HD8D8D8"

    End If
 
End Sub

'---
Sub test11111()
Application.EnableEvents = True
End Sub
