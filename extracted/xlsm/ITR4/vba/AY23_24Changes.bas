Attribute VB_Name = "AY23_24Changes"
Sub a()
Application.EnableEvents = True
End Sub
Sub lockAllSheet()
Dim sheet As Worksheet

For Each sheet In ThisWorkbook.Sheets
sheet.Protect Password:=getmsgstate
Next sheet

End Sub
Public Function ChkMaxDate(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

     ChkMaxDate = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

'        If Year > 2023 Then
        'PAG_C1 AY 2024-25 Change
        'Ayush_25-26
'        If Year > 2024 Then
'Ayush_26-27
'            If Year > 2025 Then
            If Year > CInt(Sheet8.Range("DOB_Year").value) Then
            ChkMaxDate = False
            Exit Function
        Else
'            If Year = 2024 Then
'                If Year = 2025 Then
'Ayush_26-27
                    If Year = CInt(Sheet8.Range("DOB_Year").value) Then
                If month > 4 Then
                    ChkMaxDate = False
                    Exit Function
                Else
                    If month = 4 Then
                        If dat > 1 Then
                            ChkMaxDate = False
                            Exit Function
                        Else
                            If dat = 1 Then
                                ChkMaxDate = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function

Public Function ChkMaxDatePrev(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

     ChkMaxDatePrev = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

        'If Year > 2022 Then
        
        'PAG_C1 AY 2024-25 Change
        If Year > 2023 Then
            ChkMaxDatePrev = False
            Exit Function
        Else
            If Year = 2023 Then
                If month > 4 Then
                    ChkMaxDatePrev = False
                    Exit Function
                Else
                    If month = 4 Then
                        If dat > 1 Then
                            ChkMaxDatePrev = False
                            Exit Function
                        Else
                            If dat = 1 Then
                                ChkMaxDatePrev = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function



Public Function CheckDate(dateEntered As Variant, Optional HUF As Variant = 0) As Long
'0: Date Ok
'1: Invalid Date Format
'2: Date not in Range

On Error Resume Next
    CheckDate = 0
    If Trim(dateEntered) = "" Or Not IsEmpty(dateEntered) Then
        If Not FormatNCheckDate(dateEntered, HUF) Then
            CheckDate = 1
            Exit Function
        End If

'        If Not ChkMaxDate(dateEntered, "02/04/2023") Then

       'PAG_C1 AY 2024-25 Change
       'Ayush_25-26
'        If Not ChkMaxDate(dateEntered, "02/04/2024") Then
'Ayush_26-27
 '       If Not ChkMaxDate(dateEntered, "02/04/2025") Then
        If Not ChkMaxDate(dateEntered, Sheet8.Range("DOB_1").value) Then
        
            CheckDate = 2
            Exit Function
        Else

        End If
    End If
 End Function

Sub ID_Country()
    If UCase(Sheet1.Range("sheet1.CountryCode")) = "91-INDIA" Then
    
          'Malli 31/03/2026
            If Sheet1.Range("sheet1.StateCode1").value = "" Then
            Sheet1.Range("sheet1.StateCode1") = "(Select)"
            End If
            If Sheet1.Range("sheet1.PinCode").value = "" Then
            Sheet1.Range("sheet1.PinCode") = ""
            End If
            '-----------------------
            
            If UCase(Sheet1.Range("sheet1.StateCode1")) = "99-FOREIGN" Then
                MsgBox "* Country Should Not be India For Selected State in Sheet Income Details" & Chr(13)
                Application.EnableEvents = False
                Sheet1.Range("sheet1.CountryCode").value = "(Select)"
                Sheet1.Range("sheet1.CountryCode").Select
                Application.EnableEvents = True
            End If
            
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.mobileCountryCode").MergeArea.value = "91"
            
            'Malli 31/03/2026
            If Not Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False Then
            Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
            Sheet1.Range("sheet1.PinCode").MergeArea.ClearContents
            Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)
            End If
            
            If Not Sheet1.Range("HASZIP").MergeArea.Locked = True Then
            Sheet1.Range("HASZIP").MergeArea.ClearContents
            Sheet1.Range("HASZIP").MergeArea.Locked = True
            Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)
            End If
            
            If Not Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True Then
            Sheet1.Range("sheet1.ZipCode").MergeArea.ClearContents
            Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
            Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
            End If
            '----------------------
            Sheet1.Protect Password:=getmsgstate
            
        ElseIf Sheet1.Range("sheet1.CountryCode").value = "(Select)" Or Sheet1.Range("sheet1.CountryCode").value = "" Then
            Application.EnableEvents = False
            Sheet1.Range("sheet1.StateCode1") = "(Select)"
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.mobileCountryCode").MergeArea.value = ""
            Sheet1.Range("sheet1.mobileCountryCode2").MergeArea.value = ""
            
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
            Application.EnableEvents = True

        ElseIf UCase(Sheet1.Range("sheet1.CountryCode").value) <> "91-INDIA" Then
        
            If UCase(Sheet1.Range("sheet1.StateCode1")) <> "99-FOREIGN" And UCase(Sheet1.Range("sheet1.StateCode1")) <> "(SELECT)" And UCase(Sheet1.Range("sheet1.StateCode1")) <> "" Then
                MsgBox "* Country cannot be other than India as you have selected an Indian state in Sheet Income Details" & Chr(13)
                Sheet1.Range("sheet1.CountryCode").value = "91-INDIA"
                Return
            End If
            
            Sheet1.Range("sheet1.StateCode1").value = "99-Foreign"
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.mobileCountryCode").MergeArea.value = ""
            Sheet1.Range("sheet1.mobileCountryCode2").MergeArea.value = ""

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
'Ayush_13/01/2026_V0.2
Sub ID_Country2()
    If UCase(Sheet1.Range("sheet1.CountryCode2")) = "91-INDIA" Then
            Sheet1.Range("sheet1.StateCode2") = "(Select)"
            Sheet1.Range("sheet1.PinCode2") = ""
            
            If UCase(Sheet1.Range("sheet1.StateCode2")) = "99-FOREIGN" Then
                MsgBox "* Country Should Not be India For Selected State in Sheet Income Details" & Chr(13)
                Application.EnableEvents = False
                Sheet1.Range("sheet1.CountryCode2").value = "(Select)"
                Sheet1.Range("sheet1.CountryCode2").Select
                Application.EnableEvents = True
            End If
            
            Sheet1.Unprotect Password:=getmsgstate
           ' Sheet1.Range("sheet1.mobileCountryCode2").MergeArea.value = "91"
            Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = False
            Sheet1.Range("sheet1.PinCode2").MergeArea.ClearContents
            Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = (&HCCFFCC)
            
            Sheet1.Range("HASZIP2").MergeArea.ClearContents
            Sheet1.Range("HASZIP2").MergeArea.Locked = True
            Sheet1.Range("HASZIP2").MergeArea.Interior.Color = (&HD8D8D8)
            
            Sheet1.Range("sheet1.ZipCode2").MergeArea.ClearContents
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = (&HD8D8D8)
            
            Sheet1.Protect Password:=getmsgstate
            
        ElseIf Sheet1.Range("sheet1.CountryCode2").value = "(Select)" Or Sheet1.Range("sheet1.CountryCode2").value = "" Then
            Application.EnableEvents = False
            Sheet1.Range("sheet1.StateCode2") = "(Select)"
            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range("sheet1.mobileCountryCode").MergeArea.value = ""
'            Sheet1.Range("sheet1.mobileCountryCode2").MergeArea.value = ""
            
            Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = False
            Sheet1.Range("sheet1.PinCode2").MergeArea.value = ""
            Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = (&HCCFFCC)

            Sheet1.Range("HASZIP2").MergeArea.value = ""
            Sheet1.Range("HASZIP2").MergeArea.Locked = True
            Sheet1.Range("HASZIP2").MergeArea.Interior.Color = (&HD8D8D8)

            Sheet1.Range("sheet1.ZipCode2").MergeArea.value = ""
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = True
            Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)

            Sheet1.Protect Password:=getmsgstate
            Application.EnableEvents = True

        ElseIf UCase(Sheet1.Range("sheet1.CountryCode2").value) <> "91-INDIA" Then
        
            If UCase(Sheet1.Range("sheet1.StateCode2")) <> "99-FOREIGN" And UCase(Sheet1.Range("sheet1.StateCode2")) <> "(SELECT)" And UCase(Sheet1.Range("sheet1.StateCode1")) <> "" Then
                MsgBox "* Country cannot be other than India as you have selected an Indian state in Sheet Income Details" & Chr(13)
                Sheet1.Range("sheet1.CountryCode2").value = "91-INDIA"
                Return
            End If
            
            Sheet1.Range("sheet1.StateCode2").value = "99-Foreign"
            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range("sheet1.mobileCountryCode").MergeArea.value = ""
'            Sheet1.Range("sheet1.mobileCountryCode2").MergeArea.value = ""

            Sheet1.Range("sheet1.PinCode2").MergeArea.Locked = True
            Sheet1.Range("sheet1.PinCode2").MergeArea.value = ""
            Sheet1.Range("sheet1.PinCode2").MergeArea.Interior.Color = (&HD8D8D8)

            Sheet1.Range("HASZIP2").MergeArea.value = ""
            Sheet1.Range("HASZIP2").MergeArea.Locked = False
            Sheet1.Range("HASZIP2").MergeArea.Interior.Color = (&HCCFFCC)

            Sheet1.Range("sheet1.ZipCode2").MergeArea.value = ""
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Locked = False
            Sheet1.Range("sheet1.ZipCode2").MergeArea.Interior.Color = (&HCCFFCC)

            Sheet1.Protect Password:=getmsgstate
        End If
End Sub
'---------------------
Sub IT1_DateofDeposit(ByVal Target As Range)
                Dim temp
                temp = Dformat(Sheet202.Range(Target.Address).value, "")
'                If Not ChkMinInclusiveDate(temp, "2024-01-01") Then
'                    fmsgboxsmall ("The date of credit to Central government cannot be prior to 01/01/2024 in schedule IT.")
'                If Not ChkMinInclusiveDate(temp, "2025-01-01") Then
'                    fmsgboxsmall ("Please enter valid date in dd/mm/yyyy format in schedule IT. Also date of credit to Central government cannot be prior to 01/01/2025 in schedule IT.")
'                If Not FormatNCheckDate(temp) Then
'                    fmsgboxsmall ("""Date of credit into Govt Account must be a valid dd/mm/yyyy format in schedule IT-1""")
'                    Sheet202.Range(Target.Address).value = ""
'                 End If
                 
                 If Not ChkMinInclusiveDate(temp, "2026-01-01") Then
                    fmsgboxsmall (" ""Please enter valid date in dd/mm/yyyy format in schedule IT. Also date of credit to Central government cannot be prior to 01/01/2026 in schedule IT-1.""")
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
   msgbox_1398AATI = msgbox_1398AATI + "* ""Date of credit into Govt Account must be a valid dd/mm/yyyy format at Sr. No. " & i & " in schedule IT-1""" & Chr(13)
    Exit Function
End If

If Not chkCompulsory_1398A(DateCredit_IT1(i)) Then
    msgbox_1398AATI = msgbox_1398AATI + "* ""Please enter the Date of credit into Govt Account at Sr. No. " & i & " in schedule IT-1""" & Chr(13)
    ValidateDate_IT1 = False
    Exit Function
End If

'If Not ChkMinInclusiveDate_1398A(Dformat(DateCredit_IT1(i), ""), "2025-01-01") Then ''
'    msgbox_1398AATI = msgbox_1398AATI + "*Date of credit to Central government at Sr. No.  " & i & " cannot be prior to 01/01/2024 in schedule IT" & Chr(13)
 If Not ChkMinInclusiveDate_1398A(Dformat(DateCredit_IT1(i), ""), "2026-01-01") Then ''
    msgbox_1398AATI = msgbox_1398AATI + "* ""Date of credit to Central government at Sr. No.  " & i & " cannot be prior to 01/01/2026 in schedule IT-1""" & Chr(13)
    
    
    
    ValidateDate_IT1 = False
    Exit Function
End If
Next
End Function

Sub IT2_DateofDeposit(ByVal Target As Range)
    Dim temp1
    temp1 = Dformat(Sheet202.Range(Target.Address).value, "")
'    If Not ChkMinInclusiveDate(temp1, "2022-04-01") Then
'        fmsgboxsmall ("The date of credit to Central government cannot be prior to 01/04/2022  in schedule IT.")
'Ayush_UR

'  If Not ChkMinInclusiveDate(temp1, "2023-04-01") Then
'        fmsgboxsmall ("Please enter valid date in dd/mm/yyyy format in schedule IT. Also date of credit to Central government cannot be prior to 01/04/2023 in schedule IT.")

   If Not ChkMinInclusiveDate(temp1, "2024-04-01") Then
         fmsgboxsmall (" ""Please enter valid date in dd/mm/yyyy format in schedule IT. Also date of credit to Central government cannot be prior to 01/04/2024 in schedule IT-2.""")

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
If Not FormatNCheckDate(DateCredit_IT2(i)) Then
    ValidateDate_IT2 = False
   msgbox_1398AATI = msgbox_1398AATI + "* ""Date of credit into Govt Account must be a valid dd/mm/yyyy format at Sr. No. " & i & " in Schedule IT-2""" & Chr(13)
    Exit Function
End If

If Not chkCompulsory_1398A(DateCredit_IT2(i)) Then
    msgbox_1398AATI = msgbox_1398AATI + "* ""Please enter the Date of credit into Govt Account at Sr. No. " & i & " in Schedule IT-2""" & Chr(13)
    ValidateDate_IT2 = False
    Exit Function
End If

'If Not ChkMinInclusiveDate_1398A(Dformat(DateCredit_IT2(i), ""), "2022-04-01") Then ''
'    msgbox_1398AATI = msgbox_1398AATI + "*Date of credit to Central government at Sr. No  " & i & " cannot be prior to 01/04/2022 in Schedule IT" & Chr(13)
 If Not ChkMinInclusiveDate_1398A(Dformat(DateCredit_IT2(i), ""), "2024-04-01") Then ''
    msgbox_1398AATI = msgbox_1398AATI + "* ""Date of credit to Central government at Sr. No  " & i & " cannot be prior to 01/04/2024 in Schedule IT-2""" & Chr(13)
    
    ValidateDate_IT2 = False
    Exit Function
End If
Next
End Function


Sub ReturnDropdownA23_24(RowIndex As Long)
On Error GoTo endline
Application.EnableEvents = False
Sheet201.Unprotect Password:=getmsgstate



'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).value = "2024-25" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
'        'Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'        If (Dformat("01/01/2025", "") <= Dformat(Sheet5.Range("Ver.Date").value, "")) And (Dformat("31/03/2025", "") >= Dformat(Sheet5.Range("Ver.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AY23Return"
'        ElseIf (Dformat("31/12/2024", "") >= Dformat(Sheet5.Range("Ver.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'        End If
'    End If
'End If
If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).value = "2025-26" Then
    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).value = "No" Then
        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
        'Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
    Else
    
        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
        If (Dformat("31/12/2025", "") >= Dformat(Sheet5.Range("Ver.Date").value, "")) Then
            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
        ElseIf (Dformat("01/01/2026", "") < Dformat(Sheet5.Range("Ver.Date").value, "")) Or (Dformat("31/03/2026", "") >= Dformat(Sheet5.Range("Ver.Date").value, "")) Then
            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AfterDateReturn"
        End If
    End If
End If

If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).value = "2026-27" Then
    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).value = "No" Then
        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
        'Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
    Else
    
        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
        If (Dformat("31/12/2026", "") >= Dformat(Sheet5.Range("Ver.Date").value, "")) Then
            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
        ElseIf (Dformat("01/01/2027", "") < Dformat(Sheet5.Range("Ver.Date").value, "")) Or (Dformat("31/03/2027", "") >= Dformat(Sheet5.Range("Ver.Date").value, "")) Then
            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AfterDateReturn"
        End If
    End If
End If
'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).value = "2025-26" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
'        'Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'        If (Dformat("01/01/2026", "") <= Dformat(Sheet5.Range("Ver.Date").value, "")) And (Dformat("31/03/2026", "") >= Dformat(Sheet5.Range("Ver.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AY23Return"
'        ElseIf (Dformat("31/12/2025", "") >= Dformat(Sheet5.Range("Ver.Date").value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'        End If
'    End If
'End If
endline:
Sheet201.Protect Password:=getmsgstate
    Application.EnableEvents = True
End Sub

Sub PartA_139_LockUnlockTable(ByVal Target As Range)
    Application.EnableEvents = False
    Sheet201.Unprotect Password:=getmsgstate
'        If Target.value = "2026-27" Or Target.value = "2025-26" Then
        If Target.value = "2027-28" Or Target.value = "2026-27" Then
            Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.Locked = False
            Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"

            Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
            Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
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

Sub CmdBP1()
    Sheets("BP").Activate
    mdCommon.DefinedgridNameRange = "Sec44AD.NOBName||Sec44AD.NOBCode||Sec44AD.NOBDescription"
    ActiveCellRange = mdCommon.searchLastRow("Sec44AD.NOBName")
    mdCommon.insertRowUnderSectionWithFormula
End Sub

Sub CmdBP2()
    Sheets("BP").Activate
    mdCommon.DefinedgridNameRange = "Sec44AD.NOBName1||Sec44AD.NOBCode1||Sec44AD.NOBDescription1"
    ActiveCellRange = mdCommon.searchLastRow("Sec44AD.NOBName1")
    mdCommon.insertRowUnderSectionWithFormula
End Sub

Sub CmdBP3()
    Sheets("BP").Activate
    mdCommon.DefinedgridNameRange = "Sec44AD.NOBName2||Sec44AD.NOBCode2||Sec44AD.NOBDescription2"
    ActiveCellRange = mdCommon.searchLastRow("Sec44AD.NOBName2")
    mdCommon.insertRowUnderSectionWithFormula
End Sub

Function ValidateTdsClaimed_TDS2() As Boolean
ValidateTdsClaimed_TDS2 = True
noOfProcessSub = end_TDS2_1

Dim rangecells As Range
Dim rangecells1 As Range
Dim rangecells2 As Range
Dim rangecells3 As Range
Dim i As Long

Set rangecells = Range("TDsOthr.tdsclaimed").Cells
Set rangecells1 = Range("TDsOthr.tds").Cells
Set rangecells2 = Range("TDsOthr.tdsdeducted").Cells
Set rangecells3 = Range("TDsOthr.grossamount").Cells

 ReDim TdsClaimed_TDS2_1(end_TDS2_1)
 ReDim Tds_TDS2_1(end_TDS2_1)
 ReDim Tdsdeducted_TDS2_1(end_TDS2_1)
 ReDim Tdsgrossamount_TDS2_1(end_TDS2_1)

 For i = 1 To end_TDS2_1
    TdsClaimed_TDS2_1(i) = rangecells.item(i).value
    Tds_TDS2_1(i) = rangecells1.item(i).value
    Tdsdeducted_TDS2_1(i) = rangecells2.item(i).value
    Tdsgrossamount_TDS2_1(i) = rangecells3.item(i).value

    If Not chkCompulsory(TdsClaimed_TDS2_1(i)) Then
      MsgTDS2 = MsgTDS2 + "* Please enter Tds claimed at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
       ValidateTdsClaimed_TDS2 = False
       Exit Function

    End If

    If Tds_TDS2_1(i) <> "" Then
    If TdsClaimed_TDS2_1(i) > Tds_TDS2_1(i) Then
      'MsgTDS2 = MsgTDS2 + "* Amount entered in Column 6 cannot be more than Column 4 or Column 5 at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
      'MsgTDS2 = MsgTDS2 + "* Amount entered in Column TDS Claimed cannot be more than Column TDS b/f or Column TDS Deducted at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
      MsgTDS2 = MsgTDS2 + "* Amount entered in Column TDS Claimed cannot be more than Column TDS b/f or Column TDS Deducted or Gross Amount at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
       ValidateTdsClaimed_TDS2 = False
       Exit Function
    End If
    End If
    If Tdsdeducted_TDS2_1(i) <> "" Then
    If TdsClaimed_TDS2_1(i) > Tdsdeducted_TDS2_1(i) Then
      'MsgTDS2 = MsgTDS2 + "* Amount entered in Column 6 cannot be more than Column 4 or Column 5 at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
      'MsgTDS2 = MsgTDS2 + "* Amount entered in Column TDS Claimed cannot be more than Column TDS b/f or Column TDS Deducted at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
      MsgTDS2 = MsgTDS2 + "* Amount entered in Column TDS Claimed cannot be more than Column TDS b/f or Column TDS Deducted or Gross Amount at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
       ValidateTdsClaimed_TDS2 = False
       Exit Function
    End If
    End If
    
    If Tdsgrossamount_TDS2_1(i) <> "" Then
    If TdsClaimed_TDS2_1(i) > Tdsgrossamount_TDS2_1(i) Then
       MsgTDS2 = MsgTDS2 + "* Amount entered in Column TDS Claimed cannot be more than Column TDS b/f or Column TDS Deducted or Gross Amount at Sr.No " & i & " in Schedule TDS2(i)" & Chr(13)
       ValidateTdsClaimed_TDS2 = False
       Exit Function
    End If
    End If

 UpdateProgressBar
Next
End Function


Function ValidateTdsClaimed_TDS2ii() As Boolean
ValidateTdsClaimed_TDS2ii = True
noOfProcessSub = end_TDS2ii_1

Dim rangecells As Range
Dim rangecells1 As Range
Dim rangecells2 As Range
Dim rangecells3 As Range
Dim i As Long

Set rangecells = Range("TDsOthr2.tdsclaimed").Cells
Set rangecells1 = Range("TDsOthr2.tds").Cells
Set rangecells2 = Range("TDsOthr2.tdsdeducted").Cells
Set rangecells3 = Range("TDsOthr2.grossamount").Cells

 ReDim TdsClaimed_TDS2ii_1(end_TDS2ii_1)
 ReDim Tds_TDS2ii_1(end_TDS2ii_1)
 ReDim Tdsdeducted_TDS2ii_1(end_TDS2ii_1)
 ReDim Tdsgrossamount_TDS2ii_1(end_TDS2ii_1)
 
 For i = 1 To end_TDS2ii_1
    TdsClaimed_TDS2ii_1(i) = rangecells.item(i).value
    Tds_TDS2ii_1(i) = rangecells1.item(i).value
    Tdsdeducted_TDS2ii_1(i) = rangecells2.item(i).value
    Tdsgrossamount_TDS2ii_1(i) = rangecells3.item(i).value

    If Not chkCompulsory(TdsClaimed_TDS2ii_1(i)) Then
    
    'Error Message Changed By Jyoti for the SIT-67660 on 25/06/2024
    
     ' MsgTDS2ii = MsgTDS2ii + "* Please enter Tds claimed at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
      
      MsgTDS2ii = MsgTDS2ii + "*""Please enter Tax Claimed"" at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
       ValidateTdsClaimed_TDS2ii = False
       Exit Function
    End If

    If Tds_TDS2ii_1(i) <> "" Then
    If TdsClaimed_TDS2ii_1(i) > Tds_TDS2ii_1(i) Then
      'MsgTDS2ii = MsgTDS2ii + "* Amount entered in Column 6 cannot be more than Column 4 or Column 5 at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
      'MsgTDS2ii = MsgTDS2ii + "* Amount entered in Column 6 cannot be more than Column TDS b/f or Column TDS Deducted. at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
      MsgTDS2ii = MsgTDS2ii + "* Amount entered in Column 6 cannot be more than Column TDS b/f or Column TDS Deducted or Gross Amount at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
       ValidateTdsClaimed_TDS2ii = False
       Exit Function
    End If
    End If
    If Tdsdeducted_TDS2ii_1(i) <> "" Then
    If TdsClaimed_TDS2ii_1(i) > Tdsdeducted_TDS2ii_1(i) Then
      'MsgTDS2ii = MsgTDS2ii + "* Amount entered in Column 6 cannot be more than Column 4 or Column 5 at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
      'MsgTDS2ii = MsgTDS2ii + "* Amount entered in Column 6 cannot be more than Column TDS b/f or Column TDS Deducted. at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
       MsgTDS2ii = MsgTDS2ii + "* Amount entered in Column 6 cannot be more than Column TDS b/f or Column TDS Deducted or Gross Amount at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
       ValidateTdsClaimed_TDS2ii = False
       Exit Function
    End If
    End If
    If Tdsgrossamount_TDS2ii_1(i) <> "" Then
    If TdsClaimed_TDS2ii_1(i) > Tdsgrossamount_TDS2ii_1(i) Then
      'MsgTDS2ii = MsgTDS2ii + "* Amount entered in Column 6 cannot be more than Column 4 or Column 5 at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
      'MsgTDS2ii = MsgTDS2ii + "* Amount entered in Column 6 cannot be more than Column TDS b/f or Column TDS Deducted. at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
       MsgTDS2ii = MsgTDS2ii + "* Amount entered in Column 6 cannot be more than Column TDS b/f or Column TDS Deducted or Gross Amount at Sr.No " & i & " in Schedule TDS2(ii)" & Chr(13)
       ValidateTdsClaimed_TDS2ii = False
       Exit Function
    End If
    End If


 UpdateProgressBar
Next
End Function


Function CountDropdown23_24(ByVal sDropdownValue As String, rangecells As Range, rangeCount As Long, ByVal Target As Range) As Long
    Dim currCnt  As Long
    Dim rangeValue As Variant
    Dim i As Long

    currCnt1 = 0
     currCnt2 = 0
      currCnt3 = 0
      currCnt4 = 0
      currCnt5 = 0
      currCnt6 = 0
      currCnt7 = 0
      currCnt8 = 0
      currCnt9 = 0
     currCnt10 = 0
      currCnt11 = 0
      currCnt12 = 0
      currCnt13 = 0
      currCnt14 = 0
      currCnt15 = 0
     currCnt16 = 0
     'Ayush_Legacy4
     currCnt17 = 0
     
     
    ReDim rangeValue(rangeCount)
      For i = 1 To rangeCount
        If Not rangecells.item(i).value = "(Select)" Then
            rangeValue(i) = rangecells.item(i).value
        End If
      Next

    If Not IsEmpty(rangeValue) And UBound(rangeValue) > 0 Then
      For i = 1 To UBound(rangeValue)
       If Trim(rangeValue(i)) = Trim(sDropdownValue) And (IIf(Trim(rangeValue(i)) = "", "(Select)", Trim(rangeValue(i))) <> "(Select)") Then

       'If rangeValue(i) = sDropdownValue And (IIf(rangeValue(i) = "", "(Select)", rangeValue(i)) <> "(Select)") Then
       If Mid(Trim(sDropdownValue), 1, 4) = "Agri" Then
       currCnt1 = currCnt1 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(34)" Then
       currCnt2 = currCnt2 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(19)" Then
       currCnt3 = currCnt3 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 12) = "Sec 10(10BC)" Then
       currCnt4 = currCnt4 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 11) = "Sec 10(10D)" Then
       currCnt5 = currCnt5 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(11)" Then
       currCnt6 = currCnt6 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(26)" Then
       currCnt7 = currCnt7 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 13) = "Sec 10(26AAA)" Then
       currCnt8 = currCnt8 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(12)" Then
       currCnt9 = currCnt9 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(13)" Then
       currCnt10 = currCnt10 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(16)" Then
       currCnt11 = currCnt11 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(17)" Then
       currCnt12 = currCnt12 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 11) = "Sec 10(17A)" Then
       currCnt13 = currCnt13 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(18)" Then
       currCnt14 = currCnt14 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 4) = "Defe" Then
       currCnt15 = currCnt15 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 4) = "LTCG" Then
       currCnt16 = currCnt16 + 1
       End If
       
        'Ayush_Legacy4
       If Mid(Trim(sDropdownValue), 1, 11) = "Sec 10(12C)" Then
       currCnt17 = currCnt17 + 1
       End If
       
       End If
      Next
    End If
    CountDropdown = currCnt1
    
    
    'If currCnt1 > 1 Or currCnt2 > 1 Or currCnt3 > 1 Or currCnt4 > 1 Or currCnt5 > 1 Or currCnt6 > 1 Or currCnt7 > 1 Or currCnt8 > 1 Or currCnt9 > 1 Or currCnt10 > 1 Or currCnt11 > 1 Or currCnt12 > 1 Or currCnt13 > 1 Or currCnt14 > 1 Or currCnt15 > 1 Or currCnt16 > 1 Then
'    If currCnt1 > 1 Or currCnt2 > 1 Or currCnt3 > 1 Or currCnt4 > 1 Or currCnt5 > 1 Or currCnt6 > 1 Or currCnt7 > 1 Or currCnt8 > 1 Or currCnt9 > 1 Or currCnt10 > 1 Or currCnt11 > 1 Or currCnt12 > 1 Or currCnt13 > 1 Or currCnt14 > 1 Or currCnt15 > 1 Then
'Ayush_Legacy4
 If currCnt1 > 1 Or currCnt2 > 1 Or currCnt3 > 1 Or currCnt4 > 1 Or currCnt5 > 1 Or currCnt6 > 1 Or currCnt7 > 1 Or currCnt8 > 1 Or currCnt9 > 1 Or currCnt10 > 1 Or currCnt11 > 1 Or currCnt12 > 1 Or currCnt13 > 1 Or currCnt14 > 1 Or currCnt15 > 1 Or currCnt17 > 1 Then
        fmsgbox "* Do not enter same Nature of income twice" ',' vbOKOnly, "Error(s)!"
        Target.value = "(Select)"
    End If
End Function

Sub TAX_LTCG_AMOUNT(ByVal Target As Range)
    'If Mid(Sheet5.Range(Replace(Target.Address, "I", "G")).value, 1, 10) = "Any exempt" Then
    If Mid(Sheet5.Range(Replace(Target.Address, "I", "G")).value, 1, 4) = "LTCG" Then
        If Sheet5.Range(Target.Address).value > 100000 Then
            'fmsgbox "* Any exempt income including LTCG on which tax is not payablecannot be more than Rs 100000"
             
            fmsgbox "* LTCG u/s 112A not exceeding  Rs. 1 Lakh' should not be more than Rs 100000"
            Sheet5.Range(Target.Address).value = ""
            Sheet5.Range(Target.Address).Select
        End If
    End If
End Sub


Sub Tax_LockUnlockTRP()

    If Sheet1.Range("Sheet1.ReturnFileSec").value = "139(8A)" Then
        Sheet5.Unprotect Password:=getmsgstate
        Sheet5.Range("Ver.IdentificationNoOfTRP").MergeArea.Locked = True
        Sheet5.Range("Ver.IdentificationNoOfTRP").MergeArea.Interior.Color = "&HD8D8D8"   '&HCCFFCC
        Sheet5.Range("Ver.NameOfTRP").MergeArea.Locked = True
        Sheet5.Range("Ver.NameOfTRP").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet5.Range("Ver.ReImbFrmGov").MergeArea.Locked = True
        Sheet5.Range("Ver.ReImbFrmGov").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet5.Protect Password:=getmsgstat
    Else
        Sheet5.Unprotect Password:=getmsgstate
        Sheet5.Range("Ver.IdentificationNoOfTRP").MergeArea.Locked = False
        Sheet5.Range("Ver.IdentificationNoOfTRP").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet5.Range("Ver.NameOfTRP").MergeArea.Locked = False
        Sheet5.Range("Ver.NameOfTRP").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet5.Range("Ver.ReImbFrmGov").MergeArea.Locked = False
        Sheet5.Range("Ver.ReImbFrmGov").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet5.Protect Password:=getmsgstat
    End If

End Sub

Sub IC_LockUnlockGrossIncome(ByVal Target As Range)
    If Target.value = "Not Applicable (eg. Family pension etc)" Then
        Sheet1.Unprotect Password:=getmsgstate
        'Sheet1.Range("IncD.IncomeFromBP").MergeArea.Locked = True
        'Sheet1.Range("IncD.IncomeFromBP").MergeArea.Interior.Color = "&HD8D8D8"   '&HCCFFCC
        Sheet1.Range("IncD.IncomeFromSal_i").MergeArea.Locked = True
        Sheet1.Range("IncD.IncomeFromSal_i").MergeArea.Interior.Color = "&HD8D8D8"
        
        Sheet1.Range("IncD.IncomeFromSal_ii").MergeArea.ClearContents
        Sheet1.Range("IncD.IncomeFromSal_ii").MergeArea.Locked = True
        Sheet1.Range("IncD.IncomeFromSal_ii").MergeArea.Interior.Color = "&HD8D8D8"
        
        Sheet1.Range("IncD.IncomeFromSal_iii").MergeArea.ClearContents
        Sheet1.Range("IncD.IncomeFromSal_iii").MergeArea.Locked = True
        Sheet1.Range("IncD.IncomeFromSal_iii").MergeArea.Interior.Color = "&HD8D8D8"
        
        Sheet1.Range("IncD.IncomeFromSal_iv").MergeArea.ClearContents
        Sheet1.Range("IncD.IncomeFromSal_iv").MergeArea.Locked = True
        Sheet1.Range("IncD.IncomeFromSal_iv").MergeArea.Interior.Color = "&HD8D8D8"
        
        'Sheet1.Range("IncomeNotified89A").MergeArea.Locked = True
        'Sheet1.Range("IncomeNotified89A").MergeArea.Interior.Color = "&HD8D8D8"
        
        'Ayush_13/01/2026_V0.2_Commented
'        Sheet1.Range("IncomeNotified89A_AmountUS").MergeArea.ClearContents
'        Sheet1.Range("IncomeNotified89A_AmountUS").MergeArea.Locked = True
'        Sheet1.Range("IncomeNotified89A_AmountUS").MergeArea.Interior.Color = "&HD8D8D8"
'
'        Sheet1.Range("IncomeNotified89A_AmountUK").MergeArea.ClearContents
'        Sheet1.Range("IncomeNotified89A_AmountUK").MergeArea.Locked = True
'        Sheet1.Range("IncomeNotified89A_AmountUK").MergeArea.Interior.Color = "&HD8D8D8"
'
'        Sheet1.Range("IncomeNotified89A_AmountCan").MergeArea.ClearContents
'        Sheet1.Range("IncomeNotified89A_AmountCan").MergeArea.Locked = True
'        Sheet1.Range("IncomeNotified89A_AmountCan").MergeArea.Interior.Color = "&HD8D8D8"
        
'        Sheet1.Range("IncomeNotifiedOther89A").MergeArea.ClearContents
'        Sheet1.Range("IncomeNotifiedOther89A").MergeArea.Locked = True
'        Sheet1.Range("IncomeNotifiedOther89A").MergeArea.Interior.Color = "&HD8D8D8"
        
        'Sheet1.Range("IncD.IncomeFromSal_allow").MergeArea.Locked = True
        'Sheet1.Range("IncD.IncomeFromSal_allow").MergeArea.Interior.Color = "&HD8D8D8"
        
        Sheet1.Range("Sheet1.Nature").value = "(Select)"
        Sheet1.Range("Sheet1.Nature").Locked = True
        Sheet1.Range("Sheet1.Nature").Interior.Color = "&HD8D8D8"
        
        Sheet1.Range("Sheet1.Description").value = "Not Applicable"
        Sheet1.Range("Sheet1.Description").Locked = True
        Sheet1.Range("Sheet1.Description").Interior.Color = "&HD8D8D8"
        
        Sheet1.Range("Sheet1.Amount").ClearContents
        Sheet1.Range("Sheet1.Amount").Locked = True
        Sheet1.Range("Sheet1.Amount").Interior.Color = "&HD8D8D8"
 'Ayush_15/01/2026_V0.2_commented
'        Sheet1.Range("Increliefus89A").MergeArea.ClearContents
'        Sheet1.Range("Increliefus89A").MergeArea.Locked = True
'        Sheet1.Range("Increliefus89A").MergeArea.Interior.Color = "&HD8D8D8"
        
        'Sheet1.Range("IncD.IncomeFromSal_net").MergeArea.Locked = True
        'Sheet1.Range("IncD.IncomeFromSal_net").MergeArea.Interior.Color = "&HD8D8D8"
        'Sheet1.Range("IncD.IncomeFromSal_va").MergeArea.Locked = True
        'Sheet1.Range("IncD.IncomeFromSal_va").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("IncD.IncomeFromSal_iva").MergeArea.Locked = True
        Sheet1.Range("IncD.IncomeFromSal_iva").MergeArea.Interior.Color = "&HD8D8D8"
        

        
        Sheet1.Range("IncD.IncomeFromSal_ivc").MergeArea.ClearContents
        Sheet1.Range("IncD.IncomeFromSal_ivc").MergeArea.Locked = True
        Sheet1.Range("IncD.IncomeFromSal_ivc").MergeArea.Interior.Color = "&HD8D8D8"
        
        'Sheet1.Range("IncD.IncomeFromSal").MergeArea.Locked = True
        'Sheet1.Range("IncD.IncomeFromSal").MergeArea.Interior.Color = "&HD8D8D8"
        
        Sheet1.Protect Password:=getmsgstat
    Else
        Sheet1.Unprotect Password:=getmsgstate
        
        'Sheet1.Range("IncD.IncomeFromBP").MergeArea.Locked = True
        'Sheet1.Range("IncD.IncomeFromBP").MergeArea.Interior.Color = "&HFFFFFF"   '&HCCFFCC
        Sheet1.Range("IncD.IncomeFromSal_i").MergeArea.Locked = True
        Sheet1.Range("IncD.IncomeFromSal_i").MergeArea.Interior.Color = "&HFFFFFF"
        
        Sheet1.Range("IncD.IncomeFromSal_ii").MergeArea.Locked = False
        Sheet1.Range("IncD.IncomeFromSal_ii").MergeArea.Interior.Color = "&HCCFFCC"
        
        Sheet1.Range("IncD.IncomeFromSal_iii").MergeArea.Locked = False
        Sheet1.Range("IncD.IncomeFromSal_iii").MergeArea.Interior.Color = "&HCCFFCC"
        
        Sheet1.Range("IncD.IncomeFromSal_iv").MergeArea.Locked = False
        Sheet1.Range("IncD.IncomeFromSal_iv").MergeArea.Interior.Color = "&HCCFFCC"
        
        'Sheet1.Range("IncomeNotified89A").MergeArea.Locked = True
        'Sheet1.Range("IncomeNotified89A").MergeArea.Interior.Color = "&HFFFFFF"
        
        'Ayush_13/01/2026_V0.2_Commented
'        Sheet1.Range("IncomeNotified89A_AmountUS").MergeArea.Locked = False
'        Sheet1.Range("IncomeNotified89A_AmountUS").MergeArea.Interior.Color = "&HCCFFCC"
'
'        Sheet1.Range("IncomeNotified89A_AmountUK").MergeArea.Locked = False
'        Sheet1.Range("IncomeNotified89A_AmountUK").MergeArea.Interior.Color = "&HCCFFCC"
'
'        Sheet1.Range("IncomeNotified89A_AmountCan").MergeArea.Locked = False
'        Sheet1.Range("IncomeNotified89A_AmountCan").MergeArea.Interior.Color = "&HCCFFCC"
        
'        Sheet1.Range("IncomeNotifiedOther89A").MergeArea.Locked = False
'        Sheet1.Range("IncomeNotifiedOther89A").MergeArea.Interior.Color = "&HCCFFCC"
        
        'Sheet1.Range("IncD.IncomeFromSal_allow").MergeArea.Locked = True
        'Sheet1.Range("IncD.IncomeFromSal_allow").MergeArea.Interior.Color = "&HFFFFFF"
        
        Sheet1.Range("Sheet1.Nature").Locked = False
        Sheet1.Range("Sheet1.Nature").Interior.Color = "&HCCFFCC"
        
        'Sheet1.Range("Sheet1.Description").Locked = False
        'Sheet1.Range("Sheet1.Description").Interior.Color = "&HD8D8D8"
        
        Sheet1.Range("Sheet1.Amount").Locked = False
        Sheet1.Range("Sheet1.Amount").Interior.Color = "&HCCFFCC"
        
 'Ayush_15/01/2026_V0.2_commented
'        Sheet1.Range("Increliefus89A").MergeArea.Locked = False
'        Sheet1.Range("Increliefus89A").MergeArea.Interior.Color = "&HCCFFCC"
        
        'Sheet1.Range("IncD.IncomeFromSal_net").MergeArea.Locked = True
        'Sheet1.Range("IncD.IncomeFromSal_net").MergeArea.Interior.Color = "&HFFFFFF"
        'Sheet1.Range("IncD.IncomeFromSal_va").MergeArea.Locked = True
        'Sheet1.Range("IncD.IncomeFromSal_va").MergeArea.Interior.Color = "&HFFFFFF"
        Sheet1.Range("IncD.IncomeFromSal_iva").MergeArea.Locked = True
        Sheet1.Range("IncD.IncomeFromSal_iva").MergeArea.Interior.Color = "&HFFFFFF"
        

        
        Sheet1.Range("IncD.IncomeFromSal_ivc").MergeArea.Locked = False
        Sheet1.Range("IncD.IncomeFromSal_ivc").MergeArea.Interior.Color = "&HCCFFCC"
        
        'Sheet1.Range("IncD.IncomeFromSal").MergeArea.Locked = True
        'Sheet1.Range("IncD.IncomeFromSal").MergeArea.Interior.Color = "&HFFFFFF"
        
        Sheet1.Protect Password:=getmsgstat
    End If
End Sub


Sub reset10IEAy23_24()
On Error Resume Next
Sheet8.Unprotect Password:=getmsgstate
Sheet1.Unprotect Password:=getmsgstate
'--Below fields are all hidden

    If Sheet8.Range("OldBacValue").value = 2 Or Sheet8.Range("OldBacValue").value = "" Or Sheet1.Range("sheet1.Status") = "F - FIRM(Other than LLP)" Then
    
        Sheet1.Range("AYinOpting").MergeArea.Locked = True
        Sheet1.Range("AYinOpting").Interior.Color = (&HD8D8D8)
        Sheet1.Range("AYinOpting").value = ""
        
        Sheet1.Range("Sheet1.date10IE_Opt").MergeArea.Locked = True
        Sheet1.Range("Sheet1.date10IE_Opt").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.date10IE_Opt").value = ""
    
        Sheet1.Range("Sheet1.acknowledgementNumber_Opt").MergeArea.Locked = True
        Sheet1.Range("Sheet1.acknowledgementNumber_Opt").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.acknowledgementNumber_Opt").value = ""
    
        Sheet1.Range("Opt_Out115BAC").MergeArea.Locked = True
        Sheet1.Range("Opt_Out115BAC").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Opt_Out115BAC").value = ""
        
        Sheet1.Range("AYinOptOut").Interior.Color = (&HD8D8D8)
        Sheet1.Range("AYinOptOut").MergeArea.Locked = True
        Sheet1.Range("AYinOptOut").value = ""
        
        Sheet1.Range("Sheet1.date10IE_OptOut").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.date10IE_OptOut").MergeArea.Locked = True
        Sheet1.Range("Sheet1.date10IE_OptOut").value = ""
        
        Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").MergeArea.Locked = True
        Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").value = ""
    
    Else
        If Sheet1.Range("AYinOpting").Locked = True Then
            Sheet1.Range("AYinOpting").value = "(Select)"
            Sheet1.Range("AYinOpting").MergeArea.Locked = False
            Sheet1.Range("AYinOpting").Interior.Color = (&HCCFFCC)
        End If
        
        If Sheet1.Range("Sheet1.date10IE_Opt").Locked = True Then
            Sheet1.Range("Sheet1.date10IE_Opt").value = ""
            Sheet1.Range("Sheet1.date10IE_Opt").MergeArea.Locked = False
            Sheet1.Range("Sheet1.date10IE_Opt").Interior.Color = (&HCCFFCC)
        End If
        
        If Sheet1.Range("Sheet1.acknowledgementNumber_Opt").Locked = True Then
            Sheet1.Range("Sheet1.acknowledgementNumber_Opt").value = ""
            Sheet1.Range("Sheet1.acknowledgementNumber_Opt").MergeArea.Locked = False
            Sheet1.Range("Sheet1.acknowledgementNumber_Opt").Interior.Color = (&HCCFFCC)
        End If
        
        If Sheet1.Range("Opt_Out115BAC").Locked = True Then
            Sheet1.Range("Opt_Out115BAC").value = "(Select)"
            Sheet1.Range("Opt_Out115BAC").MergeArea.Locked = False
            Sheet1.Range("Opt_Out115BAC").Interior.Color = (&HCCFFCC)
        End If
     End If

    Sheet8.Protect Password:=getmsgstate
    Sheet1.Protect Password:=getmsgstate
End Sub

'Sub CurrentYear10IE_23_24(ByVal Target As Range)
'AY 2023-24 Old Starts--------------------------------


'If Not Application.Intersect(Sheet1.Range("Opt_Out115BAC"), Sheet1.Range(Target.Address)) Is Nothing Then
'
'Sheet1.Unprotect Password:=getmsgstate
'Application.EnableEvents = False
'
'    If Sheet1.Range("Opt_Out115BAC").value = "Yes" Then
'
'        Sheet1.Range("AYinOptOut").value = "(Select)"
'        Sheet1.Range("AYinOptOut").MergeArea.Locked = False
'        Sheet1.Range("AYinOptOut").Interior.Color = (&HCCFFCC)
'
'        Sheet1.Range("Sheet1.date10IE_OptOut").value = ""
'        Sheet1.Range("Sheet1.date10IE_OptOut").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.date10IE_OptOut").Interior.Color = (&HCCFFCC)
'
'        Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").value = ""
'        Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").Interior.Color = (&HCCFFCC)
'
'    Else
'
'        Sheet1.Range("AYinOptOut").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("AYinOptOut").MergeArea.Locked = True
'        Sheet1.Range("AYinOptOut").value = ""
'
'        Sheet1.Range("Sheet1.date10IE_OptOut").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.date10IE_OptOut").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.date10IE_OptOut").value = ""
'
'        Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").value = ""
'
'    End If
'
'CurrentYearOptionAY23_24
'Sheet1.Protect Password:=getmsgstate
'Application.EnableEvents = True
'
'  End If
'End Sub

'AY 2023-24 Old Ends--------------------------------


'AY 2023-24 OLD STARTS---------------------------------------------------------------
'Sub CurrentYearOptionAY23_24()

'Sheet1.Unprotect Password:=getmsgstate
'Application.EnableEvents = False
'
'If Mid(Trim(Sheet1.Range("sheet1.Status").value), 1, 1) <> "F" Then
'    If Sheet8.Range("OldBacValue").value = 2 Then
'
'        Sheet1.Range("OptingNewTaxRegime").Validation.Delete
'        Sheet1.Range("OptingNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current1"
'        Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
'        Sheet1.Range("OptingNewTaxRegime").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("OptingNewTaxRegime").MergeArea.Locked = False
'
'    End If
'
'    If Sheet8.Range("OldBacValue").value = 1 And Sheet1.Range("Opt_Out115BAC").value = "Yes" Then
'        Sheet1.Range("OptingNewTaxRegime").Validation.Delete
'        Sheet1.Range("OptingNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current3"
'        Application.EnableEvents = True
'        Sheet1.Range("OptingNewTaxRegime").value = "Not eligible to opt in"
'        Application.EnableEvents = False
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("OptingNewTaxRegime").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("OptingNewTaxRegime").MergeArea.Locked = True
'
'    Else
'        Sheet1.Range("OptingNewTaxRegime").Validation.Delete
'        Sheet1.Range("OptingNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current1"
'        Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
'        Sheet1.Range("OptingNewTaxRegime").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("OptingNewTaxRegime").MergeArea.Locked = False
'
'        Sheet1.Range("Sheet1.date10IE").value = ""
'        Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.date10IE").Interior.Color = (&HCCFFCC)
'
'        Sheet1.Range("Sheet1.acknowledgementNumber").value = ""
'        Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.acknowledgementNumber").Interior.Color = (&HCCFFCC)
'    End If
'
'    If Sheet8.Range("OldBacValue").value = 1 And Sheet1.Range("Opt_Out115BAC").value = "No" Then
'        Sheet1.Range("OptingNewTaxRegime").Validation.Delete
'        Sheet1.Range("OptingNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current2"
'        Sheet1.Range("OptingNewTaxRegime").value = "(Select)"
'        Sheet1.Range("OptingNewTaxRegime").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("OptingNewTaxRegime").MergeArea.Locked = False
'    End If
'End If
'Sheet1.Protect Password:=getmsgstate
'Application.EnableEvents = True
'End Sub
'AY 2023-24 OLD End---------------------------------------------------------------

'Sub CurrentYear10IEDateAckAY_23_24(ByVal Target As Range)
'PAG_E18 AY 2023-24 Old Start-----------------------------------------------------
'If Not Application.Intersect(Sheet1.Range("OptingNewTaxRegime"), Sheet1.Range(Target.Address)) Is Nothing Then
'
'Sheet1.Unprotect Password:=getmsgstate
'Application.EnableEvents = False
'
'If (Sheet1.Range("OptingNewTaxRegime").value = "Opting in now") Or (Sheet1.Range("OptingNewTaxRegime").value = "Opt out") Then
'        Sheet1.Range("Sheet1.date10IE").value = ""
'        Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.date10IE").Interior.Color = (&HCCFFCC)
'
'        Sheet1.Range("Sheet1.acknowledgementNumber").value = ""
'        Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.acknowledgementNumber").Interior.Color = (&HCCFFCC)
'
'Else
'        Sheet1.Range("Sheet1.date10IE").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.date10IE").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.date10IE").value = ""
'
'        Sheet1.Range("Sheet1.acknowledgementNumber").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.acknowledgementNumber").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.acknowledgementNumber").value = ""
'
'
'End If
'
'Sheet1.Protect Password:=getmsgstate
'Application.EnableEvents = True
'
'End If
'End Sub ''PAG_E18 AY 2023-24 Old End------------------------------------------------------------------------




'Newly addedby sai for 115BAC(6)01-12-2023

'PAG_C22 AY 2024-25 Change Starts---------------------------------------------------------------------------

Sub CurrentYear10IEDateAckAY_24_25(ByVal Target As Range) 'Below range names are not used in DE sheet v0.2

If Not Application.Intersect(Sheet1.Range("sheet1.Status"), Sheet1.Range(Target.Address)) Is Nothing Then

Sheet1.Unprotect Password:=getmsgstate
Application.EnableEvents = False


    If (Left(Sheet1.Range("sheet1.Status").value, 1) = "F") Then
         
            Sheet1.Unprotect Password:=getmsgstate

            Sheet1.Range("Sheet1.OptOutNewTaxRegime").Validation.Delete
            Sheet1.Range("Sheet1.OptOutNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current5"
            Application.EnableEvents = True
            Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Not Applicable"
            Application.EnableEvents = False
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("Sheet1.OptOutNewTaxRegime").Interior.Color = (&HD8D8D8)
            Sheet1.Range("Sheet1.OptOutNewTaxRegime").MergeArea.Locked = True
            
           '10-12-2023-----------------------------------------------------------------------
           
            Sheet1.Range("Sheet1.Dateoffiling").Interior.Color = (&HD8D8D8)
            Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Locked = True
             Sheet1.Range("Sheet1.Dateoffiling").MergeArea.value = ""

            Sheet1.Range("Sheet1.acknowledgementNumber_New").Interior.Color = (&HD8D8D8)
            Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.Locked = True
            Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.value = ""
            
            
            
            ThisWorkbook.Unprotect Password:=getmsgstate
            Sheet11.Unprotect Password:=getmsgstate
            
            Sheet11.Visible = xlSheetVisible
            Sheet11.Protect Password:=getmsgstate
            
            Sheet6.Unprotect Password:=getmsgstate
            Sheet6.Visible = xlSheetVisible
            Sheet6.Protect Password:=getmsgstate
            
            
            Sheet16.Unprotect Password:=getmsgstate
            Sheet16.Visible = xlSheetHidden
            Sheet16.Protect Password:=getmsgstate
            
            Sheet14.Unprotect Password:=getmsgstate
            Sheet14.Visible = xlSheetHidden
            Sheet14.Protect Password:=getmsgstate
            
            
            
            '10-12-2023-----------------------------------------------------------------------
        ElseIf UCase(Sheet1.Range("sheet1.Status")) = UCase("(Select)") Or Sheet1.Range("sheet1.OptOutNewTaxRegime").value = "Not Applicable" Then
            Sheet1.Unprotect Password:=getmsgstate

            Sheet1.Range("Sheet1.OptOutNewTaxRegime").Validation.Delete
            Sheet1.Range("Sheet1.OptOutNewTaxRegime").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=NTR_Current4"
'Ayush_25-26
'            Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "No"
            Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "(Select)"
            Sheet1.Range("Sheet1.OptOutNewTaxRegime").Interior.Color = (&HCCFFCC)
            Sheet1.Range("Sheet1.OptOutNewTaxRegime").MergeArea.Locked = False

            NewBacYesValueChange
            'Ayush_25-26_commented

'            If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "No" Then
'
'
'            Sheet1.Unprotect Password:=getmsgstate
'
'            Sheet1.Range("Sheet1.Dateoffiling").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Locked = True
'                Sheet1.Range("Sheet1.Dateoffiling").MergeArea.ClearContents
'                Sheet1.Range("Sheet1.acknowledgementNumber_New").Interior.Color = (&HD8D8D8)
'                Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.Locked = True
'                Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.value = ""
'
'            '10-12-2023-----------------------------------------------------------------------
'           Else
'             Sheet1.Unprotect Password:=getmsgstate
'
'             Sheet1.Range("Sheet1.Dateoffiling").Interior.Color = (&HCCFFCC)
'            Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Locked = False
'
'            Sheet1.Range("Sheet1.acknowledgementNumber_New").Interior.Color = (&HCCFFCC)
'            Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.Locked = False
'
'            End If
            '10-12-2023-----------------------------------------------------------------------
            
    End If



Sheet1.Protect Password:=getmsgstate
'Application.EnableEvents = True

End If
End Sub
'PAG_C22 AY 2024-25 Change End 01-12-2023---------------------------------------------------------------------------

'PAG_E24 AY 2023-24 Old

'Sub ID_Opt_Date10IE(ByVal Target As Range)

'If Not Application.Intersect(Sheet1.Range("Sheet1.date10IE_Opt"), Sheet1.Range(Target.Address)) Is Nothing Then
'Sheet1.Unprotect Password:=getmsgstate
'        If Trim(Range("Sheet1.date10IE_Opt").value) <> "00/00/0000" And Range("Sheet1.date10IE_Opt").value <> "" Then
'
'            If Not FormatNCheckDate(Trim(Range("Sheet1.date10IE_Opt").value)) Then
'
'                fmsgboxStatus "* Date of filing of Form 10IE Is Invalid in Sheet Income Details. Please enter date in dd/mm/yyyy format."
'                Sheet1.Range("Sheet1.date10IE_Opt").Select
'                Sheet1.Range("Sheet1.date10IE_Opt").value = ""
'
'            Else
'
'                If (Sheet1.Range("AYinOpting").value = "2021-22") Then
'
''                        If Dformat(Trim(Range("Sheet1.date10IE_Opt").value), "yyyy-mm-dd") > Dformat(Trim("31/03/2022"), "yyyy-mm-dd") Then
''
''                            fmsgboxStatus "* Date of filing of Form 10IE cannot be beyond 31st March, 2022 for AY 2021-22"
''                            Sheet1.Range("Sheet1.date10IE_Opt").Select
''                            Sheet1.Range("Sheet1.date10IE_Opt").value = ""
'
'                        If Dformat(Trim(Range("Sheet1.date10IE_Opt").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2021"), "yyyy-mm-dd") Then
'
'                            fmsgboxStatus "* Date of filing of Form 10IE cannot be prior to  01/04/2021"
'                            Sheet1.Range("Sheet1.date10IE_Opt").Select
'                            Sheet1.Range("Sheet1.date10IE_Opt").value = ""
'
'                        ElseIf Dformat(Trim(Range("Sheet1.date10IE_Opt").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'
'                            fmsgbox "* Date of filing of Form 10IE can Not be more than system Date in Income Details"
'                            Range("Sheet1.date10IE_Opt").Select
'                            Range("Sheet1.date10IE_Opt").value = ""
'
'                        End If
'
'                ElseIf (Sheet1.Range("AYinOpting").value = "2022-23") Then
'
''                        If Dformat(Trim(Range("Sheet1.date10IE_Opt").value), "yyyy-mm-dd") > Dformat(Trim("31/03/2023"), "yyyy-mm-dd") Then
''
''                            fmsgboxStatus "* Date of filing of Form 10IE cannot be beyond 31st March, 2023 for AY 2022-23"
''                            Sheet1.Range("Sheet1.date10IE_Opt").Select
''                            Sheet1.Range("Sheet1.date10IE_Opt").value = ""
'
'                        If Dformat(Trim(Range("Sheet1.date10IE_Opt").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2021"), "yyyy-mm-dd") Then
'
'                            fmsgboxStatus "* Date of filing of Form 10IE cannot be prior to  01/04/2021"
'                            Sheet1.Range("Sheet1.date10IE_Opt").Select
'                            Sheet1.Range("Sheet1.date10IE_Opt").value = ""
'
'                        ElseIf Dformat(Trim(Range("Sheet1.date10IE_Opt").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'
'                            fmsgbox "* Date of filing of Form 10IE can Not be more than system Date in Income Details"
'                            Range("Sheet1.date10IE_Opt").Select
'                            Range("Sheet1.date10IE_Opt").value = ""
'
'                        End If
'
'                End If
'            End If
'        End If
'Sheet1.Protect Password:=getmsgstate
'End If
        'PAG_E24 AY 2023-24 Old-------------------------------------------
  

'PAG_C25 AY 2024-25 Change Starts--------------------------------------------------------------------------------------------

Sub ID_Opt_Date10IE_New(ByVal Target As Range)
          If Not Application.Intersect(Sheet1.Range("Sheet1.Dateoffiling"), Sheet1.Range(Target.Address)) Is Nothing Then
            Sheet1.Unprotect Password:=getmsgstate
          If Trim(Range("Sheet1.Dateoffiling").value) <> "00/00/0000" And Range("Sheet1.Dateoffiling").value <> "" Then
        
            If Not FormatNCheckDate(Trim(Range("Sheet1.Dateoffiling").value)) Then
            
                fmsgboxStatus "* Please enter date in dd/mm/yyyy format."
                
                Sheet1.Range("Sheet1.Dateoffiling").value = ""
                
            Else
'Ayush_25-26
'         If Dformat(Trim(Range("Sheet1.Dateoffiling").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2024"), "yyyy-mm-dd") Then
If Dformat(Trim(Range("Sheet1.Dateoffiling").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2025"), "yyyy-mm-dd") Then
                        
'                            fmsgboxStatus "*Date of filing of Form 10IEA cannot be prior to  01/04/2024"
                            fmsgboxStatus """Date of filing of Form 10IEA cannot be prior to  01/04/2025"""
                            Sheet1.Range("Sheet1.Dateoffiling").Select
                            Sheet1.Range("Sheet1.Dateoffiling").value = ""
        
                        ElseIf Dformat(Trim(Range("Sheet1.Dateoffiling").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
                        
                            fmsgbox "* Date of filing of Form 10IEA can Not be more than system Date in Income Details"
                           
                            Range("Sheet1.Dateoffiling").value = ""
                            
                    End If
                End If
            End If
        End If
        Sheet1.Protect Password:=getmsgstate


End Sub
'PAG_C25 AY 2024-25 Change End---------------------------------------------------------------------------------

'PAG 2023-24 OLD------------------------------------------------------------------
'Sub ID_OptOut_Date10IE(ByVal Target As Range)
'
'If Not Application.Intersect(Sheet1.Range("Sheet1.date10IE_OptOut"), Sheet1.Range(Target.Address)) Is Nothing Then
'Sheet1.Unprotect Password:=getmsgstate
'
'        If Trim(Range("Sheet1.date10IE_OptOut").value) <> "00/00/0000" And Range("Sheet1.date10IE_OptOut").value <> "" Then
'
'            If Not FormatNCheckDate(Trim(Range("Sheet1.date10IE_OptOut").value)) Then
'
'                fmsgboxStatus "* Date of filing of Form 10IE Is Invalid in Sheet Income Details. Please enter date in dd/mm/yyyy format."
'                Sheet1.Range("Sheet1.date10IE_OptOut").Select
'                Sheet1.Range("Sheet1.date10IE_OptOut").value = ""
'
'            Else
'
'                 If (Sheet1.Range("AYinOptOut").value = "2021-22") Then
'
''                    If Dformat(Trim(Range("Sheet1.date10IE_OptOut").value), "yyyy-mm-dd") > Dformat(Trim("31/03/2022"), "yyyy-mm-dd") Then
''
''                        fmsgboxStatus "* Date of filing of Form 10IE cannot be beyond 31st March, 2022 in AY 2021-22"
''                        Sheet1.Range("Sheet1.date10IE_OptOut").Select
''                        Sheet1.Range("Sheet1.date10IE_OptOut").value = ""
'
'                    If Dformat(Trim(Range("Sheet1.date10IE_OptOut").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2021"), "yyyy-mm-dd") Then
'
'                        fmsgboxStatus "* Date of filing of Form 10IE cannot be prior to  01/04/2021"
'                        Sheet1.Range("Sheet1.date10IE_OptOut").Select
'                        Sheet1.Range("Sheet1.date10IE_OptOut").value = ""
'
'                    ElseIf Dformat(Trim(Range("Sheet1.date10IE_OptOut").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'
'                        fmsgbox "* Date of filing of Form 10IE can Not be more than system Date in Income Details"
'                        Range("Sheet1.date10IE_OptOut").Select
'                        Range("Sheet1.date10IE_OptOut").value = ""
'
'                    End If
'
'                ElseIf (Sheet1.Range("AYinOptOut").value = "2022-23") Then
'
''                    If Dformat(Trim(Range("Sheet1.date10IE_OptOut").value), "yyyy-mm-dd") > Dformat(Trim("31/03/2023"), "yyyy-mm-dd") Then
''
''                        fmsgboxStatus "* Date of filing of Form 10IE cannot be beyond 31st March, 2023 in AY 2022-23"
''                        Sheet1.Range("Sheet1.date10IE_OptOut").Select
''                        Sheet1.Range("Sheet1.date10IE_OptOut").value = ""
'
'                    If Dformat(Trim(Range("Sheet1.date10IE_OptOut").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2021"), "yyyy-mm-dd") Then
'
'                        fmsgboxStatus "* Date of filing of Form 10IE cannot be prior to  01/04/2021"
'                        Sheet1.Range("Sheet1.date10IE_OptOut").Select
'                        Sheet1.Range("Sheet1.date10IE_OptOut").value = ""
'
'                    ElseIf Dformat(Trim(Range("Sheet1.date10IE_OptOut").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'
'                        fmsgbox "* Date of filing of Form 10IE can Not be more than system Date in Income Details"
'                        Range("Sheet1.date10IE_OptOut").Select
'                        Range("Sheet1.date10IE_OptOut").value = ""
'
'                     End If
'                End If
'
'            End If
'        End If
'Sheet1.Protect Password:=getmsgstate
'End If
'
'End Sub
'----------------------------------------------------------------

'PAG 2023-24 OLD-------------------------------------------------
'Sub ID_Date10IE(ByVal Target As Range)
'
'If Not Application.Intersect(Sheet1.Range("Sheet1.date10IE"), Sheet1.Range(Target.Address)) Is Nothing Then
'Sheet1.Unprotect Password:=getmsgstate
'
'        If Trim(Range("Sheet1.date10IE").value) <> "00/00/0000" And Range("Sheet1.date10IE").value <> "" Then
'
'            If Not FormatNCheckDate(Trim(Range("Sheet1.date10IE").value)) Then
'
'                fmsgboxStatus "* Date of filing of Form 10IE Is Invalid in Sheet Income Details. Please enter date in dd/mm/yyyy format."
'                Sheet1.Range("Sheet1.date10IE").Select
'                Sheet1.Range("Sheet1.date10IE").value = ""
'
'            Else
'                If Dformat(Trim(Range("Sheet1.date10IE").value), "yyyy-mm-dd") < Dformat("01/04/2022", "yyyy-mm-dd") Then
'                        fmsgboxStatus "* Date of filing of Form 10IE cannot be prior to  01/04/2022 in Income Details"
'                        Range("Sheet1.date10IE").Select
'                        Range("Sheet1.date10IE").value = ""
'
'                ElseIf Dformat(Trim(Range("Sheet1.date10IE").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'                        fmsgboxStatus "* Date of filing of Form 10IE can Not be more than system Date in Income Details"
'                        Range("Sheet1.date10IE").Select
'                        Range("Sheet1.date10IE").value = ""
'
''                ElseIf (Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt") Then
''
''                    If Dformat(Trim(Range("Sheet1.date10IE").value), "yyyy-mm-dd") > Dformat(Trim("31/03/2023"), "yyyy-mm-dd") Then
''                        fmsgboxStatus "* Date of filing of Form 10IE cannot be beyond 31st March, 2023 in Sheet Income Details"
''                        Sheet1.Range("Sheet1.date10IE").Select
''                        Sheet1.Range("Sheet1.date10IE").value = ""
''                    End If
'
''                ElseIf (Sheet1.Range("OptingNewTaxRegime").value = "Opting in now") Or (Sheet1.Range("OptingNewTaxRegime").value = "Opt out") Then
''
''                    If Dformat(Trim(Range("Sheet1.date10IE").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2022"), "yyyy-mm-dd") Then
''                        fmsgboxStatus "* Date of filing cannot be prior to 1st April 2022"
''                        Sheet1.Range("Sheet1.date10IE").Select
''                        Sheet1.Range("Sheet1.date10IE").value = ""
''                    End If
'
'                End If
'
'            End If
'        End If
'
'Sheet1.Protect Password:=getmsgstate
'End If
'
'End Sub
'PAG 2023-24 OLD--------------------------------------------

Sub IDM_115BACFlag()
    If Mid(Trim(Sheet1.Range("sheet1.Status").value), 1, 1) <> "F" Then
        msgError = msgError & "* Selection of either of the options for Have you ever opted for new tax regime u/s 115BAC in earlier years? is mandatory in Sheet : Income Details " & errmsgID & Chr(13)
    End If
End Sub

'PAG_E25 AY 2023-24 OLD
'Sub Validate_CurrentAY23_24()
'     If Sheet8.Range("OldBacValue").value = 1 Then
'
'        If Sheet1.Range("AYinOpting").value = "(Select)" Or Sheet1.Range("AYinOpting").value = "" Then
'            msgError = msgError & "*Please select dropdown in field AY in which Option is exercised" & Chr(13)
'        End If
'
'        If Sheet1.Range("Sheet1.date10IE_Opt").value = "" Then
'            msgError = msgError & "*Please furnish date of filling of form 10-IE where you have exercised the option of new tax regime" & Chr(13)
'        End If
'
''        If Sheet1.Range("Sheet1.Dateoffiling").value = "" Then
''         msgError = msgError & "*sai sunil Please furnish date of filling of form 10-IE where you have exercised the option of new tax regime" & Chr(13)
''        End If
'
'
'        If Sheet1.Range("Sheet1.acknowledgementNumber_Opt").value = "" Then
'            msgError = msgError & "*Acknowledgement Number of Form 10IE is mandatory" & Chr(13)
'        End If
'
'        If Sheet1.Range("Opt_Out115BAC").value = "(Select)" Or Sheet1.Range("Opt_Out115BAC").value = "" Then
'            msgError = msgError & "*Selection of either of the options for 'Have you ever opted out of section 115BAC in earlier years?' is mandatory" & Chr(13)
'        End If
'
'     End If
'
'     If Sheet8.Range("OldBacValue").value = 1 And Sheet1.Range("Opt_Out115BAC").value = "Yes" Then
'
'        If Sheet1.Range("AYinOptOut").value = "(Select)" Or Sheet1.Range("AYinOptOut").value = "" Then
'            msgError = msgError & "*Please select dropdown in field AY in which Option is opted out." & Chr(13)
'        End If
'
''        If Sheet1.Range("Sheet1.date10IE_OptOut").value = "" Then
''            msgError = msgError & "*Please furnish date of filling of form 10-IE where you have withdrawn out of new tax regime" & Chr(13)
''        End If
'
'        If Sheet1.Range("Sheet1.acknowledgementNumber_OptOut").value = "" Then
'            msgError = msgError & "*Acknowledgement Number of Form 10IE is mandatory" & Chr(13)
'        End If
'
'        If Sheet1.Range("Opt_Out115BAC").value = "(Select)" Or Sheet1.Range("Opt_Out115BAC").value = "" Then
'            msgError = msgError & "*Selection of either of the options for 'Have you ever opted out of section 115BAC in earlier years?' is mandatory" & Chr(13)
'        End If
'
'     End If
'
'     If (Sheet1.Range("OptingNewTaxRegime").value = "(Select)" Or Sheet1.Range("OptingNewTaxRegime").value = "") And (Mid(Trim(Sheet1.Range("sheet1.Status").value), 1, 1) <> "F") Then
'            msgError = msgError & "*Please select the applicable option 'Option for current assessment year'" & Chr(13)
'     End If
'
'     If (Sheet1.Range("OptingNewTaxRegime").value = "Opting in now") Or (Sheet1.Range("OptingNewTaxRegime").value = "Opt out") Then
'        If Sheet1.Range("Sheet1.date10IE").value = "" Then
'            msgError = msgError & "*Please furnish date of filling of form 10-IE to opt for new tax regime" & Chr(13)
'        End If
'
'        If Sheet1.Range("Sheet1.acknowledgementNumber").value = "" Then
'            msgError = msgError & "*Acknowledgement Number of Form 10IE is mandatory" & Chr(13)
'        End If
'     End If
'
'
'End Sub
'PAG_E25 AY 2023-24 Old----------------------------------------



'PAG_C25 AY 2024-25 Change Starts------------------------------

  Sub Validate_CurrentAY24_25()
  
    'Ayush_25-26
        If (Sheet1.Range("sheet1.Status").value = "I - INDIVIDUAL") Or (Sheet1.Range("sheet1.Status").value = "H - HUF") Then
           If Sheet1.Range("Sheet1.115BAC").value = "" Or Sheet1.Range("Sheet1.115BAC").value = "(Select)" Then
                'msgError = msgError & "*Please select an option from "" Have you exercised the option u/s 115BAC(6) of Opting out of new tax regime in AY 2024-25?"" " & Chr(13)
                'Updated by Bindu on 5th Feb 2025 as per DE V3
                'Ayush_15/01/2026_V0.2
'                 msgError = msgError & "*Please select an option from "" Have you exercised the option u/s 115BAC(6) of Opting out of new tax regime in Form 10-IEA in AY 2024-25?"" " & Chr(13)
'                  msgError = msgError & "* Please select an option from ""Have you filed form 10IEA within due date  for any earlier assessment year for choosing old tax regime?"" " & Chr(13)
 'Ayush_31/03/2026
                   msgError = msgError & "* Please select an option from ""Have you filed form 10IEA within due date for any earlier assessment year for choosing old tax regime?""" & Chr(13)

           End If
        End If
  
 
'    If (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within the due date") Or (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-but beyond the due date") Then
'        If Sheet1.Range("Sheet1.Dateoffiling").value = "" Then
'            msgError = msgError & "*Please furnish date of filling " & Chr(13)
'        End If
'
'
'
'        If Sheet1.Range("Sheet1.acknowledgementNumber_New").value = "" Then
'            msgError = msgError & "*Acknowledgement Number of Form 10IEA is mandatory" & Chr(13)
'        End If
'        If Sheet1.Range("Sheet1.acknowledgementNumber_New").value <> "" Then
'        If Len(Sheet1.Range("Sheet1.acknowledgementNumber_New").value) > 15 Then
'           msgError = msgError & "*Acknowledgement Number of Form 10IEA is  cannot exceed 15 characters"
'        End If
'         If Len(Sheet1.Range("Sheet1.acknowledgementNumber_New").value) < 15 Then
'           msgError = msgError & "*Acknowledgement Number of Form 10IEA is should 15 characters"
'        End If
'        End If
'
'
'     End If
     
End Sub
'PAG_C25 AY 2024-25 Change End-----------------------------

Sub ID_Allowance_Amount(ByVal Target As Range)
    If (Sheet1.Range(Replace(Target.Address, "AA", "I")).value) = "Sec 10(13A)-Allowance to meet expenditure incurred on house rent" Then
        Sheet1.Unprotect Password:=getmsgstate
        'If Sheet1.Range(Target.Address).value > (0.33 * Sheet1.Range("IncD.IncomeFromSal_ii").value) Then
         
          
           If Sheet1.Range(Target.Address).value > 2000000 Then
                fmsgbox ("*If claim of deduction under exempt allowance under Section 10(13A) is more than Rs 20,00,000, you are required to file return in ITR 3. Please refer Rule 12 for further details.")
                Target.value = ""
           End If
           
         
         'Newly changed by Bindu as per DE & BA team
         If Sheet1.Range(Target.Address).value > Application.WorksheetFunction.RoundUp(Sheet1.Range("IncD.IncomeFromSal_ii").value / 3, 0) Then

         'If Sheet1.Range(Target.Address).value > (0.3333 * Sheet1.Range("IncD.IncomeFromSal_ii").value) Then
            
            
            'fmsgbox ("*Sec 10(13A)-Allowance To meet expenditure incurred On house rent can Not exceed 33.33% of Salary 17(1)")
            
'            Change.02.03.2023.102.IDS.79
'            fmsgbox ("*Sec 10(13A)-Allowance To meet expenditure incurred On house rent can Not exceed 50% of Salary 17(1)")

            'PAG_C34 AY 2023-24 old
'            fmsgbox ("*Exemption of HRA u/s 10(13A) cannot be more than 50% of Salary as per section 17(1)")

            'PAG_C34 AY 2024-25 Change
            'fmsgbox ("*Exemption of HRA u/s 10(13A) claimed cannot be more than limit prescribed ")
'            End Change.IDS79
            Sheet1.Range(Target.Address).value = ""
            Sheet1.Range(Target.Address).Select
        End If
        
        
        
        'sudha_2025-26
          
           
        '=============================
        
        
    End If

    If (Sheet1.Range(Replace(Target.Address, "AA", "I")).value) = "Sec 10(5)-Leave Travel concession/assistance" Then
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range(Target.Address).value > Sheet1.Range("IncD.IncomeFromSal_ii").value Then
            fmsgbox ("*Amount u/s 10(5) can not exceed Salary as per sec 17(1)")
            Sheet1.Range(Target.Address).value = ""
            Sheet1.Range(Target.Address).Select
        End If
    End If
    
    'Ayush_17/03/2026
    If (Sheet1.Range(Replace(Target.Address, "AA", "I")).value) = "Sec 10(17)-Allowance MP/MLA/MLC" Then
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range(Target.Address).value > Sheet1.Range("IncD.IncomeFromSal_ii").value Then
            fmsgbox ("*Amount u/s 10(17) can not exceed Salary as per sec 17(1)")
            Sheet1.Range(Target.Address).value = ""
            Sheet1.Range(Target.Address).Select
        End If
    End If
    
    If (Sheet1.Range(Replace(Target.Address, "AA", "I")).value) = "Sec 10(10A)-Commuted value of pension received" Then
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range(Target.Address).value > Sheet1.Range("IncD.IncomeFromSal_ii").value Then
            fmsgbox ("*Amount u/s 10(10A) can not exceed Salary as per sec 17(1)")
            Sheet1.Range(Target.Address).value = ""
            Sheet1.Range(Target.Address).Select
        End If
    End If
    
    If (Sheet1.Range(Replace(Target.Address, "AA", "I")).value) = "Sec 10(10AA)-Earned leave encashment on Retirement" Then
        Sheet1.Unprotect Password:=getmsgstate
        'Change.03.03.2023.102.IDS.110
        If Sheet1.Range(Target.Address).value > 2500000 Then
'            If Sheet1.Range("sheet1.NatureofEmployment").value <> "(Select)" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Public Sector Undertaking" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-PSU" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-Other" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Others" Then

        'Validation updated as per the Defect on 02-01-2024
        If Sheet1.Range("sheet1.NatureofEmployment").value = "(Select)" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-Other" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Others" Then
           
                'PAG_C34 AY 2023-24 Old
'                fmsgbox ("Deduction allowable for Employees other than Central or State Governments CG - Pensioners or SG - Pensioners is only Rs 3 lakh")
                
                'PAG_C34 AY 2024-25 Change
'                fmsgbox ("Deduction allowable for Employees other than Central or State Governments CG - Pensioners or SG - Pensioners is only Rs 25 lakh")
        ' Error Message updated as per the Defect on 02-01-2024
        'Ayush_11/02/2026_V0.3
                fmsgbox ("Deduction allowable for Employees other than Central or State Governments CG - Pensioners or SG - Pensioners or PSU pensioners or PSU is upto Rs 25 lakh")
'As per DE V0.4 by Bindu on 24th feb 26
'                fmsgbox ("Deduction allowable for Employees other than Central or State Governments, ""Judge as defined in The Supreme Court Judges (Salaries and Conditions of Service) Act, 1958"", CG - Pensioners or SG - Pensioners or PSU pensioners or PSU is upto Rs  25 lakh.")
                Sheet1.Range(Target.Address).value = ""
                Sheet1.Range(Target.Address).Select
            End If
        If Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-PSU" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-Other" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Public Sector Undertaking" Then

            
        End If
    End If
        'End Change IDS 110
        If Sheet1.Range(Target.Address).value > Sheet1.Range("IncD.IncomeFromSal_ii").value Then
            fmsgbox ("*Amount u/s 10(10AA) can not exceed Salary as per sec 17(1)")
            Sheet1.Range(Target.Address).value = ""
            Sheet1.Range(Target.Address).Select
        End If
    End If
    
    If (Sheet1.Range(Replace(Target.Address, "AA", "I")).value) = "Sec 10(10C)-Amount received/receivable on voluntary retirement or termination of service" Then
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range(Target.Address).value > 500000 Then
            fmsgbox ("*Amount u/s 10(10C) can not exceed Rs. 5 lakhs")
            Sheet1.Range(Target.Address).value = ""
            Sheet1.Range(Target.Address).Select
        End If
    End If
    

    If (Sheet1.Range(Replace(Target.Address, "AA", "I")).value) = "Sec 10(10CC)-Tax paid by employer on non-monetary perquisite" Then
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range(Target.Address).value > Sheet1.Range("IncD.IncomeFromSal_iii").value Then
            fmsgbox ("*Amount u/s 10(10CC) can not exceed Value of perquisites as per sec 17(2)")
            Sheet1.Range(Target.Address).value = ""
            Sheet1.Range(Target.Address).Select
        End If
    End If
    
    If (Sheet1.Range(Replace(Target.Address, "AA", "I")).value) = "Sec 10(14)(i)-Prescribed Allowances or benefits (not in a nature of perquisite) specifically granted to meet expenses wholly, necessarily and exclusively and to the extent actually incurred, in performance of duties of office or employment" Then
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range(Target.Address).value > Sheet1.Range("IncD.IncomeFromSal_ii").value Then
            fmsgbox ("*Amount u/s 10(14)(i) can not be more than Salary as per section 17(1)")
            Sheet1.Range(Target.Address).value = ""
            Sheet1.Range(Target.Address).Select
        End If
    End If
    
    If (Sheet1.Range(Replace(Target.Address, "AA", "I")).value) = "Sec 10(14)(ii)-Prescribed Allowances or benefits granted to meet personal expenses in performance of duties of office or employment or to compensate him for increased cost of living." Then
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range(Target.Address).value > Sheet1.Range("IncD.IncomeFromSal_ii").value Then
            fmsgbox ("*Amount u/s 10(14)(ii) can not be more than Salary as per section 17(1)")
            Sheet1.Range(Target.Address).value = ""
            Sheet1.Range(Target.Address).Select
        End If
    End If
    'Change.03.03.2023.102.IDS.85
    If (Sheet1.Range(Replace(Target.Address, "AA", "I")).value) = "Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee" Then
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range(Target.Address).value > 38400 Then
        
        'Error Message Changed By Jyoti for the SIT- 67674 on 25/06/2024
        'fmsgbox ("Transport allowance granted to certain physically handicapped assessee should not exxceed Rs 38,400")
        fmsgbox ("""Exempt allowance under Section 10(14)(ii) cannot exceed Rs 38,400""")
        
            Sheet1.Range(Target.Address).value = ""
            Sheet1.Range(Target.Address).Select
        End If
    End If
    'End Change IDS 85
End Sub


Sub ID_Selection_ReturnFileSec()

    If Range("sheet1.ReturnFileSec").value = "139(4)-After due date" Then
'        If Not ChkMinInclusiveDate(Dformat(Sheet5.Range("Ver.Date").value, "yyyy-mm-dd"), "2023-08-01") Then
'            fmsgbox "Filing section u/s. 139(4) cannot be selected till 31st July, 2023 Or extended due date"
            
            'PAG_C8 AY 2024-25 Change
          'Ayush_25-26
'        If Not ChkMinInclusiveDate(Dformat(Sheet5.Range("Ver.Date").value, "yyyy-mm-dd"), "2024-08-01") Then
'            fmsgbox "Filing section u/s. 139(4) cannot be selected till 31st July, 2024 of such AY Or extended due date"
'        If Not ChkMinInclusiveDate(Dformat(Sheet5.Range("Ver.Date").value, "yyyy-mm-dd"), "2025-08-01") Then
'Ayush_Due Date Changes
'Ayush_26-27_C28
    Dim dueDt As Date
    dueDt = CDate(Sheet9.Range("DueDate_Extend").value)
        'If Not ChkMinInclusiveDate(Dformat(Sheet5.Range("Ver.Date").value, "yyyy-mm-dd"), "2025-09-16") Then
        
        If Not ChkMinInclusiveDate(Dformat(Sheet5.Range("Ver.Date").value, "yyyy-mm-dd"), Dformat(Sheet9.Range("DueDate_Extend_1").value, "yyyy-mm-dd")) Then 'Bindu 8th Sep 2025
            
            
           ' fmsgbox "Filing section u/s. 139(4) cannot be selected till 31st July, 2026 of such AY or extended due date"
      
    fmsgbox "Filing section u/s. 139(4) cannot be selected till " & _
          FormatDateWithOrdinal(dueDt) & " of such AY or extended due date"

      '      fmsgbox "Filing section u/s. 139(4) cannot be selected till 15th September, 2025 of such AY or extended due date"
            Sheet1.Range("sheet1.ReturnFileSec").value = "(Select)"
        End If
    End If
    
    'AYush_27/02/2026
  
'    If Range("sheet1.ReturnFileSec").value = "139(5)-Revised Return" Then
'            If (Dformat(Sheet5.Range("Ver.Date").value, "yyyy-mm-dd")) > "2027-03-31" Then
'                MsgBox ("""Selected Filing section 139(5) cannot be selected post 31st March, 2027""")
'                Range("sheet1.ReturnFileSec1").value = "(Select)"
'            End If
'        End If
    
    
    
    
    If Range("sheet1.ReturnFileSec").value = "139(8A)" Then
'        If (Dformat(Sheet5.Range("Ver.Date").value, "yyyy-mm-dd")) <= ("2023-12-31") Then
'            fmsgbox "Filing section u/s. 139(8A) should be allowed to select only after 31.12.2023 or after extended due date u/s 139(4)"

            'PAG_C8 AY 2024-25 Change
          'Ayush_26-27
          Dim cutoff_139_8A As Date
          cutoff_139_8A = CDate(Sheet8.Range("date_139_8A").value)
          
'        If (Dformat(Sheet5.Range("Ver.Date").value, "yyyy-mm-dd")) <= ("2025-12-31") Then
'        If (Dformat(Sheet5.Range("Ver.Date").value, "yyyy-mm-dd")) <= ("2025-12-31") Then
        If (Dformat(Sheet5.Range("Ver.Date").value, "yyyy-mm-dd")) <= (Dformat(Sheet8.Range("date_139_8A").value, "yyyy-mm-dd")) Then
        
        
'            fmsgbox "139(8A) should be allowed to select only after 31.12.2025 of such AY or after extended due date u/s 139(4)"
            fmsgbox "139(8A) should be allowed to select only after " & Dformat2(cutoff_139_8A, "dd/mm/yyyy") & " of such AY or after extended due date u/s 139(4)"
            Sheet1.Range("sheet1.ReturnFileSec").value = "(Select)"
        End If
    End If
End Sub

'PAG_C33 2024-25 Change-------------------------------------------

'added by sai 0n 24-11-2023 for 115BAC(6) Development
'Sub ID_Nature_Employment_New(ByVal Target As Range)
'    'Application.EnableEvents = False
'    Sheet1.Unprotect Password:=getmsgstate
'        If (Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government") And ((Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "No" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-but beyond the due date")) Then
'            Sheet1.Range("Sheet1.Nature").Validation.Delete
'            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature1"
'        ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government") And (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within the due date") Then
'            Sheet1.Range("Sheet1.Nature").Validation.Delete
'            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature2"
'        ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-CG" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-SG") And ((Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "No" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-but beyond the due date")) Then
'            Sheet1.Range("Sheet1.Nature").Validation.Delete
'            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature3"
'        ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-CG" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-SG") And (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within the due date") Then
'            Sheet1.Range("Sheet1.Nature").Validation.Delete
'            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature4"
'        ElseIf Sheet1.Range("sheet1.NatureofEmployment").value <> "(Select)" And ((Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "No" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-but beyond the due date")) Then
'            Sheet1.Range("Sheet1.Nature").Validation.Delete
'            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature5"
'        ElseIf Sheet1.Range("sheet1.NatureofEmployment").value <> "(Select)" And (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within the due date") Then
'            Sheet1.Range("Sheet1.Nature").Validation.Delete
'            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature6"
'        End If
'    Application.EnableEvents = True
'    Sheet1.Protect Password:=getmsgstate
'End Sub
'PAG_C33 2024-25 Change------------------------------------------

'PAG_E34 2023-24 Old

'Sub ID_Nature_Employment(ByVal Target As Range)
'    Application.EnableEvents = False
'    Sheet1.Unprotect Password:=getmsgstate
'        If (Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government") And ((Sheet1.Range("OptingNewTaxRegime").value = "Opting in now" Or Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt")) Then
'            Sheet1.Range("Sheet1.Nature").Validation.Delete
'            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature1"
'        ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government") And Not ((Sheet1.Range("OptingNewTaxRegime").value = "Opting in now" Or Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt")) Then
'            Sheet1.Range("Sheet1.Nature").Validation.Delete
'            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature2"
'        ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-CG" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-SG") And ((Sheet1.Range("OptingNewTaxRegime").value = "Opting in now" Or Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt")) Then
'            Sheet1.Range("Sheet1.Nature").Validation.Delete
'            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature3"
'        ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-CG" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-SG") And Not ((Sheet1.Range("OptingNewTaxRegime").value = "Opting in now" Or Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt")) Then
'            Sheet1.Range("Sheet1.Nature").Validation.Delete
'            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature4"
'        ElseIf Sheet1.Range("sheet1.NatureofEmployment").value <> "(Select)" And ((Sheet1.Range("OptingNewTaxRegime").value = "Opting in now" Or Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt")) Then
'            Sheet1.Range("Sheet1.Nature").Validation.Delete
'            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature5"
'        ElseIf Sheet1.Range("sheet1.NatureofEmployment").value <> "(Select)" And Not ((Sheet1.Range("OptingNewTaxRegime").value = "Opting in now" Or Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt")) Then
'            Sheet1.Range("Sheet1.Nature").Validation.Delete
'            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature6"
'        End If
'    Application.EnableEvents = True
'    Sheet1.Protect Password:=getmsgstate
'End Sub

'Function ValidateDateDep_ITAY23_24(ByVal Target As Range) As Boolean
'    Dim rangecells As Range
'    Dim i As Long
'    Dim DateDep_IT
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'
'    mIntCells = Sheet18.Range("TaxP.DateDep").count
'    ValidateDateDep_IT = True
'
'    Dim end_IT
'    end_IT = 0
'    Set rangecells = Sheet18.Range("TaxP.DateDep").Cells
'
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).value = "" Then
'            end_IT = end_IT + 1
'        End If
'    Next
'
'    ReDim DateDep_IT(end_IT)
'
'    For i = 1 To end_IT
'        DateDep_IT(i) = rangecells.item(i).value
'        If Not chkCompulsory(DateDep_IT(i)) Then
'            fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is mandatory")
'            ValidateDateDep_ITAY23_24 = False
'            Exit Function
'        End If
'
'        If Not CheckDateddmmyyyy(DateDep_IT(i)) Then
'            ValidateDateDep_ITAY23_24 = False
'            fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is invalid. Pl enter in dd/mm/yyyy format")
'            Target.value = ""
'            Exit Function
'        Else
'            DateDep_IT(i) = Dformat(DateDep_IT(i), "")
'            If Not ChkMinInclusiveDate(DateDep_IT(i), "2022-04-01") Then
'                fmsgboxsmall ("Date of Deposit at Sr. No  " & i & " in Sheet IT  cannot be prior to 01-04-2022")
'                Target.value = ""
'                ValidateDateDep_ITAY23_24 = False
'                Exit Function
'            End If
'            If Sheet1.Range("Sheet1.ReturnFileSec") = "139(8A)" Then
'                If ChkMinInclusiveDate(DateDep_IT(i), "2023-01-01") Then
'                    fmsgboxsmall ("Please enter the details of 140B in Sl No 14A of Part B-ATI and not in this schedule, so as to get the correct computation of your tax liability")
'                    Target.value = ""
'                End If
'            End If
'        End If
'    Next
'
'End Function

'For PAG_E67,PAG_E68,PAG_E69 AY 2023-24 Old--------

'Sub ID_234ABCF_LockUnlock()
'If (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 11) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 12) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 17) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 18) Then
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.IntrstPayUs234A").MergeArea.Locked = False
'    Sheet1.Range("IncD.IntrstPayUs234A").MergeArea.Interior.Color = RGB(255, 255, 204)
'    Sheet1.Range("IncD.IntrstPayUs234B").MergeArea.Locked = False
'    Sheet1.Range("IncD.IntrstPayUs234B").MergeArea.Interior.Color = RGB(255, 255, 204)
''    Sheet1.Range("IncD.IntrstPayUs234C").MergeArea.Locked = False
''    Sheet1.Range("IncD.IntrstPayUs234C").MergeArea.Interior.Color = RGB(255, 255, 204)
'    Sheet1.Range("IncD.IntrstPayUs234F").MergeArea.Locked = False
'    Sheet1.Range("IncD.IntrstPayUs234F").MergeArea.Interior.Color = RGB(255, 255, 204)
'    Sheet1.Protect Password:=getmsgstate
'Else
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.IntrstPayUs234A").MergeArea.Locked = True
'    Sheet1.Range("IncD.IntrstPayUs234A").MergeArea.Interior.ColorIndex = 2
'    Sheet1.Range("IncD.IntrstPayUs234B").MergeArea.Locked = True
'    Sheet1.Range("IncD.IntrstPayUs234B").MergeArea.Interior.ColorIndex = 2
'    ' AY 2024-25 Change
'    Sheet1.Range("IncD.IntrstPayUs234F").MergeArea.Locked = True
'    Sheet1.Range("IncD.IntrstPayUs234F").MergeArea.Interior.ColorIndex = 2
'
'    Sheet1.Protect Password:=getmsgstate
'End If
'End Sub
'For PAG_E67,PAG_E68,PAG_E69 AY 2023-24 Old End------------


'For PAG_C67,PAG_C68,PAG_C69 AY 2024-25 Change Starts--------------------------
Sub ID_234ABCF_LockUnlock_New()
'If (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 11) And (Dformat(Range("Sheet1.DueDate"), "yyyy-mm-dd") < Dformat(Range("Ver.Date"), "yyyy-mm-dd")) Then
'If (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 11) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 12) And (Dformat(Range("Sheet1.DueDate"), "yyyy-mm-dd") < Dformat(Range("Ver.Date"), "yyyy-mm-dd")) Then
'Added by Aavula Naresh on 07/01/2026
If (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 11) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 12) Or (Mid(Range("sheet1.ReturnFileSec1").value, 1, 2) = 21) And (Dformat(Range("Sheet1.DueDate"), "yyyy-mm-dd") < Dformat(Range("Ver.Date"), "yyyy-mm-dd")) Then

    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.IntrstPayUs234A").MergeArea.Locked = False
    Sheet1.Range("IncD.IntrstPayUs234A").MergeArea.Interior.Color = RGB(255, 255, 204)
    Sheet1.Range("IncD.IntrstPayUs234B").MergeArea.Locked = False
    Sheet1.Range("IncD.IntrstPayUs234B").MergeArea.Interior.Color = RGB(255, 255, 204)
'    Sheet1.Range("IncD.IntrstPayUs234C").MergeArea.Locked = False
'    Sheet1.Range("IncD.IntrstPayUs234C").MergeArea.Interior.Color = RGB(255, 255, 204)
    Sheet1.Range("IncD.IntrstPayUs234F").MergeArea.Locked = False
    Sheet1.Range("IncD.IntrstPayUs234F").MergeArea.Interior.Color = RGB(255, 255, 204)
    Sheet1.Protect Password:=getmsgstate
Else
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.IntrstPayUs234A").MergeArea.Locked = True
    Sheet1.Range("IncD.IntrstPayUs234A").MergeArea.Interior.ColorIndex = 2
    Sheet1.Range("IncD.IntrstPayUs234B").MergeArea.Locked = True
    Sheet1.Range("IncD.IntrstPayUs234B").MergeArea.Interior.ColorIndex = 2
    ' AY 2024-25 Change
    Sheet1.Range("IncD.IntrstPayUs234F").MergeArea.Locked = True
    Sheet1.Range("IncD.IntrstPayUs234F").MergeArea.Interior.ColorIndex = 2
    
    Sheet1.Protect Password:=getmsgstate
End If
End Sub
'For PAG_C67,PAG_C68,PAG_C69 AY 2024-25 Change End--------------------------


'Function ValidatePAN_ARN_80GD() As Boolean
'ValidatePAN_ARN_80GD = True
''if pan and arn are same, then show error
'Dim i, j As Long
'For i = 1 To Range("Per5080G.DoneeARN").count
'    If Range("Per5080G.DoneePAN").Cells(i, 1).value <> "" And Range("Per5080G.DoneeARN").Cells(i, 1).value <> "" Then
'    For j = i + 1 To Range("Per5080G.ArnNbr").count
'        If Range("Per5080G.DoneePAN").Cells(i, 1).value = Range("Per5080G.DoneePAN").Cells(j, 1).value And _
'            Range("Per5080G.DoneeARN").Cells(i, 1).value = Range("Per5080G.DoneeARN").Cells(j, 1).value Then
'            errmsg80G = errmsg80G & "*Donation to same donee cannot be entered more than once, please enter different ARN, if applicable."
'            ValidatePAN_ARN_80GD = False
'            Exit Function
'        End If
'    Next
'    End If
'Next
'End Function
'
''Change.27.01.2023.102.80G0.2V
'Function ValidateARNnumber_80GD() As Boolean
''    ValidateARNnumber_80GD = True
''    setTableInfo_D
'    noOfProcessSub = ColCountD
'    Dim rangecells As Range
'    Set rangecells = Range("Per5080G.DoneeARN").Cells
'    Dim i, j As Long
'    Dim PANValue As Variant
'    ReDim ArnNumber_80GD(ColCountD)
'    For i = 1 To Range("Per5080G.DoneeARN").count
'        PANValue = Range("Per5080G.DoneePAN").Cells(i, 1).value
'
'        For j = 1 To Range("Per5080G.DoneeARN").count
'            If PANValue <> "" And PANValue = Range("Per5080G.DoneePAN").Cells(j, 1).value And Range("Per5080G.DoneeARN").Cells(i, 1).value = "" Then
'                errmsg80G = errmsg80G & "* Please enter ARN (Donation reference Number) at Sr.NO " & i & " in Schedule 80G_D." & Chr(13)
'                ValidateARNnumber_80GD = False
'                Exit Function
'            End If
'        Next
'    Next
'End Function



Sub DPMAddComment46(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = getmsgstate
Sheet5.Unprotect Password:=getmsgstate

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet5.Range(targetadd).Comment.Delete
With Sheet5.Range(targetadd)
.AddComment
With .Comment
.text text:="The exemption under 10(26) is available only to certain categories of tax payers in NER and Ladakh"""
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet5.Protect Password:=getmsgstate
End Sub
Sub DPMAddComment45(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = getmsgstate
Sheet5.Unprotect Password:=getmsgstate
    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet5.Range(targetadd).Comment.Delete
With Sheet5.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption is allowed to the extent such income does not exceed one thousand five hundred rupees in respect of each minor child whose income is so includible"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet5.Protect Password:=getmsgstate
End Sub
Sub DPMADeleteComment45(Target As Range)
Dim TargtVal, targetadd As Variant
On Error Resume Next
sPassword = getmsgstate
Sheet5.Unprotect Password:=getmsgstate
    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet5.Range(targetadd).Comment.Delete
Sheet5.Protect Password:=getmsgstate
End Sub



Sub AnyOtherDeductions80CCH(ByVal Target As Range)
'    If (Sheet1.Range("sheet1.DOB").value) = "" Then
'    GoTo Endline
'    End If
'
'    Dim age
'    age = calculateAge(Sheet1.Range("sheet1.DOB").value)
'
    Application.EnableEvents = False
    Sheet1.Unprotect Password:=getmsgstate
    
    'PAG_C61 AY 2023-24 Old
'    If (Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government") And (age >= 17 And age <= 22) And (Mid(Trim(Sheet1.Range("sheet1.Status").value), 1, 1) = "I") Then

    'PAG_C61 2024-25 Change
   ' If (Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government") And (Mid(Trim(Sheet1.Range("sheet1.Status").value), 1, 1) = "I") Then
    'Chandru
    'Ayush_25-26
'    If Range("sheet1.NatureofEmployment").value = "Central Government" And (Dformat((Sheet1.Range("sheet1.DOB").value), "") >= "2001-03-31" And Dformat((Sheet1.Range("sheet1.DOB").value), "") <= "2007-04-01") Then
'      If Range("sheet1.NatureofEmployment").value = "Central Government" And (Dformat((Sheet1.Range("sheet1.DOB").value), "") >= "2001-03-31" And Dformat((Sheet1.Range("sheet1.DOB").value), "") <= "2008-04-01") Then
'Ayush_Legacy51
'If Range("sheet1.NatureofEmployment").value = "Central Government" And (Dformat((Sheet1.Range("sheet1.DOB").value), "") >= "1998-03-31" And Dformat((Sheet1.Range("sheet1.DOB").value), "") <= "2008-04-01") And Sheet1.Range("sheet1.Status").value = "I - INDIVIDUAL" Then
'Ayush_26-27_C160
'Ayush_11/02/2026_V0.3
If Range("sheet1.NatureofEmployment").value = "Central Government" And (Dformat((Sheet1.Range("sheet1.DOB").value), "") >= "1999-03-31" And Dformat((Sheet1.Range("sheet1.DOB").value), "") <= "2009-04-01") And Sheet1.Range("sheet1.Status").value = "I - INDIVIDUAL" Then
'By Bindu on DE V0.4 on 24th Feb 26
'If (Range("sheet1.NatureofEmployment").value = "Central Government" Or Range("sheet1.NatureofEmployment").value = "Judge as defined in The Supreme Court Judges (Salaries and Conditions of Service) Act, 1958") And (Dformat((Sheet1.Range("sheet1.DOB").value), "") >= "1999-03-31" And Dformat((Sheet1.Range("sheet1.DOB").value), "") <= "2009-04-01") And Sheet1.Range("sheet1.Status").value = "I - INDIVIDUAL" Then
        
        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.value = 0
        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Locked = False
        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Interior.Color = "&HCCFFCC"

    Else

        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.value = 0
        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Locked = True
        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Interior.Color = "&HD8D8D8"
    End If

endline:
    Sheet1.Protect Password:=getmsgstate
    Application.EnableEvents = True
End Sub


'Added by Mallikarjun

'IT-C2 AY_2024_2025 change
Function ValidateDateDep_ITAY24_25(ByVal Target As Range) As Boolean
    Dim rangecells As Range
    Dim i As Long
    Dim DateDep_IT
    Dim mIntCells As Long
    Dim mIntCtr As Long
    
    
    mIntCells = Sheet18.Range("TaxP.DateDep").count
    ValidateDateDep_ITAY24_25 = True

    Dim end_IT
    end_IT = 0
    Set rangecells = Sheet18.Range("TaxP.DateDep").Cells

    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            end_IT = end_IT + 1
        End If
    Next

    ReDim DateDep_IT(end_IT)

    For i = 1 To end_IT
        DateDep_IT(i) = rangecells.item(i).value
        If Not chkCompulsory(DateDep_IT(i)) Then
            'fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is mandatory") '
            fmsgboxsmall ("Please enter the date of Deposit at Sr. No  " & i & "  in Sheet IT")
            ValidateDateDep_ITAY24_25 = False
            Exit Function
        End If

        If Not CheckDateddmmyyyy(DateDep_IT(i)) Then
            ValidateDateDep_ITAY24_25 = False
'            fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is invalid. Pl enter in dd/mm/yyyy format")
'Ayush_10/02
'              fmsgboxsmall ("Please enter a valid date in dd/mm/yyyy format at Sr. No  " & i & "  in Sheet IT ")
                fmsgboxsmall ("""Date of Deposit - Please enter valid date in valid format at Sr. No  " & i & "  in Sheet IT."" ")
            Target.value = ""
            Exit Function
        Else
            DateDep_IT(i) = Dformat(DateDep_IT(i), "")
            
            Dim IT_date_deposit As Date
            IT_date_deposit = CDate(Sheet8.Range("IT_Date_Deposit_PY").value)
            IT_date_deposit1 = Dformat(Sheet8.Range("IT_Date_Deposit_PY").value, "")
            'Ayush_13
'            If Not ChkMinInclusiveDate(DateDep_IT(i), "2023-04-01") Then
'             If Not ChkMinInclusiveDate(DateDep_IT(i), "2024-04-01") Then
            If Not ChkMinInclusiveDate(DateDep_IT(i), IT_date_deposit1) Then
            'Error Message Changed By Ayush for the SIT-56341
               ' fmsgboxsmall ("Date of deposit cannot be prior to 01-04-2023 at Sr. No " & i & " in Sheet IT  ")
'               fmsgboxsmall ("""Date of deposit cannot be prior to 01-04-2023""")
'                fmsgboxsmall ("""Date of deposit cannot be prior to 01-04-2024""")
                fmsgboxsmall ("""Date of deposit cannot be prior to " & Dformat3(IT_date_deposit, "dd/mm/yyyy") & """")
                'Ayush_Legacy52
                'Target.value = ""
                rangecells.item(i).ClearContents
                ValidateDateDep_ITAY24_25 = False
                Exit Function
            End If
            
             If DateDep_IT(i) > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
        
                fmsgboxsmall ("Date of Deposit at Sr. No  " & i & " in Sheet IT can not be beyond system date")
                Target.value = ""
                Target.Select
                ValidateDateDep_ITAY24_25 = False
                Exit Function
            End If

             
          '-------------------------------------------------

            If Sheet1.Range("Sheet1.ReturnFileSec") = "139(8A)" Then
'                If ChkMinInclusiveDate(DateDep_IT(i), "2024-01-01") Then
'                    fmsgboxsmall ("Please enter the details of 140B in Sl No 14A of Part B-ATI and not in this schedule, so as to get the correct computation of your tax liability")
 'Ayush_UR_17/11/2025
            If ChkMinInclusiveDate(DateDep_IT(i), "2026-01-01") Then
                   fmsgboxsmall ("Date of deposit can't be after 31-12-2025 in schedule IT. Please enter payments made after 31st December 2025 in schedule IT 1.")
                    
                    
                    Target.value = ""
                End If
            End If
        End If
    Next

End Function

Sub enable()
Application.EnableEvents = True
End Sub
'Chandru
Sub ID_Nature_Employment_New()
    Application.EnableEvents = False
    Sheet1.Unprotect Password:=getmsgstate
    'Ayush_25-26
    
'        If ((Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government") Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-Other" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-PSU") And ((Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "No" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-but beyond the due date")) Then
'         If ((Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government") Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-Other" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-PSU") And Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "New Tax Regime" Then
'Ayush_27/01
'            If (Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government") And Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "New Tax Regime" Then
            If (Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government") And Sheet8.Range("bacValue").value = 1 Then
 'BY Bindu on 24th Feb 26 as per DE V0.4
           ' If (Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Judge as defined in The Supreme Court Judges (Salaries and Conditions of Service) Act, 1958") And Sheet8.Range("bacValue").value = 1 Then

            Sheet1.Range("Sheet1.Nature").Validation.Delete
            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature1"
            
            
            'Added for SIT-84927
            'Ayush_27/01
'            ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-Other" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-PSU") And (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "New Tax Regime") Then
            ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-Other" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-PSU") And (Sheet8.Range("bacValue").value = 1) Then
            Sheet1.Range("Sheet1.Nature").Validation.Delete
            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature7"
            
            'Ayush_05/06/2025
            ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-Other" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-PSU") And (Sheet8.Range("bacValue").value = 2) Then
            Sheet1.Range("Sheet1.Nature").Validation.Delete
            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature8"
            '---------------
            
            '------------------
'        ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government") And (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within the due date") Then
'Ayush_27/01
'         ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government") And (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Old Tax Regime") Then
           'By Bindu on 24th Feb 26 As per DE V0.4
            ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government") And (Sheet8.Range("bacValue").value = 2) Then
          '  ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Central Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "State Government" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Judge as defined in The Supreme Court Judges (Salaries and Conditions of Service) Act, 1958") And (Sheet8.Range("bacValue").value = 2) Then
            Sheet1.Range("Sheet1.Nature").Validation.Delete
            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature2"
            
            
'        ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-CG" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-SG") And ((Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "No" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-but beyond the due date")) Then
'Ayush_27/01
'         ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-CG" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-SG") And ((Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "New Tax Regime")) Then
                     ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-CG" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-SG") And ((Sheet8.Range("bacValue").value = 1)) Then
            Sheet1.Range("Sheet1.Nature").Validation.Delete
            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature3"
            
            
'        ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-CG" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-SG") And (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within the due date") Then
'Ayush_27/01
'         ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-CG" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-SG") And (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Old Tax Regime") Then
                     ElseIf (Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-CG" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-SG") And (Sheet8.Range("bacValue").value = 2) Then
            Sheet1.Range("Sheet1.Nature").Validation.Delete
            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature4"
            
            
'        ElseIf Sheet1.Range("sheet1.NatureofEmployment").value <> "(Select)" And ((Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "No" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-but beyond the due date")) Then
'Ayush_27/01
'         ElseIf Sheet1.Range("sheet1.NatureofEmployment").value <> "(Select)" And ((Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "New Tax Regime")) Then
            ElseIf Sheet1.Range("sheet1.NatureofEmployment").value <> "(Select)" And ((Sheet8.Range("bacValue").value = 1)) Then
            Sheet1.Range("Sheet1.Nature").Validation.Delete
            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature5"
            
            
'        ElseIf Sheet1.Range("sheet1.NatureofEmployment").value <> "(Select)" And (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-within the due date") Then
'Ayush_27/01
'         ElseIf Sheet1.Range("sheet1.NatureofEmployment").value <> "(Select)" And (Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Old Tax Regime") Then
        ElseIf Sheet1.Range("sheet1.NatureofEmployment").value <> "(Select)" And (Sheet8.Range("bacValue").value = 2) Then
            Sheet1.Range("Sheet1.Nature").Validation.Delete
            Sheet1.Range("Sheet1.Nature").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Employment_Nature6"
'        'Newly added by Bindu
'         ElseIf Sheet1.Range("sheet1.NatureofEmployment").value = "(Select)" Or Sheet1.Range("sheet1.NatureofEmployment").value = "" Then
'            Sheet1.Range("Sheet1.Nature").Validation.Delete
'
        End If
    Application.EnableEvents = True
    Sheet1.Protect Password:=getmsgstate


End Sub

'Ayush_25-26
Sub BAC115_24_25(ByVal Target As Range)

If Not Application.Intersect(Sheet1.Range("sheet1.Status"), Sheet1.Range(Target.Address)) Is Nothing Then

Sheet1.Unprotect Password:=getmsgstate
Application.EnableEvents = False

If (Left(Sheet1.Range("sheet1.Status").value, 1) = "F") Then
         
         Sheet1.Unprotect Password:=getmsgstate

         Sheet1.Range("Sheet1.115BAC").Validation.Delete
            Sheet1.Range("Sheet1.115BAC").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BAC115.NA"
            Application.EnableEvents = True
            Sheet1.Range("Sheet1.115BAC").value = "Not Applicable"
            Application.EnableEvents = False
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("Sheet1.115BAC").Interior.Color = (&HD8D8D8)
            Sheet1.Range("Sheet1.115BAC").MergeArea.Locked = True

Else

Sheet1.Unprotect Password:=getmsgstate

            Sheet1.Range("Sheet1.115BAC").Validation.Delete
            Sheet1.Range("Sheet1.115BAC").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BAC115.NY"
            Sheet1.Range("Sheet1.115BAC").value = "(Select)"
            Sheet1.Range("Sheet1.115BAC").Interior.Color = (&HCCFFCC)
            Sheet1.Range("Sheet1.115BAC").MergeArea.Locked = False
            
        
    

End If
Sheet1.Protect Password:=getmsgstate
    Application.EnableEvents = True
    End If
    
End Sub


Sub BAC115filing_24_25(ByVal Target As Range)



If Not Application.Intersect(Sheet1.Range("Sheet1.115BAC"), Sheet1.Range(Target.Address)) Is Nothing Then

Sheet1.Unprotect Password:=getmsgstate
Application.EnableEvents = False


    
        If Sheet1.Range("Sheet1.115BAC").value = "Yes" Then
        
        Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = False
        Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.DateOfFiling").value = ""
    Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = False
        Sheet1.Range("BAC115.AcknowledgementNo").Interior.Color = (&HCCFFCC)
        Sheet1.Range("BAC115.AcknowledgementNo").value = ""
        
        Else
        
        Sheet1.Range("BAC115.DateOfFiling").MergeArea.Locked = True
        Sheet1.Range("BAC115.DateOfFiling").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.DateOfFiling").value = ""
    
        Sheet1.Range("BAC115.AcknowledgementNo").MergeArea.Locked = True
        Sheet1.Range("BAC115.AcknowledgementNo").Interior.Color = (&HD8D8D8)
        Sheet1.Range("BAC115.AcknowledgementNo").value = ""
        
        End If
Sheet1.Protect Password:=getmsgstate
    Application.EnableEvents = True
    End If
    
End Sub


'Ayush_25 -26
Function ID_Opt_Date10IEA_New() As Boolean
ID_Opt_Date10IEA_New = True
        
            Sheet1.Unprotect Password:=getmsgstate
          If Trim(Range("BAC115.DateOfFiling").value) <> "00/00/0000" And Range("BAC115.DateOfFiling").value <> "" Then
        
            If Not FormatNCheckDate(Trim(Range("BAC115.DateOfFiling").value)) Then
            'Ayush_10/02
                msgError = msgError & """Date of filing of form 10-IEA for AY 2024-25 - Please enter valid date in valid format.""" & Chr(13)
                ID_Opt_Date10IEA_New = False
                
                Sheet1.Range("BAC115.DateOfFiling").value = ""
                
            Else
        If Sheet1.Range("Sheet1.115BAC").value = "Yes" Then 'added  by sai on 27/01/2025 for AY 2025-26
    
                    If Dformat(Trim(Range("BAC115.DateOfFiling").value), "yyyy-mm-dd") < Dformat(Trim("01/04/2024"), "yyyy-mm-dd") Then
                        
                            msgError = msgError & """Date of filing of Form 10IEA cannot be prior to  01/04/2024.""" & Chr(13)
                            ID_Opt_Date10IEA_New = False

                            Sheet1.Range("BAC115.DateOfFiling").value = ""
        
                    ElseIf Dformat(Trim(Range("BAC115.DateOfFiling").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
                        
                            msgError = msgError & " Date of filing of Form 10IEA can Not be more than system Date in Income Details" & Chr(13)
                            ID_Opt_Date10IEA_New = False
                           
                            Range("BAC115.DateOfFiling").value = ""
                            
                    End If
                End If
            End If
        End If
        
        'Ayush_28/02
         If Trim(Range("BAC115.DateOfFiling").value) = "00/00/0000" Then
                msgError = msgError & """Date of filing of form 10-IEA for AY 2024-25 - Please enter valid date in valid format.""" & Chr(13)
                ID_Opt_Date10IEA_New = False
                
                Sheet1.Range("BAC115.DateOfFiling").value = ""
        End If
        
        Sheet1.Protect Password:=getmsgstate


End Function


'Ayush_25-26
Function Validate_CurrentAY25_26() As Boolean
Validate_CurrentAY25_26 = True

    'Ayush_25-26
'        If (Sheet1.Range("sheet1.Status").value = "I - INDIVIDUAL") Or (Sheet1.Range("sheet1.Status").value = "H - HUF") Then
'            If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "(Select)" Then
'
'                 msgError = msgError & "*Please select Option for current assessment year in (A23b)" & Chr(13)
'
'            End If
'        End If

    If (Sheet1.Range("Sheet1.115BAC").value = "Yes") Then
'        If Sheet1.Range("BAC115.DateOfFiling").value = "" Then
''            msgError = msgError & "*If A23a is ""Yes"" Date of filing of Form 10IEA is mandatory. " & Chr(13)
''Ayush_10/02
''            msgError = msgError & "*""Date of filing of Form 10IEA is mandatory.""" & Chr(13) 'Modified by sai on 27/01/2025
'            msgError = msgError & "*""Please enter Date of filing""" & Chr(13)
'            Validate_CurrentAY25_26 = False
'        End If

        If Sheet1.Range("BAC115.AssessmentYear_2425").value = "" Or Sheet1.Range("BAC115.AssessmentYear_2425").value = "(Select)" Then
            msgError = msgError & "* Please select a dropdown from ""Assessment year for which form 10IEA for choosing old tax regime was file""" & Chr(13)
            Validate_CurrentAY25_26 = False
        End If
        
        If Sheet1.Range("BAC115.AcknowledgementNo").value = "" Then
'            msgError = msgError & "*If A23a is ""Yes"" Acknowledgement Number of Form 10IEA is mandatory." & Chr(13)
            msgError = msgError & "* ""Acknowledgement Number of Form 10-IEA"" is mandatory." & Chr(13) 'Modified by sai on 27/01/2025
            Validate_CurrentAY25_26 = False
        End If

'Ayush_09/06
'        If Dformat(Trim(Sheet1.Range("BAC115.DateOfFiling").value), "yyyy-mm-dd") > Dformat(Trim(Sheet5.Range("Ver.Date").value), "yyyy-mm-dd") Then
'    '    If Not Sheet1.ChkMaxDate_filing_2425(Trim(Sheet1.Range("BAC115.DateOfFiling").value), "31-07-2024") Then
'                'fmsgbox "*""Please select ‘No’, if Form 10IEA was filed after due date for AY 2024-25."""
'               '  msgError = msgError & "* ""Please select ‘No’, if Form 10IEA was filed after due date for AY 2024-25.""" & Chr(13)
'                                  msgError = msgError & "* ""Date of filing of Form 10IEA for AY 2024-25 can not be more than system Date in Income Details""" & Chr(13)
'                 Validate_CurrentAY25_26 = False
'        End If
        
        If Sheet1.Range("BAC115.AcknowledgementNo").value <> "" Then
           If Len(Sheet1.Range("BAC115.AcknowledgementNo").value) > 15 Then
              msgError = msgError & "*Acknowledgement Number of Form 10-IEA is  cannot exceed 15 characters" & Chr(13)
              Validate_CurrentAY25_26 = False
          End If
          
         If Len(Sheet1.Range("BAC115.AcknowledgementNo").value) < 15 Then
           msgError = msgError & "*Acknowledgement Number of Form 10-IEA is should 15 characters" & Chr(13)
           Validate_CurrentAY25_26 = False
         End If
        End If

     End If
     'Ayush_15/01/2026_V0.2_commented
 '    If Not ID_Opt_Date10IEA_New Then Validate_CurrentAY25_26 = False
     
End Function
'Ayush_25-26
'Sub Validate_dateFiling_25_26()
''    If Sheet1.Range("Sheet1.115BAC").value = "No" And Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Old Tax Regime" Then
'    If Sheet1.Range("Sheet1.115BAC").value = "No" And Sheet8.Range("bacValue").value = 2 Then
'            'If Sheet1.Range("Sheet1.Dateoffiling") > "31/07/2025" Then
'            ' If Dformat(Trim(Sheet1.Range("Sheet1.Dateoffiling").value), "yyyy-mm-dd") >= Dformat(Trim("01/08/2025"), "yyyy-mm-dd") Then
'
'             If Sheet1.Range("Sheet1.Dateoffiling").value <> "" Then
'              If Not Sheet1.ChkMaxDate_filing(Trim(Sheet1.Range("Sheet1.Dateoffiling").value), "31-07-2025") Then
'                 msgError = msgError & "*Since form 10IEA is filed after due date option for current assessment year cannot be Old tax regime" & Chr(13)
'              End If
'             End If
'    End If
'
''    If Sheet1.Range("Sheet1.115BAC").value = "No" And Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Old Tax Regime" Or Sheet1.Range("Sheet1.115BAC").value = "Yes" And Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "New Tax Regime" Then
'    If Sheet1.Range("Sheet1.115BAC").value = "No" And Sheet8.Range("bacValue").value = 2 Or Sheet1.Range("Sheet1.115BAC").value = "Yes" And Sheet8.Range("bacValue").value = 1 Then
'         If Sheet1.Range("Sheet1.Dateoffiling") = "" Then
'                 msgError = msgError & "*Date of filing of Form 10IEA is mandatory." & Chr(13)
'         End If
'         'Malli------------
'         If Sheet1.Range("Sheet1.Dateoffiling").value <> "" Then
'            If Not FormatNCheckDate(Trim(Range("Sheet1.Dateoffiling").value)) Then
'                msgError = msgError & "* ""Please enter date in dd/mm/yyyy format.""" & Chr(13)
'            End If
'
'            If ChkMaxDate(Sheet1.Range("Sheet1.Dateoffiling").value, "") Then
'                msgError = msgError & "* ""Date of filing of Form 10IEA cannot be prior to  01/04/2025.""" & Chr(13)
'            End If
'          End If
'        '-----------------
'         If Sheet1.Range("Sheet1.acknowledgementNumber_New") = "" Then
'                 msgError = msgError & "*Acknowledgement Number of Form 10IEA is mandatory." & Chr(13)
'         End If
'    End If
'
'End Sub


Sub BAC115filing_25_26(ByVal Target As Range)
'Malli-------------------12/12/2024
If Not Application.Intersect(Sheet1.Range("Sheet1.OptOutNewTaxRegime"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("Sheet1.115BAC"), Sheet1.Range(Target.Address)) Is Nothing Then
Sheet1.Unprotect Password:=getmsgstate
Application.EnableEvents = False


If Not isdropdownblank(Sheet1.Range("Sheet1.115BAC").value) And Not isdropdownblank(Sheet1.Range("Sheet1.OptOutNewTaxRegime").value) Then
 
'If (Sheet1.Range("Sheet1.115BAC").value = "Yes" And Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "New Tax Regime") Or _
'   (Sheet1.Range("Sheet1.115BAC").value = "No" And Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Old Tax Regime") Then
If (Sheet1.Range("Sheet1.115BAC").value = "Yes" And Sheet8.Range("bacValue").value = 1) Or _
   (Sheet1.Range("Sheet1.115BAC").value = "No" And Sheet8.Range("bacValue").value = 2) Then
        Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Locked = False
        Sheet1.Range("Sheet1.Dateoffiling").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.Dateoffiling").value = ""
        Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.Locked = False
        Sheet1.Range("Sheet1.acknowledgementNumber_New").Interior.Color = (&HCCFFCC)
        Sheet1.Range("Sheet1.acknowledgementNumber_New").value = ""
        Else
        Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Dateoffiling").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Dateoffiling").value = ""
        Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.Locked = True
        Sheet1.Range("Sheet1.acknowledgementNumber_New").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.acknowledgementNumber_New").value = ""
        End If
Else
        Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Locked = True
        Sheet1.Range("Sheet1.Dateoffiling").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.Dateoffiling").value = ""
        Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.Locked = True
        Sheet1.Range("Sheet1.acknowledgementNumber_New").Interior.Color = (&HD8D8D8)
        Sheet1.Range("Sheet1.acknowledgementNumber_New").value = ""
End If
Sheet1.Protect Password:=getmsgstate
    Application.EnableEvents = True
    End If
End Sub
'
'Sub BAC115filing_25_26(ByVal target As Range)
'
'
'If Not Application.Intersect(Sheet1.Range("Sheet1.OptOutNewTaxRegime"), Sheet1.Range(target.Address)) Is Nothing Then
'
'Sheet1.Unprotect Password:=getmsgstate
'Application.EnableEvents = False
'
'
'
'        If Sheet1.Range("Sheet1.115BAC").value = "Yes" And Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "New Tax Regime" Or Sheet1.Range("Sheet1.115BAC").value = "No" And Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Old Tax Regime" Then
'        Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.Dateoffiling").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("Sheet1.Dateoffiling").value = ""
'        Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.Locked = False
'        Sheet1.Range("Sheet1.acknowledgementNumber_New").Interior.Color = (&HCCFFCC)
'        Sheet1.Range("Sheet1.acknowledgementNumber_New").value = ""
'        Else
'        Sheet1.Range("Sheet1.Dateoffiling").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.Dateoffiling").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.Dateoffiling").value = ""
'        Sheet1.Range("Sheet1.acknowledgementNumber_New").MergeArea.Locked = True
'        Sheet1.Range("Sheet1.acknowledgementNumber_New").Interior.Color = (&HD8D8D8)
'        Sheet1.Range("Sheet1.acknowledgementNumber_New").value = ""
'        End If
'Sheet1.Protect Password:=getmsgstate
'    Application.EnableEvents = True
'    End If
'End Sub

Sub ID_Allowance_Amount_2(ByVal Target As String)
    

    If (Sheet1.Range(Replace(Target, "AA", "I")).value) = "Sec 10(13A)-Allowance to meet expenditure incurred on house rent" Then
        Sheet1.Unprotect Password:=getmsgstate
        'If Sheet1.Range(Target.Address).value > (0.33 * Sheet1.Range("IncD.IncomeFromSal_ii").value) Then
         
         
         'Newly changed by Bindu as per DE & BA team
         If Sheet1.Range(Target).value > Application.WorksheetFunction.RoundUp(Sheet1.Range("IncD.IncomeFromSal_ii").value / 3, 0) Then

         'If Sheet1.Range(Target.Address).value > (0.3333 * Sheet1.Range("IncD.IncomeFromSal_ii").value) Then
            
            
            'fmsgbox ("*Sec 10(13A)-Allowance To meet expenditure incurred On house rent can Not exceed 33.33% of Salary 17(1)")
            
'            Change.02.03.2023.102.IDS.79
'            fmsgbox ("*Sec 10(13A)-Allowance To meet expenditure incurred On house rent can Not exceed 50% of Salary 17(1)")

            'PAG_C34 AY 2023-24 old
'            fmsgbox ("*Exemption of HRA u/s 10(13A) cannot be more than 50% of Salary as per section 17(1)")

            'PAG_C34 AY 2024-25 Change
            'fmsgbox ("*Exemption of HRA u/s 10(13A) claimed cannot be more than limit prescribed ")
'            End Change.IDS79
            Sheet1.Range(Target).value = ""
            Sheet1.Range(Target).Select
        End If
    End If

    If (Sheet1.Range(Replace(Target, "AA", "I")).value) = "Sec 10(5)-Leave Travel concession/assistance" Then
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range(Target).value > Sheet1.Range("IncD.IncomeFromSal_ii").value Then
            fmsgbox ("*Amount u/s 10(5) can not exceed Salary as per sec 17(1)")
            Sheet1.Range(Target).value = ""
            Sheet1.Range(Target).Select
        End If
    End If

    If (Sheet1.Range(Replace(Target, "AA", "I")).value) = "Sec 10(10A)-Commuted value of pension received" Then
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range(Target).value > Sheet1.Range("IncD.IncomeFromSal_ii").value Then
            fmsgbox ("*Amount u/s 10(10A) can not exceed Salary as per sec 17(1)")
            Sheet1.Range(Target).value = ""
            Sheet1.Range(Target).Select
        End If
    End If

    If (Trim(Sheet1.Range(Replace(Target, "AA", "I")).value)) = Trim("Sec 10(10AA)-Earned leave encashment on Retirement") Then
        Sheet1.Unprotect Password:=getmsgstate
        'Change.03.03.2023.102.IDS.110
        If Sheet1.Range(Target).value > 2500000 Then
'            If Sheet1.Range("sheet1.NatureofEmployment").value <> "(Select)" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Public Sector Undertaking" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-PSU" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-Other" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Others" Then

        'Validation updated as per the Defect on 02-01-2024
        If Sheet1.Range("sheet1.NatureofEmployment").value = "(Select)" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-Other" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Others" Then

                'PAG_C34 AY 2023-24 Old
'                fmsgbox ("Deduction allowable for Employees other than Central or State Governments CG - Pensioners or SG - Pensioners is only Rs 3 lakh")

                'PAG_C34 AY 2024-25 Change
'                fmsgbox ("Deduction allowable for Employees other than Central or State Governments CG - Pensioners or SG - Pensioners is only Rs 25 lakh")
        ' Error Message updated as per the Defect on 02-01-2024
        'Ayush_11/02/2026_V0.3
                fmsgbox ("Deduction allowable for Employees other than Central or State Governments CG - Pensioners or SG - Pensioners or PSU pensioners or PSU is upto Rs 25 lakh")
                'By Bindu on 24th Feb 26 as per DE V0.4
                'fmsgbox ("Deduction allowable for Employees other than Central or State Governments, ""Judge as defined in The Supreme Court Judges (Salaries and Conditions of Service) Act, 1958"", CG - Pensioners or SG - Pensioners or PSU pensioners or PSU is upto Rs  25 lakh.")

                Sheet1.Range(Target).value = ""
                Sheet1.Range(Target).Select
            End If
        If Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-PSU" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Pensioners-Other" Or Sheet1.Range("sheet1.NatureofEmployment").value = "Public Sector Undertaking" Then


        End If
    End If
        'End Change IDS 110
        If Sheet1.Range(Target).value > Sheet1.Range("IncD.IncomeFromSal_ii").value Then
            fmsgbox ("*Amount u/s 10(10AA) can not exceed Salary as per sec 17(1)")
            Sheet1.Range(Target).value = ""
            Sheet1.Range(Target).Select
        End If
    End If

    If (Sheet1.Range(Replace(Target, "AA", "I")).value) = "Sec 10(10C)-Amount received/receivable on voluntary retirement or termination of service" Then
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range(Target).value > 500000 Then
            fmsgbox ("*Amount u/s 10(10C) can not exceed Rs. 5 lakhs")
            Sheet1.Range(Target).value = ""
            Sheet1.Range(Target).Select
        End If
    End If


    If (Sheet1.Range(Replace(Target, "AA", "I")).value) = "Sec 10(10CC)-Tax paid by employer on non-monetary perquisite" Then
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range(Target).value > Sheet1.Range("IncD.IncomeFromSal_iii").value Then
            fmsgbox ("*Amount u/s 10(10CC) can not exceed Value of perquisites as per sec 17(2)")
            Sheet1.Range(Target).value = ""
            Sheet1.Range(Target).Select
        End If
    End If

    If (Sheet1.Range(Replace(Target, "AA", "I")).value) = "Sec 10(14)(i)-Prescribed Allowances or benefits (not in a nature of perquisite) specifically granted to meet expenses wholly, necessarily and exclusively and to the extent actually incurred, in performance of duties of office or employment" Then
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range(Target).value > Sheet1.Range("IncD.IncomeFromSal_ii").value Then
            fmsgbox ("*Amount u/s 10(14)(i) can not be more than Salary as per section 17(1)")
            Sheet1.Range(Target).value = ""
            Sheet1.Range(Target).Select
        End If
    End If

    If (Sheet1.Range(Replace(Target, "AA", "I")).value) = "Sec 10(14)(ii)-Prescribed Allowances or benefits granted to meet personal expenses in performance of duties of office or employment or to compensate him for increased cost of living." Then
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range(Target).value > Sheet1.Range("IncD.IncomeFromSal_ii").value Then
            fmsgbox ("*Amount u/s 10(14)(ii) can not be more than Salary as per section 17(1)")
            Sheet1.Range(Target).value = ""
            Sheet1.Range(Target).Select
        End If
    End If
    'Change.03.03.2023.102.IDS.85
    If (Sheet1.Range(Replace(Target, "AA", "I")).value) = "Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee" Then
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range(Target).value > 38400 Then

        'Error Message Changed By Jyoti for the SIT- 67674 on 25/06/2024
        'fmsgbox ("Transport allowance granted to certain physically handicapped assessee should not exxceed Rs 38,400")
        fmsgbox ("""Exempt allowance under Section 10(14)(ii) cannot exceed Rs 38,400""")

            Sheet1.Range(Target).value = ""
            Sheet1.Range(Target).Select
        End If
    End If
    'End Change IDS 85
End Sub

' === Formats a date as: 31st July, 2026 ===
Private Function FormatDateWithOrdinal(ByVal dt As Date) As String
    Dim d As Integer, suffix As String
    d = day(dt)
    
    ' Determine suffix (account for 11th, 12th, 13th)
    If d Mod 100 >= 11 And d Mod 100 <= 13 Then
        suffix = "th"
    Else
        
        Select Case d Mod 10
            Case 1: suffix = "st"
            Case 2: suffix = "nd"
            Case 3: suffix = "rd"
            Case Else: suffix = "th"
        End Select
    End If
    
    
 ' Example: 31st July, 2026
    FormatDateWithOrdinal = CStr(d) & suffix & " " & Format(dt, "mmmm, yyyy")
End Function

'AYush_27/02/2026
Function ChkMaxInclusiveDate_return(Maxinclusive As Variant, Maxinclusivedate As Variant) As Boolean
'Note:
'' both date must be in format yyyy-mm-dd
    
    ChkMaxInclusiveDate_return = True
    If Len(Maxinclusive) > 0 Then
        If Mid(Maxinclusive, 1, 4) > Mid(Maxinclusivedate, 1, 4) Then
            ChkMaxInclusiveDate_return = False
            Exit Function
        Else
            If Mid(Maxinclusive, 1, 4) = Mid(Maxinclusivedate, 1, 4) Then
                If (Mid(Maxinclusive, 6, 2) > Mid(Maxinclusivedate, 6, 2)) Then
                    ChkMaxInclusiveDate_return = False
                    Exit Function
                ElseIf ((Mid(Maxinclusive, 6, 2) = Mid(Maxinclusivedate, 6, 2))) Then
                    If (Mid(Maxinclusive, 9, 2) > Mid(Maxinclusivedate, 9, 2)) Then
                        ChkMaxInclusiveDate_return = False
                        Exit Function
                   End If
                End If
            End If
        End If
    End If
End Function

Public Function ChkMaxDate_return(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

     ChkMaxDate_return = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))


            If Year > 2027 Then
           
            ChkMaxDate_return = False
            Exit Function
        Else

                If Year = 2027 Then

                If month > 4 Then
                    ChkMaxDate_return = False
                    Exit Function
                Else
                    If month = 4 Then
                        If dat > 1 Then
                            ChkMaxDate_return = False
                            Exit Function
                        Else
                            If dat = 1 Then
                                ChkMaxDate_return = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function

Sub DPMAddComment47(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = getmsgstate
Sheet5.Unprotect Password:=getmsgstate

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet5.Range(targetadd).Comment.Delete
With Sheet5.Range(targetadd)
.AddComment
With .Comment
.text text:="The exemption under 10(26AAA) is available only to certain categories of Sikkimese tax payers"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet5.Protect Password:=getmsgstate
End Sub
Sub DPMAddComment48(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = getmsgstate
Sheet5.Unprotect Password:=getmsgstate

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet5.Range(targetadd).Comment.Delete
With Sheet5.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption allowed to the extent it does not exceed sixty per cent of the total amount payable  at the time of such closure or opting out of the scheme"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet5.Protect Password:=getmsgstate
End Sub
Sub DPMAddComment49(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = getmsgstate
Sheet5.Unprotect Password:=getmsgstate

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet5.Range(targetadd).Comment.Delete
With Sheet5.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption allowed to the extent that it does not exceed sixty per cent. of the individual corpus"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet5.Protect Password:=getmsgstate
End Sub
Sub DPMAddComment50(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = getmsgstate
Sheet5.Unprotect Password:=getmsgstate

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet5.Range(targetadd).Comment.Delete
With Sheet5.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption allowed to the extent it does not exceed twenty-five per cent of the amount of contributions made by him;"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet5.Protect Password:=getmsgstate
End Sub
Sub DPMAddComment51(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = getmsgstate
Sheet5.Unprotect Password:=getmsgstate

    TargtVal = Target.value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet5.Range(targetadd).Comment.Delete
With Sheet5.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption allowed to the extent it does not exceed twenty-five per cent of the amount of contributions made by him;"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet5.Protect Password:=getmsgstate
End Sub



Sub ttttttttttttttttttttttttttttttttttttttttt()
Application.EnableEvents = True
End Sub
