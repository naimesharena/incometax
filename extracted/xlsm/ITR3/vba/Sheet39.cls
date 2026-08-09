Attribute VB_Name = "Sheet39"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Private Sub Worksheet_Activate()
On Error Resume Next

    Sheet39.Unprotect Password:=getmsgstate
    ActiveSheet.Shapes("Button 3").Left = Range("Button_SAL").Left + 2
    ActiveSheet.Shapes("Button 3").Top = Range("Button_SAL").Top + 3

    ActiveSheet.Shapes("Button 78").Left = Range("Button_SAL").Left + 120
    ActiveSheet.Shapes("Button 78").Top = Range("Button_SAL").Top + 3

    ActiveSheet.Shapes("Button 80").Left = Range("Button_SAL").Left + 235
    ActiveSheet.Shapes("Button 80").Top = Range("Button_SAL").Top + 3

    ActiveSheet.Shapes("Button 79").Left = Range("Button_SAL").Left + 350
    ActiveSheet.Shapes("Button 79").Top = Range("Button_SAL").Top + 3

    ActiveSheet.Shapes("Button 16").Left = Range("Button_SAL1").Left + 2
    ActiveSheet.Shapes("Button 16").Top = Range("Button_SAL1").Top + 2

    Sheet39.Protect Password:=getmsgstate

End Sub
Private Sub Worksheet_Changeold1(ByVal Target As Range)
On Error GoTo endline

CalculateTaxCheckFlag = False

Dim count, i As Variant
Application.EnableEvents = False


calSchS
count = Sheet39.Range("SalarySectionCount").value

Dim temp_Pensioner As Variant
Dim temp_Others, temp As Variant
Dim Salcount1 As Long
Salcount1 = 0
temp_Pensioner = 0
temp_Others = 0
temp = 0
For Salcount1 = 1 To Sheet39.Range("SalarySectionCount").value

 If Not isdropdownblank(Sheet39.Range("sheet.EmployerCategory" & Salcount1).value) Then
 temp = temp + 1
 End If

' If (Mid(Sheet39.Range("sheet.EmployerCategory" & Salcount1).value, 1, 3) = "Pen") Then
 If (Sheet39.Range("sheet.EmployerCategory" & Salcount1).value = "CG-Pensioners") Or (Sheet39.Range("sheet.EmployerCategory" & Salcount1).value = "SG-Pensioners") Or (Sheet39.Range("sheet.EmployerCategory" & Salcount1).value = "PSU-Pensioners") Or (Sheet39.Range("sheet.EmployerCategory" & Salcount1).value = "Others-Pensioners") Then
 temp_Pensioner = temp_Pensioner + 1
 End If

 If Mid(Sheet39.Range("sheet.EmployerCategory" & Salcount1).value, 1, 1) = "O" Then
 temp_Others = temp_Others + 1
 End If
Next


If temp = temp_Pensioner Or temp = temp_Others Then
Sheet39.Unprotect Password:=getmsgstate
Sheet39.Range("SAL.Entertainment_allowance_1").value = ""
Sheet39.Range("SAL.Entertainment_allowance_1").Locked = True
Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HD8D8D8"
Sheet39.Protect Password:=getmsgstate
Else
Sheet39.Unprotect Password:=getmsgstate
Sheet39.Range("SAL.Entertainment_allowance_1").Locked = False
Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HCCFFCC"
Sheet39.Protect Password:=getmsgstate
End If

If temp = (temp_Pensioner + temp_Others) And temp_Pensioner > 0 And temp_Others > 0 Then
Sheet39.Unprotect Password:=getmsgstate
Sheet39.Range("SAL.Entertainment_allowance_1").value = ""
Sheet39.Range("SAL.Entertainment_allowance_1").Locked = True
Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HD8D8D8"
Sheet39.Protect Password:=getmsgstate
End If

If temp = temp_Pensioner Then
Sheet39.Unprotect Password:=getmsgstate
Sheet39.Range("SAL.DeductionUnderSection1").value = ""
Sheet39.Range("SAL.DeductionUnderSection1").Locked = True
Sheet39.Range("SAL.DeductionUnderSection1").Interior.Color = "&HD8D8D8"
Sheet39.Protect Password:=getmsgstate
Else
Sheet39.Unprotect Password:=getmsgstate
Sheet39.Range("SAL.DeductionUnderSection1").Locked = False
Sheet39.Range("SAL.DeductionUnderSection1").Interior.Color = "&HCCFFCC"
Sheet39.Protect Password:=getmsgstate
End If

If Not Intersect(Range("Salary.Amount1"), Range(Target.Address)) Is Nothing Then
'If Mid(Sheet39.Range(Replace(Target.address, "J", "G")).value, 1, 11) = "Sec 10(13A)" Then
'If Sheet39.Range(Target.address).value > WorksheetFunction.Min(0.5 * Sheet19.Range("Temp_DA_BS").value, Sheet19.Range("Temp_HRA").value) Then
'    fmsgboxStatus "Sec 10(13A)-Allowance to meet expenditure incurred on house rent- amount of exemption should be restricted to min of amount entered under  ""House Rent Allowance"" in sl no B1a or 50% of ""Basic+DA"" as under dropdowns of 1(a)"
'    Sheet39.Range(Target.address).ClearContents
'End If
'End If

    If Range(Target.Address).value > 2000000 Then '= "Sec 10(10)-Death-cum-retirement gratuity received " Then
'        If temp_Pensioner > 0 Or temp_Others > 0 Then
        If Sheet39.Range(Replace(Target.Address, "J", "G")).value = "Sec 10(10)-Death-cum-retirement gratuity received " Then
        fmsgboxStatus " For Sec 10(10)-Death-cum-retirement gratuity is selected than amount should not exceed .20 Lakhs "
        Sheet39.Range(Target.Address).value = ""
        End If
'        End If
    End If

    If Range(Target.Address).value > 500000 Then '= "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
    If Sheet39.Range(Replace(Target.Address, "J", "G")).value = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
    fmsgboxStatus "Exempt Allowance u/s 10(10B)-First Proviso-  Compensation limit notified by CG in the Official Gazette can not exceed Rs.500,000"
    Sheet39.Range(Target.Address).value = ""
    End If
    End If
    
    
    'new change
    If Range(Target.Address).value > 38400 Then '= "Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee" Then
    If Sheet39.Range(Replace(Target.Address, "J", "G")).value = "Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee" Then
    fmsgboxStatus "Exempt Allowance u/s Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee can not exceed Rs.38400"
    Sheet39.Range(Target.Address).value = ""
    End If
    End If
    'new change

'    If Range(Target.address).value > 500000 Then '= "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
'    If Sheet39.Range(Replace(Target.address, "J", "G")).value = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
'    fmsgboxStatus "Exempt allowance u/s 10(10B) Second Proviso- Compensation under scheme approved by CG cannot exceed Rs. 5 lakhs."
'    Sheet39.Range(Target.address).value = ""
'    End If
'    End If




End If


If Target.Validation.Type = 3 Then
    If Not Intersect(Range("Salary.Nature1"), Range(Target.Address)) Is Nothing Then
sPassword = EfilingCommon.getmsgstate
         Sheet39.Unprotect Password:=sPassword

    'Dim ChangeAddress As String
         ChangeAddress = Replace(Target.Address, "G", "H")
    If Mid(Sheet39.Range(Target.Address).value, 1, 1) = "A" Then
         Sheet39.Range(ChangeAddress).Locked = False
         Sheet39.Range(ChangeAddress).Interior.Color = "&HCCFFCC"
         Sheet39.Range(ChangeAddress).ClearContents
    Else
         Sheet39.Range(ChangeAddress).Locked = True
         Sheet39.Range(ChangeAddress).Interior.Color = "&HD8D8D8"
         Sheet39.Range(ChangeAddress).ClearContents
    End If

      Sheet39.Protect Password:=sPassword
    Dim nOfCount1 As Long
                nOfCount1 = CountDropdownAlwnc(Trim(Target.value), Sheet39.Range("Salary.Nature1").Cells, Sheet39.Range("Salary.Nature1").count)
'                If currCntAlwnc1 > 1 Or currCntAlwnc2 > 1 Or currCntAlwnc3 > 1 Or currCntAlwnc4 > 1 Or currCntAlwnc5 > 1 Or currCntAlwnc6 > 1 Or currCntAlwnc7 > 1 Or currCntAlwnc8 > 1 Or currCntAlwnc9 > 1 Or currCntAlwnc10 > 1 Or currCntAlwnc11 > 1 Or currCntAlwnc12 > 1 Or currCntAlwnc13 > 1 Or currCntAlwnc14 > 1 Then
'                If currCntAlwnc1 > 1 Or currCntAlwnc2 > 1 Or currCntAlwnc3 > 1 Or currCntAlwnc4 > 1 Or currCntAlwnc5 > 1 Or currCntAlwnc6 > 1 Or currCntAlwnc7 > 1 Or currCntAlwnc8 > 1 Or currCntAlwnc9 > 1 Or currCntAlwnc10 > 1 Or currCntAlwnc11 > 1 Or currCntAlwnc12 > 1 Or currCntAlwnc13 > 1 Then
                
                
                
        If currCntAlwnc1 > 1 Or currCntAlwnc2 > 1 Or currCntAlwnc3 > 1 Or currCntAlwnc4 > 1 Or currCntAlwnc5 > 1 Or currCntAlwnc6 > 1 Or currCntAlwnc7 > 1 Or currCntAlwnc8 > 1 Or currCntAlwnc9 > 1 Or currCntAlwnc10 > 1 Or currCntAlwnc11 > 1 Or currCntAlwnc12 > 1 Or currCntAlwnc13 > 1 Or currCntAlwnc15 > 1 Or currCntAlwnc16 > 1 Then
                    fmsgboxStatus "Drop down has already been selected" ', vbOKOnly, "Error(s)!"
                    Target.value = "(Select)"
                End If

    If Range(Target.Address).value = "Sec 10(10)-Death-cum-retirement gratuity received " Then
'    If temp_Pensioner > 0 Or temp_Others > 0 Then
    If Sheet39.Range(Replace(Target.Address, "G", "J")).value > 2000000 Then
    fmsgboxStatus " For Sec 10(10)-Death-cum-retirement gratuity is selected than amount should not exceed .20 Lakhs "
    Sheet39.Range(Target.Address).value = "(Select)"
'    End If
    End If
    End If

    If Range(Target.Address).value = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
    If Sheet39.Range(Replace(Target.Address, "G", "J")).value > 500000 Then
    fmsgboxStatus "Exempt Allowance u/s 10(10B)-First Proviso-  Compensation limit notified by CG in the Official Gazette can not exceed Rs.500,000"
    Sheet39.Range(Target.Address).value = "(Select)"
    End If
    End If

    If Range(Target.Address).value = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
    If Sheet39.Range(Replace(Target.Address, "G", "J")).value > 500000 Then
    fmsgboxStatus "Exempt allowance u/s 10(10B) Second Proviso- Compensation under scheme approved by CG cannot exceed Rs. 5 lakhs."
    Sheet39.Range(Target.Address).value = "(Select)"
    End If
    End If

    If Mid(Range(Target.Address).value, 1, 11) = "Sec 10(13A)" Then

    If Sheet39.Range(Replace(Target.Address, "G", "J")).value > WorksheetFunction.Min(0.5 * Sheet19.Range("Temp_DA_BS").value, Sheet19.Range("Temp_HRA").value) Then
    fmsgboxStatus "Sec 10(13A)-Allowance to meet expenditure incurred on house rent- amount of exemption should be restricted to min of amount entered under  ""House Rent Allowance"" in sl no B1a or 50% of ""Basic+DA"" as under dropdowns of 1(a)"
    Sheet39.Range(Target.Address).value = "(Select)"
    End If
    
    'new change
         If (Range(Target.Address).value) = "Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee" Then
        If Sheet39.Range(Replace(Target.Address, "G", "J")) > 38400 Then
        fmsgboxStatus "Exempt allowance u/s For Section 10(14)(ii) -  Transport allowance granted to certain physically handicapped assessee is selected than amount should not exceed 38400"
        Sheet1.Range(Target.Address).value = ""
        Sheet1.Range(Target.Address).Select
        End If
        End If
    'new change
    'new change
    
'new change

   ' LOCK80GG
    Else
   'UNLOCK80GG



    End If

If Mid(Range(Target.Address).value, 1, 11) = "Sec 10(13A)" Then
                Sheet19.Unprotect Password:=sPassword
                Sheet19.Range("scvia.Section80GG").MergeArea.value = ""
                Sheet19.Range("scvia.Section80GG").Interior.Color = "&HD8D8D8" ', "AU"
                Sheet19.Range("scvia.Section80GG").MergeArea.Locked = True
                Sheet19.Protect Password:=sPassword
    Else
                Sheet19.Unprotect Password:=sPassword
                Sheet19.Range("scvia.Section80GG").MergeArea.value = ""
                Sheet19.Range("scvia.Section80GG").MergeArea.Interior.Color = "&HCCFFCC" ', "AU"
                Sheet19.Range("scvia.Section80GG").MergeArea.Locked = False
                Sheet19.Protect Password:=sPassword
    End If


        Dim ChangeAddressFrom As String
        Dim ChangeAddressTo As String
        ChangeAddressFrom = Replace(Target.Address, "G", "H")
'        ChangeAddressTo = Replace(Target.Address, "F", "H")
'        ChangeAddressFrom = ChangeAddressFrom & ":" & ChangeAddressTo

            If Mid(Sheet39.Range(Target.Address).value, 1, 1) = "A" Then
                Sheet39.Range(ChangeAddressFrom).Locked = False
                Sheet39.Range(ChangeAddressFrom).Interior.Color = "&HCCFFCC"
                Sheet39.Range(ChangeAddressFrom).ClearContents
            Else
                Sheet39.Range(ChangeAddressFrom).Locked = True
                Sheet39.Range(ChangeAddressFrom).Interior.Color = "&HD8D8D8"
                Sheet39.Range(ChangeAddressFrom).ClearContents
            End If
    End If
End If


If Target.Validation.Type = 6 Then
    If (Target.name.name) = "SAL.PANofEmployer" & count Then
        Dim TargetRepTAN As String
        TargetRepTAN = ""
        TargetRepTAN = Target.name.name

        sPassword = EfilingCommon.getmsgstate
         Sheet39.Unprotect Password:=sPassword
            If Not ValidateTANCodes(UCase(Trim(Range(TargetRepTAN).value))) Then
                fmsgboxStatus "Invalid TAN" & Chr(10) & Chr(13) & "TAN format should be First 4 Alphabets, next 5 digits, then 1 Alphabet.Please enter Valid TAN" ', vbOKCancel
                Sheet39.Range(TargetRepTAN).Select
                Sheet39.Range(TargetRepTAN).value = ""
            Else
                Sheet39.Range(TargetRepTAN).value = UCase(Range(TargetRepTAN).value)
            End If
         Sheet39.Protect Password:=sPassword
    End If
End If




If Target.Validation.Type = 1 Then
    Application.ScreenUpdating = False




    If Intersect(Range("SAL.DeductionUnder6Section1"), Range(Target.Address)) Is Nothing Then
        AddTotalSalary
    End If
    Application.ScreenUpdating = True
End If

count = Sheet39.Range("SalarySectionCount").value
For i = 1 To count
Sheet39.Range("SAL.PANofEmployer" & i).value = UCase(Sheet39.Range("SAL.PANofEmployer" & i).value)

        sPassword = EfilingCommon.getmsgstate
         Sheet39.Unprotect Password:=sPassword


    Dim nOfCount  As Long

    If Not Intersect(Range("Salary.Section" & i), Range(Target.Address)) Is Nothing Then
                nOfCount = CountDropdown1(Trim(Target.value), Sheet39.Range("Salary.Section" & i).Cells, Sheet39.Range("Salary.Section" & i).count)
                If currCnt1 > 1 Or currCnt2 > 1 Or currCnt3 > 1 Or currCnt4 > 1 Or currCnt5 > 1 Or currCnt6 > 1 Or currCnt7 > 1 Or currCnt8 > 1 Or currCnt9 > 1 Or currCnt10 > 1 Or currCnt11 > 1 Or currCnt12 > 1 Or currCnt13 > 1 Or currCnt14 > 1 Or currCnt15 > 1 Or currCnt16 > 1 Or currCnt17 > 1 Then
                    fmsgboxStatus "Drop down has already been selected" ', vbOKOnly, "Error(s)!"
                    Target.value = "(Select)"
        End If
    End If

    If Not Intersect(Range("Salary.ValueSection" & i), Range(Target.Address)) Is Nothing Then
                nOfCount = CountDropdown2(Trim(Target.value), Sheet39.Range("Salary.ValueSection" & i).Cells, Sheet39.Range("Salary.ValueSection" & i).count)
                If currCntvalue1 > 1 Or currCntvalue2 > 1 Or currCntvalue3 > 1 Or currCntvalue4 > 1 Or currCntvalue5 > 1 Or currCntvalue6 > 1 Or currCntvalue7 > 1 Or currCntvalue8 > 1 Or _
                          currCntvalue9 > 1 Or currCntvalue10 > 1 Or currCntvalue11 > 1 Or currCntvalue12 > 1 Or currCntvalue13 > 1 Or currCntvalue14 > 1 Or currCntvalue15 > 1 Or currCntvalue16 > 1 Or currCntvalue17 > 1 Then
                    fmsgboxStatus "Drop down has already been selected" ', vbOKOnly, "Error(s)!"
                    Target.value = "(Select)"
        End If
    End If

    If Not Intersect(Range("Salary.ProfitSection" & i), Range(Target.Address)) Is Nothing Then
                nOfCount = CountDropdown3(Trim(Target.value), Sheet39.Range("Salary.ProfitSection" & i).Cells, Sheet39.Range("Salary.ProfitSection" & i).count)
                If currCntProfit1 > 1 Or currCntProfit2 > 1 Or currCntProfit3 > 1 Then
                    fmsgboxStatus "Drop down has already been selected" ', vbOKOnly, "Error(s)!"
                    Target.value = "(Select)"
        End If
    End If

   If Not Intersect(Range("Salary.Section" & i), Range(Target.Address)) Is Nothing Then
'    Dim ChangeAddress As String
         ChangeAddress = Replace(Target.Address, "G", "H")
    If Sheet39.Range(Target.Address).value = "Others" Then
         Sheet39.Range(ChangeAddress).Locked = False
         Sheet39.Range(ChangeAddress).Interior.Color = "&HCCFFCC"
         Sheet39.Range(ChangeAddress).ClearContents
    Else
         Sheet39.Range(ChangeAddress).Locked = True
         Sheet39.Range(ChangeAddress).Interior.Color = "&HD8D8D8"
         Sheet39.Range(ChangeAddress).ClearContents
    End If
   End If

   If Not Intersect(Range("Salary.ValueSection" & i), Range(Target.Address)) Is Nothing Then
    'Dim ChangeAddress As String
         ChangeAddress = Replace(Target.Address, "G", "H")
    If Sheet39.Range(Target.Address).value = "Other benefits or amenities" Then
         Sheet39.Range(ChangeAddress).Locked = False
         Sheet39.Range(ChangeAddress).Interior.Color = "&HCCFFCC"
         Sheet39.Range(ChangeAddress).ClearContents
    Else
         Sheet39.Range(ChangeAddress).Locked = True
         Sheet39.Range(ChangeAddress).Interior.Color = "&HD8D8D8"
         Sheet39.Range(ChangeAddress).ClearContents
    End If
   End If


   If Not Intersect(Range("Salary.ProfitSection" & i), Range(Target.Address)) Is Nothing Then
    'Dim ChangeAddress As String
         ChangeAddress = Replace(Target.Address, "G", "H")
    If Sheet39.Range(Target.Address).value = "Any Other" Then
         Sheet39.Range(ChangeAddress).Locked = False
         Sheet39.Range(ChangeAddress).Interior.Color = "&HCCFFCC"
         Sheet39.Range(ChangeAddress).ClearContents
    Else
         Sheet39.Range(ChangeAddress).Locked = True
         Sheet39.Range(ChangeAddress).Interior.Color = "&HD8D8D8"
         Sheet39.Range(ChangeAddress).ClearContents
    End If
   End If

   If Not Intersect(Range("SAL.StateCode" & i), Range(Target.Address)) Is Nothing Then
   If Mid(Sheet39.Range("SAL.StateCode" & i).value, 1, 2) = "99" Then
         Sheet39.Range("SAL.ZipCode" & i).Locked = False
         Sheet39.Range("SAL.ZipCode" & i).Interior.Color = "&HCCFFCC"
         Sheet39.Range("SAL.ZipCode" & i).ClearContents
         Sheet39.Range("SAL.PinCode" & i).Locked = True
         Sheet39.Range("SAL.PinCode" & i).Interior.Color = "&HD8D8D8"
         Sheet39.Range("SAL.PinCode" & i).ClearContents
   Else
         Sheet39.Range("SAL.PinCode" & i).Locked = False
         Sheet39.Range("SAL.PinCode" & i).Interior.Color = "&HCCFFCC"
         Sheet39.Range("SAL.PinCode" & i).ClearContents
         Sheet39.Range("SAL.ZipCode" & i).Locked = True
         Sheet39.Range("SAL.ZipCode" & i).Interior.Color = "&HD8D8D8"
         Sheet39.Range("SAL.ZipCode" & i).ClearContents
   End If
   End If

'   If Not Intersect(Range("sheet.EmployerCategory" & i), Range(Target.address)) Is Nothing Then
'   If Sheet39.Range("sheet.EmployerCategory" & i).value = "Pensioners" Or Mid(Sheet39.Range("sheet.EmployerCategory" & i).value, 1, 3) = "OTH" Then
'         Sheet39.Range("SAL.Entertainment_allowance_1").ClearContents
'         Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HD8D8D8"
'         Sheet39.Range("SAL.Entertainment_allowance_1").Locked = True
'
'   Else
'         'Sheet39.Range("SAL.Entertainment_allowance_" & i).ClearContents
'         Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HCCFFCC"
'         Sheet39.Range("SAL.Entertainment_allowance_1").Locked = False
'
'   End If
'
'   If Sheet39.Range("sheet.EmployerCategory" & i).value = "Pensioners" Then
'      Sheet39.Range("SAL.DeductionUnderSection1").ClearContents
'      Sheet39.Range("SAL.DeductionUnderSection1").Locked = True
'      Sheet39.Range("SAL.DeductionUnderSection1").Interior.Color = "&HD8D8D8"
'
'Else
'      'Sheet39.Range("SAL.Entertainment_allowance_" & i).ClearContents
'      Sheet39.Range("SAL.DeductionUnderSection1").Interior.Color = "&HCCFFCC"
'      Sheet39.Range("SAL.DeductionUnderSection1").Locked = False
'   End If
'
'End If
Next

  If Target.Validation.Type = 1 Then
    Application.ScreenUpdating = False
    If Intersect(Range("SAL.DeductionUnder6Section1"), Range(Target.Address)) Is Nothing Then
        AddTotalSalary

    End If
    Application.ScreenUpdating = True
End If

   Sheet39.Protect Password:=sPassword










'If Target.name.name = "SAL.DeductionUnder6Section1" Then
             If (Range("SAL.DeductionUnder6Section1").value) > 50000 Then
            fmsgboxStatus "Deduction u/s. 16(ia) cannot exceed Rs. 50000 or Net Salary whichever is lower in Schedule Salary."
            Range("SAL.DeductionUnder6Section1").value = 50000
            Range("SAL.DeductionUnder6Section1").Select
             End If

             Dim temp16a, temp16aa As Variant
            temp16a = Sheet39.Range("SAL.DeductionUnder6Section1").value
            temp16aa = Sheet39.Range("SAL.Netsalary1").value


        If temp16a <> "" Then
        If Not Sheet39.Range("SAL.DeductionUnder6Section1").value <= temp16aa Then
               fmsgboxStatus "Deduction u/s. 16(ia) cannot exceed Rs. 50000 or Net Salary whichever is lower in Schedule Salary"
               Range("SAL.DeductionUnder6Section1").value = temp16aa
               Range("SAL.DeductionUnder6Section1").Select


        End If
        End If

'End If

 If Target.name.name = "SAL.Entertainment_allowance_1" Or Target.name.name = "SAL.Netsalary1" Then
   Dim temp1 As Variant
       temp1 = 0
       temp1 = Sheet39.Range("Temp_allowance").value / 5

       If Range(Target.Address).value > WorksheetFunction.Min(temp1, 5000) Then
         fmsgboxStatus "Deduction of Entertainment allowance u/s 16(ii) Maximum amount allowable in this field shall be 1/5th of Basic Salary 1a of all employers or Rs 5000 whichever is lower."
         Range("SAL.Entertainment_allowance_1").value = ""
         Range("SAL.Entertainment_allowance_1").Select
       End If
   End If

   If Target.name.name = "SAL.DeductionUnderSection1" Then
   If Range("SAL.DeductionUnderSection1").value > 5000 Then
    fmsgboxStatus "Deduction of Professional tax u/s 16(iii) is Rs. 200/- p.m. in most of the state. Kindly reconsider the amount claimed"
'    Range("SAL.DeductionUnderSection1").value = ""
'    Range("SAL.DeductionUnderSection1").Select
   End If
   End If




endline:
Application.EnableEvents = True
End Sub
Function GetTotal89A(cnt As Variant) As Long
    Dim i As Variant
    Dim amt As Long
    amt = 0
    GetTotal89A = 0
    For i = 1 To cnt
    amt = amt + Sheet39.Range("IncomeNotified89A_AmountUS" & i).value
    amt = amt + Sheet39.Range("IncomeNotified89A_AmountUK" & i).value
    amt = amt + Sheet39.Range("IncomeNotified89A_AmountCan" & i).value
    Next
    GetTotal89A = amt
End Function

Private Sub Worksheet_Calculate()

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error Resume Next
CalculateTaxCheckFlag = False
Dim count, i As Variant
Application.EnableEvents = False

 'sai 12/09/2024----------------
            Dim PinCode As Range
            Dim DupFlag As Boolean
            Dim StateName
            Dim StateCode
            Dim state1 As Variant

'=====================================================
'Jyoti HRA 28/5/2025
''--Bindu 27/05/2025
sPassword = EfilingCommon.getmsgstate
         Sheet39.Unprotect Password:=sPassword
         
Dim BA_SAL, DA_SAL, HouseRentAllowance_count, count1
 BA_SAL = 0
 DA_SAL = 0
HouseRentAllowance_count = 0

 Totsections = Sheet39.Range("SalarySectionCount").value

 For count1 = 1 To Totsections
       BA_SAL = BA_SAL + WorksheetFunction.SumIf(Sheet39.Range("Salary.Section" & count1), "Basic Salary", Sheet39.Range("Salary.SalaryAmount" & count1))

     DA_SAL = DA_SAL + WorksheetFunction.SumIf(Sheet39.Range("Salary.Section" & count1), "Dearness Allowance(DA)", Sheet39.Range("Salary.SalaryAmount" & count1))
 
  HouseRentAllowance_count = HouseRentAllowance_count + WorksheetFunction.SumIf(Sheet39.Range("Salary.Section" & count1), "House Rent Allowance(HRA)", Sheet39.Range("Salary.SalaryAmount" & count1))

    
Next

Sheet39.Range("SAL_Basic").value = BA_SAL
Sheet39.Range("SAL_DA").value = DA_SAL
Sheet39.Range("HRA_SAL").value = HouseRentAllowance_count



''---End Bindu
'End Jyoti 28/5/2025

count = Sheet39.Range("SalarySectionCount").value
For i = 1 To count
Sheet39.Range("SAL.PANofEmployer" & i).value = UCase(Sheet39.Range("SAL.PANofEmployer" & i).value)
        sPassword = EfilingCommon.getmsgstate
         Sheet39.Unprotect Password:=sPassword
    Dim nOfCount, inInc89A As Long
   
    'Newlyy added by Bindu as its Legacy issue
    Dim leavecount

    If i > 1 Then
      If leavecount = 0 Then
           leavecount = CountDropdown_Leave(Trim(Target.value), Sheet39.Range("Salary.Section" & i).Cells, Sheet39.Range("Salary.Section" & i).count)
      End If
    Else
          leavecount = CountDropdown_Leave(Trim(Target.value), Sheet39.Range("Salary.Section" & i).Cells, Sheet39.Range("Salary.Section" & i).count)
    End If

    
    If Not Intersect(Range("IncomeNotified89A_AmountUS" & i), Range(Target.Address)) Is Nothing Then
        inInc89A = GetTotal89A(count)
        If Sheet39.Range("Increliefus89A").value > inInc89A Then
        fmsgbox "* Income claimed for relief from taxation u/s 89A  cannot be claimed more than income entered in income from retirement benefit account maintained in a notified country u/s 89A"
        Sheet39.Range("Increliefus89A").value = 0
        End If
        calSchS
    End If
    If Not Intersect(Range("IncomeNotified89A_AmountUK" & i), Range(Target.Address)) Is Nothing Then
        inInc89A = GetTotal89A(count)
        If Sheet39.Range("Increliefus89A").value > inInc89A Then
        fmsgbox "* Income claimed for relief from taxation u/s 89A  cannot be claimed more than income entered in income from retirement benefit account maintained in a notified country u/s 89A"
        Sheet39.Range("Increliefus89A").value = 0
        End If
        calSchS
    End If
    If Not Intersect(Range("IncomeNotified89A_AmountCan" & i), Range(Target.Address)) Is Nothing Then
        inInc89A = GetTotal89A(count)
        If Sheet39.Range("Increliefus89A").value > inInc89A Then
        fmsgbox "* Income claimed for relief from taxation u/s 89A  cannot be claimed more than income entered in income from retirement benefit account maintained in a notified country u/s 89A"
        Sheet39.Range("Increliefus89A").value = 0
        End If
        calSchS
    End If
    
    If Not Intersect(Range("Increliefus89A"), Range(Target.Address)) Is Nothing Then
        inInc89A = GetTotal89A(count)
        If Sheet39.Range("Increliefus89A").value > inInc89A Then
        fmsgbox "* Income claimed for relief from taxation u/s 89A  cannot be claimed more than income entered in income from retirement benefit account maintained in a notified country u/s 89A"
        Sheet39.Range("Increliefus89A").value = 0
        End If
        If Range("Increliefus89A").value > 0 Then
        fmsgbox "* Please ensure to submit Form 10EE."
    End If
    End If
    
    If Not Intersect(Range("SAL.PANofEmployer" & i), Range(Target.Address)) Is Nothing Then
        Dim TargetRepTAN As String
        TargetRepTAN = ""
        TargetRepTAN = Target.name.name
        sPassword = EfilingCommon.getmsgstate
         Sheet39.Unprotect Password:=sPassword
            If Not ValidateTANCodes(UCase(Trim(Range(TargetRepTAN).value))) Then
            If Trim(Target.value) <> "" Then
                fmsgboxStatus "Invalid TAN" & Chr(10) & Chr(13) & "TAN format should be First 4 Alphabets, next 5 digits, then 1 Alphabet.Please enter Valid TAN" ', vbOKCancel
                End If
                Sheet39.Range(TargetRepTAN).Select
                Sheet39.Range(TargetRepTAN).value = ""
            Else
                Sheet39.Range(TargetRepTAN).value = UCase(Range(TargetRepTAN).value)
            End If
         Sheet39.Protect Password:=sPassword
    End If
    If Not Intersect(Range("sheet.EmployerCategory" & i), Range(Target.Address)) Is Nothing Then
        calSchS
    End If
    
'Change-28.01.2023.101.S.3A
    If Not Intersect(Range("sheet.EmployerCategory" & i), Range(Target.Address)) Is Nothing Then
'Change-04.05.2023.101. -> V0.4 Implementation
        Sheet39.Range("Salary.Section" & i).ClearContents
        Application.EnableEvents = True
        Application.EnableEvents = False
        Sheet39.Range("Salary.SalaryAmount" & i).ClearContents
        Application.EnableEvents = False
'        Sheet39.Unprotect Password:=sPassword
'        Sheet39.Range("SAL.AllowancesNotExempt" & i).value = 0
'        Sheet39.Range("SAL.Salary" & i).value = 0
'        Sheet39.Protect Password:=sPassword
'----
        SalaryExemptAllowanceDropdown
    End If

    If Not Intersect(Range("Salary.SalaryAmount" & i), Range(Target.Address)) Is Nothing Then
        
        calSchS
    End If
    If Not Intersect(Range("Salary.Section" & i), Range(Target.Address)) Is Nothing Then
        nOfCount = CountDropdown1(Trim(Target.value), Sheet39.Range("Salary.Section" & i).Cells, Sheet39.Range("Salary.Section" & i).count)
        If currCnt1 > 1 Or currCnt2 > 1 Or currCnt3 > 1 Or currCnt4 > 1 Or currCnt5 > 1 Or currCnt6 > 1 Or currCnt7 > 1 Or currCnt8 > 1 Or currCnt9 > 1 Or currCnt10 > 1 Or currCnt11 > 1 Or currCnt12 > 1 Or currCnt13 > 1 Or currCnt14 > 1 Or currCnt15 > 1 Or currCnt16 > 1 Or currCnt17 > 1 Then
            fmsgboxStatus "Drop down has already been selected" ', vbOKOnly, "Error(s)!"
            Target.value = "(Select)"
        End If
                
''Change-04.05.2023.101. -> V0.4 Implementation
'        If (Target.value = "Contribution made by the central government towards Agnipath scheme as referred  under section 80CCH") Then
'            If Sheet39.Range("sheet.EmployerCategory" & i).value <> "Central Government" Or Dformat(Sheet1.Range("sheet1.DOB"), "") < "2001-03-31" Or Dformat(Sheet1.Range("sheet1.DOB"), "") > "2006-03-31" Then
'                fmsgboxStatus "Contribution made by the central government towards Agnipath scheme as referred  under section 80CCH can only be selected, if user has enrolled into the Agnipath Scheme"
'                Target.value = "(Select)"
'            End If
'        End If

'        If (Target.value = "Contribution made by the central government towards Agnipath scheme as referred  under section 80CCH") Then
'            'If Sheet39.Range("sheet.EmployerCategory" & i).value <> "Central Government" Or Dformat(Sheet1.Range("sheet1.DOB"), "") < "2001-03-31" Or Dformat(Sheet1.Range("sheet1.DOB"), "") > "2006-03-31" Then
'                If Range("sheet.EmployerCategory" & i) <> "Central Government" Then 'Or (Dformat((Sheet1.Range("sheet1.DOB").value), "") < "2001-03-31" Or Dformat((Sheet1.Range("sheet1.DOB").value), "") > "2006-04-01") Then
'                fmsgboxStatus "Contribution made by the central government towards Agnipath scheme as referred  under section 80CCH can only be selected, if user has enrolled into the Agnipath Scheme"
'                Target.value = ""
'            End If
'        End If



       'Newly added by Sai for 80CCH Validation on 21/03/2024

        If Target.value = "Contribution made by the central government towards Agnipath scheme as referred  under section 80CCH" Then
            'If Range("sheet.EmployerCategory" & i) <> "Central Government" Or (Dformat((Sheet1.Range("sheet1.DOB").value), "") < "2001-03-31" Or Dformat((Sheet1.Range("sheet1.DOB").value), "") > "2007-04-01") Then
               ' AY24-25 Change
               'If Range("sheet.EmployerCategory" & i) <> "Central Government" Then 'Or (Dformat((Sheet1.Range("sheet1.DOB").value), "") < "2001-03-31" Or Dformat((Sheet1.Range("sheet1.DOB").value), "") > "2006-04-01") Then
'        If Range("sheet.EmployerCategory" & i) <> "Central Government" Or (Dformat((Sheet1.Range("sheet1.DOB").value), "") < "1998-03-31" Or Dformat((Sheet1.Range("sheet1.DOB").value), "") > "2008-04-01") Then 'added
        If Range("sheet.EmployerCategory" & i) <> "Central Government" Or (Dformat((Sheet1.Range("sheet1.DOB").value), "") < "1999-03-31" Or Dformat((Sheet1.Range("sheet1.DOB").value), "") > "2009-04-01") Then 'added by Aavula
                fmsgboxStatus ("Contribution made by the central government towards Agnipath scheme as referred  under section 80CCH can only be selected, if user has enrolled into the Agnipath Scheme")
                Target.value = ""
            End If
        End If
'----
        calSchS
    End If
    If Not Intersect(Range("Salary.ValueSection" & i), Range(Target.Address)) Is Nothing Then
                nOfCount = CountDropdown2(Trim(Target.value), Sheet39.Range("Salary.ValueSection" & i).Cells, Sheet39.Range("Salary.ValueSection" & i).count)
                If currCntvalue1 > 1 Or currCntvalue2 > 1 Or currCntvalue3 > 1 Or currCntvalue4 > 1 Or currCntvalue5 > 1 Or currCntvalue6 > 1 Or currCntvalue7 > 1 Or currCntvalue8 > 1 Or _
                          currCntvalue9 > 1 Or currCntvalue10 > 1 Or currCntvalue11 > 1 Or currCntvalue12 > 1 Or currCntvalue13 > 1 Or currCntvalue14 > 1 Or currCntvalue15 > 1 Or currCntvalue16 > 1 Or currCntvalue17 > 1 _
                          Or currCntvalue18 > 1 Or currCntvalue19 > 1 Or currCntvalue20 > 1 Or currCntvalue21 > 1 Then
                    fmsgboxStatus "Drop down has already been selected" ', vbOKOnly, "Error(s)!"
                    Target.value = "(Select)"
        End If
    End If
    
    
    'Newly added by Bindu as per Legacy issue
     'leavecount
  If Not Intersect(Range("Salary.Nature1"), Range(Target.Address)) Is Nothing Then
'    If target.value = "Sec 10(5)-Leave Travel concession/assistance" Or target.value = "Sec 10(5)- Travel concession/assistance received" Then
    If Target.value = "Sec 10(5)- Travel concession/assistance received" Then 'Added by Aavula Naresh on 16/05/2025
       
       If leavecount = 0 Then
'            fmsgboxStatus ("Sec 10(5)-Leave Travel concession/assistance received can only be selected, if user has selected Leave Travel Allowance(LTA)")
            fmsgboxStatus ("Sec 10(5)- Travel concession/assistance received can only be selected, if user has selected Leave Travel Allowance(LTA)")
            Target.value = ""
       End If
    End If
  End If
  
  If Not Intersect(Range("Salary.Amount1"), Range(Target.Address)) Is Nothing Then
'    If Range(Replace(target.Address, "J", "G")).value = "Sec 10(5)-Leave Travel concession/assistance" Then
    If Range(Replace(Target.Address, "J", "G")).value = "Sec 10(5)- Travel concession/assistance received" Then 'Added by Aavula Naresh on 16/05/2025
       If leavecount = 0 Then
'            fmsgboxStatus ("Sec 10(5)-Leave Travel concession/assistance can only be selected, if user has selected Leave Travel Allowance(LTA)")
            fmsgboxStatus ("Sec 10(5)- Travel concession/assistance received can only be selected, if user has selected Leave Travel Allowance(LTA)")
            Target.value = ""
            Range(Replace(Target.Address, "J", "G")).value = ""
       End If
    End If
  End If

    If Not Intersect(Range("Salary.ProfitSection" & i), Range(Target.Address)) Is Nothing Then
                nOfCount = CountDropdown3(Trim(Target.value), Sheet39.Range("Salary.ProfitSection" & i).Cells, Sheet39.Range("Salary.ProfitSection" & i).count)
                If currCntProfit1 > 1 Or currCntProfit2 > 1 Or currCntProfit3 > 1 Then
                    fmsgboxStatus "Drop down has already been selected" ', vbOKOnly, "Error(s)!"
                    Target.value = "(Select)"
        End If
    End If

   If Not Intersect(Range("Salary.Section" & i), Range(Target.Address)) Is Nothing Then
         ChangeAddress = Replace(Target.Address, "G", "H")
    If Sheet39.Range(Target.Address).value = "Others" Then
'Change-7, ANK-152, 25.05.2022
        Sheet39.Unprotect Password:=sPassword
'---
         Sheet39.Range(ChangeAddress).Locked = False
         Sheet39.Range(ChangeAddress).Interior.Color = "&HCCFFCC"
         Sheet39.Range(ChangeAddress).ClearContents
    Else
'Change-8, ANK-152, 25.05.2022
        Sheet39.Unprotect Password:=sPassword
'---
         Sheet39.Range(ChangeAddress).Locked = True
         Sheet39.Range(ChangeAddress).Interior.Color = "&HD8D8D8"
         Sheet39.Range(ChangeAddress).ClearContents
    End If
   End If

   If Not Intersect(Range("Salary.ValueSection" & i), Range(Target.Address)) Is Nothing Then
         ChangeAddress = Replace(Target.Address, "G", "H")
    If Sheet39.Range(Target.Address).value = "Other benefits or amenities" Then
         Sheet39.Range(ChangeAddress).Locked = False
         Sheet39.Range(ChangeAddress).Interior.Color = "&HCCFFCC"
         Sheet39.Range(ChangeAddress).ClearContents
    Else
         Sheet39.Range(ChangeAddress).Locked = True
         Sheet39.Range(ChangeAddress).Interior.Color = "&HD8D8D8"
         Sheet39.Range(ChangeAddress).ClearContents
    End If
   End If


   If Not Intersect(Range("Salary.ProfitSection" & i), Range(Target.Address)) Is Nothing Then
         ChangeAddress = Replace(Target.Address, "G", "H")
    If Sheet39.Range(Target.Address).value = "Any Other" Then
         Sheet39.Range(ChangeAddress).Locked = False
         Sheet39.Range(ChangeAddress).Interior.Color = "&HCCFFCC"
         Sheet39.Range(ChangeAddress).ClearContents
    Else
         Sheet39.Range(ChangeAddress).Locked = True
         Sheet39.Range(ChangeAddress).Interior.Color = "&HD8D8D8"
         Sheet39.Range(ChangeAddress).ClearContents
    End If
   End If

   If Not Intersect(Range("SAL.StateCode" & i), Range(Target.Address)) Is Nothing Then
   If Mid(Sheet39.Range("SAL.StateCode" & i).value, 1, 2) = "99" Then
         Sheet39.Range("SAL.ZipCode" & i).Locked = False
         Sheet39.Range("SAL.ZipCode" & i).Interior.Color = "&HCCFFCC"
         Sheet39.Range("SAL.ZipCode" & i).ClearContents
         Sheet39.Range("SAL.PinCode" & i).Locked = True
         Sheet39.Range("SAL.PinCode" & i).Interior.Color = "&HD8D8D8"
         Sheet39.Range("SAL.PinCode" & i).ClearContents
   Else
         Sheet39.Range("SAL.PinCode" & i).Locked = False
         Sheet39.Range("SAL.PinCode" & i).Interior.Color = "&HCCFFCC"
         Sheet39.Range("SAL.PinCode" & i).ClearContents
         Sheet39.Range("SAL.ZipCode" & i).Locked = True
         Sheet39.Range("SAL.ZipCode" & i).Interior.Color = "&HD8D8D8"
         Sheet39.Range("SAL.ZipCode" & i).ClearContents
   End If

   End If
    
'Change 01, CDB - , 30.06.2022
'Commented by sai 12/09/2024======================================

'        If Not Intersect(Range("SAL.PinCode" & i), Range(Target.Address)) Is Nothing Then
'
'           Dim PinCode As Range
'           Dim found As Variant
'            If (Target.value <> "") Then
'           found = False
'
'           For Each PinCode In SheetALL.Range("PincodeList")
'               Dim code As String
'               code = Left(CStr(Target.value), 3)
'               If CStr(PinCode.value) = CStr(code) Then
'                   Range("SAL.StateCode" & i).value = SheetALL.Range("CX" & PinCode.row).value
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
            
'---End Change========================================================================================

'by sai on 17/09/2024
If Not Application.Intersect(Sheet39.Range("SAL.StateCode" & i), Sheet39.Range(Target.Address)) Is Nothing Then
     If Sheet39.Range(Replace(Target.Address, "H", "K")).value <> "" Then

     Dim PIN_targetadd, state_targetadd As String
      state_targetadd = Replace(Target.Address, "$", "")
      PIN_targetadd = Replace(state_targetadd, "H", "K")
    
        If state_Validation(PIN_targetadd, state_targetadd) = False Then
        Sheet39.Range(Replace(state_targetadd, "H", "K")).value = ""
        End If
     End If
End If


If Not Application.Intersect(Sheet39.Range("SAL.PinCode" & i), Sheet39.Range(Target.Address)) Is Nothing Then


     Dim pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk As String
     Dim ws As Worksheet
     
      pinrng_chk = Replace(Target.Address, "$", "")
      staterng_chk = Replace(pinrng_chk, "K", "H")
      Set ws = Worksheets("Schedule S")
      
      Call PinState_codes_validation(ws, pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk)

End If

Next


'''''''''''




'dpk1s
Dim temp_Pensioner, temp_PensionerCG, temp_PensionerSG, temp_PensionerPSU, temp_PensionerOth As Variant
Dim temp_Others, temp As Variant
Dim Salcount1 As Long
Salcount1 = 0
temp_Pensioner = 0
temp_PensionerCG = 0
temp_PensionerSG = 0
temp_PensionerPSU = 0
temp_PensionerOth = 0
temp_Others = 0
temp = 0

For Salcount1 = 1 To Sheet39.Range("SalarySectionCount").value
     If Not isdropdownblank(Sheet39.Range("sheet.EmployerCategory" & Salcount1).value) Then
     temp = temp + 1
     End If
    
     If Sheet39.Range("sheet.EmployerCategory" & Salcount1).value = "CG-Pensioners" Then
     temp_PensionerCG = temp_PensionerCG + 1
     End If
     
     If Sheet39.Range("sheet.EmployerCategory" & Salcount1).value = "SG-Pensioners" Then
     temp_PensionerSG = temp_PensionerSG + 1
     End If
     
     If Sheet39.Range("sheet.EmployerCategory" & Salcount1).value = "PSU-Pensioners" Then
     temp_PensionerPSU = temp_PensionerPSU + 1
     End If
     
     If Sheet39.Range("sheet.EmployerCategory" & Salcount1).value = "Others-Pensioners" Then
     temp_PensionerOth = temp_PensionerOth + 1
     End If
        
'Change-20.04.2023.101.S -> SIT-48785
     If Mid(Sheet39.Range("sheet.EmployerCategory" & Salcount1).value, 1, 1) = "O" Then
     temp_Others = temp_Others + 1
     End If
'    If Sheet39.Range("sheet.EmployerCategory" & Salcount1).value = "OTHERS" Then
'     temp_Others = temp_Others + 1
'     End If
'----
Next

'Change-23.05.2023.101.S -> Issue number 1 Production build 1
'If temp <> 0 And (temp = temp_PensionerCG Or temp = temp_PensionerSG Or temp = temp_PensionerPSU Or temp = temp_PensionerOth Or temp = temp_Others Or SheetALL.Range("bacValue").value = 1) Then
'    Sheet39.Unprotect Password:=getmsgstate
'    Sheet39.Range("SAL.Entertainment_allowance_1").value = ""
'    Sheet39.Range("SAL.Entertainment_allowance_1").Locked = True
'    Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HD8D8D8"
'    Sheet39.Protect Password:=getmsgstate
'Else
'    Sheet39.Unprotect Password:=getmsgstate
'    Sheet39.Range("SAL.Entertainment_allowance_1").Locked = False
'    Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HCCFFCC"
'    Sheet39.Protect Password:=getmsgstate
'End If
'
'If temp = (temp_Pensioner + temp_Others) And temp_Pensioner > 0 And temp_Others > 0 Then
'    Sheet39.Unprotect Password:=getmsgstate
'    Sheet39.Range("SAL.Entertainment_allowance_1").value = ""
'    Sheet39.Range("SAL.Entertainment_allowance_1").Locked = True
'    Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HD8D8D8"
'    Sheet39.Protect Password:=getmsgstate
'End If
'
'
'If (SheetALL.Range("bacValue").value) = 1 Then
'    Sheet39.Unprotect Password:=getmsgstate
'    Sheet39.Range("SAL.Entertainment_allowance_1").value = ""
'    Sheet39.Range("SAL.Entertainment_allowance_1").Locked = True
'    Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HD8D8D8"
'    Sheet39.Protect Password:=getmsgstate
'End If

''''''''''''




Dim Temp_CG, Temp_SG, Temp_PSU
Temp_CG = 0
Temp_SG = 0
Temp_PSU = 0

For Salcount1 = 1 To Sheet39.Range("SalarySectionCount").value

If Sheet39.Range("sheet.EmployerCategory" & Salcount1).value = "Central Government" Then
     Temp_CG = Temp_CG + 1
End If

If Sheet39.Range("sheet.EmployerCategory" & Salcount1).value = "State Government" Then
     Temp_SG = Temp_SG + 1
End If

If Sheet39.Range("sheet.EmployerCategory" & Salcount1).value = "Public Sector Undertaking" Then
     Temp_PSU = Temp_PSU + 1
End If

Next Salcount1

If Temp_CG > 0 Or Temp_SG > 0 Or Temp_PSU > 0 Then
    Sheet39.Unprotect Password:=getmsgstate
    Sheet39.Range("SAL.Entertainment_allowance_1").Locked = False
    Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HCCFFCC"
    Sheet39.Protect Password:=getmsgstate
Else
    Sheet39.Unprotect Password:=getmsgstate
    Sheet39.Range("SAL.Entertainment_allowance_1").value = ""
    Sheet39.Range("SAL.Entertainment_allowance_1").Locked = True
    Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HD8D8D8"
    Sheet39.Protect Password:=getmsgstate
End If

    'Modified by Aavula Naresh SIT-69502
'If Sheet1.Range("OptingNewTaxRegime").value = "Continue to opt" Or Sheet1.Range("OptingNewTaxRegime").value = "Opting in now" Or Sheet1.Range("sheet1.Status").value = "H-HUF" Then
    'Commented by Aavula Naresh
'  If Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "No" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime").value = "Yes-but beyond due date" Or Sheet1.Range("sheet1.Status").value = "H-HUF" Then
    
    'Modified by Aavula Naresh for AY 2025-26
    
'If Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_No").value = "No" Or Sheet1.Range("Sheet1.OptOutNewTaxRegime_New_NA").value = "No" Or Sheet1.Range("Sheet1.OptOutNewTaxRegimeSetB").value = "No" Or Sheet1.Range("sheet1.Status").value = "H-HUF" Then

'Updated by sadineni on 22/05/2026 as Naresh forgot to update it in AY 2026-27
'Newly modified by sadineni on 01/06/2026
'If Sheet1.Range("Sheet1.Form10IEAoldtax_IA").value = "Yes" Or Sheet1.Range("Sheet1.Form10IEANewTax").value = "No" Or Sheet1.Range("Sheet1.Form10IEAB").value = "No" Or Sheet1.Range("Sheet1.Form10IEANewTaxIB").value = "No" Or Sheet1.Range("sheet1.Status").value = "H-HUF" Then
If Sheet1.Range("Sheet1.Form10IEAoldtax_IA").value = "Yes" Or Sheet1.Range("Sheet1.Form10IEANewTax").value = "Yes" Or Sheet1.Range("Sheet1.Form10IEAB").value = "No" Or Sheet1.Range("Sheet1.Form10IEANewTaxIB").value = "No" Or Sheet1.Range("sheet1.Status").value = "H-HUF" Then
    Sheet39.Unprotect Password:=getmsgstate
    Sheet39.Range("SAL.Entertainment_allowance_1").value = ""
    Sheet39.Range("SAL.Entertainment_allowance_1").Locked = True
    Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HD8D8D8"
    Sheet39.Protect Password:=getmsgstate
    
End If

'---------
If temp <> 0 And (temp = temp_Pensioner Or SheetALL.Range("bacValue").value = 1) Then
    Sheet39.Unprotect Password:=getmsgstate
    Sheet39.Range("SAL.DeductionUnderSection1").value = ""
    Sheet39.Range("SAL.DeductionUnderSection1").Locked = True
    Sheet39.Range("SAL.DeductionUnderSection1").Interior.Color = "&HD8D8D8"
    Sheet39.Protect Password:=getmsgstate
Else
    If (Sheet1.Range("sheet1.Status").value <> "H-HUF") Then
        Sheet39.Unprotect Password:=getmsgstate
        Sheet39.Range("SAL.DeductionUnderSection1").Locked = False
        Sheet39.Range("SAL.DeductionUnderSection1").Interior.Color = "&HCCFFCC"
        Sheet39.Protect Password:=getmsgstate
    End If
End If

'dpk1e
'Jyoti 26_06_2025
If Not Intersect(Range("Salary.Amount1"), Range(Target.Address)) Is Nothing Then
'    If Range(target.Address).value > 2000000 Then '= "Sec 10(10)-Death-cum-retirement gratuity received " Then
'        If Trim(Sheet39.Range(Replace(target.Address, "J", "G")).value) = Trim("Sec 10(10)-Death-cum-retirement gratuity received ") Then
'        fmsgboxStatus " For Sec 10(10)-Death-cum-retirement gratuity is selected than amount should not exceed Rs.20 Lakhs "
'        Sheet39.Range(target.Address).value = ""
'        End If
'   calSchS
'    End If

    If Range(Target.Address).value > 500000 Then '= "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
        If Sheet39.Range(Replace(Target.Address, "J", "G")).value = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
        fmsgboxStatus "Exempt Allowance u/s 10(10B)-First Proviso-  Compensation limit notified by CG in the Official Gazette can not exceed Rs.500,000"
        Sheet39.Range(Target.Address).value = ""
        End If
    End If
    'Change-02.11.2023.102.Ver.0.5
    If Range(Target.Address).value > 500000 Then '= "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
        If Sheet39.Range(Replace(Target.Address, "J", "G")).value = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
        fmsgboxStatus "Exempt Allowance u/s 10(10B)-Second proviso- Compensation under scheme approved by the Central Government can not exceed Rs.500,000"
        Sheet39.Range(Target.Address).value = ""
        End If
    End If
    'End Change
    If Range(Target.Address).value > 38400 Then '= "Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee" Then
        If Sheet39.Range(Replace(Target.Address, "J", "G")).value = "Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee" Then
        fmsgboxStatus "Exempt Allowance u/s Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee can not exceed Rs.38400"
        Sheet39.Range(Target.Address).value = ""
        End If
    End If
'Change 102 28.11.2023 Ver0.6 SIT-62868 B
    If Range(Target.Address).value > 500000 Then '= "Sec 10(10C)-Amount received/receivable on voluntary retirement or termination of service" Then
        'If Sheet39.Range(Replace(Target.Address, "J", "G")).value = "Sec 10(10C)-Amount received/receivable on voluntary retirement or termination of service" Then
         If Sheet39.Range(Replace(Target.Address, "J", "G")).value = "Sec 10(10C)- Amount received on voluntary retirement or termination of service" Then
        fmsgboxStatus "Kindly restrict the exemption u/s 10(10C) to Rs. 5,00,000"
        Sheet39.Range(Target.Address).value = ""
        End If
    End If
'End Change
End If

If Not Intersect(Range("Salary.Amount1"), Range(Target.Address)) Is Nothing Then
    If Mid(Sheet39.Range(Replace(Target.Address, "J", "G")), 1, 11) = "Sec 10(13A)" Then
    If Range(Target.Address).value > WorksheetFunction.Min(0.5 * Sheet19.Range("Temp_DA_BS").value, Sheet19.Range("Temp_HRA").value) Then
    fmsgboxStatus "Sec 10(13A)-Allowance to meet expenditure incurred on house rent- amount of exemption should be restricted to min of amount entered under  ""House Rent Allowance"" in sl no B1a or 50% of ""Basic+DA"" as under dropdowns of 1(a)"
    Range(Target.Address).value = ""
    End If
    End If
   calSchS
End If



     count = Sheet39.Range("SalarySectionCount").value
For i = 1 To count
    If UCase(Sheet39.Range("sheet.EmployerCategory" & i).value = "OTHERS") Or Sheet39.Range("sheet.EmployerCategory" & i).value = "Others-Pensioners" Then
    
        If Not Application.Intersect(Sheet39.Range("Salary.Amount1"), Sheet39.Range(Target.Address)) Is Nothing Then
        
        'Newly added by Sai
         If Sheet39.Range(Target.Address).value > 2500000 Then
                If Sheet39.Range(Replace(Target.Address, "J", "G")).value = "Sec 10(10AA)-Earned leave encashment on Retirement" Then
                          fmsgboxStatus "Deduction u/s 10(10AA) shall be restricted to Rs. 25 Lakh in case of  employer catergory is selected as other than Central or State Government"
                          Target.value = ""
                 End If
             
          End If
        End If
        
        
    ElseIf UCase(Sheet39.Range("sheet.EmployerCategory" & i).value = "PSU-Pensioners") Or Sheet39.Range("sheet.EmployerCategory" & i).value = "Public Sector Undertaking" Then
        If Not Application.Intersect(Sheet39.Range("Salary.Amount1"), Sheet39.Range(Target.Address)) Is Nothing Then
            If Target.value > 2500000 Then
                If Sheet39.Range(Replace(Target.Address, "J", "G")).value = "Sec 10(10AA)-Earned leave encashment on Retirement" Then
                      fmsgboxStatus "Deduction u/s 10(10AA) shall be restricted to Rs. 25 Lakh in case of  employer catergory is selected as other than Central or State Government"
                End If
            End If
      End If
    End If
Next

'----
'Jyoti 26_06_2025
count = Sheet39.Range("SalarySectionCount").value

For i = 1 To count
 
If UCase(Sheet39.Range("sheet.EmployerCategory" & i).value = "Public Sector Undertaking" Or Sheet39.Range("sheet.EmployerCategory" & i).value = "PSU-Pensioners" Or Sheet39.Range("sheet.EmployerCategory" & i).value = "Others-Pensioners" Or Sheet39.Range("sheet.EmployerCategory" & i).value = "OTHERS") Then

    If Not Intersect(Range("Salary.Amount1"), Range(Target.Address)) Is Nothing Then

          If Range(Target.Address).value > 2000000 Then '= "Sec 10(10)-Death-cum-retirement gratuity received " Then

              If Trim(Sheet39.Range(Replace(Target.Address, "J", "G")).value) = Trim("Sec 10(10)-Death-cum-retirement gratuity received ") Then

                  'fmsgboxStatus " For Sec 10(10)-Death-cum-retirement gratuity is selected than amount should not exceed Rs.20 Lakhs "

                   fmsgboxStatus "Amount u/s 10(10) can not exceed Rs. 20 Lakhs where Nature of employement is ""Public Sector Undertaking"", ""PSU-Pensioners"", ""Others-Pensioners"" or ""Others"""

                   Sheet39.Range(Target.Address).value = ""

             End If

             calSchS

          End If

     End If
 
     ElseIf UCase(Sheet39.Range("sheet.EmployerCategory" & i).value = "CG-Pensioners") Or Sheet39.Range("sheet.EmployerCategory" & i).value = "SG-Pensioners" Or Sheet39.Range("sheet.EmployerCategory" & i).value = "Central Government" Or Sheet39.Range("sheet.EmployerCategory" & i).value = "State Government" Then

     If Not Intersect(Range("Salary.Amount1"), Range(Target.Address)) Is Nothing Then

               If Range(Target.Address).value > 2500000 Then '= "Sec 10(10)-Death-cum-retirement gratuity received " Then

                   If Trim(Sheet39.Range(Replace(Target.Address, "J", "G")).value) = Trim("Sec 10(10)-Death-cum-retirement gratuity received ") Then

                    ' fmsgboxStatus " For Sec 10(10)-Death-cum-retirement gratuity is selected than amount should not exceed Rs.20 Lakhs "

                   fmsgboxStatus "Amount u/s 10(10) can not exceed Rs. 25 Lakhs where Nature of employment is""Central Government"", ""State Government"", ""CG-Pensioners"" or ""SG-Pensioners"""

                  Sheet39.Range(Target.Address).value = ""

                   End If

                 calSchS

              End If

    End If

End If

Next
 
'----

If Target.Validation.Type = 3 Then
    If Not Application.Intersect(Sheet39.Range("Salary.Nature1"), Sheet39.Range(Target.Address)) Is Nothing Then
         sPassword = EfilingCommon.getmsgstate
         Sheet39.Unprotect Password:=sPassword
         ChangeAddress = Replace(Target.Address, "G", "H")
            If Mid(Sheet39.Range(Target.Address).value, 1, 1) = "A" Then
                 Sheet39.Range(ChangeAddress).Locked = False
                 Sheet39.Range(ChangeAddress).Interior.Color = "&HCCFFCC"
                 Sheet39.Range(ChangeAddress).ClearContents
            Else
                 Sheet39.Range(ChangeAddress).Locked = True
                 Sheet39.Range(ChangeAddress).Interior.Color = "&HD8D8D8"
                 Sheet39.Range(ChangeAddress).ClearContents
            End If
        Sheet39.Protect Password:=sPassword
        Dim nOfCount1 As Long
        nOfCount1 = CountDropdownAlwnc(Trim(Target.value), Sheet39.Range("Salary.Nature1").Cells, Sheet39.Range("Salary.Nature1").count)
'Change 102 30.11.2023 Ver 0.6 SIT-62835
'                If currCntAlwnc1 > 1 Or currCntAlwnc2 > 1 Or currCntAlwnc3 > 1 Or currCntAlwnc4 > 1 Or currCntAlwnc5 > 1 Or currCntAlwnc6 > 1 Or currCntAlwnc7 > 1 Or currCntAlwnc8 > 1 Or currCntAlwnc9 > 1 Or currCntAlwnc10 > 1 Or currCntAlwnc11 > 1 Or currCntAlwnc12 > 1 Or currCntAlwnc13 > 1 Or currCntAlwnc14 > 1 Then
                If currCntAlwnc1 > 1 Or currCntAlwnc2 > 1 Or currCntAlwnc3 > 1 Or currCntAlwnc4 > 1 Or currCntAlwnc5 > 1 Or currCntAlwnc6 > 1 Or currCntAlwnc7 > 1 Or currCntAlwnc8 > 1 Or currCntAlwnc9 > 1 Or currCntAlwnc10 > 1 Or currCntAlwnc11 > 1 Or currCntAlwnc12 > 1 Or currCntAlwnc13 > 1 Or currCntAlwnc15 > 1 Or currCntAlwnc16 > 1 Then
'End Change
            fmsgboxStatus "Drop down has already been selected" ', vbOKOnly, "Error(s)!"
            Target.value = "(Select)"
        End If


'Newly modified by Sai for SIT-69391
'If Target.Column = 7 Then
'        If Not (Target.value = "(Select)" Or Target.value = "") Then
'        If 1 < Application.CountIf(Range("Salary.Nature1"), Target.value) Then
''             fmsgbox "* Same drop downs cannot be selected more than once in Schedule BP"
'             fmsgboxStatus "Drop down has already been selected" ', vbOKOnly, "Error(s)!"
'             Range(Target.Address).value = "(Select)"
'        End If
'        End If
'End If



        If Range(Target.Address).value = "Sec 10(10)-Death-cum-retirement gratuity received " Then
            If Sheet39.Range(Replace(Target.Address, "G", "J")).value > 2000000 Then
            fmsgboxStatus " For Sec 10(10)-Death-cum-retirement gratuity is selected than amount should not exceed .20 Lakhs "
            Sheet39.Range(Target.Address).value = "(Select)"
            End If
        End If

        If Range(Target.Address).value = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
            If Sheet39.Range(Replace(Target.Address, "G", "J")).value > 500000 Then
            fmsgboxStatus "Exempt Allowance u/s 10(10B)-First Proviso-  Compensation limit notified by CG in the Official Gazette can not exceed Rs.500,000"
            Sheet39.Range(Target.Address).value = "(Select)"
            End If
        End If

        If Range(Target.Address).value = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
            If Sheet39.Range(Replace(Target.Address, "G", "J")).value > 500000 Then
            fmsgboxStatus "Exempt allowance u/s 10(10B) Second Proviso- Compensation under scheme approved by CG cannot exceed Rs. 5 lakhs."
            Sheet39.Range(Target.Address).value = "(Select)"
            End If
        End If
'Change 102 28.11.2023 Ver0.6 SIT-62868 A
        If Range(Target.Address).value = "Sec 10(10C)-Amount received/receivable on voluntary retirement or termination of service" Then
            If Sheet39.Range(Replace(Target.Address, "G", "J")).value > 500000 Then
            fmsgboxStatus "Kindly restrict the exemption u/s 10(10C) to Rs. 5,00,000"
            Sheet39.Range(Target.Address).value = "(Select)"
            End If
        End If
'End Change
        If Mid(Range(Target.Address).value, 1, 11) = "Sec 10(13A)" Then
            If Sheet39.Range(Replace(Target.Address, "G", "J")).value > WorksheetFunction.Min(0.5 * Sheet19.Range("Temp_DA_BS").value, Sheet19.Range("Temp_HRA").value) Then
            fmsgboxStatus "Sec 10(13A)-Allowance to meet expenditure incurred on house rent- amount of exemption should be restricted to min of amount entered under  ""House Rent Allowance"" in sl no B1a or 50% of ""Basic+DA"" as under dropdowns of 1(a)"
            Sheet39.Range(Target.Address).value = "(Select)"
            End If
        Else
        End If
           Dim ChangeAddressFrom As String
           Dim ChangeAddressTo As String
           ChangeAddressFrom = Replace(Target.Address, "G", "H")
                If Mid(Sheet39.Range(Target.Address).value, 1, 1) = "A" Then
                    Sheet39.Range(ChangeAddressFrom).Locked = False
                    Sheet39.Range(ChangeAddressFrom).Interior.Color = "&HCCFFCC"
                    Sheet39.Range(ChangeAddressFrom).ClearContents
                Else
                    Sheet39.Range(ChangeAddressFrom).Locked = True
                    Sheet39.Range(ChangeAddressFrom).Interior.Color = "&HD8D8D8"
                    Sheet39.Range(ChangeAddressFrom).ClearContents
                End If
        calSchS
    End If
End If

'
If Target.Validation.Type = 1 Then
    Application.ScreenUpdating = False
    If Intersect(Range("SAL.DeductionUnder6Section1"), Range(Target.Address)) Is Nothing Then
        AddTotalSalary
    End If
    Application.ScreenUpdating = True
End If
 '
If Not Intersect(Range("SAL.Entertainment_allowance_1"), Range(Target.Address)) Is Nothing Then
   Dim temp1 As Variant
       temp1 = 0
       temp1 = Sheet39.Range("Temp_allowance").value / 5
       If Range(Target.Address).value > WorksheetFunction.Min(temp1, 5000) Then
'Change 102 28.11.2023 SIT-62833
'         fmsgboxStatus "Deduction of Entertainment allowance u/s 16(ii) Maximum amount allowable in this field shall be 1/5th of Basic Salary 1a of all employers or Rs 5000 whichever is lower."
         fmsgboxStatus "Deduction of Entertainment allowance u/s 16(ii) can not exceed 1/5th of net Salary (SL no.4) or Rs. 5,000 whichever is lower"
'End Change SIT-62833
         Range("SAL.Entertainment_allowance_1").value = ""
         Range("SAL.Entertainment_allowance_1").Select
       End If
End If
'
If Not Intersect(Range("SAL.DeductionUnderSection1"), Range(Target.Address)) Is Nothing Then
   If Range("SAL.DeductionUnderSection1").value > 5000 Then
    fmsgboxStatus "Deduction of Professional tax u/s 16(iii) is Rs. 200/- p.m. in most of the state. Kindly reconsider the amount claimed"
   End If
End If

'
 If Not Application.Intersect(Sheet15.Range("Sch10of13A_50Por40Pofsalary"), Sheet39.Range(Target.Address)) Is Nothing Then
            Dim Per50Salaryaspersection17_i As Variant
            Per50Salaryaspersection17_i = Sheet39.Range("Sch10of13A_DetlsofSalpersec17of1").value * 0.5
                If Sheet39.Range("Sch10of13A_50Por40Pofsalary").value > Per50Salaryaspersection17_i Then
                  'fmsgbox "* ""The value in this field shall not be more than 50% of field 4 ""Details of Salary as per Section 17(1)"""""
                   fmsgbox "* ""The value in this Field 6 ""50% /40% of salary (C)"" shall not be more than 50% of Field 4 ""Details of Salary as per Section 17(1)"""
                   Sheet15.Range("Sch10of13A_50Por40Pofsalary").value = ""
                End If
 End If
                
''--------------------
'If Not Application.Intersect(Sheet39.Range("Salary.SalaryAmount1"), Sheet39.Range(target.Address)) Is Nothing Then
'    fmsgboxStatus """Nature of Exempt Allowance(s) will get refreshed if the field ""Salary as per Section 17(1)"" is modified""."
'End If

'Added by Aavula Naresh V0.6 Changes

Dim BasicSalary_count, DearnessAllowance_count, j As Variant
BasicSalary_count = 0
DearnessAllowance_count = 0
Dim HouseRentAllowance_cnt As Variant
HouseRentAllowance_cnt = 0

Dim basiccnt, dearallowancecnt  As Variant
Dim Basic_dear_TOTAL As Variant
basiccnt = 0
dearallowancecnt = 0
HouseRentAllowancecnt = 0

For i = 1 To count
BasicSalary_count = WorksheetFunction.SumIf(Sheet39.Range("Salary.Section" & i), "Basic Salary", Sheet39.Range("Salary.SalaryAmount" & i))
'DearnessAllowance_count = WorksheetFunction.SumIf(Sheet39.Range("Salary.Section" & i), "Dearness Allowance", Sheet39.Range("Salary.SalaryAmount" & i))

'For SIT-95254 by sai on 25/06/2025
DearnessAllowance_count = WorksheetFunction.SumIf(Sheet39.Range("Salary.Section" & i), "Dearness Allowance(DA)", Sheet39.Range("Salary.SalaryAmount" & i))

basiccnt = basiccnt + BasicSalary_count
dearallowancecnt = dearallowancecnt + DearnessAllowance_count

Next

Basic_dear_TOTAL = basiccnt + dearallowancecnt

Dim msgflag As Boolean
msgflag = True
For i = 1 To count
If Not Application.Intersect(Sheet39.Range("Salary.Section" & i), Sheet39.Range(Target.Address)) Is Nothing Or _
   Not Application.Intersect(Sheet39.Range("Salary.SalaryAmount" & i), Sheet39.Range(Target.Address)) Is Nothing Or _
   Not Application.Intersect(Sheet39.Range("Sch10of13A_DetlsofSalpersec17of1"), Sheet39.Range(Target.Address)) Is Nothing Then
       If Basic_dear_TOTAL > 0 And Sheet39.Range("Sch10of13A_DetlsofSalpersec17of1").value > 0 Then
                If msgflag = True Then
                If Sheet39.Range("Sch10of13A_DetlsofSalpersec17of1").value > Basic_dear_TOTAL Then
                    fmsgboxStatus "Salary for the purpose of this computation shall not be more than the basic salary and dearness allowance at field 17(1) in schedule salary."
                    Sheet39.Range("Sch10of13A_DetlsofSalpersec17of1").value = ""
                    msgflag = False
                End If
                End If
       End If
End If

If Not Intersect(Sheet39.Range("Salary.SalaryAmount" & i), Sheet39.Range(Target.Address)) Is Nothing Then
If Len(Target.value) > 0 Then

    fmsgboxStatus """Nature of Exempt Allowance(s) will get refreshed if the field ""Salary as per Section 17(1)"" is modified""."
End If
End If
Next

endline:
Application.EnableEvents = True
End Sub
Private Sub Worksheet_Change2122(ByVal Target As Range)
On Error GoTo endline

CalculateTaxCheckFlag = False

Dim count, i As Variant

Application.EnableEvents = False




calSchS
count = Sheet39.Range("SalarySectionCount").value

Dim temp_Pensioner, temp_PensionerCG, temp_PensionerSG, temp_PensionerPSU, temp_PensionerOth As Variant
Dim temp_Others, temp As Variant
Dim Salcount1 As Long
Salcount1 = 0
temp_Pensioner = 0
temp_PensionerCG = 0
temp_PensionerSG = 0
temp_PensionerPSU = 0
temp_PensionerOth = 0
temp_Others = 0
temp = 0

For Salcount1 = 1 To Sheet39.Range("SalarySectionCount").value

 If Not isdropdownblank(Sheet39.Range("sheet.EmployerCategory" & Salcount1).value) Then
 temp = temp + 1
 End If

' If Mid(Sheet39.Range("sheet.EmployerCategory" & Salcount1).value, 1, 3) = "Pen" Then
' temp_Pensioner = temp_Pensioner + 1
' End If

 If Sheet39.Range("sheet.EmployerCategory" & Salcount1).value = "CG-Pensioners" Then
 temp_PensionerCG = temp_Pensioner + 1
 End If
 
 If Sheet39.Range("sheet.EmployerCategory" & Salcount1).value = "SG-Pensioners" Then
 temp_PensionerSG = temp_Pensioner + 1
 End If
 
 If Sheet39.Range("sheet.EmployerCategory" & Salcount1).value = "PSU-Pensioners" Then
 temp_PensionerPSU = temp_Pensioner + 1
 End If
 
 If Sheet39.Range("sheet.EmployerCategory" & Salcount1).value = "Others-Pensioners" Then
 temp_PensionerOth = temp_Pensioner + 1
 End If
    

 If Mid(Sheet39.Range("sheet.EmployerCategory" & Salcount1).value, 1, 1) = "O" Then
 temp_Others = temp_Others + 1
 End If
Next


'If Temp = temp_Pensioner Or Temp = temp_Others Or SheetALL.Range("bacValue").value = 1 Then
If temp = temp_PensionerCG Or temp = temp_PensionerSG Or temp = temp_PensionerPSU Or temp = temp_PensionerOth Or temp = temp_Others Or SheetALL.Range("bacValue").value = 1 Then
Sheet39.Unprotect Password:=getmsgstate
Sheet39.Range("SAL.Entertainment_allowance_1").value = ""
Sheet39.Range("SAL.Entertainment_allowance_1").Locked = True
Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HD8D8D8"
Sheet39.Protect Password:=getmsgstate
Else
Sheet39.Unprotect Password:=getmsgstate
Sheet39.Range("SAL.Entertainment_allowance_1").Locked = False
Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HCCFFCC"
Sheet39.Protect Password:=getmsgstate
End If

If temp = (temp_Pensioner + temp_Others) And temp_Pensioner > 0 And temp_Others > 0 Then
Sheet39.Unprotect Password:=getmsgstate
Sheet39.Range("SAL.Entertainment_allowance_1").value = ""
Sheet39.Range("SAL.Entertainment_allowance_1").Locked = True
Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HD8D8D8"
Sheet39.Protect Password:=getmsgstate
End If


If (SheetALL.Range("bacValue").value) = 1 Then
Sheet39.Unprotect Password:=getmsgstate
Sheet39.Range("SAL.Entertainment_allowance_1").value = ""
Sheet39.Range("SAL.Entertainment_allowance_1").Locked = True
Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HD8D8D8"
Sheet39.Protect Password:=getmsgstate
End If


If temp = temp_Pensioner Or SheetALL.Range("bacValue").value = 1 Then
Sheet39.Unprotect Password:=getmsgstate
Sheet39.Range("SAL.DeductionUnderSection1").value = ""
Sheet39.Range("SAL.DeductionUnderSection1").Locked = True
Sheet39.Range("SAL.DeductionUnderSection1").Interior.Color = "&HD8D8D8"
Sheet39.Protect Password:=getmsgstate
Else
Sheet39.Unprotect Password:=getmsgstate
Sheet39.Range("SAL.DeductionUnderSection1").Locked = False
Sheet39.Range("SAL.DeductionUnderSection1").Interior.Color = "&HCCFFCC"
Sheet39.Protect Password:=getmsgstate
End If
'dpk1


If Not Intersect(Range("Salary.Amount1"), Range(Target.Address)) Is Nothing Then
'If Mid(Sheet39.Range(Replace(Target.address, "J", "G")).value, 1, 11) = "Sec 10(13A)" Then
'If Sheet39.Range(Target.address).value > WorksheetFunction.Min(0.5 * Sheet19.Range("Temp_DA_BS").value, Sheet19.Range("Temp_HRA").value) Then
'    fmsgboxStatus "Sec 10(13A)-Allowance to meet expenditure incurred on house rent- amount of exemption should be restricted to min of amount entered under  ""House Rent Allowance"" in sl no B1a or 50% of ""Basic+DA"" as under dropdowns of 1(a)"
'    Sheet39.Range(Target.address).ClearContents
'End If
'End If

    If Range(Target.Address).value > 2000000 Then '= "Sec 10(10)-Death-cum-retirement gratuity received " Then
'        If temp_Pensioner > 0 Or temp_Others > 0 Then
        If Sheet39.Range(Replace(Target.Address, "J", "G")).value = "Sec 10(10)-Death-cum-retirement gratuity received " Then
        fmsgboxStatus " For Sec 10(10)-Death-cum-retirement gratuity is selected than amount should not exceed .20 Lakhs "
        Sheet39.Range(Target.Address).value = ""
        End If
'        End If
    End If

    If Range(Target.Address).value > 500000 Then '= "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
    If Sheet39.Range(Replace(Target.Address, "J", "G")).value = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
    fmsgboxStatus "Exempt Allowance u/s 10(10B)-First Proviso-  Compensation limit notified by CG in the Official Gazette can not exceed Rs.500,000"
    Sheet39.Range(Target.Address).value = ""
    End If
    End If
    
    'new change
    If Range(Target.Address).value > 38400 Then '= "Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee" Then
    If Sheet39.Range(Replace(Target.Address, "J", "G")).value = "Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee" Then
    fmsgboxStatus "Exempt Allowance u/s Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee can not exceed Rs.38400"
    Sheet39.Range(Target.Address).value = ""
    End If
    End If
    'new change

'    If Range(Target.address).value > 500000 Then '= "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
'    If Sheet39.Range(Replace(Target.address, "J", "G")).value = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
'    fmsgboxStatus "Exempt allowance u/s 10(10B) Second Proviso- Compensation under scheme approved by CG cannot exceed Rs. 5 lakhs."
'    Sheet39.Range(Target.address).value = ""
'    End If
'    End If




End If
If Not Intersect(Range("Salary.Amount1"), Range(Target.Address)) Is Nothing Then
    If Mid(Sheet39.Range(Replace(Target.Address, "J", "G")), 1, 11) = "Sec 10(13A)" Then
    If Range(Target.Address).value > WorksheetFunction.Min(0.5 * Sheet19.Range("Temp_DA_BS").value, Sheet19.Range("Temp_HRA").value) Then
    fmsgboxStatus "Sec 10(13A)-Allowance to meet expenditure incurred on house rent- amount of exemption should be restricted to min of amount entered under  ""House Rent Allowance"" in sl no B1a or 50% of ""Basic+DA"" as under dropdowns of 1(a)"
    Range(Target.Address).value = ""
    End If
    End If
End If


If Target.Validation.Type = 3 Then
    If Not Intersect(Range("Salary.Nature1"), Range(Target.Address)) Is Nothing Then
sPassword = EfilingCommon.getmsgstate
         Sheet39.Unprotect Password:=sPassword

    'Dim ChangeAddress As String
         ChangeAddress = Replace(Target.Address, "G", "H")
    If Mid(Sheet39.Range(Target.Address).value, 1, 1) = "A" Then
         Sheet39.Range(ChangeAddress).Locked = False
         Sheet39.Range(ChangeAddress).Interior.Color = "&HCCFFCC"
         Sheet39.Range(ChangeAddress).ClearContents
    Else
         Sheet39.Range(ChangeAddress).Locked = True
         Sheet39.Range(ChangeAddress).Interior.Color = "&HD8D8D8"
         Sheet39.Range(ChangeAddress).ClearContents
    End If

      Sheet39.Protect Password:=sPassword
    Dim nOfCount1 As Long
                nOfCount1 = CountDropdownAlwnc(Trim(Target.value), Sheet39.Range("Salary.Nature1").Cells, Sheet39.Range("Salary.Nature1").count)
'Change 102 30.11.2023 Ver 0.6 SIT-62835
'                If currCntAlwnc1 > 1 Or currCntAlwnc2 > 1 Or currCntAlwnc3 > 1 Or currCntAlwnc4 > 1 Or currCntAlwnc5 > 1 Or currCntAlwnc6 > 1 Or currCntAlwnc7 > 1 Or currCntAlwnc8 > 1 Or currCntAlwnc9 > 1 Or currCntAlwnc10 > 1 Or currCntAlwnc11 > 1 Or currCntAlwnc12 > 1 Or currCntAlwnc13 > 1 Or currCntAlwnc14 > 1 Then
                If currCntAlwnc1 > 1 Or currCntAlwnc2 > 1 Or currCntAlwnc3 > 1 Or currCntAlwnc4 > 1 Or currCntAlwnc5 > 1 Or currCntAlwnc6 > 1 Or currCntAlwnc7 > 1 Or currCntAlwnc8 > 1 Or currCntAlwnc9 > 1 Or currCntAlwnc10 > 1 Or currCntAlwnc11 > 1 Or currCntAlwnc12 > 1 Or currCntAlwnc13 > 1 Then
'End Change
                    fmsgboxStatus "Drop down has already been selected" ', vbOKOnly, "Error(s)!"
                    Target.value = "(Select)"
                End If

    If Range(Target.Address).value = "Sec 10(10)-Death-cum-retirement gratuity received " Then
'    If temp_Pensioner > 0 Or temp_Others > 0 Then
    If Sheet39.Range(Replace(Target.Address, "G", "J")).value > 2000000 Then
    fmsgboxStatus " For Sec 10(10)-Death-cum-retirement gratuity is selected than amount should not exceed .20 Lakhs "
    Sheet39.Range(Target.Address).value = "(Select)"
'    End If
    End If
    End If

    If Range(Target.Address).value = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
    If Sheet39.Range(Replace(Target.Address, "G", "J")).value > 500000 Then
    fmsgboxStatus "Exempt Allowance u/s 10(10B)-First Proviso-  Compensation limit notified by CG in the Official Gazette can not exceed Rs.500,000"
    Sheet39.Range(Target.Address).value = "(Select)"
    End If
    End If

    If Range(Target.Address).value = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
    If Sheet39.Range(Replace(Target.Address, "G", "J")).value > 500000 Then
    fmsgboxStatus "Exempt allowance u/s 10(10B) Second Proviso- Compensation under scheme approved by CG cannot exceed Rs. 5 lakhs."
    Sheet39.Range(Target.Address).value = "(Select)"
    End If
    End If

    If Mid(Range(Target.Address).value, 1, 11) = "Sec 10(13A)" Then

    If Sheet39.Range(Replace(Target.Address, "G", "J")).value > WorksheetFunction.Min(0.5 * Sheet19.Range("Temp_DA_BS").value, Sheet19.Range("Temp_HRA").value) Then
    fmsgboxStatus "Sec 10(13A)-Allowance to meet expenditure incurred on house rent- amount of exemption should be restricted to min of amount entered under  ""House Rent Allowance"" in sl no B1a or 50% of ""Basic+DA"" as under dropdowns of 1(a)"
    Sheet39.Range(Target.Address).value = "(Select)"
    End If


   ' LOCK80GG
    Else
   'UNLOCK80GG



    End If




        Dim ChangeAddressFrom As String
        Dim ChangeAddressTo As String
        ChangeAddressFrom = Replace(Target.Address, "G", "H")
'        ChangeAddressTo = Replace(Target.Address, "F", "H")
'        ChangeAddressFrom = ChangeAddressFrom & ":" & ChangeAddressTo

            If Mid(Sheet39.Range(Target.Address).value, 1, 1) = "A" Then
                Sheet39.Range(ChangeAddressFrom).Locked = False
                Sheet39.Range(ChangeAddressFrom).Interior.Color = "&HCCFFCC"
                Sheet39.Range(ChangeAddressFrom).ClearContents
            Else
                Sheet39.Range(ChangeAddressFrom).Locked = True
                Sheet39.Range(ChangeAddressFrom).Interior.Color = "&HD8D8D8"
                Sheet39.Range(ChangeAddressFrom).ClearContents
            End If
    End If
End If


If Target.Validation.Type = 6 Then
    If (Target.name.name) = "SAL.PANofEmployer" & count Then
        Dim TargetRepTAN As String
        TargetRepTAN = ""
        TargetRepTAN = Target.name.name

        sPassword = EfilingCommon.getmsgstate
         Sheet39.Unprotect Password:=sPassword
            If Not ValidateTANCodes(UCase(Trim(Range(TargetRepTAN).value))) Then
                fmsgboxStatus "Invalid TAN" & Chr(10) & Chr(13) & "TAN format should be First 4 Alphabets, next 5 digits, then 1 Alphabet.Please enter Valid TAN" ', vbOKCancel
                Sheet39.Range(TargetRepTAN).Select
                Sheet39.Range(TargetRepTAN).value = ""
            Else
                Sheet39.Range(TargetRepTAN).value = UCase(Range(TargetRepTAN).value)
            End If
         Sheet39.Protect Password:=sPassword
    End If
End If




If Target.Validation.Type = 1 Then
    Application.ScreenUpdating = False




    If Intersect(Range("SAL.DeductionUnder6Section1"), Range(Target.Address)) Is Nothing Then
        AddTotalSalary
    End If
    Application.ScreenUpdating = True
End If

count = Sheet39.Range("SalarySectionCount").value
For i = 1 To count
Sheet39.Range("SAL.PANofEmployer" & i).value = UCase(Sheet39.Range("SAL.PANofEmployer" & i).value)

        sPassword = EfilingCommon.getmsgstate
         Sheet39.Unprotect Password:=sPassword


    Dim nOfCount  As Long

    If Not Intersect(Range("Salary.Section" & i), Range(Target.Address)) Is Nothing Then
                nOfCount = CountDropdown1(Trim(Target.value), Sheet39.Range("Salary.Section" & i).Cells, Sheet39.Range("Salary.Section" & i).count)
                If currCnt1 > 1 Or currCnt2 > 1 Or currCnt3 > 1 Or currCnt4 > 1 Or currCnt5 > 1 Or currCnt6 > 1 Or currCnt7 > 1 Or currCnt8 > 1 Or currCnt9 > 1 Or currCnt10 > 1 Or currCnt11 > 1 Or currCnt12 > 1 Or currCnt13 > 1 Or currCnt14 > 1 Or currCnt15 > 1 Or currCnt16 > 1 Then
                    fmsgboxStatus "Drop down has already been selected" ', vbOKOnly, "Error(s)!"
                    Target.value = "(Select)"
        End If
    End If

    If Not Intersect(Range("Salary.ValueSection" & i), Range(Target.Address)) Is Nothing Then
                nOfCount = CountDropdown2(Trim(Target.value), Sheet39.Range("Salary.ValueSection" & i).Cells, Sheet39.Range("Salary.ValueSection" & i).count)
                If currCntvalue1 > 1 Or currCntvalue2 > 1 Or currCntvalue3 > 1 Or currCntvalue4 > 1 Or currCntvalue5 > 1 Or currCntvalue6 > 1 Or currCntvalue7 > 1 Or currCntvalue8 > 1 Or _
                          currCntvalue9 > 1 Or currCntvalue10 > 1 Or currCntvalue11 > 1 Or currCntvalue12 > 1 Or currCntvalue13 > 1 Or currCntvalue14 > 1 Or currCntvalue15 > 1 Or currCntvalue16 > 1 Or currCntvalue17 > 1 _
                          Or currCntvalue18 > 1 Or currCntvalue19 > 1 Or currCntvalue20 > 1 Or currCntvalue21 > 1 Then
                    fmsgboxStatus "Drop down has already been selected" ', vbOKOnly, "Error(s)!"
                    Target.value = "(Select)"
        End If
    End If

    If Not Intersect(Range("Salary.ProfitSection" & i), Range(Target.Address)) Is Nothing Then
                nOfCount = CountDropdown3(Trim(Target.value), Sheet39.Range("Salary.ProfitSection" & i).Cells, Sheet39.Range("Salary.ProfitSection" & i).count)
                If currCntProfit1 > 1 Or currCntProfit2 > 1 Or currCntProfit3 > 1 Then
                    fmsgboxStatus "Drop down has already been selected" ', vbOKOnly, "Error(s)!"
                    Target.value = "(Select)"
        End If
    End If

   If Not Intersect(Range("Salary.Section" & i), Range(Target.Address)) Is Nothing Then
'    Dim ChangeAddress As String
         ChangeAddress = Replace(Target.Address, "G", "H")
    If Sheet39.Range(Target.Address).value = "Others" Then
         Sheet39.Range(ChangeAddress).Locked = False
         Sheet39.Range(ChangeAddress).Interior.Color = "&HCCFFCC"
         Sheet39.Range(ChangeAddress).ClearContents
    Else
         Sheet39.Range(ChangeAddress).Locked = True
         Sheet39.Range(ChangeAddress).Interior.Color = "&HD8D8D8"
         Sheet39.Range(ChangeAddress).ClearContents
    End If
   End If

   If Not Intersect(Range("Salary.ValueSection" & i), Range(Target.Address)) Is Nothing Then
    'Dim ChangeAddress As String
         ChangeAddress = Replace(Target.Address, "G", "H")
    If Sheet39.Range(Target.Address).value = "Other benefits or amenities" Then
         Sheet39.Range(ChangeAddress).Locked = False
         Sheet39.Range(ChangeAddress).Interior.Color = "&HCCFFCC"
         Sheet39.Range(ChangeAddress).ClearContents
    Else
         Sheet39.Range(ChangeAddress).Locked = True
         Sheet39.Range(ChangeAddress).Interior.Color = "&HD8D8D8"
         Sheet39.Range(ChangeAddress).ClearContents
    End If
   End If


   If Not Intersect(Range("Salary.ProfitSection" & i), Range(Target.Address)) Is Nothing Then
    'Dim ChangeAddress As String
         ChangeAddress = Replace(Target.Address, "G", "H")
    If Sheet39.Range(Target.Address).value = "Any Other" Then
         Sheet39.Range(ChangeAddress).Locked = False
         Sheet39.Range(ChangeAddress).Interior.Color = "&HCCFFCC"
         Sheet39.Range(ChangeAddress).ClearContents
    Else
         Sheet39.Range(ChangeAddress).Locked = True
         Sheet39.Range(ChangeAddress).Interior.Color = "&HD8D8D8"
         Sheet39.Range(ChangeAddress).ClearContents
    End If
   End If

'   If Not Intersect(Range("SAL.StateCode" & i), Range(Target.Address)) Is Nothing Then
'   If Mid(Sheet39.Range("SAL.StateCode" & i).value, 1, 2) = "99" Then
'         Sheet39.Range("SAL.ZipCode" & i).Locked = False
'         Sheet39.Range("SAL.ZipCode" & i).Interior.Color = "&HCCFFCC"
'         Sheet39.Range("SAL.ZipCode" & i).ClearContents
'         Sheet39.Range("SAL.PinCode" & i).Locked = True
'         Sheet39.Range("SAL.PinCode" & i).Interior.Color = "&HD8D8D8"
'         Sheet39.Range("SAL.PinCode" & i).ClearContents
'   Else
'         Sheet39.Range("SAL.PinCode" & i).Locked = False
'         Sheet39.Range("SAL.PinCode" & i).Interior.Color = "&HCCFFCC"
'         Sheet39.Range("SAL.PinCode" & i).ClearContents
'         Sheet39.Range("SAL.ZipCode" & i).Locked = True
'         Sheet39.Range("SAL.ZipCode" & i).Interior.Color = "&HD8D8D8"
'         Sheet39.Range("SAL.ZipCode" & i).ClearContents
'   End If
'   End If

















'   If Not Intersect(Range("sheet.EmployerCategory" & i), Range(Target.address)) Is Nothing Then
'   If Sheet39.Range("sheet.EmployerCategory" & i).value = "Pensioners" Or Mid(Sheet39.Range("sheet.EmployerCategory" & i).value, 1, 3) = "OTH" Then
'         Sheet39.Range("SAL.Entertainment_allowance_1").ClearContents
'         Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HD8D8D8"
'         Sheet39.Range("SAL.Entertainment_allowance_1").Locked = True
'
'   Else
'         'Sheet39.Range("SAL.Entertainment_allowance_" & i).ClearContents
'         Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HCCFFCC"
'         Sheet39.Range("SAL.Entertainment_allowance_1").Locked = False
'
'   End If
'
'   If Sheet39.Range("sheet.EmployerCategory" & i).value = "Pensioners" Then
'      Sheet39.Range("SAL.DeductionUnderSection1").ClearContents
'      Sheet39.Range("SAL.DeductionUnderSection1").Locked = True
'      Sheet39.Range("SAL.DeductionUnderSection1").Interior.Color = "&HD8D8D8"
'
'Else
'      'Sheet39.Range("SAL.Entertainment_allowance_" & i).ClearContents
'      Sheet39.Range("SAL.DeductionUnderSection1").Interior.Color = "&HCCFFCC"
'      Sheet39.Range("SAL.DeductionUnderSection1").Locked = False
'   End If
'
'End If
Next

  If Target.Validation.Type = 1 Then
    Application.ScreenUpdating = False
    If Intersect(Range("SAL.DeductionUnder6Section1"), Range(Target.Address)) Is Nothing Then
        AddTotalSalary

    End If
    Application.ScreenUpdating = True
End If

   Sheet39.Protect Password:=sPassword










'If Target.name.name = "SAL.DeductionUnder6Section1" Then
            
            
            
            
             If (Range("SAL.DeductionUnder6Section1").value) > 50000 Then
            fmsgboxStatus "Deduction u/s. 16(ia) cannot exceed Rs. 50000 or Net Salary whichever is lower in Schedule Salary."
            Range("SAL.DeductionUnder6Section1").value = 50000
            Range("SAL.DeductionUnder6Section1").Select
             End If

             Dim temp16a, temp16aa As Variant
            temp16a = Sheet39.Range("SAL.DeductionUnder6Section1").value
            temp16aa = Sheet39.Range("SAL.Netsalary1").value


        If temp16a <> "" Then
        If Not Sheet39.Range("SAL.DeductionUnder6Section1").value <= temp16aa Then
               fmsgboxStatus "Deduction u/s. 16(ia) cannot exceed Rs. 50000 or Net Salary whichever is lower in Schedule Salary"
               Range("SAL.DeductionUnder6Section1").value = temp16aa
               Range("SAL.DeductionUnder6Section1").Select


        End If
        End If
        
        If SheetALL.Range("bacValue").value = 1 Then
            Range("SAL.DeductionUnder6Section1").value = 0
            
            End If
        
'End If

 If Target.name.name = "SAL.Entertainment_allowance_1" Or Target.name.name = "SAL.Netsalary1" Then
   Dim temp1 As Variant
       temp1 = 0
       temp1 = Sheet39.Range("Temp_allowance").value / 5

       If Range(Target.Address).value > WorksheetFunction.Min(temp1, 5000) Then
         fmsgboxStatus "Deduction of Entertainment allowance u/s 16(ii) Maximum amount allowable in this field shall be 1/5th of Basic Salary 1a of all employers or Rs 5000 whichever is lower."
         Range("SAL.Entertainment_allowance_1").value = ""
         Range("SAL.Entertainment_allowance_1").Select
       End If
   End If

   If Target.name.name = "SAL.DeductionUnderSection1" Then
   If Range("SAL.DeductionUnderSection1").value > 5000 Then
    fmsgboxStatus "Deduction of Professional tax u/s 16(iii) is Rs. 200/- p.m. in most of the state. Kindly reconsider the amount claimed"
'    Range("SAL.DeductionUnderSection1").value = ""
'    Range("SAL.DeductionUnderSection1").Select
   End If
   End If


endline:
Application.EnableEvents = True
End Sub
Private Sub Worksheet_Changeold2(ByVal Target As Range)
On Error GoTo endline

CalculateTaxCheckFlag = False

Dim count, i As Variant

Application.EnableEvents = False


calSchS
count = Sheet39.Range("SalarySectionCount").value
 validate10aaAmount
Dim temp_Pensioner As Variant
Dim temp_Others, temp As Variant
Dim Salcount1 As Long
Salcount1 = 0
temp_Pensioner = 0
temp_Others = 0
temp = 0
For Salcount1 = 1 To Sheet39.Range("SalarySectionCount").value

 If Not isdropdownblank(Sheet39.Range("sheet.EmployerCategory" & Salcount1).value) Then
 temp = temp + 1
 End If

 If Mid(Sheet39.Range("sheet.EmployerCategory" & Salcount1).value, 1, 3) = "Pen" Then
 temp_Pensioner = temp_Pensioner + 1
 End If

 If Mid(Sheet39.Range("sheet.EmployerCategory" & Salcount1).value, 1, 1) = "O" Then
 temp_Others = temp_Others + 1
 End If
Next


If temp = temp_Pensioner Or temp = temp_Others Or SheetALL.Range("bacValue").value = 1 Then
Sheet39.Unprotect Password:=getmsgstate
Sheet39.Range("SAL.Entertainment_allowance_1").value = ""
Sheet39.Range("SAL.Entertainment_allowance_1").Locked = True
Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HD8D8D8"
Sheet39.Protect Password:=getmsgstate
Else
Sheet39.Unprotect Password:=getmsgstate
Sheet39.Range("SAL.Entertainment_allowance_1").Locked = False
Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HCCFFCC"
Sheet39.Protect Password:=getmsgstate
End If

If temp = (temp_Pensioner + temp_Others) And temp_Pensioner > 0 And temp_Others > 0 Then
Sheet39.Unprotect Password:=getmsgstate
Sheet39.Range("SAL.Entertainment_allowance_1").value = ""
Sheet39.Range("SAL.Entertainment_allowance_1").Locked = True
Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HD8D8D8"
Sheet39.Protect Password:=getmsgstate
End If


If (SheetALL.Range("bacValue").value) = 1 Then
Sheet39.Unprotect Password:=getmsgstate
Sheet39.Range("SAL.Entertainment_allowance_1").value = ""
Sheet39.Range("SAL.Entertainment_allowance_1").Locked = True
Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HD8D8D8"
Sheet39.Protect Password:=getmsgstate
End If


If temp = temp_Pensioner Or SheetALL.Range("bacValue").value = 1 Then
Sheet39.Unprotect Password:=getmsgstate
Sheet39.Range("SAL.DeductionUnderSection1").value = ""
Sheet39.Range("SAL.DeductionUnderSection1").Locked = True
Sheet39.Range("SAL.DeductionUnderSection1").Interior.Color = "&HD8D8D8"
Sheet39.Protect Password:=getmsgstate
Else
Sheet39.Unprotect Password:=getmsgstate
Sheet39.Range("SAL.DeductionUnderSection1").Locked = False
Sheet39.Range("SAL.DeductionUnderSection1").Interior.Color = "&HCCFFCC"
Sheet39.Protect Password:=getmsgstate
End If

If Not Intersect(Range("Salary.Amount1"), Range(Target.Address)) Is Nothing Then
'If Mid(Sheet39.Range(Replace(Target.address, "J", "G")).value, 1, 11) = "Sec 10(13A)" Then
'If Sheet39.Range(Target.address).value > WorksheetFunction.Min(0.5 * Sheet19.Range("Temp_DA_BS").value, Sheet19.Range("Temp_HRA").value) Then
'    fmsgboxStatus "Sec 10(13A)-Allowance to meet expenditure incurred on house rent- amount of exemption should be restricted to min of amount entered under  ""House Rent Allowance"" in sl no B1a or 50% of ""Basic+DA"" as under dropdowns of 1(a)"
'    Sheet39.Range(Target.address).ClearContents
'End If
'End If

    If Range(Target.Address).value > 2000000 Then '= "Sec 10(10)-Death-cum-retirement gratuity received " Then
'        If temp_Pensioner > 0 Or temp_Others > 0 Then
        If Sheet39.Range(Replace(Target.Address, "J", "G")).value = "Sec 10(10)-Death-cum-retirement gratuity received " Then
        fmsgboxStatus " For Sec 10(10)-Death-cum-retirement gratuity is selected than amount should not exceed .20 Lakhs "
        Sheet39.Range(Target.Address).value = ""
        End If
'        End If
    End If

    If Range(Target.Address).value > 500000 Then '= "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
    If Sheet39.Range(Replace(Target.Address, "J", "G")).value = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
    fmsgboxStatus "Exempt Allowance u/s 10(10B)-First Proviso-  Compensation limit notified by CG in the Official Gazette can not exceed Rs.500,000"
    Sheet39.Range(Target.Address).value = ""
    End If
    End If
    
    'new change
    If Range(Target.Address).value > 38400 Then '= "Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee" Then
    If Sheet39.Range(Replace(Target.Address, "J", "G")).value = "Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee" Then
    fmsgboxStatus "Exempt Allowance u/s Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee can not exceed Rs.38400"
    Sheet39.Range(Target.Address).value = ""
    End If
    End If
    'new change

'    If Range(Target.address).value > 500000 Then '= "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
'    If Sheet39.Range(Replace(Target.address, "J", "G")).value = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
'    fmsgboxStatus "Exempt allowance u/s 10(10B) Second Proviso- Compensation under scheme approved by CG cannot exceed Rs. 5 lakhs."
'    Sheet39.Range(Target.address).value = ""
'    End If
'    End If




End If


If Target.Validation.Type = 3 Then
    If Not Intersect(Range("Salary.Nature1"), Range(Target.Address)) Is Nothing Then
sPassword = EfilingCommon.getmsgstate
         Sheet39.Unprotect Password:=sPassword

    'Dim ChangeAddress As String
         ChangeAddress = Replace(Target.Address, "G", "H")
    If Mid(Sheet39.Range(Target.Address).value, 1, 1) = "A" Then
         Sheet39.Range(ChangeAddress).Locked = False
         Sheet39.Range(ChangeAddress).Interior.Color = "&HCCFFCC"
         Sheet39.Range(ChangeAddress).ClearContents
    Else
         Sheet39.Range(ChangeAddress).Locked = True
         Sheet39.Range(ChangeAddress).Interior.Color = "&HD8D8D8"
         Sheet39.Range(ChangeAddress).ClearContents
    End If

      Sheet39.Protect Password:=sPassword
    Dim nOfCount1 As Long
                nOfCount1 = CountDropdownAlwnc(Trim(Target.value), Sheet39.Range("Salary.Nature1").Cells, Sheet39.Range("Salary.Nature1").count)
                If currCntAlwnc1 > 1 Or currCntAlwnc2 > 1 Or currCntAlwnc3 > 1 Or currCntAlwnc4 > 1 Or currCntAlwnc5 > 1 Or currCntAlwnc6 > 1 Or currCntAlwnc7 > 1 Or currCntAlwnc8 > 1 Or currCntAlwnc9 > 1 Or currCntAlwnc10 > 1 Or currCntAlwnc11 > 1 Or currCntAlwnc12 > 1 Or currCntAlwnc13 > 1 Or currCntAlwnc14 > 1 Then
                    fmsgboxStatus "Drop down has already been selected" ', vbOKOnly, "Error(s)!"
                    Target.value = "(Select)"
                End If

    If Range(Target.Address).value = "Sec 10(10)-Death-cum-retirement gratuity received " Then
'    If temp_Pensioner > 0 Or temp_Others > 0 Then
    If Sheet39.Range(Replace(Target.Address, "G", "J")).value > 2000000 Then
    fmsgboxStatus " For Sec 10(10)-Death-cum-retirement gratuity is selected than amount should not exceed .20 Lakhs "
    Sheet39.Range(Target.Address).value = "(Select)"
'    End If
    End If
    End If

    If Range(Target.Address).value = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
    If Sheet39.Range(Replace(Target.Address, "G", "J")).value > 500000 Then
    fmsgboxStatus "Exempt Allowance u/s 10(10B)-First Proviso-  Compensation limit notified by CG in the Official Gazette can not exceed Rs.500,000"
    Sheet39.Range(Target.Address).value = "(Select)"
    End If
    End If

    If Range(Target.Address).value = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
    If Sheet39.Range(Replace(Target.Address, "G", "J")).value > 500000 Then
    fmsgboxStatus "Exempt allowance u/s 10(10B) Second Proviso- Compensation under scheme approved by CG cannot exceed Rs. 5 lakhs."
    Sheet39.Range(Target.Address).value = "(Select)"
    End If
    End If

    If Mid(Range(Target.Address).value, 1, 11) = "Sec 10(13A)" Then

    If Sheet39.Range(Replace(Target.Address, "G", "J")).value > WorksheetFunction.Min(0.5 * Sheet19.Range("Temp_DA_BS").value, Sheet19.Range("Temp_HRA").value) Then
    fmsgboxStatus "Sec 10(13A)-Allowance to meet expenditure incurred on house rent- amount of exemption should be restricted to min of amount entered under  ""House Rent Allowance"" in sl no B1a or 50% of ""Basic+DA"" as under dropdowns of 1(a)"
    Sheet39.Range(Target.Address).value = "(Select)"
    End If


   ' LOCK80GG
    Else
   'UNLOCK80GG



    End If




        Dim ChangeAddressFrom As String
        Dim ChangeAddressTo As String
        ChangeAddressFrom = Replace(Target.Address, "G", "H")
'        ChangeAddressTo = Replace(Target.Address, "F", "H")
'        ChangeAddressFrom = ChangeAddressFrom & ":" & ChangeAddressTo

            If Mid(Sheet39.Range(Target.Address).value, 1, 1) = "A" Then
                Sheet39.Range(ChangeAddressFrom).Locked = False
                Sheet39.Range(ChangeAddressFrom).Interior.Color = "&HCCFFCC"
                Sheet39.Range(ChangeAddressFrom).ClearContents
            Else
                Sheet39.Range(ChangeAddressFrom).Locked = True
                Sheet39.Range(ChangeAddressFrom).Interior.Color = "&HD8D8D8"
                Sheet39.Range(ChangeAddressFrom).ClearContents
            End If
    End If
End If


If Target.Validation.Type = 6 Then
    If (Target.name.name) = "SAL.PANofEmployer" & count Then
        Dim TargetRepTAN As String
        TargetRepTAN = ""
        TargetRepTAN = Target.name.name

        sPassword = EfilingCommon.getmsgstate
         Sheet39.Unprotect Password:=sPassword
            If Not ValidateTANCodes(UCase(Trim(Range(TargetRepTAN).value))) Then
                fmsgboxStatus "Invalid TAN" & Chr(10) & Chr(13) & "TAN format should be First 4 Alphabets, next 5 digits, then 1 Alphabet.Please enter Valid TAN" ', vbOKCancel
                Sheet39.Range(TargetRepTAN).Select
                Sheet39.Range(TargetRepTAN).value = ""
            Else
                Sheet39.Range(TargetRepTAN).value = UCase(Range(TargetRepTAN).value)
            End If
         Sheet39.Protect Password:=sPassword
    End If
End If




If Target.Validation.Type = 1 Then
    Application.ScreenUpdating = False




    If Intersect(Range("SAL.DeductionUnder6Section1"), Range(Target.Address)) Is Nothing Then
        AddTotalSalary
    End If
    Application.ScreenUpdating = True
End If

count = Sheet39.Range("SalarySectionCount").value
For i = 1 To count
Sheet39.Range("SAL.PANofEmployer" & i).value = UCase(Sheet39.Range("SAL.PANofEmployer" & i).value)

        sPassword = EfilingCommon.getmsgstate
         Sheet39.Unprotect Password:=sPassword


    Dim nOfCount  As Long

    If Not Intersect(Range("Salary.Section" & i), Range(Target.Address)) Is Nothing Then
                nOfCount = CountDropdown1(Trim(Target.value), Sheet39.Range("Salary.Section" & i).Cells, Sheet39.Range("Salary.Section" & i).count)
                If currCnt1 > 1 Or currCnt2 > 1 Or currCnt3 > 1 Or currCnt4 > 1 Or currCnt5 > 1 Or currCnt6 > 1 Or currCnt7 > 1 Or currCnt8 > 1 Or currCnt9 > 1 Or currCnt10 > 1 Or currCnt11 > 1 Or currCnt12 > 1 Or currCnt13 > 1 Or currCnt14 > 1 Or currCnt15 > 1 Or currCnt16 > 1 Then
                    fmsgboxStatus "Drop down has already been selected" ', vbOKOnly, "Error(s)!"
                    Target.value = "(Select)"
        End If
    End If

    If Not Intersect(Range("Salary.ValueSection" & i), Range(Target.Address)) Is Nothing Then
                nOfCount = CountDropdown2(Trim(Target.value), Sheet39.Range("Salary.ValueSection" & i).Cells, Sheet39.Range("Salary.ValueSection" & i).count)
                If currCntvalue1 > 1 Or currCntvalue2 > 1 Or currCntvalue3 > 1 Or currCntvalue4 > 1 Or currCntvalue5 > 1 Or currCntvalue6 > 1 Or currCntvalue7 > 1 Or currCntvalue8 > 1 Or _
                          currCntvalue9 > 1 Or currCntvalue10 > 1 Or currCntvalue11 > 1 Or currCntvalue12 > 1 Or currCntvalue13 > 1 Or currCntvalue14 > 1 Or currCntvalue15 > 1 Or currCntvalue16 > 1 Or currCntvalue17 > 1 Then
                    fmsgboxStatus "Drop down has already been selected" ', vbOKOnly, "Error(s)!"
                    Target.value = "(Select)"
        End If
    End If

    If Not Intersect(Range("Salary.ProfitSection" & i), Range(Target.Address)) Is Nothing Then
                nOfCount = CountDropdown3(Trim(Target.value), Sheet39.Range("Salary.ProfitSection" & i).Cells, Sheet39.Range("Salary.ProfitSection" & i).count)
                If currCntProfit1 > 1 Or currCntProfit2 > 1 Or currCntProfit3 > 1 Then
                    fmsgboxStatus "Drop down has already been selected" ', vbOKOnly, "Error(s)!"
                    Target.value = "(Select)"
        End If
    End If

   If Not Intersect(Range("Salary.Section" & i), Range(Target.Address)) Is Nothing Then
'    Dim ChangeAddress As String
         ChangeAddress = Replace(Target.Address, "G", "H")
    If Sheet39.Range(Target.Address).value = "Others" Then
         Sheet39.Range(ChangeAddress).Locked = False
         Sheet39.Range(ChangeAddress).Interior.Color = "&HCCFFCC"
         Sheet39.Range(ChangeAddress).ClearContents
    Else
         Sheet39.Range(ChangeAddress).Locked = True
         Sheet39.Range(ChangeAddress).Interior.Color = "&HD8D8D8"
         Sheet39.Range(ChangeAddress).ClearContents
    End If
   End If

   If Not Intersect(Range("Salary.ValueSection" & i), Range(Target.Address)) Is Nothing Then
    'Dim ChangeAddress As String
         ChangeAddress = Replace(Target.Address, "G", "H")
    If Sheet39.Range(Target.Address).value = "Other benefits or amenities" Then
         Sheet39.Range(ChangeAddress).Locked = False
         Sheet39.Range(ChangeAddress).Interior.Color = "&HCCFFCC"
         Sheet39.Range(ChangeAddress).ClearContents
    Else
         Sheet39.Range(ChangeAddress).Locked = True
         Sheet39.Range(ChangeAddress).Interior.Color = "&HD8D8D8"
         Sheet39.Range(ChangeAddress).ClearContents
    End If
   End If


   If Not Intersect(Range("Salary.ProfitSection" & i), Range(Target.Address)) Is Nothing Then
    'Dim ChangeAddress As String
         ChangeAddress = Replace(Target.Address, "G", "H")
    If Sheet39.Range(Target.Address).value = "Any Other" Then
         Sheet39.Range(ChangeAddress).Locked = False
         Sheet39.Range(ChangeAddress).Interior.Color = "&HCCFFCC"
         Sheet39.Range(ChangeAddress).ClearContents
    Else
         Sheet39.Range(ChangeAddress).Locked = True
         Sheet39.Range(ChangeAddress).Interior.Color = "&HD8D8D8"
         Sheet39.Range(ChangeAddress).ClearContents
    End If
   End If

   If Not Intersect(Range("SAL.StateCode" & i), Range(Target.Address)) Is Nothing Then
   If Mid(Sheet39.Range("SAL.StateCode" & i).value, 1, 2) = "99" Then
         Sheet39.Range("SAL.ZipCode" & i).Locked = False
         Sheet39.Range("SAL.ZipCode" & i).Interior.Color = "&HCCFFCC"
         Sheet39.Range("SAL.ZipCode" & i).ClearContents
         Sheet39.Range("SAL.PinCode" & i).Locked = True
         Sheet39.Range("SAL.PinCode" & i).Interior.Color = "&HD8D8D8"
         Sheet39.Range("SAL.PinCode" & i).ClearContents
   Else
         Sheet39.Range("SAL.PinCode" & i).Locked = False
         Sheet39.Range("SAL.PinCode" & i).Interior.Color = "&HCCFFCC"
         Sheet39.Range("SAL.PinCode" & i).ClearContents
         Sheet39.Range("SAL.ZipCode" & i).Locked = True
         Sheet39.Range("SAL.ZipCode" & i).Interior.Color = "&HD8D8D8"
         Sheet39.Range("SAL.ZipCode" & i).ClearContents
   End If
   End If

'   If Not Intersect(Range("sheet.EmployerCategory" & i), Range(Target.address)) Is Nothing Then
'   If Sheet39.Range("sheet.EmployerCategory" & i).value = "Pensioners" Or Mid(Sheet39.Range("sheet.EmployerCategory" & i).value, 1, 3) = "OTH" Then
'         Sheet39.Range("SAL.Entertainment_allowance_1").ClearContents
'         Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HD8D8D8"
'         Sheet39.Range("SAL.Entertainment_allowance_1").Locked = True
'
'   Else
'         'Sheet39.Range("SAL.Entertainment_allowance_" & i).ClearContents
'         Sheet39.Range("SAL.Entertainment_allowance_1").Interior.Color = "&HCCFFCC"
'         Sheet39.Range("SAL.Entertainment_allowance_1").Locked = False
'
'   End If
'
'   If Sheet39.Range("sheet.EmployerCategory" & i).value = "Pensioners" Then
'      Sheet39.Range("SAL.DeductionUnderSection1").ClearContents
'      Sheet39.Range("SAL.DeductionUnderSection1").Locked = True
'      Sheet39.Range("SAL.DeductionUnderSection1").Interior.Color = "&HD8D8D8"
'
'Else
'      'Sheet39.Range("SAL.Entertainment_allowance_" & i).ClearContents
'      Sheet39.Range("SAL.DeductionUnderSection1").Interior.Color = "&HCCFFCC"
'      Sheet39.Range("SAL.DeductionUnderSection1").Locked = False
'   End If
'
'End If
Next

  If Target.Validation.Type = 1 Then
    Application.ScreenUpdating = False
    If Intersect(Range("SAL.DeductionUnder6Section1"), Range(Target.Address)) Is Nothing Then
        AddTotalSalary

    End If
    Application.ScreenUpdating = True
End If

   Sheet39.Protect Password:=sPassword










'If Target.name.name = "SAL.DeductionUnder6Section1" Then
            
            
            
            
             If (Range("SAL.DeductionUnder6Section1").value) > 50000 Then
            fmsgboxStatus "Deduction u/s. 16(ia) cannot exceed Rs. 50000 or Net Salary whichever is lower in Schedule Salary."
            Range("SAL.DeductionUnder6Section1").value = 50000
            Range("SAL.DeductionUnder6Section1").Select
             End If

             Dim temp16a, temp16aa As Variant
            temp16a = Sheet39.Range("SAL.DeductionUnder6Section1").value
            temp16aa = Sheet39.Range("SAL.Netsalary1").value


        If temp16a <> "" Then
        If Not Sheet39.Range("SAL.DeductionUnder6Section1").value <= temp16aa Then
               fmsgboxStatus "Deduction u/s. 16(ia) cannot exceed Rs. 50000 or Net Salary whichever is lower in Schedule Salary"
               Range("SAL.DeductionUnder6Section1").value = temp16aa
               Range("SAL.DeductionUnder6Section1").Select


        End If
        End If
        
        If SheetALL.Range("bacValue").value = 1 Then
            Range("SAL.DeductionUnder6Section1").value = 0
            
            End If
        
'End If

 If Target.name.name = "SAL.Entertainment_allowance_1" Or Target.name.name = "SAL.Netsalary1" Then
   Dim temp1 As Variant
       temp1 = 0
       temp1 = Sheet39.Range("Temp_allowance").value / 5

       If Range(Target.Address).value > WorksheetFunction.Min(temp1, 5000) Then
         fmsgboxStatus "Deduction of Entertainment allowance u/s 16(ii) Maximum amount allowable in this field shall be 1/5th of Basic Salary 1a of all employers or Rs 5000 whichever is lower."
         Range("SAL.Entertainment_allowance_1").value = ""
         Range("SAL.Entertainment_allowance_1").Select
       End If
   End If

   If Target.name.name = "SAL.DeductionUnderSection1" Then
   If Range("SAL.DeductionUnderSection1").value > 5000 Then
    fmsgboxStatus "Deduction of Professional tax u/s 16(iii) is Rs. 200/- p.m. in most of the state. Kindly reconsider the amount claimed"
'    Range("SAL.DeductionUnderSection1").value = ""
'    Range("SAL.DeductionUnderSection1").Select
   End If
   End If


   
endline:
Application.EnableEvents = True
End Sub

Sub AddTotalSalary()
    Dim Row7RangeName As String
    Dim i As Variant
    Dim newDefinedName As String
    Dim tot, tot1, tot2, tot3, tot4, tot5, tot6, Totsections, otherTotal, Tempsal, TotalTempGrosssal As Double
    Dim ESOPTotalTempGrosssal As Double
    Dim Deduction80CCD2Total, AllowanceTotal As Double
    Dim Exemptincome As Double

    Tempsal = 0
    tot = 0
    tot1 = 0
    tot2 = 0
    tot3 = 0
    tot4 = 0
    tot5 = 0
    tot6 = 0
    otherTotal = 0
    AllowanceTotal = 0
    Deduction80CCD2Total = 0
    Exemptincome = 0
    TotalTempGrosssal = 0
    ESOPTotalTempGrosssal = 0
    'Row7RangeName = "SAL.IncomeFromSalary"
    Totsections = Sheet39.Range("SalarySectionCount").value
'    Sheet39.Range("HRA_SAL").value = 0

         sPassword = EfilingCommon.getmsgstate
         Sheet39.Unprotect Password:=sPassword

        For i = 1 To Totsections
        'newDefinedName = Row7RangeName & i

       tot2 = WorksheetFunction.Sum(Sheet39.Range("Salary.SalaryAmount" & i))
        Sheet39.Range("SAL.AllowancesNotExempt" & i).value = tot2

         tot3 = WorksheetFunction.Sum(Sheet39.Range("Salary.ValueAmount" & i))
        Sheet39.Range("SAL.ValueOfPerquisites" & i).value = tot3

         tot4 = WorksheetFunction.Sum(Sheet39.Range("Salary.ProfitAmount" & i))
        Sheet39.Range("SAL.ProfitsinLieuOfSalary" & i).value = tot4
        
        tot6 = WorksheetFunction.Sum(Sheet39.Range("IncomeNotified89A_Amount" & i))
        Sheet39.Range("IncomeNotified89A" & i).value = tot6
        
        'Sheet39.Range("SAL.Salary" & i).value = Sheet39.Range("SAL.AllowancesNotExempt" & i).value + Sheet39.Range("SAL.ValueOfPerquisites" & i).value + Sheet39.Range("SAL.ProfitsinLieuOfSalary" & i).value + Sheet39.Range("IncomeNotified89A" & i).value + Sheet39.Range("IncomeNotifiedOther89A" & i).value
        Sheet39.Range("SAL.Salary" & i).value = Sheet39.Range("SAL.AllowancesNotExempt" & i).value + Sheet39.Range("SAL.ValueOfPerquisites" & i).value + Sheet39.Range("SAL.ProfitsinLieuOfSalary" & i).value + Sheet39.Range("IncomeNotified89A" & i).value + Sheet39.Range("IncomeNotifiedOther89A" & i).value + Sheet39.Range("IncomeTaxablePrevYear89A" & i).value



        Tempsal = Tempsal + Sheet39.Range("SAL.Salary" & i).value
       ' Sheet39.Range("SAL.AllowancesNotExempt" & i).formula = "=Sum(range(""Salary.SalaryAmount"" & i))"
        tot = tot + Sheet39.Range("SAL.Salary" & i)
        tot1 = tot1 + Sheet39.Range("SAL.AllowancesNotExempt" & i) + Sheet39.Range("SAL.ProfitsinLieuOfSalary" & i)

'        tot3 = tot3 + Sheet39.Range("Salary.ValueAmount" & i)
'        tot4 = tot4 + Sheet39.Range("Salary.ProfitAmount" & i)

        TotalTempGrosssal = TotalTempGrosssal + Sheet39.Range("SAL.Salary" & i).value
        ESOPTotalTempGrosssal = ESOPTotalTempGrosssal + Sheet39.Range("SAL.ValueOfPerquisitesA" & i).value
        
        Next

Dim BA_SAL, DA_SAL, HouseRentAllowance_count, count1
 BA_SAL = 0
 DA_SAL = 0
HouseRentAllowance_count = 0

 Totsections = Sheet39.Range("SalarySectionCount").value

 For count1 = 1 To Totsections
       BA_SAL = BA_SAL + WorksheetFunction.SumIf(Sheet39.Range("Salary.Section" & count1), "Basic Salary", Sheet39.Range("Salary.SalaryAmount" & count1))

     DA_SAL = DA_SAL + WorksheetFunction.SumIf(Sheet39.Range("Salary.Section" & count1), "Dearness Allowance(DA)", Sheet39.Range("Salary.SalaryAmount" & count1))
 
  HouseRentAllowance_count = HouseRentAllowance_count + WorksheetFunction.SumIf(Sheet39.Range("Salary.Section" & count1), "House Rent Allowance(HRA)", Sheet39.Range("Salary.SalaryAmount" & count1))

    
Next

Sheet39.Range("SAL_Basic").value = BA_SAL
Sheet39.Range("SAL_DA").value = DA_SAL
Sheet39.Range("HRA_SAL").value = HouseRentAllowance_count



'    If Sheet39.Range("SAL.Netsalary1").value > 50000 Then
'    Sheet39.Range("SAL.DeductionUnder6Section1").value = 50000
'    Else
'    Sheet39.Range("SAL.DeductionUnder6Section1").value = Sheet39.Range("SAL.Netsalary1").value
'    End If
'    If SheetALL.Range("bacValue").value = 1 Then
'            Range("SAL.DeductionUnder6Section1").value = 0
'
'            End If
    
    Sheet39.Protect Password:=sPassword
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet39.Unprotect Password:=sPassword
   Sheet7.Unprotect Password:=sPassword
   Sheet19.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    Sheet39.Range("Tempgross").value = tot
    Sheet39.Range("Tempgross1").value = tot1
    Sheet7.Range("Sheet8b.Salaries").value = Sheet39.Range("SAL.IncomeFromSalary1").value
    
    Sheet39.Range("SAL.TotalGrossSalary").value = TotalTempGrosssal
    Sheet39.Range("SAL.TotESOP").value = ESOPTotalTempGrosssal
    tot5 = tot1 - Sheet39.Range("SAL.ExemptUSectionOth1").value
    Sheet19.Range("TEMP80CCD1").value = WorksheetFunction.Max(0, tot5)

    'Sheet19.Range("Salaryded").value = Deduction80CCD2Total
    'Sheet39.Range("SAL.Salary1_Temp").value = Tempsal
'----------------Lock Password-------------------START---
   Sheet7.Protect Password:=sPassword
   Sheet39.Protect Password:=sPassword
   Sheet19.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
End Sub

Private Sub Worksheet_Deactivate()

'Change-04.05.2023.101. -> V0.4 Implementation
CalcVIA80CCH
'----
'AY_2024_25'Malli
CalcVIA80CCH_NewRegme


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



