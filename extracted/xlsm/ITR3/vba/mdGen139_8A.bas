Attribute VB_Name = "mdGen139_8A"
'SUB FOR 1398A
Public msgbox_1398A As String
Public end1398A_1, end1398A_2 As Long
Public rngname_1398A_1, rngname_1398A_2 As Variant
Public dateorigfiled As Variant
Public Systemdate2324 As Variant
Function checkfieldSuperSpecialcharacterDot_1398A(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldSuperSpecialcharacterDot_1398A = True
    Dim arr As Variant
    arr = Array(".")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldSuperSpecialcharacterDot_1398A = False
            Exit Function
        End If
        Next
    Next
End Function

Public Sub fmsgboxStatus_1398A(IntMsg1 As Variant)
'With MessageBox
'    .Height = 130
'    .LMessagebox.Height = 60
'    .CommandButton1.Top = 73
'    .LMessagebox = "*" + IntMsg1
'    .LMessagebox.TextAlign = fmTextAlignLeft
'    .Show
'End With
MsgBox IntMsg1
End Sub
Public Sub fmsgboxsmall_1398A(IntMsg2 As Variant)
'With MessageBox
'     .Height = 255
'     .LMessagebox.Height = 174
'     .CommandButton1.Top = 191
'     .LMessagebox = IntMsg2
'     .LMessagebox.TextAlign = fmTextAlignLeft
'     .Show
'End With
MsgBox IntMsg2
End Sub

Sub CmdValidate_Gen_1398A_Click()
Validate_Gen_1398A
fmsgboxStatus_1398A "Sheet Part A Gen_139(8A) is OK"
CloseMsg
End Sub
Sub Validate_Gen_1398A()
subProcCaption = "Validating Part A Gen_139(8A)"

If Not ValidateSheetGen1398A Then
fmsgboxsmall_1398A (msgbox_1398A)
CloseMsg
End If
End Sub
Sub setTableInfoGen1398A_1()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet201.Range("U_UnabsorbedDepreciationYear").Cells
    mIntCells = Sheet201.Range("U_UnabsorbedDepreciationYear").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end1398A_1 = ccount
    rngname_1398A_1 = "U_UnabsorbedDepreciationYear"
End Sub
Sub setTableInfoGen1398A_2()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet201.Range("U_ReasonsForUpdatingIncome").Cells
    mIntCells = Sheet201.Range("U_ReasonsForUpdatingIncome").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    end1398A_2 = ccount
    rngname_1398A_2 = "U_ReasonsForUpdatingIncome"
End Sub

Function ValidateSheetGen1398A() As Boolean

ValidateSheetGen1398A = True
msgbox_1398A = "Part A Gen_139(8A)" & Chr(10)
setTableInfoGen1398A_1
If Mid(Sheet1.Range("sheet1.ReturnFileSec").value, 1, 7) = "139(8A)" Then
    
    If Sheet201.Range("U_UnabsorbedDepreciationYear").Cells(1).value = "2023-24" Then
        If Sheet201.Range("U_RevisedReturnFile").Cells(1).value = "" Or Sheet201.Range("U_RevisedReturnFile").Cells(1).value = "(Select)" Then
             msgbox_1398A = msgbox_1398A + "* Please Select wether revised Return has been filed for this AY" & Chr(13)
            ValidateSheetGen1398A = False
        End If
        
    End If
    
    If Sheet201.Range("U_PreviouslyFiledForThisAY").value = "(Select)" Or Sheet201.Range("U_PreviouslyFiledForThisAY").value = "" Then
    msgbox_1398A = msgbox_1398A + "* Please select from the drop downs given at field 'Whether return previously filed for this assessment year?'" & Chr(13)
    ValidateSheetGen1398A = False
'Change 102 28.11.2023 Ver.0.6 SIT-62971
'    Else
'    ElseIf Sheet201.Range("U_PreviouslyFiledForThisAY").value = "No" And Sheet1.Range("OptingNewTaxRegime").value = "Opting in now" Then
'    msgbox_1398A = msgbox_1398A + "* 'Opting in now' cannot be selected for ' Option for Current Assessment year' in Part A General if Sl.no A5 in Part A General 139(8A) is selected as ""No""." & Chr(13)
'    ValidateSheetGen1398A = False
'End Change
    End If
    
    If Trim(Sheet201.Range("U_Name").value) = "" Then
    msgbox_1398A = msgbox_1398A + "* Name is mendatory" & Chr(13)
    ValidateSheetGen1398A = False
    End If
    
    If Sheet201.Range("U_PAN").value = "" Then
    msgbox_1398A = msgbox_1398A + "* PAN is mendatory" & Chr(13)
    ValidateSheetGen1398A = False
    End If
    
    If Sheet201.Range("U_PreviouslyFiledForThisAY").value = "Yes" Then
        If Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").value = "(Select)" Or Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").value = "" Then
'Change Satya 102, 31/10/2023, ver.0.6
'        msgbox_1398A = msgbox_1398A + "* Please select from the drop downs given at field 'If Yes, Whether U/s'" & Chr(13)
        msgbox_1398A = msgbox_1398A + "* ""A6 is mandatorily required to be selected""" & Chr(13)
        ValidateSheetGen1398A = False
        End If
    
    ElseIf Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").Locked = False Then
        If Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").value = "(Select)" Or Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").value = "" Then
        msgbox_1398A = msgbox_1398A + "* Please select from the drop downs given at field 'If Yes, Whether U/s'" & Chr(13)
        
        ValidateSheetGen1398A = False
        End If
'End Change
    End If
    
    If Sheet201.Range("U_PreviouslyFiledForThisAY").value = "Yes" Then
    If Sheet201.Range("U_ITRForm").value = "(Select)" Or Sheet201.Range("U_ITRForm").value = "" Then
'Change Satya 102, 31/10/2023, ver.0.6
'    msgbox_1398A = msgbox_1398A + "* Please select from the drop downs given at field 'ITR Type'" & Chr(13)
    msgbox_1398A = msgbox_1398A + """Please select from the drop downs given in ITR Type""" & Chr(13)
'End Change
    ValidateSheetGen1398A = False
    End If
    End If
    
    

    If Sheet201.Range("U_LaidOutIn_139_8A").value = "(Select)" Or Sheet201.Range("U_LaidOutIn_139_8A").value = "(Select)" Then
    msgbox_1398A = msgbox_1398A + "* Please select from the drop downs given at field 'Are you eligible for filing an updated return as per the conditions laid out in first, second and third and fourth provisos to section 139(8A)?' " & Chr(13)
    ValidateSheetGen1398A = False
    End If
    
'    If Sheet201.Range("U_PreviouslyFiledForThisAY").Value = "Yes" Then
'    If Sheet201.Range("U_ITRForm").Value = "(Select)" Then
'    msgbox_1398A = msgbox_1398A + "* Please select ITR Type from the drop downs given  " & Chr(13)
'    ValidateSheetGen1398A = False
'    End If
'    End If
    
    If Sheet201.Range("U_LaidOutIn_139_8A").value = "No" Then
    msgbox_1398A = msgbox_1398A + "*Dear user, you will not be able to proceed to file Updated Return u/s 139(8A) as you are not eligible for filing an updated return as per the conditions laid out in first, second and third provisos to section 139(8A)" & Chr(13)
    ValidateSheetGen1398A = False
    End If
    
    
    Systemdate2324 = Sheet52.Range("sheet9.Date").value
    If Sheet201.Range("U_UpdatedReturnDuringPeriod").value = "(Select)" Or Sheet201.Range("U_UpdatedReturnDuringPeriod").value = "" Then
    msgbox_1398A = msgbox_1398A + "* Please select from the the drop downs given at field 'Are you filing the updated return during the period'" & Chr(13)
    ValidateSheetGen1398A = False
'Change Satya 102, 31/10/2023, ver.0.6
    Else
    'Uncommented by sai on 18/11/2024
        If Not ChkMinSystemdate2324(Systemdate2324, "01/04/2026") Then
           If Sheet201.Range("U_UpdatedReturnDuringPeriod").value = "Between 12 to 24 Months from the end of Relevant Assessment  Year" Then

           msgbox_1398A = msgbox_1398A + "* Please select appropriate option in 'Are you filing the updated return during the period.'" & Chr(13)
           ValidateSheetGen1398A = False
           End If
        Else
'           If Sheet201.Range("U_UpdatedReturnDuringPeriod").value = "Up to 12 months from the end of Relevant Assessment Year" Then
'
'           msgbox_1398A = msgbox_1398A + "* Please select appropriate option in 'Are you filing the updated return during the period.'" & Chr(13)
'           ValidateSheetGen1398A = False
'           End If
        End If
'End Change


    End If
    
    If Sheet201.Range("U_UnabsorbedDepreciation").value = "(Select)" Or Sheet201.Range("U_UnabsorbedDepreciation").value = "" Then
    If Sheet201.Range("Reason1").value = 1 Or Sheet201.Range("Reason2").value = 1 Or Sheet201.Range("Reason3").value = 1 Then
   ' msgbox_1398A = msgbox_1398A + "* Please select from the drop downs given at field 'Are you filing the updated return to reduce carried forward loss or unabsorbed depreciation or tax credit?'" & Chr(13)
    msgbox_1398A = msgbox_1398A + "* ""Please select from the radio buttons given at field ""Are you filing the updated return to reduce carried forward loss or unabsorbed depreciation or tax credit?""""" & Chr(13) 'Added by Shrutika(03/12/2025)UR
 
    ValidateSheetGen1398A = False
    End If
    End If
    
    If Not ValidateGen1398A_1 Then ValidateSheetGen1398A = False
    If Not ValidateGen1398A_2 Then ValidateSheetGen1398A = False
    
'    If Sheet201.Range("U_RevisedReturnFile").Value = "(Select)" Then
'    msgbox_1398A = msgbox_1398A + "* Please select from the drop downs given at 'Whether revised return has been filed for the AY'" & Chr(13)
'    ValidateSheetGen1398A = False
'    End If
'
'
'    If Sheet201.Range("U_UpdatedReturnFile").Value = "(Select)" Then
'    msgbox_1398A = msgbox_1398A + "* Please select from the drop downs given at 'Whether updated return is filed for the AY'" & Chr(13)
'    ValidateSheetGen1398A = False
'    End If
'
    
    
'    dateorigfiled = Sheet201.Range("U_OrigRetFiledDate").Value
'    If Not FormatNCheckDate_1398A(dateorigfiled) Then
'    ValidateSheetGen1398A = False
'    msgbox_1398AATI = msgbox_1398AATI + "* dddDate of Deposit into Govt Account must be a valid dd/mm/yyyy format at Sr.NO " & i & " in schedule IT-2" & Chr(13)
'    End If
    
    If Sheet201.Range("U_UnabsorbedDepreciation").value <> "Yes" Then
    If Sheet201.Range("U_UnabsorbedDepreciationYear_1").value >= 1 Or Sheet201.Range("U_UnabsorbedDepreciationYear_2").value >= 1 Then
    msgbox_1398A = msgbox_1398A + "* Year you can not select if updated return is not to reduce carried forward loss or unabsorbed depreciation or tax credit" & Chr(13)
    
    ValidateSheetGen1398A = False
    End If
    End If
End If

End Function

Function ValidateGen1398A_1()
ValidateGen1398A_1 = True



   setTableInfoGen1398A_1
    Dim rangecells As Range
    Set rangecells = Sheet201.Range("U_UnabsorbedDepreciationYear").Cells
    Dim i As Long
    ReDim U_UnabsorbedDepreciationYear(end1398A_1)
    
    subProcCaption = "Validating Part A Gen_139(8)"
    noOfProcessSub = end1398A_1
    If Sheet201.Range("U_UnabsorbedDepreciation").value = "Yes" Then
        If end1398A_1 = 0 Then
            'msgbox_1398A = msgbox_1398A + "* Please select from the AY from the dropdown provided at sl.no.(b)" & Chr(13)
            
            'Added by Aavula Naresh on 11/11/2024
            
            msgbox_1398A = msgbox_1398A + "* ""Please select from the AY from the dropdown provided at sl.no.A12(b)""" & Chr(13)
            
            ''''
            ValidateGen1398A_1 = False
            Exit Function
        End If
         If end1398A_1 > 0 Then
         Dim Year(2)
          Dim ReturnYesNo(2)
          Dim ReturnFile(2)



           For i = 1 To end1398A_1
                Year(i) = Sheet201.Range("U_UnabsorbedDepreciationYear").Cells(i, 1).value
                ReturnYesNo(i) = Sheet201.Range("U_RevisedReturnFile").Cells(i, 1).value
                ReturnFile(i) = Sheet201.Range("U_UpdatedReturnFile").Cells(i, 1).value
                If (Year(i) <> "") Then
                     If (ReturnYesNo(i) = "" Or ReturnYesNo(i) = "(Select)") Then
                        'msgbox_1398A = msgbox_1398A + "Please select Whether return has been filed at Sr No." & i & Chr(13)
                        
                        'Added by Aavula Naresh on 11/11/2024
                        'msgbox_1398A = msgbox_1398A + """Please select Whether return has been filed at sl.no.A12(c)"" in Sr No." & i & Chr(13)
'added by Konda

                         msgbox_1398A = msgbox_1398A + "* ""Please select a drop down at Sr No. " & i & " in ""Whether original/revised return has been filed"" for the AY at A12(c)""" & Chr(13)
                        
                        
                        ValidateGen1398A_1 = False
                    End If
                   ' If (ReturnFile(i) = "" Or ReturnFile(i) = "(Select)") And (ReturnYesNo(i) <> "No") Then
                    If (ReturnFile(i) = "" Or ReturnFile(i) = "(Select)") Then
                       ' msgbox_1398A = msgbox_1398A + "Please select Return Filed, if yes is selected at Sr No." & i & Chr(13)
                        
                        'Added by Aavula Naresh on 11/11/2024
                        
                        'msgbox_1398A = msgbox_1398A + """Please select Return Filed, if yes is selected at sl.no.A12(d)"" in Sr No." & i & Chr(13)

'added by Konda
                    
                    msgbox_1398A = msgbox_1398A + "* ""Please select a drop down at Sr No. " & i & " in ""Whether updated return has been filed"" for the AY at A12(d)""" & Chr(13)
                        ValidateGen1398A_1 = False
                    End If
                End If
              Next
         End If
    End If
End Function
Function FormatNCheckDate_1398A(ByRef dt As Variant, Optional Year As String = "2021") As Boolean
On Error Resume Next
    FormatNCheckDate_1398A = True
    If Len(dt) > 0 Then
        'Format the date in dd/mm/yyyy format
        If Mid(dt, 3, 1) <> "/" Then 'Checking between dd/mm
            If Mid(dt, 3, 1) <> "\" Or Mid(dt, 3, 1) <> "-" Or Mid(dt, 3, 1) <> "." Or Mid(dt, 3, 1) <> "," Then
                dt = Mid(dt, 1, 2) & "/" & Mid(dt, 4, 7)
            End If
        End If
        
        If Mid(dt, 6, 1) <> "/" Then 'Checking between mm/yyyy
            If Mid(dt, 6, 1) <> "\" Or Mid(dt, 6, 1) <> "-" Or Mid(dt, 6, 1) <> "." Or Mid(dt, 6, 1) <> "," Then
                dt = Mid(dt, 1, 5) & "/" & Mid(dt, 7, 4)
            End If
        End If
        
        'Checking Date if it is in correct format :Day(dd), Month(mm) & Year(yyyy)
        If val(Mid(dt, 1, 2)) < 0 Or val(Mid(dt, 1, 2)) > 31 Then FormatNCheckDate_1398A = False
        If val(Mid(dt, 4, 2)) < 0 Or val(Mid(dt, 4, 2)) > 12 Then FormatNCheckDate_1398A = False
        If Not IsDate(dt) Then FormatNCheckDate_1398A = False
        
    End If
    
End Function
Function ValidateGen1398A_2()
ValidateGen1398A_2 = True

    setTableInfoGen1398A_2
    Dim rangecells As Range
    Set rangecells = Sheet201.Range("U_ReasonsForUpdatingIncome").Cells
    Dim i As Long
    ReDim U_UnabsorbedDepreciationYear(end1398A_2)
    
    subProcCaption = "Validating Part A Gen_139(8)"
    noOfProcessSub = end1398A_2

        If end1398A_2 = 0 Then
            msgbox_1398A = msgbox_1398A + "* Please select Reason from the drodpown " & Chr(13)
            ValidateGen1398A_2 = False
            Exit Function
        End If
        
        If Sheet201.Range("U_PreviouslyFiledForThisAY").value = "Yes" Then
          If Sheet201.Range("U_AcknowledgementNo").value = "" Or Sheet201.Range("U_OrigRetFiledDate").value = "" Then
          msgbox_1398A = msgbox_1398A + "* Date of filing and acknowledgement no. of last valid return is mandatory." & Chr(13)
          ValidateGen1398A_2 = False
          End If
        End If



        If Len(Sheet201.Range("U_OrigRetFiledDate").value) > 0 Then
'          If Not FormatNCheckDate_1398A(Sheet201.Range("U_OrigRetFiledDate").value) Then
           If Not CheckDateddmmyyyy(Sheet201.Range("U_OrigRetFiledDate").value) Then 'by sai on 16/11/2024
              msgbox_1398A = msgbox_1398A + "* Date of filing of Original return must be a valid dd/mm/yyyy format" & Chr(13)
              ValidateGen1398A_2 = False
           End If
        End If
End Function
Sub Cmd_AddRows_Gen138A_1_Click()
Dim vRows As Long
EfilingCommon.DefinedgridNameRange = ("U_ReasonsForUpdatingIncome")
ActiveCellRange = EfilingCommon.searchLastRow("U_ReasonsForUpdatingIncome")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub Cmd_AddRows_Gen138A_2_Click()
Dim vRows As Long
EfilingCommon.DefinedgridNameRange = ("U_UnabsorbedDepreciationYear||U_RevisedReturnFile||U_UpdatedReturnFile")
ActiveCellRange = EfilingCommon.searchLastRow("U_UnabsorbedDepreciationYear")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub



