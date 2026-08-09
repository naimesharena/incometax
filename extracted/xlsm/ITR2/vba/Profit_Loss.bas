Attribute VB_Name = "Profit_Loss"
Option Explicit

Dim MsgBx_NoActPL As String
Dim MsgBx_names As Variant
'Dim MsgBx_names As Variant
Dim MsgBx_1aiii As String
Dim MsgBx_2otherInc As String, MsgBx_14xii As String, MsgBx_38 As String
Dim MsgBx_compPaid As String
Dim MsgBx_BadDept39 As String
Dim MsgBx_OtherPan As String
Dim MsgBx_GST As String

Public ActiveCellRange As String

Dim ColCount_Nature As Long, ColCount_Nature2 As Long, ColCount_Nature3 As Long
Dim ColCount_Amt As Long, ColCount_Amt2 As Long, ColCount_Amt3 As Long
Dim ColCount_PAN39 As Long, ColCount_Aadhar39 As Long, ColCount_Amt39 As Long

Public Pan39_PL As Variant, Aadhar39_PL, Amt39_PL As Variant

Public ColCount_OtherPAN1, ColCount_OtherPAN2, ColCount_OtherPAN3, ColCount_OtherPAN4, ColCount_OtherPAN5, ColCount_OtherPAN6, ColCount_OtherPAN7, ColCount_OtherPAN8, ColCount_OtherPAN9, ColCount_OtherPAN10, ColCount_OtherPAN11 As Variant
Public Pl_OtherPanName, Pl_OtherPanAmount, PL_Address_Flat, PL_Address_Village, PL_Address_Road, PL_Address_Area, PL_Address_City, PL_Address_State, PL_Address_Country, PL_Address_Pin, PL_Address_Zip As Variant
Public Nature As Variant
Public Amount As Variant
Public Nature2 As Variant
Public amount2 As Variant
Public Nature3 As Variant
Public amount3 As Variant

Public errorMsg_NOB1, errorMsg_NOB2, errorMsg_NOB3, errorMsg_NOB5, errorMsg_NOB6 As Variant
Public end_NOBC, end_NOBName, end_NOBDescription, NOBCode, NOBName, NOBDescription As Variant
Public end_NOBC1, end_NOBName1, end_NOBDescription1, NOBCode1, NOBName1, NOBDescription1 As Variant
Public end_NOBC2, end_NOBName2, end_NOBDescription2, NOBCode2, NOBName2, NOBDescription2 As Variant
Public end44AE_RegNo, end44AE_OLH, end44AE_TC, end44AE_Months, end44AE_PIncome, SEC44AERegNo, SEC44AEOLH, SEC44AETC, SEC44AEMonth, SEC44AEPIncome, Greater12_BP, Lesser12_BP As Variant

Public TempPL_61i, TempPL_61a, TempPL_61b, TempPL_61ii, TempPL_61iia, TempPL_61iib, TempPL_62i, TempPL_62ii, TempPL_63ii, TempPL_64GrossTurnOver, TempPL_64ia1, TempPL_64ia2, TempPL_64ib, TempPL_64ic, TempPL_64id, TempPL_64iia, TempPL_64iia_i, TempPL_64iia_ii, TempPL_64iib, TempPL_64iic, TempPL_64iid, TempPL_64iii, TempPL_65i, TempPL_65ii, TempPL_65iii, TempPL_65iv As Variant
Sub CmdValidate_PL_Click()
    ValidateSheet_PL
    fmsgboxStatus "Sheet Profit Loss is ok"
CloseMsg
End Sub

Sub AddRows_44AE()
    Sheet4.Activate
    EfilingCommon.DefinedgridNameRange = "Sec44AE.RegNo||Sec44AE.OLH||Sec44AE.TonnageCpct||Sheet44AE.NoOfMonths||Sheet44AE.PresumptiveIncome||greater12||lesser12"
    ActiveCellRange = EfilingCommon.searchLastRow("Sec44AE.RegNo")
    EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub Cmd_AddRows_PAN_Click()
Dim vRows As Long
Sheet4.Activate
EfilingCommon.DefinedgridNameRange = "PLBD.PAN||PLBD.Aadhar||PLBD.Amount"
ActiveCellRange = EfilingCommon.searchLastRow("PLBD.PAN")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub Cmd_AddRows_OtherPAN_Click()
Dim vRows As Long
Sheet4.Activate
EfilingCommon.DefinedgridNameRange = "PL_Name||PL_Address_Flat||PL_Address_Village||PL_Address_Road||PL_Address_Area||PL_Address_City||PL_Address_State||PL_Address_Country||PL_Address_Pin||PL_Address_Zip||PL_Amount"
ActiveCellRange = EfilingCommon.searchLastRow("PL_Name")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub


Sub Cmd_AddRows_Nature2_Click()
Dim vRows As Long
Sheet4.Activate
EfilingCommon.DefinedgridNameRange = "Nature_Name2||Nature_Amt2"
ActiveCellRange = EfilingCommon.searchLastRow("Nature_Name2")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub Cmd_AddRows_Nature3_Click()
Dim vRows As Long
Sheet4.Activate
EfilingCommon.DefinedgridNameRange = "Nature_Name3||Nature_Amt3"
ActiveCellRange = EfilingCommon.searchLastRow("Nature_Name3")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub


Sub ValidateSheet_PL()

'check for Nature and Amount 1A iii
subProcCaption = "Validating PL"
noOfProcessSub = 7

    UpdateProgressBar
    
'Check for nature and amount in 2
If Not ValidateNatureAmt2 Then
    Sheet4.Activate
    fmsgboxsmall (MsgBx_2otherInc)
    CloseMsg
End If
UpdateProgressBar

'check for
If Not validatecompPaid Then
    Sheet4.Activate
    fmsgboxsmall (MsgBx_compPaid)
    CloseMsg
End If
UpdateProgressBar
'Check for NAture AMount38
If Not ValidateNAmtureAmt_38 Then
    Sheet4.Activate
    fmsgboxsmall (MsgBx_38)
    CloseMsg
End If
UpdateProgressBar
'Check for PAn Amount 39
If Not ValidateBad_Dept_39 Then
    Sheet4.Activate
    fmsgboxsmall (MsgBx_BadDept39)
    CloseMsg
End If
UpdateProgressBar


If Not ValidateOtherPan Then
    Sheet4.Activate
    fmsgboxsmall (MsgBx_OtherPan)
    CloseMsg
End If
UpdateProgressBar
    
'Check For filling of Both 52 and 53
'    If Not ValidateNoAct_PL Then
'        Sheet4.Activate
'        MsgBox (MsgBx_NoActPL)
'        'CloseMsg
'    End If
    
    If Not checkNames Then
        Sheet4.Activate
        fmsgboxsmall (MsgBx_names)
        CloseMsg
    End If
    
    
    If Not ValidateGST Then
    Sheet4.Activate
    fmsgboxsmall (MsgBx_GST)
    CloseMsg
    End If
    
    ValidateNOB1
    ValidateNOB2
    ValidateNOB3
    ValidateSec44AE
    ValidateCinditionalcheck
If errorMsg_NOB6 <> "" Then
fmsgbox errorMsg_NOB6
End
End If
    
    UpdateProgressBar
'Check for value greater than 1 Crore
    'ValidateTotRevenue
     UpdateProgressBar


End Sub
Function ValidateGST() As Boolean
ValidateGST = True
End Function

Function ValidateTotRevenue() As Boolean
If Range("PL.TotRevenueFrmOperations").value > 10000000 Then
    fmsgboxStatus "Sales /Gross Receipts(1 of Part A-P & L) of Business & Profession is greater than 1 crore then AUDIT INFORMATION must be completely filled"
End If
End Function

'Function ValidateNoAct_PL() As Boolean

Function checkNames() As Boolean
    checkNames = True
Dim rangecell As Range
Dim tar As Variant

Set rangecell = Sheet4.Range("Nature_Name2").Cells
For Each tar In rangecell.value
    If Not checkfieldspecialcharacter(tar) Then
        checkNames = False
        MsgBx_names = MsgBx_names + "* Nature in 2x cannot contain special charachters <> ' '&' " '
        Exit Function
    End If
Next

Set rangecell = Sheet4.Range("Nature_Name3").Cells
For Each tar In rangecell.value
    If Not checkfieldspecialcharacter(tar) Then
        checkNames = False
        MsgBx_names = MsgBx_names + "* Nature in 38 cannot contain special charachters <> ' '&' " '
        Exit Function
    End If
Next
End Function

Function ValidateNatureAmt2() As Boolean
ValidateNatureAmt2 = True
    countfilledCol_Amt2
    countfilledCol_Nature2
    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Sheet4.Range("Nature_Name2").Cells
    Set rangecells1 = Sheet4.Range("Nature_Amt2").Cells
    Dim i As Long
    ReDim Nature2(ColCount_Nature2)
    ReDim amount2(ColCount_Nature2)
    MsgBx_2otherInc = ""
If (ColCount_Nature2 <> ColCount_Amt2) Then
    ValidateNatureAmt2 = False
    If (ColCount_Nature2 < ColCount_Amt2) Then
       ' Sheet4.Range("Nature_Name2").Select
        MsgBx_2otherInc = "Please specify the nature in Any Other Income in schedule Profit and Loss A/c"
    Else:
        'Sheet4.Range("Nature_Amt2").Select
        MsgBx_2otherInc = "Please enter the amount in Any Other Income in schedule Profit and Loss A/c"
    End If
End If

If ColCount_Nature2 > 0 Then

For i = 1 To ColCount_Nature2
rangecells.item(i).value = UCase(rangecells.item(i).value)
rangecells1.item(i).value = UCase(rangecells1.item(i).value)
Nature2(i) = rangecells.item(i).value
amount2(i) = rangecells1.item(i).value
Next

End If

End Function


Sub countfilledCol_Amt2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet4.Range("Nature_Amt2").count
    Set rangecells = Sheet4.Range("Nature_Amt2").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount_Amt2 = ccount
End Sub


Sub countfilledCol_Nature2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet4.Range("Nature_Name2").count
    Set rangecells = Sheet4.Range("Nature_Name2").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount_Nature2 = ccount
End Sub

Function validatecompPaid() As Boolean
validatecompPaid = True

    If Sheet4.Range("PL.TotEmployeeComp").value > 0 Then
        If isdropdownblank(Sheet4.Range("PL.AnyCompPaidToNonRes")) Then
               MsgBx_compPaid = "Please select Whether any compensation paid to non resident in Sheet : Profit and Loss" & Chr(13)
               validatecompPaid = False
               Exit Function
        End If
    End If
    
    If Mid(Sheet4.Range("PL.AnyCompPaidToNonRes").value, 1, 1) = "Y" Then
        If Sheet4.Range("PL.AmtPaidToNonRes").value = "" Or Sheet4.Range("PL.AmtPaidToNonRes").value <= 0 Then
            MsgBx_compPaid = "Please enter the amount paid as compensation to the non resident" & Chr(13)
                validatecompPaid = False
                Exit Function
        End If
    End If
    
End Function

Function ValidateNAmtureAmt_38() As Boolean
ValidateNAmtureAmt_38 = True
    countfilledCol_Amt3
    countfilledCol_Nature3
    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Sheet4.Range("Nature_Name3").Cells
    Set rangecells1 = Sheet4.Range("Nature_Amt3").Cells
    Dim i As Long
    ReDim Nature3(ColCount_Nature3)
    ReDim amount3(ColCount_Nature3)
    MsgBx_2otherInc = ""
    
If (ColCount_Nature3 <> ColCount_Amt3) Then
    ValidateNAmtureAmt_38 = False
    If (ColCount_Nature3 < ColCount_Amt3) Then
        Sheet4.Range("Nature_Name3").Select
        MsgBx_38 = "Please specify the nature in Any Other Expenses in schedule Profit and Loss A/c"
    Else:
        Sheet4.Range("Nature_Amt3").Select
        MsgBx_38 = "Please enter the amount in Any Other Expenses in schedule Profit and Loss A/c"
    End If
End If

If ColCount_Nature3 > 0 Then

For i = 1 To ColCount_Nature3
rangecells.item(i).value = UCase(rangecells.item(i).value)
rangecells1.item(i).value = UCase(rangecells1.item(i).value)
Nature3(i) = rangecells.item(i).value
amount3(i) = rangecells1.item(i).value
Next

End If


End Function



Sub countfilledCol_Amt3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet4.Range("Nature_Amt3").count
    Set rangecells = Sheet4.Range("Nature_Amt3").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount_Amt3 = ccount
End Sub


Sub countfilledCol_Nature3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet4.Range("Nature_Name3").count
    Set rangecells = Sheet4.Range("Nature_Name3").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount_Nature3 = ccount
End Sub


'Validation for Item 39 Bad Item
Function ValidateBad_Dept_39() As Boolean
ValidateBad_Dept_39 = True
MsgBx_BadDept39 = ""
    If Not ValidatePAn_39 Then ValidateBad_Dept_39 = False
    If Not ValidateAadhar_39 Then ValidateBad_Dept_39 = False
    If Not ValidateAmt_39 Then ValidateBad_Dept_39 = False
        setTableAmt39_PL
    Dim rangecells As Range
    Set rangecells = Sheet4.Range("PLBD.PAN").Cells
     Dim rangeAdharcells As Range
    Set rangeAdharcells = Sheet4.Range("PLBD.Aadhar").Cells
    Dim rangeAmountcells As Range
    Set rangeAmountcells = Sheet4.Range("PLBD.Amount").Cells
    Dim i As Long
    ReDim Pan39_PL(ColCount_PAN39)
    For i = 1 To Sheet4.Range("PLBD.PAN").count
        Dim panCheck As Boolean
        panCheck = True
        Dim adharCheck As Boolean
        adharCheck = True
        Dim amountCheck As Boolean
        Dim amountCheck1L As Boolean
        amountCheck = True
        amountCheck1L = False
        If Trim(rangecells.item(i).value) = "" Or IsEmpty(rangecells.item(i).value) Then
            panCheck = False
        End If
         If Trim(rangeAdharcells.item(i).value) = "" Or IsEmpty(rangeAdharcells.item(i).value) Then
            adharCheck = False
        End If
        If Trim(rangeAmountcells.item(i).value) = "" Or IsEmpty(rangeAmountcells.item(i).value) Then
            amountCheck = False
        End If
        If amountCheck Then
        If (rangeAmountcells.item(i).value) >= 100000 Then
        amountCheck1L = True
        End If
        End If
        
        If panCheck And Not amountCheck Then
            Sheet4.Range("PLBD.Amount").Select
            MsgBx_BadDept39 = MsgBx_BadDept39 + "Please Specify the Amt"
             ValidateBad_Dept_39 = False
            Exit Function
        End If
        If adharCheck And Not amountCheck Then
            Sheet4.Range("PLBD.Amount").Select
            MsgBx_BadDept39 = MsgBx_BadDept39 + "Please Specify the Amt"
             ValidateBad_Dept_39 = False
            Exit Function
        End If
        If amountCheck Then
        
            If Not panCheck And Not adharCheck Then
                Sheet4.Range("PLBD.PAN").Select
                Sheet4.Range("PLBD.Aadhar").Select
                MsgBx_BadDept39 = MsgBx_BadDept39 + "Please enter a PAN/ Aadhaar in schedule Profit and Loss A/c"
                ValidateBad_Dept_39 = False
                Exit Function
            End If
        End If
        If amountCheck1L Then
            If Not panCheck Then
                Sheet4.Range("PLBD.PAN").Select
                MsgBx_BadDept39 = MsgBx_BadDept39 + "Please enter PAN in schedule Profit and Loss A/c in Sr No 47i "
                ValidateBad_Dept_39 = False
                Exit Function
            End If
        End If
    Next
'If (ColCount_PAN39 <> ColCount_Amt39) Then
    '    ValidateBad_Dept_39 = False
   ' If (ColCount_PAN39 < ColCount_Amt39) Then
    '    Sheet4.Range("PLBD.PAN").Select
    '    MsgBx_BadDept39 = "Please enter a PAN in schedule Profit and Loss A/c"
   ' Else:
    '    Sheet4.Range("PLBD.Amount").Select
    '    MsgBx_BadDept39 = "Please Specify the Amt"
    'End If
'End If
        
    

End Function
Function ValidatePAn_39() As Boolean
ValidatePAn_39 = True
    setTablepan_39
    Dim rangecells As Range
    Set rangecells = Sheet4.Range("PLBD.PAN").Cells
    Dim i As Long
    ReDim Pan39_PL(ColCount_PAN39)
    For i = 1 To ColCount_PAN39
    rangecells.item(i).value = UCase(rangecells.item(i).value)
        Pan39_PL(i) = rangecells.item(i).value
        If Not chkCompulsory(Pan39_PL(i)) Then
        
            MsgBx_BadDept39 = MsgBx_BadDept39 + "* Please enter a PAN in schedule Profit and Loss A/c" & Chr(13)
            ValidatePAn_39 = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(Pan39_PL(i)) Then
             MsgBx_BadDept39 = MsgBx_BadDept39 + "* PAN No at Sr. No  " & i & " Item 47 ::Sheet PL Cannot Contain Special Characters" & Chr(13)
             ValidatePAn_39 = False
             Exit Function
        End If
        
        If Not CheckDoneePAN(UCase(Pan39_PL(i))) Then
             MsgBx_BadDept39 = MsgBx_BadDept39 + "* Please enter a valid 10 character PAN at Sr. No  " & i & "in schedule Profit and Loss A/c" & Chr(13)
             ValidatePAn_39 = False
             Exit Function
        End If
    Next
End Function
Function ValidateAadhar_39() As Boolean
ValidateAadhar_39 = True
    Dim rangecells As Range
    Set rangecells = Sheet4.Range("PLBD.Aadhar").Cells
    Dim i As Long
    ReDim Aadhar39_PL(ColCount_PAN39)
    For i = 1 To ColCount_PAN39
        Aadhar39_PL(i) = rangecells.item(i).value
'        If Not chkCompulsory(Aadhar39_PL(i)) Then
'            MsgBx_BadDept39 = MsgBx_BadDept39 + "*Please enter  Aadhaar Number at Sr. No " & i & " in schedule Profit and Loss A/c" & Chr(13)
'            ValidateAadhar_39 = False
'            Exit Function
'        End If
        
    If Aadhar39_PL(i) <> "" Then
         If Aadhar39_PL(i) = "000000000000" Then
        MsgBx_BadDept39 = MsgBx_BadDept39 + "*Aadhaar Number cannot be All Zero's at Sr. No " & i & " in schedule Profit and Loss A/c" & Chr(13)
            ValidateAadhar_39 = False
            Exit Function
        End If
        
        
         If Aadhar39_PL(i) = "111111111111" Then
            MsgBx_BadDept39 = MsgBx_BadDept39 + "*Aadhaar Number cannot be All One's at Sr. No " & i & " in schedule Profit and Loss A/c" & Chr(13)
            ValidateAadhar_39 = False
            Exit Function
        End If
        
           If Len(Aadhar39_PL(i)) <> 12 Then
            MsgBx_BadDept39 = MsgBx_BadDept39 + "*Please enter a valid 12 character Aadhaar Number at Sr. No " & i & " in schedule Profit and Loss A/c" & Chr(13)
            ValidateAadhar_39 = False
            Exit Function
        End If
      End If
    Next
End Function
Function ChkAdhaarNoPL(AadharNumberPL As Variant) As Boolean
On Error Resume Next
    ChkAdhaarNoPL = True
'    AadharNumberPL = Sheet4.Range("PLBD.Aadhar")
    If AadharNumberPL <> "" Then
        If AadharNumberPL = "000000000000" Or AadharNumberPL = "111111111111" Then
            ChkAdhaarNoPL = False
            Exit Function
        End If
      
        If Len(AadharNumberPL) <> 12 Then
            ChkAdhaarNoPL = False
            Exit Function
        End If
    End If
End Function
Function ValidateAmt_39() As Boolean
ValidateAmt_39 = True
    Dim rangecells As Range
    Set rangecells = Sheet4.Range("PLBD.Amount").Cells
    Dim i As Long
    ReDim Amt39_PL(ColCount_PAN39)
    For i = 1 To ColCount_PAN39
        Amt39_PL(i) = rangecells.item(i).value
        If Not chkCompulsory(Amt39_PL(i)) Then
            MsgBx_BadDept39 = MsgBx_BadDept39 + "* Please enter the amount of Bad Debts at Point-47  at Sr. No " & i & " in schedule Profit and Loss A/c" & Chr(13)
            ValidateAmt_39 = False
            Exit Function
        End If
    Next
    
    If Len(Sheet4.Range("PLBaddebts").value) > 14 Then
            MsgBx_BadDept39 = MsgBx_BadDept39 + "* Total at point 47 cannont be more than 14 digits schedule Profit and Loss A/c" & Chr(13)
            ValidateAmt_39 = False
            Exit Function
        End If
    
    
End Function
Sub setTablepan_39()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Sheet4.Range("PLBD.PAN").Cells
    mIntCells = Sheet4.Range("PLBD.PAN").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount_PAN39 = ccount
End Sub
Sub setTableAadhar39_PL()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Sheet4.Range("PLBD.Aadhar").Cells
    mIntCells = Sheet4.Range("PLBD.Aadhar").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount_Aadhar39 = ccount
End Sub
Sub setTableAmt39_PL()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Sheet4.Range("PLBD.Amount").Cells
    mIntCells = Sheet4.Range("PLBD.Amount").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount_Amt39 = ccount
End Sub

Function CheckDoneePAN(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.
'ITR 3 is for individuals .So,4th character of PAN should be "P"

    CheckDoneePAN = True
    If Len(PAN) > 0 Then
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
            CheckDoneePAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
            CheckDoneePAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
            CheckDoneePAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 4, 1)) Then
            CheckDoneePAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
            CheckDoneePAN = False
            Exit Function
        End If
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
            CheckDoneePAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
            CheckDoneePAN = False
            Exit Function
        End If
    End If
End Function

'new code
Sub setTableOtherpan_1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Sheet4.Range("PL_Name").Cells
    mIntCells = Sheet4.Range("PL_Name").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount_OtherPAN1 = ccount
End Sub

Sub setTableOtherpan_2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Sheet4.Range("PL_Address_Flat").Cells
    mIntCells = Sheet4.Range("PL_Address_Flat").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount_OtherPAN2 = ccount
End Sub

Sub setTableOtherpan_3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Sheet4.Range("PL_Address_Village").Cells
    mIntCells = Sheet4.Range("PL_Address_Village").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount_OtherPAN3 = ccount
End Sub

Sub setTableOtherpan_4()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Sheet4.Range("PL_Address_Road").Cells
    mIntCells = Sheet4.Range("PL_Address_Road").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount_OtherPAN4 = ccount
End Sub

Sub setTableOtherpan_5()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Sheet4.Range("PL_Address_Area").Cells
    mIntCells = Sheet4.Range("PL_Address_Area").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount_OtherPAN5 = ccount
End Sub

Sub setTableOtherpan_6()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Sheet4.Range("PL_Address_City").Cells
    mIntCells = Sheet4.Range("PL_Address_City").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount_OtherPAN6 = ccount
End Sub

Sub setTableOtherpan_7()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Sheet4.Range("PL_Address_State").Cells
    mIntCells = Sheet4.Range("PL_Address_State").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    ColCount_OtherPAN7 = ccount
End Sub

Sub setTableOtherpan_8()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Sheet4.Range("PL_Address_Country").Cells
    mIntCells = Sheet4.Range("PL_Address_Country").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).value) Then
        ccount = ccount + 1
        End If
    Next
    ColCount_OtherPAN8 = ccount
End Sub

Sub setTableOtherpan_9()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Sheet4.Range("PL_Address_Pin").Cells
    mIntCells = Sheet4.Range("PL_Address_Pin").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount_OtherPAN9 = ccount
End Sub

Sub setTableOtherpan_10()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Sheet4.Range("PL_Address_Zip").Cells
    mIntCells = Sheet4.Range("PL_Address_Zip").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount_OtherPAN10 = ccount
End Sub

Sub setTableOtherpan_11()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Sheet4.Range("PL_Amount").Cells
    mIntCells = Sheet4.Range("PL_Amount").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount_OtherPAN11 = ccount
End Sub


Function ValidateOtherPan()
ValidateOtherPan = True
setTableOtherpan_1
setTableOtherpan_2
setTableOtherpan_3
setTableOtherpan_4
setTableOtherpan_5
setTableOtherpan_6
setTableOtherpan_7
setTableOtherpan_8
setTableOtherpan_9
setTableOtherpan_10
setTableOtherpan_11

ColCount_OtherPAN1 = WorksheetFunction.Max(0, ColCount_OtherPAN1, ColCount_OtherPAN2, ColCount_OtherPAN3, ColCount_OtherPAN4, ColCount_OtherPAN5, ColCount_OtherPAN6, ColCount_OtherPAN7, ColCount_OtherPAN8, ColCount_OtherPAN9, ColCount_OtherPAN10, ColCount_OtherPAN11)
If Not ValidateOtherPanName Then ValidateOtherPan = False
If Not ValidateAddress_PL Then ValidateOtherPan = False
If Not ValidateOtherPanAmount Then ValidateOtherPan = False

End Function



Function ValidateOtherPanName() As Boolean
 
    ValidateOtherPanName = True
    
    Dim rangecells As Range
    Set rangecells = Sheet4.Range("PL_Name").Cells
    ReDim Pl_OtherPanName(ColCount_OtherPAN1)
    Dim i As Long
    Dim cellrange As Variant
    For i = 1 To ColCount_OtherPAN1
    cellrange = GetMergedAddressCell(rangecells, i)
        Pl_OtherPanName(i) = Sheet4.Range(cellrange).value
     If Not chkMandatory(Pl_OtherPanName(i)) Then
         MsgBx_OtherPan = MsgBx_OtherPan + "* Please enter Name at Sr. No  " & i & " in schedule Profit and Loss A/c" & Chr(13)
         ValidateOtherPanName = False
         Exit Function
     End If
     
     If Len(Pl_OtherPanName(i)) > 125 Then
         MsgBx_OtherPan = MsgBx_OtherPan + "* Name  at Sr. No  " & i & " cannot be more than 125 characters in schedule Profit and Loss A/c" & Chr(13)
         ValidateOtherPanName = False
         Exit Function
     End If
 Next
End Function

Function ValidateAddress_PL() As Boolean
ValidateAddress_PL = True
    Dim rangecells1, rangecells2, rangecells3, rangecells4, rangecells5, rangecells6, rangecells7, rangecells8, rangecells9 As Range
    Dim A As Long
    Dim i As Long
    
    'setTableInfo_AL1

    Set rangecells1 = Sheet4.Range("PL_Address_Flat").Cells
    Set rangecells2 = Sheet4.Range("PL_Address_Village").Cells
    Set rangecells3 = Sheet4.Range("PL_Address_Road").Cells
    Set rangecells4 = Sheet4.Range("PL_Address_Area").Cells
    Set rangecells5 = Sheet4.Range("PL_Address_City").Cells
    Set rangecells6 = Sheet4.Range("PL_Address_State").Cells
    Set rangecells7 = Sheet4.Range("PL_Address_Country").Cells
    Set rangecells8 = Sheet4.Range("PL_Address_Pin").Cells
    Set rangecells9 = Sheet4.Range("PL_Address_Zip").Cells
    
    
    ReDim PL_Address_Flat(ColCount_OtherPAN1)
    ReDim PL_Address_Village(ColCount_OtherPAN1)
    ReDim PL_Address_Road(ColCount_OtherPAN1)
    ReDim PL_Address_Area(ColCount_OtherPAN1)
    ReDim PL_Address_City(ColCount_OtherPAN1)
    ReDim PL_Address_State(ColCount_OtherPAN1)
    ReDim PL_Address_Country(ColCount_OtherPAN1)
    ReDim PL_Address_Pin(ColCount_OtherPAN1)
    ReDim PL_Address_Zip(ColCount_OtherPAN1)
    
    For i = 1 To ColCount_OtherPAN1
        PL_Address_Flat(i) = rangecells1.item(i).value
        PL_Address_Village(i) = rangecells2.item(i).value
        PL_Address_Road(i) = rangecells3.item(i).value
        PL_Address_Area(i) = rangecells4.item(i).value
        PL_Address_City(i) = rangecells5.item(i).value
        PL_Address_State(i) = rangecells6.item(i).value
        PL_Address_Country(i) = rangecells7.item(i).value
        PL_Address_Pin(i) = rangecells8.item(i).value
        PL_Address_Zip(i) = rangecells9.item(i).value
        
        'Flat
        If PL_Address_Flat(i) = "" Then
            MsgBx_OtherPan = MsgBx_OtherPan + "* Flat/ Door/ Building in Sheet is mandatory at Sr.No " & i & " in schedule Profit & Loss A/c" & Chr(13)
            ValidateAddress_PL = False
            Exit Function
        End If
        
        If Len(PL_Address_Flat(i)) > 50 Then
            MsgBx_OtherPan = MsgBx_OtherPan + "* Flat/ Door/ Building cannot be more than 50 characters at Sr.No " & i & " in schedule Profit & Loss A/c" & Chr(13)
            ValidateAddress_PL = False
            Exit Function
        End If
        
        'Villages
        If PL_Address_Village(i) = "" Then
            'MsgBx_OtherPan = MsgBx_OtherPan + "* Name of Premises/Building/Village  is Mandatory at Sr.No " & i & " in Sheet Profit and Loss" & Chr(13)
            'ValidateAddress_PL = False
            'Exit Function
        End If
        
        If Len(PL_Address_Village(i)) > 50 Then
            MsgBx_OtherPan = MsgBx_OtherPan + "* Name of Premises / Building / Village cannot be more than 50 characters  at Sr.No " & i & "in schedule Profit & Loss A/c" & Chr(13)
            ValidateAddress_PL = False
            Exit Function
        End If
        
        'Road
        If PL_Address_Road(i) = "" Then

        End If

        If Len(PL_Address_Road(i)) > 50 Then
            MsgBx_OtherPan = MsgBx_OtherPan + "* Road/ Street/Post office cannot be more than 50 characters  at Sr.No " & i & "  in schedule Profit & Loss A/c" & Chr(13)
            ValidateAddress_PL = False
            Exit Function
        End If


        'Area
        If PL_Address_Area(i) = "" Then
            MsgBx_OtherPan = MsgBx_OtherPan + "* Area/ Locality is mandatory at Sr.No " & i & " in schedule Profit & Loss A/c" & Chr(13)
            ValidateAddress_PL = False
            Exit Function
        End If
        
        If Len(PL_Address_Area(i)) > 50 Then
            MsgBx_OtherPan = MsgBx_OtherPan + "* Area/locality cannot exceed 50 characters  at Sr.No " & i & " in schedule Profit & Loss A/c" & Chr(13)
            ValidateAddress_PL = False
            Exit Function
        End If
        
        'Town
        If PL_Address_City(i) = "" Then
            MsgBx_OtherPan = MsgBx_OtherPan + "* Town/City/District is Mandatory at Sr.No " & i & " in schedule Profit & Loss A/c" & Chr(13)
            ValidateAddress_PL = False
            Exit Function
        End If
        
        If Len(PL_Address_City(i)) > 50 Then
            MsgBx_OtherPan = MsgBx_OtherPan + "* Town/City/District cannot exceed 50 characters  at Sr.No " & i & "  in schedule Profit & Loss A/c" & Chr(13)
            ValidateAddress_PL = False
            Exit Function
        End If
        
        'State
        If PL_Address_State(i) = "" Or UCase(PL_Address_State(i)) = "(SELECT)" Then
            MsgBx_OtherPan = MsgBx_OtherPan + "* State  is Mandatory at Sr.No " & i & "in schedule Profit & Loss A/c" & Chr(13)
            ValidateAddress_PL = False
            Exit Function
        End If
        
        'Country
        If PL_Address_Country(i) = "" Or UCase(PL_Address_Country(i)) = "(SELECT)" Then
            MsgBx_OtherPan = MsgBx_OtherPan + "* Country  is Mandatory at Sr.No " & i & " in schedule Profit & Loss A/c" & Chr(13)
            ValidateAddress_PL = False
            Exit Function
        End If
        
        'Pincode
        If PL_Address_State(i) <> "99-STATE OUTSIDE INDIA" Then
        If PL_Address_Pin(i) = "" Then
            MsgBx_OtherPan = MsgBx_OtherPan + "* Pin Code  is Mandatory at Sr.No " & i & "in schedule Profit & Loss A/c" & Chr(13)
            ValidateAddress_PL = False
            Exit Function
        End If
        
        If Len(PL_Address_Pin(i)) > 6 Then
            MsgBx_OtherPan = MsgBx_OtherPan + "* PinCode cannot exceed 6 digits  at Sr.No " & i & "  in schedule Profit & Loss A/c" & Chr(13)
            ValidateAddress_PL = False
            Exit Function
        End If
        
        If Not IsNumeric(PL_Address_Pin(i)) Then
            MsgBx_OtherPan = MsgBx_OtherPan + "* PinCode  at Sr.No " & i & "  is Invalid  in schedule Profit & Loss A/c" & Chr(13)
            ValidateAddress_PL = False
            Exit Function
        End If
        End If
        
        If PL_Address_State(i) = "99-STATE OUTSIDE INDIA" Then
        If PL_Address_Zip(i) = "" Then
            MsgBx_OtherPan = MsgBx_OtherPan + "* Zip Code  is Mandatory at Sr.No " & i & " in schedule Profit & Loss A/c" & Chr(13)
            ValidateAddress_PL = False
            Exit Function
        End If
        
        If Len(PL_Address_Zip(i)) > 8 Then
            MsgBx_OtherPan = MsgBx_OtherPan + "* ZipCode  cannot exceed 8 digits at Sr.No " & i & "   in schedule Profit & Loss A/c" & Chr(13)
            ValidateAddress_PL = False
            Exit Function
        End If
        
       If Not checkfieldspecialcharacter(PL_Address_Zip(i)) Then
             MsgBx_OtherPan = MsgBx_OtherPan + "* characters < > & ' " & Chr(34) & " are not allowed in ZipCode in schedule Profit & Loss A/c" & Chr(13)
            ValidateAddress_PL = False
            Exit Function
        End If
        End If
        
        
        'State & Country Check
        Dim StateCode As Variant
        Dim CountryCode As Variant
        
        StateCode = UCase(Mid(Trim(PL_Address_State(i)), 1, InStr(1, Trim(PL_Address_State(i)), "-") - 1))
        CountryCode = UCase(Mid(Trim(PL_Address_Country(i)), 1, InStr(1, Trim(PL_Address_Country(i)), "-") - 1))
        
        If (StateCode <> "99" And CountryCode <> "91") Then
            MsgBx_OtherPan = MsgBx_OtherPan + "* Country and State mapping of Address for immovable asset at Sr.No " & i & "  is Invalid in schedule Profit & Loss A/c" & Chr(13)
            ValidateAddress_PL = False
            Exit Function
        End If
        
    Next
End Function

Function ValidateOtherPanAmount() As Boolean
 
    ValidateOtherPanAmount = True
    
    Dim rangecells As Range
    Set rangecells = Sheet4.Range("PL_Amount").Cells
    ReDim Pl_OtherPanAmount(ColCount_OtherPAN1)
    Dim i As Long
     Dim cellrange As Variant
    For i = 1 To ColCount_OtherPAN1
     cellrange = GetMergedAddressCell(rangecells, i)
        Pl_OtherPanAmount(i) = Sheet4.Range(cellrange).value
       
     If Not chkMandatory(Pl_OtherPanAmount(i)) Then
         MsgBx_OtherPan = MsgBx_OtherPan + "* Please enter amount at point 47 (ii) at Sr. No  " & i & "in schedule Profit & Loss A/c " & Chr(13)
         ValidateOtherPanAmount = False
         Exit Function
     End If
 Next
 
 If Len(Sheet4.Range("PLOthers").value) > 14 Then
        MsgBx_OtherPan = MsgBx_OtherPan + "* Total at point 47 (ii) cannont be more than 14 digits schedule Profit & Loss A/c " & Chr(13)
         ValidateOtherPanAmount = False
         Exit Function
     End If
 
End Function

'end

Sub setTblinfo_NOBC()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AD.NOBCode").count
    Set rangecells = Range("Sec44AD.NOBCode").Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
 end_NOBC = ccount
 End Sub
 
 Sub setTblinfo_NOBName()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AD.NOBName").count
    Set rangecells = Range("Sec44AD.NOBName").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_NOBName = ccount
End Sub
Sub setTblinfo_NOBDescription()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AD.NOBDescription").count
    Set rangecells = Range("Sec44AD.NOBDescription").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_NOBDescription = ccount
End Sub


Function ValidateNOB1() As Boolean
On Error Resume Next

ValidateNOB1 = True

Dim cellrange As String
Dim cellrange1 As String
Dim cellrange2 As String
Dim rangecells As Range
Dim rangecells1 As Range
Dim rangecells2 As Range
Dim value As Variant
Dim i As Long
Set rangecells = Range("Sec44AD.NOBCode").Cells
Set rangecells1 = Range("Sec44AD.NOBName").Cells
Set rangecells2 = Range("Sec44AD.NOBDescription").Cells
setTblinfo_NOBC
setTblinfo_NOBName
setTblinfo_NOBDescription

ReDim NOBCode(end_NOBName)
ReDim NOBName(end_NOBName)
ReDim NOBDescription(end_NOBName)

For i = 1 To end_NOBName

 cellrange = GetMergedAddressCell(rangecells, i)
 cellrange1 = GetMergedAddressCell(rangecells1, i)
 cellrange2 = GetMergedAddressCell(rangecells2, i)
 
  value = value + 1
 NOBCode(i) = Sheet4.Range(cellrange).value
 NOBCode(i) = Mid(NOBCode(i), 1, WorksheetFunction.Search("-", NOBCode(i)) - 1)
 NOBDescription(i) = Sheet4.Range(cellrange2).value
 NOBName(i) = Sheet4.Range(cellrange1).value

Next

If value = 0 Then
    Dim income_BP As Variant
    income_BP = IIf(Sheet4.Range("PL_61i").value = "", 0, Sheet4.Range("PL_61i").value)
    If income_BP > 0 Then
       If Not end_NOBC > 0 Then
        errorMsg_NOB1 = errorMsg_NOB1 & "Please select  Business code from the drop down."
        'ValidateNOB1 = False
       End If
    End If
End If

If end_NOBName <> end_NOBC Then
        errorMsg_NOB1 = errorMsg_NOB1 & "Please fill all the mandatory fields in table under section 44AD in schedule Profit & Loss A/c"
        'ValidateNOB1 = False
End If

'If end_NOBName <> 0 And end_NOBC <> 0 Then

'End If

If errorMsg_NOB1 <> "" Then
fmsgboxsmall errorMsg_NOB1
End
'ValidateNOB1 = False
End If


End Function

Sub setTblinfo_NOBC1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AD.NOBCode1").count
    Set rangecells = Range("Sec44AD.NOBCode1").Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
 end_NOBC1 = ccount
 End Sub
 
 Sub setTblinfo_NOBName1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AD.NOBName1").count
    Set rangecells = Range("Sec44AD.NOBName1").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_NOBName1 = ccount
End Sub
Sub setTblinfo_NOBDescription1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AD.NOBDescription1").count
    Set rangecells = Range("Sec44AD.NOBDescription1").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_NOBDescription1 = ccount
End Sub

Function ValidateNOB2() As Boolean
On Error Resume Next

ValidateNOB2 = True

Dim cellrange As String
Dim cellrange1 As String
Dim cellrange2 As String
Dim rangecells As Range
Dim rangecells1 As Range
Dim rangecells2 As Range
Dim value As Variant
Dim i As Long
Set rangecells = Range("Sec44AD.NOBCode1").Cells
Set rangecells1 = Range("Sec44AD.NOBName1").Cells
Set rangecells2 = Range("Sec44AD.NOBDescription1").Cells
setTblinfo_NOBC1
setTblinfo_NOBName1
setTblinfo_NOBDescription1

ReDim NOBCode1(end_NOBC1)
ReDim NOBName1(end_NOBC1)
ReDim NOBDescription1(end_NOBC1)

For i = 1 To end_NOBC1

 cellrange = GetMergedAddressCell(rangecells, i)
 cellrange1 = GetMergedAddressCell(rangecells1, i)
 cellrange2 = GetMergedAddressCell(rangecells2, i)
 
  value = value + 1
 NOBCode1(i) = Sheet4.Range(cellrange).value
 NOBCode1(i) = Mid(NOBCode1(i), 1, WorksheetFunction.Search("-", NOBCode1(i)) - 1)
 NOBDescription1(i) = Sheet4.Range(cellrange2).value
 NOBName1(i) = Sheet4.Range(cellrange1).value

Next

If value = 0 Then
    Dim income_BP As Variant
    income_BP = IIf(Sheet4.Range("PL_62i").value = "", 0, Sheet4.Range("PL_62i").value)
    If income_BP > 0 Then
        errorMsg_NOB2 = errorMsg_NOB2 & "Please select  Business code from the drop down."
        'ValidateNOB2 = False
    End If
End If

If end_NOBC1 <> end_NOBName1 Then
        errorMsg_NOB2 = errorMsg_NOB2 & "Please fill all the mandatory fields in table under section 44ADA in schedule Profit & Loss A/c"
        'ValidateNOB2 = False
End If



If errorMsg_NOB2 <> "" Then
fmsgboxsmall errorMsg_NOB2
End
End If

End Function


Sub setTblinfo_NOBC2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AD.NOBCode2").count
    Set rangecells = Range("Sec44AD.NOBCode2").Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
 end_NOBC2 = ccount
 End Sub
 
 Sub setTblinfo_NOBName2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AD.NOBName2").count
    Set rangecells = Range("Sec44AD.NOBName2").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_NOBName2 = ccount
End Sub
Sub setTblinfo_NOBDescription2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AD.NOBDescription2").count
    Set rangecells = Range("Sec44AD.NOBDescription2").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_NOBDescription2 = ccount
End Sub

Function ValidateNOB3() As Boolean
On Error Resume Next

ValidateNOB3 = True

Dim cellrange As String
Dim cellrange1 As String
Dim cellrange2 As String
Dim rangecells As Range
Dim rangecells1 As Range
Dim rangecells2 As Range
Dim value As Variant
Dim i As Long
Set rangecells = Range("Sec44AD.NOBCode2").Cells
Set rangecells1 = Range("Sec44AD.NOBName2").Cells
Set rangecells2 = Range("Sec44AD.NOBDescription2").Cells
setTblinfo_NOBC2
setTblinfo_NOBName2
setTblinfo_NOBDescription2

ReDim NOBCode2(end_NOBC2)
ReDim NOBName2(end_NOBC2)
ReDim NOBDescription2(end_NOBC2)

For i = 1 To end_NOBC2

 cellrange = GetMergedAddressCell(rangecells, i)
 cellrange1 = GetMergedAddressCell(rangecells1, i)
 cellrange2 = GetMergedAddressCell(rangecells2, i)
 
  value = value + 1
 NOBCode2(i) = Sheet4.Range(cellrange).value
 NOBCode2(i) = Mid(NOBCode2(i), 1, WorksheetFunction.Search("-", NOBCode2(i)) - 1)
 NOBDescription2(i) = Sheet4.Range(cellrange2).value
 NOBName2(i) = Sheet4.Range(cellrange1).value

Next

If value = 0 Then
    Dim income_BP As Variant
    income_BP = IIf(Sheet4.Range("PL_63ii").value = "", 0, Sheet4.Range("PL_63ii").value)
    If income_BP > 0 Then
        errorMsg_NOB3 = errorMsg_NOB3 & "Please select  Business code from the drop down."
        'ValidateNOB3 = False
    End If
End If

If end_NOBC2 <> end_NOBName2 Then
        errorMsg_NOB3 = errorMsg_NOB3 & "Please fill all the mandatory fields in table under section 44AE in schedule Profit & Loss A/c"
        'ValidateNOB3 = False
End If

If errorMsg_NOB3 <> "" Then
fmsgboxsmall errorMsg_NOB3
End
End If

End Function

Sub setTableInfo_RegNo()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("Sec44AE.RegNo").count
 Set rangecells = Sheet4.Range("Sec44AE.RegNo").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end44AE_RegNo = ccount
End Sub
Sub setTblinfo_OLH()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Sec44AE.OLH").count
    Set rangecells = Range("Sec44AE.OLH").Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
 end44AE_OLH = ccount
 End Sub
Sub setTableInfo_TC()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("Sec44AE.TonnageCpct").count
 Set rangecells = Sheet4.Range("Sec44AE.TonnageCpct").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end44AE_TC = ccount
End Sub
Sub setTableInfo_Month()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("Sheet44AE.NoOfMonths").count
 Set rangecells = Sheet4.Range("Sheet44AE.NoOfMonths").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end44AE_Months = ccount
End Sub
Sub setTableInfo_PIncome()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet4.Range("Sheet44AE.PresumptiveIncome").count
 Set rangecells = Sheet4.Range("Sheet44AE.PresumptiveIncome").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).value = "" Then
         ccount = ccount + 1
     End If
 Next
 end44AE_PIncome = ccount
End Sub


Function ValidateSec44AE() As Boolean
On Error Resume Next
Dim cntr As Long
Dim cntr2  As Variant
Dim value As Variant
Dim i As Variant
value = 0
cntr = 0
ValidateSec44AE = True
setTableInfo_RegNo
ReDim SEC44AERegNo(end44AE_RegNo)
ReDim SEC44AEOLH(end44AE_RegNo)
ReDim SEC44AETC(end44AE_RegNo)
ReDim SEC44AEMonth(end44AE_RegNo)
ReDim SEC44AEPIncome(end44AE_RegNo)
ReDim Greater12_BP(end44AE_RegNo)
ReDim Lesser12_BP(end44AE_RegNo)


setTblinfo_OLH
setTableInfo_TC
setTableInfo_Month
setTableInfo_PIncome

end44AE_RegNo = WorksheetFunction.Max(0, end44AE_RegNo, end44AE_OLH, end44AE_TC, end44AE_Months, end44AE_PIncome)
Dim rangecells As Range
Dim rangecells1 As Range
Dim cellrange As Variant
Dim cellrange1 As Variant
Set rangecells = Sheet4.Range("Sec44AE.RegNo").Cells
Set rangecells1 = Sheet4.Range("Sheet44AE.PresumptiveIncome").Cells

For cntr = 1 To end44AE_RegNo



        cellrange = GetMergedAddressCell(rangecells, cntr)
        SEC44AERegNo(cntr) = Sheet4.Range(cellrange).value
        

 If Not CheckEmptyDDList(SEC44AERegNo(cntr)) Then
 value = value + 1
 SEC44AERegNo(cntr) = Sheet4.Range(cellrange).value
 'SEC44AEOLH(cntr2) = Mid(NOBCode3(cntr2), 1, WorksheetFunction.Search("-", NOBCode3(cntr2)) - 1)
 SEC44AEOLH(cntr) = Range("Sec44AE.OLH").item(cntr).value
 If SEC44AEOLH(cntr) = "Owned" Then
    SEC44AEOLH(cntr) = "OWN"
ElseIf SEC44AEOLH(cntr) = "Leased" Then
    SEC44AEOLH(cntr) = "LEASE"
ElseIf SEC44AEOLH(cntr) = "Hired" Then
    SEC44AEOLH(cntr) = "HIRED"
End If
 SEC44AETC(cntr) = Range("Sec44AE.TonnageCpct").item(cntr).value
 SEC44AEMonth(cntr) = Range("Sheet44AE.NoOfMonths").item(cntr).value
cellrange1 = GetMergedAddressCell(rangecells1, cntr)
SEC44AEPIncome(cntr) = Sheet4.Range(cellrange1).value
 
 Greater12_BP(cntr) = Range("greater12").item(cntr).value
 Lesser12_BP(cntr) = Range("lesser12").item(cntr).value
 Else

 End If
Next

For i = 1 To end44AE_RegNo

        If Not chkCompulsory(SEC44AERegNo(i)) Then
                errorMsg_NOB5 = errorMsg_NOB5 + "* Please enter the Registration no of goods carriage in Point 63  at Sr. No  " & i & "in schedule Profit & Loss A/c " & Chr(13)
                'ValidateSec44AE = False
                'Exit Function
        End If
        
            If Not checkfieldspecialcharacter1(SEC44AERegNo(i)) Then
                errorMsg_NOB5 = errorMsg_NOB5 + "* Special characters are not allowed in Registration no of goods carriage in Point 63  at Sr. No  " & i & " in schedule Profit & Loss A/c" & Chr(13)
                'ValidateSec44AE = False
                'Exit Function
        End If
        
        If Not chkCompulsory(SEC44AETC(i)) Then
                errorMsg_NOB5 = errorMsg_NOB5 + "* Please enter the tonnage capacity of goods carriage(in MT) in Point 63  at Sr. No  " & i & "in schedule Profit & Loss A/c " & Chr(13)
                'ValidateSec44AE = False
               ' Exit Function
        End If
        
        If Not chkCompulsory(SEC44AEMonth(i)) Then
                errorMsg_NOB5 = errorMsg_NOB5 + "* Please enter Period of Holding  at Sr. No  " & i & "in schedule Profit & Loss A/c " & Chr(13)
                'ValidateSec44AE = False
               ' Exit Function
        End If
        
        If Not chkCompulsory(SEC44AEPIncome(i)) Then
                errorMsg_NOB5 = errorMsg_NOB5 + "* Please enter Presumptive income u/s 44AE for the goods carriage   at Sr. No  " & i & " in schedule Profit & Loss A/c" & Chr(13)
                'ValidateSec44AE = False
                'Exit Function
        End If
        
        If isdropdownblank(SEC44AEOLH(i)) Then
                errorMsg_NOB5 = errorMsg_NOB5 + "* Please Select an option from the drop down whether goods carriage is owned/leased/hired at Sr. No  " & i & " in schedule Profit & Loss A/c" & Chr(13)
                'ValidateSec44AE = False
                'Exit Function
        End If
        
        If SEC44AETC(i) > 12 Then
        If SEC44AEPIncome(i) < Greater12_BP(i) Then
            errorMsg_NOB5 = errorMsg_NOB5 + "* Income cannot be less than Rs. 1000 per month in case of Heavy Goods Vehicle  at Sr. No  " & i & "in schedule Profit & Loss A/c " & Chr(13)
            'ValidateSec44AE = False
                'Exit Function
      End If
      End If
      
      If SEC44AETC(i) <= 12 Then
        If SEC44AEPIncome(i) < Lesser12_BP(i) Then
            errorMsg_NOB5 = errorMsg_NOB5 + "* Income cannot be less than Rs. 7500 per month in case of other than Heavy Goods Vehicle  at Sr. No  " & i & "in schedule Profit & Loss A/c " & Chr(13)
            'ValidateSec44AE = False
                'Exit Function
      End If
      End If

Next
If Sheet4.Range("PL_TotalMonths").value > 120 Then
 errorMsg_NOB5 = errorMsg_NOB5 + "* Total of number of months cannot exceed 120 in Schedule Profit Loss in schedule Profit & Loss A/c" & Chr(13)
End If

If end44AE_PIncome <> end44AE_RegNo Or end44AE_PIncome <> end44AE_OLH Or end44AE_PIncome <> end44AE_TC Or end44AE_PIncome <> end44AE_Months Then
        errorMsg_NOB5 = errorMsg_NOB5 & "Please enter all mandatory Fields in 44AE Table in schedule Profit & Loss A/c"
       'ValidateSec44AE = False
End If


If errorMsg_NOB5 <> "" Then
fmsgboxsmall errorMsg_NOB5
End
End If

End Function

Public Function CheckEmptyDDList(item As Variant) As Boolean
On Error Resume Next
CheckEmptyDDList = False

If IsEmpty(item) Or Mid(item, 1, 1) = "(" Or item = "" Then CheckEmptyDDList = True

End Function

Function ValidateCinditionalcheck() As Boolean
On Error Resume Next



TempPL_61i = Sheet4.Range("PL_61i").value
TempPL_61a = Sheet4.Range("PL_61a").value
TempPL_61b = Sheet4.Range("PL_61b").value

TempPL_61ii = Sheet4.Range("PL_61ii").value
TempPL_61iia = Sheet4.Range("PL_61iia").value
TempPL_61iib = Sheet4.Range("PL_61iib").value

TempPL_62i = Sheet4.Range("PL_62i").value
TempPL_62ii = Sheet4.Range("PL_62ii").value

TempPL_63ii = Sheet4.Range("PL_63ii").value

TempPL_64GrossTurnOver = Sheet4.Range("PL_64GrossTurnOver").value
TempPL_64ia1 = Sheet4.Range("PL_64ia1").value
TempPL_64ia2 = Sheet4.Range("PL_64ia2").value
TempPL_64ib = Sheet4.Range("PL_64ib").value
TempPL_64ic = Sheet4.Range("PL_64ic").value

TempPL_64id = Sheet4.Range("PL_64id").value
TempPL_64iia = Sheet4.Range("PL_64iia").value
TempPL_64iia_i = Sheet4.Range("PL_64iia_i").value
TempPL_64iia_ii = Sheet4.Range("PL_64iia_ii").value
TempPL_64iib = Sheet4.Range("PL_64iib").value

TempPL_64iic = Sheet4.Range("PL_64iic").value
TempPL_64iid = Sheet4.Range("PL_64iid").value
TempPL_64iii = Sheet4.Range("PL_64iii").value
TempPL_65i = Sheet4.Range("PL_65i").value
TempPL_65ii = Sheet4.Range("PL_65ii").value
TempPL_65iii = Sheet4.Range("PL_65iii").value
TempPL_65iv = Sheet4.Range("PL_65iv").value

If Len(TempPL_61ii) > 8 Then
errorMsg_NOB6 = errorMsg_NOB6 + "*Presumptive income under section 44AD(iia+iib) cannot be more than 8 digits in schedule Profit & Loss A/c " & Chr(13)
Exit Function
End If

If TempPL_61i > 20000000 Then
errorMsg_NOB6 = errorMsg_NOB6 + "* Since gross turnover or gross receipts exceeds 2 crore, you are not eligible to avail the provisions of Section 44AD" & Chr(13)
Exit Function
End If

If TempPL_61a < TempPL_61iia Then
errorMsg_NOB6 = errorMsg_NOB6 + "* Invalid presumptive income under 44AD, as 61(i)(a) is less than 61(ii)(a). in schedule Profit & Loss A/c" & Chr(13)
Exit Function
End If

If TempPL_61a + TempPL_61b > 20000000 Then
errorMsg_NOB6 = errorMsg_NOB6 + "* Since gross turnover or gross receipts exceeds 2 crore, you are not eligible to avail the provisions of Section 44AD" & Chr(13)
Exit Function
End If

If TempPL_61b < TempPL_61iib Then
errorMsg_NOB6 = errorMsg_NOB6 + "* Invalid presumptive income under 44AD, as 61(i)(b) is less than 61(ii)(b). in schedule Profit & Loss A/c" & Chr(13)
Exit Function
End If

If TempPL_61a > 0 Then
If Not TempPL_61iia > 0 Then
errorMsg_NOB6 = errorMsg_NOB6 + "* field 61(ii)(a) is mandatory in schedule Profit & Loss A/c" & Chr(13)
Exit Function
End If
End If

If TempPL_61iia > TempPL_61a Then
errorMsg_NOB6 = errorMsg_NOB6 + "* 61(ii)(a) can not be greater than 61(i)(a) sheet Profit and Loss in schedule Profit & Loss A/c" & Chr(13)
Exit Function
End If

If TempPL_61iia < 0.06 * TempPL_61a Then
errorMsg_NOB6 = errorMsg_NOB6 + "* Since the income is less than 6% of 61(i)(a),you are not entitled to avail the benefit of section 44AD" & Chr(13)
Exit Function
End If

If TempPL_61b > 0 Then
If Not TempPL_61iib > 0 Then
errorMsg_NOB6 = errorMsg_NOB6 + "* Field 61(ii)(b) is mandatory in schedule Profit & Loss A/c" & Chr(13)
Exit Function
End If
End If

If TempPL_61iib > TempPL_61b Then
errorMsg_NOB6 = errorMsg_NOB6 + "* 61(ii)(b) can not be greater than 61(i)(b) sheet Profit and Loss in schedule Profit & Loss A/c" & Chr(13)
Exit Function
End If

If TempPL_61iib < 0.08 * TempPL_61b Then
errorMsg_NOB6 = errorMsg_NOB6 + "* Since the income is less than 8% of 61(i)(b), you are not entitled to avail the benefit of section 44AD" & Chr(13)
Exit Function
End If

If TempPL_62i < TempPL_62ii Then
errorMsg_NOB6 = errorMsg_NOB6 + "* Invalid presumptive income under 44ADA, as 62(i) is less than 62(ii) in schedule Profit & Loss A/c" & Chr(13)
Exit Function
End If

If TempPL_62i > 5000000 Then
errorMsg_NOB6 = errorMsg_NOB6 + "* If Gross Receipts are more than Rs.50 Lakh, you cannot avail the provision of Section 44ADA" & Chr(13)
Exit Function
End If

If TempPL_62i > 0 Then
If Not TempPL_62ii > 0 Then
errorMsg_NOB6 = errorMsg_NOB6 + "* Field 62(ii)(b) is mandatory in schedule Profit & Loss A/c" & Chr(13)
Exit Function
End If
End If

If TempPL_62ii < 0.5 * TempPL_62i Then
errorMsg_NOB6 = errorMsg_NOB6 + "* If the income is less than 50% of 62(i),you are not eligible to avail the provisions of Section 44ADA"" & Chr(13)"
Exit Function
End If

'If TempPL_64ic < 0.5 * TempPL_62i Then
'errorMsg_NOB6 = errorMsg_NOB6 + "* If the income is less than 50% of 62(i), it is mandatory to maintain books of accounts and have a tax audit under 44AB" & Chr(13)
'Exit Function
'End If

If TempPL_64ic > 0 Then
If TempPL_64ib = 0 Then
errorMsg_NOB6 = errorMsg_NOB6 + "* In order to claim Expenses, Gross Profit must be greater than zero in schedule Profit & Loss A/c" & Chr(13)
Exit Function
End If
End If

If TempPL_64ib > TempPL_64GrossTurnOver Then
errorMsg_NOB6 = errorMsg_NOB6 + "* Gross profit cannot be more than Gross turnover/Gross receipts in 64(i) in schedule Profit & Loss A/c" & Chr(13)
Exit Function
End If

If TempPL_64iic > 0 Then
If TempPL_64iib = 0 Then
errorMsg_NOB6 = errorMsg_NOB6 + "* In order to claim Expenses Gross Profit must be greater than zero in schedule Profit & Loss A/c" & Chr(13)
Exit Function
End If
End If

If TempPL_64iib > TempPL_64iia Then
errorMsg_NOB6 = errorMsg_NOB6 + "* Gross profit cannot be more than Gross receipts in 64(ii) in schedule Profit & Loss A/c" & Chr(13)
Exit Function
End If



If TempPL_65i > 0 Then
If TempPL_65ii = "" Then
errorMsg_NOB6 = errorMsg_NOB6 + "* Please enter Gross Profit from Speculative Activity in schedule Profit & Loss A/c" & Chr(13)
Exit Function
End If
End If
 
If TempPL_65i > 0 Then
If TempPL_65iii = "" Then
errorMsg_NOB6 = errorMsg_NOB6 + "* Please enter Expenditure relating to Speculative Activity in schedule Profit & Loss A/c" & Chr(13)
Exit Function
End If
End If

If TempPL_65ii > TempPL_65i Then
errorMsg_NOB6 = errorMsg_NOB6 + "* Gross Profit cannot be more than Turnover From Speculative Activity in schedule Profit & Loss A/c" & Chr(13)
Exit Function
End If

End Function



Function SetStateCountryImmovableAsset1(Tval As Variant, Tadd As Variant, CallType As Variant)
On Error Resume Next
Application.EnableEvents = False


'------------------Un Protect--------------------
   sPassword = EfilingCommon.getmsgstate
   Sheet4.Unprotect Password:=sPassword
'------------------Un Protect Ends--------------------
    
    Dim rangecellsState, rangecellsCountry, rangecellsPin, rangecellsZip As Range
    Dim mIntCells, q  As Long
    Dim PrevVal As Variant

    Set rangecellsState = Range("PL_Address_State").Cells
    Set rangecellsCountry = Range("PL_Address_Country").Cells
    Set rangecellsPin = Range("PL_Address_Pin").Cells
    Set rangecellsZip = Range("PL_Address_Zip").Cells



    mIntCells = Range("PL_Name").count

If CallType = "S" Then
    If UCase(Tval) = "99-STATE OUTSIDE INDIA" Then
        Sheet4.Range(Tadd).Select
        Sheet4.Range(Tadd).Offset(0, 1) = "(Select)"
       Sheet4.Range(Tadd).Offset(0, 2) = ""
        Sheet4.Range(Tadd).Offset(0, 2).Interior.Color = (&HD8D8D8)
        Sheet4.Range(Tadd).Offset(0, 2).Locked = True
        Sheet4.Range(Tadd).Offset(0, 3).Interior.Color = (&HCCFFCC)
        Sheet4.Range(Tadd).Offset(0, 3).Locked = False
        GoTo endfd
    ElseIf UCase(Tval) = "(SELECT)" Then
        Sheet4.Range(Tadd).Select
        Sheet4.Range(Tadd).Offset(0, 1) = "(Select)"
        Sheet4.Range(Tadd).Offset(0, 2) = ""
        Sheet4.Range(Tadd).Offset(0, 2).Interior.Color = (&HCCFFCC)
        Sheet4.Range(Tadd).Offset(0, 2).Locked = False
        Sheet4.Range(Tadd).Offset(0, 3) = ""
        Sheet4.Range(Tadd).Offset(0, 3).Interior.Color = (&HD8D8D8)
        Sheet4.Range(Tadd).Offset(0, 3).Locked = True
        GoTo endfd
    Else
        Sheet4.Range(Tadd).Select
        Sheet4.Range(Tadd).Offset(0, 1) = "91-INDIA"
        Sheet4.Range(Tadd).Offset(0, 2) = ""
        Sheet4.Range(Tadd).Offset(0, 2).Interior.Color = (&HCCFFCC)
        Sheet4.Range(Tadd).Offset(0, 2).Locked = False
        Sheet4.Range(Tadd).Offset(0, 3) = ""
        Sheet4.Range(Tadd).Offset(0, 3).Interior.Color = (&HD8D8D8)
        Sheet4.Range(Tadd).Offset(0, 3).Locked = True
        GoTo endfd
    End If
Else

    If UCase(Tval) = "(SELECT)" Then
        Sheet4.Range(Tadd).Select
        Sheet4.Range(Tadd).Offset(0, -1) = "(Select)"
        Sheet4.Range(Tadd).Offset(0, 1) = ""
        Sheet4.Range(Tadd).Offset(0, 1).Interior.Color = (&HCCFFCC)
        Sheet4.Range(Tadd).Offset(0, 1).Locked = False
        Sheet4.Range(Tadd).Offset(0, 2) = ""
        Sheet4.Range(Tadd).Offset(0, 2).Interior.Color = (&HD8D8D8)
        Sheet4.Range(Tadd).Offset(0, 2).Locked = True
        GoTo endfd
    ElseIf UCase(Tval) <> "91-INDIA" Then
        Sheet4.Range(Tadd).Select
        Sheet4.Range(Tadd).Offset(0, -1) = "99-STATE OUTSIDE INDIA"
        Sheet4.Range(Tadd).Offset(0, 1) = ""
        Sheet4.Range(Tadd).Offset(0, 1).Interior.Color = (&HD8D8D8)
        Sheet4.Range(Tadd).Offset(0, 1).Locked = True
        Sheet4.Range(Tadd).Offset(0, 2).Interior.Color = (&HCCFFCC)
        Sheet4.Range(Tadd).Offset(0, 2).Locked = False
        GoTo endfd
    Else
        Sheet4.Range(Tadd).Select
        Sheet4.Range(Tadd).Offset(0, -1) = "(Select)"
        Sheet4.Range(Tadd).Offset(0, 1) = ""
        Sheet4.Range(Tadd).Offset(0, 1).Interior.Color = (&HCCFFCC)
        Sheet4.Range(Tadd).Offset(0, 1).Locked = False
        Sheet4.Range(Tadd).Offset(0, 2) = ""
        Sheet4.Range(Tadd).Offset(0, 2).Interior.Color = (&HD8D8D8)
        Sheet4.Range(Tadd).Offset(0, 2).Locked = True
        GoTo endfd
    End If

End If

    For q = 1 To mIntCells
        If UCase(rangecellsState.item(q).value) = "99-STATE OUTSIDE INDIA" Then
            rangecellsCountry.item(q).value = "(Select)"
            rangecellsPin.item(q).value = ""
            rangecellsZip.item(q).value = ""
        ElseIf UCase(rangecellsState.item(q).value) = "(SELECT)" Then
            rangecellsCountry.item(q).value = "(Select)"
            rangecellsPin.item(q).value = ""
            rangecellsZip.item(q).value = ""
        Else
            rangecellsCountry.item(q).value = "91-INDIA"
            rangecellsPin.item(q).value = ""
            rangecellsZip.item(q).value = ""
        End If
        
        If UCase(rangecellsCountry.item(q).value) <> "91-INDIA" Then
            rangecellsState.item(q).value = "99-STATE OUTSIDE INDIA"
            rangecellsPin.item(q).value = ""
            rangecellsZip.item(q).value = ""
        Else
            rangecellsCountry.item(q).value = "(Select)"
            rangecellsPin.item(q).value = ""
            rangecellsZip.item(q).value = ""
        End If
        
    Next
endfd:
'----------------Lock Password-------------------START---
  Sheet4.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
Application.EnableEvents = True
End Function

