Attribute VB_Name = "SchIT"
Option Explicit

Public msgbox_TDS2, msgbox_ESOP, msgbox_ESOP2, msgbox_ESOP3 As String

Public msgbox_ESOP4, msgbox_ESOP5, msgbox_ESOPPD As String

Public ColCount2 As Long
Public ColCount2_1 As Long
Public ColCount2_2 As Long
Public ColCount2_3 As Long

Public ColCount As Long



Public DateCredit_TDS As Variant
Public BSR_TDS As Variant
Public SerialNum_TDS As Variant
Public TaxPaid3_TDS As Variant
Public end_IT, rngname_IT As Variant


Sub Cmd_AddRows_ESOP_Click()
Dim vRows As Long
EfilingCommon.DefinedgridNameRange = ("ESOPCurrAY||ESOP.Date||ESOP.AttributedTax")
ActiveCellRange = EfilingCommon.searchLastRow("ESOPCurrAY")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
'Change- esop
Sub AddDiffRowsEsop(DiffRows As Long)
Dim vRows
Sheet46.Activate
EfilingCommon.DefinedgridNameRange = ("ESOPCurrAY||ESOP.Date||ESOP.AttributedTax")
ActiveCellRange = EfilingCommon.searchLastRow("ESOPCurrAY")
vRows = EfilingCommon.insertRowUnderSectionWithFormula(DiffRows)
End Sub
Sub Cmd_Validate_ESOP_Click()
ValidateSheet_ESOP
'fmsgboxStatus "Sheet ESOP is ok"
End Sub

 Sub ValidateSheet_ESOP()
On Error Resume Next
subProcCaption = "Validating ESOP"
 
 
If ValidateESOP Then
    If Not ValidateESOP2 Then
        fmsgboxsmall (msgbox_ESOP2)
    ElseIf Not ValidateESOP3 Then
        fmsgboxsmall (msgbox_ESOP3)
    ElseIf Not ValidateESOP4 Then
        fmsgboxsmall (msgbox_ESOP4)
    ElseIf Not ValidateESOP5 Then
        fmsgboxsmall (msgbox_ESOP5)
    ElseIf Not ValidateESOP_PDA Then
        fmsgboxsmall (msgbox_ESOPPD)
    Else
        fmsgboxsmall ("Sheet ESOP is ok")
    End If
Else
fmsgboxsmall (msgbox_ESOP)
End If
'End If
End Sub
 

'Sub ValidateSheet_ESOP()
'On Error Resume Next
'subProcCaption = "Validating ESOP"
'
'
'If ValidateESOP Then
'    If ValidateESOP2 Then
'        MsgBox "Sheet ESOP is ok"
'    Else
'        MsgBox (msgbox_ESOP2)
'    End If
'
'    'Malli==============================
'   If ValidateESOP3 Then
''       MsgBox "Sheet ESOP is ok"
'    Else
'        MsgBox (msgbox_ESOP3)
'    End If
'
'  'Newly added for AY 2025-26
'    If ValidateESOP4 Then
''       MsgBox "Sheet ESOP is ok"
'     Else
'        MsgBox (msgbox_ESOP4)
'    End If
'
'Else
''    If Not ValidateESOP2 Then
''        MsgBox (msgbox_ESOP + msgbox_ESOP2)
''        '=============
''        ElseIf Not ValidateESOP2 Then        'Malli added
''        MsgBox (msgbox_ESOP + msgbox_ESOP3)
''        '=============
''        Else
''        MsgBox (msgbox_ESOP)
''    End If
''MsgBox "Sheet ESOP is ok"
'
'    If Not ValidateESOP2 Then
'        MsgBox (msgbox_ESOP + msgbox_ESOP2)
'        '=============
'         ElseIf Not ValidateESOP2 Then        'Malli added
'         MsgBox (msgbox_ESOP + msgbox_ESOP3)
'        '=============
'
'         ElseIf Not ValidateESOP3 Then        'For AY 2025-26
'         MsgBox (msgbox_ESOP + msgbox_ESOP4)
'
'        ElseIf msgbox_ESOP = "" Then
'
'        MsgBox "Sheet ESOP is ok"
'        Else
'        MsgBox (msgbox_ESOP)
'    End If
'
'
'
'End If
'
''''''
'
'If Not ValidateESOP_PDA Then
'
'fmsgboxsmall (msgbox_ESOPPD)
'CloseMsg
'Else
'fmsgboxsmall "Sheet ESOP is ok"
'End If
'
'
''If Not (ValidateESOP And ValidateESOP2) Then
''MsgBox (msgbox_ESOP + msgbox_ESOP2)
''Else
''fmsgboxsmall "Sheet ESOP is ok"
''End If
'
'End Sub

'Sub ValidateSheet_ESOP()
'On Error Resume Next
'subProcCaption = "Validating ESOP"
'If Not ValidateESOP Then
'fmsgboxsmall (msgbox_ESOP)
'CloseMsg
'Else
'fmsgboxsmall "Sheet ESOP is ok"
'End If
'
'End Sub


Sub Cmd_AddRows_IT_Click()
Dim vRows As Long
EfilingCommon.DefinedgridNameRange = ("IT.BSRCode||IT.DateDep||IT.SrlNoOfChaln||IT.Amt||IT.FormulaOFS||FormulaOfQ||FormulaOfSAT||FormulaOfSAT1||FormulaOfExSAT||FormulaOfExSAT1||FormulaOfSATNew")
ActiveCellRange = EfilingCommon.searchLastRow("IT.BSRCode")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

Sub Cmd_Validate_IT_Click()
ValidateSheet_IT
fmsgboxStatus "Sheet IT is ok"
End Sub


Sub ValidateSheet_IT()
subProcCaption = "Validating IT"
If Not ValidateAdvanceTax Then
fmsgboxsmall (msgbox_TDS2)
CloseMsg
End If
End Sub


Function ValidateAdvanceTax() As Boolean
ValidateAdvanceTax = True
msgbox_TDS2 = "IT" & Chr(10)
    setTableInfo_Grid3
    setTableInfo1_Grid3
    setTableInfo2_Grid3
    setTableInfo3_Grid3
    

   
    If Not ValidateBSR_TDS Then ValidateAdvanceTax = False
    If Len((Sheet25.Range("IT.BSRCode").item(1).value) > 0) Then
'        If Not ValidateBSR_TDS Then ValidateAdvanceTax = False
        If Not ValidateDateCreditToGovt Then ValidateAdvanceTax = False
        If Not ValidateSerialNum Then ValidateAdvanceTax = False
        If Not ValidateTaxPaid Then ValidateAdvanceTax = False
    End If
    

    
    
'    If ((ColCount2 <> ColCount2_1) Or (ColCount2 <> ColCount2_2) Or (ColCount2 <> ColCount2_3)) Then
'    msgbox_TDS2 = msgbox_TDS2 + "* Enter All Mandatory Details schedule IT"
'    ValidateAdvanceTax = False
'    End If
End Function

Function ValidateBSR_TDS() As Boolean
ValidateBSR_TDS = True
'setTableInfo_Grid3
Dim rangecells As Range
Set rangecells = Sheet25.Range("IT.BSRCode").Cells
ColCount2 = Application.WorksheetFunction.Max(0, ColCount2, ColCount2_1, ColCount2_2, ColCount2_3)
Dim i As Long
ReDim BSR_TDS(ColCount2)
noOfProcessSub = ColCount2
For i = 1 To ColCount2
BSR_TDS(i) = rangecells.item(i).value
If Len(BSR_TDS(i)) = 0 Then
End If

If Not chkCompulsory(BSR_TDS(i)) Then
'    msgbox_TDS2 = msgbox_TDS2 + "* Please enter the BSR Code at Sr.NO " & i & "in schedule IT" & Chr(13)
    'SIT-70607 Changed by sai
    msgbox_TDS2 = msgbox_TDS2 + "* ""Please enter the BSR Code in schedule IT"" at Sr.No " & i & "" & Chr(13)
    
    ValidateBSR_TDS = False
    Exit Function
End If

'SIT-73931 Changed by sai

'Added by sudha as per SIT-94947
If Not IsNumeric(Mid(BSR_TDS(i), 1, 3)) Then


    'msgbox_TDS2 = msgbox_TDS2 + "* BSR code at Sr.No " & i & " can allow only Numeric Values in schedule IT" & Chr(13)
    msgbox_TDS2 = msgbox_TDS2 + "* BSR code must be First three is Numeric and Next Four alphanumeric in schedule IT at Sr.No " & i & "" & Chr(13)
    
    ValidateBSR_TDS = False
    Exit Function
End If

' If Not Validate_BSRCODE(BSR_TDS(i)) Then
'         msgbox_TDS2 = msgbox_TDS2 + "Invalid BSR code at Sr. No  " & i & "  in Sheet Schedule IT  BSR format should be First 3 numeric and next 4 alphanumeric." & Chr(13)
'         ValidateBSR_TDS = False
'        Exit Function
'End If
UpdateProgressBar
Next
End Function
Function Validate_BSRCODE(BsrCode As Variant) As Boolean
On Error Resume Next


    Validate_BSRCODE = True
    If Len(BsrCode) > 0 Then
        If Not IsNumeric(Mid(BsrCode, 1, 1)) Then
            Validate_BSRCODE = False
            Exit Function
        End If
        If Not IsNumeric(Mid(BsrCode, 2, 1)) Then
            Validate_BSRCODE = False
            Exit Function
        End If
        If Not IsNumeric(Mid(BsrCode, 3, 1)) Then
            Validate_BSRCODE = False
            Exit Function
        End If

        If Not checkfieldspecialcharacter_BsRCode(Mid(BsrCode, 4, 4)) Then
            Validate_BSRCODE = False
            Exit Function
        End If

    End If
End Function

Function checkfieldspecialcharacter_BsRCode(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter_BsRCode = True
    Dim arr As Variant
    arr = Array("@", "*", "!", "-", "_", "|", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter_BsRCode = False
            Exit Function
        End If
        Next
    Next
End Function

Function ValidateDateCreditToGovt() As Boolean
ValidateDateCreditToGovt = True
'setTableInfo_Grid3
Dim rangecells As Range
Set rangecells = Sheet25.Range("IT.DateDep").Cells
Dim i As Long
ReDim DateCredit_TDS(ColCount2)
For i = 1 To ColCount2
DateCredit_TDS(i) = rangecells.item(i).value
If Len(DateCredit_TDS(i)) = 0 Then
End If

'rangecells.item(i) = Replace(DateCredit_TDS(i), "-", "/")

If Not chkCompulsory(DateCredit_TDS(i)) Then
'    msgbox_TDS2 = msgbox_TDS2 + "* Please enter date of deposit at Sr.No " & i & " in schedule IT" & Chr(13)
    
    'msgbox_TDS2 = msgbox_TDS2 + "* ""Please enter date of deposit in schedule IT"" at Sr.No " & i & " " & Chr(13)
    'Added by Aavula on 09/10/2024
    
'    msgbox_TDS2 = msgbox_TDS2 + "* ""Please enter the date of Deposit"" at Sr.No " & i & " " & Chr(13)
    
    'Added by Aavula Naresh New Dev
    
    msgbox_TDS2 = msgbox_TDS2 + "* ""Please enter date of deposit in Schedule IT"" at Sr.No " & i & " " & Chr(13)

    
    ValidateDateCreditToGovt = False
    Exit Function
End If



Next
End Function


Function ValidateSerialNum() As Boolean
ValidateSerialNum = True
'setTableInfo_Grid3
Dim rangecells As Range
Set rangecells = Sheet25.Range("IT.SrlNoOfChaln").Cells
Dim i As Long
ReDim SerialNum_TDS(ColCount2)
For i = 1 To ColCount2
SerialNum_TDS(i) = rangecells.item(i).value
If Len(SerialNum_TDS(i)) = 0 Then
End If

If Not chkCompulsory(SerialNum_TDS(i)) Then

'    msgbox_TDS2 = msgbox_TDS2 + "* Please enter 5 digit serial number of challan at Sr.NO " & i & "in schedule IT" & Chr(13)
    
  'Added by sudha as per SIT-94950
  
    'msgbox_TDS2 = msgbox_TDS2 + "* ""Please enter 5 digit serial number of challan in schedule IT"" at Sr.NO " & i & "" & Chr(13)
     msgbox_TDS2 = msgbox_TDS2 + "* ""Please enter serial number of challan in Schedule IT"" at Sr.No " & i & "" & Chr(13)
    ValidateSerialNum = False
    Exit Function
End If


If Not IsNumeric(SerialNum_TDS(i)) Then
    msgbox_TDS2 = msgbox_TDS2 + "* Serial Number of Challan at Sr.NO " & i & " can allow only Numeric Values in schedule IT" & Chr(13)
    ValidateSerialNum = False
    Exit Function
End If
Next
End Function

Function ValidateTaxPaid() As Boolean
ValidateTaxPaid = True
'setTableInfo_Grid3
Dim rangecells As Range
Set rangecells = Sheet25.Range("IT.Amt").Cells
Dim i As Long
ReDim TaxPaid3_TDS(ColCount2)
For i = 1 To ColCount2
TaxPaid3_TDS(i) = rangecells.item(i).value
If Len(TaxPaid3_TDS(i)) = 0 Then
End If

If Not chkCompulsory(TaxPaid3_TDS(i)) Then
    'msgbox_TDS2 = msgbox_TDS2 + "* Please enter the Tax Paid Amount at Sr.NO " & i & " in schedule IT" & Chr(13)
    'ramya 70611
        msgbox_TDS2 = msgbox_TDS2 + "* ""Please enter amount in schedule IT"" at Sr.No " & i & "" & Chr(13)
    ValidateTaxPaid = False
    
    Exit Function
End If
Next
End Function

Sub setTableInfo_Grid3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet25.Range("IT.BSRCode").count
    Set rangecells = Sheet25.Range("IT.BSRCode").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2 = ccount
End Sub

Sub setTableInfo1_Grid3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet25.Range("IT.DateDep").count
    Set rangecells = Sheet25.Range("IT.DateDep").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_1 = ccount
End Sub


Sub setTableInfo2_Grid3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet25.Range("IT.SrlNoOfChaln").count
    Set rangecells = Sheet25.Range("IT.SrlNoOfChaln").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_2 = ccount
End Sub


Sub setTableInfo3_Grid3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet25.Range("IT.Amt").count
    Set rangecells = Sheet25.Range("IT.Amt").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_3 = ccount
End Sub



Sub ImportScheduleITCSV()
Dim filePath, text As String

'fmsgbox "Import personal/tax details from downloaded Pre-filled JSON or Import from already generated JSON of the current assessment year."
Dim answer As Integer
 
answer = MsgBox("Do you want to import CSV? If you import csv, All data will be clear.", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
If answer = vbYes Then
    msgbox_TDS2 = ""
    With Application.FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        .Title = "Please select a CSV file."
        'Filter to just the following types of files to narrow down selection options
        .Filters.add "CSV File", "*.csv", 1
        'Show the dialog box
        If .Show = True Then
            filePath = .SelectedItems.item(1)
        Else
            Exit Sub
        End If
        On Error Resume Next
    End With
    
    Open filePath For Input As #1
    text = Input$(LOF(1), 1)
    Close #1
    
    importITCSVFunc (text)
End If
End Sub
Sub importITCSVFunc(text As String)
    Const Delim As String = ","
    Dim MyData As String, strData() As String, TmpAr() As String
    Dim TwoDArray() As String
    Dim i As Long, n As Long
    n = 0
    Dim TotalExRow, TotalXMLRow, TotalDiffRow, rowcount As Variant
    strData() = Split(text, vbCrLf)
    For i = LBound(strData) To UBound(strData)
        If Len(Trim(strData(i))) <> 0 And i <> 0 Then
            TmpAr = Split(strData(i), Delim)
            n = n + 1
            ReDim Preserve TwoDArray(3, 1 To n)

            TwoDArray(0, n) = TmpAr(0)
            TwoDArray(1, n) = TmpAr(1)
            TwoDArray(2, n) = TmpAr(2)
            TwoDArray(3, n) = TmpAr(3)
            
        End If
    Next i
    
    Dim BSRCodeColNo, DateDepColNo, SrlNoOfChalnColNo, AmtColNo As Variant
    
    BSRCodeColNo = Sheet25.Range("IT.BSRCode").Column
    DateDepColNo = Sheet25.Range("IT.DateDep").Column
    SrlNoOfChalnColNo = Sheet25.Range("IT.SrlNoOfChaln").Column
    AmtColNo = Sheet25.Range("IT.Amt").Column
    
    
    TotalExRow = Sheet25.Range("IT.BSRCode").Rows.count
    TotalXMLRow = n
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    
        If Sheet25.Range("IT.BSRCode").Locked = False Then
            Sheet25.Range("IT.BSRCode").ClearContents
        End If
        If Sheet25.Range("IT.DateDep").Locked = False Then
            Sheet25.Range("IT.DateDep").ClearContents
        End If
        If Sheet25.Range("IT.SrlNoOfChaln").Locked = False Then
            Sheet25.Range("IT.SrlNoOfChaln").ClearContents
        End If
        If Sheet25.Range("IT.Amt").Locked = False Then
            Sheet25.Range("IT.Amt").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
        AddDiffRows_ScheduleIT (TotalDiffRow)
    End If
    rowcount = getRowNo(Sheet25.Range("IT.BSRCode").name)

    For i = 1 To n
        If Sheet25.Cells(rowcount, BSRCodeColNo).Locked = False Then
            Sheet25.Cells(rowcount, BSRCodeColNo).value = TwoDArray(0, i)
        End If
        If Sheet25.Cells(rowcount, DateDepColNo).Locked = False Then
            Sheet25.Cells(rowcount, DateDepColNo).value = TwoDArray(1, i)
        End If
        If Sheet25.Cells(rowcount, SrlNoOfChalnColNo).Locked = False Then
            Sheet25.Cells(rowcount, SrlNoOfChalnColNo).value = TwoDArray(2, i)
        End If
        If Sheet25.Cells(rowcount, AmtColNo).Locked = False Then
            Sheet25.Cells(rowcount, AmtColNo).value = TwoDArray(3, i)
        End If

        rowcount = rowcount + 1
               
    Next i
End Sub

'----end change---

