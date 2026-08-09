Attribute VB_Name = "importschedule112a"
'25/10/2021
Dim rngname_112A
Dim MsgBox_112A
Dim validate_112A As Boolean
'Added by Riyaz on 13/06/2025 for 2025-26 New Dev
Sub ImportSchedule112A()
validate_112A = True
Dim filePath, text As String

Dim answer As Integer
 
answer = MsgBox("Do you want to import CSV? If you import csv, All data will be clear.", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
If answer = vbYes Then
    MsgBox_112A = ""
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Please select a CSV file."
        .Filters.add "CSV File", "*.csv", 1
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
    import112aModule (text)
End If
End Sub
'Added by Riyaz on 13/06/2025 for 2025-26 New Dev
Sub import112aModule(text As String)
    Const Delim As String = ","
    Dim MyData As String, strData() As String, TmpAr() As String
    Dim TwoDArray() As String
    Dim i As Long, n As Long
    strData() = Split(text, vbCrLf)
    For i = LBound(strData) To UBound(strData)
        If Len(Trim(strData(i))) <> 0 And i <> 0 Then
            TmpAr = Split(strData(i), Delim)
            n = n + 1
            ReDim Preserve TwoDArray(13, 1 To n)
           ' Call validate112(TmpAr, n)
            TwoDArray(0, n) = TmpAr(0)
            TwoDArray(1, n) = TmpAr(1)
            TwoDArray(2, n) = TmpAr(2)
            TwoDArray(3, n) = TmpAr(3)
            TwoDArray(4, n) = TmpAr(4)
            TwoDArray(5, n) = TmpAr(5)
            TwoDArray(6, n) = TmpAr(6)
            TwoDArray(7, n) = TmpAr(7)
            TwoDArray(8, n) = TmpAr(8)
            TwoDArray(9, n) = TmpAr(9)
            TwoDArray(10, n) = TmpAr(10)
            TwoDArray(11, n) = TmpAr(11)
            TwoDArray(12, n) = TmpAr(12)
            TwoDArray(13, n) = TmpAr(13)
'            TwoDArray(14, n) = TmpAr(14)
'            TwoDArray(15, n) = TmpAr(14)
        End If
    Next i

     If validate_112A = False Then
        MsgBox MsgBox_112A, vbOKOnly, "Error(s)"
        Exit Sub
    End If

    ShareAcq_112A = Sheet53.Range("ShareAcq_112A").Column
'    ShareTra_112A = Sheet53.Range("ShareTra_112A").Column
    ISINCode_112A = Sheet53.Range("ISINCode_112A").Column
    NameofShare_112A = Sheet53.Range("NameofShare_112A").Column
    NoofShare_112A = Sheet53.Range("NoofShare_112A").Column
    SalePriceperShare_112A = Sheet53.Range("SalePriceperShare_112A").Column
    TotalSaleValue_112A = Sheet53.Range("TotalSaleValue_112A").Column 'A
    COAwithoutIndex_112A = Sheet53.Range("COAwithoutIndex_112A").Column 'A
    COAwithIndex_112A = Sheet53.Range("COAwithIndex_112A").Column
    LTCGAssetAcquired_112A = Sheet53.Range("LTCGAssetAcquired_112A").Column 'A
    FairMarketValue_112A = Sheet53.Range("FairMarketValue_112A").Column
    TotalFairMarketValue_112A = Sheet53.Range("TotalFairMarketValue_112A").Column 'A
    ExpenditureWholly_112A = Sheet53.Range("ExpenditureWholly_112A").Column
    TotalDeductions_112A = Sheet53.Range("TotalDeductions_112A").Column 'A
    Balance_112A = Sheet53.Range("Balance_112A").Column 'A
    
    TotalExRow = Sheet53.Range("ShareAcq_112A").Rows.count
    
    TotalXMLRow = n
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    
    If (TotalXMLRow > 0) Then
        If Sheet53.Range("ShareAcq_112A").Locked = False Then
            Sheet53.Range("ShareAcq_112A").ClearContents
        End If
        
'        If Sheet53.Range("ShareTra_112A").Locked = False Then
'            Sheet53.Range("ShareTra_112A").ClearContents
'        End If
        
        If Sheet53.Range("ISINCode_112A").Locked = False Then
            Sheet53.Range("ISINCode_112A").ClearContents
        End If
        
        If Sheet53.Range("NameofShare_112A").Locked = False Then
            Sheet53.Range("NameofShare_112A").ClearContents
        End If
        
        If Sheet53.Range("NoofShare_112A").Locked = False Then
            Sheet53.Range("NoofShare_112A").ClearContents
        End If
        
        If Sheet53.Range("SalePriceperShare_112A").Locked = False Then
            Sheet53.Range("SalePriceperShare_112A").ClearContents
        End If
        
        If Sheet53.Range("TotalSaleValue_112A").Locked = False Then
            Sheet53.Range("TotalSaleValue_112A").ClearContents
        End If
        
        If Sheet53.Range("COAwithIndex_112A").Locked = False Then
            Sheet53.Range("COAwithIndex_112A").ClearContents
        End If
        
        If Sheet53.Range("FairMarketValue_112A").Locked = False Then
            Sheet53.Range("FairMarketValue_112A").ClearContents
        End If
        
        If Sheet53.Range("ExpenditureWholly_112A").Locked = False Then
            Sheet53.Range("ExpenditureWholly_112A").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
        AddDiffRows_112A_IMPORT (TotalDiffRow)
    End If

    rowcount = getRowNo(Sheet53.Range("ShareAcq_112A").name)
    
    For i = 1 To n
        
        If Sheet53.Cells(rowcount, ShareAcq_112A).Locked = False Then
            If UCase(TwoDArray(0, i)) = "BE" Or TwoDArray(0, i) = "On or before 31st January 2018" Then
                TwoDArray(0, i) = "'On or before 31st January 2018"
            End If
            If UCase(TwoDArray(0, i)) = "AE" Or TwoDArray(0, i) = "After 31st January 2018" Then
                TwoDArray(0, i) = "'After 31st January 2018"
            End If
            Sheet53.Cells(rowcount, ShareAcq_112A).value = TwoDArray(0, i)
        End If
        
'        If Sheet53.Cells(rowcount, ShareTra_112A).Locked = False Then
'            If UCase(TwoDArray(1, i)) = "BE" Then
'                TwoDArray(1, i) = "'Before 23rd July 2024 "
'            End If
'            If UCase(TwoDArray(1, i)) = "AE" Then
'                TwoDArray(1, i) = "'On or after 23rd July 2024"
'            End If
'            Sheet53.Cells(rowcount, ShareTra_112A).value = TwoDArray(1, i)
'        End If
        
        If Sheet53.Cells(rowcount, ISINCode_112A).Locked = False Then
            Sheet53.Cells(rowcount, ISINCode_112A).value = TwoDArray(1, i)
        End If
        
        If Sheet53.Cells(rowcount, NameofShare_112A).Locked = False Then
            Sheet53.Cells(rowcount, NameofShare_112A).value = TwoDArray(2, i)
        End If
        
        If Sheet53.Cells(rowcount, NoofShare_112A).Locked = False Then
            Sheet53.Cells(rowcount, NoofShare_112A).value = TwoDArray(3, i)
        End If
        
        If Sheet53.Cells(rowcount, SalePriceperShare_112A).Locked = False Then
            Sheet53.Cells(rowcount, SalePriceperShare_112A).value = TwoDArray(4, i)
        End If
        
        If Sheet53.Cells(rowcount, TotalSaleValue_112A).Locked = False Then
            Sheet53.Cells(rowcount, TotalSaleValue_112A).value = TwoDArray(5, i)
        End If

        If Sheet53.Cells(rowcount, COAwithIndex_112A).Locked = False Then
            Sheet53.Cells(rowcount, COAwithIndex_112A).value = TwoDArray(7, i)
        End If
               
        If Sheet53.Cells(rowcount, FairMarketValue_112A).Locked = False Then
            Sheet53.Cells(rowcount, FairMarketValue_112A).value = TwoDArray(9, i)
        End If
        
        If Sheet53.Cells(rowcount, ExpenditureWholly_112A).Locked = False Then
            Sheet53.Cells(rowcount, ExpenditureWholly_112A).value = TwoDArray(11, i)
        End If
        
            
        rowcount = rowcount + 1
                
    Next i
End Sub

Function validate112(TmpAr() As String, n As Long)
    If Not UCase(TmpAr(0)) = "BE" And Not UCase(TmpAr(0)) = "AE" Then
        MsgBox_112A = MsgBox_112A + "Please enter valid Share/Unit Acquired Sr. No." & n & " in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
    
    If (Not chkMandatory_112a(TmpAr(1))) Then
        MsgBox_112A = MsgBox_112A + "Please enter the ISIN Code. at Sr. No." & n & " in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
        
    If Len(TmpAr(1)) > 0 Then
        If Len(TmpAr(1)) <> 12 Then
            MsgBox_112A = MsgBox_112A + "Please enter the valid ISIN Code. at Sr. No." & n & " in Schedule 112A" & Chr(13)
            validate_112A = False
        End If
    End If
        
    If UCase(Trim(Mid(TmpAr(1), 1, 2))) <> "IN" Then
        MsgBox_112A = MsgBox_112A + "ISIN Code. at Sr. No." & n & " should start with ""IN"" in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
    
    If (Not chkMandatory_112a(TmpAr(2))) Then
        MsgBox_112A = MsgBox_112A + "Please enter Name of the Share/Unit at Sr. No." & n & " in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
        
    If Len(TmpAr(2)) > 125 Then
        MsgBox_112A = MsgBox_112A + "Name of the Share/Unit at Sr. No." & n & " should not exceed 125 characters." & " in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
    
    If Not chkMandatory_112a(TmpAr(3)) Then
        MsgBox_112A = MsgBox_112A + "Please enter the No. of Shares/Units at Sr. No." & n & " in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
        
    If CDec(TmpAr(3)) < 0 Then
        MsgBox_112A = MsgBox_112A + "* No. of Shares/Units at Sr. No." & n & " cannot be negative in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
    Dim tmpArray
    tmpArray = Split(TmpAr(3), ".")
    If Len(TmpAr(3)) > 19 Or Len(tmpArray(0)) > 14 Then
        MsgBox_112A = MsgBox_112A + "* No. of Shares/Units at Sr. No." & n & " cannot exceed 19 digits in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
    
    If Not chkMandatory_112a(TmpAr(4)) Then
        MsgBox_112A = MsgBox_112A + "Please enter the Sale-price per Share/Unit at Sr. No." & n & " in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
        
    If CDec(TmpAr(4)) < 0 Then
        MsgBox_112A = MsgBox_112A + "* Sale-price per Share/Unit at Sr. No." & n & " cannot be negative in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
        
    tmpArray = Split(TmpAr(4), ".")
    If Len(TmpAr(4)) > 19 Or Len(tmpArray(0)) > 14 Then
        MsgBox_112A = MsgBox_112A + "* Sale-price per Share/Unit at Sr. No." & n & " cannot exceed 19 digits in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
    
    If Len(TmpAr(5)) > 14 Then
        MsgBox_112A = MsgBox_112A + "Total Sale Value at Sr. No " & n & " cannot exceed 14 digits." & " in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
    
    If CDec(TmpAr(5)) < 0 Then
        MsgBox_112A = MsgBox_112A + "* Total Sale Value at Sr. No." & n & " cannot be negative in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
    
    If Not chkMandatory_112a(TmpAr(7)) Then
        MsgBox_112A = MsgBox_112A + "Please enter the Cost of acquisition at Sr. No." & n & " in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
    
    If CDec(TmpAr(7)) < 0 Then
        MsgBox_112A = MsgBox_112A + "* Cost of acquisition at Sr. No." & n & " cannot be negative in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
    
    If Not chkMandatory_112a(TmpAr(9)) Then
        MsgBox_112A = MsgBox_112A + "Please enter the Fair Market Value per share/unit as on 31st January,2018. If share is Aquired after 31st Jan 2018 , please enter '0' at Sr. No." & i & " in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
        
    If CDec(TmpAr(9)) < 0 Then
        MsgBox_112A = MsgBox_112A + "* Fair Market Value per share/unit as on 31st January,2018 in Column 10,If share is Acquired after 31st Jan 2018 , please enter '0' at Sr. No." & i & " cannot be negative in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
        
    tmpArray = Split(TmpAr(9), ".")
    If Len(TmpAr(9)) > 19 Or Len(tmpArray(0)) > 14 Then
        MsgBox_112A = MsgBox_112A + "* Fair Market Value per share/unit as on 31st January,2018 in Column 10,If share is Acquired after 31st Jan 2018 , please enter '0' at Sr. No." & i & " cannot exceed 19 digits in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
    
    If Not chkMandatory_112a(TmpAr(11)) Then
        MsgBox_112A = MsgBox_112A + "Please enter the Expenditure incurred wholly and exclusively in connection with transfer at Sr. No." & i & " in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
        
    If CDec(TmpAr(11)) < 0 Then
        MsgBox_112A = MsgBox_112A + "* Expenditure incurred wholly and exclusively in connection at Sr. No." & i & " cannot be negative in Schedule 112A" & Chr(13)
        validate_112A = False
    End If
End Function

Function AddDiffRows_112A_IMPORT(DiffRows As Long)
    setTblinfo_112A_IMPORT
    Sheet53.Activate
    SelectLastRow ("ShareAcq_112A")
    insertRowUnderSectionWithFormula (DiffRows)
    Call ExendRangeNameToTable(DiffRows, rngname_112A)
End Function
Function setTblinfo_112A_IMPORT()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet53.Range("ShareAcq_112A").count
    Set rangecells = Sheet53.Range("ShareAcq_112A").Cells
    Dim natureAmount As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
    ColCount_112A = ccount
    rngname_112A = "ShareAcq_112A;ISINCode_112A;NameofShare_112A;NoofShare_112A;SalePriceperShare_112A;TotalSaleValue_112A;COAwithoutIndex_112A;COAwithIndex_112A;LTCGAssetAcquired_112A;FairMarketValue_112A;TotalFairMarketValue_112A;ExpenditureWholly_112A;TotalDeductions_112A;Balance_112A;"
End Function

Function chkMandatory_112a(field As Variant) As Boolean
chkMandatory_112a = True
If Len(Trim(field)) <= 0 Then
    chkMandatory_112a = False
End If
If IsEmpty(field) Then
chkMandatory_112a = False
End If
End Function
