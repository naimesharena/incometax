Attribute VB_Name = "importschedule115ad"
'25/10/2021
Dim rngname_115AD
Dim MsgBox_115AD
Dim validate_115AD As Boolean
Sub ImportSchedule115AD()

Dim filePath, text As String

'fmsgbox "Import personal/tax details from downloaded Pre-filled JSON or Import from already generated JSON of the current assessment year."
Dim answer As Integer
 
answer = MsgBox("Do you want to import CSV? If you import csv, All data will be clear.", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
If answer = vbYes Then
    MsgBox_115AD = ""
    validate_115AD = True
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
    import115ADModule (text)
End If
End Sub

Sub import115ADModule(text As String)
    Const Delim As String = ","
    Dim MyData As String, strData() As String, TmpAr() As String
    Dim TwoDArray() As String
    Dim i As Long, n As Long
    strData() = Split(text, vbCrLf)
    For i = LBound(strData) To UBound(strData)
        If Len(Trim(strData(i))) <> 0 And i <> 0 Then
            TmpAr = Split(strData(i), Delim)
            n = n + 1
            ReDim Preserve TwoDArray(15, 1 To n)
           ' Call validate115ADA(TmpAr, n)
            '~~> TmpAr(1) : 1 for Col B, 0 would be A
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
'            TwoDArray(14, n) = TmpAr(14)  'Commented by Shrutika(11/05/2026)
'            TwoDArray(15, n) = TmpAr(14)
        End If
    Next i

     If validate_115AD = False Then
        MsgBox MsgBox_115AD, vbOKOnly, "Error(s)"
        Exit Sub
    End If

    ShareAcq_115AD_1 = Sheet54.Range("ShareAcq_115AD_1").Column
'    ShareTrans_115AD_1 = Sheet54.Range("ShareTrans_115AD_1").Column
    ISINCode_115AD = Sheet54.Range("ISINCode_115AD").Column
    NameofShare_115AD = Sheet54.Range("NameofShare_115AD").Column
    NoofShare_115AD = Sheet54.Range("NoofShare_115AD").Column
    SalePriceperShare_115AD = Sheet54.Range("SalePriceperShare_115AD").Column
    TotalSaleValue_115AD = Sheet54.Range("TotalSaleValue_115AD").Column
    COAwithoutIndex_115AD = Sheet54.Range("COAwithoutIndex_115AD").Column
    COAwithIndex_115AD = Sheet54.Range("COAwithIndex_115AD").Column
    LTCGAssetAcquired_115AD = Sheet54.Range("LTCGAssetAcquired_115AD").Column
    FairMarketValue_115AD = Sheet54.Range("FairMarketValue_115AD").Column
    TotalFairMarketValue_115AD = Sheet54.Range("TotalFairMarketValue_115AD").Column
    ExpenditureWholly_115AD = Sheet54.Range("ExpenditureWholly_115AD").Column
    TotalDeductions_115AD = Sheet54.Range("TotalDeductions_115AD").Column
    Balance_115AD = Sheet54.Range("Balance_115AD").Column
        
    TotalExRow = Sheet54.Range("ShareAcq_115AD_1").Rows.count
    
    TotalXMLRow = n
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    
    If (TotalXMLRow > 0) Then
        If Sheet54.Range("ShareAcq_115AD_1").Locked = False Then
            Sheet54.Range("ShareAcq_115AD_1").ClearContents
        End If
                    'Commented by Shrutika(11/05/2026)
'        If Sheet54.Range("ShareTrans_115AD_1").Locked = False Then
'            Sheet54.Range("ShareTrans_115AD_1").ClearContents
'        End If
        
        If Sheet54.Range("ISINCode_115AD").Locked = False Then
            Sheet54.Range("ISINCode_115AD").ClearContents
        End If
        
        If Sheet54.Range("NameofShare_115AD").Locked = False Then
            Sheet54.Range("NameofShare_115AD").ClearContents
        End If
        
        If Sheet54.Range("NoofShare_115AD").Locked = False Then
            Sheet54.Range("NoofShare_115AD").ClearContents
        End If
        
        If Sheet54.Range("SalePriceperShare_115AD").Locked = False Then
            Sheet54.Range("SalePriceperShare_115AD").ClearContents
        End If
        
        If Sheet54.Range("TotalSaleValue_115AD").Locked = False Then
            Sheet54.Range("TotalSaleValue_115AD").ClearContents
        End If
        
        If Sheet54.Range("COAwithIndex_115AD").Locked = False Then
            Sheet54.Range("COAwithIndex_115AD").ClearContents
        End If
        
        If Sheet54.Range("FairMarketValue_115AD").Locked = False Then
            Sheet54.Range("FairMarketValue_115AD").ClearContents
        End If
        
        If Sheet54.Range("ExpenditureWholly_115AD").Locked = False Then
            Sheet54.Range("ExpenditureWholly_115AD").ClearContents
        End If
        
    End If
    
   If (TotalDiffRow > 0) Then
        AddDiffRows_115AD (TotalDiffRow)
    End If

    rowcount = getRowNo(Sheet54.Range("ShareAcq_115AD_1").name)
    
    For i = 1 To n
        If Sheet54.Cells(rowcount, ShareAcq_115AD_1).Locked = False Then
            If UCase(TwoDArray(0, i)) = "BE" Then
                TwoDArray(0, i) = "'On or before 31st January 2018"
            End If
            If UCase(TwoDArray(0, i)) = "AE" Then
                TwoDArray(0, i) = "'After 31st January 2018"
            End If
            Sheet54.Cells(rowcount, ShareAcq_115AD_1).value = TwoDArray(0, i)
        End If
                  'Commented by Shrutika(11/05/2026)
'        If Sheet54.Cells(rowcount, ShareTrans_115AD_1).Locked = False Then
'            If UCase(TwoDArray(1, i)) = "BE" Then
'                TwoDArray(1, i) = "'Before 23rd July 2024 "
'            End If
'            If UCase(TwoDArray(1, i)) = "AE" Then
'                TwoDArray(1, i) = "'On or after 23rd July 2024"
'            End If
'            Sheet54.Cells(rowcount, ShareTrans_115AD_1).value = TwoDArray(1, i)
'        End If
        
        If Sheet54.Cells(rowcount, ISINCode_115AD).Locked = False Then
            Sheet54.Cells(rowcount, ISINCode_115AD).value = TwoDArray(1, i)
        End If
        
        If Sheet54.Cells(rowcount, NameofShare_115AD).Locked = False Then
            Sheet54.Cells(rowcount, NameofShare_115AD).value = TwoDArray(2, i)
        End If
        
        If Sheet54.Cells(rowcount, NoofShare_115AD).Locked = False Then
            Sheet54.Cells(rowcount, NoofShare_115AD).value = TwoDArray(3, i)
        End If
        
        If Sheet54.Cells(rowcount, SalePriceperShare_115AD).Locked = False Then
            Sheet54.Cells(rowcount, SalePriceperShare_115AD).value = TwoDArray(4, i)
        End If
        
        If Sheet54.Cells(rowcount, TotalSaleValue_115AD).Locked = False Then
            Sheet54.Cells(rowcount, TotalSaleValue_115AD).value = TwoDArray(5, i)
        End If
        
        If Sheet54.Cells(rowcount, COAwithIndex_115AD).Locked = False Then
            Sheet54.Cells(rowcount, COAwithIndex_115AD).value = TwoDArray(7, i)
        End If
        
        If Sheet54.Cells(rowcount, FairMarketValue_115AD).Locked = False Then
            Sheet54.Cells(rowcount, FairMarketValue_115AD).value = TwoDArray(9, i)
        End If
        
        If Sheet54.Cells(rowcount, ExpenditureWholly_115AD).Locked = False Then
            Sheet54.Cells(rowcount, ExpenditureWholly_115AD).value = TwoDArray(11, i)
        End If
        
        rowcount = rowcount + 1
                
    Next i
End Sub


Function validate115ADA(TmpAr() As String, n As Long)
    If Not UCase(TmpAr(0)) = "BE" And Not UCase(TmpAr(0)) = "AE" Then
        MsgBox_115AD = MsgBox_115AD + "Please enter valid Share/Unit Acquired Sr. No." & n & " in Schedule 112A" & Chr(13)
        validate_115AD = False
    End If
    
    If Not chkMandatory_115ad(TmpAr(1)) Then
        MsgBox_115AD = MsgBox_115AD + "Please enter the ISIN Code. at Sr. No." & n & " in Schedule 115AD(1)(iii)(p)" & Chr(13)
        validate_115AD = False
    End If
        
    If Len(TmpAr(1)) > 0 Then
        If Len(TmpAr(1)) <> 12 Then
            MsgBox_115AD = MsgBox_115AD + "Please enter the valid ISIN Code. at Sr. No." & n & " in Schedule 115AD(1)(iii)(p)" & Chr(13)
            validate_115AD = False
        End If
    End If
        
    If UCase(Trim(Mid(TmpAr(1), 1, 2))) <> "IN" Then
        MsgBox_115AD = MsgBox_115AD + "ISIN Code. at Sr. No." & n & " should start with ""IN"" in Schedule 115AD(1)(iii)(p)" & Chr(13)
        validate_115AD = False
    End If
    
    If Not chkMandatory_115ad(TmpAr(2)) Then
        MsgBox_115AD = MsgBox_115AD + "Please enter Name of the Share/Unit at Sr. No." & n & " in Schedule 115AD(1)(iii)(p)" & Chr(13)
        validate_115AD = False
    End If
        
    If Len(TmpAr(2)) > 125 Then
        MsgBox_115AD = MsgBox_115AD + "Name of the Share/Unit at Sr. No." & n & " in Schedule 115AD(1)(iii)(p)" & Chr(13)
        validate_115AD = False
    End If
    
    If Not chkMandatory_115ad(TmpAr(3)) Then
        MsgBox_115AD = MsgBox_115AD + "Please enter the No. of Shares/Units at Sr. No." & n & " in Schedule 115AD(1)(iii)(p)" & Chr(13)
        validate_115AD = False
    End If
        
    If CDec(TmpAr(3)) < 0 Then
        MsgBox_115AD = MsgBox_115AD + "* No. of Shares/Units at Sr. No." & n & " cannot be negetive in Schedule 115AD(1)(iii) proviso" & Chr(13)
        validate_115AD = False
    End If
        
    tmpArray = Split(TmpAr(3), ".")
    If Len(TmpAr(3)) > 19 Or Len(tmpArray(0)) > 14 Then
        MsgBox_115AD = MsgBox_115AD + "* No. of Shares/Units at Sr. No." & n & " cannot exceed 19 digits in Schedule 115AD(1)(iii) proviso" & Chr(13)
        validate_115AD = False
    End If
    
    If Not chkMandatory_115ad(TmpAr(4)) Then
        MsgBox_115AD = MsgBox_115AD + "Please enter the Sale-price per Share/Unit at Sr. No." & i & " in Schedule 115AD(1)(iii)(p)" & Chr(13)
        validate_115AD = False
    End If
        
    If CDec(TmpAr(4)) < 0 Then
        MsgBox_115AD = MsgBox_115AD + "* Sale-price per Share/Unit at Sr. No." & i & " cannot be negetive in Schedule 115AD(1)(iii) proviso" & Chr(13)
        validate_115AD = False
    End If
        
    tmpArray = Split(TmpAr(4), ".")
    If Len(TmpAr(4)) > 19 Or Len(tmpArray(0)) > 14 Then
        MsgBox_115AD = MsgBox_115AD + "* Sale-price per Share/Unit at Sr. No." & i & " cannot exceed 19 digits in Schedule 115AD(1)(iii) proviso" & Chr(13)
        validate_115AD = False
    End If
    
    If Not chkMandatory_115ad(TmpAr(5)) Then
        MsgBox_115AD = MsgBox_115AD + "Please enter the Total Sale Value at Sr. No." & n & " in Schedule 115AD(1)(iii) proviso" & Chr(13)
        validate_115AD = False
    End If
    
    If Len(TmpAr(5)) > 14 Then
        MsgBox_115AD = MsgBox_115AD + "Total Sale Value at Sr. No " & i & " cannot exceed 14 digits." & " in Schedule 115AD(1)(iii)(p)" & Chr(13)
        validate_115AD = False
    End If
    
    If CDec(TmpAr(5)) < 0 Then
        MsgBox_115AD = MsgBox_115AD + "* Total Sale Value at Sr. No." & n & " cannot be negative in Schedule 115AD(1)(iii) proviso" & Chr(13)
        validate_115AD = False
    End If
    
    If Not chkMandatory_115ad(TmpAr(7)) Then
        MsgBox_115AD = MsgBox_115AD + "Please enter the Cost of acquisition at Sr. No." & n & " in Schedule 115AD(1)(iii) proviso" & Chr(13)
        validate_115AD = False
    End If
    
    If CDec(TmpAr(7)) < 0 Then
        MsgBox_115AD = MsgBox_115AD + "* Cost of acquisition at Sr. No." & n & " cannot be negative in Schedule 115AD(1)(iii) proviso" & Chr(13)
        validate_115AD = False
    End If
    
    If Not chkMandatory_115ad(TmpAr(9)) Then
        MsgBox_115AD = MsgBox_115AD + "Please enter the Fair Market Value per share/unit as on 31st January,2018  If share is Aquired after 31st Jan 2018 , please enter '0' at Sr. No." & i & " in Schedule 115AD(1)(iii)(p)" & Chr(13)
        validate_115AD = False
    End If
        
    If CDec(TmpAr(9)) < 0 Then
        MsgBox_115AD = MsgBox_115AD + "* Fair Market Value per share/unit as on 31st January,2018 in Column 10,If share is Acquired after 31st Jan 2018 , please enter '0' at Sr. No." & i & " cannot be negetive in Schedule 115AD(1)(iii) proviso" & Chr(13)
        validate_115AD = False
    End If
        
    tmpArray = Split(TmpAr(9), ".")
    If Len(TmpAr(9)) > 19 Or Len(tmpArray(0)) > 14 Then
        MsgBox_115AD = MsgBox_115AD + "* Fair Market Value per share/unit as on 31st January,2018 in Column 10,If share is Acquired after 31st Jan 2018 , please enter '0' at Sr. No." & i & " cannot exceed 19 digits in Schedule 115AD(1)(iii) proviso" & Chr(13)
        validate_115AD = False
    End If
    
    If Not chkMandatory_115ad(TmpAr(11)) Then
        MsgBox_115AD = MsgBox_115AD + "Please enter the Expenditure incurred wholly and exclusively in connection with transfer at Sr. No." & i & " in Schedule 115AD(1)(iii)(p)" & Chr(13)
        validate_115AD = False
    End If
    
    If CDec(TmpAr(11)) < 0 Then
        MsgBox_115AD = MsgBox_115AD + "* Expenditure incurred wholly and exclusively in connection at Sr. No." & i & " cannot be negative in Schedule 115AD(1)(iii)(p)" & Chr(13)
        validate_115AD = False
    End If
End Function

Function AddDiffRows_115AD(DiffRows As Long)
    setTblinfo_115AD
    Sheet54.Activate
    SelectLastRow ("ShareAcq_115AD_1")
    insertRowUnderSectionWithFormula (DiffRows)
    Call ExendRangeNameToTable(DiffRows, rngname_115AD)
End Function

Function setTblinfo_115AD()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("ShareAcq_115AD_1").count
    Set rangecells = Range("ShareAcq_115AD_1").Cells
    Dim natureAmount As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).value = "" Or rangecells.item(mIntCtr).value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
    ColCount_115AD = ccount
    rngname_115AD = "ShareAcq_115AD_1;ISINCode_115AD;NameofShare_115AD;NoofShare_115AD;SalePriceperShare_115AD;TotalSaleValue_115AD;COAwithoutIndex_115AD;COAwithIndex_115AD;LTCGAssetAcquired_115AD;FairMarketValue_115AD;TotalFairMarketValue_115AD;ExpenditureWholly_115AD;TotalDeductions_115AD;Balance_115AD;"
End Function

Function chkMandatory_115ad(field As Variant) As Boolean
chkMandatory_115ad = True
If Len(Trim(field)) <= 0 Then
    chkMandatory_115ad = False
End If
If IsEmpty(field) Then
chkMandatory_115ad = False
End If
End Function
