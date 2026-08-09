Attribute VB_Name = "md115AD"
Option Explicit

Public MsgBox_115AD As String

Public TotalSaleValue_115AD_1T As Variant
Public ShareAcq_115AD_1 As Variant
'Public ShareTrans_115AD_1 As Variant   'commented by Chetan C M on 20/01/2026 for AY 2026-27
Public ISINCode_115AD As Variant
Public NameofShare_115AD As Variant
Public NoofShare_115AD As Variant
Public SalePriceperShare_115AD As Variant
Public TotalSaleValue_115AD As Variant
Public COAwithoutIndex_115AD As Variant
Public COAwithIndex_115AD As Variant
Public LTCGAssetAcquired_115AD As Variant
Public FairMarketValue_115AD As Variant
Public TotalFairMarketValue_115AD As Variant
'Public FullValueConsideration_115AD As Variant
'Public CostofImprovement_115AD As Variant
Public ExpenditureWholly_115AD As Variant
Public TotalDeductions_115AD As Variant
Public Balance_115AD As Variant

Public rngname_115AD As Variant
Public Total_TotalSaleValue_115AD As Variant
Public Total_COAwithoutIndex_115AD As Variant
Public Total_COAwithIndex_115AD As Variant
Public Total_LTCGAssetAcquired_115AD As Variant
Public Total_TotalFairMarketValue_115AD As Variant
'Public Total_FullValueConsideration_115AD As Variant
'Public Total_CostofImprovement_115AD As Variant
Public Total_ExpenditureWholly_115AD As Variant
Public Total_TotalDeductions_115AD As Variant
Public Total_Balance_115AD As Variant

'commented by Chetan C M on 20/01/2026 for AY 2026-27
    'start--
''added by Chetan C M on 03/03/2025
''start--
'Public Total_Balance_115ADBE As Variant
'Public Total_Balance_115ADAE As Variant
'Public TotBalance_115_ABE_AAE As Variant    'added by Chetan C M on 29/10/2025 for SIT-89185
''--end
    '--end
    
Public end115AD, end115AD1, end115AD2, end115AD3, end115AD4, end115AD5, end115AD7, end115AD8, end115AD9, end115AD10 As Long

Sub ValidateSheet115AD_Click()
validate115AD
fmsgboxStatus " Schedule 115AD(1)(iii) proviso is ok" ', vbOKOnly, "ITR-3"
End Sub

Sub validate115AD()
    If Not Validate115AD_1 Then
        Sheet54.Activate
        fmsgboxsmall MsgBox_115AD ', vbOKOnly, "Error(s)"
        CloseMsg
    End If
End Sub
Sub AddRows115AD()
Dim vRows  As Long
Sheets("Schedule 115AD(1)(iii) proviso").Activate

'changed by Chetan C M on 20/01/2026 for AY 2026-27
    'start--
'EfilingCommon.DefinedgridNameRange = "ShareAcq_115AD_1||ShareTrans_115AD_1||ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||TotalSaleValue_115AD||COAwithoutIndex_115AD||COAwithIndex_115AD||LTCGAssetAcquired_115AD||FairMarketValue_115AD||TotalFairMarketValue_115AD||ExpenditureWholly_115AD||TotalDeductions_115AD||Balance_115AD"
EfilingCommon.DefinedgridNameRange = "ShareAcq_115AD_1||ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||TotalSaleValue_115AD||COAwithoutIndex_115AD||COAwithIndex_115AD||LTCGAssetAcquired_115AD||FairMarketValue_115AD||TotalFairMarketValue_115AD||ExpenditureWholly_115AD||TotalDeductions_115AD||Balance_115AD"
    '--end

ActiveCellRange = EfilingCommon.searchLastRow("ISINCode_115AD")
EfilingCommon.insertRowUnderSectionWithFormula f_112A_115AD_1:=1
End Sub
Sub setTableInfo115AD()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet54.Range("ISINCode_115AD").Cells
    mIntCells = Sheet54.Range("ISINCode_115AD").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end115AD = ccount
    
'changed by Chetan C M on 20/01/2026 for AY 2026-27
    'start--
'    rngname_115AD = "ShareAcq_115AD_1;ShareTrans_115AD_1;ISINCode_115AD;NameofShare_115AD;NoofShare_115AD;SalePriceperShare_115AD;TotalSaleValue_115AD;COAwithoutIndex_115AD;COAwithIndex_115AD;LTCGAssetAcquired_115AD;FairMarketValue_115AD;TotalFairMarketValue_115AD;ExpenditureWholly_115AD;TotalDeductions_115AD;Balance_115AD;"
    rngname_115AD = "ShareAcq_115AD_1;ISINCode_115AD;NameofShare_115AD;NoofShare_115AD;SalePriceperShare_115AD;TotalSaleValue_115AD;COAwithoutIndex_115AD;COAwithIndex_115AD;LTCGAssetAcquired_115AD;FairMarketValue_115AD;TotalFairMarketValue_115AD;ExpenditureWholly_115AD;TotalDeductions_115AD;Balance_115AD;"
    '--end

End Sub
Sub setTableInfo115AD8()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet54.Range("ShareAcq_115AD_1").Cells
    mIntCells = Sheet54.Range("ShareAcq_115AD_1").count
    For mIntCtr = 1 To mIntCells
        If Not (rangecells.item(mIntCtr).value = "(Select)" Or rangecells.item(mIntCtr).value = "") And rangecells.item(mIntCtr).Locked = False Then
        ccount = ccount + 1
        End If
    Next
    end115AD8 = ccount
End Sub
Sub setTableInfo115AD9()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet54.Range("TotalSaleValue_115AD").Cells
    mIntCells = Sheet54.Range("TotalSaleValue_115AD").count
    For mIntCtr = 1 To mIntCells
        If rangecells.item(mIntCtr).value = "" And rangecells.item(mIntCtr).Locked = False Then
        ccount = ccount + 1
        End If
    Next
    end115AD9 = ccount
End Sub
Function ValidateTotalSaleValue_115AD_1T() As Boolean
    ValidateTotalSaleValue_115AD_1T = True
    Dim rangecells As Range
    Set rangecells = Sheet54.Range("TotalSaleValue_115AD").Cells
    Dim i As Long
    ReDim TotalSaleValue_115AD_1T(end115AD)
    
    For i = 1 To end115AD
        TotalSaleValue_115AD_1T(i) = rangecells.item(i).value
        
        If TotalSaleValue_115AD_1T(i) = "" Then
            MsgBox_115AD = MsgBox_115AD + "* Please enter Total Sale Value at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateTotalSaleValue_115AD_1T = False
            Exit Function
        End If
    Next
End Function
Function ValidateShareAcq_115AD_1() As Boolean
    ValidateShareAcq_115AD_1 = True
    Dim rangecells As Range
    Set rangecells = Sheet54.Range("ShareAcq_115AD_1").Cells
    Dim i As Long
    ReDim ShareAcq_115AD_1(end115AD)
    
    For i = 1 To end115AD
        ShareAcq_115AD_1(i) = rangecells.item(i).value
          'Malli----SIT-89058
        If ShareAcq_115AD_1(i) = "(Select)" Or ShareAcq_115AD_1(i) = "" Then
           ' MsgBox_115AD = MsgBox_115AD + "* Please Select Whether Share acquired on or before / After 31st January,2018 at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
            ' MsgBox_115AD = MsgBox_115AD + "* """Please select the dropdown option in Column 1a of Schedule 115AD(1)(iii) proviso at Sr. No." & i & """ & Chr(13)
            
        'Added by Shrutika(04/03/2025)
             MsgBox_115AD = MsgBox_115AD + "* Please select the dropdown option in Column 1a of Schedule 115AD(1)(iii) proviso at Sr. No." & i & "" & Chr(13)
            ValidateShareAcq_115AD_1 = False
            Exit Function
        End If
    Next
End Function
Sub setTableInfo115AD1()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet54.Range("NameofShare_115AD").Cells
    mIntCells = Sheet54.Range("NameofShare_115AD").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end115AD1 = ccount
End Sub

Sub setTableInfo115AD2()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet54.Range("NoofShare_115AD").Cells
    mIntCells = Sheet54.Range("NoofShare_115AD").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end115AD2 = ccount
End Sub

Sub setTableInfo115AD3()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet54.Range("SalePriceperShare_115AD").Cells
    mIntCells = Sheet54.Range("SalePriceperShare_115AD").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end115AD3 = ccount
End Sub

Sub setTableInfo115AD4()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet54.Range("COAwithIndex_115AD").Cells
    mIntCells = Sheet54.Range("COAwithIndex_115AD").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end115AD4 = ccount
End Sub

Sub setTableInfo115AD5()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet54.Range("FairMarketValue_115AD").Cells
    mIntCells = Sheet54.Range("FairMarketValue_115AD").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end115AD5 = ccount
End Sub


Sub setTableInfo115AD7()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet54.Range("ExpenditureWholly_115AD").Cells
    mIntCells = Sheet54.Range("ExpenditureWholly_115AD").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end115AD7 = ccount
End Sub

Function ValidateISINCode_115AD() As Boolean
    ValidateISINCode_115AD = True
    Dim rangecells As Range
    Set rangecells = Sheet54.Range("ISINCode_115AD").Cells
    Dim i As Long
    ReDim ISINCode_115AD(end115AD)
    
    subProcCaption = "Validating 115AD"
    noOfProcessSub = end115AD
    
    For i = 1 To end115AD
        ISINCode_115AD(i) = rangecells.item(i).value
        
        If Not chkMandatory(ISINCode_115AD(i)) Then
            MsgBox_115AD = MsgBox_115AD + "* Please enter the ISIN Code. at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateISINCode_115AD = False
            Exit Function
        End If
        
        If Len(ISINCode_115AD(i)) > 0 Then
            If Len(ISINCode_115AD(i)) <> 12 Then
                MsgBox_115AD = MsgBox_115AD + "* Please enter the valid ISIN Code. at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
                ValidateISINCode_115AD = False
                Exit Function
            End If
        End If
        
        If UCase(Trim(Mid(ISINCode_115AD(i), 1, 2))) <> "IN" Then
            MsgBox_115AD = MsgBox_115AD + "* Please enter the valid ISIN Code. at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateISINCode_115AD = False
            Exit Function
        End If
    Next
End Function

Function ValidateNameofShare_115AD() As Boolean
    ValidateNameofShare_115AD = True
    Dim rangecells As Range
    Set rangecells = Sheet54.Range("NameofShare_115AD").Cells
    Dim i As Long
    ReDim NameofShare_115AD(end115AD)
    
    For i = 1 To end115AD
        NameofShare_115AD(i) = rangecells.item(i).value
        
        If Not chkMandatory(NameofShare_115AD(i)) Then
            MsgBox_115AD = MsgBox_115AD + "* Please enter Name of the Share/Unit at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateNameofShare_115AD = False
            Exit Function
        End If
        
        If Len(NameofShare_115AD(i)) > 125 Then
            MsgBox_115AD = MsgBox_115AD + "* Name of the Share/Unit at Sr. No." & i & " should not exceed 125 characters.in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateNameofShare_115AD = False
            Exit Function
        End If
    Next
End Function

Function ValidateNoofShare_115AD() As Boolean
    ValidateNoofShare_115AD = True
    Dim rangecells As Range
    Set rangecells = Sheet54.Range("NoofShare_115AD").Cells
    Dim i As Long
    ReDim NoofShare_115AD(end115AD)
    'added by Chetan C M on 25/10/2025 for SIT-89185=======start
     Dim int_num As String, decimal_num As String, find_num As Integer
    'end======
    For i = 1 To end115AD
        NoofShare_115AD(i) = rangecells.item(i).value
        
        If Not chkMandatory(NoofShare_115AD(i)) Then
            MsgBox_115AD = MsgBox_115AD + "* Please enter the No. of Shares/Units at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateNoofShare_115AD = False
            Exit Function
        End If
        
        'added by Chetan C M on 25/10/2025 for SIT-89185
        'start--
        Dim tmpArray
        If CDec(NoofShare_115AD(i)) < 0 Then
            MsgBox_115AD = MsgBox_115AD + "* No. of Shares/Units at Sr. No." & i & " cannot be negetive in Schedule 115AD(1)(iii) proviso" & Chr(13)
            ValidateNoofShare_115AD = False
            Exit Function
        End If
        
        tmpArray = Split(NoofShare_115AD(i), ".")
        If Len(NoofShare_115AD(i)) > 19 Or Len(tmpArray(0)) > 14 Then
            MsgBox_115AD = MsgBox_115AD + "* No. of Shares/Units at Sr. No." & i & " cannot exceed 19 digits in Schedule 115AD(1)(iii) proviso" & Chr(13)
            ValidateNoofShare_115AD = False
            Exit Function
        End If
        
        find_num = InStr(1, NoofShare_115AD(i), ".")

        If find_num = 1 Then
            MsgBox_115AD = MsgBox_115AD + "* No. of Shares/Units should start with a number." & Chr(13)
            ValidateNoofShare_115AD = False
            Exit Function
        End If

        If find_num > 0 Then
            int_num = Left(NoofShare_115AD(i), find_num - 1)
            decimal_num = Mid(NoofShare_115AD(i), find_num + 1)
        Else
            int_num = NoofShare_115AD(i)
            decimal_num = ""
        End If

        If Not IsNumeric(int_num) Then
            MsgBox_115AD = MsgBox_115AD + "* No. of Shares/Units should be a numeric value." & Chr(13)
            ValidateNoofShare_115AD = False
            Exit Function
        End If

        If Len(int_num) > 14 Or val(int_num) > 99999999999999# Then
            MsgBox_115AD = MsgBox_115AD + "* No. of Shares/Units cannot exceed 14 digits before decimal." & Chr(13)
            ValidateNoofShare_115AD = False
            Exit Function
        End If

        If Len(decimal_num) > 0 Then
            If Not IsNumeric(decimal_num) Then
                MsgBox_115AD = MsgBox_115AD + "* No. of Shares/Units should be a numeric value after decimal." & Chr(13)
                ValidateNoofShare_115AD = False
                Exit Function
            End If

            If Len(decimal_num) > 4 Then
                MsgBox_115AD = MsgBox_115AD + "* No. of Shares/Units cannot exceed 4 digits after decimal." & Chr(13)
                ValidateNoofShare_115AD = False
                Exit Function
            End If
        End If
        
        '--end
    Next
End Function

Function ValidateSalePriceperShare_115AD() As Boolean
    ValidateSalePriceperShare_115AD = True
    Dim rangecells As Range
    Set rangecells = Sheet54.Range("SalePriceperShare_115AD").Cells
    Dim i As Long
    ReDim SalePriceperShare_115AD(end115AD)
    'added by Chetan C M on 25/10/2025 for SIT-89185=======start
    Dim int_num As String, decimal_num As String, find_num As Integer
    '====end
    
    For i = 1 To end115AD
        SalePriceperShare_115AD(i) = rangecells.item(i).value
        
        If Not chkMandatory(SalePriceperShare_115AD(i)) Then
            MsgBox_115AD = MsgBox_115AD + "* Please enter the Sale-price per Share/Unit at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateSalePriceperShare_115AD = False
            Exit Function
        End If
        
        'added by Chetan C M on 25/10/2025 for SIT-89185
        'start--
        Dim tmpArray
        If CDec(SalePriceperShare_115AD(i)) < 0 Then
            MsgBox_115AD = MsgBox_115AD + "* Sale-price per Share/Unit at Sr. No." & i & " cannot be negetive in Schedule 115AD(1)(iii) proviso" & Chr(13)
            ValidateSalePriceperShare_115AD = False
            Exit Function
        End If
        
        tmpArray = Split(SalePriceperShare_115AD(i), ".")
        If Len(SalePriceperShare_115AD(i)) > 19 Or Len(tmpArray(0)) > 14 Then
            MsgBox_115AD = MsgBox_115AD + "* Sale-price per Share/Unit at Sr. No." & i & " cannot exceed 19 digits in Schedule 115AD(1)(iii) proviso" & Chr(13)
            ValidateSalePriceperShare_115AD = False
            Exit Function
        End If
        
        find_num = InStr(1, SalePriceperShare_115AD(i), ".")

        If find_num = 1 Then
            MsgBox_115AD = MsgBox_115AD + "* Sale-price per Share/Unit should starts with number." & Chr(13)
            ValidateSalePriceperShare_115AD = False
            Exit Function
        End If

        If find_num > 0 Then
            int_num = Left(SalePriceperShare_115AD(i), find_num - 1)
            decimal_num = Mid(SalePriceperShare_115AD(i), find_num + 1)
        Else
            int_num = SalePriceperShare_115AD(i)
            decimal_num = ""
        End If

        If Not IsNumeric(int_num) Then
            MsgBox_115AD = MsgBox_115AD + "* Sale-price per Share/Unit should be a numeric value." & Chr(13)
            ValidateSalePriceperShare_115AD = False
            Exit Function
        End If

        If Len(int_num) > 14 Or val(int_num) > 99999999999999# Then
            MsgBox_115AD = MsgBox_115AD + "* Sale-price per Share/Unit cannot exceed 14 digits before decimal." & Chr(13)
            ValidateSalePriceperShare_115AD = False
            Exit Function
        End If

        If Len(decimal_num) > 0 Then
            If Not IsNumeric(decimal_num) Then
                MsgBox_115AD = MsgBox_115AD + "* Sale-price per Share/Unit should be a numeric value after decimal." & Chr(13)
                ValidateSalePriceperShare_115AD = False
                Exit Function
            End If

            If Len(decimal_num) > 4 Then
                MsgBox_115AD = MsgBox_115AD + "* Sale-price per Share/Unit cannot exceed 4 digits after decimal." & Chr(13)
                ValidateSalePriceperShare_115AD = False
                Exit Function
            End If
        End If

        '--end
        
    Next
End Function

Function ValidateCOAwithIndex_115AD() As Boolean
    ValidateCOAwithIndex_115AD = True
    Dim rangecells As Range
    Set rangecells = Sheet54.Range("COAwithIndex_115AD").Cells
    Dim i As Long
    ReDim COAwithIndex_115AD(end115AD)
    'added by Chetan C M on 25/10/2025 for SIT-89185=======start
    Dim int_num As String, decimal_num As String, find_num As Integer
    '===end
    
    For i = 1 To end115AD
        COAwithIndex_115AD(i) = rangecells.item(i).value
        
        If Not chkMandatory(COAwithIndex_115AD(i)) Then
            MsgBox_115AD = MsgBox_115AD + "* Please enter the Cost of acquisition at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateCOAwithIndex_115AD = False
            Exit Function
        End If
        
    'added by Chetan C M on 25/10/2025 for SIT-89185=======start
        find_num = InStr(1, COAwithIndex_115AD(i), ".")

        If find_num = 1 Then
            MsgBox_115AD = MsgBox_115AD + "* Cost of acquisition should start with a number." & Chr(13)
            ValidateCOAwithIndex_115AD = False
            Exit Function
        End If

        If find_num > 0 Then
            int_num = Left(COAwithIndex_115AD(i), find_num - 1)
            decimal_num = Mid(COAwithIndex_115AD(i), find_num + 1)
        Else
            int_num = COAwithIndex_115AD(i)
            decimal_num = ""
        End If

        If Not IsNumeric(int_num) Then
            MsgBox_115AD = MsgBox_115AD + "* Cost of acquisition should be a numeric value." & Chr(13)
            ValidateCOAwithIndex_115AD = False
            Exit Function
        End If

        If Len(int_num) > 14 Or val(int_num) > 99999999999999# Then
            MsgBox_115AD = MsgBox_115AD + "* Cost of acquisition cannot exceed 14 digits before decimal." & Chr(13)
            ValidateCOAwithIndex_115AD = False
            Exit Function
        End If

        If Len(decimal_num) > 0 Then
            If Not IsNumeric(decimal_num) Then
                MsgBox_115AD = MsgBox_115AD + "* Cost of acquisition should be a numeric value after decimal." & Chr(13)
                ValidateCOAwithIndex_115AD = False
                Exit Function
            End If

            If Len(decimal_num) > 4 Then
                MsgBox_115AD = MsgBox_115AD + "* Cost of acquisition cannot exceed 4 digits after decimal." & Chr(13)
                ValidateCOAwithIndex_115AD = False
                Exit Function
            End If
        End If
    '===end
        
    Next
End Function

Function ValidateFairMarketValue_115AD() As Boolean
    ValidateFairMarketValue_115AD = True
    Dim rangecells As Range
    Set rangecells = Sheet54.Range("FairMarketValue_115AD").Cells
    Dim i As Long
    ReDim FairMarketValue_115AD(end115AD)
     'added by Chetan C M on 25/10/2025 for SIT-89185=======start
    Dim int_num As String, decimal_num As String, find_num As Integer
    '===end
    
    For i = 1 To end115AD
        FairMarketValue_115AD(i) = rangecells.item(i).value
        
        If Not chkMandatory(FairMarketValue_115AD(i)) Then
            MsgBox_115AD = MsgBox_115AD + "* Please enter the Fair Market Value per share/unit as on 31st January,2018 .115AD(1)(b)(iii) Proviso . If share is Aquired after 31st Jan 2018 , please enter ""0"" at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateFairMarketValue_115AD = False
            Exit Function
        End If
        
        'added by Chetan C M on 25/10/2025 for SIT-89185
        'start--
        Dim tmpArray
        If CDec(FairMarketValue_115AD(i)) < 0 Then
            MsgBox_115AD = MsgBox_115AD + "* Fair Market Value per share/unit as on 31st January,2018 in Column 10,If share is Acquired after 31st Jan 2018 , please enter '0' at Sr. No." & i & " cannot be negetive in Schedule 115AD(1)(iii) proviso" & Chr(13)
            ValidateFairMarketValue_115AD = False
            Exit Function
        End If
        
        tmpArray = Split(FairMarketValue_115AD(i), ".")
        If Len(FairMarketValue_115AD(i)) > 19 Or Len(tmpArray(0)) > 14 Then
            MsgBox_115AD = MsgBox_115AD + "* Fair Market Value per share/unit as on 31st January,2018 in Column 10,If share is Acquired after 31st Jan 2018 , please enter '0' at Sr. No." & i & " cannot exceed 19 digits in Schedule 115AD(1)(iii) proviso" & Chr(13)
            ValidateFairMarketValue_115AD = False
            Exit Function
        End If
        
        find_num = InStr(1, FairMarketValue_115AD(i), ".")

        If find_num = 1 Then
            MsgBox_115AD = MsgBox_115AD + "* Fair Market Value per share/unit should start with a number." & Chr(13)
            ValidateFairMarketValue_115AD = False
            Exit Function
        End If

        If find_num > 0 Then
            int_num = Left(FairMarketValue_115AD(i), find_num - 1)
            decimal_num = Mid(FairMarketValue_115AD(i), find_num + 1)
        Else
            int_num = FairMarketValue_115AD(i)
            decimal_num = ""
        End If

        If Not IsNumeric(int_num) Then
            MsgBox_115AD = MsgBox_115AD + "* Fair Market Value per share/unit should be a numeric value." & Chr(13)
            ValidateFairMarketValue_115AD = False
            Exit Function
        End If

        If Len(int_num) > 14 Or val(int_num) > 99999999999999# Then
            MsgBox_115AD = MsgBox_115AD + "* Fair Market Value per share/unit cannot exceed 14 digits before decimal." & Chr(13)
            ValidateFairMarketValue_115AD = False
            Exit Function
        End If

        If Len(decimal_num) > 0 Then
            If Not IsNumeric(decimal_num) Then
                MsgBox_115AD = MsgBox_115AD + "* Fair Market Value per share/unit should be a numeric value after decimal." & Chr(13)
                ValidateFairMarketValue_115AD = False
                Exit Function
            End If

            If Len(decimal_num) > 4 Then
                MsgBox_115AD = MsgBox_115AD + "* Fair Market Value per share/unit cannot exceed 4 digits after decimal." & Chr(13)
                ValidateFairMarketValue_115AD = False
                Exit Function
            End If
        End If

        '--end
        
    Next
End Function

Function ValidateExpenditureWholly_115AD() As Boolean
    ValidateExpenditureWholly_115AD = True
    Dim rangecells As Range
    Set rangecells = Sheet54.Range("ExpenditureWholly_115AD").Cells
    Dim i As Long
    ReDim ExpenditureWholly_115AD(end115AD)
    'added by Chetan C M on 25/10/2025 for SIT-89185=======start
    Dim int_num As String, decimal_num As String, find_num As Integer
    '====end
    
    For i = 1 To end115AD
        ExpenditureWholly_115AD(i) = rangecells.item(i).value
        
        If Not chkMandatory(ExpenditureWholly_115AD(i)) Then
            MsgBox_115AD = MsgBox_115AD + "* Please enter the Expenditure incurred wholly and exclusively in connection with transfer at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateExpenditureWholly_115AD = False
            Exit Function
        End If
        
        'added by Chetan C M on 25/10/2025 for SIT-89185=======start
        find_num = InStr(1, ExpenditureWholly_115AD(i), ".")

        If find_num = 1 Then
            MsgBox_115AD = MsgBox_115AD + "* Expenditure incurred wholly and exclusively in connection with transfer should start with a number." & Chr(13)
            ValidateExpenditureWholly_115AD = False
            Exit Function
        End If

        If find_num > 0 Then
            int_num = Left(ExpenditureWholly_115AD(i), find_num - 1)
            decimal_num = Mid(ExpenditureWholly_115AD(i), find_num + 1)
        Else
            int_num = ExpenditureWholly_115AD(i)
            decimal_num = ""
        End If

        If Not IsNumeric(int_num) Then
            MsgBox_115AD = MsgBox_115AD + "* Expenditure incurred wholly and exclusively in connection with transfer should be a numeric value." & Chr(13)
            ValidateExpenditureWholly_115AD = False
            Exit Function
        End If

        If Len(int_num) > 14 Or val(int_num) > 99999999999999# Then
            MsgBox_115AD = MsgBox_115AD + "* Expenditure incurred wholly and exclusively in connection with transfer cannot exceed 14 digits before decimal." & Chr(13)
            ValidateExpenditureWholly_115AD = False
            Exit Function
        End If

        If Len(decimal_num) > 0 Then
            If Not IsNumeric(decimal_num) Then
                MsgBox_115AD = MsgBox_115AD + "* Expenditure incurred wholly and exclusively in connection with transfer should be a numeric value after decimal." & Chr(13)
                ValidateExpenditureWholly_115AD = False
                Exit Function
            End If

            If Len(decimal_num) > 4 Then
                MsgBox_115AD = MsgBox_115AD + "* Expenditure incurred wholly and exclusively in connection with transfer cannot exceed 4 digits after decimal." & Chr(13)
                ValidateExpenditureWholly_115AD = False
                Exit Function
            End If
        End If

        '===end
        
    Next
End Function

Function ValidateNonEditableFields_115AD() As Boolean
    ValidateNonEditableFields_115AD = True
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Dim rangecells4 As Range
    Dim rangecells5 As Range
    Dim rangecells6 As Range
    Dim tmpArray() As String
    
    Set rangecells = Sheet54.Range("TotalSaleValue_115AD").Cells
    Set rangecells1 = Sheet54.Range("COAwithoutIndex_115AD").Cells
    Set rangecells2 = Sheet54.Range("LTCGAssetAcquired_115AD").Cells
    Set rangecells3 = Sheet54.Range("TotalFairMarketValue_115AD").Cells
'    Set rangecells4 = Sheet54.Range("FullValueConsideration_115AD").Cells
    Set rangecells5 = Sheet54.Range("TotalDeductions_115AD").Cells
    Set rangecells6 = Sheet54.Range("Balance_115AD").Cells
    
    Dim i As Long
    ReDim TotalSaleValue_115AD(end115AD)
    ReDim COAwithoutIndex_115AD(end115AD)
    ReDim LTCGAssetAcquired_115AD(end115AD)
    ReDim TotalFairMarketValue_115AD(end115AD)
'    ReDim FullValueConsideration_115AD(end115AD)
    ReDim TotalDeductions_115AD(end115AD)
    ReDim Balance_115AD(end115AD)
    
    For i = 1 To end115AD
        TotalSaleValue_115AD(i) = rangecells.item(i).value
        
        If Len(TotalSaleValue_115AD(i)) > 0 Then
            TotalSaleValue_115AD(i) = Application.WorksheetFunction.Round(TotalSaleValue_115AD(i), 0)
        End If
        
        'Tmparray = Split(TotalSaleValue_115AD(i), ".")
        If Len(TotalSaleValue_115AD(i)) > 14 Then ' Or Len(Tmparray(0)) > 11 Then
        'changed by Chetan C M on 29/10/2025 for SIT-89185
        'start--
'            MsgBox_115AD = MsgBox_115AD + "* Total sale Value (4*5) cannot be more than 14 digits at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
            MsgBox_115AD = MsgBox_115AD + "* Full Value of Consideration cannot be more than 14 digits at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
        '--end
            ValidateNonEditableFields_115AD = False
            'Exit Function
        End If

        COAwithoutIndex_115AD(i) = rangecells1.item(i).value
        
        
        If Len(COAwithoutIndex_115AD(i)) > 0 Then
        COAwithoutIndex_115AD(i) = Application.WorksheetFunction.Round(COAwithoutIndex_115AD(i), 0)
        End If
        
        'Tmparray = Split(COAwithoutIndex_115AD(i), ".")
        If Len(COAwithoutIndex_115AD(i)) > 14 Then ' Or Len(Tmparray(0)) > 11 Then
            MsgBox_115AD = MsgBox_115AD + "* Cost of acquisition without indexation cannot be more than 14 digits at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateNonEditableFields_115AD = False
            'Exit Function
        End If
        
        LTCGAssetAcquired_115AD(i) = rangecells2.item(i).value
        
        If Len(LTCGAssetAcquired_115AD(i)) > 0 Then
        LTCGAssetAcquired_115AD(i) = Application.WorksheetFunction.Round(LTCGAssetAcquired_115AD(i), 0)
        End If
        
        'Tmparray = Split(LTCGAssetAcquired_115AD(i), ".")
        If Len(LTCGAssetAcquired_115AD(i)) > 14 Then ' Or Len(Tmparray(0)) > 11 Then
            MsgBox_115AD = MsgBox_115AD + "* Long term capital asset acquired before 01.02.2018 cannot be more than 14 digits at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateNonEditableFields_115AD = False
            'Exit Function
        End If
        
        
        
        TotalFairMarketValue_115AD(i) = rangecells3.item(i).value
        
         If Len(TotalFairMarketValue_115AD(i)) > 0 Then
        TotalFairMarketValue_115AD(i) = Application.WorksheetFunction.Round(TotalFairMarketValue_115AD(i), 0)
        End If
        
        
         'Tmparray = Split(TotalFairMarketValue_115AD(i), ".")
        If Len(TotalFairMarketValue_115AD(i)) > 14 Then ' Or Len(Tmparray(0)) > 11 Then
            MsgBox_115AD = MsgBox_115AD + "* Total Fair Market Value of capital asset cannot be more than 14 digits at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateNonEditableFields_115AD = False
            'Exit Function
        End If

        TotalDeductions_115AD(i) = rangecells5.item(i).value
        
        If Len(TotalDeductions_115AD(i)) > 0 Then
        TotalDeductions_115AD(i) = Application.WorksheetFunction.Round(TotalDeductions_115AD(i), 0)
        End If
        
                'Tmparray = Split(TotalDeductions_115AD(i), ".")
        If Len(TotalDeductions_115AD(i)) > 14 Then ' Or Len(Tmparray(0)) > 11 Then
            MsgBox_115AD = MsgBox_115AD + "* Total deductions cannot be more than 14 digits at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateNonEditableFields_115AD = False
            'Exit Function
        End If
        
        Balance_115AD(i) = rangecells6.item(i).value
        
         If Len(Balance_115AD(i)) > 0 Then
        Balance_115AD(i) = Application.WorksheetFunction.Round(Balance_115AD(i), 0)
        End If
        
'        'Tmparray = Split(Balance_115AD(i), ".")
'
'        If Len(Balance_115AD(i)) > 14 then ' Or Len(Tmparray(0)) > 11 Then
'            MsgBox_115AD = MsgBox_115AD + "* Total at Column number 14 cannot be more than 14 digits at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
'            ValidateNonEditableFields_115AD = False
'            Exit Function
'        End If
        
        If Not Balance_115AD(i) < 0 Then
            If Len(Balance_115AD(i)) > 14 Then ' Or Len(Tmparray(0)) > 11 Then
                MsgBox_115AD = MsgBox_115AD + "* Balance cannot be more than 14 digits at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
                ValidateNonEditableFields_115AD = False
                'Exit Function
            End If
        Else
            'Tmparray = Split(CStr((Balance_115AD(i) * -1)), ".")
            If Len((Balance_115AD(i) * -1)) > 14 Then ' Or Len(Tmparray(0)) > 11 Then
                MsgBox_115AD = MsgBox_115AD + "* Balance cannot be more than 14 digits at Sr. No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
                ValidateNonEditableFields_115AD = False
                'Exit Function
            End If
        End If
        
    Next

End Function

Function ValidateTotalFields_115AD() As Boolean
 ValidateTotalFields_115AD = True
 Dim tmpArray() As String
 
 Total_TotalSaleValue_115AD = Sheet54.Range("Total_TotalSaleValue_115AD").value
 
 If Len(Total_TotalSaleValue_115AD) > 0 Then
 Total_TotalSaleValue_115AD = Application.WorksheetFunction.Round(Total_TotalSaleValue_115AD, 0)
 End If
 
 'Tmparray = Split(Total_TotalSaleValue_115AD, ".")
 If Len(Total_TotalSaleValue_115AD) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
            MsgBox_115AD = MsgBox_115AD + "* Total of column number 6 cannot be more than 14 digits in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateTotalFields_115AD = False
'            Exit Function
 End If
 
 Total_COAwithoutIndex_115AD = Sheet54.Range("Total_COAwithoutIndex_115AD").value
 
 If Len(Total_COAwithoutIndex_115AD) > 0 Then
 Total_COAwithoutIndex_115AD = Application.WorksheetFunction.Round(Total_COAwithoutIndex_115AD, 0)
 End If
 
 'Tmparray = Split(Total_COAwithoutIndex_115AD, ".")
 If Len(Total_COAwithoutIndex_115AD) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
            MsgBox_115AD = MsgBox_115AD + "* Total of column number 7 cannot be more than 14 digits in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateTotalFields_115AD = False
'            Exit Function
 End If
 
 Total_COAwithIndex_115AD = Sheet54.Range("Total_COAwithIndex_115AD").value
 
  If Len(Total_COAwithIndex_115AD) > 0 Then
 Total_COAwithIndex_115AD = Application.WorksheetFunction.Round(Total_COAwithIndex_115AD, 0)
 End If
 
  'Tmparray = Split(Total_COAwithIndex_115AD, ".")
 If Len(Total_COAwithIndex_115AD) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
            MsgBox_115AD = MsgBox_115AD + "* Total of column number 8 cannot be more than 14 digits in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateTotalFields_115AD = False
'            Exit Function
 End If
 
 Total_LTCGAssetAcquired_115AD = Sheet54.Range("Total_LTCGAssetAcquired_115AD").value
 
 If Len(Total_LTCGAssetAcquired_115AD) > 0 Then
 Total_LTCGAssetAcquired_115AD = Application.WorksheetFunction.Round(Total_LTCGAssetAcquired_115AD, 0)
 End If
 
   'Tmparray = Split(Total_LTCGAssetAcquired_115AD, ".")
 If Len(Total_LTCGAssetAcquired_115AD) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
            MsgBox_115AD = MsgBox_115AD + "* Total of column number 9 cannot be more than 14 digits in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateTotalFields_115AD = False
'            Exit Function
 End If
 
 Total_TotalFairMarketValue_115AD = Sheet54.Range("Total_TotalFairMarketValue_115AD").value
 
 If Len(Total_TotalFairMarketValue_115AD) > 0 Then
 Total_TotalFairMarketValue_115AD = Application.WorksheetFunction.Round(Total_TotalFairMarketValue_115AD, 0)
 End If
 
    'Tmparray = Split(Total_TotalFairMarketValue_115AD, ".")
 If Len(Total_TotalFairMarketValue_115AD) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
            MsgBox_115AD = MsgBox_115AD + "* Total of column number 11 cannot be more than 14 digits in sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateTotalFields_115AD = False
'            Exit Function
 End If

 
 Total_ExpenditureWholly_115AD = Sheet54.Range("Total_ExpenditureWholly_115AD").value
 
 If Len(Total_ExpenditureWholly_115AD) > 0 Then
 Total_ExpenditureWholly_115AD = Application.WorksheetFunction.Round(Total_ExpenditureWholly_115AD, 0)
 End If
 
     'Tmparray = Split(Total_ExpenditureWholly_115AD, ".")
 If Len(Total_ExpenditureWholly_115AD) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
            MsgBox_115AD = MsgBox_115AD + "* Total of column number 12 cannot be more than 14 digits sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateTotalFields_115AD = False
'            Exit Function
 End If
 
 Total_TotalDeductions_115AD = Sheet54.Range("Total_TotalDeductions_115AD").value
 
 If Len(Total_TotalDeductions_115AD) > 0 Then
 Total_TotalDeductions_115AD = Application.WorksheetFunction.Round(Total_TotalDeductions_115AD, 0)
 End If
 
      'Tmparray = Split(Total_TotalDeductions_115AD, ".")
 If Len(Total_TotalDeductions_115AD) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
            MsgBox_115AD = MsgBox_115AD + "* Total of column number 13 cannot be more than 14 digits sheet 115AD(1)(iii) proviso" & Chr(13)
            ValidateTotalFields_115AD = False
'            Exit Function
 End If
 
 Total_Balance_115AD = Sheet54.Range("Total_Balance_115AD").value
 
 If Len(Total_Balance_115AD) > 0 Then
 Total_Balance_115AD = Application.WorksheetFunction.Round(Total_Balance_115AD, 0)
 End If
 
 'Tmparray = Split(Total_Balance_115AD, ".")
    If Not Total_Balance_115AD < 0 Then
        If Len(Total_Balance_115AD) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
            MsgBox_115AD = MsgBox_115AD + "* Total of column number 14 cannot be more than 14 digits in Schedule 115AD(1)(iii)(p)" & Chr(13)
            ValidateTotalFields_115AD = False
'            Exit Function
        End If
    Else
        'Tmparray = Split(CStr((Total_Balance_115AD * -1)), ".")
        If Len((Total_Balance_115AD * -1)) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
            MsgBox_115AD = MsgBox_115AD + "* Total of column number 14 cannot be more than 14 digits in Schedule 115AD(1)(iii)(p)" & Chr(13)
            ValidateTotalFields_115AD = False
'            Exit Function
        End If
    End If


    Total_Balance_115AD = Sheet54.Range("Total_Balance_115AD").value    'added on 11/03/2025
    
    'commented by Chetan C M on 20/01/2026 for AY2026-27
        'start--
'    'added by Chetan C M on 03/03/2025
'    'start--
'
'    Total_Balance_115ADBE = Sheet54.Range("Total_Balance_115ADBE").value
'    Total_Balance_115ADAE = Sheet54.Range("Total_Balance_115ADAE").value
'    TotBalance_115_ABE_AAE = Sheet54.Range("TotBalance_115_ABE_AAE").value  'added by Chetan C M on 29/10/2025 for SIT-89185
'    '--end
'
'    'added by Chetan C M on 29/10/2025
'    'start--
'
'    If Len(Total_Balance_115ADBE) > 14 Then
'            MsgBox_115AD = MsgBox_115AD + "* Total of Col 14 where transfer was before 23rd July 2024 cannot be more than 14 digits in sheet 115AD(1)(iii) proviso" & Chr(13)
'            ValidateTotalFields_115AD = False
'
'    End If
'
'    If Len(Total_Balance_115ADAE) > 14 Then
'            MsgBox_115AD = MsgBox_115AD + "* Total of Col 14 where transfer was on or after 23rd July 2024 cannot be more than 14 digits in sheet 115AD(1)(iii) proviso" & Chr(13)
'            ValidateTotalFields_115AD = False
'
'    End If
'
'    If Len(TotBalance_115_ABE_AAE) > 14 Then
'            MsgBox_115AD = MsgBox_115AD + "* Total of LTCG u/s 112A rws 115AD(1)(b)(iii) proviso cannot be more than 14 digits in sheet 115AD(1)(iii) proviso" & Chr(13)
'            ValidateTotalFields_115AD = False
'
'    End If
'    '--end
        '--end
        
End Function

Function Validate115AD_1() As Boolean
    Validate115AD_1 = True
    
    setTableInfo115AD
    setTableInfo115AD1
    setTableInfo115AD2
    setTableInfo115AD3
    setTableInfo115AD4
    setTableInfo115AD5
'    setTableInfo115AD6
    setTableInfo115AD7
    setTableInfo115AD8
    setTableInfo115AD9
    
'Added by sudha on 25th dec 24
'Commented by Riyaz on 09/02/2026
'    setTableInfo115AD10
    
'-----------------------------------
    
    MsgBox_115AD = "Sheet 115AD(1)(iii) proviso : " & Chr(10)
    
  'Changed by sudha on 25th Dec 24
  
    'end115AD = WorksheetFunction.Max(end115AD, end115AD1, end115AD2, end115AD3, end115AD4, end115AD5, end115AD7, end115AD8, end115AD9)
    
    end115AD = WorksheetFunction.Max(end115AD, end115AD1, end115AD2, end115AD3, end115AD4, end115AD5, end115AD7, end115AD8, end115AD9, end115AD10)
 '-----------------------------
    
    If Not ValidateShareAcq_115AD_1 Then Validate115AD_1 = False
    
'Added by sudha on 25th dec 24
'Commented by Riyaz on 09/02/2026
'    If Not ValidateShareTra_115AD_1 Then Validate115AD_1 = False
     
'-------------------------------

    If Not ValidateISINCode_115AD Then Validate115AD_1 = False
    If Not ValidateNameofShare_115AD Then Validate115AD_1 = False
    If Not ValidateNoofShare_115AD Then Validate115AD_1 = False
    If Not ValidateSalePriceperShare_115AD Then Validate115AD_1 = False
    If Not ValidateTotalSaleValue_115AD_1T Then Validate115AD_1 = False
    If Not ValidateCOAwithIndex_115AD Then Validate115AD_1 = False
    If Not ValidateFairMarketValue_115AD Then Validate115AD_1 = False
'    If Not ValidateCostofImprovement_115AD Then Validate115AD_1 = False
    If Not ValidateExpenditureWholly_115AD Then Validate115AD_1 = False
    If Not ValidateNonEditableFields_115AD Then Validate115AD_1 = False
    If Not ValidateTotalFields_115AD Then Validate115AD_1 = False
    
End Function


Function ValidateLock_115AD_1()
    Dim rangecells As Range
    Dim targetadd As Variant
    Set rangecells = Sheet54.Range("TotalSaleValue_115AD").Cells
    Dim i As Long
    
    Sheet54.Unprotect Password:=getmsgstate
    Application.EnableEvents = False
    
    For i = 1 To rangecells.Rows.count
        
        targetadd = rangecells.item(i).Address

        If Sheet54.Range(Replace(targetadd, "J", "I")).Locked = True And (Not rangecells.item(i).value = 0) Then
        
            rangecells.item(i).Locked = True
            rangecells.item(i).Interior.Color = vbWhite
            rangecells.item(i).Font.Color = RGB(51, 102, 255)
            rangecells.item(i).value = 0
            
        End If
        
    Next
    
    Sheet54.Protect Password:=getmsgstate
    Application.EnableEvents = True
    
End Function

'commented by Chetan C M on 20/01/2026 for AY2026-27
        'start--
'Function ValidateShareTra_115AD_1() As Boolean
'    ValidateShareTra_115AD_1 = True
'    Dim rangecells As Range
'    Set rangecells = Sheet54.Range("ShareTrans_115AD_1").Cells
'    Dim i As Long
'    ReDim ShareTrans_115AD_1(end115AD)
'
'    For i = 1 To end115AD
'        ShareTrans_115AD_1(i) = rangecells.item(i).value
'        'SIT-89058 Malli
'        If ShareTrans_115AD_1(i) = "(Select)" Or ShareTrans_115AD_1(i) = "" Then
'            'MsgBox_115AD = MsgBox_115AD + "* Please select the dropdown option in Column 1b at Sr.No." & i & " in sheet 115AD(1)(iii) proviso" & Chr(13)
'              'MsgBox_115AD = MsgBox_115AD + "* ""Please select the dropdown option in Column 1b of Schedule 115AD(1)(iii) proviso at Sr.No." & i & """" & Chr(13)
'             'Added by Shrutika(04/03/2025)
'              MsgBox_115AD = MsgBox_115AD + "* Please select the dropdown option in Column 1b of Schedule 115AD(1)(iii) proviso at Sr.No." & i & "" & Chr(13)
'            ValidateShareTra_115AD_1 = False
'            Exit Function
'        End If
'    Next
'End Function
'Sub setTableInfo115AD10()
'    Dim rangecells As Range
'    Dim mIntCells  As Long
'    Dim mIntCtr  As Long
'    Dim ccount  As Long
'    ccount = 0
'    Set rangecells = Sheet54.Range("ShareTrans_115AD_1").Cells
'    mIntCells = Sheet54.Range("ShareTrans_115AD_1").count
'    For mIntCtr = 1 To mIntCells
'        If Not (rangecells.item(mIntCtr).value = "(Select)" Or rangecells.item(mIntCtr).value = "") And rangecells.item(mIntCtr).Locked = False Then
'        ccount = ccount + 1
'        End If
'    Next
'    end115AD10 = ccount
'End Sub
        '--end

Function validateImport115AD_1()
    Dim targetadd As Variant
    Dim rangecells As Range
    Set rangecells = Sheet54.Range("ShareAcq_115AD_1").Cells
    Dim i As Long
    Sheet54.Unprotect Password:=getmsgstate
    Application.EnableEvents = False
    
    For i = 1 To rangecells.Rows.count
        targetadd = rangecells.item(i).Address
        targetadd = Replace(targetadd, "$", "")
        If rangecells.item(i).value = "After 31st January 2018" And Sheet54.Range(Replace(targetadd, "E", "F")).Locked = False Then

            Sheet54.Range(Replace(targetadd, "E", "F")).Locked = True
            Sheet54.Range(Replace(targetadd, "E", "F")).Interior.Color = "&HD8D8D8"
            Sheet54.Range(Replace(targetadd, "E", "F")).value = "INNOTREQUIRD"
            
            Sheet54.Range(Replace(targetadd, "E", "G")).Locked = True
            Sheet54.Range(Replace(targetadd, "E", "G")).Interior.Color = "&HD8D8D8"
            Sheet54.Range(Replace(targetadd, "E", "G")).value = "CONSOLIDATED"
            
            Sheet54.Range(Replace(targetadd, "E", "H")).Locked = True
            Sheet54.Range(Replace(targetadd, "E", "H")).Interior.Color = "&HD8D8D8"
            Sheet54.Range(Replace(targetadd, "E", "H")).value = 0
            
            Sheet54.Range(Replace(targetadd, "E", "I")).Locked = True
            Sheet54.Range(Replace(targetadd, "E", "I")).Interior.Color = "&HD8D8D8"
            Sheet54.Range(Replace(targetadd, "E", "I")).value = 0
            
            Sheet54.Range(Replace(targetadd, "E", "N")).Locked = True
            Sheet54.Range(Replace(targetadd, "E", "N")).Interior.Color = "&HD8D8D8"
            Sheet54.Range(Replace(targetadd, "E", "N")).value = 0
            
            Sheet54.Range(Replace(targetadd, "E", "J")).Locked = False
            Sheet54.Range(Replace(targetadd, "E", "J")).Interior.Color = "&HCCFFCC"
            Sheet54.Range(Replace(targetadd, "E", "J")).Font.Color = vbBlack
            Sheet54.Range(Replace(targetadd, "E", "J")).value = Application.WorksheetFunction.Round(Sheet54.Range(Replace(targetadd, "E", "J")).value, 0)
        Else
            Sheet54.Range(Replace(targetadd, "E", "J")).Locked = True
            Sheet54.Range(Replace(targetadd, "E", "J")).Interior.Color = vbWhite
            Sheet54.Range(Replace(targetadd, "E", "J")).Font.Color = RGB(51, 102, 255)
            'MsgBox ("=MAX(0," & "ROUND(" & Replace(TargetAdd, "E", "H") & ",4)" & "*" & "ROUND(" & Replace(TargetAdd, "E", "I") & ",4)" & ")")
            Sheet54.Range(Replace(targetadd, "E", "J")).formula = "=ROUND(MAX(0," & "ROUND(" & Replace(targetadd, "E", "H") & ",4)" & "*" & "ROUND(" & Replace(targetadd, "E", "I") & ",4)" & "),0)"
        End If
    Next
    'Non editable fields
    Sheet54.Range("TotalSaleValue_115AD").NumberFormat = "0"
    Sheet54.Range("COAwithoutIndex_115AD").NumberFormat = "0"
    Sheet54.Range("LTCGAssetAcquired_115AD").NumberFormat = "0"
    Sheet54.Range("TotalFairMarketValue_115AD").NumberFormat = "0"
    Sheet54.Range("TotalDeductions_115AD").NumberFormat = "0"
    Sheet54.Range("Balance_115AD").NumberFormat = "0"
    'Total fields
    Sheet54.Range("Total_TotalSaleValue_115AD").NumberFormat = "0"
    Sheet54.Range("Total_COAwithoutIndex_115AD").NumberFormat = "0"
    Sheet54.Range("Total_COAwithIndex_115AD").NumberFormat = "0"
    Sheet54.Range("Total_LTCGAssetAcquired_115AD").NumberFormat = "0"
    Sheet54.Range("Total_TotalFairMarketValue_115AD").NumberFormat = "0"
    Sheet54.Range("Total_ExpenditureWholly_115AD").NumberFormat = "0"
    Sheet54.Range("Total_TotalDeductions_115AD").NumberFormat = "0"
    Sheet54.Range("Total_Balance_115AD").NumberFormat = "0"
    
    Sheet54.Protect Password:=getmsgstate
    Application.EnableEvents = True
End Function



