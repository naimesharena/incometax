Attribute VB_Name = "md112A"
Option Explicit

Public MsgBox_112A As String

Public TotalSaleValue_112A_1T As Variant
Public ShareAcq_112A As Variant

'Public ShareTra_112A As Variant 'Added by Aavula 2025-26    'commented by Chetan C M on 30/01/2026 for AY 2026-27

Public ISINCode_112A As Variant
Public NameofShare_112A As Variant
Public NoofShare_112A As Variant
Public SalePriceperShare_112A As Variant
Public TotalSaleValue_112A As Variant
Public COAwithoutIndex_112A As Variant
Public COAwithIndex_112A As Variant
Public LTCGAssetAcquired_112A As Variant
Public FairMarketValue_112A As Variant
Public TotalFairMarketValue_112A As Variant
'Public FullValueConsideration_112A As Variant
'Public CostofImprovement_112A As Variant
Public ExpenditureWholly_112A As Variant
Public TotalDeductions_112A As Variant
Public Balance_112A As Variant

Public rngname_112A As Variant
Public Total_TotalSaleValue_112A As Variant
Public Total_COAwithoutIndex_112A As Variant
Public Total_COAwithIndex_112A As Variant
Public Total_LTCGAssetAcquired_112A As Variant
Public Total_TotalFairMarketValue_112A As Variant
'Public Total_FullValueConsideration_112A As Variant
'Public Total_CostofImprovement_112A As Variant
Public Total_ExpenditureWholly_112A As Variant
Public Total_TotalDeductions_112A As Variant
Public Total_Balance_112A As Variant
Public LockFlag112A, LockFlag115A As Boolean

'commented by Chetan C M on 30/01/2026 for AY 2026-27
'start--
'Public Total_Balance_112ABE As Variant
'Public Total_Balance_112AAE As Variant
'--end

Public end112A, end112A1, end112A2, end112A3, end112A4, end112A5, end112A6, end112A7, end112A8, end112A9 As Long

Sub ValidateSheet112A_Click()
Validate112A
fmsgboxStatus " Schedule 112A is ok" ', vbOKOnly, "ITR-3"
End Sub

Sub Validate112A()
    If Not Validate112A_1 Then
        Sheet53.Activate
        fmsgboxsmall MsgBox_112A ', vbOKOnly, "Error(s)"
        CloseMsg
    End If
End Sub
Sub AddRows112A()
Dim vRows  As Long
Sheets("Schedule 112A").Activate

'changed by Chetan C M on 30/01/2026 for AY 2026-27
'start--
'EfilingCommon.DefinedgridNameRange = "ShareAcq_112A||ShareTra_112A||ISINCode_112A||NameofShare_112A||NoofShare_112A||SalePriceperShare_112A||TotalSaleValue_112A||COAwithoutIndex_112A||COAwithIndex_112A||LTCGAssetAcquired_112A||FairMarketValue_112A||TotalFairMarketValue_112A||ExpenditureWholly_112A||TotalDeductions_112A||Balance_112A"
EfilingCommon.DefinedgridNameRange = "ShareAcq_112A||ISINCode_112A||NameofShare_112A||NoofShare_112A||SalePriceperShare_112A||TotalSaleValue_112A||COAwithoutIndex_112A||COAwithIndex_112A||LTCGAssetAcquired_112A||FairMarketValue_112A||TotalFairMarketValue_112A||ExpenditureWholly_112A||TotalDeductions_112A||Balance_112A"
'--end

ActiveCellRange = EfilingCommon.searchLastRow("ISINCode_112A")
EfilingCommon.insertRowUnderSectionWithFormula f_112A_115AD_1:=1
End Sub

Sub setTableInfo112A()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet53.Range("ISINCode_112A").Cells
    mIntCells = Sheet53.Range("ISINCode_112A").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end112A = ccount
    
    'changed by Chetan C M on 30/01/2026 for AY 2026-27
    'start--
    'rngname_112A = "ShareAcq_112A;ShareTra_112A;ISINCode_112A;NameofShare_112A;NoofShare_112A;SalePriceperShare_112A;TotalSaleValue_112A;COAwithoutIndex_112A;COAwithIndex_112A;LTCGAssetAcquired_112A;FairMarketValue_112A;TotalFairMarketValue_112A;ExpenditureWholly_112A;TotalDeductions_112A;Balance_112A;"
    rngname_112A = "ShareAcq_112A;ISINCode_112A;NameofShare_112A;NoofShare_112A;SalePriceperShare_112A;TotalSaleValue_112A;COAwithoutIndex_112A;COAwithIndex_112A;LTCGAssetAcquired_112A;FairMarketValue_112A;TotalFairMarketValue_112A;ExpenditureWholly_112A;TotalDeductions_112A;Balance_112A;"
    '--end
    
End Sub

Sub setTableInfo112A8()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet53.Range("ShareAcq_112A").Cells
    mIntCells = Sheet53.Range("ShareAcq_112A").count
    For mIntCtr = 1 To mIntCells
        If (Not rangecells.item(mIntCtr).value = "(Select)") And (Not rangecells.item(mIntCtr).value = "") Then
        ccount = ccount + 1
        End If
    Next
    end112A8 = ccount
End Sub

'commented by Chetan C M on 30/01/2026 for AY 2026-27
'start--
'Sub setTableInfo112B()
'    Dim rangecells As Range
'    Dim mIntCells  As Long
'    Dim mIntCtr  As Long
'    Dim ccount  As Long
'    ccount = 0
'    Set rangecells = Sheet53.Range("ShareTra_112A").Cells
'    mIntCells = Sheet53.Range("ShareTra_112A").count
'    For mIntCtr = 1 To mIntCells
'        If (Not rangecells.item(mIntCtr).value = "(Select)") And (Not rangecells.item(mIntCtr).value = "") Then
'        ccount = ccount + 1
'        End If
'    Next
'    end112A8 = ccount
'End Sub
'--end


Sub setTableInfo112A9()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet53.Range("TotalSaleValue_112A").Cells
    mIntCells = Sheet53.Range("TotalSaleValue_112A").count
    For mIntCtr = 1 To mIntCells
        If rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end112A9 = ccount
End Sub
Function ValidateTotalSaleValue_112A_1T() As Boolean
    ValidateTotalSaleValue_112A_1T = True
    Dim rangecells As Range
    Set rangecells = Sheet53.Range("TotalSaleValue_112A").Cells
    Dim i As Long
    ReDim TotalSaleValue_112A_1T(end112A)
    
    For i = 1 To end112A
        TotalSaleValue_112A_1T(i) = rangecells.item(i).value
        
        If TotalSaleValue_112A_1T(i) = "" Then
            MsgBox_112A = MsgBox_112A + "* Please enter Total Sale Value at Sr. No." & i & " in sheet 112A" & Chr(13)
            ValidateTotalSaleValue_112A_1T = False
            Exit Function
        End If
    Next
End Function
Function ValidateShareAcq_112A() As Boolean
    ValidateShareAcq_112A = True
    Dim rangecells As Range
    Set rangecells = Sheet53.Range("ShareAcq_112A").Cells
    Dim i As Long
    ReDim ShareAcq_112A(end112A)
    
    For i = 1 To end112A
        ShareAcq_112A(i) = rangecells.item(i).value
        
        If ShareAcq_112A(i) = "(Select)" Then
           'changed by Chetan C M on 10/09/2025
        'start--
            'MsgBox_112A = MsgBox_112A + "* Please select Whether Share acquired on or before / After 31st January,2018 at Sr. No." & i & " in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Please select the dropdown option in Column 1a of Schedule 112A." & Chr(13)
        '--end
        ValidateShareAcq_112A = False
            Exit Function
        End If
    Next
End Function

'commented by Chetan C M on 30/01/2026 for AY 2026-27
'start--
'Function ValidateShareTra_112A() As Boolean
'    ValidateShareTra_112A = True
'    Dim rangecells As Range
'    Set rangecells = Sheet53.Range("ShareTra_112A").Cells
'    Dim i As Long
'    ReDim ShareTra_112A(end112A)
'
'    For i = 1 To end112A
'        ShareTra_112A(i) = rangecells.item(i).value
'
'        If ShareTra_112A(i) = "(Select)" Or ShareTra_112A(i) = "" Then
'
'            MsgBox_112A = MsgBox_112A + "* Please select the dropdown option in Column 1b of Schedule 112A" & Chr(13)
'            ValidateShareTra_112A = False
'            Exit Function
'        End If
'    Next
'End Function
'--end

Sub setTableInfo112A1()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet53.Range("NameofShare_112A").Cells
    mIntCells = Sheet53.Range("NameofShare_112A").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end112A1 = ccount
End Sub

Sub setTableInfo112A2()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet53.Range("NoofShare_112A").Cells
    mIntCells = Sheet53.Range("NoofShare_112A").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end112A2 = ccount
End Sub

Sub setTableInfo112A3()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet53.Range("SalePriceperShare_112A").Cells
    mIntCells = Sheet53.Range("SalePriceperShare_112A").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end112A3 = ccount
End Sub

Sub setTableInfo112A4()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet53.Range("COAwithIndex_112A").Cells
    mIntCells = Sheet53.Range("COAwithIndex_112A").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end112A4 = ccount
End Sub

Sub setTableInfo112A5()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet53.Range("FairMarketValue_112A").Cells
    mIntCells = Sheet53.Range("FairMarketValue_112A").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end112A5 = ccount
End Sub

Sub setTableInfo112A7()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet53.Range("ExpenditureWholly_112A").Cells
    mIntCells = Sheet53.Range("ExpenditureWholly_112A").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
        ccount = ccount + 1
        End If
    Next
    end112A7 = ccount
End Sub

Function ValidateISINCode_112A() As Boolean
    ValidateISINCode_112A = True
    Dim rangecells As Range
    Set rangecells = Sheet53.Range("ISINCode_112A").Cells
    Dim i As Long
    ReDim ISINCode_112A(end112A)
    
    subProcCaption = "Validating 112A"
    noOfProcessSub = end112A
    
    For i = 1 To end112A
        ISINCode_112A(i) = rangecells.item(i).value
        
        If Not chkMandatory(ISINCode_112A(i)) Then
        'changed by Chetan C M on 10/09/2025
        'start--
            'MsgBox_112A = MsgBox_112A + "* Please enter the ISIN Code. at Sr. No." & i & " in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Please enter the ISIN Code. at Sr. No." & i & " in Schedule 112A" & Chr(13)
        '--end
            ValidateISINCode_112A = False
            Exit Function
        End If
        
        If Len(ISINCode_112A(i)) > 0 Then
            If Len(ISINCode_112A(i)) <> 12 Then
            'changed by Chetan C M on 10/09/2025
            'start--
                'MsgBox_112A = MsgBox_112A + "* Please enter the valid ISIN Code. at Sr. No." & i & " in sheet 112A" & Chr(13)
                MsgBox_112A = MsgBox_112A + "* Please enter the valid ISIN Code. at Sr. No." & i & " in Schedule 112A" & Chr(13)
            '--end
                ValidateISINCode_112A = False
                Exit Function
            End If
        End If
        
        If UCase(Trim(Mid(ISINCode_112A(i), 1, 2))) <> "IN" Then
        'changed by Chetan C M on 10/09/2025
        'start--
            'MsgBox_112A = MsgBox_112A + "* Please enter the valid ISIN Code. at Sr. No." & i & " in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Please enter the valid ISIN Code. at Sr. No." & i & " in Schedule 112A" & Chr(13)
        '--end
            ValidateISINCode_112A = False
            Exit Function
        End If
    Next
End Function

Function ValidateNameofShare_112A() As Boolean
    ValidateNameofShare_112A = True
    Dim rangecells As Range
    Set rangecells = Sheet53.Range("NameofShare_112A").Cells
    Dim i As Long
    ReDim NameofShare_112A(end112A)
    
    For i = 1 To end112A
        NameofShare_112A(i) = rangecells.item(i).value
        
        If Not chkMandatory(NameofShare_112A(i)) Then
        'changed by Chetan C M on 10/09/2025
        'start--
            'MsgBox_112A = MsgBox_112A + "* Please enter Name of the Share/Unit at Sr. No." & i & " in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Please enter Name of the Share/Unit at Sr. No." & i & " in Schedule 112A" & Chr(13)
        '--end
            ValidateNameofShare_112A = False
            Exit Function
        End If
        
        If Len(NameofShare_112A(i)) > 125 Then
        'changed by Chetan C M on 10/09/2025
        'start--
            'MsgBox_112A = MsgBox_112A + "* Name of the Share/Unit at Sr. No." & i & " should not exceed 125 characters.in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Name of the Share/Unit at Sr. No." & i & " should not exceed 125 characters.in Schedule 112A" & Chr(13)
        '--end
            ValidateNameofShare_112A = False
            Exit Function
        End If
    Next
End Function

Function ValidateNoofShare_112A() As Boolean
    ValidateNoofShare_112A = True
    Dim rangecells As Range
    Set rangecells = Sheet53.Range("NoofShare_112A").Cells
    Dim i As Long
    ReDim NoofShare_112A(end112A)
    
    For i = 1 To end112A
        NoofShare_112A(i) = rangecells.item(i).value
        
        If Not chkMandatory(NoofShare_112A(i)) Then
        'changed by Chetan C M on 10/09/2025
        'start--
            'MsgBox_112A = MsgBox_112A + "* Please enter the No. of Shares/Units at Sr. No." & i & " in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Please enter the No. of Shares/Units at Sr. No." & i & " in Schedule 112A" & Chr(13)
        '--end
            ValidateNoofShare_112A = False
            Exit Function
        End If

    Next
End Function

Function ValidateSalePriceperShare_112A() As Boolean
    ValidateSalePriceperShare_112A = True
    Dim rangecells As Range
    Set rangecells = Sheet53.Range("SalePriceperShare_112A").Cells
    Dim i As Long
    ReDim SalePriceperShare_112A(end112A)
    
    For i = 1 To end112A
        SalePriceperShare_112A(i) = rangecells.item(i).value
        
        If Not chkMandatory(SalePriceperShare_112A(i)) Then
        ''changed by Chetan C M on 11/09/2025
        'start--
            'MsgBox_112A = MsgBox_112A + "* Please enter the Sale-price per Share/Unit at Sr. No." & i & " in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Please enter the Sale-price per Share/Unit in Column 5 of Schedule 112A" & Chr(13)
        '--end
            ValidateSalePriceperShare_112A = False
            Exit Function
        End If

    Next
End Function

Function ValidateCOAwithIndex_112A() As Boolean
    ValidateCOAwithIndex_112A = True
    Dim rangecells As Range
    Set rangecells = Sheet53.Range("COAwithIndex_112A").Cells
    Dim i As Long
    ReDim COAwithIndex_112A(end112A)
    
    For i = 1 To end112A
        COAwithIndex_112A(i) = rangecells.item(i).value
        
        If Not chkMandatory(COAwithIndex_112A(i)) Then
        'changed by Chetan C M on 06/10/2025 for SIT-91306(TRC-627747)
        'start--
            'MsgBox_112A = MsgBox_112A + "* Please enter the Cost of acquisition at Sr. No." & i & " in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Please enter the Cost of acquisition in Column 8 of Schedule 112A" & Chr(13)
        '--end
            ValidateCOAwithIndex_112A = False
            Exit Function
        End If

    Next
End Function

Function ValidateFairMarketValue_112A() As Boolean
    ValidateFairMarketValue_112A = True
    Dim rangecells As Range
    Set rangecells = Sheet53.Range("FairMarketValue_112A").Cells
    Dim i As Long
    ReDim FairMarketValue_112A(end112A)
    
    For i = 1 To end112A
        FairMarketValue_112A(i) = rangecells.item(i).value
        
        If Not chkMandatory(FairMarketValue_112A(i)) Then
        'changed by Chetan C M on 11/09/2025
        'start--
            'MsgBox_112A = MsgBox_112A + "* Please enter the Fair Market Value per share/unit as on 31st January,2018. If share is Aquired after 31st Jan 2018 , please enter ""0"" at Sr. No." & i & " in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Please enter the Fair Market Value per share/unit as on 31st January,2018 in Column 10 of Schedule 112A. If share is Acquired after 31st Jan 2018 , please enter ""0""" & Chr(13)
        '--end
            ValidateFairMarketValue_112A = False
            Exit Function
        End If

    Next
End Function


Function ValidateExpenditureWholly_112A() As Boolean
    ValidateExpenditureWholly_112A = True
    Dim rangecells As Range
    Set rangecells = Sheet53.Range("ExpenditureWholly_112A").Cells
    Dim i As Long
    ReDim ExpenditureWholly_112A(end112A)
    
    For i = 1 To end112A
        ExpenditureWholly_112A(i) = rangecells.item(i).value
        
        If Not chkMandatory(ExpenditureWholly_112A(i)) Then
        'changed by Chetan C M on 11/09/2025
        'start--
            'MsgBox_112A = MsgBox_112A + "* Please enter the Expenditure incurred wholly and exclusively in connection with transfer at Sr. No." & i & " in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Please enter the Expenditure wholly and exclusively in connection with transfer in Column 14 of Schedule 112A" & Chr(13)
        '--end
            ValidateExpenditureWholly_112A = False
            Exit Function
        End If

    Next
End Function

Function ValidateNonEditableFields_112A() As Boolean
    ValidateNonEditableFields_112A = True
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Dim rangecells4 As Range
    Dim rangecells5 As Range
    Dim rangecells6 As Range
    Dim tmpArray() As String
    
    Set rangecells = Sheet53.Range("TotalSaleValue_112A").Cells
    Set rangecells1 = Sheet53.Range("COAwithoutIndex_112A").Cells
    Set rangecells2 = Sheet53.Range("LTCGAssetAcquired_112A").Cells
    Set rangecells3 = Sheet53.Range("TotalFairMarketValue_112A").Cells
    'Set rangecells4 = Sheet53.Range("FullValueConsideration_112A").Cells
    Set rangecells5 = Sheet53.Range("TotalDeductions_112A").Cells
    Set rangecells6 = Sheet53.Range("Balance_112A").Cells
    
    Dim i As Long
    ReDim TotalSaleValue_112A(end112A)
    ReDim COAwithoutIndex_112A(end112A)
    ReDim LTCGAssetAcquired_112A(end112A)
    ReDim TotalFairMarketValue_112A(end112A)
    'ReDim FullValueConsideration_112A(end112A)
    ReDim TotalDeductions_112A(end112A)
    ReDim Balance_112A(end112A)
    
    For i = 1 To end112A
        TotalSaleValue_112A(i) = rangecells.item(i).value
        COAwithoutIndex_112A(i) = rangecells1.item(i).value
        LTCGAssetAcquired_112A(i) = rangecells2.item(i).value
        TotalFairMarketValue_112A(i) = rangecells3.item(i).value
        'FullValueConsideration_112A(i) = rangecells4.item(i).value
        TotalDeductions_112A(i) = rangecells5.item(i).value
        Balance_112A(i) = rangecells6.item(i).value
        
        If Len(TotalSaleValue_112A(i)) > 0 Then
            TotalSaleValue_112A(i) = Application.WorksheetFunction.Round(TotalSaleValue_112A(i), 0)
        End If
        
'        Tmparray = Split(TotalSaleValue_112A(i), ".")
        If Len(TotalSaleValue_112A(i)) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
        'changed by Chetan C M on 10/09/2025
        'start--
            'MsgBox_112A = MsgBox_112A + "* Total Sale Value at Sr. No " & i & " cannot exceed 14 digits in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Full value of consideration at Sr. No " & i & " cannot exceed 14 digits in Schedule 112A" & Chr(13)
        '--end
        ValidateNonEditableFields_112A = False
        'commented by Chetan C M on 13/10/2025 for SIT-SIT-95268 (issue 5)
'            Exit Function
        End If
        
        If Len(COAwithoutIndex_112A(i)) > 0 Then
            COAwithoutIndex_112A(i) = Application.WorksheetFunction.Round(COAwithoutIndex_112A(i), 0)
        End If
        
'        Tmparray = Split(COAwithoutIndex_112A(i), ".")
        If Len(COAwithoutIndex_112A(i)) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
        'changed by Chetan C M on 10/09/2025
        'start--
            'MsgBox_112A = MsgBox_112A + "* Cost of acquisition without indexation at Sr. No " & i & " cannot exceed 14 digits in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Cost of acquisition without indexation at Sr. No " & i & " cannot exceed 14 digits in Schedule 112A" & Chr(13)
        '--end
            ValidateNonEditableFields_112A = False
            'commented by Chetan C M on 13/10/2025 for SIT-SIT-95268 (issue 5)
'            Exit Function
        End If
        
        If Len(LTCGAssetAcquired_112A(i)) > 0 Then
            LTCGAssetAcquired_112A(i) = Application.WorksheetFunction.Round(LTCGAssetAcquired_112A(i), 0)
        End If

'        Tmparray = Split(LTCGAssetAcquired_112A(i), ".")
        If Len(LTCGAssetAcquired_112A(i)) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
        'changed by Chetan C M on 10/09/2025
        'start--
            'MsgBox_112A = MsgBox_112A + "* Long term capital asset acquired before 01.02.2018 at Sr. No " & i & " cannot exceed 14 digits in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Long term capital asset acquired before 01.02.2018 at Sr. No " & i & " cannot exceed 14 digits in Schedule 112A" & Chr(13)
        '--end
            ValidateNonEditableFields_112A = False
            'commented by Chetan C M on 13/10/2025 for SIT-SIT-95268 (issue 5)
'            Exit Function
        End If
        
        If Len(TotalFairMarketValue_112A(i)) > 0 Then
        TotalFairMarketValue_112A(i) = Application.WorksheetFunction.Round(TotalFairMarketValue_112A(i), 0)
        End If

'        Tmparray = Split(TotalFairMarketValue_112A(i), ".")
        If Len(TotalFairMarketValue_112A(i)) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
        'changed by Chetan C M on 10/09/2025
        'start--
            'MsgBox_112A = MsgBox_112A + "* Total Fair Market Value of capital asset at Sr. No " & i & " cannot exceed 14 digits in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Total Fair Market Value of capital asset at Sr. No " & i & " cannot exceed 14 digits in Schedule 112A" & Chr(13)
        '--end
            ValidateNonEditableFields_112A = False
            'commented by Chetan C M on 13/10/2025 for SIT-SIT-95268 (issue 5)
'            Exit Function
        End If

        If Len(TotalDeductions_112A(i)) > 0 Then
            TotalDeductions_112A(i) = Application.WorksheetFunction.Round(TotalDeductions_112A(i), 0)
        End If

'        Tmparray = Split(TotalDeductions_112A(i), ".")
        If Len(TotalDeductions_112A(i)) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
        'changed by Chetan C M on 10/09/2025
        'start--
            'MsgBox_112A = MsgBox_112A + "* Total deductions at Sr. No " & i & " cannot exceed 14 digits in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Total deductions at Sr. No " & i & " cannot exceed 14 digits in Schedule 112A" & Chr(13)
        '--end
            ValidateNonEditableFields_112A = False
            'commented by Chetan C M on 13/10/2025 for SIT-SIT-95268 (issue 5)
'            Exit Function
        End If
        
        If Len(Balance_112A(i)) > 0 Then
            Balance_112A(i) = Application.WorksheetFunction.Round(Balance_112A(i), 0)
        End If
        
'        Tmparray = Split(Balance_112A(i), ".")
        If Not Balance_112A(i) < 0 Then
            If Len(Balance_112A(i)) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
                MsgBox_112A = MsgBox_112A + "Balance at Sr. No " & i & " cannot exceed 14 digits " & " in Schedule 112A" & Chr(13)
                ValidateNonEditableFields_112A = False
                'commented by Chetan C M on 13/10/2025 for SIT-SIT-95268 (issue 5)
'                Exit Function
            End If
        Else
'            Tmparray = Split(CStr((Balance_112A(i) * -1)), ".")
            If Len((Balance_112A(i) * -1)) > 14 Then ' Or Len(Tmparray(0)) > 11 Then
                MsgBox_112A = MsgBox_112A + "Balance at Sr. No " & i & " cannot exceed 14 digits " & " in Schedule 112A" & Chr(13)
                ValidateNonEditableFields_112A = False
                'commented by Chetan C M on 13/10/2025 for SIT-SIT-95268 (issue 5)
'                Exit Function
            End If
        End If
    Next

End Function

Function ValidateTotalFields_112A() As Boolean
 ValidateTotalFields_112A = True
 Dim tmpArray() As String
 
 Total_TotalSaleValue_112A = Sheet53.Range("Total_TotalSaleValue_112A").value
 Total_COAwithoutIndex_112A = Sheet53.Range("Total_COAwithoutIndex_112A").value
 Total_COAwithIndex_112A = Sheet53.Range("Total_COAwithIndex_112A").value
 Total_LTCGAssetAcquired_112A = Sheet53.Range("Total_LTCGAssetAcquired_112A").value
 Total_TotalFairMarketValue_112A = Sheet53.Range("Total_TotalFairMarketValue_112A").value
' Total_FullValueConsideration_112A = Sheet53.Range("Total_FullValueConsideration_112A").value
' Total_CostofImprovement_112A = Sheet53.Range("Total_CostofImprovement_112A").value
 Total_ExpenditureWholly_112A = Sheet53.Range("Total_ExpenditureWholly_112A").value
 Total_TotalDeductions_112A = Sheet53.Range("Total_TotalDeductions_112A").value
 Total_Balance_112A = Sheet53.Range("Total_Balance_112A").value
 
'commented by Chetan C M on 30/01/2026 for AY 2026-27
'start--
' 'Added by Aavula 2025-26
' Total_Balance_112ABE = Sheet53.Range("Total_Balance_112ABE").value
' Total_Balance_112AAE = Sheet53.Range("Total_Balance_112AAE").value
 '''''''''
'--end
 
    If Len(Total_TotalSaleValue_112A) > 0 Then
        Total_TotalSaleValue_112A = Application.WorksheetFunction.Round(Total_TotalSaleValue_112A, 0)
    End If
 
    'Tmparray = Split(Total_TotalSaleValue_112A, ".")
    If Len(Total_TotalSaleValue_112A) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
    'changed by Chetan C M on 10/09/2025
    'start--
       'MsgBox_112A = MsgBox_112A + "* Total_TotalSaleValue_112A cannot be greater than 14 digits in sheet 112A" & Chr(13)
       MsgBox_112A = MsgBox_112A + "* Total Full value of consideration cannot be greater than 14 digits in sheet 112A" & Chr(13)
    '--end
       ValidateTotalFields_112A = False
       'commented by Chetan C M on 13/10/2025 for SIT-SIT-95268 (issue 5)
'       Exit Function
    End If
 
    If Len(Total_COAwithoutIndex_112A) > 0 Then
        Total_COAwithoutIndex_112A = Application.WorksheetFunction.Round(Total_COAwithoutIndex_112A, 0)
    End If

    'Tmparray = Split(Total_COAwithoutIndex_112A, ".")
    If Len(Total_COAwithoutIndex_112A) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
    'changed by Chetan C M on 10/09/2025
    'start--
       'MsgBox_112A = MsgBox_112A + "* Total_COAwithoutIndex_112A cannot be greater than 14 digits in sheet 112A" & Chr(13)
       MsgBox_112A = MsgBox_112A + "* Total_COAwithoutIndex_112A cannot be greater than 14 digits in Schedule 112A" & Chr(13)
    '--end
       ValidateTotalFields_112A = False
       'commented by Chetan C M on 13/10/2025 for SIT-SIT-95268 (issue 5)
'       Exit Function
    End If
 
    If Len(Total_COAwithIndex_112A) > 0 Then
        Total_COAwithIndex_112A = Application.WorksheetFunction.Round(Total_COAwithIndex_112A, 0)
    End If

    'Tmparray = Split(Total_COAwithIndex_112A, ".")
    If Len(Total_COAwithIndex_112A) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
    'changed by Chetan C M on 10/09/2025
    'start--
       MsgBox_112A = MsgBox_112A + "* Total_COAwithIndex_112A cannot be greater than 14 digits in sheet 112A" & Chr(13)
       MsgBox_112A = MsgBox_112A + "* Total_COAwithIndex_112A cannot be greater than 14 digits in Schedule 112A" & Chr(13)
    '--end
       ValidateTotalFields_112A = False
       'commented by Chetan C M on 13/10/2025 for SIT-SIT-95268 (issue 5)
'       Exit Function
    End If
 
    If Len(Total_LTCGAssetAcquired_112A) > 0 Then
        Total_LTCGAssetAcquired_112A = Application.WorksheetFunction.Round(Total_LTCGAssetAcquired_112A, 0)
    End If

    'Tmparray = Split(Total_LTCGAssetAcquired_112A, ".")
    If Len(Total_LTCGAssetAcquired_112A) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
    'changed by Chetan C M on 10/09/2025
    'start--
       MsgBox_112A = MsgBox_112A + "* Total_LTCGAssetAcquired_112A cannot be greater than 14 digits in sheet 112A" & Chr(13)
       MsgBox_112A = MsgBox_112A + "* Total_LTCGAssetAcquired_112A cannot be greater than 14 digits in Schedule 112A" & Chr(13)
    '--end
       ValidateTotalFields_112A = False
       'commented by Chetan C M on 13/10/2025 for SIT-SIT-95268 (issue 5)
'       Exit Function
    End If
 
    If Len(Total_TotalFairMarketValue_112A) > 0 Then
        Total_TotalFairMarketValue_112A = Application.WorksheetFunction.Round(Total_TotalFairMarketValue_112A, 0)
    End If

    'Tmparray = Split(Total_TotalFairMarketValue_112A, ".")
    If Len(Total_TotalFairMarketValue_112A) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
    'changed by Chetan C M on 10/09/2025
    'start--
       'MsgBox_112A = MsgBox_112A + "* Total_TotalFairMarketValue_112A cannot be greater than 14 digits in sheet 112A" & Chr(13)
       MsgBox_112A = MsgBox_112A + "* Total_TotalFairMarketValue_112A cannot be greater than 14 digits in Schedule 112A" & Chr(13)
    '--end
       ValidateTotalFields_112A = False
       'commented by Chetan C M on 13/10/2025 for SIT-SIT-95268 (issue 5)
'       Exit Function
    End If

    If Len(Total_ExpenditureWholly_112A) > 0 Then
        Total_ExpenditureWholly_112A = Application.WorksheetFunction.Round(Total_ExpenditureWholly_112A, 0)
    End If

    'Tmparray = Split(Total_ExpenditureWholly_112A, ".")
    If Len(Total_ExpenditureWholly_112A) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
    'changed by Chetan C M on 10/09/2025
    'start--
       'MsgBox_112A = MsgBox_112A + "* Total_ExpenditureWholly_112A cannot be greater than 14 digits in sheet 112A" & Chr(13)
       MsgBox_112A = MsgBox_112A + "* Total_ExpenditureWholly_112A cannot be greater than 14 digits in Schedule 112A" & Chr(13)
    '--end
       ValidateTotalFields_112A = False
       'commented by Chetan C M on 13/10/2025 for SIT-SIT-95268 (issue 5)
'       Exit Function
    End If
 
    If Len(Total_TotalDeductions_112A) > 0 Then
        Total_TotalDeductions_112A = Application.WorksheetFunction.Round(Total_TotalDeductions_112A, 0)
    End If

    'Tmparray = Split(Total_TotalDeductions_112A, ".")
    If Len(Total_TotalDeductions_112A) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
    'changed by Chetan C M on 10/09/2025
    'start--
       'MsgBox_112A = MsgBox_112A + "* Total_TotalDeductions_112A cannot be greater than 14 digits in sheet 112A" & Chr(13)
       MsgBox_112A = MsgBox_112A + "* Total_TotalDeductions_112A cannot be greater than 14 digits in Schedule 112A" & Chr(13)
    '--end
       ValidateTotalFields_112A = False
       'commented by Chetan C M on 13/10/2025 for SIT-SIT-95268 (issue 5)
'       Exit Function
    End If
 
    If Len(Total_Balance_112A) > 0 Then
        Total_Balance_112A = Application.WorksheetFunction.Round(Total_Balance_112A, 0)
    End If
    
    'Tmparray = Split(Total_Balance_112A, ".")
    If Not Total_Balance_112A < 0 Then
        If Len(Total_Balance_112A) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
            MsgBox_112A = MsgBox_112A + "* Total_Balance_112A cannot be greater than 14 digits in Schedule 112A" & Chr(13)
            ValidateTotalFields_112A = False
            'commented by Chetan C M on 13/10/2025 for SIT-SIT-95268 (issue 5)
'            Exit Function
        End If
    Else
        'Tmparray = Split(CStr((Total_Balance_112A * -1)), ".")
        If Len((Total_Balance_112A * -1)) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
            MsgBox_112A = MsgBox_112A + "* Total_Balance_112A cannot be greater than 14 digits in Schedule 112A" & Chr(13)
            ValidateTotalFields_112A = False
            'commented by Chetan C M on 13/10/2025 for SIT-SIT-95268 (issue 5)
'            Exit Function
        End If
    End If

End Function

Function Validate112A_1() As Boolean
    Validate112A_1 = True
    
    setTableInfo112A
    setTableInfo112A1
    setTableInfo112A2
    setTableInfo112A3
    setTableInfo112A4
    setTableInfo112A5
'    setTableInfo112A6
    setTableInfo112A7
    setTableInfo112A8
    setTableInfo112A9
    
    MsgBox_112A = "Sheet 112A : " & Chr(10)
    
    end112A = WorksheetFunction.Max(end112A, end112A1, end112A2, end112A3, end112A4, end112A5, end112A7, end112A8, end112A9)
    
    If Not ValidateShareAcq_112A Then Validate112A_1 = False
    
'    If Not ValidateShareTra_112A Then Validate112A_1 = False 'Added by Aavula Naresh   'commented by Chetan C M on 30/01/2026 for AY 2026-27

    If Not ValidateISINCode_112A Then Validate112A_1 = False
    If Not ValidateNameofShare_112A Then Validate112A_1 = False
    If Not ValidateNoofShare_112A Then Validate112A_1 = False
    If Not ValidateSalePriceperShare_112A Then Validate112A_1 = False
    If Not ValidateTotalSaleValue_112A_1T Then Validate112A_1 = False
    If Not ValidateCOAwithIndex_112A Then Validate112A_1 = False
    If Not ValidateFairMarketValue_112A Then Validate112A_1 = False
'    If Not ValidateCostofImprovement_112A Then Validate112A_1 = False
    If Not ValidateExpenditureWholly_112A Then Validate112A_1 = False
    If Not ValidateNonEditableFields_112A Then Validate112A_1 = False
    If Not ValidateTotalFields_112A Then Validate112A_1 = False
    
End Function

Function validateImport112A()
    Dim targetadd As Variant
    Dim rangecells As Range
    Set rangecells = Sheet53.Range("ShareAcq_112A").Cells
    Dim i As Long
    Sheet53.Unprotect Password:=getmsgstate
    Application.EnableEvents = False
    
    For i = 1 To rangecells.Rows.count
        targetadd = rangecells.item(i).Address
        targetadd = Replace(targetadd, "$", "")

        If rangecells.item(i).value = "After 31st January 2018" And Sheet53.Range(Replace(targetadd, "E", "F")).Locked = False Then

            Sheet53.Range(Replace(targetadd, "E", "F")).Locked = True
            Sheet53.Range(Replace(targetadd, "E", "F")).Interior.Color = "&HD8D8D8"
            Sheet53.Range(Replace(targetadd, "E", "F")).value = "INNOTREQUIRD"
            
            Sheet53.Range(Replace(targetadd, "E", "G")).Locked = True
            Sheet53.Range(Replace(targetadd, "E", "G")).Interior.Color = "&HD8D8D8"
            Sheet53.Range(Replace(targetadd, "E", "G")).value = "CONSOLIDATED"
            
            Sheet53.Range(Replace(targetadd, "E", "H")).Locked = True
            Sheet53.Range(Replace(targetadd, "E", "H")).Interior.Color = "&HD8D8D8"
            Sheet53.Range(Replace(targetadd, "E", "H")).value = 0
            
            Sheet53.Range(Replace(targetadd, "E", "I")).Locked = True
            Sheet53.Range(Replace(targetadd, "E", "I")).Interior.Color = "&HD8D8D8"
            Sheet53.Range(Replace(targetadd, "E", "I")).value = 0
            
            Sheet53.Range(Replace(targetadd, "E", "N")).Locked = True
            Sheet53.Range(Replace(targetadd, "E", "N")).Interior.Color = "&HD8D8D8"
            Sheet53.Range(Replace(targetadd, "E", "N")).value = 0
            
            Sheet53.Range(Replace(targetadd, "E", "J")).Locked = False
            Sheet53.Range(Replace(targetadd, "E", "J")).Interior.Color = "&HCCFFCC"
            Sheet53.Range(Replace(targetadd, "E", "J")).Font.Color = vbBlack
            Sheet53.Range(Replace(targetadd, "E", "J")).value = Application.WorksheetFunction.Round(Sheet53.Range(Replace(targetadd, "E", "J")).value, 0)
        Else
            Sheet53.Range(Replace(targetadd, "E", "J")).Locked = True
            Sheet53.Range(Replace(targetadd, "E", "J")).Interior.Color = vbWhite
            Sheet53.Range(Replace(targetadd, "E", "J")).Font.Color = RGB(51, 102, 255)
            'MsgBox ("=MAX(0," & "ROUND(" & Replace(TargetAdd, "E", "H") & ",4)" & "*" & "ROUND(" & Replace(TargetAdd, "E", "I") & ",4)" & ")")
            Sheet53.Range(Replace(targetadd, "E", "J")).formula = "=ROUND(MAX(0," & "ROUND(" & Replace(targetadd, "E", "H") & ",4)" & "*" & "ROUND(" & Replace(targetadd, "E", "I") & ",4)" & "),0)"
        End If
    Next
    
    'Non Editable fields
    Sheet53.Range("TotalSaleValue_112A").NumberFormat = "0"
    Sheet53.Range("COAwithoutIndex_112A").NumberFormat = "0"
    Sheet53.Range("LTCGAssetAcquired_112A").NumberFormat = "0"
    Sheet53.Range("TotalFairMarketValue_112A").NumberFormat = "0"
    Sheet53.Range("TotalDeductions_112A").NumberFormat = "0"
    Sheet53.Range("Balance_112A").NumberFormat = "0"
    'Total Fields
    Sheet53.Range("Total_TotalSaleValue_112A").NumberFormat = "0"
    Sheet53.Range("Total_COAwithoutIndex_112A").NumberFormat = "0"
    Sheet53.Range("Total_COAwithIndex_112A").NumberFormat = "0"
    Sheet53.Range("Total_LTCGAssetAcquired_112A").NumberFormat = "0"
    Sheet53.Range("Total_TotalFairMarketValue_112A").NumberFormat = "0"
    Sheet53.Range("Total_ExpenditureWholly_112A").NumberFormat = "0"
    Sheet53.Range("Total_TotalDeductions_112A").NumberFormat = "0"
    Sheet53.Range("Total_Balance_112A").NumberFormat = "0"

    
    Sheet53.Protect Password:=getmsgstate
    Application.EnableEvents = True
End Function

Sub LOCKRANGE115AD(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet54.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = True
                    chcell.MergeArea.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                Else
                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                End If
            Next
        Next
LockFlag115A = True
End Sub
Sub UNLOCKRANGE115AD(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet54.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = False
                    chcell.MergeArea.Interior.Color = "&HCCFFCC"
                Else
                    chcell.Locked = False
                    chcell.Interior.Color = "&HCCFFCC"
                    
                    'changed by Chetan C M on 30/01/2026 for AY 2026-27
                    'start--
                     'If rangenamestring(i) = "ShareAcq_115AD_1" Or rangenamestring(i) = "ShareTrans_115AD_1" Then 'Added by Aavula 2025-26
                     If rangenamestring(i) = "ShareAcq_115AD_1" Then
                     '--end
                     
                        chcell.value = "(Select)"
                    End If
                End If
            Next
        Next
LockFlag115A = False
End Sub

Sub LOCKRANGE112AD(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet53.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = True
                    chcell.MergeArea.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                Else
                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.value = ""
                End If
            Next
        Next
LockFlag112A = True
End Sub
Sub UNLOCKRANGE112AD(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet53.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = False
                    chcell.MergeArea.Interior.Color = "&HCCFFCC"
                Else
                    chcell.Locked = False
                    chcell.Interior.Color = "&HCCFFCC"
                    
                    'changed by Chetan C M on 30/01/2026 for AY 2026-27
                    'start--
''                    If rangenamestring(i) = "ShareAcq_115AD_1" Then
'                     If rangenamestring(i) = "ShareAcq_112A" Or rangenamestring(i) = "ShareTra_112A" Then 'Added by Aavula 2025-26
                      If rangenamestring(i) = "ShareAcq_112A" Then
                     '--end
                     
                        chcell.value = "(Select)"
                    End If
                End If
            Next
        Next
LockFlag112A = False
End Sub
