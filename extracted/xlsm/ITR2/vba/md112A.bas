Attribute VB_Name = "md112A"
Option Explicit

Public MsgBox_112A As String

Public TotalSaleValue_112A_1T As Variant
Public ShareAcq_112A As Variant
'Added by Riyaz on 23/12/2024
'Public ShareTra_112A As Variant        'commented by Chetan C M on 20/01/2026 for AY 2026-27
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

'commented by Chetan C M on 20/01/2026 for AY2026-27
    'start--
''added by Chetan C M on 03/03/2025
''start--
'Public Total_Balance_112ABE As Variant
'Public Total_Balance_112AAE As Variant
'Public TotBalance_112_ABE_AAE As Variant    'added by Chetan C M on 29/10/2025 for SIT-89185
''--end
    '--end

Public LockFlag112A, LockFlag115A As Boolean

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
'changed by Chetan C M on 20/01/2026 for AY 2026-27
'start--
'Changed by Riyaz on 23/12/2024
'EfilingCommon.DefinedgridNameRange = "ShareAcq_112A||ShareTra_112A||ISINCode_112A||NameofShare_112A||NoofShare_112A||SalePriceperShare_112A||TotalSaleValue_112A||COAwithoutIndex_112A||COAwithIndex_112A||LTCGAssetAcquired_112A||FairMarketValue_112A||TotalFairMarketValue_112A||ExpenditureWholly_112A||TotalDeductions_112A||Balance_112A"
EfilingCommon.DefinedgridNameRange = "ShareAcq_112A||ISINCode_112A||NameofShare_112A||NoofShare_112A||SalePriceperShare_112A||TotalSaleValue_112A||COAwithoutIndex_112A||COAwithIndex_112A||LTCGAssetAcquired_112A||FairMarketValue_112A||TotalFairMarketValue_112A||ExpenditureWholly_112A||TotalDeductions_112A||Balance_112A"
'--end
ActiveCellRange = EfilingCommon.searchLastRow("ISINCode_112A")
EfilingCommon.insertRowUnderSectionWithFormula f_112A_115AD_1:=1
End Sub

'Ankita(23-Jan)

Function IBANCOUNTRY7()
Sheet53.Range("ShareAcq_112A").value = "(Select)"
End Function

'commented by Chetan C M on 20/01/2026 for AY 2026-27
    'start--
'Function IBANTra()
'Sheet53.Range("ShareTra_112A").value = "(Select)"
'End Function
    '--end
    
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
    rngname_112A = "ShareAcq_112A;ISINCode_112A;NameofShare_112A;NoofShare_112A;SalePriceperShare_112A;TotalSaleValue_112A;COAwithoutIndex_112A;COAwithIndex_112A;LTCGAssetAcquired_112A;FairMarketValue_112A;TotalFairMarketValue_112A;ExpenditureWholly_112A;TotalDeductions_112A;Balance_112A;"
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
        If Not (rangecells.item(mIntCtr).value = "(Select)" Or rangecells.item(mIntCtr).value = "") Then
        ccount = ccount + 1
        End If
    Next
    end112A8 = ccount
End Sub
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
        
        If ShareAcq_112A(i) = "(Select)" Or ShareAcq_112A(i) = "" Then
           ' MsgBox_112A = MsgBox_112A + "* Please select the dropdown option in Column 1b of Schedule 112A" & Chr(13)
        'Added by Shrutika(AY-25)
           ' MsgBox_112A = MsgBox_112A + "* Please select the dropdown option in Column 1b of Schedule 112A." & Chr(13)
'Changed by sudha on 5th march 2025
        MsgBox_112A = MsgBox_112A + "* Please select the dropdown option in Column 1a of Schedule 112A." & Chr(13)
            ValidateShareAcq_112A = False
            Exit Function
        End If
    Next
End Function

'commented by Chetan C M on 20/01/2026 for AY 2026-27
    'start--
''Function added by Riyaz 23/12/2024
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
           ' MsgBox_112A = MsgBox_112A + "* Please enter the ISIN Code. at Sr. No." & i & " in sheet 112A" & Chr(13)
            'changed by Chetan C M on 25/10/2025
            'start--
'            MsgBox_112A = MsgBox_112A + "* Please enter the ISIN Code in Column 2 of Schedule 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Please enter the ISIN Code" & Chr(13)
            '--end
            
            ValidateISINCode_112A = False
            Exit Function
        End If
        
        If Len(ISINCode_112A(i)) > 0 Then
            If Len(ISINCode_112A(i)) <> 12 Then
               ' MsgBox_112A = MsgBox_112A + "* Please enter the valid ISIN Code. at Sr. No." & i & " in sheet 112A" & Chr(13)
                MsgBox_112A = MsgBox_112A + "* Please enter the valid ISIN Code in Column 2 of Schedule 112A" & Chr(13)
                ValidateISINCode_112A = False
                Exit Function
            End If
        End If
        
         
        If UCase(Trim(Mid(ISINCode_112A(i), 1, 2))) <> "IN" Then
           ' MsgBox_112A = MsgBox_112A + "* Please enter the valid ISIN Code. at Sr. No." & i & " in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Please enter the valid ISIN Code in Column 2 of Schedule 112A" & Chr(13)
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
           ' MsgBox_112A = MsgBox_112A + "* Please enter Name of the Share/Unit at Sr. No." & i & " in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Please enter Name of the Share/Unit in Column 3 of Schedule 112A" & Chr(13)
            ValidateNameofShare_112A = False
            Exit Function
        End If
        
        If Len(NameofShare_112A(i)) > 125 Then
            MsgBox_112A = MsgBox_112A + "* Name of the Share/Unit at Sr. No." & i & " should not exceed 125 characters.in sheet 112A" & Chr(13)
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
    'added by Chetan C M on 25/10/2025 for SIT-89185====start
    Dim int_num As String, decimal_num As String, find_num As Integer
    '===end
    
    For i = 1 To end112A
        NoofShare_112A(i) = rangecells.item(i).value
        
        If Not chkMandatory(NoofShare_112A(i)) Then
          '  MsgBox_112A = MsgBox_112A + "* Please enter the No. of Shares/Units at Sr. No." & i & " in sheet 112A" & Chr(13)
             MsgBox_112A = MsgBox_112A + "* Please enter the No. of Shares/Units in Column 4 of Schedule 112A" & Chr(13)
            ValidateNoofShare_112A = False
            Exit Function
        End If
        
        'added by Chetan C M on 25/10/2025 for SIT-89185
        'start--
        Dim tmpArray
        If CDec(NoofShare_112A(i)) < 0 Then
            MsgBox_112A = MsgBox_112A + "* No. of Shares/Units at Sr. No." & i & " cannot be negative in Schedule 112A" & Chr(13)
            ValidateNoofShare_112A = False
            Exit Function
        End If
        
        tmpArray = Split(NoofShare_112A(i), ".")
        If Len(NoofShare_112A(i)) > 19 Or Len(tmpArray(0)) > 14 Then
            MsgBox_112A = MsgBox_112A + "* No. of Shares/Units at Sr. No." & i & " cannot exceed 19 digits in Schedule 112A" & Chr(13)
            ValidateNoofShare_112A = False
            Exit Function
        End If
        
         find_num = InStr(1, NoofShare_112A(i), ".")

        If find_num = 1 Then
            MsgBox_112A = MsgBox_112A + "* No. of Shares/Units should start with a number." & Chr(13)
            ValidateNoofShare_112A = False
            Exit Function
        End If

        If find_num > 0 Then
            int_num = Left(NoofShare_112A(i), find_num - 1)
            decimal_num = Mid(NoofShare_112A(i), find_num + 1)
        Else
            int_num = NoofShare_112A(i)
            decimal_num = ""
        End If

        If Not IsNumeric(int_num) Then
            MsgBox_112A = MsgBox_112A + "* No. of Shares/Units should be a numeric value." & Chr(13)
            ValidateNoofShare_112A = False
            Exit Function
        End If

        If Len(int_num) > 14 Or val(int_num) > 99999999999999# Then
            MsgBox_112A = MsgBox_112A + "* No. of Shares/Units cannot exceed 14 digits before decimal." & Chr(13)
            ValidateNoofShare_112A = False
            Exit Function
        End If

        If Len(decimal_num) > 0 Then
            If Not IsNumeric(decimal_num) Then
                MsgBox_112A = MsgBox_112A + "* No. of Shares/Units should be a numeric value after decimal." & Chr(13)
                ValidateNoofShare_112A = False
                Exit Function
            End If

            If Len(decimal_num) > 4 Then
                MsgBox_112A = MsgBox_112A + "* No. of Shares/Units cannot exceed 4 digits after decimal." & Chr(13)
                ValidateNoofShare_112A = False
                Exit Function
            End If
        End If
        '--end
        
    Next
End Function

Function ValidateSalePriceperShare_112A() As Boolean
    ValidateSalePriceperShare_112A = True
    Dim rangecells As Range
    Set rangecells = Sheet53.Range("SalePriceperShare_112A").Cells
    Dim i As Long
    ReDim SalePriceperShare_112A(end112A)
    'added by Chetan C M on 25/10/2025 for SIT-89185====start
    Dim int_num As String, decimal_num As String, find_num As Integer
    '===end
    
    For i = 1 To end112A
        SalePriceperShare_112A(i) = rangecells.item(i).value
        
        If Not chkMandatory(SalePriceperShare_112A(i)) Then
           ' MsgBox_112A = MsgBox_112A + "* Please enter the Sale-price per Share/Unit at Sr. No." & i & " in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Please enter the Sale-price per Share/Unit in Column 5 of Schedule 112A" & Chr(13)
            ValidateSalePriceperShare_112A = False
            Exit Function
        End If
        
        'added by Chetan C M on 25/10/2025 for SIT-89185
        'start--
        Dim tmpArray
        If CDec(SalePriceperShare_112A(i)) < 0 Then
            MsgBox_112A = MsgBox_112A + "* Sale-price per Share/Unit at Sr. No." & i & " cannot be negative in Schedule 112A" & Chr(13)
            ValidateSalePriceperShare_112A = False
            Exit Function
        End If
        
        tmpArray = Split(SalePriceperShare_112A(i), ".")
        If Len(SalePriceperShare_112A(i)) > 19 Or Len(tmpArray(0)) > 14 Then
            MsgBox_112A = MsgBox_112A + "* Sale-price per Share/Unit at Sr. No." & i & " cannot exceed 19 digits in Schedule 112A" & Chr(13)
            ValidateSalePriceperShare_112A = False
            Exit Function
        End If
        
        find_num = InStr(1, SalePriceperShare_112A(i), ".")

        If find_num = 1 Then
            MsgBox_112A = MsgBox_112A + "* Sale-price per Share/Unit should starts with number." & Chr(13)
            ValidateSalePriceperShare_112A = False
            Exit Function
        End If

        If find_num > 0 Then
            int_num = Left(SalePriceperShare_112A(i), find_num - 1)
            decimal_num = Mid(SalePriceperShare_112A(i), find_num + 1)
        Else
            int_num = SalePriceperShare_112A(i)
            decimal_num = ""
        End If

        If Not IsNumeric(int_num) Then
            MsgBox_112A = MsgBox_112A + "* Sale-price per Share/Unit should be a numeric value." & Chr(13)
            ValidateSalePriceperShare_112A = False
            Exit Function
        End If

        If Len(int_num) > 14 Or val(int_num) > 99999999999999# Then
            MsgBox_112A = MsgBox_112A + "* Sale-price per Share/Unit cannot exceed 14 digits before decimal." & Chr(13)
            ValidateSalePriceperShare_112A = False
            Exit Function
        End If

        If Len(decimal_num) > 0 Then
            If Not IsNumeric(decimal_num) Then
                MsgBox_112A = MsgBox_112A + "* Sale-price per Share/Unit should be a numeric value after decimal." & Chr(13)
                ValidateSalePriceperShare_112A = False
                Exit Function
            End If

            If Len(decimal_num) > 4 Then
                MsgBox_112A = MsgBox_112A + "* Sale-price per Share/Unit cannot exceed 4 digits after decimal." & Chr(13)
                ValidateSalePriceperShare_112A = False
                Exit Function
            End If
        End If
        
        '--end
        
    Next
End Function

Function ValidateCOAwithIndex_112A() As Boolean
    ValidateCOAwithIndex_112A = True
    Dim rangecells As Range
    Set rangecells = Sheet53.Range("COAwithIndex_112A").Cells
    Dim i As Long
    ReDim COAwithIndex_112A(end112A)
    'added by Chetan C M on 25/10/2025 for SIT-89185====start
    Dim int_num As String, decimal_num As String, find_num As Integer
    '===end
    
    For i = 1 To end112A
        COAwithIndex_112A(i) = rangecells.item(i).value
        
        If Not chkMandatory(COAwithIndex_112A(i)) Then
           ' MsgBox_112A = MsgBox_112A + "* Please enter the Cost of acquisition at Sr. No." & i & " in sheet 112A" & Chr(13)
             MsgBox_112A = MsgBox_112A + "* Please enter the Cost of acquisition in Column 8 of Schedule 112A" & Chr(13)
            ValidateCOAwithIndex_112A = False
            Exit Function
        End If
        
        'added by Chetan C M on 25/10/2025 for SIT-89185====start
        find_num = InStr(1, COAwithIndex_112A(i), ".")

        If find_num = 1 Then
            MsgBox_112A = MsgBox_112A + "* Cost of acquisition should start with a number." & Chr(13)
            ValidateCOAwithIndex_112A = False
            Exit Function
        End If

        If find_num > 0 Then
            int_num = Left(COAwithIndex_112A(i), find_num - 1)
            decimal_num = Mid(COAwithIndex_112A(i), find_num + 1)
        Else
            int_num = COAwithIndex_112A(i)
            decimal_num = ""
        End If

        If Not IsNumeric(int_num) Then
            MsgBox_112A = MsgBox_112A + "* Cost of acquisition should be a numeric value." & Chr(13)
            ValidateCOAwithIndex_112A = False
            Exit Function
        End If

        If Len(int_num) > 14 Or val(int_num) > 99999999999999# Then
            MsgBox_112A = MsgBox_112A + "* Cost of acquisition cannot exceed 14 digits before decimal." & Chr(13)
            ValidateCOAwithIndex_112A = False
            Exit Function
        End If

        If Len(decimal_num) > 0 Then
            If Not IsNumeric(decimal_num) Then
                MsgBox_112A = MsgBox_112A + "* Cost of acquisition should be a numeric value after decimal." & Chr(13)
                ValidateCOAwithIndex_112A = False
                Exit Function
            End If

            If Len(decimal_num) > 4 Then
                MsgBox_112A = MsgBox_112A + "* Cost of acquisition cannot exceed 4 digits after decimal." & Chr(13)
                ValidateCOAwithIndex_112A = False
                Exit Function
            End If
        End If

        '===end
        
    Next
End Function

Function ValidateFairMarketValue_112A() As Boolean
    ValidateFairMarketValue_112A = True
    Dim rangecells As Range
    Set rangecells = Sheet53.Range("FairMarketValue_112A").Cells
    Dim i As Long
    ReDim FairMarketValue_112A(end112A)
    'added by Chetan C M on 25/10/2025 for SIT-89185====start
    Dim int_num As String, decimal_num As String, find_num As Integer
    '===end
    
    For i = 1 To end112A
        FairMarketValue_112A(i) = rangecells.item(i).value
        
        If Not chkMandatory(FairMarketValue_112A(i)) Then
         '   MsgBox_112A = MsgBox_112A + "* Please enter the Fair Market Value per share/unit as on 31st January,2018. If share is Aquired after 31st Jan 2018 , please enter ""0"" at Sr. No." & i & " in sheet 112A" & Chr(13)
           'Commented by Shrutika
           '  MsgBox_112A = MsgBox_112A + "* Please enter the Fair Market Value per share/unit as on 31st January,2018 in Column 10 of Schedule 112A" & Chr(13)
            'Added by Shrutika
            MsgBox_112A = MsgBox_112A + "* ""Please enter the Fair Market Value per share/unit as on 31st January,2018 in Column 10 of Schedule 112A."" If share is Acquired after 31st Jan 2018 , please enter ""0"""
            
            ValidateFairMarketValue_112A = False
            Exit Function
        End If
        
        'added by Chetan C M on 25/10/2025 for SIT-89185
        'start--
        Dim tmpArray
        If CDec(FairMarketValue_112A(i)) < 0 Then
            MsgBox_112A = MsgBox_112A + "* Fair Market Value per share/unit as on 31st January,2018 in Column 10,If share is Acquired after 31st Jan 2018 , please enter '0' at Sr. No." & i & " cannot be negative in Schedule 112A" & Chr(13)
            ValidateFairMarketValue_112A = False
            Exit Function
        End If
        
        tmpArray = Split(FairMarketValue_112A(i), ".")
        If Len(FairMarketValue_112A(i)) > 19 Or Len(tmpArray(0)) > 14 Then
            MsgBox_112A = MsgBox_112A + "* Fair Market Value per share/unit as on 31st January,2018 in Column 10,If share is Acquired after 31st Jan 2018 , please enter '0' at Sr. No." & i & " cannot exceed 19 digits in Schedule 112A" & Chr(13)
            ValidateFairMarketValue_112A = False
            Exit Function
        End If
        
        find_num = InStr(1, FairMarketValue_112A(i), ".")

        If find_num = 1 Then
            MsgBox_112A = MsgBox_112A + "* Fair Market Value per share/unit should start with a number." & Chr(13)
            ValidateFairMarketValue_112A = False
            Exit Function
        End If

        If find_num > 0 Then
            int_num = Left(FairMarketValue_112A(i), find_num - 1)
            decimal_num = Mid(FairMarketValue_112A(i), find_num + 1)
        Else
            int_num = FairMarketValue_112A(i)
            decimal_num = ""
        End If

        If Not IsNumeric(int_num) Then
            MsgBox_112A = MsgBox_112A + "* Fair Market Value per share/unit should be a numeric value." & Chr(13)
            ValidateFairMarketValue_112A = False
            Exit Function
        End If

        If Len(int_num) > 14 Or val(int_num) > 99999999999999# Then
            MsgBox_112A = MsgBox_112A + "* Fair Market Value per share/unit cannot exceed 14 digits before decimal." & Chr(13)
            ValidateFairMarketValue_112A = False
            Exit Function
        End If

        If Len(decimal_num) > 0 Then
            If Not IsNumeric(decimal_num) Then
                MsgBox_112A = MsgBox_112A + "* Fair Market Value per share/unit should be a numeric value after decimal." & Chr(13)
                ValidateFairMarketValue_112A = False
                Exit Function
            End If

            If Len(decimal_num) > 4 Then
                MsgBox_112A = MsgBox_112A + "* Fair Market Value per share/unit cannot exceed 4 digits after decimal." & Chr(13)
                ValidateFairMarketValue_112A = False
                Exit Function
            End If
        End If
        
        '--end
        
    Next
End Function


Function ValidateExpenditureWholly_112A() As Boolean
    ValidateExpenditureWholly_112A = True
    Dim rangecells As Range
    Set rangecells = Sheet53.Range("ExpenditureWholly_112A").Cells
    Dim i As Long
    ReDim ExpenditureWholly_112A(end112A)
    'added by Chetan C M on 25/10/2025 for SIT-89185====start
    Dim int_num As String, decimal_num As String, find_num As Integer
    '===end
    
    For i = 1 To end112A
        ExpenditureWholly_112A(i) = rangecells.item(i).value
        
        If Not chkMandatory(ExpenditureWholly_112A(i)) Then
          '  MsgBox_112A = MsgBox_112A + "* Please enter the Expenditure incurred wholly and exclusively in connection with transfer at Sr. No." & i & " in sheet 112A" & Chr(13)
             MsgBox_112A = MsgBox_112A + "* Please enter the Expenditure wholly and exclusively in connection with transfer in Column 14 of Schedule 112A" & Chr(13)
            ValidateExpenditureWholly_112A = False
            Exit Function
        End If
        
        'added by Chetan C M on 25/10/2025 for SIT-89185====start
        find_num = InStr(1, ExpenditureWholly_112A(i), ".")

        If find_num = 1 Then
            MsgBox_112A = MsgBox_112A + "* Expenditure incurred wholly and exclusively in connection with transfer should start with a number." & Chr(13)
            ValidateExpenditureWholly_112A = False
            Exit Function
        End If

        If find_num > 0 Then
            int_num = Left(ExpenditureWholly_112A(i), find_num - 1)
            decimal_num = Mid(ExpenditureWholly_112A(i), find_num + 1)
        Else
            int_num = ExpenditureWholly_112A(i)
            decimal_num = ""
        End If

        If Not IsNumeric(int_num) Then
            MsgBox_112A = MsgBox_112A + "* Expenditure incurred wholly and exclusively in connection with transfer should be a numeric value." & Chr(13)
            ValidateExpenditureWholly_112A = False
            Exit Function
        End If

        If Len(int_num) > 14 Or val(int_num) > 99999999999999# Then
            MsgBox_112A = MsgBox_112A + "* Expenditure incurred wholly and exclusively in connection with transfer cannot exceed 14 digits before decimal." & Chr(13)
            ValidateExpenditureWholly_112A = False
            Exit Function
        End If

        If Len(decimal_num) > 0 Then
            If Not IsNumeric(decimal_num) Then
                MsgBox_112A = MsgBox_112A + "* Expenditure incurred wholly and exclusively in connection with transfer should be a numeric value after decimal." & Chr(13)
                ValidateExpenditureWholly_112A = False
                Exit Function
            End If

            If Len(decimal_num) > 4 Then
                MsgBox_112A = MsgBox_112A + "* Expenditure incurred wholly and exclusively in connection with transfer cannot exceed 4 digits after decimal." & Chr(13)
                ValidateExpenditureWholly_112A = False
                Exit Function
            End If
        End If
        '===end
        
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
        
        'changed by Chetan C M on 28/10/2025 for SIT-89185 (issue 19)
'            MsgBox_112A = MsgBox_112A + "* Total Sale Value at Sr. No " & i & " cannot exceed 14 digits in sheet 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Full Value of Consideration at Sr. No " & i & " cannot exceed 14 digits in sheet 112A" & Chr(13)
            '--end
            ValidateNonEditableFields_112A = False
'            Exit Function
        End If
        
        If Len(COAwithoutIndex_112A(i)) > 0 Then
            COAwithoutIndex_112A(i) = Application.WorksheetFunction.Round(COAwithoutIndex_112A(i), 0)
        End If
        
'        Tmparray = Split(COAwithoutIndex_112A(i), ".")
        If Len(COAwithoutIndex_112A(i)) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
            MsgBox_112A = MsgBox_112A + "* Cost of acquisition without indexation at Sr. No " & i & " cannot exceed 14 digits in sheet 112A" & Chr(13)
            ValidateNonEditableFields_112A = False
'            Exit Function
        End If
        
        If Len(LTCGAssetAcquired_112A(i)) > 0 Then
            LTCGAssetAcquired_112A(i) = Application.WorksheetFunction.Round(LTCGAssetAcquired_112A(i), 0)
        End If

'        Tmparray = Split(LTCGAssetAcquired_112A(i), ".")
        If Len(LTCGAssetAcquired_112A(i)) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
            MsgBox_112A = MsgBox_112A + "* Long term capital asset acquired before 01.02.2018 at Sr. No " & i & " cannot exceed 14 digits in sheet 112A" & Chr(13)
            ValidateNonEditableFields_112A = False
'            Exit Function
        End If
        
        If Len(TotalFairMarketValue_112A(i)) > 0 Then
        TotalFairMarketValue_112A(i) = Application.WorksheetFunction.Round(TotalFairMarketValue_112A(i), 0)
        End If

'        Tmparray = Split(TotalFairMarketValue_112A(i), ".")
        If Len(TotalFairMarketValue_112A(i)) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
            MsgBox_112A = MsgBox_112A + "* Total Fair Market Value of capital asset at Sr. No " & i & " cannot exceed 14 digits in sheet 112A" & Chr(13)
            ValidateNonEditableFields_112A = False
'            Exit Function
        End If

        If Len(TotalDeductions_112A(i)) > 0 Then
            TotalDeductions_112A(i) = Application.WorksheetFunction.Round(TotalDeductions_112A(i), 0)
        End If

'        Tmparray = Split(TotalDeductions_112A(i), ".")
        If Len(TotalDeductions_112A(i)) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
            MsgBox_112A = MsgBox_112A + "* Total deductions at Sr. No " & i & " cannot exceed 14 digits in sheet 112A" & Chr(13)
            ValidateNonEditableFields_112A = False
'            Exit Function
        End If
        
        If Len(Balance_112A(i)) > 0 Then
            Balance_112A(i) = Application.WorksheetFunction.Round(Balance_112A(i), 0)
        End If
        
'        Tmparray = Split(Balance_112A(i), ".")
        If Not Balance_112A(i) < 0 Then
            If Len(Balance_112A(i)) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
                MsgBox_112A = MsgBox_112A + "* Balance at Sr. No " & i & " cannot exceed 14 digits " & " in Schedule 112A" & Chr(13)
                ValidateNonEditableFields_112A = False
'                Exit Function
            End If
        Else
'            Tmparray = Split(CStr((Balance_112A(i) * -1)), ".")
            If Len((Balance_112A(i) * -1)) > 14 Then ' Or Len(Tmparray(0)) > 11 Then
                MsgBox_112A = MsgBox_112A + "Balance at Sr. No " & i & " cannot exceed 14 digits " & " in Schedule 112A" & Chr(13)
                ValidateNonEditableFields_112A = False
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
 Total_Balance_112A = Sheet53.Range("Total_Balance_112A").value 'added on 11/03/2025
 
 'changed by Chetan C M on 20/01/2026 for AY2026-27
    'start--
' 'added by Chetan C M on 03/03/2025
' 'start--
' Total_Balance_112ABE = Sheet53.Range("Total_Balance_112ABE").value
' Total_Balance_112AAE = Sheet53.Range("Total_Balance_112AAE").value
' TotBalance_112_ABE_AAE = Sheet53.Range("TotBalance_112_ABE_AAE").value 'added by Chetan C M on 29/10/2025 for SIT-89185
' '--end
    '--end
    
    If Len(Total_TotalSaleValue_112A) > 0 Then
        Total_TotalSaleValue_112A = Application.WorksheetFunction.Round(Total_TotalSaleValue_112A, 0)
    End If
 
    'Tmparray = Split(Total_TotalSaleValue_112A, ".")
    If Len(Total_TotalSaleValue_112A) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
    'changed by Chetan C M on 28/10/2025 for SIT-89185 (issue 19)
        'start--
'       MsgBox_112A = MsgBox_112A + "* Total_TotalSaleValue_112A cannot be greater than 14 digits in sheet 112A" & Chr(13)
        MsgBox_112A = MsgBox_112A + "* Total Full Value of Consideration cannot be greater than 14 digits in sheet 112A" & Chr(13)
        '--end
       ValidateTotalFields_112A = False
'       Exit Function
    End If
 
    If Len(Total_COAwithoutIndex_112A) > 0 Then
        Total_COAwithoutIndex_112A = Application.WorksheetFunction.Round(Total_COAwithoutIndex_112A, 0)
    End If

    'Tmparray = Split(Total_COAwithoutIndex_112A, ".")
    If Len(Total_COAwithoutIndex_112A) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
    'changed by Chetan C M on 28/10/2025 for SIT-89185 (issue 19)
    'start--
'       MsgBox_112A = MsgBox_112A + "* Total_COAwithoutIndex_112A cannot be greater than 14 digits in sheet 112A" & Chr(13)
        MsgBox_112A = MsgBox_112A + "* Total Cost of acquisition without indexation cannot be greater than 14 digits in sheet 112A" & Chr(13)
    '--end
       ValidateTotalFields_112A = False
'       Exit Function
    End If
 
    If Len(Total_COAwithIndex_112A) > 0 Then
        Total_COAwithIndex_112A = Application.WorksheetFunction.Round(Total_COAwithIndex_112A, 0)
    End If

    'Tmparray = Split(Total_COAwithIndex_112A, ".")
    If Len(Total_COAwithIndex_112A) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
    'changed by Chetan C M on 28/10/2025 for SIT-89185 (issue 19)
    'start--
'       MsgBox_112A = MsgBox_112A + "* Total_COAwithIndex_112A cannot be greater than 14 digits in sheet 112A" & Chr(13)
        MsgBox_112A = MsgBox_112A + "* Total Cost of acquisition with indexation cannot be greater than 14 digits in sheet 112A" & Chr(13)
    '--end
       ValidateTotalFields_112A = False
'       Exit Function
    End If
 
    If Len(Total_LTCGAssetAcquired_112A) > 0 Then
        Total_LTCGAssetAcquired_112A = Application.WorksheetFunction.Round(Total_LTCGAssetAcquired_112A, 0)
    End If

    'Tmparray = Split(Total_LTCGAssetAcquired_112A, ".")
    If Len(Total_LTCGAssetAcquired_112A) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
    'changed by Chetan C M on 28/10/2025 for SIT-89185 (issue 19)
    'start--
'       MsgBox_112A = MsgBox_112A + "* Total_LTCGAssetAcquired_112A cannot be greater than 14 digits in sheet 112A" & Chr(13)
        MsgBox_112A = MsgBox_112A + "* Total LTCG Asset Acquired cannot be greater than 14 digits in sheet 112A" & Chr(13)
    '--end
       ValidateTotalFields_112A = False
'       Exit Function
    End If
 
    If Len(Total_TotalFairMarketValue_112A) > 0 Then
        Total_TotalFairMarketValue_112A = Application.WorksheetFunction.Round(Total_TotalFairMarketValue_112A, 0)
    End If

    'Tmparray = Split(Total_TotalFairMarketValue_112A, ".")
    If Len(Total_TotalFairMarketValue_112A) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
    'changed by Chetan C M on 28/10/2025 for SIT-89185 (issue 19)
    'start--
'       MsgBox_112A = MsgBox_112A + "* Total_TotalFairMarketValue_112A cannot be greater than 14 digits in sheet 112A" & Chr(13)
        MsgBox_112A = MsgBox_112A + "* Total of Total Fair Market Value cannot be greater than 14 digits in sheet 112A" & Chr(13)
    '--end
       ValidateTotalFields_112A = False
'       Exit Function
    End If

    If Len(Total_ExpenditureWholly_112A) > 0 Then
        Total_ExpenditureWholly_112A = Application.WorksheetFunction.Round(Total_ExpenditureWholly_112A, 0)
    End If

    'Tmparray = Split(Total_ExpenditureWholly_112A, ".")
    If Len(Total_ExpenditureWholly_112A) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
    'changed by Chetan C M on 28/10/2025 for SIT-89185 (issue 19)
    'start--
'       MsgBox_112A = MsgBox_112A + "* Total_ExpenditureWholly_112A cannot be greater than 14 digits in sheet 112A" & Chr(13)
        MsgBox_112A = MsgBox_112A + "* Total Expenditure Wholly cannot be greater than 14 digits in sheet 112A" & Chr(13)
    '--end
       ValidateTotalFields_112A = False
'       Exit Function
    End If
 
    If Len(Total_TotalDeductions_112A) > 0 Then
        Total_TotalDeductions_112A = Application.WorksheetFunction.Round(Total_TotalDeductions_112A, 0)
    End If

    'Tmparray = Split(Total_TotalDeductions_112A, ".")
    If Len(Total_TotalDeductions_112A) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
    'changed by Chetan C M on 28/10/2025 for SIT-89185 (issue 19)
    'start--
'       MsgBox_112A = MsgBox_112A + "* Total_TotalDeductions_112A cannot be greater than 14 digits in sheet 112A" & Chr(13)
        MsgBox_112A = MsgBox_112A + "* Total of Total Deductions cannot be greater than 14 digits in sheet 112A" & Chr(13)
    '--end
       ValidateTotalFields_112A = False
'       Exit Function
    End If
 
    If Len(Total_Balance_112A) > 0 Then
        Total_Balance_112A = Application.WorksheetFunction.Round(Total_Balance_112A, 0)
    End If
    
    'Tmparray = Split(Total_Balance_112A, ".")
    If Not Total_Balance_112A < 0 Then
        If Len(Total_Balance_112A) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
        'changed by Chetan C M on 28/10/2025 for SIT-89185 (issue 19)
        'start--
'            MsgBox_112A = MsgBox_112A + "* Total_Balance_112A cannot be greater than 14 digits in Schedule 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Total Balance cannot be greater than 14 digits in Schedule 112A" & Chr(13)
        '--end
            ValidateTotalFields_112A = False
'            Exit Function
        End If
    Else
        'Tmparray = Split(CStr((Total_Balance_112A * -1)), ".")
        If Len((Total_Balance_112A * -1)) > 14 Then 'Or Len(Tmparray(0)) > 11 Then
        'changed by Chetan C M on 28/10/2025 for SIT-89185 (issue 19)
        'start--
'            MsgBox_112A = MsgBox_112A + "* Total_Balance_112A cannot be greater than 14 digits in Schedule 112A" & Chr(13)
            MsgBox_112A = MsgBox_112A + "* Total Balance cannot be greater than 14 digits in Schedule 112A" & Chr(13)
        '--end
            ValidateTotalFields_112A = False
'            Exit Function
        End If
    End If
    
    'commented by Chaten C M on 20/01/2026
        'start--
'    'added by Chetan C M on 29/10/2025
'    'start--
'
'    If Len(Total_Balance_112ABE) > 14 Then
'            MsgBox_112A = MsgBox_112A + "* Total of column (14) where transfer was before 23rd July 2024 cannot be more than 14 digits in Schedule 112A" & Chr(13)
'            ValidateTotalFields_112A = False
'
'    End If
'
'    If Len(Total_Balance_112AAE) > 14 Then
'            MsgBox_112A = MsgBox_112A + "* Total of column (14) where transfer was on or after 23rd July 2024 cannot be more than 14 digits in Schedule 112A" & Chr(13)
'            ValidateTotalFields_112A = False
'
'    End If
'
'    If Len(TotBalance_112_ABE_AAE) > 14 Then
'            MsgBox_112A = MsgBox_112A + "* Total of LTCG u/s 112A cannot be more than 14 digits in Schedule 112A" & Chr(13)
'            ValidateTotalFields_112A = False
'
'    End If
'    '--end
        '-end
        
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
'Added by Riyaz on 23/12/2024
    'If Not ValidateShareTra_112A Then Validate112A_1 = False   'commented by Chetan C M on 20/01/2026 for AY 2026-27
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
                    
                'changed by Chetan C M on 22/01/2026 for AY 2026-27
                'start--
'                    If rangenamestring(i) = "ShareAcq_115AD_1" Or rangenamestring(i) = "ShareTrans_115AD_1" Then
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
                    'Commented by Riyaz on 02/01/2025
'                    If rangenamestring(i) = "ShareAcq_115AD_1" Then
                    'changed by Chetan C M on 20/01/2026 for AY 2026-27
                        'start-
'                     If rangenamestring(i) = "ShareAcq_112A" Or rangenamestring(i) = "ShareTra_112A" Then
                    If rangenamestring(i) = "ShareAcq_112A" Then
                        '--end
                        chcell.value = "(Select)"
                    End If
                End If
            Next
        Next
LockFlag112A = False
End Sub

'added by Chetan C M on 25/10/2025 for SIT-89185
'start====

Function RoundToFourDigits(EnteredValue As String) As Variant
Dim tmpStr As String
Dim tmpArray
tmpStr = EnteredValue
tmpArray = Split(tmpStr, ".")
tmpArray(1) = CStr(Application.WorksheetFunction.Round(CDec("0." & tmpArray(1)), 4))
If Len(tmpArray(1)) <= 2 Then
    tmpArray(0) = CStr(CDec(tmpArray(0)) + CDec(tmpArray(1)))
    tmpArray(1) = "0000"
Else
    tmpArray(1) = tmpArray(1) & String(6 - (Len(tmpArray(1))), "0")
    tmpArray(1) = Right(tmpArray(1), 4)
End If
tmpStr = tmpArray(0) & "." & tmpArray(1)
RoundToFourDigits = tmpStr
End Function

'====end


