Attribute VB_Name = "SchUD"
Option Explicit

Public AssYear_UD As Variant
Public BF_UD As Variant
Public Setoff_UD As Variant
Public Balance_UD As Variant
Public Balance_UDCY As Variant
Public Balance_UDCY2 As Variant

Public BF_UD2 As Variant
Public Setoff_UD2 As Variant
Public Balance_UD2 As Variant

Public TotBalance_UD As Variant
Public TotSetoff_UD As Variant
Public TotBF_UD As Variant

Public TotBalance_UD2 As Variant
Public TotSetoff_UD2 As Variant
Public TotBF_UD2 As Variant

Public rngname_UD As Variant
Public end_UD As Variant

Public rngname_UD2 As Variant
Public end_UD2 As Variant
Public end_UD3 As Variant
Public end_UD4 As Variant
Public end_UD5 As Variant

Dim msgValidateSheetUD As String

Sub AddRows_UD()
    Dim numberofrows As Long
    setTblinfo_UD
    SelectLastRow ("UD.AssessmentYear")
    numberofrows = InsertRowsAndFillFormulasBPA()
    Call ExendRangeNameToTable(numberofrows, rngname_UD)
End Sub

Sub setTblinfo_UD()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet15.Range("UD.AssessmentYear").count
    Set rangecells = Sheet15.Range("UD.AssessmentYear").Cells
    
    For mIntCtr = 1 To mIntCells
    
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    
    end_UD = ccount
    rngname_UD = "UD.AssessmentYear;UD.BF;UD.115BAC;UD.Setoff;UD.Balance;UD2.BF;UD2.Setoff;UD2.Balance;"
End Sub

Sub setTblinfo_UD2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet15.Range("UD.BF").count
    Set rangecells = Sheet15.Range("UD.BF").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    
    end_UD2 = ccount
    rngname_UD2 = "UD.AssessmentYear;UD.BF;UD.Setoff;UD.Balance;UD2.BF;UD2.Setoff;UD2.Balance;"
End Sub
Sub setTblinfo_UD3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet15.Range("UD.Setoff").count
    Set rangecells = Sheet15.Range("UD.Setoff").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    
    end_UD3 = ccount
End Sub

Sub setTblinfo_UD4()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet15.Range("UD2.BF").count
    Set rangecells = Sheet15.Range("UD2.BF").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    
    end_UD4 = ccount
End Sub

Sub setTblinfo_UD5()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet15.Range("UD2.Setoff").count
    Set rangecells = Sheet15.Range("UD2.Setoff").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).value = "" Then
            ccount = ccount + 1
        End If
    Next
    
    end_UD5 = ccount
End Sub

Sub ValidateSheetUnabsorbedDepreciationClick()
ValidateSheetUnabsorbedDepreciation
fmsgboxStatus "Sheet Unabsorbed Depreciation(UD) is OK"
End Sub

Sub ValidateSheetUnabsorbedDepreciation()
subProcCaption = "Validating UD"
    If Not ValidatesheetUD Then
         Sheet15.Activate
         fmsgboxsmall msgValidateSheetUD ', vbOKOnly, "Error(s)!"
         CloseMsg
    Else
        'MsgBox "Sheet is Ok", vbOKOnly, "ITR-3"
    End If
End Sub


Function msgbox_UD(strmsg As String) As String
     msgValidateSheetUD = msgValidateSheetUD & strmsg & Chr(13)
End Function

Function ValidatesheetUD() As Boolean
    ValidatesheetUD = True
    
    If Not ValidateAssYear_UD() Then ValidatesheetUD = False
    
    If (Len(Sheet15.Range("UD.AssessmentYear").item(1).value) > 0) Then
        If Not ValidateBF_UD() Then ValidatesheetUD = False
        If Not ValidateSetoff_UD() Then ValidatesheetUD = False
        If Not ValidateBalance_UD() Then ValidatesheetUD = False
        If Not ValidateTotBalance_UD() Then ValidatesheetUD = False
    End If
    
    setTblinfo_UD2
    If (end_UD <> end_UD2) Then
        msgbox_UD ("* Details of Assesment Year in Sch Unabsorbed Depriciation are Mandatory")
        ValidatesheetUD = False
        Exit Function
    End If
End Function


Function ValidateAssYear_UD() As Boolean
    ValidateAssYear_UD = True
    setTblinfo_UD
    setTblinfo_UD2
    setTblinfo_UD3
    setTblinfo_UD4
    setTblinfo_UD5
    
    Dim rangecells As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Dim rangecells4 As Range
    Dim rangecells5 As Range
    Dim i As Long
    
    Set rangecells = Sheet15.Range("UD.AssessmentYear").Cells
    Set rangecells2 = Sheet15.Range("UD.BF").Cells
    Set rangecells3 = Sheet15.Range("UD.Setoff").Cells
    Set rangecells4 = Sheet15.Range("UD2.BF").Cells
    Set rangecells5 = Sheet15.Range("UD2.Setoff").Cells
    
    end_UD = WorksheetFunction.Max(0, end_UD, end_UD2, end_UD3, end_UD4, end_UD5)
    noOfProcessSub = end_UD
    
    ReDim AssYear_UD(end_UD)
    ReDim AssYear_BF(end_UD)
    ReDim AssYear_Setoff(end_UD)
    ReDim AssYear_BF2(end_UD)
    ReDim AssYear_Setoff2(end_UD)
     
    For i = 1 To end_UD
     
        AssYear_UD(i) = rangecells.item(i).value
        AssYear_BF(i) = rangecells2.item(i).value
        AssYear_Setoff(i) = rangecells3.item(i).value
        AssYear_BF2(i) = rangecells4.item(i).value
        AssYear_Setoff2(i) = rangecells5.item(i).value
        
        If Len(AssYear_UD(i)) = 0 Then
            'Exit For
        End If
        
        If (AssYear_BF(i) + AssYear_Setoff(i) + AssYear_BF2(i) + AssYear_Setoff2(i)) > 0 Then
        If Not chkCompulsory(AssYear_UD(i)) Then
            msgbox_UD ("* Assessment Year at Sr. No  " & i + 1 & "  in Sheet Unabsorbed Depriciation  is Mandatory")
            ValidateAssYear_UD = False
            Exit Function
        End If
        End If
        
        If Mid(AssYear_UD(i), 5, 1) <> "-" Then
            msgbox_UD ("* Assessment Year at Sr. No  " & i + 1 & "  in Sheet Unabsorbed Depriciation  must be of appropriate format such as 2022-23 , 2023-24 etc.")
            ValidateAssYear_UD = False
            Exit Function
        End If
                
        
        If IsNumeric(Mid(AssYear_UD(i), 6, 2)) And IsNumeric(Mid(AssYear_UD(i), 3, 2)) Then
       
         If Mid(Mid(AssYear_UD(i), 1, 4) + 1, 3, 2) <> Mid(AssYear_UD(i), 6, 2) Then
                msgbox_UD ("* Assessment Year at Sr. No  " & i + 1 & "  in Sheet Unabsorbed Depriciation  must be a consecutive year such as 2021-22 etc.")
                ValidateAssYear_UD = False
                Exit Function
            End If
        Else
            msgbox_UD ("* Assessment Year at Sr. No  " & i + 1 & "  in Sheet Unabsorbed Depriciation  must be of appropriate format such as 2018-19 , 2019-20 etc.")
            ValidateAssYear_UD = False
            Exit Function
        End If
        
        If Mid(AssYear_UD(i), 1, 4) > 2021 Then
                msgbox_UD ("* Assessment Year at Sr. No  " & i + 1 & "  in Sheet Unabsorbed Depriciation  should not exceed the AY 2021-22.")
                ValidateAssYear_UD = False
                Exit Function
            End If
        UpdateProgressBar
    Next
End Function

Function ValidateAssYear_UDold1() As Boolean
    Dim rangecells As Range
    Dim i As Long
    
    ValidateAssYear_UD = True
    setTblinfo_UD
    
    Set rangecells = Sheet15.Range("UD.AssessmentYear").Cells
    ReDim AssYear_UD(end_UD)
    noOfProcessSub = end_UD
    For i = 1 To end_UD
        AssYear_UD(i) = rangecells.item(i).value
        If Len(AssYear_UD(i)) = 0 Then
            Exit For
        End If
        
        If Not chkCompulsory(AssYear_UD(i)) Then
        
            msgbox_UD ("* Assessment Year at Sr. No  " & i + 1 & "  in Sheet Unabsorbed Depriciation  is Mandatory")
            ValidateAssYear_UD = False
            Exit Function
        End If
        
        If Mid(AssYear_UD(i), 5, 1) <> "-" Then
            msgbox_UD ("* Assessment Year at Sr. No  " & i + 1 & "  in Sheet Unabsorbed Depriciation  must be of appropriate format such as 2022-23 , 2023-24 etc.")
            ValidateAssYear_UD = False
            Exit Function
        End If
        
        If IsNumeric(Mid(AssYear_UD(i), 6, 2)) And IsNumeric(Mid(AssYear_UD(i), 3, 2)) Then
       
         If Mid(Mid(AssYear_UD(i), 1, 4) + 1, 3, 2) <> Mid(AssYear_UD(i), 6, 2) Then
                msgbox_UD ("* Assessment Year at Sr. No  " & i + 1 & "  in Sheet Unabsorbed Depriciation  must be a consecutive year such as 2021-22 etc.")
                ValidateAssYear_UD = False
                Exit Function
            End If
        Else
            msgbox_UD ("* Assessment Year at Sr. No  " & i + 1 & "  in Sheet Unabsorbed Depriciation  must be of appropriate format such as 2018-19 , 2019-20 etc.")
            ValidateAssYear_UD = False
            Exit Function
        End If
        
        If Mid(AssYear_UD(i), 1, 4) > 2021 Then
                msgbox_UD ("* Assessment Year at Sr. No  " & i + 1 & "  in Sheet Unabsorbed Depriciation  should not exceed the AY 2021-22.")
                ValidateAssYear_UD = False
                Exit Function
            End If
        UpdateProgressBar
    Next
End Function

Function ValidateBF_UD() As Boolean
    Dim i As Long
    Dim rangecells As Range
    Dim rangecells2 As Range
    ValidateBF_UD = True
    setTblinfo_UD
    
    Set rangecells = Sheet15.Range("UD.BF").Cells
    
    Set rangecells2 = Sheet15.Range("UD2.BF").Cells
    
    ReDim BF_UD(end_UD)
    ReDim BF_UD2(end_UD)
    
    For i = 1 To end_UD
        BF_UD(i) = rangecells.item(i).value
        BF_UD2(i) = rangecells2.item(i).value
        
        If Len(BF_UD(i)) = 0 Then
        'Exit For
        End If
        
        If BF_UD2(i) = "" Then
        If Not chkCompulsory(BF_UD(i)) Then
            msgbox_UD ("* Amount of brought forward unabsorbed depreciation at Sr. No  " & i + 1 & "  in Sheet Unabsorbed Depriciation is Mandatory")
            ValidateBF_UD = False
            Exit Function
        End If
        End If
        
        If BF_UD(i) = "" Then
        If Not chkCompulsory(BF_UD2(i)) Then
            msgbox_UD ("* Amount of brought forward unabsorbed allowance at Sr. No  " & i + 1 & "  in Sheet Unabsorbed Depriciation is Mandatory")
            ValidateBF_UD = False
            Exit Function
        End If
        End If

        If (BF_UD(i) < 0) Then
            msgbox_UD ("* Amount of brought forward unabsorbed depreciation at Sr. No  " & i + 1 & "  in Sheet Unabsorbed Depriciation should be Non negative, no decimal, upto 99,999,999,999,999")
            ValidateBF_UD = False
            Exit Function
        End If
        
        If (BF_UD2(i) < 0) Then
            msgbox_UD ("* Amount of brought forward unabsorbed allownace at Sr. No  " & i + 1 & "  in Sheet Unabsorbed Depriciation should be Non negative, no decimal, upto 99,999,999,999,999")
            ValidateBF_UD = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(BF_UD(i)) Then
            msgbox_UD ("* Amount of brought forward unabsorbed depreciation at Sr. No  " & i + 1 & " in Sheet Unabsorbed Depriciation  characters < > & ' " & Chr(34) & " are not allowed")
            ValidateBF_UD = False
            Exit Function
        End If
    Next
End Function

Function ValidateSetoff_UD() As Boolean
    Dim i As Long
    Dim rangecells As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Dim rangecells6 As Range
    
    ValidateSetoff_UD = True
    setTblinfo_UD

    Set rangecells = Sheet15.Range("UD.Setoff").Cells
    Set rangecells3 = Sheet15.Range("UD.BF").Cells
    
    ReDim Setoff_UD(end_UD)
    ReDim BF_UD(end_UD)

    Set rangecells2 = Sheet15.Range("UD2.Setoff").Cells
    Set rangecells6 = Sheet15.Range("UD2.BF").Cells
    
    ReDim Setoff_UD2(end_UD)
    ReDim BF_UD2(end_UD)
    
    For i = 1 To end_UD
        Setoff_UD(i) = rangecells.item(i).value
        Setoff_UD2(i) = rangecells2.item(i).value
        BF_UD(i) = rangecells3.item(i).value
        BF_UD2(i) = rangecells6.item(i).value
        
        If Len(Setoff_UD(i)) = 0 Then
        'Exit For
        End If
        
        If (Setoff_UD(i) < 0) Then
            msgbox_UD ("* Amount of depreciation set-off against current year income at Sr. No  " & i + 1 & "  in Sheet Unabsorbed Depriciation should be Non negative, no decimal, upto 99,999,999,999,999")
            ValidateSetoff_UD = False
            Exit Function
        End If
        
        If (Setoff_UD2(i) < 0) Then
            msgbox_UD ("* Amount of allowance set-off against current year income at Sr. No  " & i + 1 & "  in Sheet Unabsorbed Depriciation should be Non negative, no decimal, upto 99,999,999,999,999")
            ValidateSetoff_UD = False
            Exit Function
        End If
 
        If Not checkfieldspecialcharacter(Setoff_UD(i)) Then
            msgbox_UD ("* Amount of depreciation set-off against current year income at Sr. No  " & i + 1 & " in Sheet Unabsorbed Depriciation  characters < > & ' " & Chr(34) & " are not allowed")
            ValidateSetoff_UD = False
            Exit Function
        End If
        
        If Setoff_UD(i) > BF_UD(i) Then
            msgbox_UD ("* Amount of depreciation set-off against current year income at Sr. No  " & i + 1 & " in Sheet Unabsorbed Depriciation  cannot exceed Amount Brought forward unabsorbed depreciation ")
            ValidateSetoff_UD = False
            Exit Function
        End If
        
        If Setoff_UD2(i) > BF_UD2(i) Then
            msgbox_UD ("* Amount of allowance set-off against current year income at Sr. No  " & i + 1 & " in Sheet Unabsorbed Depriciation  cannot exceed Amount Brought forward unabsorbed allowance.")
            ValidateSetoff_UD = False
            Exit Function
        End If
    Next
End Function

Function ValidateBalance_UD() As Boolean
    Dim rangecells As Range
    Dim rangecells2 As Range
    Dim i As Long
    
    ValidateBalance_UD = True
    setTblinfo_UD
    
    Set rangecells = Sheet15.Range("UD.Balance").Cells
    ReDim Balance_UD(end_UD)
    
    Set rangecells2 = Sheet15.Range("UD2.Balance").Cells
    ReDim Balance_UD2(end_UD)
    For i = 1 To end_UD
        Balance_UD(i) = rangecells.item(i).value
        Balance_UD2(i) = rangecells2.item(i).value
        
        If Len(Balance_UD(i)) = 0 Then
        'Exit For
        End If
        
        If Not chkCompulsory(Balance_UD(i)) Then
            msgbox_UD ("* Balance Carried forward to the next year at Sr. No  " & i + 1 & "  in Sheet Unabsorbed Depriciation is Mandatory")
            ValidateBalance_UD = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(Balance_UD(i)) Then
            msgbox_UD ("* Balance Carried forward to the next year at Sr. No  " & i + 1 & " in Sheet Unabsorbed Depriciation  characters < > & ' " & Chr(34) & " are not allowed")
            ValidateBalance_UD = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotBalance_UD() As Boolean
    ValidateTotBalance_UD = True
    
    Balance_UDCY = Sheet15.Range("UDCY.Balance").value
    Balance_UDCY2 = Sheet15.Range("UDCY2.Balance").value
    
    TotBalance_UD = Sheet15.Range("UD.TotBalance").value
    TotSetoff_UD = Sheet15.Range("UD.TotSetoff").value
    TotBF_UD = Sheet15.Range("UD.TotBF").value
    TotBalance_UD2 = Sheet15.Range("UD2.TotBalance").value
    TotSetoff_UD2 = Sheet15.Range("UD2.TotSetoff").value
    TotBF_UD2 = Sheet15.Range("UD2.TotBF").value
    
    Dim BFLA_TOTDepreciationLossSetOff
    Dim BFLA_TOTAllowanceLossSetOff
    
    BFLA_TOTDepreciationLossSetOff = 0
    BFLA_TOTAllowanceLossSetOff = 0
    
    BFLA_TOTDepreciationLossSetOff = Sheet16.Range("sheet16.TotUnabsorbedDeprSetoff").value
    BFLA_TOTAllowanceLossSetOff = Sheet16.Range("sheet16.TotAllUs35cl4Setoff").value
    
    
    If (Balance_UDCY < 0) Then
        msgbox_UD ("* (5). Balance carried forward to next year in Sheet Unabsorbed Depriciation should be Non negative, no decimal, upto 99,999,999,999,999")
        ValidateTotBalance_UD = False
        Exit Function
    End If
    
    If (Balance_UDCY2 < 0) Then
        msgbox_UD ("* (8). Balance carried forward to next year  in Sheet Unabsorbed Depriciation should be Non negative, no decimal, upto 99,999,999,999,999")
        ValidateTotBalance_UD = False
        Exit Function
    End If
    
    If IIf(TotSetoff_UD = "", 0, TotSetoff_UD) <> BFLA_TOTDepreciationLossSetOff Then
        msgbox_UD ("* Amount of depreciation setoff against current year income must match with Total of brought forward depreciation set off in Sch BFLA")
        ValidateTotBalance_UD = False
        Exit Function
    End If
    
    If IIf(TotSetoff_UD2 = "", 0, TotSetoff_UD2) <> BFLA_TOTAllowanceLossSetOff Then
        msgbox_UD ("* Amount of allowance setoff against current year income must match with Total of brought forward allowance u/s 35(4) set off in Sch BFLA")
        ValidateTotBalance_UD = False
        Exit Function
    End If
    
    If Len(TotBalance_UD) > 14 Then
    msgbox_UD ("* TotBalance of Balance carried forward to the next year cannot be greater than 14 digits.")
        ValidateTotBalance_UD = False
        Exit Function
    End If
    If Len(TotSetoff_UD) > 14 Then
    msgbox_UD ("* TotBalance of Amount of depreciation set-off against the current year income cannot be greater than 14 digits.")
        ValidateTotBalance_UD = False
        Exit Function
    End If
    
    If Len(TotBF_UD) > 14 Then
    msgbox_UD ("* TotBalance of Amount of brought forward unabsorbed depreciation cannot be greater than 14 digits.")
        ValidateTotBalance_UD = False
        Exit Function
    End If
    
    If Len(TotBalance_UD2) > 14 Then
    msgbox_UD ("* TotBalance of Balance Carried forward to the next year cannot be greater than 14 digits.")
        ValidateTotBalance_UD = False
        Exit Function
    End If
    
    If Len(TotSetoff_UD2) > 14 Then
    msgbox_UD ("* TotBalance of Amount of allowance  set-off against the current year income cannot be greater than 14 digits.")
        ValidateTotBalance_UD = False
        Exit Function
    End If
    
    If Len(TotBF_UD2) > 14 Then
    msgbox_UD ("* TotBalance of amount brought forward unabsorbed depriciation cannot be greater than 14 digits.")
        ValidateTotBalance_UD = False
        Exit Function
    End If
    
End Function

