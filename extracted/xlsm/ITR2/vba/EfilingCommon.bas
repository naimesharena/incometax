Attribute VB_Name = "EfilingCommon"
Option Explicit
Public CheckSection13A As Boolean

Public Total_SheetCount As Long
Public ScheduleName As Variant
Public EnabledSchedule As Variant
Public PrintYN As Variant

'For Adding Extra Rows
Public strCurrActiveCellRange, strNewActiveCellRange As String
Public DefinedgridNameRange As Variant

Public sPassword As String

Public AssesseeDob As String
Public EfilingDate As String
Public noticeDate As String
Public DateFi92cd As String
Public UniqueNumber As String

Public MsgPastAGeneralSheet, MsgPastAGeneralSheet1 As String

Public CalculateTaxCheckFlag As Boolean

Dim RangeAddress As Variant

Dim rowcount As Long 'added by Shrutika 18/11/2025

Public pwd As String
Public currCnt1, currCnt2, currCnt3, currCnt4, currCnt5, currCnt6, currCnt7, currCnt8, currCnt9, currCnt10, currCnt11, currCnt12, currCnt13, currCnt14, currCnt15, currCnt16, currCnt17 As Variant
Public currCntvalue1, currCntvalue2, currCntvalue3, currCntvalue4, currCntvalue5, currCntvalue6, currCntvalue7, currCntvalue8, currCntvalue9, currCntvalue10, currCntvalue11, currCntvalue12, currCntvalue13, currCntvalue14, currCntvalue15, currCntvalue16, currCntvalue17, currCntvalue18, currCntvalue19, currCntvalue20, currCntvalue21 As Variant
Public currCntProfit1, currCntProfit2, currCntProfit3 As Variant
Public currCntAlwnc1, currCntAlwnc2, currCntAlwnc3, currCntAlwnc4, currCntAlwnc5, currCntAlwnc6, currCntAlwnc7, currCntAlwnc8, currCntAlwnc9, currCntAlwnc10, currCntAlwnc11, currCntAlwnc12, currCntAlwnc13, currCntAlwnc14, currCntAlwnc15, currCntAlwnc16 As Variant  'missed by Chetan added by Shrutika(05/09/2025)



      
Function CheckPANRep(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.
'ITR 1 is for individuals .So,4th character of PAN should be "P" or "H"

    CheckPANRep = True
    If Len(PAN) > 0 Then
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
            CheckPANRep = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
            CheckPANRep = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
            CheckPANRep = False
            Exit Function
        End If
        
          If Not ChkAlphabet(Mid(PAN, 4, 1)) Then
            CheckPANRep = False
            Exit Function
        End If
            
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
            CheckPANRep = False
            Exit Function
        End If
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
            CheckPANRep = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
            CheckPANRep = False
            Exit Function
        End If
    End If
End Function


'uncommented by Chetan C M for AY 2025-26
'start--
Sub AddPropertyTenant(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long


    Application.EnableEvents = False
    Sheets("House Property").Activate

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet8.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    numberofrows = iRows
    For te = 1 To Sheet8.Range("PropertySectionCOunt").value


        EfilingCommon.DefinedgridNameRange = "HP.NameofTenant||HP.PANofTenant||HP.AadharofTenant||HP.TANofTenant"
        ActiveCellRange = EfilingCommon.searchLastRow("HP.NameofTenant" & te)

        If te > 1 Then
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(numberofrows, True, te)
        Else
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(iRows, True, te)
        End If
    Next



'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet8.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    Application.EnableEvents = False
    For i = 1 To Sheet8.Range("PropertySectionCOunt").value
        newfrmsize = Sheet8.Range("NumRowsPropertyBlock").value
        newfrmsize = newfrmsize + numberofrows
        Sheet8.Range("NumRowsPropertyBlock").value = newfrmsize
    Next


endline:

'----------------Lock Password-------------------START---
   Sheet8.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub
'--end


'added by Chetan C M for AY 2025-26
'start--
Sub AddSection24b(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long


    Application.EnableEvents = False
    Sheets("House Property").Activate

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet8.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    numberofrows = iRows
    For te = 1 To Sheet8.Range("PropertySectionCOunt").value


        EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.24b||bankName.24b||loanAccNum.24b||loanDate.24b||loanAmt.24b||loanOutstanding.24b||Intrst.24b"
        ActiveCellRange = EfilingCommon.searchLastRow("LoanfrmBankOrInstitute.24b" & te)

        If te > 1 Then
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(numberofrows, True, te)
        Else
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(iRows, True, te)
        End If
    Next



'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet8.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    Application.EnableEvents = False
    For i = 1 To Sheet8.Range("PropertySectionCOunt").value
        newfrmsize = Sheet8.Range("NumRowsPropertyBlock").value
        newfrmsize = newfrmsize + numberofrows
        Sheet8.Range("NumRowsPropertyBlock").value = newfrmsize
    Next


endline:

'----------------Lock Password-------------------START---
   Sheet8.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub
'--end

Sub AddRows_Others(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long
    
    
    Application.EnableEvents = False
    Sheets("Schedule S").Activate
        
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet39.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    'numberofrows = iRows
    'For te = 1 To Sheet39.Range("SalarySectionCount").value
    
        ActiveCellRange = EfilingCommon.searchLastRow("Salary.Nature1")
        EfilingCommon.DefinedgridNameRange = "Salary.Nature1||Salary.Description1||Salary.Amount1"
        'ActiveCellRange = EfilingCommon.searchLastRow("Salary.Nature" & te)
        
       ' If te > 1 Then
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula
            
        'Else
          '  numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(iRows, True, te)
       ' End If
    'Next
    
    

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet39.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    Application.EnableEvents = False
'    For i = 1 To Sheet39.Range("SalarySectionCount").value
       ' newfrmsize = Sheet39.Range("SalarySectionItem").value
      '  newfrmsize = newfrmsize + numberofrows
        'Sheet39.Range("SalarySectionItem").value = newfrmsize
'    Next


endline:
   
'----------------Lock Password-------------------START---
   Sheet39.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub

Sub AddRowsSchCG_LTCGB8()
Dim vRows As Long
Sheets("CG").Activate
'Malli-----
'AY_2024_25_OLD
'EfilingCommon.DefinedgridNameRange = "LTCG.B9_PreviousYer_1||LTCG.B9_DeductionSection_1||LTCG.B9_AsstAcquiredYer_1||LTCG.B9_AmtUtilised_1||LTCG.B9_AmtNotUsed_1"
'CG_C283_AY_2025_26
EfilingCommon.DefinedgridNameRange = "LTCG.B9_PreviousYer_1||LTCG.B9_DeductionSection_1||LTCG.B9_AsstAcquiredYer_1||LTCG.B9_AmtUtilised_1||LTCG.B9_AmtNotUsed_1"
'---------------------------
ActiveCellRange = EfilingCommon.searchLastRow("LTCG.B9_PreviousYer_1")
vRows = EfilingCommon.insertRowUnderSectionWithFormula_CG

'prepopulating year and select for the added rows
Dim rangecells, rangecells1, rangecells2 As Range
    Dim mIntCells, q As Long
Set rangecells = Range("LTCG.B9_PreviousYer_1").Cells
Set rangecells1 = Range("LTCG.B9_DeductionSection_1").Cells
Set rangecells2 = Range("LTCG.B9_Sl_1").Cells

mIntCells = Range("LTCG.B9_PreviousYer_1").count

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Application.EnableEvents = False
For q = 1 To mIntCells
    If rangecells.item(q).value = "" Then rangecells.item(q).value = "(Select)"
    If rangecells1.item(q).value = "" Then rangecells1.item(q).value = "(Select)"
    If rangecells2.item(q).value = "" Then rangecells2.item(q).value = rangecells2.item(q).Offset(-1, 0).value + 1
Next

'Sheet13.Range("CG_LTCGB9.Sl").value = ((Sheet13.Range("CG_LTCGB9.Sl").Offset(-1, 0).value) + 1)

Application.EnableEvents = True
'----------------Lock Password-------------------START---
  Sheet13.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
End Sub

Sub AddRowsCG_LTCG()
Dim vRows As Long
Sheets("CG").Activate
EfilingCommon.DefinedgridNameRange = "LTCG.B10_LtcgAmt||LTCG.B10_ItemNo||LTCG.B10_Country||LTCG.B10_DTAA_Article||LTCG.B10_RatepTreaty||LTCG.B10_CertiStatus||LTCG.B10_SecIT||LTCG.B10_RatePITAct||LTCG.B10_ApplRate||Section_CodeSLtcgB10||Section_AmountLtcgB10||Section_LTCGbreakup"
ActiveCellRange = EfilingCommon.searchLastRow("LTCG.B10_LtcgAmt")
vRows = EfilingCommon.insertRowUnderSectionWithFormula

'prepopulating  select for the added rows
Dim rangecells, rangecells1, rangecells2 As Range
Dim mIntCells, q As Long
Set rangecells = Range("LTCG.B10_ItemNo").Cells
Set rangecells1 = Range("LTCG.B10_Country").Cells
Set rangecells2 = Range("LTCG.B10_CertiStatus").Cells
Application.EnableEvents = False
mIntCells = Range("LTCG.B10_LtcgAmt").count
For q = 1 To mIntCells
    If rangecells.item(q).value = "" Then rangecells.item(q).value = "(Select)"
    If rangecells1.item(q).value = "" Then rangecells1.item(q).value = "(Select)"
'    If rangecells2.item(q).value = "" Then rangecells2.item(q).value = "(Select)"
Next
Application.EnableEvents = True
End Sub


Sub AddRowsCG_STCG()
Dim vRows As Long
Sheets("CG").Activate
EfilingCommon.DefinedgridNameRange = "STCG.A8_AmtI||STCG.A8_IA1tA6||STCG.A8_Country||STCG.A8_DTAA_Article||STCG.A8_RatePTreaty||STCG.A8_CertiStatus||STCG.A8_SecIT||STCG.A8_RateIT||STCG.A8_ApplRate||Section_CodeStcgA8||Section_AmountStcgA8||Section_STCGbreakup"
ActiveCellRange = EfilingCommon.searchLastRow("STCG.A8_AmtI")
vRows = EfilingCommon.insertRowUnderSectionWithFormula


'prepopulating  select for the added rows
Dim rangecells, rangecells1, rangecells2 As Range
Dim mIntCells, q As Long
Set rangecells = Range("STCG.A8_IA1tA6").Cells
Set rangecells1 = Range("STCG.A8_Country").Cells
Set rangecells2 = Range("STCG.A8_CertiStatus").Cells

mIntCells = Range("STCG.A8_AmtI").count
Application.EnableEvents = False
For q = 1 To mIntCells
    If rangecells.item(q).value = "" Then rangecells.item(q).value = "(Select)"
    If rangecells1.item(q).value = "" Then rangecells1.item(q).value = "(Select)"
'    If rangecells2.item(q).value = "" Then rangecells2.item(q).value = "(Select)"
Next
Application.EnableEvents = True
End Sub
Function checkfieldspecialcharacter_Bank(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter_Bank = True
    Dim arr As Variant
    arr = Array("@", "*", "!", ".", "&", "#", "~", ";", "?", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter_Bank = False
            Exit Function
        End If
        Next
    Next
End Function
'added by Chetan C M for AY 2025-26
'start--
'method to calculate the age from dateOfBirth
'Function calculateAge(dob As Variant) As Long
'
'On Error Resume Next
''Changed year from 2023 to 2024
''Ankita_25-26
''calculateAge = (2024) - val(Mid(dob, 7, 4))
' calculateAge = (2025) - val(Mid(dob, 7, 4))
'
'    If 4 < val(Mid(dob, 4, 2)) Then
'        calculateAge = calculateAge - 1
'    ElseIf val(Mid(dob, 4, 2)) = 4 And 1 < val(Mid(dob, 1, 2)) Then
'                calculateAge = calculateAge - 1
'    End If
'
'End Function
'--end
'added by Chetan C M for AY 2025-26
'start--
Function insertRowUnderSectionWithFormula_80C(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
    '----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
'Method: DefineName after no of rows insert Method
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormula_80C = nRows
'----------------Lock Password-------------------START---
'----------------Lock Password-------------------END-----

    Application.EnableEvents = True
    ActiveSheet.Protect Password:=sPassword
End Function
'--end

Function checkListIFSC(searchvalue As String) As Boolean
    checkListIFSC = True
    

End Function

Sub apply_click()
On Error GoTo endline
    Dim ws As Worksheets
    Dim i As Long

    Application.ScreenUpdating = False
'ThisWorkbook.Unprotect Password:=EfilingCommon.getmsgstate
    'For Each ws In ActiveWorkbook.Worksheets
       ' ws.Protect Password:=EfilingCommon.getmsgstate
    'Next
    Call CreateArray
    For i = 5 To Total_SheetCount
        If EnabledSchedule(i) = "N" Then
            ThisWorkbook.Unprotect Password:=getmsgstate
            ActiveWorkbook.Sheets(ScheduleName(i)).Visible = xlSheetVeryHidden
            ThisWorkbook.Protect Password:=getmsgstate
        End If
        
        If EnabledSchedule(i) = "Y" Then
            ThisWorkbook.Unprotect Password:=getmsgstate
            ActiveWorkbook.Sheets(ScheduleName(i)).Visible = xlSheetVisible
            ThisWorkbook.Protect Password:=getmsgstate
        End If
        'ThisWorkbook.Protect Password:=EfilingCommon.getmsgstate

    Next
endline:
    Application.ScreenUpdating = True

End Sub

Sub Instruction_Sheet()
'ThisWorkbook.Unprotect Password:=EfilingCommon.getmsgstate
ThisWorkbook.Unprotect Password:=getmsgstate
Sheet34.Visible = xlSheetVisible
ThisWorkbook.Protect Password:=getmsgstate
Sheet34.Activate
'ThisWorkbook.Protect Password:=EfilingCommon.getmsgstate

End Sub
Sub HelpCSV()
'ThisWorkbook.Unprotect Password:=EfilingCommon.getmsgstate
ThisWorkbook.Unprotect Password:=getmsgstate
Sheet45.Visible = xlSheetVisible
ThisWorkbook.Protect Password:=getmsgstate
Sheet45.Activate
'ThisWorkbook.Protect Password:=EfilingCommon.getmsgstate

End Sub
'added by Chetan C M for AY 2025-26
'start--
Function insertRowUnderSectionWithFormula24b(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0, Optional HOIflag As Variant = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim sPassword As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue, sAddingRangeValue, sNewCellValueNext As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
    ''''''''''malli
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
     
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
    If HOIflag = 1 Then
        Dim rr, rrCount As Variant
        Dim RngHOI As Range
        'Set RngHOI = Selection.Resize(rowsize:=nRows + 1)
        Set RngHOI = Selection.Offset(1).Resize(rowsize:=1)
        rrCount = 1
        For Each rr In RngHOI.Rows
            If Not rrCount = 0 Then
                Range("D" & rr.row).Locked = False
                Range("D" & rr.row).Interior.Color = (&HCCFFCC)
                'Commented by Ayush_05/05/2025 for 0.7 changes
'                Range("E" & rr.row).Locked = False
'                Range("E" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("F" & rr.row).Locked = False
'                Range("F" & rr.row).Interior.Color = (&HCCFFCC)

                Range("G" & rr.row).Locked = False
                Range("G" & rr.row).Interior.Color = (&HCCFFCC)
                Range("H" & rr.row).Locked = False
                Range("H" & rr.row).Interior.Color = (&HCCFFCC)
                Range("I" & rr.row).Locked = False
                Range("I" & rr.row).Interior.Color = (&HCCFFCC)
                Range("J" & rr.row).Locked = False
                Range("J" & rr.row).Interior.Color = (&HCCFFCC)
                Range("K" & rr.row).Locked = False
                Range("K" & rr.row).Interior.Color = (&HCCFFCC)
                Range("L" & rr.row).Locked = False
                Range("L" & rr.row).Interior.Color = (&HCCFFCC)
                
            End If
            rrCount = rrCount + 1
        Next
        
    End If

    
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    sNewCellValueNext = ""
    sAddingRangeValue = ""
    
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            
            sNewCellValueNext = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + 1) 'C6 + 1 => C7
            If nRows > 1 Then
            sAddingRangeValue = sNewCellValueNext + ":" + sNewCellValue
            ElseIf nRows = 1 Then
            sAddingRangeValue = sNewCellValue
            End If
            
'            If gridRange(iCount) = "TDS2.TdsCredit" Or _
'                gridRange(iCount) = "TDS3.TdsCredit" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
'
'            If gridRange(iCount) = "TDSoth.DeductedYear" Or _
'                gridRange(iCount) = "TDSoth2.DeductedYear" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
            
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            
            sNewCellValueNext = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + 1) 'C6 + 1 => C7
            If nRows > 1 Then
            sAddingRangeValue = sNewCellValueNext + ":" + sNewCellValue
            ElseIf nRows = 1 Then
            sAddingRangeValue = sNewCellValue
            End If
            
'            If gridRange(iCount) = "TDS2.TdsCredit" Or _
'                gridRange(iCount) = "TDS3.TdsCredit" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
'
'            If gridRange(iCount) = "TDSoth.DeductedYear" Or _
'                gridRange(iCount) = "TDSoth2.DeductedYear" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
            
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormula24b = nRows

'----------------Lock Password-------------------START---
  ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
    Application.EnableEvents = True
End Function
'--end

Sub Prev_Sheet()
Dim sheet As Variant

'Malli------------
  If Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" And (ActiveSheet.name = "Part A Gen_139(8A)" Or ActiveSheet.name = "Schedule S" Or ActiveSheet.name = "Part B ATI" Or ActiveSheet.name = "Verification") Then
        'MsgBox "Malli"  11/11/2024
        'Debug.Print ActiveSheet.name
          Application.EnableEvents = False
           If ActiveSheet.name = "Part A Gen_139(8A)" Then
                    Worksheets("PART A - General").Activate
           ElseIf ActiveSheet.name = "Schedule S" Then
                    Worksheets("Part A Gen_139(8A)").Activate
           ElseIf ActiveSheet.name = "Part B ATI" Then
                    Worksheets("TDS").Activate
           ElseIf ActiveSheet.name = "Verification" Then
                    Worksheets("Part B ATI").Activate
           End If
          Application.EnableEvents = True
  Else:
'------------------
    sheet = GetPreviousSheet(ActiveSheet.name)
    If sheet <> "" Then
    'Added by Shrutika
    Application.EnableEvents = False
        Sheets(sheet).Activate
    Application.EnableEvents = True
    End If
    End If 'Malli
End Sub

Sub Next_Sheet()
 Dim sheet As Variant
 
 'Malli------------
  If Sheet1.Range("sheet1.ReturnFileSec").value = "139(8A)" And (ActiveSheet.name = "PART A - General" Or ActiveSheet.name = "Part A Gen_139(8A)" Or ActiveSheet.name = "TDS" Or ActiveSheet.name = "Part B ATI") Then
        'MsgBox "Malli"  11/11/2024
        'Debug.Print ActiveSheet.name
          Application.EnableEvents = False
           If ActiveSheet.name = "PART A - General" Then
                    Worksheets("Part A Gen_139(8A)").Activate
           ElseIf ActiveSheet.name = "Part A Gen_139(8A)" Then
                     Worksheets("Schedule S").Activate
           ElseIf ActiveSheet.name = "TDS" Then
                    Worksheets("Part B ATI").Activate
           ElseIf ActiveSheet.name = "Part B ATI" Then
                    Worksheets("Verification").Activate
           End If
          Application.EnableEvents = True
  Else:
'------------------
    sheet = GetNextSheet(ActiveSheet.name)
    If sheet <> "" Then
    'Added by Shrutika
    Application.EnableEvents = False
        Sheets(sheet).Activate
    Application.EnableEvents = True
    End If
End If 'Malli
End Sub

Function GetPreviousSheet(SheetName As Variant) As Variant
    Dim i As Long
    Dim A As Long

    Call CreateArray

    For i = 0 To Total_SheetCount
        If ScheduleName(i) = SheetName Then
            For A = i - 1 To 0 Step -1
            'Added by Riyaz
            If Sheets(ScheduleName(A)).Visible = True Then
                If EnabledSchedule(A) = "Y" Then
                    GetPreviousSheet = ScheduleName(A)
                    Exit For
                End If
                End If
            Next
            End If
    Next
End Function
Sub index_Click()
Sheet35.Activate
End Sub

Function GetNextSheet(SheetName As Variant) As Variant
    Dim i As Long
    Dim A As Long
 
    Call CreateArray
    For i = 0 To Total_SheetCount
        If UCase(ScheduleName(i)) = UCase(SheetName) Then
            For A = i + 1 To Total_SheetCount
                If UCase(EnabledSchedule(A)) = "Y" Then
                If Sheets(ScheduleName(A)).Visible = True Then
                    GetNextSheet = ScheduleName(A)
                    Exit For
                End If
                End If
            Next
        End If
    Next
End Function

Sub CreateArray()
    Dim i, s As Long

    Total_SheetCount = 0
    
'Modify by Shrutika
   ' For i = 5 To 52
    For i = 5 To 54 'Added by Shrutika(26/05/2025)
        s = i
        If Not Sheet35.Range("C" & s).value = "" Then
           Total_SheetCount = Total_SheetCount + 1
        End If
    Next
    Total_SheetCount = Total_SheetCount + 4

    ReDim ScheduleName(Total_SheetCount)
        For i = 5 To Total_SheetCount
        s = i
        ScheduleName(i) = Sheet35.Range("D" & s)
    Next

    ReDim EnabledSchedule(Total_SheetCount)
    For i = 5 To Total_SheetCount
        s = i
        EnabledSchedule(i) = Sheet35.Range("G" & s)
    Next
    
    ReDim PrintYN(Total_SheetCount)
        For i = 5 To Total_SheetCount
            s = i
            PrintYN(i) = Sheet35.Range("H" & s)
        Next
End Sub
'-------------------------------------------

'Sub WorksheetsPrint()
'On Error Resume Next
'    Dim SheetArray() As Variant
'    Dim i, iCnt As Long
'    Dim PrintMsgOP As Long
'
'    Application.ScreenUpdating = False
'
'    Call CreateArray
'
'    PrintMsgOP = MsgBox("Do you want to preview the workbook for printing?", vbOKCancel, _
'                 "Print preview")
'
'    If PrintMsgOP = 2 Then
'        CloseMsg
'    End If
'
'    iCnt = -1
'
'    For i = 5 To Total_SheetCount - 2
'        If PrintYN(i) = "Y" Then
'            iCnt = iCnt + 1
'            ReDim Preserve SheetArray(iCnt)
'            SheetArray(iCnt) = ScheduleName(i)
'
'            With Sheets(ScheduleName(i)).PageSetup
'                .BlackAndWhite = True
'                .CenterHorizontally = True
'                .CenterVertically = False
'                .LeftMargin = 0.4
'                .RightMargin = 0.4
'                .PaperSize = xlPaperA4
'                .FitToPagesWide = 1
'                .Orientation = xlLandscape
'            End With
'
'        End If
'    Next
'
'    Sheets(SheetArray).PrintPreview
'
'    Sheets(1).Select
'
'    ThisWorkbook.PrintOut , preview:=True
'
'    Application.ScreenUpdating = True
'End Sub

Sub WorksheetsPrint()
    Dim PrintMsgOP As Long
    Sheet1.Activate
    
    PrintMsgOP = MsgBox("Do you want to preview the workbook for printing?", vbOKCancel, "Print preview")
    If PrintMsgOP = 1 Then
'           With Sheet1.PageSetup
'               ' .PrintArea = Sheet1.Range("PART A - General").Address
'                .BlackAndWhite = True
'                .CenterHorizontally = True
'                .CenterVertically = False
'                .LeftMargin = 0.4
'                .RightMargin = 0.4
'                .PaperSize = xlPaperA3
'                .FitToPagesWide = 1
'                .FitToPagesTall = False
'                .Zoom = False
'                .Orientation = xlLandscape
'            End With

             With Sheet35.PageSetup 'Home
               ' .PrintArea = Sheet35.Range("Home").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With

                With Sheet1.PageSetup 'PART A GENERAL
               ' .PrintArea = Sheet1.Range("PART A - General").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            With Sheet201.PageSetup 'Part A Gen_139(8A)
               ' .PrintArea = Sheet201.Range("Part A Gen_139(8A)").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
             With Sheet39.PageSetup 'Schedule S
              '  .PrintArea = Sheet39.Range("Schedule S").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            With Sheet8.PageSetup 'House Property
                '.PrintArea = Sheet8.Range("House Property").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            With Sheet13.PageSetup 'CG
                '.PrintArea = Sheet13.Range("CG").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            With Sheet53.PageSetup 'Schedule 112A
               ' .PrintArea = Sheet53.Range("Schedule 112A").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
          With Sheet54.PageSetup 'Schedule 115AD(1)(iii) proviso
               ' .PrintArea = Sheet54.Range("Schedule 115AD(1)(iii) proviso").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
        With Sheet47.PageSetup 'VDA
               ' .PrintArea = Sheet47.Range("VDA").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
        With Sheet14.PageSetup 'OS
               ' .PrintArea = Sheet14.Range("OS").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
        With Sheet16.PageSetup 'CYLA - BFLA
                '.PrintArea = Sheet16.Range("CYLA - BFLA").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
             With Sheet17.PageSetup 'CFL
               ' .PrintArea = Sheet17.Range("CFL").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
             With Sheet61.PageSetup '80C
               ' .PrintArea = Sheet61.Range("80C").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
             With Sheet20.PageSetup '80G
               ' .PrintArea = Sheet20.Range("80G").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
             With Sheet55.PageSetup '80D
               ' .PrintArea = Sheet55.Range("80D").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
             With Sheet56.PageSetup '80GGA
               ' .PrintArea = Sheet56.Range("80GGA").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
             With Sheet60.PageSetup '80E_80EE_80EEA_80EEB
               ' .PrintArea = Sheet60.Range("80E_80EE_80EEA_80EEB").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            With Sheet19.PageSetup 'VI-A
               ' .PrintArea = Sheet19.Range("VI-A").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            With Sheet21.PageSetup 'SPI - SI
               '.PrintArea = Sheet201.Range("SPI - SI").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            With Sheet23.PageSetup 'AMTC
                '.PrintArea = Sheet23.Range("AMTC").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            With Sheet24.PageSetup 'EI
               ' .PrintArea = Sheet24.Range("EI").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
             With Sheet41.PageSetup 'PTI
               ' .PrintArea = Sheet41.Range("PTI").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
        With Sheet40.PageSetup 'FSI
               ' .PrintArea = Sheet40.Range("FSI").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
        With Sheet29.PageSetup 'Sch 5A
               ' .PrintArea = Sheet29.Range("Sch 5A").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
        With Sheet27.PageSetup 'TR_FA
                '.PrintArea = Sheet27.Range("TR_FA").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            With Sheet43.PageSetup 'AL
               ' .PrintArea = Sheet43.Range("AL").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            With Sheet7.PageSetup 'Part B - TI TTI
               ' .PrintArea = Sheet7.Range("Part B - TI TTI").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            With Sheet46.PageSetup 'ESOP
              '  .PrintArea = Sheet46.Range("ESOP").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            With Sheet28.PageSetup 'TDS
               ' .PrintArea = Sheet28.Range("TDS").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            With Sheet202.PageSetup 'Part B ATI
               ' .PrintArea = Sheet202.Range("Part B ATI").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            With Sheet52.PageSetup 'Verification
               ' .PrintArea = Sheet52.Range("Verification").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.6
                .RightMargin = 0.6
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
     rowcount = ThisWorkbook.Worksheets.count
    ThisWorkbook.PrintOut , Preview:=True
    Sheet1.Activate
 Else
 
 End If

End Sub


Function InsertRowsAndFillFormulasBPA(Optional vRows As Long = 0) As Long
    Dim x As Long
    Dim sht As Worksheet, shts() As String, i As Long
    
    sPassword = EfilingCommon.getmsgstate
    ActiveSheet.Unprotect Password:=sPassword
   
    ActiveCell.EntireRow.Select
    If vRows = 0 Then
        vRows = Application.InputBox(prompt:= _
        "Enter the number of rows you want to add below selected cell", Title:="Add Rows below the selected cell", _
        Default:=1, Type:=1) 'Default for 1 row, type 1 is number
    
        If vRows = False Then
            InsertRowsAndFillFormulasBPA = 0
            Exit Function
        End If
    End If
    

    ReDim shts(1 To Worksheets.Application.ActiveWorkbook. _
    Windows(1).SelectedSheets.count)
    i = 0
    For Each sht In _
        Application.ActiveWorkbook.Windows(1).SelectedSheets
        Sheets(sht.name).Select
            i = i + 1
            shts(i) = sht.name
            x = Sheets(sht.name).UsedRange.Rows.count 'lastcell fixup
        
            Selection.Resize(rowsize:=2).Rows(2).EntireRow. _
            Resize(rowsize:=vRows).Insert Shift:=xlDown
            
            sPassword = EfilingCommon.getmsgstate
            ActiveSheet.Unprotect Password:=sPassword
            
            Selection.AutoFill Selection.Resize( _
            rowsize:=vRows + 1), xlFillDefault
        
            On Error Resume Next
            Selection.Offset(1).Resize(vRows).EntireRow. _
            SpecialCells(xlConstants).ClearContents
    Next sht
    
    Worksheets(shts).Select
    InsertRowsAndFillFormulasBPA = vRows
    ActiveSheet.Protect Password:=sPassword
End Function
 
 Sub SelectLastRow(rngname As String)
    Dim i As Long
    RangeAddress = Range(rngname).AddressLocal
    RangeAddress = Replace(RangeAddress, "$", "")
    If InStr(1, RangeAddress, ":") > 0 Then
        RangeAddress = Mid(RangeAddress, InStr(1, RangeAddress, ":") + 1, Len(RangeAddress))
        Range(RangeAddress).Select
    Else
        Range(RangeAddress).Select
    End If
End Sub
 
 
Function chkNumeric(field As Variant) As Boolean
    Dim k As Long
    Dim chkchar As String

    chkNumeric = True
    For k = 1 To Len(field)
    chkchar = Mid(field, k, 1)
        If Not IsNumeric(chkchar) Then
            chkNumeric = False
            Exit Function
        End If
    Next
End Function
 Function insertRowUnderSectionWithFormula_CG(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
'Method : Insert no of Rows
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         
        If nRows = 0 Then
         Exit Function
        End If
    Else
      nRows = nOfRows
    End If
    
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    'Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlCellTypeAllValidation).ClearContents
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
'Method: DefineName after no of rows insert Method
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$D$7:$D$9 => D7:D9
             
            If (InStr(sTempCellValue, ":") = 0) Then
                sTempCellValue = sTempCellValue & ":" & sTempCellValue
            End If
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'D7:D9 => D7
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'D7:D9 => D9
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'D9 => 9
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'D9 + nRows(eg:1) => D10
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'D7:D11
            Set srange = Range(sRangeValue)  '$D7&:$d$11
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$D$7:$D$9 => D7:D9
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'D7:D9 => D7
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'D7:D9 => D9
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'D9 => 9
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'D9 + nRows(eg:1) => D10
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'D7:D11
            Set srange = Range(sRangeValue)  '$D7&:$d$11
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormula_CG = nRows
'----------------Lock Password-------------------START---
   ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

    Application.EnableEvents = True
End Function
 'To Search Last Row of the Grid
Function searchLastRow(ByVal gridRangeName As String) As String
    On Error Resume Next
    strCurrActiveCellRange = Replace(ActiveSheet.Range(gridRangeName).AddressLocal, "$", "")
    strNewActiveCellRange = Mid(strCurrActiveCellRange, InStr(1, strCurrActiveCellRange, ":") + 1, Len(strCurrActiveCellRange))
    ActiveSheet.Range(strNewActiveCellRange).Select
    searchLastRow = strCurrActiveCellRange
    tempcheck = strNewActiveCellRange
End Function

'Bascically to insert Rows by giving input box - Yogesh Common for grids


Function insertRowUnderSectionWithFormula(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0, Optional f_112A_115AD_1 As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
'Method : Insert no of Rows
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
    
     
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    
    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    'Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlCellTypeAllValidation).ClearContents
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
    If f_112A_115AD_1 = 1 Then
        Dim rr, rrCount As Variant
        Dim RngHOI As Range
        Set RngHOI = Selection.Resize(rowsize:=nRows + 1)
        rrCount = 0
        For Each rr In RngHOI.Rows
            If Not rrCount = 0 Then
                Range("E" & rr.row).value = "(Select)"
                

 'commented by Chetan C M for AY 2026-27
  'start--
  
''Added by Riyaz on 24/12/2024
'                Range("F" & rr.row).value = "(Select)"
'
''Added by sudha on 25th dec 24
'
'                Range("F" & rr.row).value = "(Select)"
                 
' '----------------

    '--end

                
            End If
            rrCount = rrCount + 1
        Next
    End If

    
'Method: DefineName after no of rows insert Method
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue
            sRangeValue1 = sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue
            sRangeValue1 = sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormula = nRows
'----------------Lock Password-------------------START---
   ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

    Application.EnableEvents = True
End Function
Function insertRowUnderSectionWithFormula_80GGC(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0, Optional f_112A_115AD_1 As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
'Method : Insert no of Rows
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
    
     
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    
    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    'Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlCellTypeAllValidation).ClearContents
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    

    
'Method: DefineName after no of rows insert Method
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue
            sRangeValue1 = sNewCellValue 'C4:C7
            
          If gridRange(iCount) <> "TotalDonation_80GGC" And gridRange(iCount) <> "EligibleAmountofDonation_80GGC" Then
            Range(sNewCellValue).Locked = False
            Range(sNewCellValue).Interior.Color = &HCCFFCC
            
        End If
            
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
            
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue
            sRangeValue1 = sNewCellValue 'C4:C7
            
        If gridRange(iCount) <> "TotalDonation_80GGC" Or gridRange(iCount) <> "EligibleAmountofDonation_80GGC" Then
           
            Range(sNewCellValue).Locked = False
            Range(sNewCellValue).Interior.Color = &HCCFFCC
        End If
            
            
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
            
        Next
    
    End If
    insertRowUnderSectionWithFormula_80GGC = nRows
'----------------Lock Password-------------------START---
   ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

    Application.EnableEvents = True
End Function

Function insertRowUnderSectionWithFormula1(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
'Method : Insert no of Rows
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
    
     
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    
    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    'Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlCellTypeAllValidation).ClearContents
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
'Method: DefineName after no of rows insert Method
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, ";")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormula1 = nRows
'----------------Lock Password-------------------START---
   ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

    Application.EnableEvents = True
End Function

'For assigning names to all the extra rows added - Yogesh Common for grids

Function onlyDigits(ByVal s As String, ExtracType As String) As String
'On Error Resume Next
    Dim retval As String    ' This is the return string.
    Dim i As Long        ' Counter for character position.

    retval = ""
    For i = 1 To Len(s)
        If (Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9") Then
         If ExtracType = "I" Then
            retval = retval + Mid(s, i, 1)
         End If
        Else
         If ExtracType = "S" Then
          retval = retval + Mid(s, i, 1)
         End If
        End If
    Next
    onlyDigits = retval
End Function



Function addblock(prevrangename As Variant, blockname As Variant, frmcounter As Variant, totalblocksize As Variant)
On Error GoTo endline
    Dim i, commindex, j, k, b, rangeCount, offset1, offset2 As Long
    Dim newnamerefersto As Variant
    Dim gapbtn, counter, dcounter, lenn, startrow, endrow, endrow1, endrow2 As Long
    Dim destinationrowindex, destinationcolumnindex, lastdestinationrow As String
    Dim rangearr, newranges, precrange, oldrangeaddress, newrangeaddress, newrangeaddress1, newrangeaddress2 As Variant
    Dim ndestinationrowindex, ndestinationcolumnindex As Long
    Dim noofrowsadded As Long
    
    Application.EnableEvents = False

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    gapbtn = 0
    dcounter = 0
    counter = Range(frmcounter).value
    If counter > 0 Then
    
        noofrowsadded = (totalblocksize - totalblocksize) / counter
        offset1 = totalblocksize + (1) * noofrowsadded
        offset2 = totalblocksize + (counter) * noofrowsadded
        ndestinationrowindex = Range(blockname & counter).row
        ndestinationrowindex = ndestinationrowindex + gapbtn + offset1
        ndestinationcolumnindex = Range(blockname & counter).Column
        newnamerefersto = Range(blockname & counter).Address
        lastdestinationrow = ndestinationrowindex + offset1 - 1
        
        Cells(CLng(ndestinationrowindex - 1), CLng(ndestinationcolumnindex)).Select
        
        Call InsertBlock(offset1)
        
        For i = 1 To Len(newnamerefersto)
            If (Mid(newnamerefersto, i, 1) = "$") Then
                 dcounter = dcounter + 1
    
                 If dcounter = 2 Then
                     lenn = (InStr(1, newnamerefersto, ":") - 1) - i
                     startrow = Mid(newnamerefersto, i + 1, lenn)
                     newnamerefersto = Replace(newnamerefersto, startrow, ndestinationrowindex)
                 End If
    
                 If dcounter = 4 Then
                     endrow = Mid(newnamerefersto, i + 1, Len(newnamerefersto) - i)
                     newnamerefersto = Replace(newnamerefersto, endrow, lastdestinationrow)
                 End If
    
            End If
        Next
            
        Application.EnableEvents = False
        
        counter = counter + 1
        ThisWorkbook.Names.add name:=blockname & counter, _
                RefersTo:="=" & newnamerefersto, Visible:=True
    
    
    
       
'----------------Unlock Password-------------------START---
   'sPassword = EfilingCommon.getmsgstate
   'ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
'        Range(frmcounter).Locked = False
        
        Range(frmcounter).value = counter
    
 '       Range(frmcounter).Locked = True
        
    Application.EnableEvents = False

'----------------Unlock Password-------------------START---
   'sPassword = EfilingCommon.getmsgstate
   'ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
        Range(blockname & counter - 1).Copy Destination:=Cells(ndestinationrowindex, ndestinationcolumnindex)
    End If


    commindex = 2
    j = 0
    
    prevrangename = Split(prevrangename, "||")
    rangeCount = UBound(prevrangename)
    
    ReDim rangearr(rangeCount)
    commindex = 0
    j = 1
    
    For i = 0 To UBound(prevrangename)
        rangearr(i) = prevrangename(i)
    Next
    
    ReDim newranges(rangeCount)
    For i = 0 To UBound(rangearr)
       newranges(i) = Replace(rangearr(i), CStr(1), CStr(counter))
    Next
    
    If counter > 2 Then
        ReDim precrange(rangeCount)
        For i = 0 To UBound(rangearr)
             precrange(i) = Replace(rangearr(i), CStr(1), CStr(counter - 1))
        Next
    
        ReDim oldrangeaddress(rangeCount)
        For i = 0 To UBound(precrange)
            If Not (precrange(i) = "") Then
               oldrangeaddress(i) = Range(precrange(i)).Address
            End If
        Next
    Else
        ReDim oldrangeaddress(rangeCount)
        For i = 0 To UBound(rangearr)
            If Not (rangearr(i) = "") Then
                oldrangeaddress(i) = Range(rangearr(i)).Address
            End If
        Next
    End If
    
    ReDim newrangeaddress(rangeCount)
    
    For i = 0 To UBound(oldrangeaddress)
        dcounter = 0
        Dim myrangearr() As String
        myrangearr = Split(oldrangeaddress(i), ":")
        If UBound(myrangearr) <= 0 Then
        For k = 1 To Len(oldrangeaddress(i))
            If (Mid(oldrangeaddress(i), k, 1) = "$") Then
                dcounter = dcounter + 1
                If (dcounter = 2) Then
                endrow = Mid(oldrangeaddress(i), k + 1, Len(oldrangeaddress(i)) - k)
                newrangeaddress(i) = Replace(oldrangeaddress(i), endrow, (endrow + gapbtn + offset1))
                End If
            End If
        Next
        
        Else
        
        Dim myoldrangeaddress() As String
        myoldrangeaddress = Split(oldrangeaddress(i), ":")
        
        For k = 1 To Len(myrangearr(0))
            If (Mid(myoldrangeaddress(0), k, 1) = "$") Then
                dcounter = dcounter + 1
                If (dcounter = 2) Then
                endrow1 = Mid(myoldrangeaddress(0), k + 1, Len(myoldrangeaddress(0)) - k)
                newrangeaddress1 = Replace(myoldrangeaddress(0), endrow1, (endrow1 + gapbtn + offset1))
                End If
            End If
        Next
        dcounter = 0
        For k = 1 To Len(myrangearr(1))
            If (Mid(myoldrangeaddress(1), k, 1) = "$") Then
                dcounter = dcounter + 1
                If (dcounter = 2) Then
                endrow2 = Mid(myoldrangeaddress(1), k + 1, Len(myoldrangeaddress(1)) - k)
                newrangeaddress2 = Replace(myoldrangeaddress(1), endrow2, (endrow2 + gapbtn + offset1))
                End If
            End If
    
        Next
            newrangeaddress(i) = newrangeaddress1 + ":" + newrangeaddress2
        End If
        
    
      Next
    
    
    'On Error Resume Next
    For b = 0 To UBound(newrangeaddress)
    
        If Not newranges(b) = "" Then
            ThisWorkbook.Names.add name:=newranges(b), _
                        RefersTo:="=" & newrangeaddress(b), Visible:=True
            If Range(newranges(b)).Interior.ColorIndex = 35 Then
                Range(newranges(b)).ClearContents
            End If
        End If
    Next
    
endline:

'----------------Lock Password-------------------START---
   ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
Application.EnableEvents = True
End Function

Function InsertBlock(vRows1 As Variant)
On Error GoTo endline
    Dim x As Long
    Dim vRows As Long
    Dim sht As Worksheet, shts() As String, i As Long
    Application.EnableEvents = False
    ActiveCell.EntireRow.Select

    vRows = CLng(vRows1)
    
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    ReDim shts(1 To Worksheets.Application.ActiveWorkbook. _
       Windows(1).SelectedSheets.count)
    i = 0
    For Each sht In _
       Application.ActiveWorkbook.Windows(1).SelectedSheets
    Sheets(sht.name).Select
    i = i + 1
    shts(i) = sht.name

    x = Sheets(sht.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow. _
     Resize(rowsize:=vRows).Insert Shift:=xlDown

    Selection.AutoFill Selection.Resize( _
     rowsize:=vRows + 1), xlFillDefault

    On Error Resume Next
    
    Selection.Offset(1).Resize(vRows).EntireRow. _
     SpecialCells(xlConstants).ClearContents
   Next sht
   
   Worksheets(shts).Select
endline:
'----------------Lock Password-------------------START---
   'ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
Application.EnableEvents = True
   
End Function

Function insertRowUnderSection(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
'Method : Insert no of Rows
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
    
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    'Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlCellTypeAllValidation).ClearContents
    
'Method: DefineName after no of rows insert Method
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSection = nRows
'----------------Lock Password-------------------START---
   ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
    Application.EnableEvents = True
End Function


Function checkfieldspecialcharacter_FSI(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter_FSI = True
    Dim arr As Variant
    'arr = Array("@", "*", "!", "-", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<")
    arr = Array(">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter_FSI = False
            Exit Function
        End If
        Next
    Next
End Function

Function checkfieldspecialcharacter(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter = True
    Dim arr As Variant
    'arr = Array("@", "*", "!", "-", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<")
    arr = Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        'For j = 0 To UBound(arr)
        If Mid(field, i, 1) = "&" Or Mid(field, i, 1) = """" Or Mid(field, i, 1) = "'" Or Mid(field, i, 1) = ">" Or Mid(field, i, 1) = "<" Then
            checkfieldspecialcharacter = False
            Exit Function
        End If
        'Next
    Next
End Function

Function checkallfieldspecialcharacter(field As Variant) As Boolean
    Dim i, j As Long
    checkallfieldspecialcharacter = True
    Dim arr As Variant
    arr = Array("@", "*", "!", "-", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<", "&", """", "'", ">", "<")
    'arr = Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkallfieldspecialcharacter = False
            Exit Function
        End If
        Next
    Next
End Function
Function checkallfieldspecialcharacter_80PRAN(field As Variant) As Boolean
    Dim i, j As Long
    checkallfieldspecialcharacter_80PRAN = True
    Dim arr As Variant
    arr = Array("@", "*", "!", "-", "&", "#", "~", ".", ",", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<", "&", """", "'", ">", "<")
    'arr = Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkallfieldspecialcharacter_80PRAN = False
            Exit Function
        End If
        Next
    Next
End Function

Function checkfieldspecialcharacter80G(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter80G = True
    Dim arr As Variant
    arr = Array("@", "|", "_", "*", "!", "&", "#", "~", ";", "?", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter80G = False
            Exit Function
        End If
        Next
    Next
End Function

Function checkallfieldspecialcharacter_80U_80DD(field As Variant) As Boolean
    Dim i, j As Long
    checkallfieldspecialcharacter_80U_80DD = True
    Dim arr As Variant
    arr = Array("@", "*", "!", "-", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<", "&", """", "'", ">", "<")
    'arr = Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkallfieldspecialcharacter_80U_80DD = False
            Exit Function
        End If
        Next
    Next
End Function
'Function isdropdownblank(dropdown As Variant) As Boolean
'    isdropdownblank = False
'    If Mid(dropdown, 1, 1) = "(" Then
'        isdropdownblank = True
'    End If
'    If Mid(dropdown, 1, 1) = "" Or IsEmpty(dropdown) Then
'        isdropdownblank = True
'    End If
'End Function

Function InsertRowsAndFillFormulas80G(Optional vRows As Long = 0) As Long
    Dim x As Long
    Dim sht As Worksheet, shts() As String, i As Long
    
    sPassword = EfilingCommon.getmsgstate
    ActiveSheet.Unprotect Password:=sPassword
    
    ActiveCell.EntireRow.Select
    If vRows = 0 Then
        vRows = Application.InputBox(prompt:= _
        "Enter the number of rows you want to add below selected cell", Title:="Add Rows below the selected cell", _
        Default:=1, Type:=1) 'Default for 1 row, type 1 is number
        
        If vRows = False Then
            InsertRowsAndFillFormulas80G = 0
            Exit Function
        End If
    End If
    
    
    ReDim shts(1 To Worksheets.Application.ActiveWorkbook. _
    Windows(1).SelectedSheets.count)
    
    i = 0
    For Each sht In _
        Application.ActiveWorkbook.Windows(1).SelectedSheets
        Sheets(sht.name).Select
        i = i + 1
        shts(i) = sht.name
        x = Sheets(sht.name).UsedRange.Rows.count 'lastcell fixup
        
        Selection.Resize(rowsize:=2).Rows(2).EntireRow. _
        Resize(rowsize:=vRows).Insert Shift:=xlDown
        
        Selection.AutoFill Selection.Resize( _
        rowsize:=vRows + 1), xlFillDefault
        
        On Error Resume Next
        Selection.Offset(1).Resize(vRows).EntireRow. _
        SpecialCells(xlConstants).ClearContents
    Next sht
    Worksheets(shts).Select
    InsertRowsAndFillFormulas80G = vRows
    
    ActiveSheet.Protect Password:=sPassword
End Function



Function CountDropdown(ByVal sDropdownValue As String, rangecells As Range, rangeCount As Long) As Long
    Dim currCnt As Long
    Dim tcountry_Fsi As Variant
    Dim rangeValue As Variant
    Dim i As Long
    
    currCnt = 0
    
    ReDim rangeValue(rangeCount)
      For i = 1 To rangeCount
        If Not rangecells.item(i).value = "(Select)" Then
            rangeValue(i) = rangecells.item(i).value
        End If
      Next
    
    If Not IsEmpty(rangeValue) And UBound(rangeValue) > 0 Then
      For i = 1 To UBound(rangeValue)
       If rangeValue(i) = sDropdownValue And (IIf(rangeValue(i) = "", "(Select)", rangeValue(i)) <> "(Select)") Then
        currCnt = currCnt + 1
       End If
      Next
    End If
    CountDropdown = currCnt
End Function
Function CountDropdown_EI(ByVal sDropdownValue As String, rangecells As Range, rangeCount As Long) As Long
    Dim currCnt As Long
    Dim tcountry_Fsi As Variant
    Dim rangeValue As Variant
    Dim i As Long
    
    currCnt = 0
    
    ReDim rangeValue(rangeCount)
      For i = 1 To rangeCount
        If Not rangecells.item(i).value = "(Select)" Then
            rangeValue(i) = rangecells.item(i).value
        End If
      Next
    
    If Not IsEmpty(rangeValue) And UBound(rangeValue) > 0 Then
      For i = 1 To UBound(rangeValue)
       If rangeValue(i) = sDropdownValue And (IIf(rangeValue(i) = "", "(Select)", rangeValue(i)) <> "(Select)") And (IIf(rangeValue(i) = "", "Any other", rangeValue(i)) <> "Any other") Then
        currCnt = currCnt + 1
       End If
      Next
    End If
    CountDropdown_EI = currCnt
End Function

Function CountDropdown1(ByVal sDropdownValue As String, rangecells As Range, rangeCount As Long) As Long
    Dim currCnt  As Long
    Dim rangeValue As Variant
    Dim i As Long

    currCnt1 = 0
     currCnt2 = 0
      currCnt3 = 0
      currCnt4 = 0
      currCnt5 = 0
      currCnt6 = 0
      currCnt7 = 0
      currCnt8 = 0
      currCnt9 = 0
      currCnt10 = 0
      currCnt11 = 0
      currCnt12 = 0
      currCnt13 = 0
      currCnt14 = 0
      currCnt15 = 0
      currCnt16 = 0
      currCnt17 = 0


    ReDim rangeValue(rangeCount)
      For i = 1 To rangeCount
        If Not rangecells.item(i).value = "(Select)" Then
            rangeValue(i) = rangecells.item(i).value
        End If
      Next

    If Not IsEmpty(rangeValue) And UBound(rangeValue) > 0 Then
      For i = 1 To UBound(rangeValue)
       If Trim(rangeValue(i)) = Trim(sDropdownValue) And (IIf(Trim(rangeValue(i)) = "", "(Select)", Trim(rangeValue(i))) <> "(Select)") Then

       'If rangeValue(i) = sDropdownValue And (IIf(rangeValue(i) = "", "(Select)", rangeValue(i)) <> "(Select)") Then
       If Trim(sDropdownValue) = "Basic Salary" Then
       currCnt1 = currCnt1 + 1
       End If

       If Trim(sDropdownValue) = "Dearness Allowance" Then
       currCnt2 = currCnt2 + 1
       End If

       If Trim(sDropdownValue) = "Conveyance Allowance" Then
       currCnt3 = currCnt3 + 1
       End If

       If Trim(sDropdownValue) = "House Rent Allowance" Then
       currCnt4 = currCnt4 + 1
       End If

       If Trim(sDropdownValue) = "Leave Travel Allowance" Then
       currCnt5 = currCnt5 + 1
       End If

       If Trim(sDropdownValue) = "Children Education Allowance" Then
       currCnt6 = currCnt6 + 1
       End If

       If Trim(sDropdownValue) = "Other Allowance" Then
       currCnt7 = currCnt7 + 1
       End If
       
       If Trim(sDropdownValue) = "The contribution made  by the Employer  towards  pension scheme as referred u/s 80CCD" Then
       currCnt8 = currCnt8 + 1
       End If
       
       
'       If Trim(sDropdownValue) = "Amount deemed to be income under rule 11 of Fourth Schedule" Then
'       currCnt9 = currCnt9 + 1
'       End If
       
       If Trim(sDropdownValue) = "Amount deemed to be income under rule 11(4) of Part-A of Fourth Schedule" Then
       currCnt9 = currCnt9 + 1
       End If
       
       'Newly changed as per DE sheet by Bindu
'       If Trim(sDropdownValue) = "Amount deemed to be income under rule 6 of Fourth Schedule" Then
'       currCnt10 = currCnt10 + 1
'       End If
       
       If Trim(sDropdownValue) = "Amount deemed to be income under rule 6 of Part-A of Fourth Schedule" Then
       currCnt10 = currCnt10 + 1
       End If
       
       
       If Trim(sDropdownValue) = "Annuity or pension" Then
       currCnt11 = currCnt11 + 1
       End If
       
       
       
       If Trim(sDropdownValue) = "Commuted Pension" Then
       currCnt12 = currCnt12 + 1
       End If
       
       If Trim(sDropdownValue) = "Gratuity" Then
       currCnt13 = currCnt13 + 1
       End If
       
       If Trim(sDropdownValue) = "Fees/ commission" Then
       currCnt14 = currCnt14 + 1
       End If
       
       If Trim(sDropdownValue) = "Advance of salary" Then
       currCnt15 = currCnt15 + 1
       End If
       
       If Trim(sDropdownValue) = "Leave Encashment" Then
       currCnt16 = currCnt16 + 1
       End If
       
       If Trim(sDropdownValue) = "Contribution made by the central government towards Agnipath scheme as referred  under section 80CCH" Then
       currCnt17 = currCnt17 + 1
       End If


       End If
      Next
    End If
    CountDropdown1 = currCnt1
End Function

Sub LOCK80GG()
On Error Resume Next
 Application.ScreenUpdating = False
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet19.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Sheet19.Range("scvia.Section80GG").value = ""
Sheet19.Range("scvia.Section80GG").Locked = True
Sheet19.Range("scvia.Section80GG").Interior.Color = (&HD8D8D8)


'sudha_2025-26

Sheet19.Range("Sheet19.AckNum").value = ""
Sheet19.Range("Sheet19.AckNum").Locked = True
Sheet19.Range("Sheet19.AckNum").Interior.Color = (&HD8D8D8)

'----------------Lock Password-------------------START---
   Sheet19.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

 Application.ScreenUpdating = True
 End Sub
Sub UNLOCK80GG()
On Error Resume Next
Application.ScreenUpdating = False
 'Application.EnableEvents = False
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet19.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Sheet19.Range("scvia.Section80GG").value = ""
Sheet19.Range("scvia.Section80GG").Locked = False
Sheet19.Range("scvia.Section80GG").Interior.Color = (&HCCFFCC)

'sudha_2025-26

Sheet19.Range("Sheet19.AckNum").value = ""
Sheet19.Range("Sheet19.AckNum").Locked = False
Sheet19.Range("Sheet19.AckNum").Interior.Color = (&HCCFFCC)

'----------------Lock Password-------------------START---
   Sheet19.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

 'Application.EnableEvents = True
 Application.ScreenUpdating = True
End Sub
Function CountDropdownAlwnc(ByVal sDropdownValue As String, rangecells As Range, rangeCount As Long) As Long
    Dim currCnt  As Long
    Dim rangeValue As Variant
    Dim i As Long

    currCntAlwnc1 = 0
     currCntAlwnc2 = 0
      currCntAlwnc3 = 0
      currCntAlwnc4 = 0
      currCntAlwnc5 = 0
      currCntAlwnc6 = 0
      currCntAlwnc7 = 0
      currCntAlwnc8 = 0
      currCntAlwnc9 = 0
      currCntAlwnc10 = 0
      currCntAlwnc11 = 0
      currCntAlwnc12 = 0
      currCntAlwnc13 = 0
      currCntAlwnc14 = 0
      'added by Chetan C M on 28/08/2025
      'start--
      currCntAlwnc15 = 0
      currCntAlwnc16 = 0
      '--end
      
    ReDim rangeValue(rangeCount)
      For i = 1 To rangeCount
        If Not rangecells.item(i).value = "(Select)" Then
            rangeValue(i) = rangecells.item(i).value
        End If
      Next
      

    If Not IsEmpty(rangeValue) And UBound(rangeValue) > 0 Then
      For i = 1 To UBound(rangeValue)
       If Trim(rangeValue(i)) = Trim(sDropdownValue) And (IIf(Trim(rangeValue(i)) = "", "(Select)", Trim(rangeValue(i))) <> "(Select)") Then

       'If rangeValue(i) = sDropdownValue And (IIf(rangeValue(i) = "", "(Select)", rangeValue(i)) <> "(Select)") Then
       If Trim(sDropdownValue) = "Sec 10(5)-Leave Travel concession/assistance" Then
       currCntAlwnc1 = currCntAlwnc1 + 1
       End If

       If Trim(sDropdownValue) = "Sec 10(6)-Remuneration received as an official, by whatever name called, of an embassy, high commission etc." Then
       currCntAlwnc2 = currCntAlwnc2 + 1
       End If

       If Trim(sDropdownValue) = "Sec 10(7)-Allowances or perquisites paid or allowed as such outside India by the Government to a citizen of India for rendering service outside India" Then
       currCntAlwnc3 = currCntAlwnc3 + 1
       End If
        
    'changed by Chetan C M on 28/08/2025
    'start--
        'If Trim(sDropdownValue) = "Sec 10(10)-Death-cum-retirement gratuity received " Then
        If Trim(sDropdownValue) = "Sec 10(10)-Death-cum-retirement gratuity received" Then
        currCntAlwnc4 = currCntAlwnc4 + 1
        End If
    '--end

       If Trim(sDropdownValue) = "Sec 10(10A)-Commuted value of pension received" Then
       currCntAlwnc5 = currCntAlwnc5 + 1
       End If

       If Trim(sDropdownValue) = "Sec 10(10AA)-Earned leave encashment on Retirement" Then
       currCntAlwnc6 = currCntAlwnc6 + 1
       End If

       If Trim(sDropdownValue) = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
       currCntAlwnc7 = currCntAlwnc7 + 1
       End If
       
       If Trim(sDropdownValue) = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
       currCntAlwnc8 = currCntAlwnc8 + 1
       End If
       
       If Trim(sDropdownValue) = "Sec 10(10C)-Amount received/receivable on voluntary retirement or termination of service" Then
       currCntAlwnc9 = currCntAlwnc9 + 1
       End If
       
       If Trim(sDropdownValue) = "Sec 10(10CC)-Tax paid by employer on non-monetary perquisite" Then
       currCntAlwnc10 = currCntAlwnc10 + 1
       End If
       
       If Trim(sDropdownValue) = "Sec 10(13A)-Allowance to meet expenditure incurred on house rent" Then
       currCntAlwnc11 = currCntAlwnc11 + 1
       End If
       
    'changed by Chetan C M on 24/10/2025 for SIT-89185-(issue9)
    'start--
'       If Trim(sDropdownValue) = "Sec 10(14)(i)-Prescribed Allowances or benefits (not in a nature of perquisite) specifically granted to meet expenses wholly, necessarily and exclusively and to the extent actually incurred, in performance of duties of office or employment" Then
        If Trim(sDropdownValue) = "Sec 10(14)(i)-Allowances referred in sub-clauses (a) to (c) of sub-rule (1) in Rule 2BB" Then
       currCntAlwnc12 = currCntAlwnc12 + 1
       End If
    '--end
    
    'changed by Chetan C Mon 28/08/2025
    'start--
       'If Trim(sDropdownValue) = "Sec 10(14)(ii)-Prescribed Allowances or benefits granted to meet personal expenses in performance of duties of office or employment or to compensate him for increased cost of living. " Then
        If Trim(sDropdownValue) = "Sec 10(14)(ii)-Prescribed Allowances or benefits granted to meet personal expenses in performance of duties of office or employment or to compensate him for increased cost of living." Then
        currCntAlwnc13 = currCntAlwnc13 + 1
        End If
        
        If Trim(sDropdownValue) = "Sec 10(14)(ii)-Transport allowance granted to certain physically handicapped assessee" Then
        currCntAlwnc14 = currCntAlwnc14 + 1
        End If
        
        If Trim(sDropdownValue) = "Exempt income received by a judge covered under the payment of salaries to Supreme Court/High Court judges Act /Rules" Then
        currCntAlwnc15 = currCntAlwnc15 + 1
        End If
        
    '--end
       
       If Trim(sDropdownValue) = "Any Other" Then
       currCntAlwnc16 = currCntAlwnc16 + 1
       End If
       
       
       End If
      Next
    End If
    CountDropdownAlwnc = rangeCount
End Function


Function CountDropdown2(ByVal sDropdownValue As String, rangecells As Range, rangeCount As Long) As Long
    Dim currCnt  As Long
    Dim rangeValue As Variant
    Dim i As Long

    currCntvalue1 = 0
     currCntvalue2 = 0
      currCntvalue3 = 0
      currCntvalue4 = 0
      currCntvalue5 = 0
      currCntvalue6 = 0
      currCntvalue7 = 0
      currCntvalue8 = 0
      currCntvalue9 = 0
      currCntvalue10 = 0
      currCntvalue11 = 0
      currCntvalue12 = 0
      currCntvalue13 = 0
      currCntvalue14 = 0
      currCntvalue15 = 0
      currCntvalue16 = 0
      currCntvalue17 = 0
      currCntvalue18 = 0
      currCntvalue19 = 0
      currCntvalue20 = 0
      currCntvalue21 = 0
      


    ReDim rangeValue(rangeCount)
      For i = 1 To rangeCount
        If Not rangecells.item(i).value = "(Select)" Then
            rangeValue(i) = rangecells.item(i).value
        End If
      Next

    If Not IsEmpty(rangeValue) And UBound(rangeValue) > 0 Then
      For i = 1 To UBound(rangeValue)
       If Trim(rangeValue(i)) = Trim(sDropdownValue) And (IIf(Trim(rangeValue(i)) = "", "(Select)", Trim(rangeValue(i))) <> "(Select)") Then

       'If rangeValue(i) = sDropdownValue And (IIf(rangeValue(i) = "", "(Select)", rangeValue(i)) <> "(Select)") Then
       If Trim(sDropdownValue) = "Accommodation" Then
       currCntvalue1 = currCntvalue1 + 1
       End If

       If Trim(sDropdownValue) = "Cars / Other Automotive" Then
       currCntvalue2 = currCntvalue2 + 1
       End If

       If Trim(sDropdownValue) = "Sweeper, gardener, watchman or personal attendant" Then
       currCntvalue3 = currCntvalue3 + 1
       End If

       If Trim(sDropdownValue) = "Gas, electricity, water" Then
       currCntvalue4 = currCntvalue4 + 1
       End If

       If Trim(sDropdownValue) = "Interest free or concessional loans" Then
       currCntvalue5 = currCntvalue5 + 1
       End If

       If Trim(sDropdownValue) = "Holiday expenses" Then
       currCntvalue6 = currCntvalue6 + 1
       End If

       If Trim(sDropdownValue) = "Free or concessional travel" Then
       currCntvalue7 = currCntvalue7 + 1
       End If
       
       If Trim(sDropdownValue) = "Free meals" Then
       currCntvalue8 = currCntvalue8 + 1
       End If
       
       If Trim(sDropdownValue) = "Free education" Then
       currCntvalue9 = currCntvalue9 + 1
       End If
       If Trim(sDropdownValue) = "Gifts, vouchers, etc." Then
       currCntvalue10 = currCntvalue10 + 1
       End If
       
       If Trim(sDropdownValue) = "Credit card expenses" Then
       currCntvalue11 = currCntvalue11 + 1
       End If
       
       If Trim(sDropdownValue) = "Club expenses" Then
       currCntvalue12 = currCntvalue12 + 1
       End If
       
       If Trim(sDropdownValue) = "Use of movable assets by employees" Then
       currCntvalue13 = currCntvalue13 + 1
       End If
       
       If Trim(sDropdownValue) = "Transfer of assets to employee" Then
       currCntvalue14 = currCntvalue14 + 1
       End If
       
       If Trim(sDropdownValue) = "Value of any other benefit/amenity/service/privilege" Then
       currCntvalue15 = currCntvalue15 + 1
       End If
       
'changes done by Chetan C M on 12/08/2025
    'start--
       
'        If Trim(sDropdownValue) = "Stock options (non-qualified options)" Then
'       currCntvalue16 = currCntvalue16 + 1
'       End If
'
'       If Trim(sDropdownValue) = "Tax paid by employer on non-monetary perquisite" Then
'       currCntvalue17 = currCntvalue17 + 1
'       End If
'
'       If Trim(sDropdownValue) = "Stock options allotted or transferred by employer being an eligible start-up referred to in section 80-IAC" Then
'       currCntvalue18 = currCntvalue18 + 1
'       End If

         If Trim(sDropdownValue) = "Stock options allotted or transferred by employer being an eligible start-up referred to in section 80-IAC-Tax to be deferred" Then
       currCntvalue16 = currCntvalue16 + 1
       End If
       
       If Trim(sDropdownValue) = "Other benefits or amenities" Then
       currCntvalue17 = currCntvalue17 + 1
       End If
       
       If Trim(sDropdownValue) = "Stock options allotted or transferred by employer being an eligible start-up referred to in section 80-IAC-Tax not to be deferred" Then
       currCntvalue18 = currCntvalue18 + 1
       End If
       
    '--end

       If Trim(sDropdownValue) = "Stock options (non-qualified options) other than ESOP in col 16 above." Then
       currCntvalue19 = currCntvalue19 + 1
       End If

       If Trim(sDropdownValue) = "Contribution by employer to fund and scheme taxable under section 17(2)(vii)" Then
       currCntvalue20 = currCntvalue20 + 1
       End If

       If Trim(sDropdownValue) = "Annual accretion by way of interest, dividend etc. to the balance at the credit of fund and scheme referred to in section 17(2)(vii) and taxable under section 17(2)(viia)" Then
       currCntvalue21 = currCntvalue21 + 1
       End If
       
       End If
      Next
    End If
    CountDropdown2 = currCnt1
End Function
Function CountDropdown3(ByVal sDropdownValue As String, rangecells As Range, rangeCount As Long) As Long
    Dim currCnt  As Long
    Dim rangeValue As Variant
    Dim i As Long

    currCntProfit1 = 0
     currCntProfit2 = 0
      currCntProfit3 = 0
      
    ReDim rangeValue(rangeCount)
      For i = 1 To rangeCount
        If Not rangecells.item(i).value = "(Select)" Then
            rangeValue(i) = rangecells.item(i).value
        End If
      Next

    If Not IsEmpty(rangeValue) And UBound(rangeValue) > 0 Then
      For i = 1 To UBound(rangeValue)
       If Trim(rangeValue(i)) = Trim(sDropdownValue) And (IIf(Trim(rangeValue(i)) = "", "(Select)", Trim(rangeValue(i))) <> "(Select)") Then

       'If rangeValue(i) = sDropdownValue And (IIf(rangeValue(i) = "", "(Select)", rangeValue(i)) <> "(Select)") Then
'       If Trim(sDropdownValue) = "Compensation due/received by an assessee from his employer or former employer in connection with the  termination of his employment or modification thereto" Then
        If Trim(sDropdownValue) = "Any compensation due or received by an assessee from an employer or former employer in connection with the termination of his employment or modification thereto." Then
       
       currCntProfit1 = currCntProfit1 + 1
       End If

       If Trim(sDropdownValue) = "Any payment due/received by an assessee from his employer or a former employer or from a provident or other fund, sum received under Keyman Insurance Policy including Bonus thereto" Then
       currCntProfit2 = currCntProfit2 + 1
       End If

       If Trim(sDropdownValue) = "Any amount due/received by assessee from any person before joining or after cessation of employment with that person" Then
       currCntProfit3 = currCntProfit3 + 1
       End If

       
      
       End If
      Next
    End If
    CountDropdown3 = currCnt1
End Function
'
'Function calculateAge(dob As Variant) As Long
'
'On Error Resume Next
'calculateAge = (2022) - val(Mid(dob, 7, 4))
'    If 4 < val(Mid(dob, 4, 2)) Then
'        calculateAge = calculateAge - 1
'    ElseIf val(Mid(dob, 4, 2)) = 4 And 1 < val(Mid(dob, 1, 2)) Then
'                calculateAge = calculateAge - 1
'    End If
'
'End Function

Function checkFirstDateBefore(firstDate As Variant, secondDate As Variant) As Boolean
On Error Resume Next


If val(Mid(firstDate, 7, 4)) < val(Mid(secondDate, 7, 4)) Then
        checkFirstDateBefore = True
ElseIf val(Mid(firstDate, 7, 4)) = val(Mid(secondDate, 7, 4)) Then

        If val(Mid(firstDate, 4, 2)) < val(Mid(secondDate, 4, 2)) Then
        
             checkFirstDateBefore = True
             
        ElseIf val(Mid(firstDate, 4, 2)) = val(Mid(secondDate, 4, 2)) Then
        
                If val(Mid(firstDate, 1, 2)) < val(Mid(secondDate, 1, 2)) Then
                
                        checkFirstDateBefore = True
                        
                ElseIf val(Mid(firstDate, 1, 2)) = val(Mid(secondDate, 1, 2)) Then
                
                        checkFirstDateBefore = True
            
                Else
                      checkFirstDateBefore = False
                End If
        Else
            checkFirstDateBefore = False
        End If
Else

    checkFirstDateBefore = False
End If
End Function


Sub UNLOCKFA()
 Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    Application.ScreenUpdating = False
  Application.EnableEvents = False

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   sourceSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Dim rngname_A1_FA As Variant
Dim rngname_A2_FA As Variant
Dim rngname_A3_FA As Variant
Dim rngname_A4_FA As Variant
Dim rngname_F_FA As Variant
Dim rngname_B_FA As Variant
Dim rngname_C_FA As Variant
Dim rngname_D_FA As Variant
Dim rngname_E_FA As Variant
Dim rngname_G_FA As Variant


'rngname_A_FA = "FA_A_Country||FA_A_ZipCode||FA_A_BankName||FA_A_BankAdd||FA_A_AccountName||FA_A_StatusBeneficiary||FA_A_ForeignAccountNumber||FA_A_IBAN||FA_A_AccOpeningDate||FA_A_PeakBal||FA_A_AccountInterest||FA_A_Amount||FA_A_ScheduleOffered||FA_A_ItemNumOfSchedule"
rngname_A1_FA = "FA_A1_Country||FA_A1_BankName||FA_A1_BankAdd||FA_A1_ZipCode||FA_A1_ForeignAccountNumber||FA_A1_StatusBeneficiary||FA_A1_AccOpeningDate||FA_A1_PeakBal||FA_A1_ClosingBal||FA_A1_Grossinterest"
rngname_A2_FA = "FA_A2_Country||FA_A2_BankName||FA_A2_BankAdd||FA_A2_ZipCode||FA_A2_ForeignAccountNumber||FA_A2_StatusBeneficiary||FA_A2_AccOpeningDate||FA_A2_PeakBal||FA_A2_ClosingBal||FA_A2_Grossinterest_Nature"
rngname_A3_FA = "FA_A3_Country||FA_A3_BankName||FA_A3_BankAdd||FA_A3_ZipCode||FA_A3_NatureOfEntity||FA_A3_AccOpeningDate||FA_A3_initialvalue||FA_A3_PeakBal||FA_A3_ClosingBal||FA_A3_Totalgrossamount||FA_A3_Totalgrosproceeds"
rngname_A4_FA = "FA_A4_Country||FA_A4_BankName||FA_A4_BankAdd||FA_A4_ZipCode||FA_A4_AccOpeningDate||FA_A4_the_cashvalue||FA_A4_Total_grossamount"
rngname_B_FA = "FA_B_Country||FA_B_ZipCode||FA_B_EntityNature||FA_B_EntityName||FA_B_EntityAdd||FA_B_NatureOfInterest||FA_B_DateSinceHeld||FA_B_TotalInv||FA_B_IncomeInterest||FA_B_NatureOfIncome||FA_B_Amount||FA_B_ScheduleOffered||FA_B_ItemNumOfSchedule"
rngname_C_FA = "FA_C_Country||FA_C_ZipCode||FA_C_PropertyAdd||FA_C_OwnershipBeneficiary||FA_C_DateOfAcquisition||FA_C_TotalInv||FA_C_IncomeFromProperty||FA_C_NatureOfIncome||FA_C_Amount||FA_C_ScheduleOffered||FA_C_ItemNumOfSchedule"
rngname_D_FA = "FA_D_Country||FA_D_ZipCode||FA_D_AssetNature||FA_D_OwnershipBeneficiary||FA_D_DateOfAcquisition||FA_D_TotalInv||FA_D_IncomeFromProperty||FA_D_NatureOfIncome||FA_D_Amount||FA_D_ScheduleOffered||FA_D_ItemNumOfSchedule"
rngname_E_FA = "FA_E_InstName||FA_E_InstAdd||FA_E_Country||FA_E_ZipCode||FA_E_AccountName||FA_E_InstitutionAccountNumber||FA_E_PeakBalInv||FA_E_IncomeAccured||FA_E_IncomeAccuredAmount||FA_E_Amount||FA_E_Scheduleoffered||FA_E_ItemNumberSchedule"

rngname_F_FA = "FA_F_Country||FA_F_ZipCode||FA_F_NameOfTrust||FA_F_AddressOfTrust||FA_F_NameOfOtherTrustees||FA_F_AddressOfOtherTrustees||FA_F_NameOfSettlor||FA_F_AddressOfSettlor||FA_F_NameOfBeneficiaries||FA_F_AddressOfBeneficiaries||FA_F_DateSinceHeld||FA_F_IncomeDerivedIsTaxable||FA_F_IncomeDerivedAmount||FA_F_Amount||FA_F_Scheduleoffered||FA_F_ItemNumberSchedule"
rngname_G_FA = "FA_G_Country||FA_G_ZipCode||FA_G_NameOfPerson||FA_G_AddressOfPerson||FA_G_IncomeDerived||FA_G_NatureOfIncome||FA_G_WhetherTaxable||FA_G_Amount||FA_G_Scheduleoffered||FA_G_ItemNumberSchedule"


UNLOCKRANGE (rngname_A1_FA)
UNLOCKRANGE (rngname_A2_FA)
UNLOCKRANGE (rngname_A3_FA)
UNLOCKRANGE (rngname_A4_FA)
UNLOCKRANGE (rngname_B_FA)
UNLOCKRANGE (rngname_C_FA)
UNLOCKRANGE (rngname_D_FA)
UNLOCKRANGE (rngname_E_FA)
UNLOCKRANGE (rngname_F_FA)
UNLOCKRANGE (rngname_G_FA)

'----------------Lock Password-------------------START---
   sourceSheet.Protect Password:=sPassword
   '----------------Lock Password-------------------END-----
Application.ScreenUpdating = True
 Application.EnableEvents = True
End Sub

Sub UNLOCKRANGE(rangenamestring As Variant)
Dim i As Long
    rangenamestring = Split(rangenamestring, "||")
    Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    For i = 0 To UBound(rangenamestring)
        sourceSheet.Range(rangenamestring(i)).Locked = False
       sourceSheet.Range(rangenamestring(i)).Interior.Color = (&HCCFFCC)
    'Added by Riyaz
        If rangenamestring(i) = "FA_D_OwnershipBeneficiary" Or rangenamestring(i) = "FA_C_OwnershipBeneficiary" Or rangenamestring(i) = "FA_B_NatureOfInterest" Or rangenamestring(i) = "FA_A1_Country" Or rangenamestring(i) = "FA_A1_StatusBeneficiary" Or rangenamestring(i) = "FA_A2_Country" Or rangenamestring(i) = "FA_A2_StatusBeneficiary" Or rangenamestring(i) = "FA_A3_Country" Or rangenamestring(i) = "FA_A4_Country" Or rangenamestring(i) = "FA_B_Country" Or rangenamestring(i) = "FA_C_Country" Or rangenamestring(i) = "FA_D_Country" Or rangenamestring(i) = "FA_E_InstName" Or rangenamestring(i) = "FA_F_Country" Or rangenamestring(i) = "FA_G_Country" Or rangenamestring(i) = "FA_B_ScheduleOffered" Or rangenamestring(i) = "FA_C_ScheduleOffered" Or rangenamestring(i) = "FA_D_ScheduleOffered" Or rangenamestring(i) = "FA_E_Scheduleoffered" Or rangenamestring(i) = "FA_F_Scheduleoffered" Or rangenamestring(i) = "FA_G_Scheduleoffered" Then
           Sheet27.Range(rangenamestring(i)).value = "(Select)"
        End If
    'End
    Next

End Sub
'Sub LOCKCG()
'
'Application.ScreenUpdating = False
' Application.EnableEvents = False
''----------------Unlock Password-------------------START---
'   sPassword = EfilingCommon.getmsgstate
'   Sheet13.Unprotect Password:=sPassword
''----------------Unlock Password-------------------END-----
'ThisWorkbook.Unprotect Password:=getmsgstate
'Dim rngname_CG As Variant
'
'rngname_CG = "STCG.A2iia_FullConsideration||STCG.A2iibi_AquisitCost||STCG.A2iibii_ImproveCost||STCG.A2iibiii_ExpOnTrans||STCG.A2iid_LossSec94of7Or94of8"
'LOCKRANGEcg (rngname_CG)
''Malli_AY_2025_26 V0.2
'rngname_CG = "STCG.A2iia_FullConsideration_ii||STCG.A2iibi_AquisitCost_ii||STCG.A2iibii_ImproveCost_ii||STCG.A2iibiii_ExpOnTrans_ii||STCG.A2iid_LossSec94of7Or94of8_ii"
'LOCKRANGEcg (rngname_CG)
''-----------------------
'
''Malli--AY_2024_25 OLD
''rngname_CG = "STCG.A3a_NRItaxSTTPaid||STCG.A3b_NRItaxSTTNotPaid||STCG.A4aia_FullConsideration||STCG.A4aib_FullMarketvalue||STCG.A4aii_FullConsiderationS||STCG.A4bi_AquisitCost||STCG.A4bii_ImproveCost||STCG.A4biii_ExpOnTrans||STCG.A4d_LossSec94of7Or94of8"
''Malli--CG_C51&52&53__AY_2025_26
'
''rngname_CG = "STCG.A3a_NRItaxSTTPaid||STCG.A3ai_NRItaxSTTPaid||STCG.A3aii_NRItaxSTTPaid||STCG.A3b_NRItaxSTTNotPaid||STCG.A4aia_FullConsideration||STCG.A4aib_FullMarketvalue||STCG.A4aii_FullConsiderationS||STCG.A4bi_AquisitCost||STCG.A4bii_ImproveCost||STCG.A4biii_ExpOnTrans||STCG.A4d_LossSec94of7Or94of8"
''LOCKRANGEcg (rngname_CG)
'
''Unit test change by sai on 22/02/2025=======================================================
'rngname_CG = "STCG.A3a_NRItaxSTTPaid||STCG.A3aii_NRItaxSTTPaid||STCG.A3b_NRItaxSTTNotPaid||STCG.A4aia_FullConsideration||STCG.A4aib_FullMarketvalue||STCG.A4aii_FullConsiderationS||STCG.A4bi_AquisitCost||STCG.A4bii_ImproveCost||STCG.A4biii_ExpOnTrans||STCG.A4d_LossSec94of7Or94of8"
''=============================================================================================
'
'
''------------------------------
'rngname_CG = "LTCG.B4a_LTCGWithoutBenefit||LTCG.B4bii_ExemptionAmount"
'LOCKRANGEcg (rngname_CG)
'rngname_CG = "LTCG.Bi5bi_AquisitCost||LTCG.Bi5bii_ImproveCost||LTCG.Bi5biii_ExpOnTrans||LTCG.Bi5dii_ExemptionAmount||LTCG.Bii5bi_AquisitCost||LTCG.Bii5bii_ImproveCost||LTCG.Bii5biii_ExpOnTrans||LTCG.Bii5dii_ExemptionAmount||LTCG.Biii5bi_AquisitCost||LTCG.Biii5bii_ImproveCost||LTCG.Biii5biii_ExpOnTrans||LTCG.Biii5dii_ExemptionAmount||LTCG.B6a_SaleonSpecAsset||LTCG.B6b_DednSpecAssetus115||LTCG.B6d_SaleOtherSpecAsset||LTCG.B6e_DednOtherSpecAssetus115||LTCG.B6aia_FullConsideration||LTCG.B6aib_FullMarketValue||LTCG.B6aii_FullConsiderationS||LTCG.B6ibia_FullConsideration||LTCG.B6ibib_FullMarketValue||LTCG.B6ibii_FullConsideration50CA||LTCG.B6iiiaia_FullConsideration||LTCG.B6iiiaib_FullMarketValue||LTCG.B6iiiaii_FullConsiderationS"
'LOCKRANGEcg (rngname_CG)
'rngname_CG = "LTCG.8.DeductionUnder54F||LTCG.B06d_ExemptionGrandTotal||LTCG.B07dii_ExemptionGrandTotal||LTCG.B07div_ExemptionGrandTotal||LTCG.B07dv_ExemptionGrandTotal"
'LOCKRANGEcg (rngname_CG)
'
'''Malli-----AY_2025_26
''rngname_CG = "STCG.A8_AmtI||STCG.A8_IA1tA6||STCG.A8_Country||STCG.A8_DTAA_Article||STCG.A8_RatePTreaty||STCG.A8_CertiStatus||STCG.A8_SecIT||STCG.A8_RateIT"
''LOCKRANGEcg (rngname_CG)
''LOCKRANGEcg_8tab
'''----------------------
'
'md115AD.ValidateLock_115AD_1
'
'Sheet13.Range("STCG.A.NRI").EntireRow.Hidden = True
'Sheet13.Range("LTCG.B.NRI").EntireRow.Hidden = True
''Sheet13.Shapes("Button 89").Visible = msoFalse
''Sheet13.Shapes("Button 160").Visible = msoFalse
''Sheet13.Range("STCG.DTAA").EntireRow.Hidden = True
''Sheet13.Range("LTCG.DTAA").EntireRow.Hidden = True
'
'
''Sheet13.Range("LTCG.B4d_LTCGOnListedSecurity").EntireRow.Hidden = True
'
''Sheet13.Range("LTCG.B4bi_ExemptionAmount").EntireRow.Hidden = True
''Sheet13.Range("LTCG.B4bii_ExemptionAmount").EntireRow.Hidden = True
''Sheet13.Range("LTCG.B4c_BalanceCG").EntireRow.Hidden = True
'
'
'
'
''----------------Lock Password-------------------START---
'   Sheet13.Protect Password:=sPassword
''----------------Lock Password-------------------END-----
'ThisWorkbook.Protect Password:=getmsgstate
' 'Application.EnableEvents = True
' Application.ScreenUpdating = True
'End Sub

Sub LOCKCG()
Application.ScreenUpdating = False
 Application.EnableEvents = False
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
ThisWorkbook.Unprotect Password:=getmsgstate
Dim rngname_CG As Variant

rngname_CG = "STCG.A2iia_FullConsideration||STCG.A2iibi_AquisitCost||STCG.A2iibii_ImproveCost||STCG.A2iibiii_ExpOnTrans||STCG.A2iid_LossSec94of7Or94of8"
LOCKRANGEcg (rngname_CG)
'Malli_AY_2025_26 V0.2
'Commented by Riyaz on 22/01/2026
'rngname_CG = "STCG.A2iia_FullConsideration_ii||STCG.A2iibi_AquisitCost_ii||STCG.A2iibii_ImproveCost_ii||STCG.A2iibiii_ExpOnTrans_ii||STCG.A2iid_LossSec94of7Or94of8_ii"
'LOCKRANGEcg (rngname_CG)
''-----------------------

'Malli--AY_2024_25 OLD
'rngname_CG = "STCG.A3a_NRItaxSTTPaid||STCG.A3b_NRItaxSTTNotPaid||STCG.A4aia_FullConsideration||STCG.A4aib_FullMarketvalue||STCG.A4aii_FullConsiderationS||STCG.A4bi_AquisitCost||STCG.A4bii_ImproveCost||STCG.A4biii_ExpOnTrans||STCG.A4d_LossSec94of7Or94of8"
'Malli--CG_C51&52&53__AY_2025_26
'rngname_CG = "STCG.A3a_NRItaxSTTPaid||STCG.A3ai_NRItaxSTTPaid||STCG.A3aii_NRItaxSTTPaid||STCG.A3b_NRItaxSTTNotPaid||STCG.A4aia_FullConsideration||STCG.A4aib_FullMarketvalue||STCG.A4aii_FullConsiderationS||STCG.A4bi_AquisitCost||STCG.A4bii_ImproveCost||STCG.A4biii_ExpOnTrans||STCG.A4d_LossSec94of7Or94of8"

'Unit test change by sai on 22/02/2025=======================================================
'Added one range name by Riyaz on 12/02/2026
rngname_CG = "STCG.A3a_NRItaxSTTPaid_i_ii||STCG.A3a_NRItaxSTTPaid||STCG.A3aii_NRItaxSTTPaid||STCG.A3b_NRItaxSTTNotPaid||STCG.A4aia_FullConsideration||STCG.A4aib_FullMarketvalue||STCG.A4aii_FullConsiderationS||STCG.A4bi_AquisitCost||STCG.A4bii_ImproveCost||STCG.A4biii_ExpOnTrans||STCG.A4d_LossSec94of7Or94of8"
'=============================================================================================


LOCKRANGEcg (rngname_CG)
'------------------------------
'Commented by Riyaz on 05/03/2025
'rngname_CG = "LTCG.B4a_LTCGWithoutBenefit||LTCG.B4bii_ExemptionAmount"
'LOCKRANGEcg (rngname_CG)
'rngname_CG = "LTCG.Bi5bi_AquisitCost||LTCG.Bi5bii_ImproveCost||LTCG.Bi5biii_ExpOnTrans||LTCG.Bi5dii_ExemptionAmount||LTCG.Bii5bi_AquisitCost||LTCG.Bii5bii_ImproveCost||LTCG.Bii5biii_ExpOnTrans||LTCG.Bii5dii_ExemptionAmount||LTCG.Biii5bi_AquisitCost||LTCG.Biii5bii_ImproveCost||LTCG.Biii5biii_ExpOnTrans||LTCG.Biii5dii_ExemptionAmount||LTCG.B6a_SaleonSpecAsset||LTCG.B6b_DednSpecAssetus115||LTCG.B6d_SaleOtherSpecAsset||LTCG.B6e_DednOtherSpecAssetus115||LTCG.B6aia_FullConsideration||LTCG.B6aib_FullMarketValue||LTCG.B6aii_FullConsiderationS||LTCG.B6ibia_FullConsideration||LTCG.B6ibib_FullMarketValue||LTCG.B6ibii_FullConsideration50CA||LTCG.B6iiiaia_FullConsideration||LTCG.B6iiiaib_FullMarketValue||LTCG.B6iiiaii_FullConsiderationS"
'LTCG.B6e_DednOtherSpecAssetus115 - REMOVED IN AY 25-26
'Commented by Riyaz on 05/03/2025
'rngname_CG = "LTCG.Bi5bi_AquisitCost||LTCG.Bi5bii_ImproveCost||LTCG.Bi5biii_ExpOnTrans||LTCG.Bi5dii_ExemptionAmount||LTCG.Bii5bi_AquisitCost||LTCG.Bii5bii_ImproveCost||LTCG.Bii5biii_ExpOnTrans||LTCG.Bii5dii_ExemptionAmount||LTCG.Biii5bi_AquisitCost||LTCG.Biii5bii_ImproveCost||LTCG.Biii5biii_ExpOnTrans||LTCG.Biii5dii_ExemptionAmount||LTCG.B6a_SaleonSpecAsset||LTCG.B6b_DednSpecAssetus115||LTCG.B6d_SaleOtherSpecAsset||LTCG.B6aia_FullConsideration||LTCG.B6aib_FullMarketValue||LTCG.B6aii_FullConsiderationS||LTCG.B6ibia_FullConsideration||LTCG.B6ibib_FullMarketValue||LTCG.B6ibii_FullConsideration50CA||LTCG.B6iiiaia_FullConsideration||LTCG.B6iiiaib_FullMarketValue||LTCG.B6iiiaii_FullConsiderationS"
'
'
'LOCKRANGEcg (rngname_CG)
'rngname_CG = "LTCG.8.DeductionUnder54F||LTCG.B06d_ExemptionGrandTotal||LTCG.B07dii_ExemptionGrandTotal||LTCG.B07div_ExemptionGrandTotal||LTCG.B07dv_ExemptionGrandTotal"
'LTCG.B07dv_ExemptionGrandTotal - REMOVED IN AY 25-26
'Commented by Riyaz on 05/03/2025
'rngname_CG = "LTCG.8.DeductionUnder54F||LTCG.B06d_ExemptionGrandTotal||LTCG.B07dii_ExemptionGrandTotal||LTCG.B07div_ExemptionGrandTotal"
'LOCKRANGEcg (rngname_CG)


''Malli-----AY_2025_26
'rngname_CG = "STCG.A8_AmtI||STCG.A8_IA1tA6||STCG.A8_Country||STCG.A8_DTAA_Article||STCG.A8_RatePTreaty||STCG.A8_CertiStatus||STCG.A8_SecIT||STCG.A8_RateIT"
'LOCKRANGEcg (rngname_CG)
'LOCKRANGEcg_8tab
''----------------------


'Newly added by Bindu
'B5 TABLE
'2 Range name added by Riyaz on 12/02/2026
Sheet13.Unprotect Password:=sPassword
rngname_CG = "LTCG.B4a_LTCGWithoutBenefit||LTCG.B06d_ExemptionGrandTotal||LTCG.B4ai_LTCGWithoutBenefit||LTCG.B4aii_LTCGWithoutBenefit||LTCG.B4aiii_LTCGWithoutBenefit||LTCG.B06d_i_ExemptionGrandTotal||LTCG.B06d_ii_ExemptionGrandTotal||LTCG.B06d_iii_ExemptionGrandTotal"
LOCKRANGEcg (rngname_CG)

'B6i Table_BE
'Commented by Riyaz on 23/01/2026
'Sheet13.Unprotect Password:=sPassword
'rngname_CG = "LTCG.B6aia_FullConsideration_ii||LTCG.B6aib_FullMarketValue_ii||LTCG.B6aii_FullConsiderationS_ii||LTCG.Bi5bi_AquisitCost_ii||LTCG.Bi5bii_ImproveCost_ii||LTCG.Bi5biii_ExpOnTrans_ii||LTCG.B07dii_ExemptionGrandTotal_ii"
'LOCKRANGEcg (rngname_CG)


'B6i Table_AE
Sheet13.Unprotect Password:=sPassword
rngname_CG = "LTCG.B6aia_FullConsideration||LTCG.B6aib_FullMarketValue||LTCG.B6aii_FullConsiderationS||LTCG.Bi5bi_AquisitCost||LTCG.Bi5bii_ImproveCost||LTCG.Bi5biii_ExpOnTrans||LTCG.B07dii_ExemptionGrandTotal"
LOCKRANGEcg (rngname_CG)


'NEWLY ADDED BY BINDU FOR TABLE B6ii_BE
'Commented by Riyaz on 23/01/2026
'Sheet13.Unprotect Password:=sPassword
'rngname_CG = "LTCG.B6ibia_FullConsideration_ii||LTCG.B6ibib_FullMarketValue_ii||LTCG.B6ibii_FullConsideration50CA_ii||LTCG.Bii5bi_AquisitCost_ii||LTCG.Bii5bii_ImproveCost_ii||LTCG.Bii5biii_ExpOnTrans_ii||LTCG.B07div_ExemptionGrandTotal_ii"
'LOCKRANGEcg (rngname_CG)

'NEWLY ADDED BY BINDU FOR TABLE B6ii_AE
rngname_CG = "LTCG.B6ibia_FullConsideration||LTCG.B6ibib_FullMarketValue||LTCG.B6ibii_FullConsideration50CA||LTCG.Bii5bi_AquisitCost||LTCG.Bii5bii_ImproveCost||LTCG.Bii5biii_ExpOnTrans||LTCG.B07div_ExemptionGrandTotal"
LOCKRANGEcg (rngname_CG)

'NEWLY ADDED BY BINDU FOR TABLE B6iii
Sheet13.Unprotect Password:=sPassword
rngname_CG = "LTCG.B6iiiaia_FullConsideration||LTCG.B6iiiaib_FullMarketValue||LTCG.B6iiiaii_FullConsiderationS||LTCG.Biii5bi_AquisitCost||LTCG.Biii5bii_ImproveCost||LTCG.Biii5biii_ExpOnTrans||LTCG.Biii5dii_ExemptionAmount"
LOCKRANGEcg (rngname_CG)

'NEWLY ADDED BY BINDU FOR TABLE B7
'1 Range name added by Riyaz on 12/02/2026
Sheet13.Unprotect Password:=sPassword
rngname_CG = "LTCG.8.DeductionUnder54F||LTCG.8_i.DeductionUnder54F||LTCG.8_ii.DeductionUnder54F"
LOCKRANGEcg (rngname_CG)

'NEWLY ADDED BY BINDU FOR TABLE B8
'2 Range name added by Riyaz on12 02 2026
Sheet13.Unprotect Password:=sPassword
rngname_CG = "LTCG.B6a_SaleonSpecAsset||LTCG.B6b_DednSpecAssetus115||LTCG.B6ai_SaleonSpecAsset||LTCG.B6aii_SaleonSpecAsset||LTCG.B6bi_DednSpecAssetus115||LTCG.B6bii_DednSpecAssetus115"
LOCKRANGEcg (rngname_CG)


md115AD.ValidateLock_115AD_1

Sheet13.Range("STCG.A.NRI").EntireRow.Hidden = True
Sheet13.Range("LTCG.B.NRI").EntireRow.Hidden = True

'Sheet13.Shapes("Button 89").Visible = msoFalse
'Sheet13.Shapes("Button 160").Visible = msoFalse
'Sheet13.Range("STCG.DTAA").EntireRow.Hidden = True
'Sheet13.Range("LTCG.DTAA").EntireRow.Hidden = True


'Sheet13.Range("LTCG.B4d_LTCGOnListedSecurity").EntireRow.Hidden = True

'Sheet13.Range("LTCG.B4bi_ExemptionAmount").EntireRow.Hidden = True
'Sheet13.Range("LTCG.B4bii_ExemptionAmount").EntireRow.Hidden = True
'Sheet13.Range("LTCG.B4c_BalanceCG").EntireRow.Hidden = True




'----------------Lock Password-------------------START---
   Sheet13.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
ThisWorkbook.Protect Password:=getmsgstate
 'Application.EnableEvents = True
 Application.ScreenUpdating = True

End Sub


Sub LOCKRANGEcg(rangenamestring As Variant)
    
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet13.Range(rangenamestring(i))
            chcell.value = ""
            chcell.Locked = True
            chcell.Interior.Color = (&HD8D8D8)

            Next
        Next
End Sub

'Sub LOCKRANGE115AD(rangenamestring As Variant)
'
'Dim i As Long
'Dim chcell As Range
'    rangenamestring = Split(rangenamestring, "||")
'        For i = 0 To UBound(rangenamestring)
'            For Each chcell In Sheet54.Range(rangenamestring(i))
'            chcell.value = ""
'            chcell.Locked = True
'            chcell.Interior.Color = (&HD8D8D8)
'
'            Next
'        Next
'End Sub
Function isdropdownblank(dropdown As Variant) As Boolean
    isdropdownblank = False
    
    If Mid(dropdown, 1, 1) = "(" Then
    ' MsgBox_80GGANew = MsgBox_80GGANew + "* Please select relevant clause at Sr. No." & i & " from Drop-Down in Schedule 80GGA" & Chr(13)
        isdropdownblank = True
    End If
    If Mid(dropdown, 1, 1) = "" Or IsEmpty(dropdown) Then
    ' MsgBox_80GGANew = MsgBox_80GGANew + "* Please select relevant clause at Sr. No." & i & " from Drop-Down in Schedule 80GGA" & Chr(13)
        isdropdownblank = True
    End If
End Function
Function chkCompulsory(field As Variant) As Boolean
    chkCompulsory = True
    If Len(Trim(field)) <= 0 Or IsEmpty(field) Then
        chkCompulsory = False
    End If
    
    'If IsEmpty(field) Then
    '    chkCompulsory = False
    'End If
End Function
Function checkfieldSuperSpecialcharacter(field As Variant) As Boolean
    checkfieldSuperSpecialcharacter = True
    Dim i, j  As Long
    Dim arr As Variant
        arr = Array(">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldSuperSpecialcharacter = False
            Exit Function
        End If
        Next
    Next
End Function

Sub LockOS()
Application.ScreenUpdating = False
  Application.EnableEvents = False
  Dim Rangename, name As Variant
  Dim i As Variant
  Rangename = "os.Country||os.Article_DTAA||os.Rate_DTAA||os.TRC||os.SectionAct||os.Amount_income"
  Rangename = Split(Rangename, "||")
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet14.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
For i = 0 To UBound(Rangename)
    Sheet14.Range(Rangename(i)).value = ""
    Sheet14.Range(Rangename(i)).Interior.Color = (&HD8D8D8)
    Sheet14.Range(Rangename(i)).Locked = True
Next
Sheet14.Range("OS_DTAA_hiding").EntireRow.Hidden = True
Sheet14.Shapes("Button 40").Visible = msoFalse

Sheet14.Protect Password:=sPassword
 
 Application.ScreenUpdating = True
 Application.EnableEvents = True
End Sub

Sub UnLockOS()

Application.ScreenUpdating = False
Application.EnableEvents = False
Dim Rangename, name As Variant
Dim i As Variant

Rangename = "os.Country||os.Article_DTAA||os.Rate_DTAA||os.TRC||os.SectionAct||os.Amount_income"
Rangename = Split(Rangename, "||")
'----------------Unlock Password-------------------START---
sPassword = EfilingCommon.getmsgstate
Sheet14.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
For i = 0 To UBound(Rangename)
    Sheet14.Range(Rangename(i)).value = ""
    Sheet14.Range(Rangename(i)).Interior.Color = (&HCCFFCC)
    Sheet14.Range(Rangename(i)).Locked = False
Next

Sheet14.Range("OS_DTAA_hiding").EntireRow.Hidden = False
Sheet14.Shapes("Button 40").Visible = msoTrue
Sheet14.Protect Password:=sPassword

Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub


Sub LockCG_STCGLTCG()
 Application.ScreenUpdating = False
  Application.EnableEvents = False
  Dim Rangename, name As Variant
  Dim i As Variant
  Rangename = "STCG.A8_Country||STCG.A8_DTAA_Article||STCG.A8_CertiStatus||STCG.A8_ItemNo||STCG.A8_StcgAmt||LTCG.B10_Country||LTCG.B10_DTAA_Article||LTCG.B10_CertiStatus||LTCG.B10_ItemNo||LTCG.B10_StcgAmt"
  Rangename = Split(Rangename, "||")
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
For i = 0 To UBound(Rangename)
    Sheet13.Range(Rangename(i)).value = ""
    Sheet13.Range(Rangename(i)).Interior.Color = (&HD8D8D8)
    Sheet13.Range(Rangename(i)).Locked = True
Next

 Sheet13.Protect Password:=sPassword

 Application.ScreenUpdating = True
 'Application.EnableEvents = True
 End Sub

Sub UnLockCG_STCGLTCG()
Application.ScreenUpdating = False
  Application.EnableEvents = False
  Dim Rangename, name As Variant
  Dim i As Variant
  Rangename = "STCG.A8_Country||STCG.A8_DTAA_Article||STCG.A8_CertiStatus||STCG.A8_ItemNo||STCG.A8_StcgAmt||LTCG.B10_Country||LTCG.B10_DTAA_Article||LTCG.B10_CertiStatus||LTCG.B10_ItemNo||LTCG.B10_StcgAmt"
  Rangename = Split(Rangename, "||")
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
For i = 0 To UBound(Rangename)
    Sheet13.Range(Rangename(i)).value = ""
    Sheet13.Range(Rangename(i)).Interior.Color = (&HCCFFCC)
    Sheet13.Range(Rangename(i)).Locked = False
Next

'-----------------Lock Password-----------------------------
 Sheet13.Protect Password:=sPassword
'-----------------------------------------------------------
 Application.ScreenUpdating = True
 'Application.EnableEvents = True
End Sub
Sub LOCKFA()
 
 Dim sourceSheet As Worksheet
 Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
  
  Application.ScreenUpdating = False
  Application.EnableEvents = False
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   sourceSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

Dim rngname_A1_FA As Variant
Dim rngname_A2_FA As Variant
Dim rngname_A3_FA As Variant
Dim rngname_A4_FA As Variant
Dim rngname_F_FA As Variant
Dim rngname_B_FA As Variant
Dim rngname_C_FA As Variant
Dim rngname_D_FA As Variant
Dim rngname_E_FA As Variant
Dim rngname_G_FA As Variant


rngname_A1_FA = "FA_A1_Country||FA_A1_BankName||FA_A1_BankAdd||FA_A1_ZipCode||FA_A1_ForeignAccountNumber||FA_A1_StatusBeneficiary||FA_A1_AccOpeningDate||FA_A1_PeakBal||FA_A1_ClosingBal||FA_A1_Grossinterest"
rngname_A2_FA = "FA_A2_Country||FA_A2_BankName||FA_A2_BankAdd||FA_A2_ZipCode||FA_A2_ForeignAccountNumber||FA_A2_StatusBeneficiary||FA_A2_AccOpeningDate||FA_A2_PeakBal||FA_A2_ClosingBal||FA_A2_Grossinterest||FA_A2_Grossinterest_Nature"
rngname_A3_FA = "FA_A3_Country||FA_A3_BankName||FA_A3_BankAdd||FA_A3_ZipCode||FA_A3_NatureOfEntity||FA_A3_AccOpeningDate||FA_A3_initialvalue||FA_A3_PeakBal||FA_A3_ClosingBal||FA_A3_Totalgrossamount||FA_A3_Totalgrosproceeds"
rngname_A4_FA = "FA_A4_Country||FA_A4_BankName||FA_A4_BankAdd||FA_A4_ZipCode||FA_A4_AccOpeningDate||FA_A4_the_cashvalue||FA_A4_Total_grossamount"
rngname_B_FA = "FA_B_Country||FA_B_ZipCode||FA_B_EntityNature||FA_B_EntityName||FA_B_EntityAdd||FA_B_NatureOfInterest||FA_B_DateSinceHeld||FA_B_TotalInv||FA_B_IncomeInterest||FA_B_NatureOfIncome||FA_B_Amount||FA_B_ScheduleOffered||FA_B_ItemNumOfSchedule"
rngname_C_FA = "FA_C_Country||FA_C_ZipCode||FA_C_PropertyAdd||FA_C_OwnershipBeneficiary||FA_C_DateOfAcquisition||FA_C_TotalInv||FA_C_IncomeFromProperty||FA_C_NatureOfIncome||FA_C_Amount||FA_C_ScheduleOffered||FA_C_ItemNumOfSchedule"
rngname_D_FA = "FA_D_Country||FA_D_ZipCode||FA_D_AssetNature||FA_D_OwnershipBeneficiary||FA_D_DateOfAcquisition||FA_D_TotalInv||FA_D_IncomeFromProperty||FA_D_NatureOfIncome||FA_D_Amount||FA_D_ScheduleOffered||FA_D_ItemNumOfSchedule"
rngname_E_FA = "FA_E_InstName||FA_E_InstAdd||FA_E_Country||FA_E_ZipCode||FA_E_AccountName||FA_E_InstitutionAccountNumber||FA_E_PeakBalInv||FA_E_IncomeAccured||FA_E_IncomeAccuredAmount||FA_E_Amount||FA_E_Scheduleoffered||FA_E_ItemNumberSchedule"

rngname_F_FA = "FA_F_Country||FA_F_ZipCode||FA_F_NameOfTrust||FA_F_AddressOfTrust||FA_F_NameOfOtherTrustees||FA_F_AddressOfOtherTrustees||FA_F_NameOfSettlor||FA_F_AddressOfSettlor||FA_F_NameOfBeneficiaries||FA_F_AddressOfBeneficiaries||FA_F_DateSinceHeld||FA_F_IncomeDerivedIsTaxable||FA_F_IncomeDerivedAmount||FA_F_Amount||FA_F_Scheduleoffered||FA_F_ItemNumberSchedule"
rngname_G_FA = "FA_G_Country||FA_G_ZipCode||FA_G_NameOfPerson||FA_G_AddressOfPerson||FA_G_IncomeDerived||FA_G_NatureOfIncome||FA_G_WhetherTaxable||FA_G_Amount||FA_G_Scheduleoffered||FA_G_ItemNumberSchedule"

LOCKRANGE (rngname_A1_FA)
LOCKRANGE (rngname_A2_FA)
LOCKRANGE (rngname_A3_FA)
LOCKRANGE (rngname_A4_FA)
LOCKRANGE (rngname_B_FA)
LOCKRANGE (rngname_C_FA)
LOCKRANGE (rngname_D_FA)
LOCKRANGE (rngname_E_FA)
LOCKRANGE (rngname_F_FA)
LOCKRANGE (rngname_G_FA)
'----------------Lock Password-------------------START---
   sourceSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub


Sub LOCKRANGE(rangenamestring As Variant)
Dim i As Long
    rangenamestring = Split(rangenamestring, "||")
    Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("TR_FA")
    For i = 0 To UBound(rangenamestring)
    sourceSheet.Range(rangenamestring(i)).value = ""
    sourceSheet.Range(rangenamestring(i)).Interior.Color = (&HD8D8D8)
        sourceSheet.Range(rangenamestring(i)).Locked = True
       
       
    Next

End Sub


'Sub UNLOCKcg()
'
'Application.ScreenUpdating = False
'  Application.EnableEvents = False
''----------------Unlock Password-------------------START---
'   sPassword = EfilingCommon.getmsgstate
'   Sheet13.Unprotect Password:=sPassword
''----------------Unlock Password-------------------END-----
'Dim rngname_CG As Variant
'
'rngname_CG = "STCG.A2iia_FullConsideration||STCG.A2iibi_AquisitCost||STCG.A2iibii_ImproveCost||STCG.A2iibiii_ExpOnTrans||STCG.A2iid_LossSec94of7Or94of8"
'UNLOCKRANGEcg (rngname_CG)
''Malli_AY_2025_26 V0.2 Change
'rngname_CG = "STCG.A2iia_FullConsideration_ii||STCG.A2iibi_AquisitCost_ii||STCG.A2iibii_ImproveCost_ii||STCG.A2iibiii_ExpOnTrans_ii||STCG.A2iid_LossSec94of7Or94of8_ii"
'UNLOCKRANGEcg (rngname_CG)
''-------------------------------
'
''Malli--AY_2024_25 OLd
''rngname_CG = "STCG.A3a_NRItaxSTTPaid||STCG.A3b_NRItaxSTTNotPaid||STCG.A4aia_FullConsideration||STCG.A4aib_FullMarketvalue||STCG.A4aii_FullConsiderationS||STCG.A4bi_AquisitCost||STCG.A4bii_ImproveCost||STCG.A4biii_ExpOnTrans||STCG.A4d_LossSec94of7Or94of8"
''Malli---CG_C51&52&53----AY_2025_26
'
''rngname_CG = "STCG.A3a_NRItaxSTTPaid||STCG.A3ai_NRItaxSTTPaid||STCG.A3aii_NRItaxSTTPaid||STCG.A3b_NRItaxSTTNotPaid||STCG.A4aia_FullConsideration||STCG.A4aib_FullMarketvalue||STCG.A4aii_FullConsiderationS||STCG.A4bi_AquisitCost||STCG.A4bii_ImproveCost||STCG.A4biii_ExpOnTrans||STCG.A4d_LossSec94of7Or94of8"
''UNLOCKRANGEcg (rngname_CG)
'
''Unit test change by sai on 22/02/2025=======================================================
'rngname_CG = "STCG.A3a_NRItaxSTTPaid||STCG.A3aii_NRItaxSTTPaid||STCG.A3b_NRItaxSTTNotPaid||STCG.A4aia_FullConsideration||STCG.A4aib_FullMarketvalue||STCG.A4aii_FullConsiderationS||STCG.A4bi_AquisitCost||STCG.A4bii_ImproveCost||STCG.A4biii_ExpOnTrans||STCG.A4d_LossSec94of7Or94of8"
''=============================================================================================
'
'
''-----------------------------
'rngname_CG = "LTCG.B4a_LTCGWithoutBenefit||LTCG.B4bii_ExemptionAmount"
'UNLOCKRANGEcg (rngname_CG)
'rngname_CG = "LTCG.Bi5bi_AquisitCost||LTCG.Bi5bii_ImproveCost||LTCG.Bi5biii_ExpOnTrans||LTCG.Bi5dii_ExemptionAmount||LTCG.Bii5bi_AquisitCost||LTCG.Bii5bii_ImproveCost||LTCG.Bii5biii_ExpOnTrans||LTCG.Bii5dii_ExemptionAmount||LTCG.Biii5bi_AquisitCost||LTCG.Biii5bii_ImproveCost||LTCG.Biii5biii_ExpOnTrans||LTCG.Biii5dii_ExemptionAmount||LTCG.B6a_SaleonSpecAsset||LTCG.B6b_DednSpecAssetus115||LTCG.B6d_SaleOtherSpecAsset||LTCG.B6e_DednOtherSpecAssetus115||LTCG.B6aia_FullConsideration||LTCG.B6aib_FullMarketValue||LTCG.B6aii_FullConsiderationS||LTCG.B6ibia_FullConsideration||LTCG.B6ibib_FullMarketValue||LTCG.B6ibii_FullConsideration50CA||LTCG.B6iiiaia_FullConsideration||LTCG.B6iiiaib_FullMarketValue||LTCG.B6iiiaii_FullConsiderationS"
'UNLOCKRANGEcg (rngname_CG)
'rngname_CG = "LTCG.8.DeductionUnder54F||LTCG.B06d_ExemptionGrandTotal||LTCG.B07dii_ExemptionGrandTotal||LTCG.B07div_ExemptionGrandTotal||LTCG.B07dv_ExemptionGrandTotal"
'UNLOCKRANGEcg (rngname_CG)
'
'''Malli-------AY_2025_26
''rngname_CG = "STCG.A8_AmtI||STCG.A8_IA1tA6||STCG.A8_Country||STCG.A8_DTAA_Article||STCG.A8_RatePTreaty||STCG.A8_CertiStatus||STCG.A8_SecIT||STCG.A8_RateIT"
''UNLOCKRANGEcg_8tab (rngname_CG)
''UNLOCKRANGEcg_8tab
'''------------------------
'
'Sheet13.Range("STCG.A.NRI").EntireRow.Hidden = False
'Sheet13.Range("LTCG.B.NRI").EntireRow.Hidden = False
'
''
''           Sheet54.Visible = xlSheetVisible
''   Sheet54.Unprotect Password:=sPassword
''rngname_CG = "ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||COAwithIndex_115AD||FairMarketValue_115AD||CostofImprovement_115AD||ExpenditureWholly_115AD"
''UNLOCKRANGE115AD (rngname_CG)
''Sheet54.Protect Password:=sPassword
' 'ThisWorkbook.Protect Password:=getmsgstate
'
'
'
''Sheet13.Range("STCG.DTAA").EntireRow.Hidden = False
''Sheet13.Range("LTCG.DTAA").EntireRow.Hidden = False
''Sheet13.Shapes("Button 89").Visible = msoTrue
''Sheet13.Shapes("Button 160").Visible = msoTrue
''Sheet13.Range("LTCG.B4d_LTCGOnListedSecurity").EntireRow.Hidden = True
'
'
''Sheet13.Range("LTCG.B4bi_ExemptionAmount").EntireRow.Hidden = True
''Sheet13.Range("LTCG.B4bii_ExemptionAmount").EntireRow.Hidden = True
''Sheet13.Range("LTCG.B4c_BalanceCG").EntireRow.Hidden = True
'
''----------------Lock Password-------------------START---
'   Sheet13.Protect Password:=sPassword
''----------------Lock Password-------------------END-----
'
' Application.EnableEvents = True
' Application.ScreenUpdating = True
'End Sub

Sub UNLOCKcg()
Application.ScreenUpdating = False
  Application.EnableEvents = False
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Dim rngname_CG As Variant

rngname_CG = "STCG.A2iia_FullConsideration||STCG.A2iibi_AquisitCost||STCG.A2iibii_ImproveCost||STCG.A2iibiii_ExpOnTrans||STCG.A2iid_LossSec94of7Or94of8"
UNLOCKRANGEcg (rngname_CG)
'Malli_AY_2025_26 V0.2 Change
'Commented by Riyaz on 22/01/2026
'rngname_CG = "STCG.A2iia_FullConsideration_ii||STCG.A2iibi_AquisitCost_ii||STCG.A2iibii_ImproveCost_ii||STCG.A2iibiii_ExpOnTrans_ii||STCG.A2iid_LossSec94of7Or94of8_ii"
'UNLOCKRANGEcg (rngname_CG)
'-------------------------------

'Malli--AY_2024_25 OLd
'rngname_CG = "STCG.A3a_NRItaxSTTPaid||STCG.A3b_NRItaxSTTNotPaid||STCG.A4aia_FullConsideration||STCG.A4aib_FullMarketvalue||STCG.A4aii_FullConsiderationS||STCG.A4bi_AquisitCost||STCG.A4bii_ImproveCost||STCG.A4biii_ExpOnTrans||STCG.A4d_LossSec94of7Or94of8"
'Malli---CG_C51&52&53----AY_2025_26

'rngname_CG = "STCG.A3a_NRItaxSTTPaid||STCG.A3ai_NRItaxSTTPaid||STCG.A3aii_NRItaxSTTPaid||STCG.A3b_NRItaxSTTNotPaid||STCG.A4aia_FullConsideration||STCG.A4aib_FullMarketvalue||STCG.A4aii_FullConsiderationS||STCG.A4bi_AquisitCost||STCG.A4bii_ImproveCost||STCG.A4biii_ExpOnTrans||STCG.A4d_LossSec94of7Or94of8"


'Unit test change by sai on 22/02/2025=======================================
'Added one range name by Riyaz on 12/02/2026
rngname_CG = "STCG.A3a_NRItaxSTTPaid_i_ii||STCG.A3a_NRItaxSTTPaid||STCG.A3aii_NRItaxSTTPaid||STCG.A3b_NRItaxSTTNotPaid||STCG.A4aia_FullConsideration||STCG.A4aib_FullMarketvalue||STCG.A4aii_FullConsiderationS||STCG.A4bi_AquisitCost||STCG.A4bii_ImproveCost||STCG.A4biii_ExpOnTrans||STCG.A4d_LossSec94of7Or94of8"
'=============================================================================

UNLOCKRANGEcg (rngname_CG)

'-----------------------------
'Commented by Riyaz on 05/03/2025
'rngname_CG = "LTCG.B4a_LTCGWithoutBenefit||LTCG.B4bii_ExemptionAmount"
'UNLOCKRANGEcg (rngname_CG)
'rngname_CG = "LTCG.Bi5bi_AquisitCost||LTCG.Bi5bii_ImproveCost||LTCG.Bi5biii_ExpOnTrans||LTCG.Bi5dii_ExemptionAmount||LTCG.Bii5bi_AquisitCost||LTCG.Bii5bii_ImproveCost||LTCG.Bii5biii_ExpOnTrans||LTCG.Bii5dii_ExemptionAmount||LTCG.Biii5bi_AquisitCost||LTCG.Biii5bii_ImproveCost||LTCG.Biii5biii_ExpOnTrans||LTCG.Biii5dii_ExemptionAmount||LTCG.B6a_SaleonSpecAsset||LTCG.B6b_DednSpecAssetus115||LTCG.B6d_SaleOtherSpecAsset||LTCG.B6e_DednOtherSpecAssetus115||LTCG.B6aia_FullConsideration||LTCG.B6aib_FullMarketValue||LTCG.B6aii_FullConsiderationS||LTCG.B6ibia_FullConsideration||LTCG.B6ibib_FullMarketValue||LTCG.B6ibii_FullConsideration50CA||LTCG.B6iiiaia_FullConsideration||LTCG.B6iiiaib_FullMarketValue||LTCG.B6iiiaii_FullConsiderationS"
'Commented by Riyaz on 05/03/2025
'rngname_CG = "LTCG.Bi5bi_AquisitCost||LTCG.Bi5bii_ImproveCost||LTCG.Bi5biii_ExpOnTrans||LTCG.Bi5dii_ExemptionAmount||LTCG.Bii5bi_AquisitCost||LTCG.Bii5bii_ImproveCost||LTCG.Bii5biii_ExpOnTrans||LTCG.Bii5dii_ExemptionAmount||LTCG.Biii5bi_AquisitCost||LTCG.Biii5bii_ImproveCost||LTCG.Biii5biii_ExpOnTrans||LTCG.Biii5dii_ExemptionAmount||LTCG.B6a_SaleonSpecAsset||LTCG.B6b_DednSpecAssetus115||LTCG.B6d_SaleOtherSpecAsset||LTCG.B6aia_FullConsideration||LTCG.B6aib_FullMarketValue||LTCG.B6aii_FullConsiderationS||LTCG.B6ibia_FullConsideration||LTCG.B6ibib_FullMarketValue||LTCG.B6ibii_FullConsideration50CA||LTCG.B6iiiaia_FullConsideration||LTCG.B6iiiaib_FullMarketValue||LTCG.B6iiiaii_FullConsiderationS"
'
'UNLOCKRANGEcg (rngname_CG)
'rngname_CG = "LTCG.8.DeductionUnder54F||LTCG.B06d_ExemptionGrandTotal||LTCG.B07dii_ExemptionGrandTotal||LTCG.B07div_ExemptionGrandTotal||LTCG.B07dv_ExemptionGrandTotal"
'Commented by Riyaz on 05/03/2025
'rngname_CG = "LTCG.8.DeductionUnder54F||LTCG.B06d_ExemptionGrandTotal||LTCG.B07dii_ExemptionGrandTotal||LTCG.B07div_ExemptionGrandTotal"
'UNLOCKRANGEcg (rngname_CG)

''Malli-------AY_2025_26
'rngname_CG = "STCG.A8_AmtI||STCG.A8_IA1tA6||STCG.A8_Country||STCG.A8_DTAA_Article||STCG.A8_RatePTreaty||STCG.A8_CertiStatus||STCG.A8_SecIT||STCG.A8_RateIT"
'UNLOCKRANGEcg_8tab (rngname_CG)
'UNLOCKRANGEcg_8tab
''------------------------


'NEWLY ADDED BY BINDU FOR TABLE B5
'2 Range name added by Riyaz on 12/02/2026
Sheet13.Unprotect Password:=sPassword
rngname_CG = "LTCG.B4a_LTCGWithoutBenefit||LTCG.B06d_ExemptionGrandTotal||LTCG.B4ai_LTCGWithoutBenefit||LTCG.B4aii_LTCGWithoutBenefit||LTCG.B4aiii_LTCGWithoutBenefit||LTCG.B06d_i_ExemptionGrandTotal||LTCG.B06d_ii_ExemptionGrandTotal||LTCG.B06d_iii_ExemptionGrandTotal"
UNLOCKRANGEcg (rngname_CG)


'NEWLY ADDED BY BINDU FOR TABLE B6i_BE
'Commented by Riyaz on 23/01/2026
'Sheet13.Unprotect Password:=sPassword
'rngname_CG = "LTCG.B6aia_FullConsideration_ii||LTCG.B6aib_FullMarketValue_ii||LTCG.B6aii_FullConsiderationS_ii||LTCG.Bi5bi_AquisitCost_ii||LTCG.Bi5bii_ImproveCost_ii||LTCG.Bi5biii_ExpOnTrans_ii||LTCG.B07dii_ExemptionGrandTotal_ii"
'UNLOCKRANGEcg (rngname_CG)

'NEWLY ADDED BY BINDU FOR TABLE B6i_AE
Sheet13.Unprotect Password:=sPassword
rngname_CG = "LTCG.B6aia_FullConsideration||LTCG.B6aib_FullMarketValue||LTCG.B6aii_FullConsiderationS||LTCG.Bi5bi_AquisitCost||LTCG.Bi5bii_ImproveCost||LTCG.Bi5biii_ExpOnTrans||LTCG.B07dii_ExemptionGrandTotal"
UNLOCKRANGEcg (rngname_CG)

'NEWLY ADDED BY BINDU FOR TABLE B6ii_BE
'Commented by Riyaz on 23/01/2026
'Sheet13.Unprotect Password:=sPassword
'rngname_CG = "LTCG.B6ibia_FullConsideration_ii||LTCG.B6ibib_FullMarketValue_ii||LTCG.B6ibii_FullConsideration50CA_ii||LTCG.Bii5bi_AquisitCost_ii||LTCG.Bii5bii_ImproveCost_ii||LTCG.Bii5biii_ExpOnTrans_ii||LTCG.B07div_ExemptionGrandTotal_ii"
'UNLOCKRANGEcg (rngname_CG)

'NEWLY ADDED BY BINDU FOR TABLE B6ii_AE
Sheet13.Unprotect Password:=sPassword
rngname_CG = "LTCG.B6ibia_FullConsideration||LTCG.B6ibib_FullMarketValue||LTCG.B6ibii_FullConsideration50CA||LTCG.Bii5bi_AquisitCost||LTCG.Bii5bii_ImproveCost||LTCG.Bii5biii_ExpOnTrans||LTCG.B07div_ExemptionGrandTotal"
UNLOCKRANGEcg (rngname_CG)

'NEWLY ADDED BY BINDU FOR TABLE B6iii
Sheet13.Unprotect Password:=sPassword
rngname_CG = "LTCG.B6iiiaia_FullConsideration||LTCG.B6iiiaib_FullMarketValue||LTCG.B6iiiaii_FullConsiderationS||LTCG.Biii5bi_AquisitCost||LTCG.Biii5bii_ImproveCost||LTCG.Biii5biii_ExpOnTrans||LTCG.Biii5dii_ExemptionAmount"
UNLOCKRANGEcg (rngname_CG)

'NEWLY ADDED BY BINDU FOR TABLE B7
'1 Range name added by Riyaz on 12/02/2026
Sheet13.Unprotect Password:=sPassword
rngname_CG = "LTCG.8.DeductionUnder54F||LTCG.8_i.DeductionUnder54F||LTCG.8_ii.DeductionUnder54F"
UNLOCKRANGEcg (rngname_CG)

'NEWLY ADDED BY BINDU FOR TABLE B8
'2 Range name added by Riyaz on 12/02/2026
Sheet13.Unprotect Password:=sPassword
rngname_CG = "LTCG.B6a_SaleonSpecAsset||LTCG.B6b_DednSpecAssetus115||LTCG.B6ai_SaleonSpecAsset||LTCG.B6aii_SaleonSpecAsset||LTCG.B6bi_DednSpecAssetus115||LTCG.B6bii_DednSpecAssetus115"
UNLOCKRANGEcg (rngname_CG)



Sheet13.Range("STCG.A.NRI").EntireRow.Hidden = False
Sheet13.Range("LTCG.B.NRI").EntireRow.Hidden = False
Sheet13.Range("STCG.a3aia3aii").EntireRow.Hidden = True

Sheet13.Range("LTCG.B4a").EntireRow.Hidden = True
Sheet13.Range("LTCG.B4b").EntireRow.Hidden = True
Sheet13.Range("LTCG.B4c").EntireRow.Hidden = True
Sheet13.Range("LTCG.B5a").EntireRow.Hidden = True
Sheet13.Range("LTCG.B5b").EntireRow.Hidden = True
Sheet13.Range("LTCG.B5c").EntireRow.Hidden = True
Sheet13.Range("LTCG.B6a").EntireRow.Hidden = True
Sheet13.Range("LTCG.B6b").EntireRow.Hidden = True
Sheet13.Range("LTCG.B6c").EntireRow.Hidden = True
Sheet13.Range("LTCG.B7a").EntireRow.Hidden = True
Sheet13.Range("LTCG.B7b").EntireRow.Hidden = True
Sheet13.Range("LTCG.B7c").EntireRow.Hidden = True

'
'           Sheet54.Visible = xlSheetVisible
'   Sheet54.Unprotect Password:=sPassword
'rngname_CG = "ISINCode_115AD||NameofShare_115AD||NoofShare_115AD||SalePriceperShare_115AD||COAwithIndex_115AD||FairMarketValue_115AD||CostofImprovement_115AD||ExpenditureWholly_115AD"
'UNLOCKRANGE115AD (rngname_CG)
'Sheet54.Protect Password:=sPassword
 'ThisWorkbook.Protect Password:=getmsgstate


           
'Sheet13.Range("STCG.DTAA").EntireRow.Hidden = False
'Sheet13.Range("LTCG.DTAA").EntireRow.Hidden = False
'Sheet13.Shapes("Button 89").Visible = msoTrue
'Sheet13.Shapes("Button 160").Visible = msoTrue
'Sheet13.Range("LTCG.B4d_LTCGOnListedSecurity").EntireRow.Hidden = True


'Sheet13.Range("LTCG.B4bi_ExemptionAmount").EntireRow.Hidden = True
'Sheet13.Range("LTCG.B4bii_ExemptionAmount").EntireRow.Hidden = True
'Sheet13.Range("LTCG.B4c_BalanceCG").EntireRow.Hidden = True

'----------------Lock Password-------------------START---
   Sheet13.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

 Application.EnableEvents = True
 Application.ScreenUpdating = True

End Sub



Sub UNLOCKRANGEcg(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet13.Range(rangenamestring(i))
            chcell.Locked = False
            chcell.Interior.Color = (&HCCFFCC)
'            chcell.value = 0
            Next
        Next
End Sub


Function ChkMinDOBDate(dob As Variant, minDefinedDOB As Variant) As Boolean
On Error Resume Next
     ChkMinDOBDate = True
     If Len(dob) > 0 Then
        If val(Mid(dob, 7, 4)) <= 2024 Then 'dpk1601
              If val(Mid(dob, 4, 2)) <= 3 Then
                If val(Mid(dob, 1, 2)) <= 31 Then
                    ChkMinDOBDate = False
                    Exit Function
                End If
            End If
        End If
     End If
End Function
'Malli_08/09/2025
Function Enddateofthemonth(startDate As Variant) As Boolean
Enddateofthemonth = True
 
If Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 1 Then
    Enddateofthemonth = False
ElseIf (Mid(startDate, 1, 2) = 28 Or Mid(startDate, 1, 2) = 29) And Mid(startDate, 4, 2) = 2 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 3 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 30 And Mid(startDate, 4, 2) = 4 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 5 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 30 And Mid(startDate, 4, 2) = 6 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 7 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 8 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 30 And Mid(startDate, 4, 2) = 9 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 10 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 30 And Mid(startDate, 4, 2) = 11 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 12 Then
    Enddateofthemonth = False
Else
    Enddateofthemonth = True
End If
End Function
Function calcNoOfMonths(currentdate As Variant, startDate As Variant) As Long
On Error Resume Next

Dim currentYear As Variant
Dim startyear As Variant
Dim currentMonth As Variant
Dim startmonth As Variant

currentYear = val(Mid(currentdate, 7, 4))
startyear = val(Mid(startDate, 7, 4))

'Added by Shrutika 15/07/2025
'If getDueDate = "15/09/2025" Then

'Malli_08/09/2025
If getDueDate = Sheet33.Range("Duedate_Extended").value Then
If Enddateofthemonth(startDate) Then
    currentMonth = val(Mid(currentdate, 4, 2)) + 1
Else
    currentMonth = val(Mid(currentdate, 4, 2))
End If
'----------------

End If

startmonth = val(Mid(startDate, 4, 2))
calcNoOfMonths = 0
 
 
 If currentYear = startyear And currentMonth = startmonth And _
        val(Mid(currentdate, 1, 2)) = val(Mid(startDate, 1, 2)) Then
       
        calcNoOfMonths = 0
ElseIf (checkFirstDateBefore(currentdate, startDate)) Then
  
        calcNoOfMonths = 0
Else
    If currentYear = (startyear + 1) Then
    
        If currentMonth < startmonth Then
    
            'calcNoOfMonths = 12 - (startmonth - currentMonth) + 1
            calcNoOfMonths = 12 - (startmonth - currentMonth)
        Else
        'tobetested
    
            'calcNoOfMonths = 12 + (currentMonth - startmonth) + 1
            calcNoOfMonths = 12 + (currentMonth - startmonth)
        End If
   ElseIf (currentYear = startyear) Then
        'calcNoOfMonths = currentMonth - startmonth + 1
        calcNoOfMonths = currentMonth - startmonth
   If (val(Mid(currentdate, 1, 2)) > val(Mid(startDate, 1, 2))) And (currentMonth = startmonth) Then
        calcNoOfMonths = calcNoOfMonths + 1
        
   'ElseIf val(Mid(startDate, 1, 2)) = 15 And (currentMonth >= startmonth) Then
   ElseIf val(Mid(startDate, 1, 2)) = 7 And (currentMonth >= startmonth) Then '15/07/2025
        calcNoOfMonths = calcNoOfMonths + 1
     End If
   ElseIf currentMonth < startmonth Then
   
        'calcNoOfMonths = Round((currentYear - startyear - 1) * 12) + 12 - startmonth + currentMonth + 1
        calcNoOfMonths = Round((currentYear - startyear - 1) * 12) + 12 - startmonth + currentMonth
        
   ElseIf (currentMonth > startmonth) Then
   
        'calcNoOfMonths = Round(((currentYear - startyear) * 12)) + currentMonth - startmonth + 1
        calcNoOfMonths = Round(((currentYear - startyear) * 12)) + currentMonth - startmonth
    Else
   
        'calcNoOfMonths = Round(((currentYear - startyear) * 12)) + 1
        calcNoOfMonths = Round(((currentYear - startyear) * 12))
    End If
End If
'Added by Shrutika 15/07/2025
If currentYear = (startyear + 1) Then

If val(Mid(startDate, 1, 2)) = 7 And (currentYear > startyear) Then
        calcNoOfMonths = calcNoOfMonths + 1
End If
End If

End Function

Function calcNoOfMonthsold1(currentdate As Variant, startDate As Variant) As Long
On Error Resume Next

Dim currentYear As Variant
Dim startyear As Variant
Dim currentMonth As Variant
Dim startmonth As Variant

currentYear = val(Mid(currentdate, 7, 4))
startyear = val(Mid(startDate, 7, 4))
currentMonth = val(Mid(currentdate, 4, 2))
startmonth = val(Mid(startDate, 4, 2))

 calcNoOfMonths = 0
 
 
 If currentYear = startyear And currentMonth = startmonth And _
        val(Mid(currentdate, 1, 2)) = val(Mid(startDate, 1, 2)) Then
        calcNoOfMonths = 1
ElseIf (checkFirstDateBefore(currentdate, startDate)) Then
        calcNoOfMonths = 0
Else
    If currentYear = (startyear + 1) Then
        If currentMonth < startmonth Then
            calcNoOfMonths = 12 - (startmonth - currentMonth) + 1
        Else
            calcNoOfMonths = 12 + (currentMonth - startmonth) + 1
        End If
   ElseIf (currentYear = startyear) Then
        'calcNoOfMonths = currentMonth - startmonth + 1
        calcNoOfMonths = currentMonth - startmonth
   ElseIf currentMonth < startmonth Then
        calcNoOfMonths = Round((currentYear - startyear - 1) * 12) + 12 - startmonth + currentMonth + 1
   ElseIf (currentMonth > startmonth) Then
        calcNoOfMonths = Round(((currentYear - startyear) * 12)) + currentMonth - startmonth + 1
    Else
        calcNoOfMonths = Round(((currentYear - startyear) * 12)) + 1
    End If
End If

End Function

Function checkfieldspecialcharacter1(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter1 = True
    Dim arr As Variant
    arr = Array("@", ".", "*", "!", "&", "#", "~", ";", "?", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
    'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter1 = False
            Exit Function
        End If
        Next
    Next
End Function
Function checkfieldspecialcharacter101(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter101 = True
    Dim arr As Variant
    arr = Array("@", ".", "*", "!", "&", "#", "~", ";", "?", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<", ",", "-", "_", "/", "|") 'Array("&", """", "'", ">", "<")
    'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter101 = False
            Exit Function
        End If
        Next
    Next
End Function

Function CheckDateddmmyyyy(dt As Variant) As Boolean
    CheckDateddmmyyyy = True
    
    If InStr(1, dt, ".") > 0 Then
        CheckDateddmmyyyy = False
    End If
    
    If Len(dt) > 0 Then
        
        If Mid(dt, 3, 1) <> "/" Then
            If Mid(dt, 3, 1) <> "\" Then
                If Mid(dt, 3, 1) <> "-" Then
                    If Mid(dt, 3, 1) <> "." Then
                        CheckDateddmmyyyy = False
                    Else
                        dt = Mid(dt, 1, 2) & "/" & Mid(dt, 4, 7)
                        CheckDateddmmyyyy = False
                    End If
                Else
                    dt = Mid(dt, 1, 2) & "/" & Mid(dt, 4, 7)
                    CheckDateddmmyyyy = False
                End If
            Else
                dt = Mid(dt, 1, 2) & "/" & Mid(dt, 4, 7)
                CheckDateddmmyyyy = False
            End If
        End If
        
        If Mid(dt, 6, 1) <> "/" Then
            If Mid(dt, 6, 1) <> "-" Then
                If Mid(dt, 6, 1) <> "\" Then
                    If Mid(dt, 6, 1) <> "." Then
                        CheckDateddmmyyyy = False
                    Else
                        dt = Mid(dt, 1, 5) & "/" & Mid(dt, 7, 4)
                        CheckDateddmmyyyy = False
                    End If
                Else
                    dt = Mid(dt, 1, 5) & "/" & Mid(dt, 7, 4)
                    CheckDateddmmyyyy = False
                End If
            Else
                dt = Mid(dt, 1, 5) & "/" & Mid(dt, 7, 4)
                CheckDateddmmyyyy = False
            End If
        End If
        
        If Not IsDate(dt) Then CheckDateddmmyyyy = False
        If val(Mid(dt, 1, 2)) < 0 Then CheckDateddmmyyyy = False
        If val(Mid(dt, 1, 2)) > 31 Then CheckDateddmmyyyy = False
        If val(Mid(dt, 4, 2)) < 0 Then CheckDateddmmyyyy = False
        If val(Mid(dt, 4, 2)) > 12 Then CheckDateddmmyyyy = False
        If val(Mid(dt, 7, 4)) < 1 Then CheckDateddmmyyyy = False
        If val(Mid(dt, 7, 4)) > 3000 Then CheckDateddmmyyyy = False
        
'If Mid(dt1, 3, 1) = "-" Then CheckDateddmmyyyy = False
'If Mid(dt1, 3, 1) = "\" Then CheckDateddmmyyyy = False
'If Mid(dt1, 3, 1) = "." Then CheckDateddmmyyyy = False
'
'If Mid(dt1, 6, 1) = "-" Then CheckDateddmmyyyy = False
'If Mid(dt1, 6, 1) = "\" Then CheckDateddmmyyyy = False
'If Mid(dt1, 6, 1) = "." Then CheckDateddmmyyyy = False
    End If
End Function

Function CheckDateMinDDMMYYYY(ByVal dt As String, ByVal minday As Long, ByVal minmonth As Long, ByVal minyear As Long, ByVal errormsg As String) As Boolean
    CheckDateMinDDMMYYYY = True

    If (val(Mid(dt, 1, 4)) < minyear) Then
        'validatedt = False
        CheckDateMinDDMMYYYY = False
        If errormsg <> "" Then fmsgboxStatus "INVALID DATE, " & errormsg
        GoTo exit1
        
    End If
    
    If (val(Mid(dt, 1, 4)) = minyear) And (val(Mid(dt, 6, 2)) < minmonth) Then
        CheckDateMinDDMMYYYY = False
        If errormsg <> "" Then fmsgboxStatus "INVALID DATE, " & errormsg
        GoTo exit1
    End If
    If (val(Mid(dt, 1, 4)) = minyear) And (val(Mid(dt, 6, 2)) < minmonth) And (val(Mid(dt, 9, 2)) < minday) Then
        CheckDateMinDDMMYYYY = False
        If errormsg <> "" Then fmsgboxStatus "INVALID DATE, " & errormsg
        GoTo exit1
    End If

exit1:
End Function

Function CheckAtoZ(chr1) As Boolean
    CheckAtoZ = True
    If ((asc(chr1) < 65) Or (asc(chr1) > 90)) Then
        CheckAtoZ = False
    End If
End Function

'Function calcTaxPayableOnTINTR(totInc As Double) As Double
'On Error Resume Next
'
'Dim tempTax, taxPayable As Double
'taxPayable = 0
'
'If (totInc <= 250000) Then
'       taxPayable = 0
'ElseIf totInc >= 250001 And totInc <= 500000 Then
'        tempTax = (totInc - 250000) * (0.05)
'        taxPayable = Round(tempTax, 0)
'
'ElseIf totInc >= 500001 And totInc <= 750000 Then
'        tempTax = (totInc - 500000) * (0.1)
'        taxPayable = Round((tempTax + 12500), 0)
'
'ElseIf totInc >= 750001 And totInc <= 1000000 Then
'        tempTax = (totInc - 750000) * (0.15)
'        taxPayable = Round((tempTax + 37500), 0)
'ElseIf totInc >= 1000001 And totInc <= 1250000 Then
'        tempTax = (totInc - 1000000) * (0.2)
'        taxPayable = Round((tempTax + 75000), 0)
'ElseIf totInc >= 1250001 And totInc <= 1500000 Then
'        tempTax = (totInc - 1250000) * (0.25)
'        taxPayable = Round((tempTax + 125000), 0)
'ElseIf totInc >= 1500001 Then
'        tempTax = (totInc - 1500000) * (0.3)
'        taxPayable = Round((tempTax + 187500), 0)
'End If
'
'Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
'calcTaxPayableOnTINTR = taxPayable
'Sheet1.Protect Password:=EfilingCommon.getmsgstate
'
'End Function



'Chandru AY2024-25

Function calcTaxPayableOnTINTR(totInc As Double) As Double
On Error Resume Next

Dim tempTax, taxPayable As Double
taxPayable = 0
           
If (totInc <= 400000) Then
       taxPayable = 0
'Malli-------
'AY_2024_25 OLD
'ElseIf totInc >= 300001 And totInc <= 600000 Then
'AY_2025_26 NEW
ElseIf totInc >= 400001 And totInc <= 800000 Then
'--------------------
        tempTax = (totInc - 400000) * (0.05)
        taxPayable = Round(tempTax, 0)
 'AY_2024_25 OLD
'ElseIf totInc >= 600001 And totInc <= 900000 Then
'        tempTax = (totInc - 600000) * (0.1)
'        taxPayable = Round((tempTax + 15000), 0)
  'AY_2025_26 NEW
  ElseIf totInc >= 800001 And totInc <= 1200000 Then
        tempTax = (totInc - 800000) * (0.1)
        taxPayable = Round((tempTax + 20000), 0)
  '-----------------------
  'AY_2024_25 OLD
'ElseIf totInc >= 900001 And totInc <= 1200000 Then
'        tempTax = (totInc - 900000) * (0.15)
'        taxPayable = Round((tempTax + 45000), 0)
  'AY_2025_26 NEW
  ElseIf totInc >= 1200001 And totInc <= 1600000 Then
        tempTax = (totInc - 1200000) * (0.15)
        taxPayable = Round((tempTax + 60000), 0)
  '----------------------
ElseIf totInc >= 1600001 And totInc <= 2000000 Then
        tempTax = (totInc - 1600000) * (0.2)
        'AY_2024_25 OLD
        'taxPayable = Round((tempTax + 90000), 0)
         'AY_2025_26 NEW
         taxPayable = Round((tempTax + 120000), 0)
         '-----------------
'ElseIf totInc >= 1250001 And totInc <= 1500000 Then
'        tempTax = (totInc - 1250000) * (0.25)
'        taxPayable = Round((tempTax + 125000), 0)

'Added by Riyaz on 03/02/2026
ElseIf totInc >= 2000001 And totInc <= 2400000 Then
'        tempTax = (totInc - 200000) * (0.25)
         tempTax = (totInc - 2000000) * (0.25) 'by sadineni on 14/04/2026 as per Riyaz
        'AY_2024_25 OLD
        'taxPayable = Round((tempTax + 90000), 0)
         'AY_2025_26 NEW
         taxPayable = Round((tempTax + 200000), 0)

ElseIf totInc >= 2400001 Then
        tempTax = (totInc - 2400000) * (0.3)
        'AY_2024_25 OLD
       ' taxPayable = Round((tempTax + 150000), 0)
        'AY_2025_26 NEW
        taxPayable = Round((tempTax + 300000), 0)
        '----------------
End If

Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
calcTaxPayableOnTINTR = taxPayable
Sheet1.Protect Password:=EfilingCommon.getmsgstate

End Function

Function calculateSlabbedTax(netTxblIncome As Double) As Double
Dim exemption As Double
Dim resStatus As String
Dim age As String
Dim Status As String

 netTxblIncome = netTxblIncome
 age = Sheet31.Range("age").value
 resStatus = Sheet33.Range("resStatus").value
 Status = Sheet33.Range("taxPayerStatus").value
           
If (Status = "I" And (resStatus = "RES" Or resStatus = "NOR") And (age > 59 And age < 80)) Then
                If (netTxblIncome >= 0 And netTxblIncome <= 300000) Then
                    calculateSlabbedTax = 0
                ElseIf (netTxblIncome >= 300001 And netTxblIncome <= 500000) Then
                    calculateSlabbedTax = (netTxblIncome - 300000) * 0.05
                ElseIf (netTxblIncome >= 500001 And netTxblIncome <= 1000000) Then
                    calculateSlabbedTax = (netTxblIncome - 500000) * 0.2 + 10000
                ElseIf (netTxblIncome >= 1000001) Then
                    calculateSlabbedTax = (netTxblIncome - 1000000) * 0.3 + 110000
                End If
           
ElseIf (Status = "I" And (resStatus = "RES" Or resStatus = "NOR") And (age >= 80)) Then
                If (netTxblIncome >= 0 And netTxblIncome <= 500000) Then
                    calculateSlabbedTax = 0
                ElseIf (netTxblIncome >= 500001 And netTxblIncome <= 1000000) Then
                    calculateSlabbedTax = (netTxblIncome - 500000) * 0.2
                ElseIf (netTxblIncome >= 1000001) Then
                    calculateSlabbedTax = (netTxblIncome - 1000000) * 0.3 + 100000
                End If
Else
                If (netTxblIncome >= 0 And netTxblIncome <= 250000) Then
                    calculateSlabbedTax = 0
                ElseIf (netTxblIncome >= 250001 And netTxblIncome <= 500000) Then
                    calculateSlabbedTax = (netTxblIncome - 250000) * 0.05  '
                ElseIf (netTxblIncome >= 500001 And netTxblIncome <= 1000000) Then
                    calculateSlabbedTax = (netTxblIncome - 500000) * 0.2 + 12500
                ElseIf (netTxblIncome >= 1000001) Then
                    calculateSlabbedTax = (netTxblIncome - 1000000) * 0.3 + 112500
                End If
End If

End Function
Function getmsgstate() As String
    pwd = SheetALL.Range("pwd")
    getmsgstate = pwd
End Function
Sub ProtectAll()
ThisWorkbook.Protect Password:=EfilingCommon.getmsgstate
Sheet1.Protect Password:=getmsgstate
Sheet10.Protect Password:=getmsgstate
Sheet11.Protect Password:=getmsgstate
Sheet12.Protect Password:=getmsgstate
Sheet13.Protect Password:=getmsgstate
Sheet14.Protect Password:=getmsgstate
Sheet15.Protect Password:=getmsgstate
Sheet16.Protect Password:=getmsgstate
Sheet17.Protect Password:=getmsgstate
Sheet18.Protect Password:=getmsgstate
Sheet19.Protect Password:=getmsgstate
Sheet2.Protect Password:=getmsgstate
Sheet20.Protect Password:=getmsgstate
Sheet21.Protect Password:=getmsgstate
Sheet22.Protect Password:=getmsgstate
Sheet23.Protect Password:=getmsgstate
Sheet24.Protect Password:=getmsgstate
Sheet25.Protect Password:=getmsgstate
Sheet27.Protect Password:=getmsgstate
Sheet28.Protect Password:=getmsgstate
Sheet29.Protect Password:=getmsgstate
Sheet3.Protect Password:=getmsgstate
Sheet30.Protect Password:=getmsgstate
Sheet31.Protect Password:=getmsgstate
Sheet32.Protect Password:=getmsgstate
Sheet33.Protect Password:=getmsgstate
Sheet34.Protect Password:=getmsgstate
Sheet35.Protect Password:=getmsgstate
Sheet36.Protect Password:=getmsgstate
Sheet37.Protect Password:=getmsgstate
Sheet39.Protect Password:=getmsgstate
Sheet4.Protect Password:=getmsgstate
Sheet40.Protect Password:=getmsgstate
Sheet41.Protect Password:=getmsgstate
Sheet42.Protect Password:=getmsgstate
Sheet43.Protect Password:=getmsgstate
Sheet44.Protect Password:=getmsgstate
Sheet46.Protect Password:=getmsgstate
Sheet48.Protect Password:=getmsgstate
Sheet49.Protect Password:=getmsgstate
Sheet5.Protect Password:=getmsgstate
Sheet50.Protect Password:=getmsgstate
Sheet51.Protect Password:=getmsgstate
Sheet52.Protect Password:=getmsgstate
Sheet53.Protect Password:=getmsgstate
Sheet54.Protect Password:=getmsgstate
Sheet55.Protect Password:=getmsgstate
Sheet56.Protect Password:=getmsgstate
Sheet57.Protect Password:=getmsgstate
Sheet6.Protect Password:=getmsgstate
Sheet7.Protect Password:=getmsgstate
Sheet8.Protect Password:=getmsgstate
Sheet9.Protect Password:=getmsgstate
Sheet60.Protect Password:=getmsgstate
SheetALL.Protect Password:=getmsgstate



End Sub
Sub Trymed()
SheetALL.Visible = False

Sheet33.Visible = False
Sheet31.Visible = False


End Sub


Function applyExemption(exemption As Double, x As Long, QtrIncm As Range) As Variant
Dim Total As Double
Dim i As Long
Dim part As Double
Dim original(1 To 5) As Variant
Dim count As Long
Dim A As Variant
 i = 1
For Each A In QtrIncm
    original(i) = A
    i = i + 1
Next
Line1:
Total = 0

For i = 1 To 5
    If (original(i) > 0) Then
        count = count + 1
    End If
Next
If (exemption > 0) Then
 For i = 1 To 5
    If (original(i) > 0) Then
        'part = (exemption / Count)
        part = (exemption)
        count = count - 1
        If (original(i) > part) Then
            exemption = exemption - part
            original(i) = original(i) - part
        Else
             exemption = exemption - original(i)
             original(i) = 0
        End If
    End If
    Total = Total + original(i)
 Next
  If (exemption > 0 And Total > 0) Then
    GoTo Line1
End If
End If
applyExemption = original(x)
'original (x)
'part (x)
End Function


Function InsertRowsAndFillFormulas(Optional vRows As Long = 0) As Long
    
    Dim x As Long
    Dim sht As Worksheet
    Dim shts() As String
    Dim i As Long
    ActiveCell.EntireRow.Select  'So you do not have to preselect entire row
    If vRows = 0 Then
        vRows = Application.InputBox(prompt:= _
        "Enter the number of rows you want to add below selected cell", Title:="Add Rows below the selected cell", _
        Default:=1, Type:=1) 'Default for 1 row, type 1 is number
        If vRows = False Then
            InsertRowsAndFillFormulas = 0
            Exit Function
        End If
    End If
    
    sPassword = EfilingCommon.getmsgstate
    ActiveSheet.Unprotect Password:=sPassword
    
   
    ReDim shts(1 To Worksheets.Application.ActiveWorkbook. _
        Windows(1).SelectedSheets.count)
    i = 0
    For Each sht In _
        Application.ActiveWorkbook.Windows(1).SelectedSheets
        Sheets(sht.name).Select
        
        i = i + 1
        shts(i) = sht.name
    
        x = Sheets(sht.name).UsedRange.Rows.count 'lastcell fixup
    
        Selection.Resize(rowsize:=2).Rows(2).EntireRow. _
        Resize(rowsize:=vRows).Insert Shift:=xlDown
         Sheets(sht.name).Unprotect Password:=sPassword
        Selection.AutoFill Selection.Resize( _
        rowsize:=vRows + 1), xlFillDefault
    
        On Error Resume Next
         
         sPassword = EfilingCommon.getmsgstate
    ActiveSheet.Unprotect Password:=sPassword
    
        Selection.Offset(1).Resize(vRows).EntireRow. _
        SpecialCells(xlCellTypeAllValidation).ClearContents
    
    Next sht
    Worksheets(shts).Select
    InsertRowsAndFillFormulas = vRows

    ActiveSheet.Protect Password:=sPassword
End Function

Sub ExendRangeNameToTable(numberofrows As Long, rangenamestring As Variant)
    Dim i As Long
    Dim x As Long
    Dim firstbound As String
    Dim temp As Variant
    Dim upperbound As String
    Dim lastbound As String
    
    rangenamestring = Split(rangenamestring, ";")
    For i = 0 To UBound(rangenamestring) - 1
        firstbound = Range(rangenamestring(i)).Address
        temp = Split(firstbound, "$")
        upperbound = UBound(temp)
        temp = temp(UBound(temp))
        x = CLng(temp) + numberofrows
        lastbound = Replace(firstbound, temp, x)
        If upperbound < 3 Then
            RangeAddress = firstbound & ":" & lastbound
        Else
            RangeAddress = lastbound
        End If
        ThisWorkbook.Names.add name:=rangenamestring(i), _
                 RefersTo:="=" & RangeAddress, Visible:=True
    Next
End Sub

Function InsertDiffRowsAndFillFormulas(Optional vRows As Long = 0) As Long
    Dim x As Long
    Dim sht As Worksheet
    Dim shts() As String
    Dim i As Long
    
    ActiveCell.EntireRow.Select  'So you do not have to preselect entire row
    If vRows = 0 Then
        vRows = Application.InputBox(prompt:= _
        "Enter the number of rows you want to add below selected cell", Title:="Add Rows below the selected cell", _
        Default:=1, Type:=1) 'Default for 1 row, type 1 is number
        If vRows = False Then
            InsertDiffRowsAndFillFormulas = 0
            Exit Function
        End If
    End If
    
    sPassword = EfilingCommon.getmsgstate
    ActiveSheet.Unprotect Password:=sPassword
    
    ReDim shts(1 To Worksheets.Application.ActiveWorkbook. _
    Windows(1).SelectedSheets.count)
    i = 0
    For Each sht In _
        Application.ActiveWorkbook.Windows(1).SelectedSheets
        Sheets(sht.name).Select
        i = i + 1
        shts(i) = sht.name
        
        x = Sheets(sht.name).UsedRange.Rows.count 'lastcell fixup
        
        Selection.Resize(rowsize:=2).Rows(2).EntireRow. _
        Resize(rowsize:=vRows).Insert Shift:=xlDown
        
        ActiveSheet.Unprotect Password:=sPassword
        
        Selection.AutoFill Selection.Resize( _
        rowsize:=vRows + 1), xlFillDefault
        
        On Error Resume Next
        
        Selection.Offset(1).Resize(vRows).EntireRow. _
        SpecialCells(xlCellTypeAllValidation).ClearContents
    Next sht
    Worksheets(shts).Select
    InsertDiffRowsAndFillFormulas = vRows
    ActiveSheet.Protect Password:=sPassword
End Function


Function Dformat(dt As Variant, timepass As String) As String
'yyyy-mm-dd'

Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(dt) > 0 Then

Year = Mid(dt, 7, 4)
month = Mid(dt, 4, 2)
day = Mid(dt, 1, 2)
formateddate = Year & "-" & month & "-" & day
Dformat = formateddate

Else
Dformat = ""
End If
End Function

'Added by Shrutika- 16/01/2026 AY-26-27

Function Dformat1(dt As Variant, timepass As String) As String
'yyyy-mm-dd'
 
Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(SheetALL.Range("DOB_1").value) > 0 Then
 
Year = Mid(SheetALL.Range("DOB_1").value, 7, 4)
month = Mid(SheetALL.Range("DOB_1").value, 4, 2)
day = Mid(SheetALL.Range("DOB_1").value, 1, 2)
formateddate = day & "/" & month & "/" & Year
Dformat1 = formateddate
 
Else
Dformat1 = ""
End If
End Function

'Added by Shrutika- 16/01/2026 AY-26-27

Function Dformat3(dt As Variant, timepass As String) As String
'yyyy-mm-dd'
 
Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(SheetALL.Range("date_purchase1").value) > 0 Then
 
Year = Mid(SheetALL.Range("date_purchase1").value, 7, 4)
month = Mid(SheetALL.Range("date_purchase1").value, 4, 2)
day = Mid(SheetALL.Range("date_purchase1").value, 1, 2)
formateddate = day & "/" & month & "/" & Year
Dformat3 = formateddate
 
Else
Dformat3 = ""
End If
End Function

Function Dformat4(dt As Variant, timepass As String) As String
'yyyy-mm-dd'
 
Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(SheetALL.Range("IT_Date").value) > 0 Then
 
Year = Mid(SheetALL.Range("IT_Date").value, 7, 4)
month = Mid(SheetALL.Range("IT_Date").value, 4, 2)
day = Mid(SheetALL.Range("IT_Date").value, 1, 2)
formateddate = Year & "-" & month & "-" & day
Dformat4 = formateddate
 
Else
Dformat4 = ""
End If
End Function

Function Dformat5(dt As Variant, timepass As String) As String
'yyyy-mm-dd'
 
Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(SheetALL.Range("IT_Date").value) > 0 Then
 
Year = Mid(SheetALL.Range("IT_Date").value, 7, 4)
month = Mid(SheetALL.Range("IT_Date").value, 4, 2)
day = Mid(SheetALL.Range("IT_Date").value, 1, 2)
formateddate = day & "/" & month & "/" & Year
Dformat5 = formateddate
 
Else
Dformat5 = ""
End If
End Function



Function GetBankName(IFSC As Variant) As String
On Error Resume Next
GetBankName = ""
Dim rangecells As Range
Dim Address As Variant
For Each rangecells In SheetALL.Range("IFSC_StartLetter")

If UCase(rangecells.value) = UCase(IFSC) Then
Address = rangecells.Address
Address = Replace(Address, "CS", "CR")
GetBankName = SheetALL.Range(Address).value
Exit Function
End If

Next

'If GetBankName = "" Then
'GetBankName = False
'End If

End Function

Sub lockall()
Dim wor As Worksheet

For Each wor In ThisWorkbook.Worksheets
    wor.Protect Password:=getmsgstate
Next

End Sub


Function CheckfewSpecialCharacter(field As Variant) As Boolean
On Error Resume Next
    Dim specialCharArray As Variant
    Dim iCharCount, iSpecialChar As Long
    
    CheckfewSpecialCharacter = True
    specialCharArray = Array("&", """", "'", ">", "<")
    For iCharCount = 1 To Len(field)
        For iSpecialChar = 0 To UBound(specialCharArray)
        If Mid(field, iCharCount, 1) = specialCharArray(iSpecialChar) Then
            CheckfewSpecialCharacter = False
            Exit Function
        End If
        Next
    Next
End Function



Function ValidateAadharEnrol(aadhar As Variant) As Boolean
On Error Resume Next
ValidateAadharEnrol = True
    If Len(aadhar) > 0 Then
        If Not IsNumeric(Mid(aadhar, 1, 4)) Then
            ValidateAadharEnrol = False
            Exit Function
        End If
        
        If Not (Mid(aadhar, 5, 1)) = "/" Then
            ValidateAadharEnrol = False
            Exit Function
        End If
        
 
        If Not IsNumeric(Mid(aadhar, 6, 5)) Then
            ValidateAadharEnrol = False
            Exit Function
        End If
        
        If Not (Mid(aadhar, 11, 1)) = "/" Then
            ValidateAadharEnrol = False
            Exit Function
        End If
        
        If Not IsNumeric(Mid(aadhar, 12, 5)) Then
            ValidateAadharEnrol = False
            Exit Function
        End If
        
        If Not (Mid(aadhar, 17, 1)) = " " Then
            ValidateAadharEnrol = False
            Exit Function
        End If
        
        If Not IsNumeric(Mid(aadhar, 18, 2)) Then
            ValidateAadharEnrol = False
            Exit Function
        End If
        
        If Not (Mid(aadhar, 20, 1)) = "/" Then
            ValidateAadharEnrol = False
            Exit Function
        End If
        
        If Not IsNumeric(Mid(aadhar, 21, 2)) Then
            ValidateAadharEnrol = False
            Exit Function
        End If
        
        If Not (Mid(aadhar, 23, 1)) = "/" Then
            ValidateAadharEnrol = False
            Exit Function
        End If
        
        If Not IsNumeric(Mid(aadhar, 24, 4)) Then
            ValidateAadharEnrol = False
            Exit Function
        End If
        
        If Not (Mid(aadhar, 28, 1)) = " " Then
            ValidateAadharEnrol = False
            Exit Function
        End If
        
        If Not IsNumeric(Mid(aadhar, 29, 2)) Then
            ValidateAadharEnrol = False
            Exit Function
        End If
        
        If Not (Mid(aadhar, 31, 1)) = ":" Then
            ValidateAadharEnrol = False
            Exit Function
        End If
        
        If Not IsNumeric(Mid(aadhar, 32, 2)) Then
            ValidateAadharEnrol = False
            Exit Function
        End If
        
        If Not (Mid(aadhar, 34, 1)) = ":" Then
            ValidateAadharEnrol = False
            Exit Function
        End If
        
        If Not IsNumeric(Mid(aadhar, 35, 2)) Then
            ValidateAadharEnrol = False
            Exit Function
        End If
    End If
End Function

Function GetAadharEnrol(aadhar As Variant) As Variant
On Error Resume Next
    GetAadharEnrol = Mid(aadhar, 1, 4) & Mid(aadhar, 6, 5) & Mid(aadhar, 12, 5) & _
                     Mid(aadhar, 18, 2) & Mid(aadhar, 21, 2) & Mid(aadhar, 24, 4) & _
                     Mid(aadhar, 29, 2) & Mid(aadhar, 32, 2) & Mid(aadhar, 35, 2)
End Function

Function GetMergedAddressCell(rnge As Range, i As Long) As Variant
    Dim sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue As String
    Dim sTempCellValueInt As String

    sTempCellValue = Replace(rnge.AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
    sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
    sTempLastCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C6
    sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
    
    sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + i - 1)
    GetMergedAddressCell = sNewCellValue
End Function

'Sub RadioButton1_Click()
'Sheet1.Range("sheet1.ReturnFileSec").value = "139(1)-On or before due date"
'Sheet1.Shapes("Option Button 451").OLEFormat.Object = 1
'Sheet1.Shapes("Option Button 452").OLEFormat.Object = 0
'Sheet1.Range("Temp_RadioButton") = 1
'Sheet1.Activate
'End Sub
'
'Sub RadioButton2_Click()
'Sheet1.Range("sheet1.ReturnFileSec").value = "(Select)"
'Sheet1.Shapes("Option Button 451").OLEFormat.Object = 0
'Sheet1.Shapes("Option Button 452").OLEFormat.Object = 1
'Sheet1.Range("Temp_RadioButton") = 2
'Sheet1.Activate
'End Sub

Function ValidateTenantPan(tenantPan As String) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.
'It is not Validatiing 4th character to p or H
'status = Sheet1.Range("sheet1.Status").value

    ValidateTenantPan = True
    If Len(tenantPan) > 0 Then
        If Not ChkAlphabet(Mid(tenantPan, 1, 1)) Then
            ValidateTenantPan = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(tenantPan, 2, 1)) Then
            ValidateTenantPan = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(tenantPan, 3, 1)) Then
            ValidateTenantPan = False
            Exit Function
        End If

        If Not ChkAlphabet(Mid(tenantPan, 4, 1)) Then
            ValidateTenantPan = False
            Exit Function
        End If

        If Not ChkAlphabet(Mid(tenantPan, 5, 1)) Then
            ValidateTenantPan = False
            Exit Function
        End If
        If Not IsNumeric(Mid(tenantPan, 6, 4)) Then
            ValidateTenantPan = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(tenantPan, 10, 1)) Then
            ValidateTenantPan = False
            Exit Function
        End If
    End If
End Function

Private Sub ButtonHyperlink()
On Error Resume Next
fmsgboxsmall "After e-payment of Tax, the details of amount paid, Challan No., BSR Code, etc. should be filled in Schedule IT of Income Tax Return before submission of the return to claim this challan."
ActiveWorkbook.FollowHyperlink _
Address:="https://eportal.incometax.gov.in/iec/foservices/#/e-pay-tax-prelogin/user-details" 'added by shrutika(10/06/25)
'Address:="https://onlineservices.tin.egov-nsdl.com/etaxnew/tdsnontds.jsp"



End Sub
Sub LinkCheckBoxes()
Dim chk As CheckBox
Dim lCol As Long
lCol = 8 'number of columns to the right for link
For Each chk In ActiveSheet.CheckBoxes
   With chk
      .LinkedCell = _
         .TopLeftCell.Offset(0, lCol).Address
   End With
Next chk
End Sub

Sub LockUnlockCoOwnersDetails(HpCount As Long)
On Error Resume Next
'Dim HpCount  As Long
'For HpCount = 1 To Sheet3.Range("HousePropertySectionCount").value

   If Sheet8.Range("HP.CoOwnedYN" & HpCount).value = "No" Then
            
                 Sheet8.Range("Co_OwnnerBlock" & HpCount).ClearContents
                 Sheet8.Range("Co_OwnnerBlock" & HpCount).Locked = True
                 Sheet8.Range("Co_OwnnerBlock" & HpCount).Interior.Color = (&HD8D8D8)
                 Sheet8.Range("HP.SharePercent" & HpCount).value = "100"
                 Sheet8.Range("HP.SharePercent" & HpCount).Locked = True
                 Sheet8.Range("HP.SharePercent" & HpCount).Interior.ColorIndex = 2
  ElseIf Sheet8.Range("HP.CoOwnedYN" & HpCount).value = "Yes" Then
                 Sheet8.Range("Co_OwnnerBlock" & HpCount).Locked = False
                 Sheet8.Range("Co_OwnnerBlock" & HpCount).Interior.Color = (&HCCFFCC)
                 Sheet8.Range("HP.SharePercent" & HpCount).value = WorksheetFunction.Max(0, 100 - WorksheetFunction.Sum(Sheet8.Range("HP.Co.Share" & HpCount).value))
                 Sheet8.Range("HP.SharePercent" & HpCount).Locked = True
                 Sheet8.Range("HP.SharePercent" & HpCount).Interior.ColorIndex = 2
   ElseIf Sheet8.Range("HP.CoOwnedYN" & HpCount).value <> "No" Then
            
                Sheet8.Range("Co_OwnnerBlock" & HpCount).ClearContents
                Sheet8.Range("Co_OwnnerBlock" & HpCount).Locked = False
                Sheet8.Range("Co_OwnnerBlock" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet8.Range("HP.SharePercent" & HpCount).value = ""
                Sheet8.Range("HP.SharePercent" & HpCount).Locked = False
                Sheet8.Range("HP.SharePercent" & HpCount).Interior.Color = (&HCCFFCC)
    End If
'Next
End Sub

Function insertRowNatureOfBusiness(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0, Optional f_112A_115AD_1 As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
'Method : Insert no of Rows
    nRows = 0
    
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
    
     
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    ActiveSheet.Range("D" & ActiveCell.row & ":G" & ActiveCell.row).Select
    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    'Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlCellTypeAllValidation).ClearContents
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
    If f_112A_115AD_1 = 1 Then
        Dim rr, rrCount As Variant
        Dim RngHOI As Range
        Set RngHOI = Selection.Resize(rowsize:=nRows + 1)
        rrCount = 0
        For Each rr In RngHOI.Rows
            If Not rrCount = 0 Then
                Range("E" & rr.row).value = "(Select)"
            End If
            rrCount = rrCount + 1
        Next
    End If

    
'Method: DefineName after no of rows insert Method
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue
            sRangeValue1 = sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue
            sRangeValue1 = sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowNatureOfBusiness = nRows
'----------------Lock Password-------------------START---
   ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----

    Application.EnableEvents = True
End Function

Function CheckNameAlpha(field As Variant) As Boolean
    Dim k As Long
    Dim chkchar As String
    CheckNameAlpha = True
    
    For k = 1 To Len(field)
    chkchar = Mid(field, k, 1)
    
    If ((asc(chkchar) < 65) Or (asc(chkchar) > 90)) Then
        CheckNameAlpha = False
    
            Exit Function
        End If
    Next
End Function

Function ValidateTANFirstThreeLetters(fieldValue As Variant) As Boolean
ValidateTANFirstThreeLetters = False
Dim i, j As Long
Dim arr As Variant

arr = Array("HYD", "VPN", "BBN", "BPL", "JBP", "CHE", "CMB", "MRI", "DEL", "CAL", "MRT", "AHM", "BRD", "RKT", "SRT", "BLR", "AGR", "KNP", "CHN", "TVD", _
        "ALD", "LKN", "MUM", "NGP", "AMR", "JLD", "PTL", "RTK", "KLP", "NSK", "PNE", "PTN", "RCH", "JDH", "JPR", "SHL")

For i = 1 To Len(fieldValue)
    For j = 0 To UBound(arr)
        If Mid(fieldValue, i, 3) = arr(j) Then
            ValidateTANFirstThreeLetters = True
            Exit Function
        End If
    Next
Next

End Function

Sub x()
Application.EnableEvents = True
End Sub
'Malli-------------------07/22/2024
Sub ExendRangeNameToTable_M_VDA(numberofrows As Long, rangenamestring As Variant)
    Dim i As Long
    Dim x As Long
    Dim firstbound As String
    Dim temp As Variant
    Dim upperbound As String
    Dim lastbound As String
    
    rangenamestring = Split(rangenamestring, "||")
    For i = 0 To UBound(rangenamestring) - 1
        firstbound = Range(rangenamestring(i)).Address
        temp = Split(firstbound, "$")
        upperbound = UBound(temp)
        temp = temp(UBound(temp))
        x = CLng(temp) + numberofrows
        lastbound = Replace(firstbound, temp, x)
        If upperbound < 3 Then
            RangeAddress = firstbound & ":" & lastbound
        Else
            RangeAddress = lastbound
        End If
        ThisWorkbook.Names.add name:=rangenamestring(i), _
                 RefersTo:="=" & RangeAddress, Visible:=True
    Next
End Sub
'Malli--------------------
Function StateMatchesPin(StateCode As Variant, Pincode_pag2) As Boolean
    StateMatchesPin = False
    Dim state1, state2 As Variant
    
    
    state1 = Application.VLookup(Pincode_pag2, SheetALL.Range("Duplicate_V"), 2, False)
    
    state2 = Application.VLookup(Pincode_pag2, SheetALL.Range("Duplicate_V"), 3, False)

    If IsError(state1) Or IsError(state2) Then
    StateMatchesPin = False
    Exit Function
    End If
    
    If UCase(Mid(StateCode, 4)) = state1 Or UCase(Mid(StateCode, 4)) = state2 Then
    'MsgBox UCase(Mid(StateCode, 4))
    StateMatchesPin = True
    End If
    
End Function
'Malli---------------------
 Function state_Validation_B(ByVal pin_trgt_adrs As Variant, ByVal state_trgt_adrs As Variant)
    state_Validation = True
    Application.EnableEvents = False
    Dim PinCode As Range
        
            If Range(pin_trgt_adrs).value <> "" Then
              Dim state1 As Variant
                state1 = Application.VLookup(Range(pin_trgt_adrs).value, SheetALL.Range("All_PinCode_V"), 2, False)
                If Not (state1 = UCase(Mid(Range(state_trgt_adrs).value, 4)) Or EfilingCommon.StateMatchesPin(Range(state_trgt_adrs).value, Range(pin_trgt_adrs).value)) Then
                Range(pin_trgt_adrs).value = ""
                End If
            End If
         
                   
End Function
  
      'Malli----------
Function PinState_codes_validation_B(ByVal pin_trgt_adrs As Variant, ByVal state_trgt_adrs As Variant, ByVal country_trgt_adrs As Variant, ByVal Zip_trgt_adrs As Variant)
    PinState_codes_validation = True
    Application.EnableEvents = False
    Dim PinCode As Range
    Dim DupFlag As Boolean
    DupFlag = False
                 
                 For Each PinCode In SheetALL.Range("Duplicate_Pincode_list")
                    If Range(pin_trgt_adrs).value = PinCode Then
                        DupFlag = True
                        Exit For
                    End If
                 Next PinCode

                 If DupFlag = True Then
                    If Range(state_trgt_adrs).value <> "" Then
                        If Not EfilingCommon.StateMatchesPin(Range(state_trgt_adrs).value, Range(pin_trgt_adrs).value) Then
                            Range(state_trgt_adrs).value = "(Select)"
                        End If
                    End If
                 End If
                 
                 If DupFlag = False Then
                    Dim StateName
                    Dim StateCode

                    StateName = Application.VLookup(Range(pin_trgt_adrs).value, SheetALL.Range("All_PinCode_V"), 2, False)
                    If IsError(StateName) Then
                        If Range(pin_trgt_adrs).value <> "" Then
                        MsgBox ("Invalid Pincode")
                     End If
                        Application.EnableEvents = False
                        pin_trgt_adrs.value = ""
                        Application.EnableEvents = True

                    Else
                        StateCode = Application.VLookup(StateName, SheetALL.Range("EB:EC"), 2, False)
                        Range(state_trgt_adrs).value = StateCode
                        Range(country_trgt_adrs).value = "91-INDIA"
                        
                      
                         ActiveSheet.Unprotect Password:=EfilingCommon.getmsgstate
                        If Mid(Range(state_trgt_adrs).value, 1, 2) = "99" Then
                                Range(Zip_trgt_adrs).Interior.Color = (&HCCFFCC)
                                Range(Zip_trgt_adrs).value = ""
                                Range(Zip_trgt_adrs).Locked = False
                        ElseIf Mid(Range(state_trgt_adrs).value, 1, 2) <> "99" Then
                                Range(Zip_trgt_adrs).Interior.Color = (&HD8D8D8)
                                Range(Zip_trgt_adrs).value = ""
                                Range(Zip_trgt_adrs).Locked = True
                        End If
                    End If
                 End If
     
    
     
End Function
'Malli-----------------
Function SetStateCountryImmovableAsset_M(Tval As Variant, Tadd As Variant, CallType As Variant)
On Error Resume Next
Application.EnableEvents = False


'------------------Un Protect--------------------
   sPassword = EfilingCommon.getmsgstate
   Sheet43.Unprotect Password:=sPassword
'------------------Un Protect Ends--------------------
    
    Dim rangecellsState, rangecellsCountry, rangecellsPin, rangecellsZip As Range
    Dim mIntCells, q  As Long
    Dim PrevVal As Variant

    Set rangecellsState = Range("SchAL.A.Address_State").Cells
    Set rangecellsCountry = Range("SchAL.A.Address_Country").Cells
    Set rangecellsPin = Range("SchAL.A.Address_Pin").Cells
    Set rangecellsZip = Range("SchAL.A.Address_Zip").Cells

     'MsgBox rangecellsZip.Address

    mIntCells = Range("SchAL.A.Description").count

If CallType = "S" Then
    'If UCase(Tval) = "99-STATE OUTSIDE INDIA" Then
    'Mid(UCase(Tval), 1, 2) = "99"
      If Mid(UCase(Tval), 1, 2) = "99" Then
        Sheet43.Range(Tadd).Select
        Sheet43.Range(Tadd).Offset(0, 1) = "(Select)"
       Sheet43.Range(Tadd).Offset(0, 2) = ""
        Sheet43.Range(Tadd).Offset(0, 2).Interior.Color = (&HD8D8D8)
        Sheet43.Range(Tadd).Offset(0, 2).Locked = True
        Sheet43.Range(Tadd).Offset(0, 3).Interior.Color = (&HCCFFCC)
        Sheet43.Range(Tadd).Offset(0, 3).Locked = False
        GoTo endfd
    ElseIf UCase(Tval) = "(SELECT)" Then
        Sheet43.Range(Tadd).Select
        Sheet43.Range(Tadd).Offset(0, 1) = "(Select)"
        Sheet43.Range(Tadd).Offset(0, 1).Interior.Color = (&HCCFFCC)
        Sheet43.Range(Tadd).Offset(0, 1).Locked = False
        
        Sheet43.Range(Tadd).Offset(0, 2) = ""
        Sheet43.Range(Tadd).Offset(0, 2).Interior.Color = (&HCCFFCC)
        Sheet43.Range(Tadd).Offset(0, 2).Locked = False
        
        Sheet43.Range(Tadd).Offset(0, 3) = ""
        Sheet43.Range(Tadd).Offset(0, 3).Interior.Color = (&HCCFFCC)
        Sheet43.Range(Tadd).Offset(0, 3).Locked = False
        GoTo endfd
    Else
        Sheet43.Range(Tadd).Select
        Sheet43.Range(Tadd).Offset(0, 1) = "91-INDIA"
        
      If Sheet43.Range(Tadd).Offset(0, 2) = "" Then
        Sheet43.Range(Tadd).Offset(0, 2) = ""
        Sheet43.Range(Tadd).Offset(0, 2).Interior.Color = (&HCCFFCC)
        Sheet43.Range(Tadd).Offset(0, 2).Locked = False
        End If
        
        Sheet43.Range(Tadd).Offset(0, 3) = ""
        Sheet43.Range(Tadd).Offset(0, 3).Interior.Color = (&HD8D8D8)
        Sheet43.Range(Tadd).Offset(0, 3).Locked = True
        GoTo endfd
    End If
Else

    If UCase(Tval) = "(SELECT)" Then
        Sheet43.Range(Tadd).Select
        Sheet43.Range(Tadd).Offset(0, -1) = "(Select)"
        Sheet43.Range(Tadd).Offset(0, 1) = ""
        Sheet43.Range(Tadd).Offset(0, 1).Interior.Color = (&HCCFFCC)
        Sheet43.Range(Tadd).Offset(0, 1).Locked = False
        Sheet43.Range(Tadd).Offset(0, 2) = ""
        Sheet43.Range(Tadd).Offset(0, 2).Interior.Color = (&HD8D8D8)
        Sheet43.Range(Tadd).Offset(0, 2).Locked = True
        GoTo endfd
    ElseIf UCase(Tval) <> "91-INDIA" Then
        Sheet43.Range(Tadd).Select
        '17/10/2024----------------------Malli
        'Sheet43.Range(Tadd).Offset(0, -1) = "99-STATE OUTSIDE INDIA"
         Sheet43.Range(Tadd).Offset(0, -1) = "99-Foreign"
        '-----------------------------------
        Sheet43.Range(Tadd).Offset(0, 1) = ""
        Sheet43.Range(Tadd).Offset(0, 1).Interior.Color = (&HD8D8D8)
        Sheet43.Range(Tadd).Offset(0, 1).Locked = True
        Sheet43.Range(Tadd).Offset(0, 2).Interior.Color = (&HCCFFCC)
        Sheet43.Range(Tadd).Offset(0, 2).Locked = False
        GoTo endfd
    Else
        Sheet43.Range(Tadd).Select
        Sheet43.Range(Tadd).Offset(0, -1) = "(Select)"
        Sheet43.Range(Tadd).Offset(0, 1) = ""
        Sheet43.Range(Tadd).Offset(0, 1).Interior.Color = (&HCCFFCC)
        Sheet43.Range(Tadd).Offset(0, 1).Locked = False
        Sheet43.Range(Tadd).Offset(0, 2) = ""
        Sheet43.Range(Tadd).Offset(0, 2).Interior.Color = (&HD8D8D8)
        Sheet43.Range(Tadd).Offset(0, 2).Locked = True
        GoTo endfd
    End If

End If

    For q = 1 To mIntCells
    'Malli------------17/10/2024
        'If UCase(rangecellsState.item(q).value) = "99-STATE OUTSIDE INDIA" Then
        If UCase(rangecellsState.item(q).value) = "99-FOREIGN" Then    '99-Foreign
        '--------------------------
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
        '17/10/2024-----------------------Malli
            'rangecellsState.item(q).value = "99-STATE OUTSIDE INDIA"
            rangecellsState.item(q).value = "99-Foreign"
            '--------------------------
            rangecellsPin.item(q).value = ""
            rangecellsZip.item(q).value = ""
        Else
            rangecellsCountry.item(q).value = "(Select)"
            rangecellsPin.item(q).value = ""
            rangecellsZip.item(q).value = ""
        End If
        '99-Foreign
    Next
endfd:
'----------------Lock Password-------------------START---
  Sheet43.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
Application.EnableEvents = True
End Function
 'Malli----------------
Function PinState_codes_validation_80G(ByVal pin_trgt_adrs As Variant, ByVal state_trgt_adrs As Variant)
    PinState_codes_validation_80G = True
    Application.EnableEvents = False
    Dim PinCode As Range
    Dim DupFlag As Boolean
    DupFlag = False
                 
                 For Each PinCode In SheetALL.Range("Duplicate_Pincode_list")
                    If Range(pin_trgt_adrs).value = PinCode Then
                        DupFlag = True
                        Exit For
                    End If
                 Next PinCode

                 If DupFlag = True Then
                    If Range(state_trgt_adrs).value <> "" Then
                        If Not EfilingCommon.StateMatchesPin(Range(state_trgt_adrs).value, Range(pin_trgt_adrs).value) Then
                            Range(state_trgt_adrs).value = "(Select)"
                        End If
                    End If
                 End If
                 
                 If DupFlag = False Then
                    Dim StateName
                    Dim StateCode

                    StateName = Application.VLookup(Range(pin_trgt_adrs).value, SheetALL.Range("All_PinCode_V"), 2, False)
                     If IsError(StateName) Then
                        If Range(pin_trgt_adrs).value <> "" Then
                        MsgBox ("Invalid Pincode")
                        End If
                        Application.EnableEvents = False
                        pin_trgt_adrs.value = ""
                        Application.EnableEvents = True

                     Else
                        StateCode = Application.VLookup(StateName, SheetALL.Range("EB:EC"), 2, False)
                        Range(state_trgt_adrs).value = StateCode
                        
                     End If
                 End If
     
    
     
End Function
'Bindu Modified one final 16_09_2024
Function state_Validation(sheetname2 As Worksheet, ByVal pin_trgt_adrs As String, ByVal state_trgt_adrs As String) As Boolean
    state_Validation = True
 
    Dim PinCode As Range
    
    sheetname2.Activate
        
            If Range(pin_trgt_adrs).value <> "" Then
              Dim state1 As String
                state1 = Application.VLookup(Range(pin_trgt_adrs).value, SheetALL.Range("All_PinCode_V"), 2, False)
                If Not (state1 = UCase(Mid(Range(state_trgt_adrs).value, 4)) Or EfilingCommon.StateMatchesPin(Range(state_trgt_adrs).value, Range(pin_trgt_adrs).value)) Then
                    state_Validation = False
    
                    'Range(pin_trgt_adrs).value = ""
                End If
            End If
         
                   
End Function
      'Bindu Modified one final 16_09_2024
Function PinState_codes_validation(sheetname1 As Worksheet, ByVal pin_trgt_adrs As String, ByVal state_trgt_adrs As String, ByVal country_trgt_adrs As String, ByVal Zip_trgt_adrs As String)
    Dim DupFlag As Boolean
Dim PinCode1 As String
DupFlag = False
PinCode1 = Application.IsError(Application.VLookup(sheetname1.Range(pin_trgt_adrs).value, SheetALL.Range("Duplicate_Pincode_list"), 1, False))
If PinCode1 = "False" Then DupFlag = True

'changed by Chetan C M on 30/07/2025
'start--

'If DupFlag = True Then
'   If sheetname1.Range(state_trgt_adrs).value <> "" Then
'       If Not EfilingCommon.StateMatchesPin(Range(state_trgt_adrs).value, sheetname1.Range(pin_trgt_adrs).value) Then
'                sheetname1.Range(state_trgt_adrs).value = "(Select)"
'       End If
'   End If
'End If
If DupFlag = True Then
   If sheetname1.Range(state_trgt_adrs).value <> "" Or sheetname1.Range(state_trgt_adrs).value = "" Then
       If Not StateMatchesPin(Range(state_trgt_adrs).value, sheetname1.Range(pin_trgt_adrs).value) Then
           sheetname1.Unprotect Password:=EfilingCommon.getmsgstate
           sheetname1.Range(state_trgt_adrs).value = "(Select)"
           sheetname1.Range(country_trgt_adrs).value = "(Select)"
           sheetname1.Protect Password:=EfilingCommon.getmsgstate
       End If
   End If
End If
'--end
                 
If DupFlag = False Then
   Dim StateName
   Dim StateCode

   StateName = Application.VLookup(sheetname1.Range(pin_trgt_adrs).value, SheetALL.Range("All_PinCode_V"), 2, False)
   If IsError(StateName) Then
       If sheetname1.Range(pin_trgt_adrs).value <> "" Then
           
           MsgBox ("Invalid Pincode")
           sheetname1.Range(pin_trgt_adrs).ClearContents
       End If
   Else
       StateCode = Application.VLookup(StateName, SheetALL.Range("EB:EC"), 2, False)
       sheetname1.Range(state_trgt_adrs).value = StateCode
       sheetname1.Range(country_trgt_adrs).value = "91-INDIA"
       sheetname1.Unprotect Password:=EfilingCommon.getmsgstate
       'added by Chetan C M for on 30/07/2025
       'start--
        If Mid(Range(state_trgt_adrs).value, 1, 2) = "99" Then
                Range(pin_trgt_adrs).Interior.Color = (&HD8D8D8)
                Range(pin_trgt_adrs).value = ""
                Range(country_trgt_adrs).value = "(Select)"
                Range(pin_trgt_adrs).Locked = True
                Range(Zip_trgt_adrs).Interior.Color = (&HCCFFCC)
                Range(Zip_trgt_adrs).value = ""
                Range(Zip_trgt_adrs).Locked = False
        ElseIf Mid(Range(state_trgt_adrs).value, 1, 2) <> "99" Then
                Range(pin_trgt_adrs).Interior.Color = (&HCCFFCC)
                Range(pin_trgt_adrs).Locked = False
                Range(Zip_trgt_adrs).Interior.Color = (&HD8D8D8)
                Range(Zip_trgt_adrs).value = ""
                Range(Zip_trgt_adrs).Locked = True
        End If
        sheetname1.Protect Password:=EfilingCommon.getmsgstate
       '--end
   End If
End If


End Function


 'Bindu Modified one final 16_09_2024
Function PIN_Validation(ByVal pin_trgt_adrs As String, ByVal state_trgt_adrs As String) As Boolean
    state_Validation = True
    
    Dim PinCode As Range
        
            If Range(pin_trgt_adrs).value <> "" Then
              Dim state1 As String
                state1 = Application.VLookup(Range(pin_trgt_adrs).value, SheetALL.Range("All_PinCode_V"), 2, False)
                If Not (state1 = UCase(Mid(Range(state_trgt_adrs).value, 4)) Or EfilingCommon.StateMatchesPin(Range(state_trgt_adrs).value, Range(pin_trgt_adrs).value)) Then
                    state_Validation = False
    
                    'Range(pin_trgt_adrs).value = ""
                End If
            End If
         
                   
End Function

 

'17/10/2024----------Malli
'Bindu Modified one final 16_09_2024
Function PINstate_ModualValidation(sheetname2 As Worksheet, ByVal pin_trgt_adrs As String, ByVal state_trgt_adrs As String) As Boolean
    PINstate_ModualValidation = True
    
    Dim PinCode As Range
    
    sheetname2.Activate
        
            If Range(pin_trgt_adrs).value <> "" Then
                Dim state1 As String
                
                state1 = Application.IsError(Application.VLookup(Range(pin_trgt_adrs).value, SheetALL.Range("All_PinCode_V"), 2, False))
                
                If state1 = True Then
                    PINstate_ModualValidation = False
                Else
                    state1 = Application.VLookup(Range(pin_trgt_adrs).value, SheetALL.Range("All_PinCode_V"), 2, False)
                    If Not (state1 = UCase(Mid(Range(state_trgt_adrs).value, 4)) Or EfilingCommon.StateMatchesPin(Range(state_trgt_adrs).value, Range(pin_trgt_adrs).value)) Then
                    PINstate_ModualValidation = False
                    End If
                 End If
                End If
                   
End Function

'Malli-------------AY_2025_26
Sub UNLOCKRANGEcg_8tabA1toA7()
Application.EnableEvents = False
 sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
Dim i As Long
Dim chcell As Range
Dim rangenamestring As Variant
rangenamestring = "STCG.A8_AmtI||STCG.A8_IA1tA6||STCG.A8_Country||STCG.A8_DTAA_Article||STCG.A8_RatePTreaty||STCG.A8_CertiStatus||STCG.A8_SecIT||STCG.A8_RateIT"
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet13.Range(rangenamestring(i))
            chcell.Locked = False
            chcell.Interior.Color = (&HCCFFCC)
'            chcell.value = 0'
                    If rangenamestring(i) = "STCG.A8_IA1tA6" Or rangenamestring(i) = "STCG.A8_Country" Or rangenamestring(i) = "STCG.A8_CertiStatus" Then
                        chcell.value = "(Select)"
                    End If
            Next
        Next
        Application.EnableEvents = True
End Sub
'Malli-------------AY_2025_26
Sub LOCKRANGEcg_8tabA1toA7()
     sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
Dim i As Long
Application.EnableEvents = False
Dim chcell As Range
Dim rangenamestring As Variant
rangenamestring = "STCG.A8_AmtI||STCG.A8_IA1tA6||STCG.A8_Country||STCG.A8_DTAA_Article||STCG.A8_RatePTreaty||STCG.A8_CertiStatus||STCG.A8_SecIT||STCG.A8_RateIT"
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet13.Range(rangenamestring(i))
            chcell.value = ""
            chcell.Locked = True
            chcell.Interior.Color = (&HD8D8D8)

            Next
        Next
Application.EnableEvents = False
End Sub
'Malli-------------AY_2025_26
Sub UNLOCKRANGEcg_12tbB1toB11()
 sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
Application.EnableEvents = False
Dim i As Long
Dim chcell As Range
Dim rangenamestring As Variant
rangenamestring = "LTCG.B10_LtcgAmt||LTCG.B10_ItemNo||LTCG.B10_Country||LTCG.B10_DTAA_Article||LTCG.B10_RatepTreaty||LTCG.B10_CertiStatus||LTCG.B10_SecIT||LTCG.B10_RatePITAct"
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet13.Range(rangenamestring(i))
            chcell.Locked = False
            chcell.Interior.Color = (&HCCFFCC)
'            chcell.value = 0'
                    If rangenamestring(i) = "LTCG.B10_ItemNo" Or rangenamestring(i) = "LTCG.B10_Country" Or rangenamestring(i) = "LTCG.B10_CertiStatus" Then
                        chcell.value = "(Select)"
                    End If
            Next
        Next
        Application.EnableEvents = True
End Sub
'Malli-------------AY_2025_26
Sub LOCKRANGEcg_12tbB1toB11()
     sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword
Dim i As Long
Application.EnableEvents = False
Dim chcell As Range
Dim rangenamestring As Variant
rangenamestring = "LTCG.B10_LtcgAmt||LTCG.B10_ItemNo||LTCG.B10_Country||LTCG.B10_DTAA_Article||LTCG.B10_RatepTreaty||LTCG.B10_CertiStatus||LTCG.B10_SecIT||LTCG.B10_RatePITAct"
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet13.Range(rangenamestring(i))
            chcell.value = ""
            chcell.Locked = True
            chcell.Interior.Color = (&HD8D8D8)

            Next
        Next
Application.EnableEvents = False
End Sub

'Ankita_09/04
Function insertRowUnderSectionWithFormula80D(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0, Optional HOIflag As Variant = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim sPassword As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue, sAddingRangeValue, sNewCellValueNext As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
    ''''''''''malli
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
     
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
    If HOIflag = 1 Then
        Dim rr, rrCount As Variant
        Dim RngHOI As Range
        'Set RngHOI = Selection.Resize(rowsize:=nRows + 1)
        Set RngHOI = Selection.Offset(1).Resize(rowsize:=1)
        rrCount = 1
        For Each rr In RngHOI.Rows
            If Not rrCount = 0 Then
'                Range("D" & rr.row).Locked = False
'                Range("D" & rr.row).Interior.Color = (&HCCFFCC)
                Range("E" & rr.row).Locked = False
                Range("E" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("F" & rr.row).Locked = False
'                Range("F" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("G" & rr.row).Locked = False
'                Range("G" & rr.row).Interior.Color = (&HCCFFCC)
                Range("H" & rr.row).Locked = False
                Range("H" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("I" & rr.row).Locked = False
'                Range("I" & rr.row).Interior.Color = (&HCCFFCC)
                Range("J" & rr.row).Locked = False
                Range("J" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("K" & rr.row).Locked = False
'                Range("K" & rr.row).Interior.Color = (&HCCFFCC)
                Range("L" & rr.row).Locked = False
                Range("L" & rr.row).Interior.Color = (&HCCFFCC)

            End If
            rrCount = rrCount + 1
        Next
        
    End If

    
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    sNewCellValueNext = ""
    sAddingRangeValue = ""
    
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            
            sNewCellValueNext = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + 1) 'C6 + 1 => C7
            If nRows > 1 Then
            sAddingRangeValue = sNewCellValueNext + ":" + sNewCellValue
            ElseIf nRows = 1 Then
            sAddingRangeValue = sNewCellValue
            End If
            
'            If gridRange(iCount) = "TDS2.TdsCredit" Or _
'                gridRange(iCount) = "TDS3.TdsCredit" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
'
'            If gridRange(iCount) = "TDSoth.DeductedYear" Or _
'                gridRange(iCount) = "TDSoth2.DeductedYear" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
            
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            
            sNewCellValueNext = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + 1) 'C6 + 1 => C7
            If nRows > 1 Then
            sAddingRangeValue = sNewCellValueNext + ":" + sNewCellValue
            ElseIf nRows = 1 Then
            sAddingRangeValue = sNewCellValue
            End If
            
'            If gridRange(iCount) = "TDS2.TdsCredit" Or _
'                gridRange(iCount) = "TDS3.TdsCredit" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
'
'            If gridRange(iCount) = "TDSoth.DeductedYear" Or _
'                gridRange(iCount) = "TDSoth2.DeductedYear" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
            
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormula80D = nRows

'----------------Lock Password-------------------START---
  ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
    Application.EnableEvents = True
End Function

Function ChkMinDOBDate_2526(dob As Variant, minDefinedDOB As Variant) As Boolean
On Error Resume Next
     ChkMinDOBDate_2526 = True
     If Len(dob) > 0 Then
        If val(Mid(dob, 7, 4)) <= 2025 Then
              If val(Mid(dob, 4, 2)) <= 3 Then
                If val(Mid(dob, 1, 2)) <= 31 Then
                    ChkMinDOBDate_2526 = False
                    Exit Function
                End If
            End If
        End If
     End If
End Function
Function insertRowUnderSectionWithFormula_80CCC_VIA(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
    '----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
'Method: DefineName after no of rows insert Method
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormula_80CCC_VIA = nRows
'----------------Lock Password-------------------START---
'----------------Lock Password-------------------END-----

    Application.EnableEvents = True
    ActiveSheet.Protect Password:=sPassword
End Function

